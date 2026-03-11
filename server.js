const express = require('express');
const session = require('express-session');
const path = require('path');
const fs = require('fs');
const https = require('https');
const bcrypt = require('bcryptjs');
const QRCode = require('qrcode');
const multer = require('multer');
const XLSX = require('xlsx');
const crypto = require('crypto');
const db = require('./db');
const { generateBadgePdf, generateBadgePdfFromTemplate } = require('./lib/badge-pdf');
const { sendWhatsAppOTP, normalizePhone } = require('./lib/whatsapp');
const EventEmitter = require('events');
const notificationEmitter = new EventEmitter();
const notifications = [];

const MAX_NOTIFICATIONS = 50;

function addNotification(type, employe) {
  const notification = {
    id: Date.now(),
    type,
    employe,
    message: type === 'entrer' ? `Arrivée de ${employe.prenom} ${employe.nom}` : `Sortie de ${employe.prenom} ${employe.nom}`,
    time: new Date().toISOString()
  };
  notifications.unshift(notification);
  if (notifications.length > MAX_NOTIFICATIONS) {
    notifications.pop();
  }
  notificationEmitter.emit('pointage', notification);
}

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 5 * 1024 * 1024 } });

const app = express();
const PORT = process.env.PORT || 3000;

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));
app.get('/logo.png', (req, res) => {
  const logoPath = path.join(__dirname, 'logo.png');
  if (!fs.existsSync(logoPath)) {
    res.status(404).type('png').send(Buffer.alloc(0));
    return;
  }
  res.sendFile(logoPath);
});
const badgePdfPath = path.join(__dirname, 'Badge.pdf');
app.get('/Badge.pdf', (req, res) => {
  if (fs.existsSync(badgePdfPath)) {
    res.sendFile(badgePdfPath);
  } else {
    res.status(404).send('Badge.pdf non trouvé. Placez le fichier Badge.pdf à la racine du projet.');
  }
});
app.use(session({
  secret: process.env.SESSION_SECRET || 'presence-aris-secret-key-2024',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 24 * 60 * 60 * 1000 }
}));

function requireAuth(req, res, next) {
  if (!req.session.userId) return res.redirect('/login');
  if (!req.session.role && req.session.userId) {
    const u = db.prepare('SELECT role FROM users WHERE id = ?').get(req.session.userId);
    if (u) req.session.role = u.role;
  }
  next();
}
function requireAdmin(req, res, next) {
  if (!req.session.userId) return res.redirect('/login');
  const user = db.prepare('SELECT role FROM users WHERE id = ?').get(req.session.userId);
  if (!user || user.role !== 'admin') return res.redirect('/');
  next();
}

// ---------- Interface Employé (Scan Badge) ----------
app.get('/employe-login', (req, res) => {
  if (req.session.employeId) return res.redirect('/employe');
  res.render('employe-login', { error: null });
});

app.post('/employe-login', (req, res) => {
  const { code } = req.body || {};
  if (!code) {
    return res.render('employe-login', { error: 'Veuillez entrer votre code' });
  }
  const employe = db.prepare('SELECT * FROM employees WHERE badge_id = ? OR id_affichage = ?').get(code, code);
  if (!employe) {
    return res.render('employe-login', { error: 'Code invalide' });
  }
  req.session.employeId = employe.id;
  req.session.employeName = employe.prenom + ' ' + employe.nom;
  res.redirect('/employe');
});

// ---------- Interface Responsable Scan (Login) ----------
app.get('/scan-login', (req, res) => {
  if (req.session.scanUserId) return res.redirect('/scan');
  res.render('scan-login', { error: null });
});

app.post('/scan-login', (req, res) => {
  const { username, password } = req.body || {};
  const user = db.prepare('SELECT * FROM users WHERE username = ?').get(username);
  if (!user || !bcrypt.compareSync(password, user.password_hash)) {
    return res.render('scan-login', { error: 'Identifiants incorrects' });
  }
  req.session.scanUserId = user.id;
  req.session.scanUserName = user.username;
  res.redirect('/scan');
});

app.get('/scan', (req, res) => {
  if (!req.session.scanUserId) return res.redirect('/scan-login');
  res.render('scan-dashboard', { user: req.session });
});

app.post('/api/scan-badge', (req, res) => {
  if (!req.session.scanUserId) {
    return res.json({ ok: false, error: 'Non connecté' });
  }
  
  const { badgeCode } = req.body || {};
  if (!badgeCode) {
    return res.json({ ok: false, error: 'Code badge requis' });
  }
  
  const employe = db.prepare('SELECT * FROM employees WHERE badge_id = ? OR id_affichage = ?').get(badgeCode, badgeCode);
  if (!employe) {
    return res.json({ ok: false, error: 'Employé non trouvé', type: 'invalide' });
  }

  // Use UTC for consistent time handling
  const now = new Date();
  const today = now.toISOString().split('T')[0];

  // On vérifie le dernier scan de l'employé pour AUJOURD'HUI
  const lastPresence = db.prepare(`
    SELECT * FROM presence 
    WHERE employee_id = ? 
    AND date(scanned_at) = ?
    ORDER BY scanned_at DESC LIMIT 1
  `).get(employe.id, today);
  
  const type = (lastPresence && lastPresence.type === 'entrer') ? 'sortie' : 'entrer';
  
  // Insertion avec l'heure locale du serveur
  const now = new Date();
  const localTime = now.getFullYear() + '-' + 
    String(now.getMonth() + 1).padStart(2, '0') + '-' + 
    String(now.getDate()).padStart(2, '0') + ' ' + 
    String(now.getHours()).padStart(2, '0') + ':' + 
    String(now.getMinutes()).padStart(2, '0') + ':' + 
    String(now.getSeconds()).padStart(2, '0');
  db.prepare('INSERT INTO presence (employee_id, type, scanned_at) VALUES (?, ?, ?)').run(employe.id, type, localTime);
  
  notificationEmitter.emit('pointage', {
    type: type,
    employe: { id: employe.id, nom: employe.nom, prenom: employe.prenom }
  });
  addNotification(type, { id: employe.id, nom: employe.nom, prenom: employe.prenom });
  
  res.json({ 
    ok: true, 
    type,
    employe: {
      id: employe.id,
      nom: employe.nom,
      prenom: employe.prenom,
      poste: employe.poste,
      photo: employe.photo || null,
      badge_id: employe.badge_id,
      id_affichage: employe.id_affichage
    }
  });
});
app.get('/scan-logout', (req, res) => {
  req.session.scanUserId = null;
  req.session.scanUserName = null;
  res.redirect('/scan-login');
});

// Scanner password change
app.get('/scan-password', (req, res) => {
  if (!req.session.scanUserId) return res.redirect('/scan-login');
  res.render('scan-password', { user: req.session, error: null, success: null });
});

// Scanner: view presences by date
app.get('/scan-presences', (req, res) => {
  if (!req.session.scanUserId) return res.redirect('/scan-login');
  
  const now = new Date();
  const date = req.query.date || now.toISOString().split('T')[0];
  const presences = db.prepare(`
    SELECT p.*, e.badge_id, e.nom, e.prenom, e.poste
    FROM presence p
    JOIN employees e ON e.id = p.employee_id
    WHERE date(p.scanned_at) = ?
    ORDER BY p.scanned_at DESC
  `).all(date);
  
  const totalEmployees = db.prepare('SELECT COUNT(*) as c FROM employees').get().c;
  const presentToday = db.prepare(`
    SELECT COUNT(DISTINCT employee_id) as c FROM presence
    WHERE date(scanned_at) = ? AND type = 'entrer'
  `).get(date).c;
  
  res.render('scan-presences', {
    user: req.session,
    presences,
    date,
    presentToday,
    totalEmployees,
    absentToday: totalEmployees - presentToday
  });
});

app.post('/scan-password', (req, res) => {
  if (!req.session.scanUserId) return res.redirect('/scan-login');
  
  const { current_password, new_password, confirm_password } = req.body || {};
  
  if (new_password !== confirm_password) {
    return res.render('scan-password', { user: req.session, error: 'Les mots de passe ne correspondent pas', success: null });
  }
  
  const user = db.prepare('SELECT password_hash FROM users WHERE id = ?').get(req.session.scanUserId);
  if (!user || !bcrypt.compareSync(current_password, user.password_hash)) {
    return res.render('scan-password', { user: req.session, error: 'Mot de passe actuel incorrect', success: null });
  }
  
  const newHash = bcrypt.hashSync(new_password, 10);
  db.prepare('UPDATE users SET password_hash = ? WHERE id = ?').run(newHash, req.session.scanUserId);
  
  res.render('scan-password', { user: req.session, error: null, success: 'Mot de passe modifié avec succès!' });
});

// API: Today's presences
app.get('/api/today-presences', (req, res) => {
  const now = new Date();
  const today = now.toISOString().split('T')[0];
  const presences = db.prepare(`
    SELECT p.*, e.nom, e.prenom, e.poste, e.badge_id
    FROM presence p
    JOIN employees e ON p.employee_id = e.id
    WHERE date(p.scanned_at) = ?
    ORDER BY p.scanned_at DESC
  `).all(today);
  res.json(presences);
});

// API: All employees with presence status
app.get('/api/employees-status', (req, res) => {
  const now = new Date();
  const today = now.toISOString().split('T')[0];
  
  // Get all employees
  const employees = db.prepare(`
    SELECT e.*, 
      (SELECT type FROM presence WHERE employee_id = e.id AND date(scanned_at) = ? ORDER BY scanned_at DESC LIMIT 1) as last_status
    FROM employees e
    ORDER BY e.nom, e.prenom
  `).all(today);
  
  // Transform to show current status
  // If last_status is 'entrer' → Présent
  // If last_status is 'sortie' → Sortie
  // If last_status is null (no scan today) → Absent
  const result = employees.map(emp => {
    let status = 'absent';
    if (emp.last_status === 'entrer') {
      status = 'present';
    } else if (emp.last_status === 'sortie') {
      status = 'sortie';
    }
    
    return {
      id: emp.id,
      nom: emp.nom,
      prenom: emp.prenom,
      poste: emp.poste,
      badge_id: emp.badge_id,
      status: status,
      isPresent: status === 'present'
    };
  });
  
  res.json(result);
});

// SSE endpoint for admin scan interface (/scan)
app.get('/api/notifications/scan', (req, res) => {
  if (!req.session.scanUserId) {
    res.status(401).json({ error: 'Non autorisé' });
    return;
  }
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  
  const onNotification = (data) => {
    res.write(`data: ${JSON.stringify(data)}\n\n`);
  };
  
  notificationEmitter.on('pointage', onNotification);
  
  req.on('close', () => {
    notificationEmitter.off('pointage', onNotification);
  });
});

// SSE endpoint for user scanner interface (/scanner)
app.get('/api/notifications/scanner', (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  
  const onNotification = (data) => {
    res.write(`data: ${JSON.stringify(data)}\n\n`);
  };
  
  notificationEmitter.on('pointage', onNotification);
  
  req.on('close', () => {
    notificationEmitter.off('pointage', onNotification);
  });
});

// Legacy endpoint - redirect to scan notifications
app.get('/api/notifications', (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  
  const onNotification = (data) => {
    res.write(`data: ${JSON.stringify(data)}\n\n`);
  };
  
  notificationEmitter.on('pointage', onNotification);
  
  req.on('close', () => {
    notificationEmitter.off('pointage', onNotification);
  });
});

// API: Get stored notifications
app.get('/api/notifications/list', requireAuth, (req, res) => {
  res.json(notifications);
});

// API: Clear notifications
app.post('/api/notifications/clear', requireAuth, (req, res) => {
  notifications.length = 0;
  res.json({ ok: true });
});

app.get('/employe', async (req, res) => {
  if (!req.session.employeId) return res.redirect('/employe-login');
  
  const employe = db.prepare('SELECT * FROM employees WHERE id = ?').get(req.session.employeId);
  if (!employe) {
    req.session.employeId = null;
    return res.redirect('/employe-login');
  }
  
  const now = new Date();
  const today = now.toISOString().split('T')[0];
  const presences = db.prepare(`
    SELECT * FROM presence 
    WHERE employee_id = ? 
    AND date(scanned_at) = ?
    ORDER BY scanned_at DESC
  `).all(employe.id, today);
  
  const lastWeek = db.prepare(`
    SELECT date(scanned_at) as date, type, scanned_at
    FROM presence 
    WHERE employee_id = ?
    AND scanned_at >= datetime('now', '-7 days')
    ORDER BY scanned_at DESC
  `).all(employe.id);
  
  let qrDataUrl = null;
  try {
    qrDataUrl = await QRCode.toDataURL(String(employe.id_affichage || employe.id), { width: 250, margin: 2 });
  } catch (_) {}
  
  res.render('employe-dashboard', { 
    employe, 
    presences, 
    lastWeek, 
    qrDataUrl,
    today,
    hasPointedToday: presences.length > 0,
    lastPresence: presences[0] || null
  });
});

app.post('/employe-pointage', (req, res) => {
  if (!req.session.employeId) {
    return res.json({ ok: false, error: 'Non connecté' });
  }
  
  const employe = db.prepare('SELECT * FROM employees WHERE id = ?').get(req.session.employeId);
  if (!employe) {
    return res.json({ ok: false, error: 'Employé non trouvé' });
  }
  
  const now = new Date();
  const today = now.toISOString().split('T')[0];
  const lastPresence = db.prepare(`
    SELECT * FROM presence 
    WHERE employee_id = ? 
    AND date(scanned_at) = ?
    ORDER BY scanned_at DESC LIMIT 1
  `).get(employe.id, today);
  
  const type = (lastPresence && lastPresence.type === 'entrer') ? 'sortie' : 'entrer';
  
  const now2 = new Date();
  const localTime = now2.getFullYear() + '-' + 
    String(now2.getMonth() + 1).padStart(2, '0') + '-' + 
    String(now2.getDate()).padStart(2, '0') + ' ' + 
    String(now2.getHours()).padStart(2, '0') + ':' + 
    String(now2.getMinutes()).padStart(2, '0') + ':' + 
    String(now2.getSeconds()).padStart(2, '0');
  db.prepare('INSERT INTO presence (employee_id, type, scanned_at) VALUES (?, ?, ?)').run(employe.id, type, localTime);
  
  notificationEmitter.emit('pointage', {
    type: type,
    employe: { id: employe.id, nom: employe.nom, prenom: employe.prenom }
  });
  addNotification(type, { id: employe.id, nom: employe.nom, prenom: employe.prenom });
  
  res.json({ ok: true, type });
});

app.get('/employe-logout', (req, res) => {
  req.session.employeId = null;
  req.session.employeName = null;
  res.redirect('/employe-login');
});

// ---------- Badge exemple (route prioritaire, sans auth, toujours répond) ----------
app.get('/badge-exemple', async (req, res) => {
  const adresseSociete = 'Lot II T 104 A lavoloha, Antananarivo 102';
  const exemple = { badge_id: 'ARIS-0001', id_affichage: 1, nom: 'RAHARISON', prenom: 'Michaël', poste: 'GERANT', adresse: adresseSociete, email: null, photo: null };
  let qrDataUrl = null;
  try {
    qrDataUrl = await QRCode.toDataURL(String(exemple.id_affichage), { width: 300, margin: 2 });
  } catch (_) {}
  try {
    return res.render('badge-fiche', { user: req.session || {}, employee: exemple, qrDataUrl, hasBadgePdfTemplate: false, adresseSociete, isExemple: true });
  } catch (err) {
    console.error('badge-exemple render:', err);
    res.status(500).send('Erreur affichage badge. Vérifiez les vues.');
  }
});

app.use((req, res, next) => {
  const p = (req.path || '').replace(/%C3%A9/g, '\u00e9');
  if ((p === '/fiches-pr\u00e9sence' || (req.originalUrl && req.originalUrl.indexOf('fiches-pr') !== -1 && req.originalUrl.indexOf('fiches-presence') === -1))) {
    return res.redirect(302, '/fiches-presence');
  }
  next();
});

// ---------- Auth ----------
app.get('/login', (req, res) => {
  if (req.session.userId) return res.redirect('/');
  res.render('login', { error: null, useOtp: false });
});

app.get('/login-scanner', (req, res) => {
  if (req.session.userId) {
    const u = db.prepare('SELECT role FROM users WHERE id = ?').get(req.session.userId);
    return res.redirect((u && u.role === 'scanner') ? '/scanner' : '/');
  }
  res.render('login-scanner', { error: null });
});
app.post('/login-scanner', (req, res) => {
  const { username, password } = req.body || {};
  const user = db.prepare('SELECT * FROM users WHERE username = ?').get(username);
  if (!user || !bcrypt.compareSync(password, user.password_hash)) {
    return res.render('login-scanner', { error: 'Identifiants incorrects' });
  }
  req.session.userId = user.id;
  req.session.username = user.username;
  req.session.role = user.role;
  res.redirect('/scanner');
});

app.post('/login', (req, res) => {
  const { username, password } = req.body || {};
  const user = db.prepare('SELECT * FROM users WHERE username = ?').get(username);
  if (!user || !bcrypt.compareSync(password, user.password_hash)) {
    return res.render('login', { error: 'Identifiants incorrects', useOtp: false });
  }

  // Si on saisit un compte "scanner" dans le login admin, basculer automatiquement
  // vers l'interface scan (sans ouvrir une session admin).
  if (user.role === 'scanner') {
    req.session.scanUserId = user.id;
    req.session.scanUserName = user.username;
    // Nettoyer l'éventuelle session admin
    req.session.userId = null;
    req.session.username = null;
    req.session.role = null;
    return res.redirect('/scan');
  }

  req.session.userId = user.id;
  req.session.username = user.username;
  req.session.role = user.role;
  res.redirect('/');
});

function findUserByPhone(phone) {
  const clean = normalizePhone(phone);
  if (!clean) return null;
  const users = db.prepare('SELECT * FROM users WHERE telephone IS NOT NULL AND telephone != ""').all();
  return users.find(u => normalizePhone(u.telephone) === clean || normalizePhone(u.telephone).endsWith(clean.slice(-9)));
}

// Connexion par WhatsApp OTP — envoi au numéro configuré (sans afficher le numéro)
app.post('/api/login-request-otp', async (req, res) => {
  const user = db.prepare('SELECT * FROM users WHERE telephone IS NOT NULL AND telephone != "" AND (callmebot_apikey IS NOT NULL AND callmebot_apikey != "")').get();
  if (!user) return res.json({ ok: false, error: 'WhatsApp non configuré. Configurez numéro et clé API dans Paramètres.' });
  const apikey = user.callmebot_apikey || process.env.CALLMEBOT_APIKEY;
  if (!apikey) return res.json({ ok: false, error: 'Clé API CallMeBot manquante. Configurez dans Paramètres.' });
  const phone = user.telephone;
  const clean = normalizePhone(phone);
  const code = String(crypto.randomInt(100000, 999999));
  const expiresAt = new Date(Date.now() + 5 * 60 * 1000).toISOString();
  db.prepare('DELETE FROM otp_codes WHERE phone = ?').run(clean);
  db.prepare('INSERT INTO otp_codes (phone, code, expires_at) VALUES (?, ?, ?)').run(clean, code, expiresAt);
  try {
    await sendWhatsAppOTP(phone, code, apikey);
    res.json({ ok: true });
  } catch (err) {
    console.error('WhatsApp OTP:', err);
    res.json({ ok: false, error: 'Erreur d\'envoi WhatsApp. Vérifiez la clé API.' });
  }
});

app.post('/api/login-verify-otp', (req, res) => {
  const { code } = req.body || {};
  const codeStr = String(code || '').trim();
  if (!codeStr) return res.json({ ok: false, error: 'Code requis' });
  const row = db.prepare('SELECT * FROM otp_codes WHERE code = ? AND expires_at > datetime("now") ORDER BY id DESC LIMIT 1').get(codeStr);
  if (!row) return res.json({ ok: false, error: 'Code incorrect ou expiré' });
  const user = findUserByPhone(row.phone);
  if (!user) return res.json({ ok: false, error: 'Utilisateur non trouvé' });
  db.prepare('DELETE FROM otp_codes WHERE phone = ?').run(row.phone);

  if (user.role === 'scanner') {
    req.session.scanUserId = user.id;
    req.session.scanUserName = user.username;
    req.session.userId = null;
    req.session.username = null;
    req.session.role = null;
    return res.json({ ok: true, redirect: '/scan' });
  }

  req.session.userId = user.id;
  req.session.username = user.username;
  req.session.role = user.role;
  res.json({ ok: true, redirect: '/' });
});

app.get('/logout', (req, res) => {
  req.session.destroy();
  res.redirect('/login');
});

// ---------- Pages (protégées) ----------
app.get('/', requireAuth, (req, res) => {
  let empCount = db.prepare('SELECT COUNT(*) as c FROM employees').get().c;
  if (empCount === 0) {
    try { runSeedEmployees(); } catch (err) { console.error('Seed:', err); }
    empCount = db.prepare('SELECT COUNT(*) as c FROM employees').get().c;
  }
  const stats = {
    employees: db.prepare('SELECT COUNT(*) as c FROM employees').get().c,
    todayPresence: db.prepare(`
      SELECT COUNT(DISTINCT employee_id) as c FROM presence
      WHERE date(scanned_at) = date('now')
    `).get().c,
    lastPresences: db.prepare(`
      SELECT p.*, e.id as emp_id, e.id_affichage, e.badge_id, e.nom, e.prenom
      FROM presence p
      JOIN employees e ON e.id = p.employee_id
      ORDER BY p.scanned_at DESC
      LIMIT 10
    `).all()
  };
  const employees = db.prepare('SELECT id, id_affichage, badge_id, nom, prenom, poste FROM employees ORDER BY nom, prenom').all();
  res.render('dashboard', { user: req.session, stats, employees });
});

app.get('/parametres', requireAuth, (req, res) => {
  const user = db.prepare('SELECT id, username, role, telephone, callmebot_apikey FROM users WHERE id = ?').get(req.session.userId);
  if (!user) return res.redirect('/login');
  const u = { ...req.session, ...user, userId: user.id };
  res.render('parametres', { user: u });
});
app.post('/parametres', requireAuth, (req, res) => {
  const { telephone, callmebot_apikey } = req.body || {};
  const tel = (telephone || '').trim();
  const key = (callmebot_apikey || '').trim();
  if (tel) {
    if (key) {
      db.prepare('UPDATE users SET telephone = ?, callmebot_apikey = ? WHERE id = ?').run(tel, key, req.session.userId);
    } else {
      db.prepare('UPDATE users SET telephone = ? WHERE id = ?').run(tel, req.session.userId);
    }
  }
  res.redirect('/parametres?saved=1');
});

// Changer mot de passe
app.post('/parametres/password', requireAuth, (req, res) => {
  const { current_password, new_password, confirm_password } = req.body || {};
  
  if (new_password !== confirm_password) {
    const u = db.prepare('SELECT id, username, role, telephone, callmebot_apikey FROM users WHERE id = ?').get(req.session.userId);
    const userData = { ...req.session, ...u, userId: u.id };
    return res.render('parametres', { user: userData, pwderror: 'Les mots de passe ne correspondent pas' });
  }
  
  const user = db.prepare('SELECT password_hash FROM users WHERE id = ?').get(req.session.userId);
  if (!user || !bcrypt.compareSync(current_password, user.password_hash)) {
    const u = db.prepare('SELECT id, username, role, telephone, callmebot_apikey FROM users WHERE id = ?').get(req.session.userId);
    const userData = { ...req.session, ...u, userId: u.id };
    return res.render('parametres', { user: userData, pwderror: 'Mot de passe actuel incorrect' });
  }
  
  const newHash = bcrypt.hashSync(new_password, 10);
  db.prepare('UPDATE users SET password_hash = ? WHERE id = ?').run(newHash, req.session.userId);
  
  const u = db.prepare('SELECT id, username, role, telephone, callmebot_apikey FROM users WHERE id = ?').get(req.session.userId);
  const userData = { ...req.session, ...u, userId: u.id };
  res.render('parametres', { user: userData, pwdsaved: true });
});
// ---------- Utilisateurs (admin uniquement) ----------
app.get('/utilisateurs', requireAdmin, (req, res) => {
  const users = db.prepare('SELECT id, username, role, created_at FROM users ORDER BY username').all();
  res.render('utilisateurs', { user: req.session, users });
});
app.get('/utilisateurs/nouveau', requireAdmin, (req, res) => {
  res.render('utilisateur-form', { user: req.session, editUser: null });
});
app.get('/utilisateurs/:id/modifier', requireAdmin, (req, res) => {
  const u = db.prepare('SELECT id, username, role FROM users WHERE id = ?').get(req.params.id);
  if (!u) return res.redirect('/utilisateurs');
  res.render('utilisateur-form', { user: req.session, editUser: u });
});
app.post('/utilisateurs', requireAdmin, (req, res) => {
  const { username, password, role } = req.body || {};
  const u = (username || '').trim();
  const p = (password || '').trim();
  if (!u || !p) return res.redirect('/utilisateurs?error=1');
  const hash = bcrypt.hashSync(p, 10);
  try {
    db.prepare('INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)').run(u, hash, role || 'user');
    res.redirect('/utilisateurs');
  } catch (e) {
    res.redirect('/utilisateurs?error=dup');
  }
});
app.post('/utilisateurs/:id', requireAdmin, (req, res) => {
  const { password, role } = req.body || {};
  const p = (password || '').trim();
  if (p) {
    const hash = bcrypt.hashSync(p, 10);
    db.prepare('UPDATE users SET password_hash = ?, role = ? WHERE id = ?').run(hash, role || 'user', req.params.id);
  } else {
    db.prepare('UPDATE users SET role = ? WHERE id = ?').run(role || 'user', req.params.id);
  }
  res.redirect('/utilisateurs');
});
app.post('/utilisateurs/:id/supprimer', requireAdmin, (req, res) => {
  db.prepare('DELETE FROM users WHERE id = ? AND username != ?').run(req.params.id, 'admin');
  res.redirect('/utilisateurs');
});
app.get('/employes', requireAuth, (req, res) => {
  const employees = db.prepare('SELECT * FROM employees ORDER BY nom, prenom').all();
  res.render('employes', { user: req.session, employees, imported: req.query.imported, import_error: req.query.import_error, error: req.query.error });
});

app.get('/employes/nouveau', requireAuth, (req, res) => {
  res.render('employe-form', { user: req.session, employee: null });
});

app.get('/employes/:id/modifier', requireAuth, (req, res) => {
  const employee = db.prepare('SELECT * FROM employees WHERE id = ?').get(req.params.id);
  if (!employee) return res.redirect('/employes');
  res.render('employe-form', { user: req.session, employee });
});

app.post('/employes', requireAuth, (req, res) => {
  const { nom, prenom, poste, departement, email, adresse, telephone } = req.body;
  const badge_id = 'ARIS-' + Date.now().toString(36).toUpperCase();
  try {
    db.prepare(`
      INSERT INTO employees (badge_id, nom, prenom, poste, departement, email, adresse, telephone)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `).run(badge_id, nom, prenom, poste || null, departement || null, email || null, adresse || null, telephone || null);
    res.redirect('/employes');
  } catch (e) {
    res.redirect('/employes?error=1');
  }
});

app.post('/employes/:id', requireAuth, (req, res) => {
  const { nom, prenom, poste, departement, email, adresse, telephone } = req.body;
  db.prepare(`
    UPDATE employees SET nom=?, prenom=?, poste=?, departement=?, email=?, adresse=?, telephone=? WHERE id=?
  `).run(nom, prenom, poste || null, departement || null, email || null, adresse || null, telephone || null, req.params.id);
  res.redirect('/employes');
});

app.post('/employes/:id/supprimer', requireAuth, (req, res) => {
  db.prepare('DELETE FROM presence WHERE employee_id = ?').run(req.params.id);
  db.prepare('DELETE FROM employees WHERE id = ?').run(req.params.id);
  res.redirect('/employes');
});

// ---------- Badges & QR ----------
app.get('/badges', requireAuth, (req, res) => {
  const employees = db.prepare('SELECT * FROM employees ORDER BY nom, prenom').all();
  res.render('badges', { user: req.session, employees });
});
const SEED_EMPLOYEES = [
  { id_affichage: 1, nom: 'RAHARISON', prenom: 'Michaël', equipe: 'Gérant', email: 'michael@aris-cc.com', mdp_mail: 'michael171102!', date_naissance: '10/11/1994', date_embauche: '5/12/2023', poste: 'GERANT', categorie: 'HC', adresse: 'Lot II B 128 TER Mahalavolona Andoharanofotsy 102', cin: '101.211.216.824', num_cnaps: '941110005606', telephone: '038 53 405 34' },
  { id_affichage: 2, nom: 'RASOANIRINA', prenom: 'Arlette', equipe: 'Agent de Sécurité', email: null, mdp_mail: null, date_naissance: '21/04/1977', date_embauche: '5/12/2023', poste: 'Sécurité', categorie: '2B', adresse: 'FA 243 TER Ambohimanatrika Mivoatra commune Tanjombato Antananarivo 102', cin: '210.012.012.871', num_cnaps: '772421000797', telephone: '034 75 819 13' },
  { id_affichage: 3, nom: 'RANAIVOARIMANANA', prenom: 'Ravakinionja Jean Valérie', equipe: 'Ingénieur BTP', email: 'onja@aris-cc.com', mdp_mail: 'onjabtp171102!', date_naissance: '08/02/1995', date_embauche: '2/1/2024', poste: 'Ingénieur BTP', categorie: 'HC', adresse: 'Lot TSF 505/A Antsahafohy Ambohitrimanjaka', cin: '103.131.015.114', num_cnaps: '950208004812', telephone: '033 05 059 33' },
  { id_affichage: 4, nom: 'RAZAFINDRAIBE', prenom: 'Harimalala Vololoniaina Annie', equipe: 'Ingénieur BTP', email: 'annie@aris-cc.com', mdp_mail: 'anniebtp171102!', date_naissance: '24/09/1996', date_embauche: '1/2/2024', poste: 'Ingénieur BTP', categorie: 'HC', adresse: 'III H 105 B BIS Avaratanana Antananarivo VI', cin: '101.982.094.987', num_cnaps: '962924004850', telephone: '034 25 903 79' },
  { id_affichage: 5, nom: 'RAHANTARIMALALA', prenom: 'Mamisoa Felicia', equipe: 'Assistant Technique', email: 'mamisoa@aris-cc.com', mdp_mail: 'mamisoa171102!', date_naissance: '02/10/1993', date_embauche: '12/4/2024', poste: 'TECHNICIEN ASSISTANT', categorie: '2B', adresse: 'III F 138 Antohomadinika Afovoany Antananarivo I', cin: '101.211.214.901', num_cnaps: '931002005967', telephone: '034 02 213 54' },
  { id_affichage: 7, nom: 'ANDRIANARISOA', prenom: 'Lalarimina Tahiry', equipe: 'Manager Google Maps', email: 'tahiry@aris-cc.com', mdp_mail: 'a1^9IM]HR&9U', date_naissance: '16/09/1995', date_embauche: '19/08/2024', poste: 'MANAGER CALL', categorie: 'HC', adresse: 'IC 189 TER D ANKADILALAMPOTSY ANKARAOBATO', cin: '101.252.184.456', num_cnaps: '952916002009', telephone: '032 52 771 41' },
  { id_affichage: 8, nom: 'FANOMEZANTSOA', prenom: 'Maminiaina Sarobidy', equipe: 'Technicien réseau', email: null, mdp_mail: null, date_naissance: '12/12/2002', date_embauche: '01/07/2025', poste: 'TECHNICIEN RESEAU', categorie: '2B', adresse: 'LOT IC 110 TER A ANKADILALAMPOTSY ANKARAOBATO', cin: '117.191.018.397', num_cnaps: '021212002606', telephone: '033 34 755 64' },
  { id_affichage: 10, nom: 'RASOAMBOLAMANANA', prenom: 'Aimée Eliane', equipe: 'Femme de ménage', email: null, mdp_mail: null, date_naissance: '08/06/1993', date_embauche: '25/09/2024', poste: 'FEMME DE MENAGE', categorie: '2B', adresse: 'II A 299 BIS K Tanjombato Iraitsimivaky Antananarivo 102', cin: '117.152.016.626', num_cnaps: '932608003092', telephone: '034 30 933 55' },
  { id_affichage: 12, nom: 'RAZANATSIMBA', prenom: 'Brigitte', equipe: 'Ebay', email: 'brigitterazanatsimba@aris-cc.com', mdp_mail: 'concept_rigi', date_naissance: '08/08/1980', date_embauche: '01/05/2025', poste: 'TELEOPERATEUR', categorie: '2B', adresse: 'II T 29 Ambohibao Iavoloha Bongatsara', cin: '117.392.002.118', num_cnaps: '802808003871', telephone: '034 95 432 10' },
];
function runSeedEmployees() {
  SEED_EMPLOYEES.forEach((e) => {
    const badge_id = 'ARIS-' + String(e.id_affichage).padStart(4, '0');
    const existing = db.prepare('SELECT id FROM employees WHERE badge_id = ?').get(badge_id);
    if (existing) {
      db.prepare('UPDATE employees SET id_affichage=?, nom=?, prenom=?, poste=?, departement=?, email=?, adresse=?, telephone=?, equipe=?, date_naissance=?, date_embauche=?, categorie=?, cin=?, num_cnaps=?, mdp_mail=? WHERE badge_id=?').run(e.id_affichage, e.nom, e.prenom, e.poste || null, e.equipe || null, e.email || null, e.adresse || null, e.telephone || null, e.equipe || null, e.date_naissance || null, e.date_embauche || null, e.categorie || null, e.cin || null, e.num_cnaps || null, e.mdp_mail || null, badge_id);
    } else {
      db.prepare('INSERT INTO employees (badge_id, id_affichage, nom, prenom, poste, departement, email, adresse, telephone, equipe, date_naissance, date_embauche, categorie, cin, num_cnaps, mdp_mail) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)').run(badge_id, e.id_affichage, e.nom, e.prenom, e.poste || null, e.equipe || null, e.email || null, e.adresse || null, e.telephone || null, e.equipe || null, e.date_naissance || null, e.date_embauche || null, e.categorie || null, e.cin || null, e.num_cnaps || null, e.mdp_mail || null);
    }
  });
}
function seedEmployees(req, res) {
  try {
    runSeedEmployees();
  } catch (err) {
    console.error('Seed employees:', err);
    return res.redirect('/?seed_error=1');
  }
  res.redirect('/employes');
}
// Seed employés (liste type) — accessible sans auth pour éviter "Impossible d'obtenir"
app.get('/api/seed-employees', seedEmployees);
app.get('/api/seed-eployes', seedEmployees);
app.post('/api/seed-employees', seedEmployees);
app.post('/api/seed-eployes', seedEmployees);

// Import employés depuis fichier Excel (.xlsx)
function mapExcelRow(row) {
  const get = (keys) => {
    for (const k of keys) {
      const v = row[k];
      if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
    }
    return null;
  };
  const id = get(['ID', 'id']);
  if (!id) return null;
  const idNum = parseInt(id, 10) || id;
  return {
    id_affichage: isNaN(idNum) ? null : idNum,
    nom: get(['NOM', 'Nom', 'nom']) || '',
    prenom: get(['PRENOM', 'Prénom', 'Prenom', 'prenom']) || '',
    poste: get(['FONCTION', 'Poste', 'poste', 'Fonction']),
    equipe: get(['Equipe', 'equipe', 'EQUIPE', 'Département', 'departement']),
    email: get(['MAIL', 'Mail', 'email', 'Email']),
    adresse: get(['HABITATION', 'Adresse', 'adresse', 'Habitation']),
    telephone: get(['TELEPHONE', 'Téléphone', 'telephone']),
    date_naissance: get(['DATE DE NAISSANCE', 'Date de naissance', 'date_naissance']),
    date_embauche: get(['DATE D EMBAUCHE', 'Date d\'embauche', 'date_embauche']),
    categorie: get(['CATEGORIE', 'Catégorie', 'categorie']),
    cin: get(['CIN', 'Cin']),
    num_cnaps: get(['NUM CNAPS', 'Num CNAPS', 'num_cnaps']),
    mdp_mail: get(['MDP MAIL', 'Mdp mail', 'mdp_mail'])
  };
}
app.post('/api/import-employees', requireAuth, upload.single('file'), (req, res) => {
  if (!req.file || !req.file.buffer) {
    return res.redirect('/employes?import_error=no_file');
  }
  try {
    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const firstSheet = wb.SheetNames[0];
    const ws = wb.Sheets[firstSheet];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    let imported = 0;
    for (let i = 0; i < rows.length; i++) {
      const e = mapExcelRow(rows[i]);
      if (!e || !e.nom) continue;
      const num = (e.id_affichage != null && !isNaN(e.id_affichage)) ? e.id_affichage : (i + 1);
      const badge_id = 'ARIS-' + String(num).padStart(4, '0');
      const existing = db.prepare('SELECT id FROM employees WHERE badge_id = ?').get(badge_id);
      if (existing) {
        db.prepare('UPDATE employees SET id_affichage=?, nom=?, prenom=?, poste=?, departement=?, email=?, adresse=?, telephone=?, equipe=?, date_naissance=?, date_embauche=?, categorie=?, cin=?, num_cnaps=?, mdp_mail=? WHERE badge_id=?').run(e.id_affichage, e.nom, e.prenom, e.poste || null, e.equipe || null, e.email || null, e.adresse || null, e.telephone || null, e.equipe || null, e.date_naissance || null, e.date_embauche || null, e.categorie || null, e.cin || null, e.num_cnaps || null, e.mdp_mail || null, badge_id);
      } else {
        db.prepare('INSERT INTO employees (badge_id, id_affichage, nom, prenom, poste, departement, email, adresse, telephone, equipe, date_naissance, date_embauche, categorie, cin, num_cnaps, mdp_mail) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)').run(badge_id, e.id_affichage, e.nom, e.prenom, e.poste || null, e.equipe || null, e.email || null, e.adresse || null, e.telephone || null, e.equipe || null, e.date_naissance || null, e.date_embauche || null, e.categorie || null, e.cin || null, e.num_cnaps || null, e.mdp_mail || null);
      }
      imported++;
    }
    res.redirect('/employes?imported=' + imported);
  } catch (err) {
    console.error('Import Excel:', err);
    res.redirect('/employes?import_error=1');
  }
});

app.get('/badge/:badgeId/qr', requireAuth, async (req, res) => {
  try {
    const url = await QRCode.toDataURL(req.params.badgeId, { width: 300, margin: 2 });
    res.json({ qr: url });
  } catch (e) {
    res.status(400).json({ error: 'Badge invalide' });
  }
});

app.get('/api/employee/:id/qr-id', requireAuth, async (req, res) => {
  const employee = db.prepare('SELECT id, id_affichage, badge_id FROM employees WHERE id = ?').get(req.params.id);
  if (!employee) return res.status(404).json({ error: 'Employé non trouvé' });
  const qrContent = employee.id_affichage != null ? String(employee.id_affichage) : employee.badge_id;
  try {
    const url = await QRCode.toDataURL(qrContent, { width: 120, margin: 1 });
    res.json({ qr: url, id: qrContent });
  } catch (e) {
    res.status(500).json({ error: 'Erreur QR' });
  }
});

app.get('/badge/:badgeId/fiche', requireAuth, async (req, res) => {
  const employee = db.prepare('SELECT * FROM employees WHERE badge_id = ?').get(req.params.badgeId);
  if (!employee) return res.status(404).send('Employé non trouvé');
  const qrContent = employee.id_affichage != null ? String(employee.id_affichage) : employee.badge_id;
  let qrDataUrl = null;
  try {
    qrDataUrl = await QRCode.toDataURL(qrContent, { width: 300, margin: 2 });
  } catch (_) {}
  const hasBadgePdfTemplate = fs.existsSync(path.join(__dirname, 'Badge.pdf'));
  const adresseSociete = 'Lot II T 104 A lavoloha, Antananarivo 102';
  res.render('badge-fiche', { user: req.session, employee, qrDataUrl, hasBadgePdfTemplate, adresseSociete });
});

// Génération du badge au format PDF (nouveau design)
app.get('/badge/:badgeId/badge.pdf', requireAuth, async (req, res) => {
  const employee = db.prepare('SELECT * FROM employees WHERE badge_id = ?').get(req.params.badgeId);
  if (!employee) return res.status(404).send('Employé non trouvé');
  try {
    const adresseSociete = 'Lot II T 104 A lavoloha, Antananarivo 102';
    const pdfBytes = await generateBadgePdf(employee, 'logo.png', { adresseSociete });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="badge-${employee.badge_id}.pdf"`);
    res.send(Buffer.from(pdfBytes));
  } catch (err) {
    console.error(err);
    res.status(500).send('Erreur lors de la génération du badge PDF.');
  }
});

// ---------- Scanner (page publique pour enregistrer présence) ----------
app.get('/scanner', (req, res) => {
  res.render('scanner', { user: req.session });
});

app.post('/api/presence', (req, res) => {
  const { badge_id, id } = req.body || {};
  const raw = (badge_id != null ? badge_id : id);
  if (raw === undefined || raw === null || raw === '') {
    return res.status(400).json({ error: 'badge_id ou id requis' });
  }
  const str = String(raw).trim();
  let employee = null;
  const num = parseInt(str, 10);
  if (!isNaN(num)) {
    employee = db.prepare('SELECT * FROM employees WHERE id_affichage = ?').get(num);
  }
  if (!employee) {
    employee = db.prepare('SELECT * FROM employees WHERE badge_id = ?').get(str);
  }
  if (!employee) {
    return res.status(404).json({ error: 'Badge / ID non reconnu' });
  }
  const last = db.prepare(`
    SELECT type FROM presence WHERE employee_id = ? ORDER BY scanned_at DESC LIMIT 1
  `).get(employee.id);
  const nextType = (!last || last.type === 'sortie') ? 'entrer' : 'sortie';
  const now3 = new Date();
  const localTime = now3.getFullYear() + '-' + 
    String(now3.getMonth() + 1).padStart(2, '0') + '-' + 
    String(now3.getDate()).padStart(2, '0') + ' ' + 
    String(now3.getHours()).padStart(2, '0') + ':' + 
    String(now3.getMinutes()).padStart(2, '0') + ':' + 
    String(now3.getSeconds()).padStart(2, '0');
  db.prepare('INSERT INTO presence (employee_id, type, scanned_at) VALUES (?, ?, ?)').run(employee.id, nextType, localTime);
  
  notificationEmitter.emit('pointage', {
    type: nextType,
    employe: { id: employee.id, nom: employee.nom, prenom: employee.prenom }
  });
  
  res.json({
    ok: true,
    type: nextType,
    employee: { 
      id: employee.id,
      nom: employee.nom, 
      prenom: employee.prenom, 
      badge_id: employee.badge_id,
      poste: employee.poste,
      photo: employee.photo
    }
  });
});

// ---------- Présences par date ----------
app.get('/presences-par-date', requireAuth, (req, res) => {
  const now = new Date();
  const localOffset = now.getTimezoneOffset() * 60000;
  const localDate = new Date(now.getTime() - localOffset);
  const date = req.query.date || localDate.toISOString().split('T')[0];
  const filter = req.query.filter; // 'presents', 'absents', or 'sorties'
  
  const totalEmployees = db.prepare('SELECT COUNT(*) as c FROM employees').get().c;
  
  // Get employees who have entered (present)
  const presentEmployees = db.prepare(`
    SELECT DISTINCT employee_id FROM presence
    WHERE date(scanned_at) = ? AND type = 'entrer'
  `).all(date).map(p => p.employee_id);
  
  // Get employees who have both entered and left (sorted)
  const sortedEmployees = db.prepare(`
    SELECT DISTINCT employee_id FROM presence p1
    WHERE date(p1.scanned_at) = ?
    AND p1.type = 'sortie'
    AND EXISTS (
      SELECT 1 FROM presence p2 
      WHERE p2.employee_id = p1.employee_id 
      AND date(p2.scanned_at) = ? 
      AND p2.type = 'entrer'
      AND p2.scanned_at < p1.scanned_at
    )
  `).all(date, date).map(p => p.employee_id);
  
  const presentCount = presentEmployees.length;
  const sortieCount = sortedEmployees.length;
  const absentCount = totalEmployees - presentCount;
  
  let presences = [];
  let pageTitle = 'Pointages';
  let showPresenceList = false;
  
  if (filter === 'presents') {
    // Show list of present employees (those who entered but not left)
    const presentEmps = db.prepare(`
      SELECT e.*, p.scanned_at as last_scan, 'entrer' as last_type
      FROM employees e
      JOIN presence p ON p.employee_id = e.id
      WHERE date(p.scanned_at) = ? AND p.type = 'entrer'
      GROUP BY e.id
      ORDER BY p.scanned_at DESC
    `).all(date);
    
    presences = presentEmps.map(e => ({
      employee_id: e.id,
      nom: e.nom,
      prenom: e.prenom,
      poste: e.poste,
      badge_id: e.badge_id,
      type: 'entrer',
      scanned_at: e.last_scan,
      is_present_list: true
    }));
    pageTitle = 'Employés présents';
    showPresenceList = true;
  } else if (filter === 'sorties') {
    // Show list of sorted employees (those who entered and left)
    const sortedEmps = db.prepare(`
      SELECT e.*, MAX(p.scanned_at) as last_scan
      FROM employees e
      JOIN presence p ON p.employee_id = e.id
      WHERE date(p.scanned_at) = ? AND p.type = 'sortie'
      AND EXISTS (
        SELECT 1 FROM presence p2 
        WHERE p2.employee_id = e.id 
        AND date(p2.scanned_at) = ? 
        AND p2.type = 'entrer'
        AND p2.scanned_at < p.scanned_at
      )
      GROUP BY e.id
      ORDER BY last_scan DESC
    `).all(date, date);
    
    presences = sortedEmps.map(e => ({
      employee_id: e.id,
      nom: e.nom,
      prenom: e.prenom,
      poste: e.poste,
      badge_id: e.badge_id,
      type: 'sortie',
      scanned_at: e.last_scan,
      is_sorted_list: true
    }));
    pageTitle = 'Employés partis';
    showPresenceList = true;
  } else if (filter === 'absents') {
    // Show list of absent employees (those who didn't enter)
    const absentEmps = db.prepare(`
      SELECT * FROM employees e
      WHERE e.id NOT IN (SELECT DISTINCT employee_id FROM presence WHERE date(scanned_at) = ? AND type = 'entrer')
      ORDER BY e.nom, e.prenom
    `).all(date);
    
    presences = absentEmps.map(e => ({
      employee_id: e.id,
      nom: e.nom,
      prenom: e.prenom,
      poste: e.poste,
      badge_id: e.badge_id,
      type: null,
      scanned_at: null,
      is_absent_list: true
    }));
    pageTitle = 'Employés absents';
    showPresenceList = true;
  } else {
    // Show all presences (pointages)
    presences = db.prepare(`
      SELECT p.*, e.badge_id, e.nom, e.prenom, e.poste
      FROM presence p
      JOIN employees e ON e.id = p.employee_id
      WHERE date(p.scanned_at) = ?
      ORDER BY p.scanned_at DESC
    `).all(date);
  }
  
  res.render('presences-par-date', {
    user: req.session,
    presences,
    date,
    presentToday: presentCount,
    sortieToday: sortieCount,
    totalEmployees,
    absentToday: absentCount,
    filter,
    pageTitle,
    showPresenceList
  });
});

// ---------- Présences (historique) ----------
app.get('/presences', requireAuth, (req, res) => {
  const page = Math.max(1, parseInt(req.query.page) || 1);
  const perPage = 50;
  const now = new Date();
  const localOffset = now.getTimezoneOffset() * 60000;
  const localDate = new Date(now.getTime() - localOffset);
  const today = localDate.toISOString().split('T')[0];
  const dateFilter = req.query.date || today;
  
  let whereClause = '';
  let params = [];
  
  if (dateFilter) {
    whereClause = "WHERE DATE(p.scanned_at) = ?";
    params.push(dateFilter);
  }
  
  const total = db.prepare('SELECT COUNT(*) as c FROM presence p ' + whereClause).get(...params).c;
  
  let query = `
    SELECT p.*, e.badge_id, e.nom, e.prenom
    FROM presence p
    JOIN employees e ON e.id = p.employee_id
    ${whereClause}
    ORDER BY e.badge_id ASC, p.scanned_at ASC
  `;
  
  if (dateFilter) {
    params.push(perPage, (page - 1) * perPage);
    query += ' LIMIT ? OFFSET ?';
  } else {
    query += ' LIMIT ? OFFSET ?';
    params.push(perPage, (page - 1) * perPage);
  }
  
  const presences = db.prepare(query).all(...params);
  res.render('presences', {
    user: req.session,
    presences,
    page,
    totalPages: Math.ceil(total / perPage),
    total,
    dateFilter,
    today
  });
});

app.get('/api/presences/export', requireAuth, (req, res) => {
  const now = new Date();
  const localOffset = now.getTimezoneOffset() * 60000;
  const localDate = new Date(now.getTime() - localOffset);
  const today = localDate.toISOString().split('T')[0];
  const dateFilter = req.query.date || today;
  
  let whereClause = '';
  let params = [];
  
  if (dateFilter) {
    whereClause = "WHERE DATE(p.scanned_at) = ?";
    params.push(dateFilter);
  }
  
  const presences = db.prepare(`
    SELECT p.scanned_at, p.type, e.badge_id, e.nom, e.prenom
    FROM presence p
    JOIN employees e ON e.id = p.employee_id
    ${whereClause}
    ORDER BY e.badge_id ASC, p.scanned_at ASC
  `).all(...params);
  res.setHeader('Content-Type', 'application/json');
  const filename = dateFilter ? `presences-${dateFilter}.json` : 'presences.json';
  res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
  res.send(JSON.stringify(presences, null, 2));
});

// ---------- Fiches de présence (imprimables, auto-sauvegarde) ----------
const MOIS = ['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'];
app.get('/fiches-presence', requireAuth, (req, res) => {
  const fiches = db.prepare('SELECT * FROM fiche_presence ORDER BY annee DESC, mois DESC').all();
  res.render('fiches-presence-list', { user: req.session, fiches, MOIS, error: req.query.error });
});
app.get('/fiches-presence/nouvelle', requireAuth, (req, res) => {
  const d = new Date();
  res.render('fiches-presence-new', { user: req.session, mois: d.getMonth() + 1, annee: d.getFullYear(), MOIS });
});
app.post('/fiches-presence', requireAuth, (req, res) => {
  const { titre, mois, annee } = req.body || {};
  const m = parseInt(mois, 10); const a = parseInt(annee, 10);
  if (!titre || !m || !a) return res.redirect('/fiches-presence?error=1');
  const titreNorm = titre.trim() || `${MOIS[m - 1]} ${a}`;
  db.prepare('INSERT INTO fiche_presence (titre, mois, annee, donnees) VALUES (?, ?, ?, ?)').run(titreNorm, m, a, '{}');
  const row = db.prepare('SELECT id FROM fiche_presence ORDER BY id DESC LIMIT 1').get();
  res.redirect('/fiches-presence/' + row.id);
});
app.get('/fiches-presence/:id', requireAuth, (req, res) => {
  const fiche = db.prepare('SELECT * FROM fiche_presence WHERE id = ?').get(req.params.id);
  if (!fiche) return res.redirect('/fiches-presence');
  const employees = db.prepare('SELECT id, badge_id, nom, prenom, poste FROM employees ORDER BY nom, prenom').all();
  const jours = new Date(fiche.annee, fiche.mois, 0).getDate();
  const donnees = (fiche.donnees && fiche.donnees !== '{}') ? JSON.parse(fiche.donnees) : {};
  res.render('fiche-presence', {
    user: req.session,
    fiche,
    employees,
    jours: Array.from({ length: jours }, (_, i) => i + 1),
    donnees,
    MOIS
  });
});
app.patch('/api/fiches-presence/:id', requireAuth, (req, res) => {
  const fiche = db.prepare('SELECT id FROM fiche_presence WHERE id = ?').get(req.params.id);
  if (!fiche) return res.status(404).json({ error: 'Fiche non trouvée' });
  const { donnees } = req.body || {};
  const json = typeof donnees === 'string' ? donnees : JSON.stringify(donnees || {});
  db.prepare('UPDATE fiche_presence SET donnees = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?').run(json, req.params.id);
  res.json({ ok: true });
});

// ---------- Fiche de présence manuelle (imprimable, employés Aris Concept, jours ouvrables) ----------
function getJoursOuvrables(mois, annee) {
  const jours = [];
  const n = new Date(annee, mois, 0).getDate();
  for (let j = 1; j <= n; j++) {
    const d = new Date(annee, mois - 1, j);
    if (d.getDay() >= 1 && d.getDay() <= 5) jours.push({ jour: j });
  }
  return jours;
}
app.get('/fiche-presence-manuelle', requireAuth, (req, res) => {
  const d = new Date();
  let mois = parseInt(req.query.mois, 10) || (d.getMonth() + 1);
  let annee = parseInt(req.query.annee, 10) || d.getFullYear();
  mois = Math.max(1, Math.min(12, mois || 1));
  annee = Math.max(2020, Math.min(2030, annee || d.getFullYear()));
  const employees = db.prepare(`
    SELECT id, id_affichage, badge_id, nom, prenom
    FROM employees
    WHERE badge_id LIKE 'ARIS-%'
    ORDER BY id_affichage, nom, prenom
  `).all();
  const joursOuvrables = getJoursOuvrables(mois, annee);
  const MOIS = ['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'];
  const moisLabel = MOIS[mois - 1] + ' ' + annee;
  res.render('fiche-presence-manuelle', {
    user: req.session,
    employees,
    joursOuvrables,
    mois,
    annee,
    moisLabel
  });
});

// ---------- 404 ----------
app.use((req, res) => {
  res.status(404).send(`
    <!DOCTYPE html>
    <html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
    <title>404 - PresenceAris</title>
    <style>body{font-family:system-ui,sans-serif;background:#0d0d0d;color:#e8e8e8;margin:0;min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:2rem;} a{color:#4da6ff;} a:hover{text-decoration:underline;} .links{margin-top:1rem;}</style>
    </head><body>
      <h1>404 — Page non trouvée</h1>
      <p>L'URL demandée n'existe pas.</p>
      <div class="links"><a href="/login">Connexion</a> &middot; <a href="/">Tableau de bord</a> &middot; <a href="/badge-exemple">Exemple badge</a></div>
    </body></html>
  `);
});

// ---------- Démarrage ----------
const httpsOptions = {
  key: fs.readFileSync(path.join(__dirname, 'key.pem')),
  cert: fs.readFileSync(path.join(__dirname, 'cert.pem'))
};

// HTTPS uniquement
https.createServer(httpsOptions, app).listen(PORT, () => {
  console.log('PresenceAris démarré sur https://localhost:' + PORT);
  console.log('Sur le réseau: https://192.168.4.250:' + PORT);
});
