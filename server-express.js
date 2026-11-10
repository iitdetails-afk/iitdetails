const express = require('express');
const fs = require('fs');
const path = require('path');
const https = require('https');
const crypto = require('crypto');
const os = require('os');
const excelDb = require('./excel-db');
const GoogleSheetsSync = require('./google-sheets-sync');

const app = express();
const port = process.env.PORT || 5000;
// On Vercel, only /tmp is writable. Use it for data files in production.
const isVercel = !!process.env.VERCEL;
const dataDir = isVercel ? '/tmp/data' : path.join(__dirname, 'data');
const submissionsFile = path.join(dataDir, 'submissions.json');
const detailsFile = path.join(dataDir, 'details.json');
const adminConfigFile = path.join(dataDir, 'admin_config.json');
const staffFile = path.join(dataDir, 'staff.json');
const leavesFile = path.join(dataDir, 'leaves.json');
const tasksFile = path.join(dataDir, 'tasks.json');
const notificationsFile = path.join(dataDir, 'notifications.json');
const addressesFile = path.join(dataDir, 'addresses.json');
const excelDatabaseFile = path.join(dataDir, 'database.xlsx');
const excelDatabasePublicFile = path.join(__dirname, 'database.xlsx');
const googleSheetsConfigFile = path.join(dataDir, 'google_sheets_config.json');
const backendPortFile = path.join(__dirname, '.backend-port');

const googleSheetsSync = new GoogleSheetsSync(googleSheetsConfigFile);

// ── Single admin session (in-memory) ──────────────────────────────────
let adminSession = { token: null, loginAt: null };

function generateToken() {
  return crypto.randomBytes(32).toString('hex');
}

const TOKEN_SECRET = 'iit_super_secret_token_key_2026';

function hashPassword(password, salt) {
  return crypto.pbkdf2Sync(password, salt, 1000, 64, 'sha512').toString('hex');
}

function generateSessionToken(payload) {
  const header = Buffer.from(JSON.stringify({ alg: 'HS256', typ: 'JWT' })).toString('base64').replace(/=/g, '');
  const body = Buffer.from(JSON.stringify(payload)).toString('base64').replace(/=/g, '');
  const signature = crypto.createHmac('sha256', TOKEN_SECRET).update(`${header}.${body}`).digest('base64').replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
  return `${header}.${body}.${signature}`;
}

function verifySessionToken(token) {
  if (!token) return null;
  const parts = token.split('.');
  if (parts.length !== 3) return null;
  const [header, body, signature] = parts;
  const expectedSig = crypto.createHmac('sha256', TOKEN_SECRET).update(`${header}.${body}`).digest('base64').replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
  if (signature !== expectedSig) return null;
  try {
    return JSON.parse(Buffer.from(body, 'base64').toString('utf8'));
  } catch {
    return null;
  }
}

function getAdminConfig() {
  return readJson(adminConfigFile, { pin: '1234', officeStartTime: '10:00', officeEndTime: '18:30' });
}

function getAdminPin() {
  const cfg = getAdminConfig();
  return cfg.pin || '1234';
}

function isValidToken(token) {
  if (!token) return false;
  const payload = verifySessionToken(token);
  return payload && payload.role === 'admin';
}

function getAuthToken(req) {
  const auth = req.headers['authorization'] || '';
  if (auth.startsWith('Bearer ')) return auth.slice(7).trim();
  return '';
}

function requireAdmin(req, res, next) {
  const token = getAuthToken(req);
  if (!isValidToken(token)) {
    return res.status(401).json({ success: false, error: 'Unauthorized — admin login required' });
  }
  next();
}

function requireStaff(req, res, next) {
  const token = getAuthToken(req);
  const payload = verifySessionToken(token);
  if (!payload || payload.role !== 'staff') {
    return res.status(401).json({ success: false, error: 'Unauthorized — staff login required' });
  }
  req.staffId = payload.staffId;
  next();
}

app.use(express.json({ limit: '50mb' }));
app.use(express.static(__dirname));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/details', (req, res) => {
  res.redirect('/details.html');
});

app.get('/form', (req, res) => {
  res.redirect('/form.html');
});

app.get('/ss', (req, res) => {
  res.redirect('/ss.html');
});

function syncDatabaseToExcel() {
  try {
    const submissions = readJson(submissionsFile, []);
    const staff = readJson(staffFile, []);
    const addresses = readJson(addressesFile, []);
    excelDb.saveToExcelFile(excelDatabaseFile, submissions, staff, addresses);
    if (excelDatabaseFile !== excelDatabasePublicFile) {
      try { excelDb.saveToExcelFile(excelDatabasePublicFile, submissions, staff, addresses); } catch(e) {}
    }
  } catch (err) {
    console.error('[ExcelDB] Error in syncDatabaseToExcel:', err);
  }
}

function mergeById(localList, remoteList, idKey) {
  const map = new Map();
  if (Array.isArray(remoteList)) {
    remoteList.forEach(item => {
      const key = item[idKey];
      if (key) map.set(String(key).trim().toLowerCase(), item);
    });
  }
  if (Array.isArray(localList)) {
    localList.forEach(item => {
      const key = item[idKey];
      if (key) {
        const k = String(key).trim().toLowerCase();
        const existing = map.get(k);
        map.set(k, existing ? { ...existing, ...item } : item);
      } else {
        map.set('TEMP_' + Math.random(), item);
      }
    });
  }
  return Array.from(map.values());
}

async function autoRestoreFromGoogleSheets() {
  if (!googleSheetsSync.isConfigured()) return;
  try {
    const pulled = await googleSheetsSync.pullFromGoogleSheets();
    if (pulled.success) {
      if (pulled.students && pulled.students.length > 0) {
        const localSubs = readJson(submissionsFile, []);
        const merged = localSubs.length === 0 ? pulled.students : mergeById(localSubs, pulled.students, 'submissionId');
        fs.writeFileSync(submissionsFile, JSON.stringify(merged, null, 2), 'utf8');
      }
      if (pulled.staff && pulled.staff.length > 0) {
        const localStaff = readJson(staffFile, []);
        const merged = localStaff.length === 0 ? pulled.staff : mergeById(localStaff, pulled.staff, 'staffId');
        fs.writeFileSync(staffFile, JSON.stringify(merged, null, 2), 'utf8');
      }
      if (pulled.addresses && pulled.addresses.length > 0) {
        const localAddr = readJson(addressesFile, []);
        const merged = localAddr.length === 0 ? pulled.addresses : mergeById(localAddr, pulled.addresses, 'id');
        fs.writeFileSync(addressesFile, JSON.stringify(merged, null, 2), 'utf8');
      }
      syncDatabaseToExcel();
      console.log(`[GoogleSheets] Live auto-restore successful: ${pulled.students?.length || 0} students, ${pulled.staff?.length || 0} staff, ${pulled.addresses?.length || 0} addresses.`);
    }
  } catch (err) {
    console.warn('[GoogleSheets] Auto-restore warning:', err.message);
  }
}

function ensureDataFiles() {
  if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });
  if (!fs.existsSync(submissionsFile)) fs.writeFileSync(submissionsFile, '[]', 'utf8');
  if (!fs.existsSync(detailsFile)) fs.writeFileSync(detailsFile, '{}', 'utf8');
  if (!fs.existsSync(adminConfigFile)) fs.writeFileSync(adminConfigFile, JSON.stringify({ pin: '1234' }, null, 2), 'utf8');
  if (!fs.existsSync(staffFile)) fs.writeFileSync(staffFile, '[]', 'utf8');
  if (!fs.existsSync(leavesFile)) fs.writeFileSync(leavesFile, '[]', 'utf8');
  if (!fs.existsSync(tasksFile)) fs.writeFileSync(tasksFile, '[]', 'utf8');
  if (!fs.existsSync(notificationsFile)) fs.writeFileSync(notificationsFile, '[]', 'utf8');
  if (!fs.existsSync(addressesFile)) fs.writeFileSync(addressesFile, '[]', 'utf8');

  // Check if permanent Excel database exists to restore empty files
  const excelTarget = fs.existsSync(excelDatabaseFile) ? excelDatabaseFile : (fs.existsSync(excelDatabasePublicFile) ? excelDatabasePublicFile : null);
  const currentSubmissions = readJson(submissionsFile, []);
  const currentStaff = readJson(staffFile, []);
  const currentAddresses = readJson(addressesFile, []);

  if (excelTarget && currentSubmissions.length === 0 && currentStaff.length === 0 && currentAddresses.length === 0) {
    const restored = excelDb.readFromExcelFile(excelTarget);
    if (restored) {
      if (restored.staff && restored.staff.length > 0) fs.writeFileSync(staffFile, JSON.stringify(restored.staff, null, 2), 'utf8');
      if (restored.submissions && restored.submissions.length > 0) fs.writeFileSync(submissionsFile, JSON.stringify(restored.submissions, null, 2), 'utf8');
      if (restored.addresses && restored.addresses.length > 0) fs.writeFileSync(addressesFile, JSON.stringify(restored.addresses, null, 2), 'utf8');
      console.log(`[ExcelDB] Auto-restored data from Excel database (${restored.staff.length} staff, ${restored.submissions.length} students, ${restored.addresses.length} addresses)`);
    }
  }

  // Ensure Excel database is synced with current data
  syncDatabaseToExcel();

  // If Google Sheets is configured, trigger auto-restore in background on startup
  if (googleSheetsSync.isConfigured()) {
    autoRestoreFromGoogleSheets();
  }
}

function readJson(filePath, fallback) {
  try {
    return JSON.parse(fs.readFileSync(filePath, 'utf8'));
  } catch {
    return fallback;
  }
}

function writeJson(filePath, data) {
  fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf8');
  if (filePath === submissionsFile || filePath === staffFile || filePath === addressesFile) {
    syncDatabaseToExcel();
    try {
      googleSheetsSync.triggerAutoSync(
        () => readJson(submissionsFile, []),
        () => readJson(staffFile, []),
        () => readJson(addressesFile, [])
      );
    } catch (e) {
      console.warn('[GoogleSheets] AutoSync error:', e);
    }
  }
}

function sendIfFileExists(res, fileName) {
  const safePath = path.join(__dirname, fileName);
  if (fs.existsSync(safePath) && fs.statSync(safePath).isFile()) {
    res.sendFile(safePath);
    return true;
  }
  return false;
}

ensureDataFiles();

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

app.get('/api/sync/status', (req, res) => {
  const cfg = googleSheetsSync.getConfig();
  res.json({
    success: true,
    configured: googleSheetsSync.isConfigured(),
    spreadsheetName: cfg.spreadsheetName || 'IIT Database Master',
    spreadsheetUrl: cfg.spreadsheetUrl || '',
    autoSync: !!cfg.autoSync,
    lastSyncTime: cfg.lastSyncTime || null,
    lastSyncStatus: cfg.lastSyncStatus || null,
    lastSyncCounts: cfg.lastSyncCounts || null
  });
});

app.post('/api/details', (req, res) => {
  writeJson(detailsFile, req.body);
  res.json({ success: true, saved: req.body });
});

app.get('/api/details', (req, res) => {
  res.json(readJson(detailsFile, {}));
});

function getNextSubmissionId(submissions) {
  let maxId = 0;
  submissions.forEach(s => {
    if (s.submissionId && s.submissionId.startsWith('ID-')) {
      const num = parseInt(s.submissionId.substring(3), 10);
      if (!isNaN(num) && num > maxId) {
        maxId = num;
      }
    }
  });
  return `ID-${String(maxId + 1).padStart(4, '0')}`;
}



app.post('/api/submit', async (req, res) => {
  const submissions = readJson(submissionsFile, []);
  const payload = req.body;
  
  if (payload.submissionId && payload.submissionId !== 'N/A' && payload.submissionId !== '') {
    const existingIndex = submissions.findIndex(s => s.submissionId === payload.submissionId);
    if (existingIndex !== -1) {
      submissions[existingIndex] = {
        ...submissions[existingIndex],
        ...payload,
        submittedAt: payload.submittedAt || submissions[existingIndex].submittedAt || new Date().toLocaleString()
      };
      writeJson(submissionsFile, submissions);

      let liveSynced = false;
      if (googleSheetsSync.isConfigured()) {
        try {
          const syncRes = await Promise.race([
            googleSheetsSync.upsertStudent(submissions[existingIndex]),
            new Promise((_, reject) => setTimeout(() => reject(new Error('Sync timeout')), 7000))
          ]);
          if (syncRes && syncRes.success) liveSynced = true;
        } catch (e) {
          console.warn('[GoogleSheets] Edit sync warning:', e.message);
        }
      }

      res.json({ success: true, saved: submissions[existingIndex], liveSynced });
      return;
    }
  }

  // Validate email uniqueness for new submissions
  if (payload.email && payload.email.trim()) {
    const emailExists = submissions.some(s => s.email && s.email.trim().toLowerCase() === payload.email.trim().toLowerCase());
    if (emailExists) {
      return res.status(400).json({ success: false, error: 'This email is already registered. Each email can only be used once.' });
    }
  }

  const generatedId = getNextSubmissionId(submissions);
  const newPayload = { 
    ...payload, 
    submissionId: payload.submissionId && payload.submissionId !== 'N/A' && payload.submissionId !== '' ? payload.submissionId : generatedId,
    submittedAt: payload.submittedAt || new Date().toLocaleString() 
  };
  submissions.push(newPayload);
  writeJson(submissionsFile, submissions);

  // Live permanent sync to Google Sheets
  let liveSynced = false;
  if (googleSheetsSync.isConfigured()) {
    try {
      const syncRes = await Promise.race([
        googleSheetsSync.upsertStudent(newPayload),
        new Promise((_, reject) => setTimeout(() => reject(new Error('Sync timeout')), 7000))
      ]);
      if (syncRes && syncRes.success) liveSynced = true;
    } catch (e) {
      console.warn('[GoogleSheets] Submission sync warning:', e.message);
    }
  }

  // Notify admin
  try {
    const notifications = readJson(notificationsFile, []);
    notifications.push({
      id: 'NTF-' + crypto.randomBytes(4).toString('hex').toUpperCase(),
      staffId: 'admin',
      message: `New student application submitted by "${newPayload.name || 'Unnamed'}" (ID: ${newPayload.submissionId})`,
      type: 'submission',
      read: false,
      createdAt: new Date().toISOString()
    });
    writeJson(notificationsFile, notifications);
  } catch (err) {
    console.error('Failed to create admin notification for submission', err);
  }

  res.json({ success: true, saved: newPayload, liveSynced });
});

app.post('/api/delete', requireAdmin, async (req, res) => {
  const { submissionId } = req.body;
  const submissions = readJson(submissionsFile, []);
  const filtered = submissions.filter(s => s.submissionId !== submissionId);
  writeJson(submissionsFile, filtered);

  if (googleSheetsSync.isConfigured()) {
    try {
      await googleSheetsSync.deleteStudent(submissionId);
    } catch (e) {
      console.warn('[GoogleSheets] Delete sync warning:', e.message);
    }
  }

  res.json({ success: true });
});

// ── Staff Management Endpoints ──
function getNextStaffId(staff) {
  let maxId = 0;
  const year = new Date().getFullYear();
  const prefix = `STF${year}`;
  staff.forEach(s => {
    if (s.staffId && s.staffId.startsWith(prefix)) {
      const num = parseInt(s.staffId.substring(prefix.length), 10);
      if (!isNaN(num) && num > maxId) {
        maxId = num;
      }
    }
  });
  return `${prefix}${String(maxId + 1).padStart(4, '0')}`;
}

app.get('/api/staff', async (req, res) => {
  let staffList = readJson(staffFile, []);
  const forceRefresh = req.query.refresh === 'true';
  if (googleSheetsSync.isConfigured() && (staffList.length === 0 || forceRefresh || googleSheetsSync.shouldRefresh())) {
    try {
      const pulled = await googleSheetsSync.pullFromGoogleSheets();
      if (pulled.success && Array.isArray(pulled.staff) && pulled.staff.length > 0) {
        staffList = mergeById(staffList, pulled.staff, 'staffId');
        fs.writeFileSync(staffFile, JSON.stringify(staffList, null, 2), 'utf8');
        syncDatabaseToExcel();
      }
    } catch (e) {
      console.warn('[GoogleSheets] Fetch staff from live sheet failed:', e.message);
    }
  }
  // Strip sensitive salt/hash from response
  const publicList = staffList.map(s => {
    const { passwordHash, salt, ...publicData } = s;
    return publicData;
  });
  res.json(publicList);
});

// Admin registers or edits staff
app.post('/api/staff', requireAdmin, async (req, res) => {
  const staffList = readJson(staffFile, []);
  const payload = req.body;
  const { staffId, name, mobile, email, department, role, password, status, photo, docName, docBase64, isEdit } = payload;
  
  if (!name || !mobile || !email || !role || !department) {
    return res.status(400).json({ success: false, error: 'Name, Mobile, Email, Department, and Role/Designation are required fields.' });
  }

  // Email unique check
  const emailIndex = staffList.findIndex(s => s.email.toLowerCase() === email.toLowerCase() && s.staffId !== staffId);
  if (emailIndex !== -1) {
    return res.status(400).json({ success: false, error: 'A staff member with this Email already exists.' });
  }
  
  // Mobile unique check
  const mobileIndex = staffList.findIndex(s => s.mobile === mobile && s.staffId !== staffId);
  if (mobileIndex !== -1) {
    return res.status(400).json({ success: false, error: 'A staff member with this Mobile number already exists.' });
  }

  if (isEdit) {
    const existingIndex = staffList.findIndex(s => s.staffId === staffId);
    if (existingIndex === -1) {
      return res.status(404).json({ success: false, error: 'Staff member to edit was not found.' });
    }
    const existing = staffList[existingIndex];
    let updatedHash = existing.passwordHash;
    let updatedSalt = existing.salt;
    
    if (password && password.trim() !== '') {
      updatedSalt = crypto.randomBytes(16).toString('hex');
      updatedHash = hashPassword(password, updatedSalt);
    }
    
    staffList[existingIndex] = {
      ...existing,
      name,
      mobile,
      email,
      department,
      role,
      status: status || existing.status,
      photo: photo !== undefined ? photo : existing.photo,
      docName: docName !== undefined ? docName : existing.docName,
      docBase64: docBase64 !== undefined ? docBase64 : existing.docBase64,
      passwordHash: updatedHash,
      salt: updatedSalt,
      updatedAt: new Date().toLocaleString()
    };
    
    writeJson(staffFile, staffList);

    if (googleSheetsSync.isConfigured()) {
      try {
        await googleSheetsSync.upsertStaff(staffList[existingIndex]);
      } catch (e) {
        console.warn('[GoogleSheets] Staff edit sync warning:', e.message);
      }
    }

    const { passwordHash, salt, ...publicData } = staffList[existingIndex];
    return res.json({ success: true, saved: publicData });
  }

  // Create new staff member
  if (!staffId || staffId.trim() === '') {
    return res.status(400).json({ success: false, error: 'Staff ID is required.' });
  }

  const idExists = staffList.some(s => s.staffId.toUpperCase() === staffId.trim().toUpperCase());
  if (idExists) {
    return res.status(400).json({ success: false, error: 'A staff member with this Staff ID already exists.' });
  }

  if (!password || password.trim() === '') {
    return res.status(400).json({ success: false, error: 'Password is required for new registration.' });
  }
  
  const salt = crypto.randomBytes(16).toString('hex');
  const passwordHash = hashPassword(password, salt);
  
  const newStaff = {
    staffId: staffId.trim(),
    name,
    mobile,
    email,
    department,
    role,
    status: status || 'Active',
    photo: photo || '',
    docName: docName || '',
    docBase64: docBase64 || '',
    passwordHash,
    salt,
    createdAt: new Date().toLocaleString()
  };
  
  staffList.push(newStaff);
  writeJson(staffFile, staffList);

  if (googleSheetsSync.isConfigured()) {
    try {
      await googleSheetsSync.upsertStaff(newStaff);
    } catch (e) {
      console.warn('[GoogleSheets] New staff sync warning:', e.message);
    }
  }
  
  const { passwordHash: ph, salt: st, ...publicData } = newStaff;
  res.json({ success: true, saved: publicData });
});

// Admin toggle status
app.post('/api/staff/toggle', requireAdmin, (req, res) => {
  const { staffId, status } = req.body;
  const staffList = readJson(staffFile, []);
  const index = staffList.findIndex(s => s.staffId === staffId);
  if (index === -1) return res.status(404).json({ success: false, error: 'Staff not found.' });
  
  staffList[index].status = status;
  writeJson(staffFile, staffList);
  res.json({ success: true, status });
});

// Admin reset password
app.post('/api/staff/reset-password', requireAdmin, (req, res) => {
  const { staffId, newPassword } = req.body;
  const staffList = readJson(staffFile, []);
  const index = staffList.findIndex(s => s.staffId === staffId);
  if (index === -1) return res.status(404).json({ success: false, error: 'Staff not found.' });
  
  const salt = crypto.randomBytes(16).toString('hex');
  staffList[index].salt = salt;
  staffList[index].passwordHash = hashPassword(newPassword, salt);
  writeJson(staffFile, staffList);
  res.json({ success: true });
});

app.post('/api/staff/delete', requireAdmin, (req, res) => {
  const { staffId } = req.body;
  const staffList = readJson(staffFile, []);
  const filtered = staffList.filter(s => s.staffId !== staffId);
  writeJson(staffFile, filtered);
  res.json({ success: true });
});

// Admin login as staff (without password)
app.post('/api/admin/login-as-staff', requireAdmin, (req, res) => {
  const { staffId } = req.body;
  if (!staffId) return res.status(400).json({ success: false, error: 'Staff ID is required.' });
  
  const staffList = readJson(staffFile, []);
  const staff = staffList.find(s => s.staffId.toUpperCase() === staffId.trim().toUpperCase());
  if (!staff) {
    return res.status(404).json({ success: false, error: 'Staff member not found.' });
  }
  
  if (staff.status === 'Inactive' || staff.status === 'Disabled') {
    return res.status(403).json({ success: false, error: 'This staff account is inactive or disabled.' });
  }
  
  const token = generateSessionToken({ staffId: staff.staffId, role: 'staff', name: staff.name });
  res.json({ success: true, token, name: staff.name, staffId: staff.staffId });
});


// Staff Authentication Login
app.post('/api/staff/login', (req, res) => {
  const { staffId, password } = req.body;
  if (!staffId || !password) return res.status(400).json({ success: false, error: 'Staff ID and password are required.' });
  
  const staffList = readJson(staffFile, []);
  const staff = staffList.find(s => s.staffId.toUpperCase() === staffId.trim().toUpperCase());
  
  if (!staff) {
    return res.status(401).json({ success: false, error: 'Incorrect Staff ID or password.' });
  }
  
  if (staff.status === 'Inactive' || staff.status === 'Disabled') {
    return res.status(403).json({ success: false, error: 'Your account is disabled. Contact Admin.' });
  }
  
  const calculatedHash = hashPassword(password, staff.salt);
  if (calculatedHash !== staff.passwordHash) {
    return res.status(401).json({ success: false, error: 'Incorrect Staff ID or password.' });
  }
  
  // Generate signed session token
  const token = generateSessionToken({ staffId: staff.staffId, role: 'staff', name: staff.name });
  res.json({ success: true, token, name: staff.name, staffId: staff.staffId });
});

// Fetch Profile
app.get('/api/staff/profile', requireStaff, (req, res) => {
  const staffList = readJson(staffFile, []);
  const staff = staffList.find(s => s.staffId === req.staffId);
  if (!staff) return res.status(404).json({ success: false, error: 'Staff not found.' });
  const { passwordHash, salt, ...profile } = staff;
  res.json(profile);
});

// ── Leave & OD Management Endpoints ──
app.post('/api/staff/leave', requireStaff, (req, res) => {
  const { type, startDate, endDate, reason, docName, docBase64 } = req.body;
  if (!type || !startDate || !endDate || !reason) {
    return res.status(400).json({ success: false, error: 'Leave type, start date, end date, and reason are required.' });
  }

  const leaves = readJson(leavesFile, []);
  const newRequest = {
    id: 'REQ-' + crypto.randomBytes(4).toString('hex').toUpperCase(),
    staffId: req.staffId,
    type, // 'Leave' or 'On-Duty'
    startDate,
    endDate,
    reason,
    docName: docName || '',
    docBase64: docBase64 || '',
    status: 'Pending',
    remarks: '',
    createdAt: new Date().toISOString()
  };

  leaves.push(newRequest);
  writeJson(leavesFile, leaves);

  // Notify admin
  try {
    const staffList = readJson(staffFile, []);
    const staff = staffList.find(s => s.staffId === req.staffId) || {};
    const staffName = staff.name || req.staffId;

    const notifications = readJson(notificationsFile, []);
    notifications.push({
      id: 'NTF-' + crypto.randomBytes(4).toString('hex').toUpperCase(),
      staffId: 'admin',
      message: `New ${type} request from ${staffName} (Reason: ${reason.slice(0, 30)}${reason.length > 30 ? '...' : ''})`,
      type: 'leave',
      read: false,
      createdAt: new Date().toISOString()
    });
    writeJson(notificationsFile, notifications);
  } catch (err) {
    console.error('Failed to create admin notification for leave request', err);
  }

  res.json({ success: true, request: newRequest });
});

// Fetch leaves for staff
app.get('/api/staff/leaves', requireStaff, (req, res) => {
  const leaves = readJson(leavesFile, []);
  const filtered = leaves.filter(l => l.staffId === req.staffId);
  res.json(filtered);
});

// Fetch leaves for admin
app.get('/api/admin/leaves', requireAdmin, (req, res) => {
  const leaves = readJson(leavesFile, []);
  const staffList = readJson(staffFile, []);
  
  const detailed = leaves.map(item => {
    const s = staffList.find(x => x.staffId === item.staffId) || {};
    return {
      ...item,
      staffName: s.name || 'Unknown',
      staffRole: s.role || 'N/A',
      staffDept: s.department || 'N/A'
    };
  });
  res.json(detailed);
});

// Admin approve/reject leave or OD
app.post('/api/admin/leaves/action', requireAdmin, (req, res) => {
  const { id, status, remarks } = req.body;
  if (!id || !status) return res.status(400).json({ success: false, error: 'Request ID and action status are required.' });
  
  const leaves = readJson(leavesFile, []);
  const index = leaves.findIndex(l => l.id === id);
  if (index === -1) return res.status(404).json({ success: false, error: 'Request not found.' });

  leaves[index].status = status; // 'Approved' or 'Rejected'
  leaves[index].remarks = remarks || '';
  writeJson(leavesFile, leaves);

  // Add notification for staff
  const notifications = readJson(notificationsFile, []);
  notifications.push({
    id: 'NTF-' + crypto.randomBytes(4).toString('hex').toUpperCase(),
    staffId: leaves[index].staffId,
    message: `Your ${leaves[index].type} request from ${leaves[index].startDate} has been ${status.toLowerCase()}.${remarks ? ' Remarks: ' + remarks : ''}`,
    type: 'leave',
    read: false,
    createdAt: new Date().toISOString()
  });
  writeJson(notificationsFile, notifications);

  res.json({ success: true });
});

// ── Work Assignment / Task Endpoints ──
// Admin assign task
app.post('/api/admin/tasks', requireAdmin, (req, res) => {
  const { title, description, priority, dueDate, department, assignedStaffId, attachmentName, attachmentBase64 } = req.body;
  if (!title || !description || !priority || !dueDate || !assignedStaffId) {
    return res.status(400).json({ success: false, error: 'Title, description, priority, due date, and assigned staff are required.' });
  }

  const tasks = readJson(tasksFile, []);
  const staffList = readJson(staffFile, []);
  const s = staffList.find(x => x.staffId === assignedStaffId);
  if (!s) return res.status(400).json({ success: false, error: 'Assigned staff member not found.' });

  const newTask = {
    id: 'TSK-' + crypto.randomBytes(4).toString('hex').toUpperCase(),
    title,
    description,
    priority, // 'Low', 'Medium', 'High'
    assignedDate: new Date().toLocaleDateString('en-IN'),
    dueDate,
    department: department || s.department || 'N/A',
    assignedStaffId,
    assignedStaffName: s.name,
    attachmentName: attachmentName || '',
    attachmentBase64: attachmentBase64 || '',
    status: 'Pending', // 'Pending', 'In Progress', 'Completed'
    completionTime: '',
    comments: '',
    progress: 0,
    progressUpdates: [] // array of updates
  };

  tasks.push(newTask);
  writeJson(tasksFile, tasks);

  // Add notification
  const notifications = readJson(notificationsFile, []);
  notifications.push({
    id: 'NTF-' + crypto.randomBytes(4).toString('hex').toUpperCase(),
    staffId: assignedStaffId,
    message: `New task assigned: "${title}" (Due: ${dueDate})`,
    type: 'task',
    read: false,
    createdAt: new Date().toISOString()
  });
  writeJson(notificationsFile, notifications);

  res.json({ success: true, task: newTask });
});

// Fetch tasks for admin
app.get('/api/admin/tasks', requireAdmin, (req, res) => {
  res.json(readJson(tasksFile, []));
});

// Fetch tasks for staff
app.get('/api/staff/tasks', requireStaff, (req, res) => {
  const tasks = readJson(tasksFile, []);
  const filtered = tasks.filter(t => t.assignedStaffId === req.staffId);
  res.json(filtered);
});

// Staff update task status
app.post('/api/staff/tasks/status', requireStaff, (req, res) => {
  const { id, status, comments } = req.body;
  if (!id || !status) return res.status(400).json({ success: false, error: 'Task ID and status are required.' });

  const tasks = readJson(tasksFile, []);
  const index = tasks.findIndex(t => t.id === id && t.assignedStaffId === req.staffId);
  if (index === -1) return res.status(404).json({ success: false, error: 'Task not found.' });

  tasks[index].status = status;
  if (comments !== undefined) tasks[index].comments = comments;
  
  if (status === 'Completed') {
    tasks[index].completionTime = new Date().toLocaleString();
    tasks[index].progress = 100;
  } else if (status === 'In Progress' && tasks[index].progress === 0) {
    tasks[index].progress = 10;
  }
  
  writeJson(tasksFile, tasks);

  // Notify admin
  try {
    const staffList = readJson(staffFile, []);
    const staff = staffList.find(s => s.staffId === req.staffId) || {};
    const staffName = staff.name || req.staffId;

    const notifications = readJson(notificationsFile, []);
    notifications.push({
      id: 'NTF-' + crypto.randomBytes(4).toString('hex').toUpperCase(),
      staffId: 'admin',
      message: `${staffName} updated task status for "${tasks[index].title}": ${status}.`,
      type: 'task',
      read: false,
      createdAt: new Date().toISOString()
    });
    writeJson(notificationsFile, notifications);
  } catch (err) {
    console.error('Failed to create admin notification for task status', err);
  }

  res.json({ success: true, task: tasks[index] });
});

// Staff post progress update
app.post('/api/staff/tasks/progress', requireStaff, (req, res) => {
  const { id, progress, comments, fileName, fileBase64 } = req.body;
  if (!id || progress === undefined) return res.status(400).json({ success: false, error: 'Task ID and progress percentage are required.' });

  const tasks = readJson(tasksFile, []);
  const index = tasks.findIndex(t => t.id === id && t.assignedStaffId === req.staffId);
  if (index === -1) return res.status(404).json({ success: false, error: 'Task not found.' });

  const progressPct = parseInt(progress, 10);
  tasks[index].progress = progressPct;
  if (progressPct === 100) {
    tasks[index].status = 'Completed';
    tasks[index].completionTime = new Date().toLocaleString();
  }

  const updateRecord = {
    progress: progressPct,
    comments: comments || '',
    fileName: fileName || '',
    fileBase64: fileBase64 || '',
    timestamp: new Date().toLocaleString()
  };

  tasks[index].progressUpdates.push(updateRecord);
  writeJson(tasksFile, tasks);

  // Notify admin
  try {
    const staffList = readJson(staffFile, []);
    const staff = staffList.find(s => s.staffId === req.staffId) || {};
    const staffName = staff.name || req.staffId;

    const notifications = readJson(notificationsFile, []);
    notifications.push({
      id: 'NTF-' + crypto.randomBytes(4).toString('hex').toUpperCase(),
      staffId: 'admin',
      message: `${staffName} updated task progress for "${tasks[index].title}": ${progressPct}%.`,
      type: 'task',
      read: false,
      createdAt: new Date().toISOString()
    });
    writeJson(notificationsFile, notifications);
  } catch (err) {
    console.error('Failed to create admin notification for task progress', err);
  }

  res.json({ success: true, task: tasks[index] });
});

// Staff notifications list
app.get('/api/staff/notifications', requireStaff, (req, res) => {
  const list = readJson(notificationsFile, []);
  const filtered = list.filter(n => n.staffId === req.staffId);
  res.json(filtered.reverse());
});

// Staff mark notification read
app.post('/api/staff/notifications/read', requireStaff, (req, res) => {
  const { id } = req.body;
  const list = readJson(notificationsFile, []);
  const index = list.findIndex(n => n.id === id && n.staffId === req.staffId);
  if (index !== -1) {
    list[index].read = true;
    writeJson(notificationsFile, list);
  }
  res.json({ success: true });
});

// ── Admin Endpoints ──
// Admin notifications list
app.get('/api/admin/notifications', requireAdmin, (req, res) => {
  const list = readJson(notificationsFile, []);
  const filtered = list.filter(n => n.staffId === 'admin');
  res.json(filtered.reverse());
});

// Admin mark notification read
app.post('/api/admin/notifications/read', requireAdmin, (req, res) => {
  const { id } = req.body;
  const list = readJson(notificationsFile, []);
  const index = list.findIndex(n => n.id === id && n.staffId === 'admin');
  if (index !== -1) {
    list[index].read = true;
    writeJson(notificationsFile, list);
  }
  res.json({ success: true });
});

// Admin mark all notifications read
app.post('/api/admin/notifications/read-all', requireAdmin, (req, res) => {
  const list = readJson(notificationsFile, []);
  list.forEach(n => {
    if (n.staffId === 'admin') n.read = true;
  });
  writeJson(notificationsFile, list);
  res.json({ success: true });
});

// Fetch admin config (Office Timings and details)
app.get('/api/admin/config', requireAdmin, (req, res) => {
  const config = getAdminConfig();
  res.json(config);
});

// Update admin config
app.post('/api/admin/config', requireAdmin, (req, res) => {
  const { pin, officeStartTime, officeEndTime } = req.body;
  const config = getAdminConfig();
  
  if (pin) config.pin = pin;
  if (officeStartTime) config.officeStartTime = officeStartTime;
  if (officeEndTime) config.officeEndTime = officeEndTime;
  
  writeJson(adminConfigFile, config);
  res.json({ success: true, config });
});

app.post('/api/admin/login', (req, res) => {
  const { pin } = req.body;
  const correctPin = getAdminPin();
  if (!pin || pin !== correctPin) {
    return res.status(401).json({ success: false, error: 'Incorrect PIN' });
  }
  const token = generateSessionToken({ role: 'admin' });
  console.log(`[Admin] New JWT session started at ${new Date().toISOString()}`);
  res.json({ success: true, token });
});

app.post('/api/admin/logout', (req, res) => {
  console.log('[Admin] Session logout');
  res.json({ success: true });
});

app.get('/api/admin/check', (req, res) => {
  const token = getAuthToken(req);
  res.json({ valid: isValidToken(token) });
});

app.post('/api/admin/pin', requireAdmin, (req, res) => {
  const { currentPin, newPin } = req.body;
  if (currentPin !== getAdminPin()) {
    return res.status(401).json({ success: false, error: 'Current PIN is incorrect' });
  }
  if (!newPin || !/^\d{4}$/.test(newPin)) {
    return res.status(400).json({ success: false, error: 'New PIN must be 4 digits' });
  }
  writeJson(adminConfigFile, { pin: newPin });
  res.json({ success: true });
});

// Admin restore data endpoint (to sync localStorage data back to server after restart)
app.post('/api/admin/restore-data', requireAdmin, (req, res) => {
  const { type, data } = req.body;
  if (!type || !Array.isArray(data)) {
    return res.status(400).json({ success: false, error: 'Invalid restore payload.' });
  }

  let filePath;
  switch (type) {
    case 'submissions':
      filePath = submissionsFile;
      break;
    case 'staff':
      filePath = staffFile;
      break;
    case 'leaves':
      filePath = leavesFile;
      break;
    case 'tasks':
      filePath = tasksFile;
      break;
    case 'notifications':
      filePath = notificationsFile;
      break;
    default:
      return res.status(400).json({ success: false, error: 'Unknown data type.' });
  }

  // Restore the data to the JSON file
  writeJson(filePath, data);
  console.log(`[Admin] Restored ${data.length} records for ${type}`);
  res.json({ success: true, count: data.length });
});

app.get('/api/submissions', async (req, res) => {
  let submissions = readJson(submissionsFile, []);
  const forceRefresh = req.query.refresh === 'true';
  if (googleSheetsSync.isConfigured() && (submissions.length === 0 || forceRefresh || googleSheetsSync.shouldRefresh())) {
    try {
      const pulled = await googleSheetsSync.pullFromGoogleSheets();
      if (pulled.success && Array.isArray(pulled.students) && pulled.students.length > 0) {
        submissions = mergeById(submissions, pulled.students, 'submissionId');
        fs.writeFileSync(submissionsFile, JSON.stringify(submissions, null, 2), 'utf8');
        syncDatabaseToExcel();
      }
    } catch (e) {
      console.warn('[GoogleSheets] Fetch submissions from live sheet failed:', e.message);
    }
  }
  res.json(submissions);
});

app.get('/api/pages', (req, res) => {
  res.json(['details.html', 'form.html', 'ss.html']);
});

app.get('/details.html', (req, res) => {
  if (!sendIfFileExists(res, 'details.html')) {
    res.status(404).send('Not found');
  }
});

app.get('/form.html', (req, res) => {
  if (!sendIfFileExists(res, 'form.html')) {
    res.status(404).send('Not found');
  }
});

app.get('/ss.html', (req, res) => {
  if (!sendIfFileExists(res, 'ss.html')) {
    res.status(404).send('Not found');
  }
});

// ── Address Management API Endpoints ─────────────────────────────────
app.get('/api/addresses', async (req, res) => {
  let addresses = readJson(addressesFile, []);
  const forceRefresh = req.query.refresh === 'true';
  if (googleSheetsSync.isConfigured() && (addresses.length === 0 || forceRefresh || googleSheetsSync.shouldRefresh())) {
    try {
      const pulled = await googleSheetsSync.pullFromGoogleSheets();
      if (pulled.success && Array.isArray(pulled.addresses) && pulled.addresses.length > 0) {
        addresses = mergeById(addresses, pulled.addresses, 'id');
        fs.writeFileSync(addressesFile, JSON.stringify(addresses, null, 2), 'utf8');
        syncDatabaseToExcel();
      }
    } catch (e) {
      console.warn('[GoogleSheets] Fetch addresses from live sheet failed:', e.message);
    }
  }
  res.json({ success: true, addresses });
});

app.post('/api/addresses', async (req, res) => {
  const payload = req.body;
  const addresses = readJson(addressesFile, []);

  let recordToSave;
  if (payload.to && payload.from) {
    const to = payload.to;
    if (!to.name || !to.line1 || !to.city || !to.pincode || !to.phone) {
      return res.status(400).json({ success: false, error: 'Recipient Name, Address Line 1, City, Pincode, and Phone are required.' });
    }
    const recordId = payload.id || ('ADDR-' + String(Date.now()).slice(-4) + '-' + String(Math.floor(100 + Math.random() * 900)));
    recordToSave = {
      id: recordId,
      dateTime: payload.dateTime || new Date().toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' }),
      to: payload.to,
      from: payload.from,
      createdAt: payload.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
  } else {
    // Flat address structure support
    const { name, city, pincode, contactPerson, phone, email, addressLine1, addressLine2, state, category, notes, title } = payload;
    if (!name || !city || !pincode) {
      return res.status(400).json({ success: false, error: 'Name, City, and Pincode are required.' });
    }
    const recordId = payload.id || ('ADDR-' + String(Date.now()).slice(-4) + '-' + String(Math.floor(100 + Math.random() * 900)));
    recordToSave = {
      id: recordId,
      dateTime: new Date().toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' }),
      to: {
        name: name.trim(),
        line1: (addressLine1 || '').trim(),
        line2: (addressLine2 || '').trim(),
        line3: '',
        pincode: (pincode || '').trim(),
        phone: (phone || '').trim(),
        extraPhone: '',
        city: (city || '').trim(),
        state: (state || '').trim(),
        country: 'India'
      },
      from: {
        name: 'Indian Info Tech',
        line1: '101-C, KM Complex,',
        line2: 'Amma Unavagam Back Side,',
        line3: '',
        pincode: '637001',
        phone: '9003446612',
        extraPhone: '',
        city: 'Namakkal',
        state: 'Tamilnadu',
        country: 'India'
      },
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
  }

  // Smart Duplicate Detection: Match by exact ID, or by Recipient Name + Phone / Line 1
  const toSaveName = (recordToSave.to?.name || '').trim().toLowerCase();
  const toSavePhone = (recordToSave.to?.phone || '').replace(/\D/g, '');
  const toSaveLine1 = (recordToSave.to?.line1 || '').trim().toLowerCase();

  const existingIdx = addresses.findIndex(a => {
    if (!a) return false;
    if (a.id && a.id === recordToSave.id) return true;
    const aTo = a.to || {};
    const aName = (aTo.name || a.name || '').trim().toLowerCase();
    const aPhone = (aTo.phone || a.phone || '').replace(/\D/g, '');
    const aLine1 = (aTo.line1 || a.addressLine1 || '').trim().toLowerCase();

    // Match if same Name and same Phone
    if (toSaveName && aName === toSaveName && toSavePhone && aPhone === toSavePhone) return true;
    // Or if same Name and same Address Line 1
    if (toSaveName && aName === toSaveName && toSaveLine1 && aLine1 === toSaveLine1) return true;
    return false;
  });

  if (existingIdx !== -1) {
    // Retain original ID and update in-place to prevent repeated duplicate entries
    const origId = addresses[existingIdx].id || recordToSave.id;
    recordToSave.id = origId;
    recordToSave.updatedAt = new Date().toISOString();
    addresses[existingIdx] = recordToSave;
  } else {
    addresses.unshift(recordToSave);
  }

  writeJson(addressesFile, addresses);

  if (googleSheetsSync.isConfigured()) {
    googleSheetsSync.upsertAddress(recordToSave).catch(e => {
      console.warn('[GoogleSheets] Address upsert background warning:', e.message);
    });
  }

  res.json({ success: true, address: recordToSave });
});

app.post('/api/addresses/clean-duplicates', (req, res) => {
  let addresses = readJson(addressesFile, []);
  const initialLen = addresses.length;
  const seen = new Set();
  const cleaned = [];

  for (const a of addresses) {
    if (!a) continue;
    const to = a.to || {};
    const name = (to.name || a.name || '').trim().toLowerCase();
    const phone = (to.phone || a.phone || '').replace(/\D/g, '');
    const line1 = (to.line1 || a.addressLine1 || '').trim().toLowerCase();
    const sig = phone ? `${name}::${phone}` : `${name}::${line1}`;

    if (!sig || sig === '::') {
      cleaned.push(a);
      continue;
    }
    if (!seen.has(sig)) {
      seen.add(sig);
      cleaned.push(a);
    }
  }

  const removedCount = initialLen - cleaned.length;
  if (removedCount > 0) {
    writeJson(addressesFile, cleaned);
  }
  res.json({ success: true, removed: removedCount, total: cleaned.length, addresses: cleaned });
});

app.put('/api/addresses/:id', async (req, res) => {
  const { id } = req.params;
  const addresses = readJson(addressesFile, []);
  const index = addresses.findIndex(a => a.id === id);
  if (index === -1) {
    return res.status(404).json({ success: false, error: 'Address record not found.' });
  }

  const existing = addresses[index];
  const updated = {
    ...existing,
    ...req.body,
    id: existing.id,
    updatedAt: new Date().toISOString()
  };

  addresses[index] = updated;
  writeJson(addressesFile, addresses);

  if (googleSheetsSync.isConfigured()) {
    googleSheetsSync.upsertAddress(updated).catch(e => {
      console.warn('[GoogleSheets] Address update background warning:', e.message);
    });
  }

  res.json({ success: true, address: updated });
});

app.delete('/api/addresses/:id', (req, res) => {
  const { id } = req.params;
  let addresses = readJson(addressesFile, []);
  const initialLen = addresses.length;
  addresses = addresses.filter(a => a.id !== id);

  if (addresses.length === initialLen) {
    return res.status(404).json({ success: false, error: 'Address record not found.' });
  }

  writeJson(addressesFile, addresses);
  res.json({ success: true, id });
});

// ── Master Excel Database API Endpoints (3 Sheets: Staff, Students, Addresses) ──
app.get('/api/excel/download', (req, res) => {
  try {
    const submissions = readJson(submissionsFile, []);
    const staff = readJson(staffFile, []);
    const addresses = readJson(addressesFile, []);
    const buffer = excelDb.getExcelBuffer(submissions, staff, addresses);
    if (!buffer) {
      return res.status(500).json({ success: false, error: 'Failed to generate Excel database' });
    }
    const timestamp = new Date().toISOString().slice(0, 10);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="IIT_Database_Staff_Students_Addresses_${timestamp}.xlsx"`);
    res.send(buffer);
  } catch (err) {
    console.error('[ExcelDB] Download error:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/excel/status', (req, res) => {
  try {
    const submissions = readJson(submissionsFile, []);
    const staff = readJson(staffFile, []);
    const addresses = readJson(addressesFile, []);
    let fileSize = 0;
    let lastModified = null;
    const excelTarget = fs.existsSync(excelDatabaseFile) ? excelDatabaseFile : (fs.existsSync(excelDatabasePublicFile) ? excelDatabasePublicFile : null);
    if (excelTarget) {
      const stat = fs.statSync(excelTarget);
      fileSize = stat.size;
      lastModified = stat.mtime;
    }
    res.json({
      success: true,
      fileSize,
      lastModified,
      counts: {
        staff: staff.length,
        students: submissions.length,
        addresses: addresses.length
      },
      sheets: ['Staff', 'Students', 'Addresses']
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/excel/upload', requireAdmin, (req, res) => {
  try {
    const { base64 } = req.body;
    if (!base64) {
      return res.status(400).json({ success: false, error: 'Excel file base64 data is required.' });
    }
    const buffer = Buffer.from(base64, 'base64');
    const parsed = excelDb.readFromExcelFile(buffer);
    if (!parsed) {
      return res.status(400).json({ success: false, error: 'Could not parse Excel database file.' });
    }

    let updatedStaffCount = 0;
    let updatedSubmissionsCount = 0;
    let updatedAddressesCount = 0;

    if (parsed.staff && parsed.staff.length > 0) {
      writeJson(staffFile, parsed.staff);
      updatedStaffCount = parsed.staff.length;
    }
    if (parsed.submissions && parsed.submissions.length > 0) {
      writeJson(submissionsFile, parsed.submissions);
      updatedSubmissionsCount = parsed.submissions.length;
    }
    if (parsed.addresses && parsed.addresses.length > 0) {
      writeJson(addressesFile, parsed.addresses);
      updatedAddressesCount = parsed.addresses.length;
    }

    syncDatabaseToExcel();

    res.json({
      success: true,
      imported: {
        staff: updatedStaffCount,
        students: updatedSubmissionsCount,
        addresses: updatedAddressesCount
      }
    });
  } catch (err) {
    console.error('[ExcelDB] Upload error:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/excel/sync', requireAdmin, (req, res) => {
  syncDatabaseToExcel();
  res.json({ success: true, message: 'Excel database synchronized successfully.' });
});

// ── Google Sheets / Google Drive Integration Endpoints ──
app.get('/api/google-sheets/config', requireAdmin, (req, res) => {
  try {
    const config = googleSheetsSync.getConfig();
    res.json({ success: true, config });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/google-sheets/config', requireAdmin, async (req, res) => {
  try {
    const { webAppUrl, spreadsheetUrl, autoSync, enabled } = req.body;
    const updated = googleSheetsSync.saveConfig({
      webAppUrl: (webAppUrl !== undefined) ? webAppUrl.trim() : undefined,
      spreadsheetUrl: (spreadsheetUrl !== undefined) ? spreadsheetUrl.trim() : undefined,
      autoSync: autoSync !== undefined ? !!autoSync : true,
      enabled: enabled !== undefined ? !!enabled : true
    });

    let testResult = null;
    let initialSync = null;

    if (updated.webAppUrl) {
      testResult = await googleSheetsSync.testConnection(updated.webAppUrl, updated.spreadsheetUrl);
      if (testResult.success) {
        // Automatically pull existing data into site
        const pullRes = await googleSheetsSync.pullFromGoogleSheets();
        if (pullRes.success) {
          if (pullRes.students && pullRes.students.length > 0) {
            const localSubs = readJson(submissionsFile, []);
            const merged = localSubs.length === 0 ? pullRes.students : mergeById(localSubs, pullRes.students, 'submissionId');
            fs.writeFileSync(submissionsFile, JSON.stringify(merged, null, 2), 'utf8');
          }
          if (pullRes.staff && pullRes.staff.length > 0) {
            const localStaff = readJson(staffFile, []);
            const merged = localStaff.length === 0 ? pullRes.staff : mergeById(localStaff, pullRes.staff, 'staffId');
            fs.writeFileSync(staffFile, JSON.stringify(merged, null, 2), 'utf8');
          }
          if (pullRes.addresses && pullRes.addresses.length > 0) {
            const localAddr = readJson(addressesFile, []);
            const merged = localAddr.length === 0 ? pullRes.addresses : mergeById(localAddr, pullRes.addresses, 'id');
            fs.writeFileSync(addressesFile, JSON.stringify(merged, null, 2), 'utf8');
          }
          syncDatabaseToExcel();
        }

        // Push combined dataset to Google Sheets so both have full master data
        const finalSubs = readJson(submissionsFile, []);
        const finalStaff = readJson(staffFile, []);
        const finalAddr = readJson(addressesFile, []);
        initialSync = await googleSheetsSync.pushToGoogleSheets(finalSubs, finalStaff, finalAddr);
      }
    }

    res.json({ success: true, config: updated, testResult, initialSync });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/google-sheets/test', requireAdmin, async (req, res) => {
  try {
    const { webAppUrl, spreadsheetUrl } = req.body;
    const result = await googleSheetsSync.testConnection(webAppUrl, spreadsheetUrl);
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/google-sheets/sync', requireAdmin, async (req, res) => {
  try {
    const submissions = readJson(submissionsFile, []);
    const staff = readJson(staffFile, []);
    const addresses = readJson(addressesFile, []);
    const result = await googleSheetsSync.pushToGoogleSheets(submissions, staff, addresses);
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/google-sheets/pull', requireAdmin, async (req, res) => {
  try {
    const result = await googleSheetsSync.pullFromGoogleSheets();
    if (!result.success) {
      return res.status(400).json(result);
    }

    let updatedStaffCount = 0;
    let updatedSubmissionsCount = 0;
    let updatedAddressesCount = 0;

    if (result.staff && result.staff.length > 0) {
      writeJson(staffFile, result.staff);
      updatedStaffCount = result.staff.length;
    }
    if (result.students && result.students.length > 0) {
      writeJson(submissionsFile, result.students);
      updatedSubmissionsCount = result.students.length;
    }
    if (result.addresses && result.addresses.length > 0) {
      writeJson(addressesFile, result.addresses);
      updatedAddressesCount = result.addresses.length;
    }

    syncDatabaseToExcel();

    res.json({
      success: true,
      message: 'Successfully pulled and restored data from Google Sheets',
      spreadsheetName: result.spreadsheetName,
      imported: {
        staff: updatedStaffCount,
        students: updatedSubmissionsCount,
        addresses: updatedAddressesCount
      }
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/google-sheets/script-code', requireAdmin, (req, res) => {
  res.json({ success: true, scriptCode: googleSheetsSync.getScriptCode() });
});

app.get('/address', (req, res) => {
  res.redirect('/address.html');
});

app.get('/address.html', (req, res) => {
  if (!sendIfFileExists(res, 'address.html')) {
    res.status(404).send('Address Manager page not found');
  }
});

function writeBackendPort(portNumber) {
  try {
    fs.writeFileSync(backendPortFile, String(portNumber), 'utf8');
  } catch (e) {
    // Ignore write errors (e.g. read-only filesystem on Vercel)
  }
}

function getLocalIpAddresses() {
  const interfaces = os.networkInterfaces();
  const addresses = [];
  for (const name of Object.keys(interfaces)) {
    for (const iface of interfaces[name]) {
      if (iface.family === 'IPv4' && !iface.internal) {
        addresses.push(iface.address);
      }
    }
  }
  return addresses;
}

function startServer(currentPort) {
  const server = app.listen(currentPort, '0.0.0.0', () => {
    writeBackendPort(currentPort);
    const ips = getLocalIpAddresses();
    console.log('========================================================');
    console.log('   INDIAN INFO TECH - SERVER RUNNING');
    console.log(`   Local Machine Access   : http://localhost:${currentPort}`);
    ips.forEach(ip => {
      console.log(`   Other Devices / Wi-Fi  : http://${ip}:${currentPort}`);
    });
    console.log('   Excel Database         : database.xlsx (Active)');
    console.log('   Google Sheets Status   : ' + (googleSheetsSync.isConfigured() ? 'LIVE SYNC ACTIVE' : 'Configure in Admin -> Google Sheets'));
    console.log('========================================================');
  });

  server.on('error', (error) => {
    if (error.code === 'EADDRINUSE') {
      const nextPort = currentPort + 1;
      console.warn(`Port ${currentPort} is busy. Trying ${nextPort}...`);
      startServer(nextPort);
    } else {
      throw error;
    }
  });
}

// Only start the server when run directly (e.g. node server-express.js)
// When imported by Vercel serverless, just export the app
if (require.main === module) {
  startServer(process.env.PORT || 5000);
}

module.exports = app;