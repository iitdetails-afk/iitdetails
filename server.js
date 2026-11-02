const http = require('http');
const fs = require('fs');
const path = require('path');
const https = require('https');
const crypto = require('crypto');

const port = process.env.PORT || 3233;
const rootDir = __dirname;
const dataDir = path.join(rootDir, 'data');
const submissionsFile = path.join(dataDir, 'submissions.json');
const detailsFile = path.join(dataDir, 'details.json');
const adminConfigFile = path.join(dataDir, 'admin_config.json');

// ── Single admin session (in-memory) ──────────────────────────────────
let adminSession = { token: null, loginAt: null };

function generateToken() {
  return crypto.randomBytes(32).toString('hex');
}

function getAdminPin() {
  const cfg = readJson(adminConfigFile, { pin: '1234' });
  return cfg.pin || '1234';
}

function isValidToken(token) {
  return token && adminSession.token && token === adminSession.token;
}

function getAuthToken(req) {
  const auth = req.headers['authorization'] || '';
  if (auth.startsWith('Bearer ')) return auth.slice(7).trim();
  return '';
}

// ─────────────────────────────────────────────────────────────────────

const mimeTypes = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.txt': 'text/plain; charset=utf-8'
};

function ensureDataFiles() {
  if (!fs.existsSync(dataDir)) {
    fs.mkdirSync(dataDir, { recursive: true });
  }
  if (!fs.existsSync(submissionsFile)) {
    fs.writeFileSync(submissionsFile, '[]', 'utf8');
  }
  if (!fs.existsSync(detailsFile)) {
    fs.writeFileSync(detailsFile, '{}', 'utf8');
  }

  if (!fs.existsSync(adminConfigFile)) {
    fs.writeFileSync(adminConfigFile, JSON.stringify({ pin: '1234' }, null, 2), 'utf8');
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
}

function sendJson(res, statusCode, payload) {
  const cors = {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*'
  };
  res.writeHead(statusCode, cors);
  res.end(JSON.stringify(payload));
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', (chunk) => { body += chunk; });
    req.on('end', () => {
      try { resolve(body ? JSON.parse(body) : {}); }
      catch (err) { reject(err); }
    });
    req.on('error', reject);
  });
}

function serveStaticFile(res, filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const contentType = mimeTypes[ext] || 'application/octet-stream';
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('File not found');
      return;
    }
    res.writeHead(200, { 'Content-Type': contentType });
    res.end(data);
  });
}



function getNextSubmissionId(submissions) {
  let maxId = 0;
  submissions.forEach(s => {
    if (s.submissionId && s.submissionId.startsWith('ID-')) {
      const num = parseInt(s.submissionId.substring(3), 10);
      if (!isNaN(num) && num > maxId) maxId = num;
    }
  });
  return `ID-${String(maxId + 1).padStart(4, '0')}`;
}

// ─────────────────────────────────────────────────────────────────────

const server = http.createServer(async (req, res) => {
  ensureDataFiles();

  // Handle preflight
  if (req.method === 'OPTIONS') {
    res.writeHead(204, { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET,POST,OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type,Authorization' });
    res.end();
    return;
  }

  const requestUrl = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  const pathname = requestUrl.pathname;

  if (pathname === '/') {
    res.writeHead(302, { Location: '/index.html' });
    res.end();
    return;
  }

  // ── Health ──
  if (pathname === '/api/health') {
    sendJson(res, 200, { status: 'ok' });
    return;
  }

  // ── ADMIN AUTH ─────────────────────────────────────────────────────

  // POST /api/admin/login — validate PIN, issue token (single session)
  if (pathname === '/api/admin/login' && req.method === 'POST') {
    try {
      const { pin } = await readBody(req);
      const correctPin = getAdminPin();
      if (!pin || pin !== correctPin) {
        sendJson(res, 401, { success: false, error: 'Incorrect PIN' });
        return;
      }
      // Invalidate any previous session, issue new token
      const token = generateToken();
      adminSession = { token, loginAt: new Date().toISOString() };
      console.log(`[Admin] New session started at ${adminSession.loginAt}`);
      sendJson(res, 200, { success: true, token });
    } catch {
      sendJson(res, 400, { success: false, error: 'Bad request' });
    }
    return;
  }

  // POST /api/admin/logout — clear session
  if (pathname === '/api/admin/logout' && req.method === 'POST') {
    adminSession = { token: null, loginAt: null };
    console.log('[Admin] Session cleared');
    sendJson(res, 200, { success: true });
    return;
  }

  // GET /api/admin/check — verify token is still active
  if (pathname === '/api/admin/check' && req.method === 'GET') {
    const token = getAuthToken(req);
    sendJson(res, 200, { valid: isValidToken(token) });
    return;
  }

  // POST /api/admin/pin — change PIN (requires valid token)
  if (pathname === '/api/admin/pin' && req.method === 'POST') {
    const token = getAuthToken(req);
    if (!isValidToken(token)) {
      sendJson(res, 401, { success: false, error: 'Unauthorized' });
      return;
    }
    try {
      const { currentPin, newPin } = await readBody(req);
      if (currentPin !== getAdminPin()) {
        sendJson(res, 401, { success: false, error: 'Current PIN is incorrect' });
        return;
      }
      if (!newPin || !/^\d{4}$/.test(newPin)) {
        sendJson(res, 400, { success: false, error: 'New PIN must be 4 digits' });
        return;
      }
      writeJson(adminConfigFile, { pin: newPin });
      sendJson(res, 200, { success: true });
    } catch {
      sendJson(res, 400, { success: false, error: 'Bad request' });
    }
    return;
  }

  // ── PROTECTED ADMIN APIs (require valid token) ──────────────────────

  const protectedAdminRoutes = [
    { path: '/api/delete', method: 'POST' },
  ];

  const isProtected = protectedAdminRoutes.some(r => pathname === r.path && req.method === r.method);
  if (isProtected) {
    const token = getAuthToken(req);
    if (!isValidToken(token)) {
      sendJson(res, 401, { success: false, error: 'Unauthorized — admin login required' });
      return;
    }
  }

  // ── SUBMISSIONS ────────────────────────────────────────────────────

  if (pathname === '/api/submit' && req.method === 'POST') {
    try {
      const payload = await readBody(req);
      const submissions = readJson(submissionsFile, []);

      if (payload.submissionId && payload.submissionId !== 'N/A' && payload.submissionId !== '') {
        const existingIndex = submissions.findIndex(s => s.submissionId === payload.submissionId);
        if (existingIndex !== -1) {
          submissions[existingIndex] = {
            ...submissions[existingIndex],
            ...payload,
            submittedAt: payload.submittedAt || submissions[existingIndex].submittedAt || new Date().toLocaleString()
          };
          writeJson(submissionsFile, submissions);
          sendJson(res, 200, { success: true, saved: submissions[existingIndex] });
          return;
        }
      }

      const generatedId = getNextSubmissionId(submissions);
      const newPayload = {
        ...payload,
        submissionId: (payload.submissionId && payload.submissionId !== 'N/A' && payload.submissionId !== '')
          ? payload.submissionId : generatedId,
        submittedAt: payload.submittedAt || new Date().toLocaleString()
      };
      submissions.push(newPayload);
      writeJson(submissionsFile, submissions);
      sendJson(res, 200, { success: true, saved: newPayload });
    } catch {
      sendJson(res, 400, { success: false, error: 'Invalid JSON body' });
    }
    return;
  }



  if (pathname === '/api/delete' && req.method === 'POST') {
    try {
      const payload = await readBody(req);
      const submissionId = payload.submissionId;
      const submissions = readJson(submissionsFile, []);
      const filtered = submissions.filter(s => s.submissionId !== submissionId);
      writeJson(submissionsFile, filtered);
      sendJson(res, 200, { success: true });
    } catch {
      sendJson(res, 400, { success: false, error: 'Invalid JSON body' });
    }
    return;
  }

  if (pathname === '/api/submissions' && req.method === 'GET') {
    sendJson(res, 200, readJson(submissionsFile, []));
    return;
  }

  if (pathname === '/api/details' && req.method === 'POST') {
    try {
      const payload = await readBody(req);
      writeJson(detailsFile, payload);
      sendJson(res, 200, { success: true, saved: payload });
    } catch {
      sendJson(res, 400, { success: false, error: 'Invalid JSON body' });
    }
    return;
  }

  if (pathname === '/api/details' && req.method === 'GET') {
    sendJson(res, 200, readJson(detailsFile, {}));
    return;
  }

  // ── Static files ───────────────────────────────────────────────────
  const safePath = path.normalize(path.join(rootDir, pathname.replace(/^\/+/, '')));
  if (!safePath.startsWith(rootDir)) {
    res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('Forbidden');
    return;
  }

  if (fs.existsSync(safePath) && fs.statSync(safePath).isFile()) {
    serveStaticFile(res, safePath);
  } else {
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('Page not found');
  }
});

// Ensure data files exist before starting
ensureDataFiles();

server.listen(port, () => {
  console.log(`Node server running at http://localhost:${port}`);
  console.log(`Admin PIN: ${readJson(adminConfigFile, { pin: '1234' }).pin}`);
});
