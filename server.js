const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const jwt = require('jsonwebtoken');
const session = require('express-session');
const FileStore = require('session-file-store')(session);
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------
const PORT = parseInt(process.env.PORT, 10) || 3000;
const ONLYOFFICE_URL = (process.env.ONLYOFFICE_URL || 'http://localhost:8080').replace(/\/+$/, '');
const APP_URL = (process.env.APP_URL || `http://host.docker.internal:${PORT}`).replace(/\/+$/, '');
const JWT_SECRET = process.env.JWT_SECRET || '';
const UPLOADS_DIR = path.join(__dirname, 'uploads');

if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Extension → documentType / fileType mappings
// ---------------------------------------------------------------------------
const EXT_MAP = {
  // Word
  docx: { documentType: 'word', fileType: 'docx' },
  doc:  { documentType: 'word', fileType: 'doc' },
  odt:  { documentType: 'word', fileType: 'odt' },
  rtf:  { documentType: 'word', fileType: 'rtf' },
  txt:  { documentType: 'word', fileType: 'txt' },
  pdf:  { documentType: 'word', fileType: 'pdf' },
  // Spreadsheet
  xlsx: { documentType: 'cell', fileType: 'xlsx' },
  xls:  { documentType: 'cell', fileType: 'xls' },
  ods:  { documentType: 'cell', fileType: 'ods' },
  csv:  { documentType: 'cell', fileType: 'csv' },
  // Presentation
  pptx: { documentType: 'slide', fileType: 'pptx' },
  ppt:  { documentType: 'slide', fileType: 'ppt' },
  odp:  { documentType: 'slide', fileType: 'odp' },
};

function extInfo(filename) {
  const ext = path.extname(filename).replace('.', '').toLowerCase();
  return EXT_MAP[ext] || null;
}

function validateTemplateFile(templateFile, type) {
  if (!fs.existsSync(templateFile)) {
    return { ok: false, reason: `Template blank.${type} not found` };
  }

  const buf = fs.readFileSync(templateFile);
  if (buf.length < 512) {
    return { ok: false, reason: `Template blank.${type} is too small to be a valid Office file` };
  }

  // Office Open XML files are ZIP containers and should start with PK\x03\x04.
  if (!(buf[0] === 0x50 && buf[1] === 0x4B && buf[2] === 0x03 && buf[3] === 0x04)) {
    return { ok: false, reason: `Template blank.${type} is not a valid OOXML (ZIP) file` };
  }

  const content = buf.toString('latin1');
  const baseRequired = ['[Content_Types].xml', '_rels/.rels'];
  const perType = {
    docx: ['word/document.xml'],
    xlsx: ['xl/workbook.xml'],
    pptx: ['ppt/presentation.xml'],
  };
  const required = [...baseRequired, ...(perType[type] || [])];
  const missing = required.filter((entry) => !content.includes(entry));
  if (missing.length) {
    return { ok: false, reason: `Template blank.${type} is missing required OOXML parts: ${missing.join(', ')}` };
  }

  return { ok: true };
}

function resolveAppUrl(req) {
  if (process.env.APP_URL && process.env.APP_URL.trim()) return APP_URL;
  const protoHeader = req.headers['x-forwarded-proto'];
  const forwardedProto = Array.isArray(protoHeader) ? protoHeader[0] : (protoHeader || '');
  const proto = forwardedProto.split(',')[0].trim() || req.protocol;
  const host = req.headers['x-forwarded-host'] || req.get('host');
  return `${proto}://${host}`.replace(/\/+$/, '');
}

// ---------------------------------------------------------------------------
// File-metadata store (simple JSON on disk)
// ---------------------------------------------------------------------------
const META_FILE = path.join(UPLOADS_DIR, '_meta.json');

function loadMeta() {
  if (!fs.existsSync(META_FILE)) return {};
  try { return JSON.parse(fs.readFileSync(META_FILE, 'utf8')); }
  catch { return {}; }
}

function saveMeta(meta) {
  fs.writeFileSync(META_FILE, JSON.stringify(meta, null, 2));
}

// ---------------------------------------------------------------------------
// Multer setup – sanitise filenames
// ---------------------------------------------------------------------------
const ALLOWED_EXTENSIONS = new Set(Object.keys(EXT_MAP));

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, UPLOADS_DIR),
  filename: (_req, file, cb) => {
    const safeName = file.originalname
      .replace(/[^a-zA-Z0-9._-]/g, '_')
      .substring(0, 200);
    const unique = crypto.randomUUID().slice(0, 8);
    cb(null, `${unique}_${safeName}`);
  },
});

const upload = multer({
  storage,
  limits: { fileSize: 100 * 1024 * 1024 }, // 100 MB
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname).replace('.', '').toLowerCase();
    if (!ALLOWED_EXTENSIONS.has(ext)) {
      return cb(new Error(`Unsupported file type: .${ext}`));
    }
    cb(null, true);
  },
});

// ---------------------------------------------------------------------------
// Express app
// ---------------------------------------------------------------------------
const app = express();

// Trust Cloudflare / reverse-proxy forwarded headers (X-Forwarded-Proto, etc.)
// Required for secure session cookies to work behind a Cloudflare tunnel.
app.set('trust proxy', 1);

// ---------------------------------------------------------------------------
// Security headers (helmet) — must come before routes
// ---------------------------------------------------------------------------
app.use(helmet({
  // Allow OnlyOffice iframe and inline scripts needed by the editor
  contentSecurityPolicy: false,
}));

// ---------------------------------------------------------------------------
// Rate limiting — protect auth and API endpoints
// ---------------------------------------------------------------------------
const authLimiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 20,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many requests, please try again later.' },
});
app.use('/auth/', authLimiter);

const apiLimiter = rateLimit({
  windowMs: 60 * 1000, // 1 minute
  max: 120,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many requests, please try again later.' },
});
app.use('/api/', apiLimiter);

app.use(express.json());

// ---------------------------------------------------------------------------
// Session middleware
// ---------------------------------------------------------------------------
const SESSION_DIR = path.join(__dirname, 'data', 'sessions');
if (!fs.existsSync(SESSION_DIR)) fs.mkdirSync(SESSION_DIR, { recursive: true });

app.use(session({
  store: new FileStore({ path: SESSION_DIR, ttl: 7 * 24 * 3600, retries: 1 }),
  secret: process.env.SESSION_SECRET || 'officeui-change-me-in-production',
  resave: false,
  saveUninitialized: false,
  cookie: { httpOnly: true, sameSite: 'lax', maxAge: 7 * 24 * 60 * 60 * 1000, secure: process.env.APP_URL?.startsWith('https') },
}));

// ---------------------------------------------------------------------------
// Auth helpers
// ---------------------------------------------------------------------------
/** Redirects to /login for page requests; 401 JSON for API requests. */
function requireLogin(req, res, next) {
  if (req.session?.user) return next();
  if (req.path.startsWith('/api/') || req.xhr) {
    return res.status(401).json({ error: 'Not authenticated' });
  }
  res.redirect('/login');
}

// Google OAuth2 + Gmail routes (public — handles /auth/google*)
app.use(require('./routes/google'));

// Login page (public)
app.get('/login', (req, res) => {
  if (req.session?.user) return res.redirect('/');
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// Gate the dashboard index — must come before express.static picks it up
app.get('/', requireLogin, (_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Serve static public assets (CSS, JS, etc.)
app.use(express.static(path.join(__dirname, 'public')));

// Protect file downloads — allow logged-in users OR a valid signed file token
// (OnlyOffice Document Server fetches files directly and has no session cookie)
const FILE_TOKEN_SECRET = JWT_SECRET || process.env.SESSION_SECRET || 'file-token-fallback';
app.use('/uploads', (req, res, next) => {
  if (req.session?.user) return next();
  const t = req.query.t;
  if (t) {
    try { jwt.verify(t, FILE_TOKEN_SECRET); return next(); } catch {}
  }
  res.status(401).send('Unauthorised');
}, express.static(UPLOADS_DIR));

// ---- Health / reachability check (public) ----
app.get('/ping', (_req, res) => {
  res.json({ ok: true, time: new Date().toISOString(), onlyofficeUrl: ONLYOFFICE_URL, appUrl: APP_URL });
});

// ---- API: OnlyOffice save callback (server-to-server — exempt from session auth) ----
// OnlyOffice Document Server calls this with no user session; verify via JWT if configured.
app.post('/api/callback/:id', (req, res) => {
  // If JWT is enabled, verify the Authorization header sent by OnlyOffice
  if (JWT_SECRET) {
    const authHeader = req.headers['authorization'] || '';
    const token = authHeader.replace(/^Bearer\s+/i, '');
    if (!token) return res.status(401).json({ error: 1 });
    try { jwt.verify(token, JWT_SECRET); } catch { return res.status(401).json({ error: 1 }); }
  }

  const { status, url } = req.body;
  console.log(`[callback] id=${req.params.id} status=${status} hasUrl=${!!url}`);

  if ((status === 2 || status === 6) && url) {
    const meta = loadMeta();
    const entry = meta[req.params.id];
    if (entry) {
      const httpLib = url.startsWith('https') ? require('https') : require('http');
      const filePath = path.join(UPLOADS_DIR, entry.storedName);
      httpLib.get(url, (stream) => {
        if (stream.statusCode && stream.statusCode >= 400) {
          console.error(`[callback] failed download for ${req.params.id} - HTTP ${stream.statusCode}`);
          stream.resume(); return;
        }
        const writeStream = fs.createWriteStream(filePath);
        stream.pipe(writeStream);
        writeStream.on('finish', () => writeStream.close());
        writeStream.on('error', (err) => console.error(`[callback] write error:`, err.message));
      }).on('error', (err) => console.error(`[callback] download error:`, err.message));
    } else {
      console.warn(`[callback] metadata missing for id=${req.params.id}`);
    }
  }
  res.json({ error: 0 });
});

// ---- Protect all remaining /api/* routes ----
app.use('/api', requireLogin);

// ---- API: list files ----
app.get('/api/files', (_req, res) => {
  const meta = loadMeta();
  const files = Object.entries(meta).map(([id, m]) => ({
    id,
    name: m.originalName,
    storedName: m.storedName,
    size: m.size,
    type: m.documentType,
    uploadedAt: m.uploadedAt,
  }));
  files.sort((a, b) => new Date(b.uploadedAt) - new Date(a.uploadedAt));
  res.json(files);
});

// ---- API: upload file ----
app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file provided' });

  const info = extInfo(req.file.originalname);
  if (!info) return res.status(400).json({ error: 'Unsupported file type' });

  const id = crypto.randomUUID();
  const meta = loadMeta();
  meta[id] = {
    originalName: req.file.originalname,
    storedName: req.file.filename,
    size: req.file.size,
    documentType: info.documentType,
    fileType: info.fileType,
    uploadedAt: new Date().toISOString(),
  };
  saveMeta(meta);

  res.json({ id, name: req.file.originalname });
});

// ---- API: delete file ----
app.delete('/api/files/:id', (req, res) => {
  const meta = loadMeta();
  const entry = meta[req.params.id];
  if (!entry) return res.status(404).json({ error: 'Not found' });

  const filePath = path.join(UPLOADS_DIR, entry.storedName);
  if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
  delete meta[req.params.id];
  saveMeta(meta);
  res.json({ ok: true });
});

// ---- API: get editor config for a file ----
app.get('/api/editor-config/:id', (req, res) => {
  const meta = loadMeta();
  const entry = meta[req.params.id];
  if (!entry) return res.status(404).json({ error: 'Not found' });

  const appBaseUrl = resolveAppUrl(req);
  // Sign a short-lived token so OnlyOffice can fetch the file without a session
  const fileToken = jwt.sign({ id: req.params.id }, FILE_TOKEN_SECRET, { expiresIn: '2h' });
  const fileUrl = `${appBaseUrl}/uploads/${encodeURIComponent(entry.storedName)}?t=${fileToken}`;
  const callbackUrl = `${appBaseUrl}/api/callback/${req.params.id}`;

  const config = {
    document: {
      fileType: entry.fileType,
      key: req.params.id + '_' + Date.now(),
      title: entry.originalName,
      url: fileUrl,
      permissions: {
        download: true,
        edit: true,
        print: true,
        review: true,
      },
    },
    documentType: entry.documentType,
    editorConfig: {
      callbackUrl,
      lang: 'en',
      mode: 'edit',
      customization: {
        autosave: true,
        forcesave: true,
      },
    },
  };

  // Sign with JWT – OnlyOffice uses the token both in the config object
  // and expects the callback to carry it in the Authorization header.
  if (JWT_SECRET) {
    config.token = jwt.sign(config, JWT_SECRET);
  }

  console.log(`[editor-config] id=${req.params.id} fileUrl=${fileUrl} callbackUrl=${callbackUrl} jwt=${JWT_SECRET ? 'on' : 'off'}`);

  res.json({
    config,
    onlyofficeUrl: ONLYOFFICE_URL,
    jwtEnabled: !!JWT_SECRET,
  });
});

// ---- API: debug a file integration ----
app.get('/api/debug/:id', (req, res) => {
  const meta = loadMeta();
  const entry = meta[req.params.id];
  if (!entry) return res.status(404).json({ error: 'Not found' });

  const appBaseUrl = resolveAppUrl(req);
  const fileToken = jwt.sign({ id: req.params.id }, FILE_TOKEN_SECRET, { expiresIn: '2h' });
  const fileUrl = `${appBaseUrl}/uploads/${encodeURIComponent(entry.storedName)}?t=${fileToken}`;
  const callbackUrl = `${appBaseUrl}/api/callback/${req.params.id}`;
  const filePath = path.join(UPLOADS_DIR, entry.storedName);

  res.json({
    id: req.params.id,
    onlyofficeUrl: ONLYOFFICE_URL,
    appUrlUsed: appBaseUrl,
    jwtEnabled: !!JWT_SECRET,
    fileUrl,
    callbackUrl,
    fileExists: fs.existsSync(filePath),
    fileSize: fs.existsSync(filePath) ? fs.statSync(filePath).size : 0,
    storedName: entry.storedName,
    originalName: entry.originalName,
    documentType: entry.documentType,
    fileType: entry.fileType,
  });
});

// ---- API: create blank document ----
app.post('/api/create', (req, res) => {
  const { name, type } = req.body; // type: "docx" | "xlsx" | "pptx"
  const allowed = { docx: 'word', xlsx: 'cell', pptx: 'slide' };
  if (!allowed[type]) return res.status(400).json({ error: 'Invalid type' });

  const templateDir = path.join(__dirname, 'templates');
  const templateFile = path.join(templateDir, `blank.${type}`);
  const templateValidation = validateTemplateFile(templateFile, type);
  if (!templateValidation.ok) {
    return res.status(500).json({
      error: `${templateValidation.reason}. Replace templates/blank.${type} with a real blank Office file created by Excel/Word/PowerPoint (or LibreOffice/OnlyOffice) and retry.`,
    });
  }

  const id = crypto.randomUUID();
  const safeName = (name || `Untitled.${type}`).replace(/[^a-zA-Z0-9._-]/g, '_').substring(0, 200);
  const storedName = `${crypto.randomUUID().slice(0, 8)}_${safeName}`;

  fs.copyFileSync(templateFile, path.join(UPLOADS_DIR, storedName));

  const meta = loadMeta();
  const stat = fs.statSync(path.join(UPLOADS_DIR, storedName));
  meta[id] = {
    originalName: name || `Untitled.${type}`,
    storedName,
    size: stat.size,
    documentType: allowed[type],
    fileType: type,
    uploadedAt: new Date().toISOString(),
  };
  saveMeta(meta);
  res.json({ id, name: meta[id].originalName });
});

// ---- API: rename file ----
app.patch('/api/files/:id', (req, res) => {
  const { name } = req.body;
  if (!name || typeof name !== 'string') return res.status(400).json({ error: 'Name required' });

  const meta = loadMeta();
  const entry = meta[req.params.id];
  if (!entry) return res.status(404).json({ error: 'Not found' });

  entry.originalName = name.substring(0, 200);
  saveMeta(meta);
  res.json({ ok: true });
});

// ---- Serve editor page (protected) ----
app.get('/editor/:id', requireLogin, (_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'editor.html'));
});

// ---- Global error handler ----
// eslint-disable-next-line no-unused-vars
app.use((err, req, res, _next) => {
  console.error('[error]', err.message);
  const status = err.status || err.statusCode || 500;
  if (req.path.startsWith('/api/') || req.xhr) {
    return res.status(status).json({ error: err.message || 'Internal server error' });
  }
  res.status(status).send('Server error — please go back and try again.');
});

// ---- Start ----
if (!process.env.SESSION_SECRET) {
  console.warn('⚠️  WARNING: SESSION_SECRET is not set. Using an insecure default. Set SESSION_SECRET in your environment before deploying.');
}
app.listen(PORT, '0.0.0.0', () => {
  console.log(`OnlyOffice Dashboard running on http://0.0.0.0:${PORT}`);
  console.log(`OnlyOffice URL: ${ONLYOFFICE_URL}`);
  console.log(`App callback URL: ${APP_URL}`);
});
