const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const jwt = require('jsonwebtoken');
const Database = require('better-sqlite3');
const bcrypt = require('bcrypt');
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
const DATA_DIR = path.join(__dirname, 'data');
const UPLOADS_DIR = path.join(__dirname, 'uploads');
const DB_FILE = path.join(DATA_DIR, 'dashboard.db');
const ROOT_FOLDER_ID = 'root';

if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

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
// SQLite (files + local users)
// ---------------------------------------------------------------------------
const db = new Database(DB_FILE);
db.pragma('journal_mode = WAL');
db.exec(`
  CREATE TABLE IF NOT EXISTS files (
    id TEXT PRIMARY KEY,
    original_name TEXT NOT NULL,
    stored_name TEXT NOT NULL,
    size INTEGER NOT NULL,
    document_type TEXT NOT NULL,
    file_type TEXT NOT NULL,
    uploaded_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS users (
    id TEXT PRIMARY KEY,
    email TEXT NOT NULL UNIQUE,
    name TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    created_at TEXT NOT NULL
  );
`);

function ensureFolderSchema() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS folders (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      parent_id TEXT,
      created_at TEXT NOT NULL,
      FOREIGN KEY (parent_id) REFERENCES folders(id)
    );
  `);

  const columns = db.prepare('PRAGMA table_info(files)').all();
  const hasFolderId = columns.some((col) => col.name === 'folder_id');
  if (!hasFolderId) {
    db.exec(`ALTER TABLE files ADD COLUMN folder_id TEXT NOT NULL DEFAULT '${ROOT_FOLDER_ID}'`);
  }

  db.prepare(`
    INSERT OR IGNORE INTO folders (id, name, parent_id, created_at)
    VALUES (?, ?, NULL, ?)
  `).run(ROOT_FOLDER_ID, 'Root', new Date().toISOString());

  db.prepare("UPDATE files SET folder_id = ? WHERE folder_id IS NULL OR folder_id = ''").run(ROOT_FOLDER_ID);
}

ensureFolderSchema();

const LEGACY_META_FILE = path.join(UPLOADS_DIR, '_meta.json');
function migrateLegacyMetaIfNeeded() {
  const fileCount = db.prepare('SELECT COUNT(*) AS count FROM files').get().count;
  if (fileCount > 0 || !fs.existsSync(LEGACY_META_FILE)) return;

  try {
    const meta = JSON.parse(fs.readFileSync(LEGACY_META_FILE, 'utf8'));
    const insert = db.prepare(`
      INSERT OR IGNORE INTO files (id, original_name, stored_name, size, document_type, file_type, uploaded_at, folder_id)
      VALUES (@id, @original_name, @stored_name, @size, @document_type, @file_type, @uploaded_at, @folder_id)
    `);

    const tx = db.transaction((entries) => {
      for (const [id, m] of entries) {
        insert.run({
          id,
          original_name: m.originalName,
          stored_name: m.storedName,
          size: Number(m.size || 0),
          document_type: m.documentType,
          file_type: m.fileType,
          uploaded_at: m.uploadedAt || new Date().toISOString(),
          folder_id: ROOT_FOLDER_ID,
        });
      }
    });

    tx(Object.entries(meta));
    console.log(`[migration] Imported ${Object.keys(meta).length} file records from uploads/_meta.json into SQLite`);
  } catch (err) {
    console.error('[migration] Failed to import legacy _meta.json:', err.message);
  }
}

function toApiFile(row) {
  return {
    id: row.id,
    name: row.original_name,
    storedName: row.stored_name,
    size: row.size,
    type: row.document_type,
    uploadedAt: row.uploaded_at,
    folderId: row.folder_id,
  };
}

function toApiFolder(row) {
  return {
    id: row.id,
    name: row.name,
    parentId: row.parent_id,
    createdAt: row.created_at,
  };
}

const getFileByIdStmt = db.prepare('SELECT * FROM files WHERE id = ?');
const listFilesStmt = db.prepare('SELECT * FROM files ORDER BY uploaded_at DESC');
const listFilesByFolderStmt = db.prepare('SELECT * FROM files WHERE folder_id = ? ORDER BY uploaded_at DESC');
const insertFileStmt = db.prepare(`
  INSERT INTO files (id, original_name, stored_name, size, document_type, file_type, uploaded_at, folder_id)
  VALUES (@id, @original_name, @stored_name, @size, @document_type, @file_type, @uploaded_at, @folder_id)
`);
const deleteFileStmt = db.prepare('DELETE FROM files WHERE id = ?');
const renameFileStmt = db.prepare('UPDATE files SET original_name = ? WHERE id = ?');
const moveFileStmt = db.prepare('UPDATE files SET folder_id = ? WHERE id = ?');

const getFolderByIdStmt = db.prepare('SELECT * FROM folders WHERE id = ?');
const listFoldersByParentStmt = db.prepare(`
  SELECT * FROM folders
  WHERE ((? IS NULL AND parent_id IS NULL) OR parent_id = ?)
  ORDER BY name COLLATE NOCASE
`);
const insertFolderStmt = db.prepare(`
  INSERT INTO folders (id, name, parent_id, created_at)
  VALUES (@id, @name, @parent_id, @created_at)
`);
const renameFolderStmt = db.prepare('UPDATE folders SET name = ? WHERE id = ?');
const deleteFolderStmt = db.prepare('DELETE FROM folders WHERE id = ?');
const countFolderChildrenStmt = db.prepare('SELECT COUNT(*) AS count FROM folders WHERE parent_id = ?');
const countFolderFilesStmt = db.prepare('SELECT COUNT(*) AS count FROM files WHERE folder_id = ?');

function buildFolderPath(folderId) {
  const pathItems = [];
  let cursor = getFolderByIdStmt.get(folderId || ROOT_FOLDER_ID);
  while (cursor) {
    pathItems.push(toApiFolder(cursor));
    if (!cursor.parent_id) break;
    cursor = getFolderByIdStmt.get(cursor.parent_id);
  }
  return pathItems.reverse();
}

const getUserByEmailStmt = db.prepare('SELECT * FROM users WHERE email = ?');
const countUsersStmt = db.prepare('SELECT COUNT(*) AS count FROM users');
const insertUserStmt = db.prepare(`
  INSERT INTO users (id, email, name, password_hash, created_at)
  VALUES (@id, @email, @name, @password_hash, @created_at)
`);

function normaliseEmail(email) {
  return String(email || '').trim().toLowerCase();
}

async function ensureLocalAdminFromEnv() {
  const email = normaliseEmail(process.env.LOCAL_ADMIN_EMAIL);
  const password = process.env.LOCAL_ADMIN_PASSWORD || '';
  const name = (process.env.LOCAL_ADMIN_NAME || 'Local Admin').trim();
  if (!email || !password) return;

  const existing = getUserByEmailStmt.get(email);
  if (existing) return;

  const password_hash = await bcrypt.hash(password, 12);
  insertUserStmt.run({
    id: crypto.randomUUID(),
    email,
    name,
    password_hash,
    created_at: new Date().toISOString(),
  });
  console.log(`[auth] Seeded local admin user: ${email}`);
}

migrateLegacyMetaIfNeeded();

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
// Local auth routes (email/password with bcrypt)
// ---------------------------------------------------------------------------
app.post('/auth/local/login', async (req, res) => {
  const email = normaliseEmail(req.body?.email);
  const password = String(req.body?.password || '');
  if (!email || !password) {
    return res.status(400).json({ error: 'Email and password are required.' });
  }

  const user = getUserByEmailStmt.get(email);
  if (!user) return res.status(401).json({ error: 'Invalid email or password.' });

  const ok = await bcrypt.compare(password, user.password_hash);
  if (!ok) return res.status(401).json({ error: 'Invalid email or password.' });

  req.session.regenerate((err) => {
    if (err) return res.status(500).json({ error: 'Could not start session.' });
    req.session.user = {
      id: user.id,
      email: user.email,
      name: user.name,
      provider: 'local',
    };
    req.session.save((saveErr) => {
      if (saveErr) return res.status(500).json({ error: 'Could not save session.' });
      res.json({ ok: true });
    });
  });
});

app.post('/auth/local/register', async (req, res) => {
  const email = normaliseEmail(req.body?.email);
  const name = String(req.body?.name || '').trim() || email;
  const password = String(req.body?.password || '');
  if (!email || password.length < 8) {
    return res.status(400).json({ error: 'Valid email and a password of at least 8 characters are required.' });
  }

  const allowRegistration = process.env.ALLOW_LOCAL_REGISTRATION === 'true';
  const hasUsers = countUsersStmt.get().count > 0;
  if (hasUsers && !allowRegistration) {
    return res.status(403).json({ error: 'Local registration is disabled.' });
  }

  if (getUserByEmailStmt.get(email)) {
    return res.status(409).json({ error: 'An account with this email already exists.' });
  }

  const password_hash = await bcrypt.hash(password, 12);
  insertUserStmt.run({
    id: crypto.randomUUID(),
    email,
    name,
    password_hash,
    created_at: new Date().toISOString(),
  });

  res.status(201).json({ ok: true });
});

app.get('/auth/local/register-status', (_req, res) => {
  const allowRegistration = process.env.ALLOW_LOCAL_REGISTRATION === 'true';
  const hasUsers = countUsersStmt.get().count > 0;
  const allowed = !hasUsers || allowRegistration;
  res.json({
    allowed,
    hasUsers,
    requiresEnvFlag: hasUsers && !allowRegistration,
  });
});

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
    const entry = getFileByIdStmt.get(req.params.id);
    if (entry) {
      const httpLib = url.startsWith('https') ? require('https') : require('http');
      const filePath = path.join(UPLOADS_DIR, entry.stored_name);
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
  const folderId = String(_req.query.folderId || '').trim();
  const files = folderId
    ? listFilesByFolderStmt.all(folderId).map(toApiFile)
    : listFilesStmt.all().map(toApiFile);
  res.json(files);
});

app.get('/api/folders/:id/contents', (req, res) => {
  const folderId = req.params.id || ROOT_FOLDER_ID;
  const folder = getFolderByIdStmt.get(folderId);
  if (!folder) return res.status(404).json({ error: 'Folder not found' });

  const folders = listFoldersByParentStmt.all(folderId, folderId).map(toApiFolder);
  const files = listFilesByFolderStmt.all(folderId).map(toApiFile);
  res.json({
    folder: toApiFolder(folder),
    path: buildFolderPath(folderId),
    folders,
    files,
  });
});

app.get('/api/folders', (req, res) => {
  const rawParent = req.query.parentId;
  const parentId = typeof rawParent === 'undefined' ? ROOT_FOLDER_ID : String(rawParent || '').trim();
  const key = parentId === '' ? null : parentId;
  const folders = listFoldersByParentStmt.all(key, key).map(toApiFolder);
  res.json(folders);
});

app.post('/api/folders', (req, res) => {
  const name = String(req.body?.name || '').trim();
  const parentId = String(req.body?.parentId || ROOT_FOLDER_ID).trim();
  if (!name) return res.status(400).json({ error: 'Folder name is required' });

  const parent = getFolderByIdStmt.get(parentId);
  if (!parent) return res.status(404).json({ error: 'Parent folder not found' });

  const id = crypto.randomUUID();
  insertFolderStmt.run({
    id,
    name: name.substring(0, 120),
    parent_id: parentId,
    created_at: new Date().toISOString(),
  });
  const created = getFolderByIdStmt.get(id);
  res.status(201).json(toApiFolder(created));
});

app.patch('/api/folders/:id', (req, res) => {
  const id = req.params.id;
  if (id === ROOT_FOLDER_ID) return res.status(400).json({ error: 'Root folder cannot be renamed' });

  const name = String(req.body?.name || '').trim();
  if (!name) return res.status(400).json({ error: 'Folder name is required' });

  const folder = getFolderByIdStmt.get(id);
  if (!folder) return res.status(404).json({ error: 'Folder not found' });

  renameFolderStmt.run(name.substring(0, 120), id);
  res.json({ ok: true });
});

app.delete('/api/folders/:id', (req, res) => {
  const id = req.params.id;
  if (id === ROOT_FOLDER_ID) return res.status(400).json({ error: 'Root folder cannot be deleted' });

  const folder = getFolderByIdStmt.get(id);
  if (!folder) return res.status(404).json({ error: 'Folder not found' });

  const childFolders = countFolderChildrenStmt.get(id).count;
  const childFiles = countFolderFilesStmt.get(id).count;
  if (childFolders > 0 || childFiles > 0) {
    return res.status(409).json({ error: 'Folder is not empty' });
  }

  deleteFolderStmt.run(id);
  res.json({ ok: true });
});

// ---- API: upload file ----
app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file provided' });

  const info = extInfo(req.file.originalname);
  if (!info) return res.status(400).json({ error: 'Unsupported file type' });

  const folderId = String(req.body?.folderId || ROOT_FOLDER_ID).trim();
  const folder = getFolderByIdStmt.get(folderId);
  if (!folder) return res.status(400).json({ error: 'Invalid folder ID' });

  const id = crypto.randomUUID();
  insertFileStmt.run({
    id,
    original_name: req.file.originalname,
    stored_name: req.file.filename,
    size: req.file.size,
    document_type: info.documentType,
    file_type: info.fileType,
    uploaded_at: new Date().toISOString(),
    folder_id: folderId,
  });

  res.json({ id, name: req.file.originalname });
});

// ---- API: delete file ----
app.delete('/api/files/:id', (req, res) => {
  const entry = getFileByIdStmt.get(req.params.id);
  if (!entry) return res.status(404).json({ error: 'Not found' });

  const filePath = path.join(UPLOADS_DIR, entry.stored_name);
  if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
  deleteFileStmt.run(req.params.id);
  res.json({ ok: true });
});

// ---- API: get editor config for a file ----
app.get('/api/editor-config/:id', (req, res) => {
  const entry = getFileByIdStmt.get(req.params.id);
  if (!entry) return res.status(404).json({ error: 'Not found' });

  const appBaseUrl = resolveAppUrl(req);
  // Sign a short-lived token so OnlyOffice can fetch the file without a session
  const fileToken = jwt.sign({ id: req.params.id }, FILE_TOKEN_SECRET, { expiresIn: '2h' });
  const fileUrl = `${appBaseUrl}/uploads/${encodeURIComponent(entry.stored_name)}?t=${fileToken}`;
  const callbackUrl = `${appBaseUrl}/api/callback/${req.params.id}`;

  const sessionUser = req.session?.user;
  const editorUser = {
    id: sessionUser?.id || sessionUser?.email || 'anonymous',
    name: sessionUser?.name || sessionUser?.email || 'Anonymous User',
    group: sessionUser?.provider || 'dashboard',
  };

  const config = {
    document: {
      fileType: entry.file_type,
      key: req.params.id + '_' + Date.now(),
      title: entry.original_name,
      url: fileUrl,
      permissions: {
        download: true,
        edit: true,
        print: true,
        review: true,
      },
    },
    documentType: entry.document_type,
    editorConfig: {
      callbackUrl,
      lang: 'en',
      mode: 'edit',
      user: editorUser,
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
  const entry = getFileByIdStmt.get(req.params.id);
  if (!entry) return res.status(404).json({ error: 'Not found' });

  const appBaseUrl = resolveAppUrl(req);
  const fileToken = jwt.sign({ id: req.params.id }, FILE_TOKEN_SECRET, { expiresIn: '2h' });
  const fileUrl = `${appBaseUrl}/uploads/${encodeURIComponent(entry.stored_name)}?t=${fileToken}`;
  const callbackUrl = `${appBaseUrl}/api/callback/${req.params.id}`;
  const filePath = path.join(UPLOADS_DIR, entry.stored_name);

  res.json({
    id: req.params.id,
    onlyofficeUrl: ONLYOFFICE_URL,
    appUrlUsed: appBaseUrl,
    jwtEnabled: !!JWT_SECRET,
    fileUrl,
    callbackUrl,
    fileExists: fs.existsSync(filePath),
    fileSize: fs.existsSync(filePath) ? fs.statSync(filePath).size : 0,
    storedName: entry.stored_name,
    originalName: entry.original_name,
    documentType: entry.document_type,
    fileType: entry.file_type,
  });
});

// ---- API: create blank document ----
app.post('/api/create', (req, res) => {
  const { name, type } = req.body; // type: "docx" | "xlsx" | "pptx"
  const allowed = { docx: 'word', xlsx: 'cell', pptx: 'slide' };
  if (!allowed[type]) return res.status(400).json({ error: 'Invalid type' });

  const folderId = String(req.body?.folderId || ROOT_FOLDER_ID).trim();
  const folder = getFolderByIdStmt.get(folderId);
  if (!folder) return res.status(400).json({ error: 'Invalid folder ID' });

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

  const stat = fs.statSync(path.join(UPLOADS_DIR, storedName));
  const originalName = name || `Untitled.${type}`;
  insertFileStmt.run({
    id,
    original_name: originalName,
    stored_name: storedName,
    size: stat.size,
    document_type: allowed[type],
    file_type: type,
    uploaded_at: new Date().toISOString(),
    folder_id: folderId,
  });
  res.json({ id, name: originalName });
});

// ---- API: rename file ----
app.patch('/api/files/:id', (req, res) => {
  const { name } = req.body;
  if (!name || typeof name !== 'string') return res.status(400).json({ error: 'Name required' });

  const entry = getFileByIdStmt.get(req.params.id);
  if (!entry) return res.status(404).json({ error: 'Not found' });

  renameFileStmt.run(name.substring(0, 200), req.params.id);
  res.json({ ok: true });
});

app.patch('/api/files/:id/move', (req, res) => {
  const folderId = String(req.body?.folderId || '').trim();
  if (!folderId) return res.status(400).json({ error: 'folderId is required' });

  const file = getFileByIdStmt.get(req.params.id);
  if (!file) return res.status(404).json({ error: 'File not found' });

  const folder = getFolderByIdStmt.get(folderId);
  if (!folder) return res.status(404).json({ error: 'Folder not found' });

  moveFileStmt.run(folderId, req.params.id);
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
ensureLocalAdminFromEnv()
  .catch((err) => {
    console.error('[auth] Failed to seed local admin user:', err.message);
  })
  .finally(() => {
    app.listen(PORT, '0.0.0.0', () => {
      console.log(`OnlyOffice Dashboard running on http://0.0.0.0:${PORT}`);
      console.log(`OnlyOffice URL: ${ONLYOFFICE_URL}`);
      console.log(`App callback URL: ${APP_URL}`);
      console.log(`SQLite DB: ${DB_FILE}`);
    });
  });
