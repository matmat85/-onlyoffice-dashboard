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
const DATA_DIR = path.join(__dirname, 'data');
const UPLOADS_DIR = path.join(__dirname, 'uploads');
const DB_FILE = path.join(DATA_DIR, 'dashboard.db');
const ROOT_FOLDER_ID = 'root';
const SPACE_ROOT_IDS = {
  private: 'root-private',
  shared: ROOT_FOLDER_ID,
  library: 'root-library',
};
const TRASH_FOLDER_IDS = {
  private: 'root-private-trash',
  shared: 'root-trash',
  library: 'root-library-trash',
};
const PROTECTED_ROOTS = new Set([...Object.values(SPACE_ROOT_IDS), ...Object.values(TRASH_FOLDER_IDS)]);
const RUNTIME_CONFIG_FILE = path.join(DATA_DIR, 'runtime-config.json');

if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

function loadRuntimeConfig() {
  if (!fs.existsSync(RUNTIME_CONFIG_FILE)) return {};
  try {
    const parsed = JSON.parse(fs.readFileSync(RUNTIME_CONFIG_FILE, 'utf8'));
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    return {};
  }
}

function saveRuntimeConfig(updates) {
  const current = loadRuntimeConfig();
  const next = { ...current, ...updates };
  fs.writeFileSync(RUNTIME_CONFIG_FILE, JSON.stringify(next, null, 2));
  return next;
}

function applyRuntimeConfigToEnv() {
  const runtime = loadRuntimeConfig();
  const keys = [
    'APP_URL',
    'ONLYOFFICE_URL',
    'GOOGLE_CLIENT_ID',
    'GOOGLE_CLIENT_SECRET',
    'SESSION_SECRET',
    'JWT_SECRET',
    'ALLOWED_EMAILS',
    'ADMIN_EMAILS',
  ];
  for (const key of keys) {
    if (runtime[key] !== undefined && runtime[key] !== null) {
      process.env[key] = String(runtime[key]);
    }
  }
}

// Apply persisted overrides from the data volume before reading config vars.
applyRuntimeConfigToEnv();

let ONLYOFFICE_URL = (process.env.ONLYOFFICE_URL || 'http://localhost:8080').replace(/\/+$/, '');
let APP_URL = (process.env.APP_URL || `http://host.docker.internal:${PORT}`).replace(/\/+$/, '');
let JWT_SECRET = process.env.JWT_SECRET || '';

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

function ensureAdminColumn() {
  const columns = db.prepare('PRAGMA table_info(users)').all();
  const hasIsAdmin = columns.some((col) => col.name === 'is_admin');
  if (!hasIsAdmin) {
    db.exec('ALTER TABLE users ADD COLUMN is_admin INTEGER NOT NULL DEFAULT 0');
  }
}
ensureAdminColumn();

function ensureSpaceColumns() {
  const fileColumns = db.prepare('PRAGMA table_info(files)').all();
  if (!fileColumns.some((c) => c.name === 'space'))
    db.exec("ALTER TABLE files ADD COLUMN space TEXT NOT NULL DEFAULT 'shared'");
  if (!fileColumns.some((c) => c.name === 'owner_email'))
    db.exec("ALTER TABLE files ADD COLUMN owner_email TEXT NOT NULL DEFAULT ''");

  const folderColumns = db.prepare('PRAGMA table_info(folders)').all();
  if (!folderColumns.some((c) => c.name === 'space'))
    db.exec("ALTER TABLE folders ADD COLUMN space TEXT NOT NULL DEFAULT 'shared'");

  // Ensure the 3 space root folders exist
  db.prepare("UPDATE folders SET space = 'shared' WHERE id = 'root'").run();
  db.prepare("INSERT OR IGNORE INTO folders (id, name, parent_id, created_at, space) VALUES ('root-private', 'My Files', NULL, ?, 'private')")
    .run(new Date().toISOString());
  db.prepare("INSERT OR IGNORE INTO folders (id, name, parent_id, created_at, space) VALUES ('root-library', 'Business Library', NULL, ?, 'library')")
    .run(new Date().toISOString());
  db.prepare("INSERT OR IGNORE INTO folders (id, name, parent_id, created_at, space) VALUES ('root-trash', 'Team Shared Bin', NULL, ?, 'shared')")
    .run(new Date().toISOString());
  db.prepare("INSERT OR IGNORE INTO folders (id, name, parent_id, created_at, space) VALUES ('root-private-trash', 'My Files Bin', NULL, ?, 'private')")
    .run(new Date().toISOString());
  db.prepare("INSERT OR IGNORE INTO folders (id, name, parent_id, created_at, space) VALUES ('root-library-trash', 'Business Library Bin', NULL, ?, 'library')")
    .run(new Date().toISOString());
}
ensureSpaceColumns();

function ensureFileLifecycleColumns() {
  const fileColumns = db.prepare('PRAGMA table_info(files)').all();
  if (!fileColumns.some((c) => c.name === 'previous_folder_id')) {
    db.exec('ALTER TABLE files ADD COLUMN previous_folder_id TEXT');
  }
  if (!fileColumns.some((c) => c.name === 'deleted_at')) {
    db.exec('ALTER TABLE files ADD COLUMN deleted_at TEXT');
  }
}
ensureFileLifecycleColumns();

function parseEmailList(raw) {
  return String(raw || '')
    .split(',')
    .map((e) => e.trim().toLowerCase())
    .filter(Boolean);
}

function listToCsv(list) {
  return (Array.isArray(list) ? list : []).join(',');
}

function normaliseUrl(raw, fallback) {
  const value = String(raw || '').trim();
  return (value || fallback).replace(/\/+$/, '');
}

// ALLOWED_EMAILS: comma-separated list of emails that can access Shared.
// If unset, all authenticated users can access Shared and Library.
let ALLOWED_EMAILS_LIST = parseEmailList(process.env.ALLOWED_EMAILS);
let ALLOWED_EMAIL_SET = ALLOWED_EMAILS_LIST.length ? new Set(ALLOWED_EMAILS_LIST) : null;

// ADMIN_EMAILS: comma-separated list of emails that can access admin features.
let ADMIN_EMAILS_LIST = parseEmailList(process.env.ADMIN_EMAILS);
let ADMIN_EMAIL_SET = ADMIN_EMAILS_LIST.length ? new Set(ADMIN_EMAILS_LIST) : null;

function refreshAccessLists() {
  ALLOWED_EMAILS_LIST = parseEmailList(process.env.ALLOWED_EMAILS);
  ALLOWED_EMAIL_SET = ALLOWED_EMAILS_LIST.length ? new Set(ALLOWED_EMAILS_LIST) : null;

  ADMIN_EMAILS_LIST = parseEmailList(process.env.ADMIN_EMAILS);
  ADMIN_EMAIL_SET = ADMIN_EMAILS_LIST.length ? new Set(ADMIN_EMAILS_LIST) : null;
}

function isAllowedForShared(email) {
  if (!ALLOWED_EMAIL_SET) return true;
  return ALLOWED_EMAIL_SET.has(String(email || '').toLowerCase());
}

function isAdminByEmail(email) {
  if (!ADMIN_EMAIL_SET) return false;
  return ADMIN_EMAIL_SET.has(String(email || '').toLowerCase());
}

function isAdminRequest(req) {
  const email = req.session?.user?.email || '';
  if (isAdminByEmail(email)) return true;

  // Also honour legacy DB admin role so existing admins are not locked out.
  const dbUser = getUserByIdStmt.get(req.session?.user?.id || '');
  return !!dbUser?.is_admin;
}

function maskSecret(value) {
  const str = String(value || '');
  if (!str) return '';
  if (str.length <= 4) return '*'.repeat(str.length);
  return `${'*'.repeat(str.length - 4)}${str.slice(-4)}`;
}

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
    previousFolderId: row.previous_folder_id || null,
    deletedAt: row.deleted_at || null,
    space: row.space || 'shared',
    ownerEmail: row.owner_email || '',
  };
}

function toApiFolder(row) {
  return {
    id: row.id,
    name: row.name,
    parentId: row.parent_id,
    createdAt: row.created_at,
    space: row.space || 'shared',
    isTrash: PROTECTED_ROOTS.has(row.id) && Object.values(TRASH_FOLDER_IDS).includes(row.id),
  };
}

const getFileByIdStmt = db.prepare('SELECT * FROM files WHERE id = ?');
const listFilesStmt = db.prepare('SELECT * FROM files ORDER BY uploaded_at DESC');
const listFilesByFolderStmt = db.prepare('SELECT * FROM files WHERE folder_id = ? ORDER BY uploaded_at DESC');
const insertFileStmt = db.prepare(`
  INSERT INTO files (id, original_name, stored_name, size, document_type, file_type, uploaded_at, folder_id, space, owner_email)
  VALUES (@id, @original_name, @stored_name, @size, @document_type, @file_type, @uploaded_at, @folder_id, @space, @owner_email)
`);
const deleteFileStmt = db.prepare('DELETE FROM files WHERE id = ?');
const renameFileStmt = db.prepare('UPDATE files SET original_name = ? WHERE id = ?');
const updateFileLocationStmt = db.prepare(`
  UPDATE files
  SET folder_id = @folder_id,
      previous_folder_id = @previous_folder_id,
      deleted_at = @deleted_at
  WHERE id = @id
`);

const getFolderByIdStmt = db.prepare('SELECT * FROM folders WHERE id = ?');
const listFoldersByParentStmt = db.prepare(`
  SELECT * FROM folders
  WHERE ((? IS NULL AND parent_id IS NULL) OR parent_id = ?)
  ORDER BY name COLLATE NOCASE
`);
const insertFolderStmt = db.prepare(`
  INSERT INTO folders (id, name, parent_id, created_at, space)
  VALUES (@id, @name, @parent_id, @created_at, @space)
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

function isTrashFolderId(folderId) {
  return Object.values(TRASH_FOLDER_IDS).includes(folderId);
}

function getTrashFolderIdForSpace(space) {
  return TRASH_FOLDER_IDS[space] || TRASH_FOLDER_IDS.shared;
}

function getRootFolderIdForSpace(space) {
  return SPACE_ROOT_IDS[space] || SPACE_ROOT_IDS.shared;
}

const getUserByEmailStmt = db.prepare('SELECT * FROM users WHERE email = ?');
const getUserByIdStmt = db.prepare('SELECT id, email, name, is_admin, created_at FROM users WHERE id = ?');
const countUsersStmt = db.prepare('SELECT COUNT(*) AS count FROM users');
const insertUserStmt = db.prepare(`
  INSERT INTO users (id, email, name, password_hash, created_at)
  VALUES (@id, @email, @name, @password_hash, @created_at)
`);
const listUsersAdminStmt = db.prepare('SELECT id, email, name, is_admin, created_at FROM users ORDER BY created_at ASC');
const deleteUserStmt = db.prepare('DELETE FROM users WHERE id = ?');
const setUserAdminStmt = db.prepare('UPDATE users SET is_admin = ? WHERE id = ?');
const setUserNameStmt = db.prepare('UPDATE users SET name = ? WHERE id = ?');

function canReadSpace(req, space, ownerEmail = '') {
  const userEmail = req.session?.user?.email || '';
  if (space === 'private') return !ownerEmail || ownerEmail === userEmail;
  if (space === 'shared') return isAllowedForShared(userEmail);
  return true;
}

function canManageSpace(req, space, ownerEmail = '') {
  const userEmail = req.session?.user?.email || '';
  if (space === 'private') return ownerEmail === userEmail;
  if (space === 'shared') return isAllowedForShared(userEmail);
  if (space === 'library') return ownerEmail === userEmail || isAdminRequest(req);
  return false;
}

function getFileReadError(req, entry) {
  if (!entry) return 'File not found';
  if (!canReadSpace(req, entry.space || 'shared', entry.owner_email || '')) {
    return 'Access denied.';
  }
  return null;
}

function getFileManageError(req, entry) {
  if (!entry) return 'File not found';
  if (!canManageSpace(req, entry.space || 'shared', entry.owner_email || '')) {
    return entry.space === 'library'
      ? 'Only admins or the file owner can manage library files.'
      : 'Access denied.';
  }
  return null;
}

function getFolderAccessError(req, folder, ownerEmail = '') {
  if (!folder) return 'Folder not found';
  if (!canReadSpace(req, folder.space || 'shared', ownerEmail)) {
    return 'Access denied.';
  }
  return null;
}

function updateFileLocation(entry, folderId, previousFolderId = null, deletedAt = null) {
  updateFileLocationStmt.run({
    id: entry.id,
    folder_id: folderId,
    previous_folder_id: previousFolderId,
    deleted_at: deletedAt,
  });
}

function normaliseEmail(email) {
  return String(email || '').trim().toLowerCase();
}

async function ensureLocalAdminFromEnv() {
  const email = normaliseEmail(process.env.LOCAL_ADMIN_EMAIL);
  const password = process.env.LOCAL_ADMIN_PASSWORD || '';
  const name = (process.env.LOCAL_ADMIN_NAME || 'Local Admin').trim();
  if (!email || !password) return;

  const existing = getUserByEmailStmt.get(email);
  if (existing) {
    if (!existing.is_admin) {
      setUserAdminStmt.run(1, existing.id);
      console.log(`[auth] Promoted existing user to admin: ${email}`);
    }
    return;
  }

  const password_hash = await bcrypt.hash(password, 12);
  const newId = crypto.randomUUID();
  insertUserStmt.run({
    id: newId,
    email,
    name,
    password_hash,
    created_at: new Date().toISOString(),
  });
  setUserAdminStmt.run(1, newId);
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

/** Returns 403 if the authenticated user is not an admin. */
function requireAdmin(req, res, next) {
  if (!isAdminRequest(req)) return res.status(403).json({ error: 'Admin access required' });
  next();
}

// Google OAuth2 + Gmail routes (public — handles /auth/google*)
app.use(require('./routes/google'));

// Login page (public)
app.get('/login', (req, res) => {
  if (req.session?.user) return res.redirect('/');
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// Unified logout for local and Google-authenticated sessions.
app.post('/auth/logout', (req, res) => {
  req.session.destroy(() => {
    res.clearCookie('connect.sid');
    res.json({ ok: true });
  });
});

app.get('/auth/status', (req, res) => {
  const user = req.session?.user;
  if (!user) return res.json({ authenticated: false });
  res.json({
    authenticated: true,
    email: user.email,
    name: user.name,
    provider: user.provider || 'local',
    isAdmin: isAdminRequest(req),
  });
});

// Admin panel (protected — must come before express.static)
app.get('/admin', requireLogin, requireAdmin, (_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});

// Gate the dashboard index — must come before express.static picks it up
app.get('/', requireLogin, (_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Serve static public assets (CSS, JS, etc.)
app.use(express.static(path.join(__dirname, 'public')));

// Protect file downloads — allow logged-in users OR a valid signed file token
// (OnlyOffice Document Server fetches files directly and has no session cookie)
let FILE_TOKEN_SECRET = JWT_SECRET || process.env.SESSION_SECRET || 'file-token-fallback';
app.use('/uploads', (req, res, next) => {
  if (req.session?.user) return next();

  // Preferred for direct links and browser probes.
  const t = req.query.t;
  if (t) {
    try { jwt.verify(t, FILE_TOKEN_SECRET); return next(); } catch {}
  }

  // Fallback for OnlyOffice when outbox JWT is sent in Authorization
  // and intermediary proxies alter/drop query parameters.
  const authHeader = String(req.headers['authorization'] || '');
  const bearer = authHeader.replace(/^Bearer\s+/i, '').trim();
  if (bearer) {
    try { jwt.verify(bearer, JWT_SECRET || FILE_TOKEN_SECRET); return next(); } catch {}
  }

  const requestPath = req.originalUrl || req.url;
  console.warn(`[uploads] unauthorised request path=${requestPath} ua=${req.get('user-agent') || 'n/a'}`);
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
app.get('/api/files', (req, res) => {
  const folderId = String(req.query.folderId || '').trim();
  const userEmail = req.session.user.email;
  let files = folderId
    ? listFilesByFolderStmt.all(folderId).map(toApiFile)
    : listFilesStmt.all().map(toApiFile);
  // Enforce space visibility
  files = files.filter((f) => {
    if (f.space === 'private') return f.ownerEmail === userEmail;
    if (f.space === 'shared') return isAllowedForShared(userEmail);
    return true; // library: all authenticated
  });
  res.json(files);
});

app.get('/api/folders/:id/contents', (req, res) => {
  const folderId = req.params.id || ROOT_FOLDER_ID;
  const folder = getFolderByIdStmt.get(folderId);
  if (!folder) return res.status(404).json({ error: 'Folder not found' });

  const space = folder.space || 'shared';
  const userEmail = req.session.user.email;

  const folderError = getFolderAccessError(req, folder);
  if (folderError) {
    return res.status(403).json({ error: folderError });
  }

  const folders = listFoldersByParentStmt.all(folderId, folderId).map(toApiFolder);
  let files = listFilesByFolderStmt.all(folderId).map(toApiFile);

  // Private: only show owner's files
  if (space === 'private') {
    files = files.filter((f) => f.ownerEmail === userEmail);
  }

  res.json({
    folder: toApiFolder(folder),
    path: buildFolderPath(folderId),
    folders,
    files,
    space,
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
  if (isTrashFolderId(parentId)) return res.status(400).json({ error: 'Cannot create folders inside the bin' });

  // Folders inherit the space of their parent
  const space = parent.space || 'shared';

  const id = crypto.randomUUID();
  insertFolderStmt.run({
    id,
    name: name.substring(0, 120),
    parent_id: parentId,
    created_at: new Date().toISOString(),
    space,
  });
  const created = getFolderByIdStmt.get(id);
  res.status(201).json(toApiFolder(created));
});

app.patch('/api/folders/:id', (req, res) => {
  const id = req.params.id;
  if (PROTECTED_ROOTS.has(id)) return res.status(400).json({ error: 'Root folders cannot be renamed' });

  const name = String(req.body?.name || '').trim();
  if (!name) return res.status(400).json({ error: 'Folder name is required' });

  const folder = getFolderByIdStmt.get(id);
  if (!folder) return res.status(404).json({ error: 'Folder not found' });

  renameFolderStmt.run(name.substring(0, 120), id);
  res.json({ ok: true });
});

app.delete('/api/folders/:id', (req, res) => {
  const id = req.params.id;
  if (PROTECTED_ROOTS.has(id)) return res.status(400).json({ error: 'Root folders cannot be deleted' });

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
  if (isTrashFolderId(folderId)) return res.status(400).json({ error: 'Cannot upload directly into the bin' });

  const id = crypto.randomUUID();
  const VALID_SPACES = new Set(['private', 'shared', 'library']);
  const space = VALID_SPACES.has(req.body?.space) ? req.body.space : (folder.space || 'shared');
  insertFileStmt.run({
    id,
    original_name: req.file.originalname,
    stored_name: req.file.filename,
    size: req.file.size,
    document_type: info.documentType,
    file_type: info.fileType,
    uploaded_at: new Date().toISOString(),
    folder_id: folderId,
    space,
    owner_email: req.session.user.email,
  });

  res.json({ id, name: req.file.originalname });
});

// ---- API: delete file ----
app.delete('/api/files/:id', (req, res) => {
  const entry = getFileByIdStmt.get(req.params.id);
  if (!entry) return res.status(404).json({ error: 'Not found' });

  const manageError = getFileManageError(req, entry);
  if (manageError) return res.status(403).json({ error: manageError });

  if (!isTrashFolderId(entry.folder_id)) {
    updateFileLocation(entry, getTrashFolderIdForSpace(entry.space || 'shared'), entry.folder_id, new Date().toISOString());
    return res.json({ ok: true, trashed: true });
  }

  const filePath = path.join(UPLOADS_DIR, entry.stored_name);
  if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
  deleteFileStmt.run(req.params.id);
  res.json({ ok: true, deleted: true });
});

// ---- API: get editor config for a file ----
app.get('/api/editor-config/:id', (req, res) => {
  const entry = getFileByIdStmt.get(req.params.id);
  if (!entry) return res.status(404).json({ error: 'Not found' });

  // Space-based access control
  const userEmail = req.session.user.email;
  const space = entry.space || 'shared';
  if (space === 'private' && entry.owner_email !== userEmail) {
    return res.status(403).json({ error: 'Access denied.' });
  }
  if (space === 'shared' && !isAllowedForShared(userEmail)) {
    return res.status(403).json({ error: 'You are not in the allowed users list.' });
  }

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
  if (isTrashFolderId(folderId)) return res.status(400).json({ error: 'Cannot create files directly in the bin' });

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
  const VALID_SPACES = new Set(['private', 'shared', 'library']);
  const space = VALID_SPACES.has(req.body?.space) ? req.body.space : (folder.space || 'shared');
  insertFileStmt.run({
    id,
    original_name: originalName,
    stored_name: storedName,
    size: stat.size,
    document_type: allowed[type],
    file_type: type,
    uploaded_at: new Date().toISOString(),
    folder_id: folderId,
    space,
    owner_email: req.session.user.email,
  });
  res.json({ id, name: originalName });
});

// ---- API: rename file ----
app.patch('/api/files/:id', (req, res) => {
  const { name } = req.body;
  if (!name || typeof name !== 'string') return res.status(400).json({ error: 'Name required' });

  const entry = getFileByIdStmt.get(req.params.id);
  if (!entry) return res.status(404).json({ error: 'Not found' });
  const manageError = getFileManageError(req, entry);
  if (manageError) return res.status(403).json({ error: manageError });

  renameFileStmt.run(name.substring(0, 200), req.params.id);
  res.json({ ok: true });
});

app.patch('/api/files/:id/move', (req, res) => {
  const folderId = String(req.body?.folderId || '').trim();
  if (!folderId) return res.status(400).json({ error: 'folderId is required' });

  const file = getFileByIdStmt.get(req.params.id);
  if (!file) return res.status(404).json({ error: 'File not found' });
  const manageError = getFileManageError(req, file);
  if (manageError) return res.status(403).json({ error: manageError });

  const folder = getFolderByIdStmt.get(folderId);
  if (!folder) return res.status(404).json({ error: 'Folder not found' });
  const folderError = getFolderAccessError(req, folder, file.owner_email || '');
  if (folderError) return res.status(403).json({ error: folderError });
  if ((folder.space || 'shared') !== (file.space || 'shared')) {
    return res.status(400).json({ error: 'Files can only be moved within the same space' });
  }

  if (isTrashFolderId(folderId)) {
    updateFileLocation(file, folderId, file.folder_id, new Date().toISOString());
  } else {
    updateFileLocation(file, folderId, null, null);
  }

  res.json({ ok: true });
});

app.post('/api/files/:id/copy', (req, res) => {
  const source = getFileByIdStmt.get(req.params.id);
  if (!source) return res.status(404).json({ error: 'File not found' });
  const manageError = getFileManageError(req, source);
  if (manageError) return res.status(403).json({ error: manageError });

  const folderId = String(req.body?.folderId || source.folder_id || getRootFolderIdForSpace(source.space || 'shared')).trim();
  const folder = getFolderByIdStmt.get(folderId);
  if (!folder) return res.status(404).json({ error: 'Folder not found' });
  if (isTrashFolderId(folderId)) return res.status(400).json({ error: 'Cannot copy files into the bin' });
  if ((folder.space || 'shared') !== (source.space || 'shared')) {
    return res.status(400).json({ error: 'Files can only be copied within the same space' });
  }

  const sourcePath = path.join(UPLOADS_DIR, source.stored_name);
  if (!fs.existsSync(sourcePath)) return res.status(404).json({ error: 'Stored file is missing' });

  const copyId = crypto.randomUUID();
  const safeName = source.original_name.replace(/[^a-zA-Z0-9._-]/g, '_').substring(0, 200);
  const storedName = `${crypto.randomUUID().slice(0, 8)}_${safeName}`;
  fs.copyFileSync(sourcePath, path.join(UPLOADS_DIR, storedName));

  const ext = path.extname(source.original_name || '');
  const base = path.basename(source.original_name || 'Copy', ext);
  const originalName = `${base} Copy${ext}`.substring(0, 200);

  insertFileStmt.run({
    id: copyId,
    original_name: originalName,
    stored_name: storedName,
    size: source.size,
    document_type: source.document_type,
    file_type: source.file_type,
    uploaded_at: new Date().toISOString(),
    folder_id: folderId,
    space: source.space || 'shared',
    owner_email: req.session.user.email,
  });

  res.status(201).json({ ok: true, id: copyId, name: originalName });
});

app.post('/api/files/:id/restore', (req, res) => {
  const entry = getFileByIdStmt.get(req.params.id);
  if (!entry) return res.status(404).json({ error: 'File not found' });
  const manageError = getFileManageError(req, entry);
  if (manageError) return res.status(403).json({ error: manageError });
  if (!isTrashFolderId(entry.folder_id)) return res.status(400).json({ error: 'File is not in the bin' });

  const targetFolderId = entry.previous_folder_id || getRootFolderIdForSpace(entry.space || 'shared');
  const targetFolder = getFolderByIdStmt.get(targetFolderId) || getFolderByIdStmt.get(getRootFolderIdForSpace(entry.space || 'shared'));
  updateFileLocation(entry, targetFolder.id, null, null);
  res.json({ ok: true, folderId: targetFolder.id });
});

app.get('/api/files/:id/download', (req, res) => {
  const entry = getFileByIdStmt.get(req.params.id);
  const readError = getFileReadError(req, entry);
  if (readError) {
    return res.status(readError === 'File not found' ? 404 : 403).json({ error: readError });
  }

  const filePath = path.join(UPLOADS_DIR, entry.stored_name);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'Stored file is missing' });
  res.download(filePath, entry.original_name);
});

// ---- Admin API: users ----
app.get('/api/admin/config', requireAdmin, (_req, res) => {
  res.json({
    appUrl: APP_URL,
    onlyofficeUrl: ONLYOFFICE_URL,
    googleClientId: process.env.GOOGLE_CLIENT_ID || '',
    allowedEmails: ALLOWED_EMAILS_LIST,
    adminEmails: ADMIN_EMAILS_LIST,
    persistedTo: 'data/runtime-config.json',
    secrets: {
      googleClientSecret: maskSecret(process.env.GOOGLE_CLIENT_SECRET),
      sessionSecret: maskSecret(process.env.SESSION_SECRET),
      jwtSecret: maskSecret(process.env.JWT_SECRET),
    },
  });
});

app.patch('/api/admin/config', requireAdmin, (req, res) => {
  const updates = {};

  if (req.body?.appUrl !== undefined) {
    APP_URL = normaliseUrl(req.body.appUrl, APP_URL || `http://host.docker.internal:${PORT}`);
    process.env.APP_URL = APP_URL;
    updates.APP_URL = APP_URL;
  }

  if (req.body?.onlyofficeUrl !== undefined) {
    ONLYOFFICE_URL = normaliseUrl(req.body.onlyofficeUrl, ONLYOFFICE_URL || 'http://localhost:8080');
    process.env.ONLYOFFICE_URL = ONLYOFFICE_URL;
    updates.ONLYOFFICE_URL = ONLYOFFICE_URL;
  }

  if (req.body?.googleClientId !== undefined) {
    const googleClientId = String(req.body.googleClientId || '').trim();
    process.env.GOOGLE_CLIENT_ID = googleClientId;
    updates.GOOGLE_CLIENT_ID = googleClientId;
  }

  if (req.body?.allowedEmails !== undefined) {
    const allowedList = parseEmailList(req.body.allowedEmails);
    process.env.ALLOWED_EMAILS = listToCsv(allowedList);
    updates.ALLOWED_EMAILS = process.env.ALLOWED_EMAILS;
  }

  if (req.body?.adminEmails !== undefined) {
    const adminList = parseEmailList(req.body.adminEmails);
    process.env.ADMIN_EMAILS = listToCsv(adminList);
    updates.ADMIN_EMAILS = process.env.ADMIN_EMAILS;
  }

  if (req.body?.googleClientSecret !== undefined && String(req.body.googleClientSecret || '').trim()) {
    const secret = String(req.body.googleClientSecret);
    process.env.GOOGLE_CLIENT_SECRET = secret;
    updates.GOOGLE_CLIENT_SECRET = secret;
  }

  if (req.body?.sessionSecret !== undefined && String(req.body.sessionSecret || '').trim()) {
    const secret = String(req.body.sessionSecret);
    process.env.SESSION_SECRET = secret;
    updates.SESSION_SECRET = secret;
  }

  if (req.body?.jwtSecret !== undefined && String(req.body.jwtSecret || '').trim()) {
    const secret = String(req.body.jwtSecret);
    JWT_SECRET = secret;
    process.env.JWT_SECRET = secret;
    updates.JWT_SECRET = secret;
  }

  saveRuntimeConfig(updates);
  refreshAccessLists();
  FILE_TOKEN_SECRET = JWT_SECRET || process.env.SESSION_SECRET || 'file-token-fallback';

  res.json({
    ok: true,
    appUrl: APP_URL,
    onlyofficeUrl: ONLYOFFICE_URL,
    googleClientId: process.env.GOOGLE_CLIENT_ID || '',
    allowedEmails: ALLOWED_EMAILS_LIST,
    adminEmails: ADMIN_EMAILS_LIST,
    persistedTo: 'data/runtime-config.json',
    secrets: {
      googleClientSecret: maskSecret(process.env.GOOGLE_CLIENT_SECRET),
      sessionSecret: maskSecret(process.env.SESSION_SECRET),
      jwtSecret: maskSecret(process.env.JWT_SECRET),
    },
  });
});

app.get('/api/admin/users', requireAdmin, (_req, res) => {
  res.json(listUsersAdminStmt.all());
});

app.post('/api/admin/users', requireAdmin, async (req, res) => {
  const email = normaliseEmail(req.body?.email);
  const name = String(req.body?.name || '').trim() || email;
  const password = String(req.body?.password || '');
  const isAdmin = req.body?.isAdmin ? 1 : 0;
  if (!email || password.length < 8) {
    return res.status(400).json({ error: 'Valid email and a password of at least 8 characters are required.' });
  }
  if (getUserByEmailStmt.get(email)) {
    return res.status(409).json({ error: 'An account with this email already exists.' });
  }
  const password_hash = await bcrypt.hash(password, 12);
  const id = crypto.randomUUID();
  insertUserStmt.run({ id, email, name, password_hash, created_at: new Date().toISOString() });
  if (isAdmin) setUserAdminStmt.run(1, id);
  res.status(201).json({ ok: true, id });
});

app.patch('/api/admin/users/:id', requireAdmin, async (req, res) => {
  const user = getUserByIdStmt.get(req.params.id);
  if (!user) return res.status(404).json({ error: 'User not found' });

  if (req.body?.name !== undefined) {
    const name = String(req.body.name).trim();
    if (name) setUserNameStmt.run(name.substring(0, 200), req.params.id);
  }
  if (req.body?.isAdmin !== undefined) {
    if (req.params.id === req.session.user.id && !req.body.isAdmin) {
      return res.status(400).json({ error: 'Cannot remove your own admin privileges.' });
    }
    setUserAdminStmt.run(req.body.isAdmin ? 1 : 0, req.params.id);
  }
  if (req.body?.password !== undefined) {
    const password = String(req.body.password);
    if (password.length < 8) return res.status(400).json({ error: 'Password must be at least 8 characters.' });
    const hash = await bcrypt.hash(password, 12);
    db.prepare('UPDATE users SET password_hash = ? WHERE id = ?').run(hash, req.params.id);
  }
  res.json({ ok: true });
});

app.delete('/api/admin/users/:id', requireAdmin, (req, res) => {
  if (req.params.id === req.session.user.id) {
    return res.status(400).json({ error: 'Cannot delete your own account.' });
  }
  const user = getUserByIdStmt.get(req.params.id);
  if (!user) return res.status(404).json({ error: 'User not found' });
  deleteUserStmt.run(req.params.id);
  res.json({ ok: true });
});

// ---- Admin API: files ----
app.get('/api/admin/files', requireAdmin, (_req, res) => {
  res.json(listFilesStmt.all().map(toApiFile));
});

app.delete('/api/admin/files/:id', requireAdmin, (req, res) => {
  const entry = getFileByIdStmt.get(req.params.id);
  if (!entry) return res.status(404).json({ error: 'Not found' });
  const filePath = path.join(UPLOADS_DIR, entry.stored_name);
  if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
  deleteFileStmt.run(req.params.id);
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
