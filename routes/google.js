/**
 * Google OAuth2 + Gmail API routes
 * Mount in server.js:  app.use(require('./routes/google'))
 */
'use strict';

const express  = require('express');
const { google } = require('googleapis');
const path     = require('path');
const fs       = require('fs');

const router = express.Router();

// ---------------------------------------------------------------------------
// Token store  (data/tokens.json  →  { "user@co.com": { ...tokens } })
// ---------------------------------------------------------------------------
const DATA_DIR   = path.join(__dirname, '..', 'data');
const TOKEN_FILE = path.join(DATA_DIR, 'tokens.json');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

function loadTokens() {
  if (!fs.existsSync(TOKEN_FILE)) return {};
  try { return JSON.parse(fs.readFileSync(TOKEN_FILE, 'utf8')); } catch { return {}; }
}
function saveTokens(tokens) {
  fs.writeFileSync(TOKEN_FILE, JSON.stringify(tokens, null, 2));
}

// ---------------------------------------------------------------------------
// OAuth2 client factory
// ---------------------------------------------------------------------------
function makeOAuth2Client() {
  const base = (process.env.APP_URL || 'http://localhost:3000').replace(/\/+$/, '');
  const redirectUri = process.env.GOOGLE_REDIRECT_URI || `${base}/auth/google/callback`;
  return new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    redirectUri
  );
}

/** Returns an authenticated OAuth2 client for the given email, or null. */
function getAuthedClient(userEmail) {
  const stored = loadTokens()[userEmail];
  if (!stored) return null;
  const client = makeOAuth2Client();
  client.setCredentials(stored);
  // Persist refreshed token automatically
  client.on('tokens', (newTokens) => {
    const all = loadTokens();
    all[userEmail] = { ...stored, ...newTokens };
    saveTokens(all);
  });
  return client;
}

// ---------------------------------------------------------------------------
// Allowlist helper
// ---------------------------------------------------------------------------
/**
 * Returns true if the given email is allowed.
 * Set ALLOWED_EMAILS=alice@co.com,bob@co.com in .env to restrict access.
 * If the variable is absent, any authenticated Google account is accepted.
 */
function isAllowed(email) {
  const raw = process.env.ALLOWED_EMAILS || '';
  if (!raw.trim()) return true; // open to any authenticated user
  return raw.split(',').map(e => e.trim().toLowerCase()).includes(email.toLowerCase());
}

// ---------------------------------------------------------------------------
// Auth routes (public — no session required)
// ---------------------------------------------------------------------------
router.get('/auth/google', (_req, res) => {
  const client = makeOAuth2Client();
  const url = client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: [
      'https://www.googleapis.com/auth/gmail.modify',
      'https://www.googleapis.com/auth/calendar.readonly',
      'https://www.googleapis.com/auth/userinfo.email',
      'https://www.googleapis.com/auth/userinfo.profile',
    ],
  });
  res.redirect(url);
});

router.get('/auth/google/callback', async (req, res) => {
  const { code } = req.query;
  if (!code) return res.redirect('/login?error=oauth');

  try {
    const client = makeOAuth2Client();
    const { tokens } = await client.getToken(code);
    client.setCredentials(tokens);

    const oauth2 = google.oauth2({ version: 'v2', auth: client });
    const { data: userInfo } = await oauth2.userinfo.get();

    // Allowlist check
    if (!isAllowed(userInfo.email)) {
      console.warn(`[google/callback] Blocked unauthorised email: ${userInfo.email}`);
      return res.redirect('/login?error=unauthorized');
    }

    // Persist OAuth tokens
    const all = loadTokens();
    all[userInfo.email] = tokens;
    saveTokens(all);

    // Create server-side session
    req.session.regenerate((err) => {
      if (err) { console.error('[google/callback] session regenerate:', err); return res.redirect('/login?error=oauth'); }
      req.session.user = { email: userInfo.email, name: userInfo.name, picture: userInfo.picture };
      req.session.save((err2) => {
        if (err2) { console.error('[google/callback] session save:', err2); return res.redirect('/login?error=oauth'); }
        res.redirect('/');
      });
    });
  } catch (e) {
    console.error('[google/callback]', e.message);
    res.redirect('/login?error=oauth');
  }
});

router.get('/auth/google/status', (req, res) => {
  const user = req.session?.user;
  if (!user) return res.json({ authenticated: false });
  const tokens = loadTokens()[user.email];
  res.json({ authenticated: !!tokens, email: user.email, name: user.name, picture: user.picture });
});

router.post('/auth/google/logout', (req, res) => {
  req.session.destroy(() => {
    res.clearCookie('connect.sid');
    res.json({ ok: true });
  });
});

// ---------------------------------------------------------------------------
// Gmail API routes  (session user required)
// ---------------------------------------------------------------------------
function requireGoogleAuth(req, res, next) {
  const email = req.session?.user?.email;
  if (!email) return res.status(401).json({ error: 'Not authenticated with Google' });
  const client = getAuthedClient(email);
  if (!client) return res.status(401).json({ error: 'No stored tokens — please sign in again' });
  req.googleClient = client;
  req.googleEmail  = email;
  next();
}

/** GET /api/email/labels  — list Gmail labels (folders) */
router.get('/api/email/labels', requireGoogleAuth, async (req, res, next) => {
  try {
    const gmail = google.gmail({ version: 'v1', auth: req.googleClient });
    const { data } = await gmail.users.labels.list({ userId: 'me' });
    res.json(data.labels || []);
  } catch (e) { next(e); }
});

/** GET /api/email/messages?label=INBOX&q=&pageToken=  — list messages */
router.get('/api/email/messages', requireGoogleAuth, async (req, res, next) => {
  try {
    const gmail = google.gmail({ version: 'v1', auth: req.googleClient });
    const { label = 'INBOX', q = '', pageToken } = req.query;
    const params = { userId: 'me', maxResults: 25, labelIds: [label] };
    if (q)         params.q = q;
    if (pageToken) params.pageToken = pageToken;

    const { data } = await gmail.users.messages.list(params);
    if (!data.messages?.length) return res.json({ messages: [], nextPageToken: null });

    const details = await Promise.all(
      data.messages.map(m =>
        gmail.users.messages.get({ userId: 'me', id: m.id, format: 'metadata',
          metadataHeaders: ['Subject', 'From', 'Date', 'To'] })
      )
    );

    const messages = details.map(({ data: msg }) => {
      const h = (name) => msg.payload.headers.find(x => x.name === name)?.value || '';
      return {
        id: msg.id, threadId: msg.threadId, snippet: msg.snippet,
        labelIds: msg.labelIds,
        subject: h('Subject'), from: h('From'), to: h('To'), date: h('Date'),
        unread: msg.labelIds?.includes('UNREAD'),
      };
    });

    res.json({ messages, nextPageToken: data.nextPageToken || null });
  } catch (e) { next(e); }
});

/** GET /api/email/message/:id  — full message with decoded body */
router.get('/api/email/message/:id', requireGoogleAuth, async (req, res, next) => {
  try {
    const gmail = google.gmail({ version: 'v1', auth: req.googleClient });
    const { data: msg } = await gmail.users.messages.get({
      userId: 'me', id: req.params.id, format: 'full',
    });

    if (msg.labelIds?.includes('UNREAD')) {
      await gmail.users.messages.modify({
        userId: 'me', id: req.params.id,
        requestBody: { removeLabelIds: ['UNREAD'] },
      }).catch(() => {});
    }

    const h = (name) => msg.payload.headers.find(x => x.name === name)?.value || '';

    function extractBody(payload) {
      if (!payload) return '';
      if (payload.body?.data) return Buffer.from(payload.body.data, 'base64').toString('utf8');
      for (const part of payload.parts || []) {
        if (part.mimeType === 'text/html') return Buffer.from(part.body.data || '', 'base64').toString('utf8');
      }
      for (const part of payload.parts || []) {
        if (part.mimeType === 'text/plain') return '<pre>' + Buffer.from(part.body.data || '', 'base64').toString('utf8') + '</pre>';
      }
      return '';
    }

    res.json({
      id: msg.id, threadId: msg.threadId,
      subject: h('Subject'), from: h('From'), to: h('To'), date: h('Date'), cc: h('Cc'),
      body: extractBody(msg.payload), labelIds: msg.labelIds,
    });
  } catch (e) { next(e); }
});

/** POST /api/email/send  — send an email */
router.post('/api/email/send', requireGoogleAuth, async (req, res, next) => {
  try {
    const { to, subject, body, replyToMessageId, threadId } = req.body;
    if (!to || !subject || !body) return res.status(400).json({ error: 'to, subject, body required' });

    const gmail = google.gmail({ version: 'v1', auth: req.googleClient });
    const mime = [
      `To: ${to}`, `Subject: ${subject}`,
      `Content-Type: text/html; charset=utf-8`, `MIME-Version: 1.0`,
      replyToMessageId ? `In-Reply-To: ${replyToMessageId}` : '',
      '', body,
    ].filter(Boolean).join('\r\n');

    const encoded = Buffer.from(mime).toString('base64url');
    const params = { userId: 'me', requestBody: { raw: encoded } };
    if (threadId) params.requestBody.threadId = threadId;

    const { data } = await gmail.users.messages.send(params);
    res.json({ id: data.id });
  } catch (e) { next(e); }
});

/** DELETE /api/email/message/:id  — move to trash */
router.delete('/api/email/message/:id', requireGoogleAuth, async (req, res, next) => {
  try {
    const gmail = google.gmail({ version: 'v1', auth: req.googleClient });
    await gmail.users.messages.trash({ userId: 'me', id: req.params.id });
    res.json({ ok: true });
  } catch (e) { next(e); }
});

module.exports = router;


