/* =============================================================
   email.js – Gmail integration panel
   ============================================================= */
'use strict';

// ── State ────────────────────────────────────────────────────
let emailState = {
  authenticated: false,
  userEmail: null,
  activeLabel: 'INBOX',
  nextPageToken: null,
  searchQuery: '',
};

// ── DOM refs ──────────────────────────────────────────────────
const emailPanel       = document.getElementById('emailPanel');
const emailLayout      = document.getElementById('emailLayout');
const emailSignInPrmt  = document.getElementById('emailSignInPrompt');
const labelList        = document.getElementById('labelList');
const emailList        = document.getElementById('emailList');
const emailReader      = document.getElementById('emailReader');
const btnCompose       = document.getElementById('btnCompose');
const composeModal     = document.getElementById('composeModal');
const emailBadge       = document.getElementById('emailBadge');
const btnEmailNext     = document.getElementById('btnEmailNext');
const userPill         = document.getElementById('userPill');
const userEmailEl      = document.getElementById('userEmail');
const btnLogout        = document.getElementById('btnLogout');
const btnGoogleSignIn  = document.getElementById('btnGoogleSignIn');

let emailPanelInitialised = false;

// ── Auth status check ─────────────────────────────────────────
async function checkAuthStatus() {
  try {
    const res = await fetch('/auth/google/status');
    const { authenticated, email, name } = await res.json();
    emailState.authenticated = authenticated;
    emailState.userEmail = email || null;
    const signedIn = !!email;
    if (signedIn) {
      userPill.style.display = 'flex';
      userEmailEl.textContent = name || email;
      btnGoogleSignIn.style.display = authenticated ? 'none' : '';
    } else {
      userPill.style.display = 'none';
      btnGoogleSignIn.style.display = '';
    }
  } catch { emailState.authenticated = false; }
}

// ── Init email panel ──────────────────────────────────────────
async function initEmailPanel() {
  if (emailPanelInitialised) return;
  emailPanelInitialised = true;

  await checkAuthStatus();
  if (!emailState.authenticated) {
    emailLayout.classList.add('hidden');
    emailSignInPrmt.classList.remove('hidden');
    return;
  }
  emailSignInPrmt.classList.add('hidden');
  emailLayout.classList.remove('hidden');
  await loadLabels();
  await loadMessages();
}

function isEmailRoute() {
  return (location.hash || '').replace(/^#\/?/, '').trim().toLowerCase() === 'email';
}

document.addEventListener('app:route-change', (event) => {
  if (event.detail?.route === 'email') {
    initEmailPanel();
  }
});

// ── Labels (folders) ─────────────────────────────────────────
const SHOW_LABELS = ['INBOX','SENT','DRAFTS','TRASH','STARRED','SPAM'];

async function loadLabels() {
  const res = await fetch('/api/email/labels');
  if (!res.ok) return;
  const labels = await res.json();
  const filtered = labels.filter(l => SHOW_LABELS.includes(l.id) || l.type === 'user');
  labelList.innerHTML = filtered.map(l => `
    <li class="label-item${l.id === emailState.activeLabel ? ' active' : ''}" data-label="${l.id}">
      ${labelIcon(l.id)} ${l.name}
      ${l.messagesUnread ? `<span class="label-unread">${l.messagesUnread}</span>` : ''}
    </li>`).join('');
  labelList.querySelectorAll('.label-item').forEach(li => {
    li.addEventListener('click', () => {
      emailState.activeLabel = li.dataset.label;
      emailState.nextPageToken = null;
      labelList.querySelectorAll('.label-item').forEach(x => x.classList.remove('active'));
      li.classList.add('active');
      loadMessages();
    });
  });
}

function labelIcon(id) {
  const icons = { INBOX:'📥', SENT:'📤', DRAFTS:'📝', TRASH:'🗑️', STARRED:'⭐', SPAM:'⚠️' };
  return `<span>${icons[id] || '📁'}</span>`;
}

// ── Message list ──────────────────────────────────────────────
async function loadMessages(append = false) {
  if (!append) { emailList.innerHTML = '<li class="email-loading">Loading…</li>'; emailState.nextPageToken = null; }
  const params = new URLSearchParams({ label: emailState.activeLabel });
  if (emailState.searchQuery) params.set('q', emailState.searchQuery);
  if (emailState.nextPageToken) params.set('pageToken', emailState.nextPageToken);

  const res = await fetch(`/api/email/messages?${params}`);
  if (!res.ok) { emailList.innerHTML = '<li class="email-loading">Error loading messages</li>'; return; }
  const { messages, nextPageToken } = await res.json();
  emailState.nextPageToken = nextPageToken || null;
  btnEmailNext.classList.toggle('hidden', !nextPageToken);

  if (!append) emailList.innerHTML = '';
  if (!messages.length && !append) {
    emailList.innerHTML = '<li class="email-loading">No messages</li>';
    return;
  }
  messages.forEach(m => {
    const li = document.createElement('li');
    li.className = 'email-item' + (m.unread ? ' unread' : '');
    li.dataset.id = m.id;
    li.innerHTML = `
      <div class="email-item-from">${escEmail(senderName(m.from))}</div>
      <div class="email-item-subject">${escEmail(m.subject || '(no subject)')}</div>
      <div class="email-item-snippet">${escEmail(m.snippet)}</div>
      <div class="email-item-date">${formatEmailDate(m.date)}</div>`;
    li.addEventListener('click', () => openMessage(m.id, li));
    emailList.appendChild(li);
  });
}

// ── Open + read message ───────────────────────────────────────
async function openMessage(id, listItem) {
  document.querySelectorAll('.email-item').forEach(el => el.classList.remove('selected'));
  listItem?.classList.add('selected');
  listItem?.classList.remove('unread');
  emailReader.innerHTML = '<div style="padding:40px;color:var(--text-muted)">Loading…</div>';

  const res = await fetch(`/api/email/message/${id}`);
  if (!res.ok) { emailReader.innerHTML = '<div style="padding:40px;color:var(--danger)">Failed to load message.</div>'; return; }
  const msg = await res.json();

  emailReader.innerHTML = `
    <div class="email-reader-header">
      <div class="email-reader-subject">${escEmail(msg.subject || '(no subject)')}</div>
      <div class="email-reader-meta">
        <span><strong>From:</strong> ${escEmail(msg.from)}</span>
        <span><strong>To:</strong> ${escEmail(msg.to)}</span>
        ${msg.cc ? `<span><strong>Cc:</strong> ${escEmail(msg.cc)}</span>` : ''}
        <span><strong>Date:</strong> ${escEmail(msg.date)}</span>
      </div>
      <div class="email-reader-actions">
        <button class="btn btn-ghost btn-sm" id="btnReply">↩ Reply</button>
        <button class="btn btn-danger btn-sm" id="btnTrash">🗑 Delete</button>
      </div>
    </div>
    <div class="email-reader-body">
      <iframe id="emailBodyFrame" sandbox="allow-same-origin" style="width:100%;min-height:400px;border:none;background:#fff;border-radius:6px;"></iframe>
    </div>`;

  // Render body safely inside a sandboxed iframe
  const frame = document.getElementById('emailBodyFrame');
  frame.onload = () => {
    try {
      frame.contentDocument.open();
      frame.contentDocument.write(msg.body || '<p style="color:#666">Empty message</p>');
      frame.contentDocument.close();
      frame.style.height = frame.contentDocument.body.scrollHeight + 40 + 'px';
    } catch {}
  };
  frame.src = 'about:blank';

  document.getElementById('btnTrash')?.addEventListener('click', () => trashMessage(id, listItem));
  document.getElementById('btnReply')?.addEventListener('click', () => openReply(msg));
}

async function trashMessage(id, listItem) {
  if (!confirm('Move this message to Trash?')) return;
  await fetch(`/api/email/message/${id}`, { method: 'DELETE' });
  listItem?.remove();
  emailReader.innerHTML = '<div class="email-reader-placeholder"><p>Message deleted</p></div>';
}

// ── Compose ───────────────────────────────────────────────────
function openCompose(defaults = {}) {
  document.getElementById('composeTo').value      = defaults.to || '';
  document.getElementById('composeSubject').value = defaults.subject || '';
  document.getElementById('composeBody').value    = defaults.body || '';
  composeModal.classList.remove('hidden');
}

function openReply(msg) {
  const replySubject = msg.subject?.startsWith('Re:') ? msg.subject : `Re: ${msg.subject}`;
  openCompose({ to: msg.from, subject: replySubject, body: '\n\n---\n' + msg.body?.replace(/<[^>]+>/g,'') });
}

btnCompose?.addEventListener('click', () => openCompose());
document.getElementById('btnCancelCompose')?.addEventListener('click', () => composeModal.classList.add('hidden'));

document.getElementById('btnSendEmail')?.addEventListener('click', async () => {
  const to      = document.getElementById('composeTo').value.trim();
  const subject = document.getElementById('composeSubject').value.trim();
  const body    = document.getElementById('composeBody').value;
  if (!to || !subject) { showToast('To and Subject are required', 'error'); return; }

  const btn = document.getElementById('btnSendEmail');
  btn.disabled = true; btn.textContent = 'Sending…';
  try {
    const res = await fetch('/api/email/send', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ to, subject, body: body.replace(/\n/g,'<br>') }),
    });
    if (!res.ok) throw new Error((await res.json()).error || 'Send failed');
    composeModal.classList.add('hidden');
    showToast('Email sent!', 'success');
  } catch (e) {
    showToast('Send failed: ' + e.message, 'error');
  } finally { btn.disabled = false; btn.textContent = 'Send'; }
});

// ── Search ────────────────────────────────────────────────────
let searchTimer;
document.getElementById('emailSearch')?.addEventListener('input', e => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => {
    emailState.searchQuery = e.target.value.trim();
    emailState.nextPageToken = null;
    loadMessages();
  }, 400);
});

document.getElementById('btnEmailRefresh')?.addEventListener('click', () => loadMessages());
btnEmailNext?.addEventListener('click', () => loadMessages(true));

// ── Logout ────────────────────────────────────────────────────
btnLogout?.addEventListener('click', async () => {
  await fetch('/auth/google/logout', { method: 'POST' });
  location.href = '/login';
});

// ── Unread badge polling (every 60s) ─────────────────────────
async function updateUnreadBadge() {
  if (!emailState.authenticated) return;
  try {
    const res = await fetch('/api/email/messages?label=INBOX');
    if (!res.ok) return;
    const { messages } = await res.json();
    const unreadCount = messages.filter(m => m.unread).length;
    if (unreadCount > 0) {
      emailBadge.textContent = unreadCount > 9 ? '9+' : unreadCount;
      emailBadge.classList.remove('hidden');
    } else {
      emailBadge.classList.add('hidden');
    }
  } catch {}
}

// ── Utilities ─────────────────────────────────────────────────
function escEmail(str) {
  return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function senderName(from) {
  const m = from?.match(/^"?([^"<]+)"?\s*</);
  return m ? m[1].trim() : (from || '');
}
function formatEmailDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  const now = new Date();
  if (d.toDateString() === now.toDateString()) return d.toLocaleTimeString([], { hour:'2-digit', minute:'2-digit' });
  return d.toLocaleDateString([], { day:'2-digit', month:'short' });
}

// ── Startup ───────────────────────────────────────────────────
(async () => {
  await checkAuthStatus();
  if (emailState.authenticated) {
    updateUnreadBadge();
    setInterval(updateUnreadBadge, 60_000);
  }
  if (isEmailRoute()) {
    initEmailPanel();
  }
})();


