/* =========================================================
   OnlyOffice Dashboard – frontend logic
   ========================================================= */
'use strict';

// ── State ──────────────────────────────────────────────────
let allFiles = [];
let allFolders = [];
let folderPath = [];
let activeFilter = 'all';
let searchQuery = '';
let currentFolderId = 'root';

// ── DOM refs ───────────────────────────────────────────────
const fileGrid = document.getElementById('fileGrid');
const emptyState = document.getElementById('emptyState');
const fileInput = document.getElementById('fileInput');
const toast = document.getElementById('toast');
const newDocModal = document.getElementById('newDocModal');
const newDocName = document.getElementById('newDocName');
const folderBreadcrumb = document.getElementById('folderBreadcrumb');
const btnNewFolder = document.getElementById('btnNewFolder');
const navBtns = document.querySelectorAll('.sidebar-nav .nav-btn');
const fileActions = document.getElementById('fileActions');
const homePanel = document.getElementById('homePanel');
const filesPanel = document.getElementById('filesPanel');
const emailPanel = document.getElementById('emailPanel');
const tasksPanel = document.getElementById('tasksPanel');
const calendarPanel = document.getElementById('calendarPanel');

const ROUTES = new Set(['home', 'files', 'email', 'tasks', 'calendar']);
let filesLoaded = false;

// ── Icons ──────────────────────────────────────────────────
const TYPE_ICON = {
  word: '📄',
  cell: '📊',
  slide: '📑',
  other: '📁',
};

const TYPE_LABEL = {
  word: 'DOC', cell: 'XLS', slide: 'PPT',
};

// ── Utilities ──────────────────────────────────────────────
function formatSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function formatDate(iso) {
  const d = new Date(iso);
  return d.toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: 'numeric' });
}

function escHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

let toastTimer;
function showToast(msg, type = '', duration = 3000) {
  clearTimeout(toastTimer);
  toast.textContent = msg;
  toast.className = `toast ${type}`;
  toastTimer = setTimeout(() => { toast.className = 'toast hidden'; }, duration);
}

function showProgressToast(msg) {
  clearTimeout(toastTimer);
  toast.innerHTML = `
    <span>${msg}</span>
    <div class="progress-bar-wrap"><div class="progress-bar" id="progressBar"></div></div>`;
  toast.className = 'toast';
}

function updateProgress(pct) {
  const bar = document.getElementById('progressBar');
  if (bar) bar.style.width = pct + '%';
}

function renderBreadcrumb() {
  if (!folderPath.length) {
    folderBreadcrumb.innerHTML = '<span class="crumb current">Root</span>';
    return;
  }

  folderBreadcrumb.innerHTML = folderPath.map((item, index) => {
    const isCurrent = index === folderPath.length - 1;
    if (isCurrent) {
      return `<span class="crumb current">${escHtml(item.name)}</span>`;
    }
    return `<button class="crumb" data-folder-jump="${item.id}" type="button">${escHtml(item.name)}</button>`;
  }).join('<span class="crumb-sep">/</span>');
}

// ── Render ─────────────────────────────────────────────────
function renderFiles() {
  const filteredFolders = allFolders.filter((folder) => (
    !searchQuery || folder.name.toLowerCase().includes(searchQuery)
  ));

  const filteredFiles = allFiles.filter((file) => {
    const matchType = activeFilter === 'all' || file.type === activeFilter;
    const matchSearch = !searchQuery || file.name.toLowerCase().includes(searchQuery);
    return matchType && matchSearch;
  });

  if (filteredFolders.length === 0 && filteredFiles.length === 0) {
    fileGrid.innerHTML = '';
    emptyState.classList.remove('hidden');
    return;
  }

  emptyState.classList.add('hidden');

  const folderCards = filteredFolders.map((folder) => `
    <div class="file-card folder-card" data-folder-id="${folder.id}">
      <div class="file-card-thumb folder">📁</div>
      <div class="file-card-body">
        <div class="file-card-name" title="${escHtml(folder.name)}">${escHtml(folder.name)}</div>
        <div class="file-card-meta">Folder</div>
      </div>
    </div>`);

  const fileCards = filteredFiles.map((file) => {
    const iconClass = file.type || 'other';
    const badgeClass = 'badge-' + iconClass;
    const label = TYPE_LABEL[file.type] || 'FILE';
    return `
      <div class="file-card" data-id="${file.id}">
        <div class="file-card-thumb ${iconClass}">
          ${TYPE_ICON[file.type] || TYPE_ICON.other}
        </div>
        <span class="type-badge ${badgeClass}">${label}</span>
        <div class="file-card-body">
          <div class="file-card-name" title="${escHtml(file.name)}">${escHtml(file.name)}</div>
          <div class="file-card-meta">${formatSize(file.size)} · ${formatDate(file.uploadedAt)}</div>
        </div>
        <div class="file-card-actions">
          <button class="btn btn-primary btn-open" data-id="${file.id}">Open</button>
          <button class="btn btn-danger btn-delete" data-id="${file.id}" title="Delete">✕</button>
        </div>
      </div>`;
  });

  fileGrid.innerHTML = [...folderCards, ...fileCards].join('');
}

// ── Load files and folders ─────────────────────────────────
async function loadFiles(folderId = currentFolderId) {
  try {
    const res = await fetch(`/api/folders/${encodeURIComponent(folderId)}/contents`);
    if (!res.ok) throw new Error('Failed to load folder');
    const payload = await res.json();

    currentFolderId = payload.folder.id;
    allFolders = payload.folders || [];
    allFiles = payload.files || [];
    folderPath = payload.path || [];

    renderBreadcrumb();
    renderFiles();
  } catch (e) {
    showToast('Could not load folder: ' + e.message, 'error');
  }
}

async function createFolder() {
  const name = prompt('Folder name');
  if (!name) return;

  try {
    const res = await fetch('/api/folders', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, parentId: currentFolderId }),
    });
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(body.error || 'Failed to create folder');

    await loadFiles(currentFolderId);
    showToast('Folder created', 'success');
  } catch (e) {
    showToast('Create folder failed: ' + e.message, 'error');
  }
}

// ── Hash routing ───────────────────────────────────────────
function getRouteFromHash() {
  const hash = (location.hash || '').replace(/^#\/?/, '').trim().toLowerCase();
  if (!ROUTES.has(hash)) return 'home';
  return hash;
}

function goToRoute(route) {
  const target = ROUTES.has(route) ? route : 'home';
  if (getRouteFromHash() === target) {
    applyRoute(target);
    return;
  }
  location.hash = `#/${target}`;
}

function applyRoute(route) {
  const current = ROUTES.has(route) ? route : 'home';

  homePanel.classList.toggle('hidden', current !== 'home');
  filesPanel.classList.toggle('hidden', current !== 'files');
  emailPanel.classList.toggle('hidden', current !== 'email');
  tasksPanel.classList.toggle('hidden', current !== 'tasks');
  calendarPanel.classList.toggle('hidden', current !== 'calendar');

  fileActions.classList.toggle('hidden', current !== 'files');

  navBtns.forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.route === current);
  });

  if (current === 'files' && !filesLoaded) {
    filesLoaded = true;
    loadFiles(currentFolderId);
  }

  document.dispatchEvent(new CustomEvent('app:route-change', {
    detail: { route: current },
  }));
}

// ── Open file ──────────────────────────────────────────────
function openFile(id) {
  window.open(`/editor/${id}`, '_blank', 'noopener');
}

// ── Delete file ────────────────────────────────────────────
async function deleteFile(id) {
  const file = allFiles.find((f) => f.id === id);
  if (!file) return;
  if (!confirm(`Delete "${file.name}"? This cannot be undone.`)) return;

  try {
    const res = await fetch(`/api/files/${id}`, { method: 'DELETE' });
    if (!res.ok) throw new Error('Delete failed');
    await loadFiles(currentFolderId);
    showToast('File deleted', 'success');
  } catch (e) {
    showToast('Delete failed: ' + e.message, 'error');
  }
}

// ── Upload file ────────────────────────────────────────────
async function uploadFile(file) {
  showProgressToast(`Uploading "${file.name}"…`);

  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    const fd = new FormData();
    fd.append('file', file);
    fd.append('folderId', currentFolderId);

    xhr.upload.addEventListener('progress', (event) => {
      if (event.lengthComputable) updateProgress(Math.round((event.loaded / event.total) * 90));
    });

    xhr.addEventListener('load', () => {
      if (xhr.status === 200) {
        updateProgress(100);
        resolve(JSON.parse(xhr.responseText));
      } else {
        let errMsg = 'Upload failed';
        try { errMsg = JSON.parse(xhr.responseText).error || errMsg; } catch {}
        reject(new Error(errMsg));
      }
    });

    xhr.addEventListener('error', () => reject(new Error('Network error')));

    xhr.open('POST', '/api/upload');
    xhr.send(fd);
  });
}

async function handleFiles(fileList) {
  const files = Array.from(fileList);
  if (!files.length) return;

  let succeeded = 0;
  for (const file of files) {
    try {
      await uploadFile(file);
      succeeded++;
    } catch (e) {
      showToast(e.message, 'error', 4000);
    }
  }

  if (succeeded > 0) {
    showToast(`${succeeded} file${succeeded > 1 ? 's' : ''} uploaded`, 'success');
    await loadFiles(currentFolderId);
  }
}

// ── New document modal ─────────────────────────────────────
let selectedDocType = 'docx';

function openNewDocModal() {
  selectedDocType = 'docx';
  newDocName.value = '';
  document.querySelectorAll('.doc-type-btn').forEach((button) => {
    button.classList.toggle('selected', button.dataset.type === selectedDocType);
  });
  newDocModal.classList.remove('hidden');
  setTimeout(() => newDocName.focus(), 50);
}

function closeNewDocModal() {
  newDocModal.classList.add('hidden');
}

async function createNewDocument() {
  const rawName = newDocName.value.trim() || `Untitled.${selectedDocType}`;
  const name = rawName.endsWith(`.${selectedDocType}`) ? rawName : `${rawName}.${selectedDocType}`;

  try {
    const res = await fetch('/api/create', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, type: selectedDocType, folderId: currentFolderId }),
    });
    if (!res.ok) {
      const body = await res.json().catch(() => ({}));
      throw new Error(body.error || 'Create failed');
    }
    const { id } = await res.json();
    closeNewDocModal();
    await loadFiles(currentFolderId);
    openFile(id);
  } catch (e) {
    showToast('Failed to create document: ' + e.message, 'error');
  }
}

// ── Event wiring ───────────────────────────────────────────
navBtns.forEach((btn) => {
  btn.addEventListener('click', () => {
    goToRoute(btn.dataset.route || 'home');
  });
});

window.addEventListener('hashchange', () => {
  applyRoute(getRouteFromHash());
});

btnNewFolder.addEventListener('click', createFolder);

folderBreadcrumb.addEventListener('click', (event) => {
  const jump = event.target.closest('[data-folder-jump]');
  if (!jump) return;
  loadFiles(jump.dataset.folderJump);
});

// Upload button
document.getElementById('btnUpload').addEventListener('click', () => fileInput.click());

fileInput.addEventListener('change', () => {
  handleFiles(fileInput.files);
  fileInput.value = '';
});

// New document button
document.getElementById('btnNewDoc').addEventListener('click', openNewDocModal);
document.getElementById('btnCancelNew').addEventListener('click', closeNewDocModal);
document.getElementById('btnConfirmNew').addEventListener('click', createNewDocument);

newDocModal.addEventListener('click', (event) => {
  if (event.target === newDocModal) closeNewDocModal();
});

// Doc type selector
document.querySelectorAll('.doc-type-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    selectedDocType = btn.dataset.type;
    document.querySelectorAll('.doc-type-btn').forEach((button) => button.classList.remove('selected'));
    btn.classList.add('selected');
    if (!newDocName.value || /^\.(docx|xlsx|pptx)$/.test(newDocName.value)) {
      newDocName.value = '';
    }
    newDocName.placeholder = `Untitled.${selectedDocType}`;
  });
});

// Enter key in new doc name
newDocName.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') createNewDocument();
  if (event.key === 'Escape') closeNewDocModal();
});

// Filter tabs
document.querySelectorAll('.filter-tab').forEach((tab) => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.filter-tab').forEach((item) => item.classList.remove('active'));
    tab.classList.add('active');
    activeFilter = tab.dataset.filter;
    renderFiles();
  });
});

// Search
document.getElementById('searchInput').addEventListener('input', (event) => {
  searchQuery = event.target.value.toLowerCase().trim();
  renderFiles();
});

// File grid – open / delete / folder enter
fileGrid.addEventListener('click', (event) => {
  const openBtn = event.target.closest('.btn-open');
  const deleteBtn = event.target.closest('.btn-delete');
  const folderCard = event.target.closest('.folder-card');
  const card = event.target.closest('.file-card');

  if (deleteBtn) {
    event.stopPropagation();
    deleteFile(deleteBtn.dataset.id);
    return;
  }

  if (openBtn) {
    event.stopPropagation();
    openFile(openBtn.dataset.id);
    return;
  }

  if (folderCard) {
    event.stopPropagation();
    loadFiles(folderCard.dataset.folderId);
    return;
  }

  if (card && !event.target.closest('.file-card-actions')) {
    openFile(card.dataset.id);
  }
});

// Drag & drop on the whole page
let dragCounter = 0;
document.addEventListener('dragenter', (event) => {
  if (event.dataTransfer.types.includes('Files')) {
    dragCounter++;
    document.body.classList.add('drag-over');
  }
});

document.addEventListener('dragleave', () => {
  dragCounter--;
  if (dragCounter <= 0) {
    dragCounter = 0;
    document.body.classList.remove('drag-over');
  }
});

document.addEventListener('dragover', (event) => event.preventDefault());

document.addEventListener('drop', (event) => {
  event.preventDefault();
  dragCounter = 0;
  document.body.classList.remove('drag-over');
  handleFiles(event.dataTransfer.files);
});

// ── Init ───────────────────────────────────────────────────
if (!location.hash || !ROUTES.has(getRouteFromHash())) {
  location.hash = '#/home';
}
applyRoute(getRouteFromHash());
