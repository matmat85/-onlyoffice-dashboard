/* =========================================================
   OnlyOffice Dashboard – frontend logic
   ========================================================= */
'use strict';

// ── State ──────────────────────────────────────────────────
let allFiles = [];
let activeFilter = 'all';
let searchQuery = '';

// ── DOM refs ───────────────────────────────────────────────
const fileGrid     = document.getElementById('fileGrid');
const emptyState   = document.getElementById('emptyState');
const fileInput    = document.getElementById('fileInput');
const toast        = document.getElementById('toast');
const newDocModal  = document.getElementById('newDocModal');
const newDocName   = document.getElementById('newDocName');

// ── Icons ──────────────────────────────────────────────────
const TYPE_ICON = {
  word:  '📄',
  cell:  '📊',
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

// ── Render ─────────────────────────────────────────────────
function renderFiles() {
  const filtered = allFiles.filter(f => {
    const matchType   = activeFilter === 'all' || f.type === activeFilter;
    const matchSearch = !searchQuery || f.name.toLowerCase().includes(searchQuery);
    return matchType && matchSearch;
  });

  if (filtered.length === 0) {
    fileGrid.innerHTML = '';
    emptyState.classList.remove('hidden');
    return;
  }

  emptyState.classList.add('hidden');

  fileGrid.innerHTML = filtered.map(f => {
    const iconClass = f.type || 'other';
    const badgeClass = 'badge-' + iconClass;
    const label = TYPE_LABEL[f.type] || 'FILE';
    return `
      <div class="file-card" data-id="${f.id}">
        <div class="file-card-thumb ${iconClass}">
          ${TYPE_ICON[f.type] || TYPE_ICON.other}
        </div>
        <span class="type-badge ${badgeClass}">${label}</span>
        <div class="file-card-body">
          <div class="file-card-name" title="${escHtml(f.name)}">${escHtml(f.name)}</div>
          <div class="file-card-meta">${formatSize(f.size)} · ${formatDate(f.uploadedAt)}</div>
        </div>
        <div class="file-card-actions">
          <button class="btn btn-primary btn-open"   data-id="${f.id}">Open</button>
          <button class="btn btn-danger  btn-delete" data-id="${f.id}" title="Delete">✕</button>
        </div>
      </div>`;
  }).join('');
}

function escHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── Load files ─────────────────────────────────────────────
async function loadFiles() {
  try {
    const res = await fetch('/api/files');
    if (!res.ok) throw new Error('Failed to load files');
    allFiles = await res.json();
    renderFiles();
  } catch (e) {
    showToast('Could not load files: ' + e.message, 'error');
  }
}

// ── Open file ──────────────────────────────────────────────
function openFile(id) {
  window.open(`/editor/${id}`, '_blank', 'noopener');
}

// ── Delete file ────────────────────────────────────────────
async function deleteFile(id) {
  const file = allFiles.find(f => f.id === id);
  if (!file) return;
  if (!confirm(`Delete "${file.name}"? This cannot be undone.`)) return;

  try {
    const res = await fetch(`/api/files/${id}`, { method: 'DELETE' });
    if (!res.ok) throw new Error('Delete failed');
    allFiles = allFiles.filter(f => f.id !== id);
    renderFiles();
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
    const fd  = new FormData();
    fd.append('file', file);

    xhr.upload.addEventListener('progress', e => {
      if (e.lengthComputable) updateProgress(Math.round((e.loaded / e.total) * 90));
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
    await loadFiles();
  }
}

// ── New document modal ─────────────────────────────────────
let selectedDocType = 'docx';

function openNewDocModal() {
  selectedDocType = 'docx';
  newDocName.value = '';
  document.querySelectorAll('.doc-type-btn').forEach(b => {
    b.classList.toggle('selected', b.dataset.type === selectedDocType);
  });
  newDocModal.classList.remove('hidden');
  setTimeout(() => newDocName.focus(), 50);
}

function closeNewDocModal() {
  newDocModal.classList.add('hidden');
}

async function createNewDocument() {
  const rawName = newDocName.value.trim() || `Untitled.${selectedDocType}`;
  // Ensure the chosen extension is appended if user didn't type it
  const name = rawName.endsWith(`.${selectedDocType}`) ? rawName : `${rawName}.${selectedDocType}`;

  try {
    const res = await fetch('/api/create', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, type: selectedDocType }),
    });
    if (!res.ok) {
      const body = await res.json().catch(() => ({}));
      throw new Error(body.error || 'Create failed');
    }
    const { id } = await res.json();
    closeNewDocModal();
    await loadFiles();
    openFile(id);
  } catch (e) {
    showToast('Failed to create document: ' + e.message, 'error');
  }
}

// ── Event wiring ───────────────────────────────────────────
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

newDocModal.addEventListener('click', e => {
  if (e.target === newDocModal) closeNewDocModal();
});

// Doc type selector
document.querySelectorAll('.doc-type-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    selectedDocType = btn.dataset.type;
    document.querySelectorAll('.doc-type-btn').forEach(b => b.classList.remove('selected'));
    btn.classList.add('selected');
    // Pre-fill extension in name input if empty or just an extension
    if (!newDocName.value || /^\.(docx|xlsx|pptx)$/.test(newDocName.value)) {
      newDocName.value = '';
    }
    newDocName.placeholder = `Untitled.${selectedDocType}`;
  });
});

// Enter key in new doc name
newDocName.addEventListener('keydown', e => {
  if (e.key === 'Enter') createNewDocument();
  if (e.key === 'Escape') closeNewDocModal();
});

// Filter tabs
document.querySelectorAll('.filter-tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.filter-tab').forEach(t => t.classList.remove('active'));
    tab.classList.add('active');
    activeFilter = tab.dataset.filter;
    renderFiles();
  });
});

// Search
document.getElementById('searchInput').addEventListener('input', e => {
  searchQuery = e.target.value.toLowerCase().trim();
  renderFiles();
});

// File grid – open / delete
fileGrid.addEventListener('click', e => {
  const openBtn   = e.target.closest('.btn-open');
  const deleteBtn = e.target.closest('.btn-delete');
  const card      = e.target.closest('.file-card');

  if (deleteBtn) {
    e.stopPropagation();
    deleteFile(deleteBtn.dataset.id);
    return;
  }
  if (openBtn) {
    e.stopPropagation();
    openFile(openBtn.dataset.id);
    return;
  }
  if (card && !e.target.closest('.file-card-actions')) {
    openFile(card.dataset.id);
  }
});

// Drag & drop on the whole page
let dragCounter = 0;
document.addEventListener('dragenter', e => {
  if (e.dataTransfer.types.includes('Files')) { dragCounter++; document.body.classList.add('drag-over'); }
});
document.addEventListener('dragleave', () => {
  dragCounter--;
  if (dragCounter <= 0) { dragCounter = 0; document.body.classList.remove('drag-over'); }
});
document.addEventListener('dragover', e => e.preventDefault());
document.addEventListener('drop', e => {
  e.preventDefault();
  dragCounter = 0;
  document.body.classList.remove('drag-over');
  handleFiles(e.dataTransfer.files);
});

// ── Init ───────────────────────────────────────────────────
loadFiles();
