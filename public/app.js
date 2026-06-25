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
let currentSpace = 'shared'; // 'private' | 'shared' | 'library'
let currentUserEmail = '';
let isCurrentUserAdmin = false;
let viewingTrash = false;

const SPACE_ROOTS = { private: 'root-private', shared: 'root', library: 'root-library' };
const TRASH_ROOTS = { private: 'root-private-trash', shared: 'root-trash', library: 'root-library-trash' };

const SPACE_META = {
  private: {
    label: 'My Files',
    desc: 'Private documents only visible to you.',
    icon: '🔒',
  },
  shared: {
    label: 'Team Shared',
    desc: 'Files shared with your whole team.',
    icon: '👥',
  },
  library: {
    label: 'Business Library',
    desc: 'Curated SOPs, instructions, and permanent company documentation accessible to everyone.',
    icon: '📚',
  },
};

// ── DOM refs ───────────────────────────────────────────────
const fileGrid = document.getElementById('fileGrid');
const emptyState = document.getElementById('emptyState');
const fileInput = document.getElementById('fileInput');
const toast = document.getElementById('toast');
const newDocModal = document.getElementById('newDocModal');
const newDocName = document.getElementById('newDocName');
const fileManageModal = document.getElementById('fileManageModal');
const fileManageTitle = document.getElementById('fileManageTitle');
const fileManageName = document.getElementById('fileManageName');
const fileManageDestination = document.getElementById('fileManageDestination');
const fileManageDestinationRow = document.getElementById('fileManageDestinationRow');
const fileManageTrashNote = document.getElementById('fileManageTrashNote');
const folderBreadcrumb = document.getElementById('folderBreadcrumb');
const btnNewFolder = document.getElementById('btnNewFolder');
const btnToggleTrash = document.getElementById('btnToggleTrash');
const userPill = document.getElementById('userPill');
const userEmailEl = document.getElementById('userEmail');
const btnGoogleSignIn = document.getElementById('btnGoogleSignIn');
const btnLogout = document.getElementById('btnLogout');
const navBtns = document.querySelectorAll('.sidebar-nav .nav-btn');
const fileActions = document.getElementById('fileActions');
const homePanel = document.getElementById('homePanel');
const filesPanel = document.getElementById('filesPanel');
const emailPanel = document.getElementById('emailPanel');
const tasksPanel = document.getElementById('tasksPanel');
const calendarPanel = document.getElementById('calendarPanel');
const adminPanel = document.getElementById('adminPanel');
const spaceTabs = document.querySelectorAll('.space-tab');
const spaceDescription = document.getElementById('spaceDescription');
const newDocSpace = document.getElementById('newDocSpace');

const ROUTES = new Set(['home', 'files', 'email', 'tasks', 'calendar', 'admin']);
let filesLoaded = false;
let adminConfigLoaded = false;
let managedFileId = null;

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

async function syncHeaderAuthState() {
  try {
    const res = await fetch('/auth/status');
    const data = await res.json();
    if (!data.authenticated) {
      userPill.style.display = 'none';
      btnGoogleSignIn.style.display = '';
      return;
    }

    userPill.style.display = 'flex';
    userEmailEl.textContent = data.name || data.email || 'Signed in';
    currentUserEmail = data.email || '';
    isCurrentUserAdmin = !!data.isAdmin;
    btnGoogleSignIn.style.display = data.provider === 'google' ? 'none' : '';
    const btnAdmin = document.getElementById('btnAdmin');
    if (btnAdmin) btnAdmin.classList.toggle('hidden', !data.isAdmin);
  } catch {
    userPill.style.display = 'none';
    btnGoogleSignIn.style.display = '';
    isCurrentUserAdmin = false;
  }
}

function setConfigValue(id, value) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = value || '(not set)';
}

function renderAdminConfig(data) {
  const allowed = Array.isArray(data.allowedEmails) && data.allowedEmails.length
    ? data.allowedEmails.join(', ')
    : '(open to all authenticated users)';
  const admins = Array.isArray(data.adminEmails) && data.adminEmails.length
    ? data.adminEmails.join(', ')
    : '(none configured)';

  setConfigValue('cfgAppUrl', data.appUrl);
  setConfigValue('cfgOnlyofficeUrl', data.onlyofficeUrl);
  setConfigValue('cfgGoogleClientId', data.googleClientId);
  setConfigValue('cfgAllowedEmails', allowed);
  setConfigValue('cfgAdminEmails', admins);
  setConfigValue('cfgGoogleClientSecret', data.secrets?.googleClientSecret || '');
  setConfigValue('cfgSessionSecret', data.secrets?.sessionSecret || '');
  setConfigValue('cfgJwtSecret', data.secrets?.jwtSecret || '');
}

function getCurrentRootFolderId() {
  return viewingTrash ? TRASH_ROOTS[currentSpace] : SPACE_ROOTS[currentSpace];
}

function getFolderById(id) {
  return allFolders.find((folder) => folder.id === id) || null;
}

function updateTrashButton() {
  if (!btnToggleTrash) return;
  btnToggleTrash.innerHTML = viewingTrash
    ? '<svg viewBox="0 0 20 20" fill="currentColor"><path d="M12.707 4.293a1 1 0 010 1.414L9.414 9H17a1 1 0 110 2H9.414l3.293 3.293a1 1 0 01-1.414 1.414l-5-5a1 1 0 010-1.414l5-5a1 1 0 011.414 0z"/></svg>Back to Files'
    : '<svg viewBox="0 0 20 20" fill="currentColor"><path d="M6 2a1 1 0 00-1 1v1H3.5a1 1 0 100 2h.538l.853 10.232A2 2 0 006.884 18h6.232a2 2 0 001.993-1.768L15.962 6h.538a1 1 0 100-2H15V3a1 1 0 00-1-1H6zm2 2V4h4V4H8z"/></svg>Bin';
}

function getFolderLookupLabel(folder, index) {
  return `${index + 1}. ${folder.path}`;
}

function getFolderOptionLabel(folder) {
  return folder.path;
}

async function fetchFolderTree(parentId, trail = []) {
  const res = await fetch(`/api/folders?parentId=${encodeURIComponent(parentId)}`);
  if (!res.ok) throw new Error('Failed to load folders');
  const folders = await res.json();
  const nested = await Promise.all(folders
    .filter((folder) => folder.space === currentSpace && !folder.isTrash)
    .map(async (folder) => {
      const pathItems = [...trail, folder.name];
      const children = await fetchFolderTree(folder.id, pathItems);
      return [{ ...folder, path: pathItems.join(' / ') }, ...children];
    }));
  return nested.flat();
}

async function chooseDestinationFolder(defaultFolderId = currentFolderId, actionLabel = 'Select a folder') {
  const rootId = SPACE_ROOTS[currentSpace];
  const rootLabel = SPACE_META[currentSpace]?.label || 'Root';
  const folders = [{ id: rootId, path: rootLabel }, ...(await fetchFolderTree(rootId, [rootLabel]))]
    .filter((folder) => folder.id !== TRASH_ROOTS[currentSpace]);

  const defaultIndex = Math.max(0, folders.findIndex((folder) => folder.id === defaultFolderId));
  const promptText = [
    `${actionLabel}:`,
    ...folders.map((folder, index) => getFolderLookupLabel(folder, index)),
  ].join('\n');
  const selection = window.prompt(promptText, String(defaultIndex + 1));
  if (selection === null) return null;

  const selectedIndex = Number.parseInt(selection, 10) - 1;
  if (!Number.isInteger(selectedIndex) || !folders[selectedIndex]) {
    throw new Error('Invalid folder selection');
  }
  return folders[selectedIndex];
}

async function getDestinationFolders() {
  const rootId = SPACE_ROOTS[currentSpace];
  const rootLabel = SPACE_META[currentSpace]?.label || 'Root';
  return [{ id: rootId, path: rootLabel }, ...(await fetchFolderTree(rootId, [rootLabel]))]
    .filter((folder) => folder.id !== TRASH_ROOTS[currentSpace]);
}

async function apiJson(url, options = {}) {
  const res = await fetch(url, options);
  const body = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(body.error || 'Request failed');
  return body;
}

function getManagedFile() {
  return allFiles.find((entry) => entry.id === managedFileId) || null;
}

function closeFileManageModal() {
  managedFileId = null;
  fileManageModal.classList.add('hidden');
}

async function populateFileDestinationOptions(selectedFolderId) {
  const folders = await getDestinationFolders();
  fileManageDestination.innerHTML = folders.map((folder) => `
    <option value="${folder.id}">${escHtml(getFolderOptionLabel(folder))}</option>
  `).join('');

  if (!folders.length) {
    fileManageDestination.value = '';
    return;
  }

  fileManageDestination.value = selectedFolderId && folders.some((folder) => folder.id === selectedFolderId)
    ? selectedFolderId
    : folders[0].id;
}

async function openFileManageModal(id) {
  const file = allFiles.find((entry) => entry.id === id);
  if (!file) return;

  managedFileId = id;
  fileManageTitle.textContent = `Manage ${file.deletedAt ? 'Bin File' : 'File'}`;
  fileManageName.value = file.name;

  const inTrash = !!file.deletedAt;
  fileManageDestinationRow.classList.toggle('hidden', inTrash);
  fileManageTrashNote.classList.toggle('hidden', !inTrash);
  document.getElementById('btnFileRename').classList.toggle('hidden', inTrash);
  document.getElementById('btnFileMove').classList.toggle('hidden', inTrash);
  document.getElementById('btnFileCopy').classList.toggle('hidden', inTrash);
  document.getElementById('btnFileRestore').classList.toggle('hidden', !inTrash);
  document.getElementById('btnFileDelete').classList.toggle('hidden', inTrash);
  document.getElementById('btnFileDeletePermanent').classList.toggle('hidden', !inTrash);

  if (!inTrash) {
    await populateFileDestinationOptions(currentFolderId);
  }

  fileManageModal.classList.remove('hidden');
  setTimeout(() => fileManageName.focus(), 50);
}

async function loadAdminConfig(force = false) {
  if (adminConfigLoaded && !force) return;

  const state = document.getElementById('adminConfigState');
  const grid = document.getElementById('adminConfigGrid');
  if (state) state.textContent = 'Loading configuration…';
  if (grid) grid.classList.add('hidden');

  try {
    const res = await fetch('/api/admin/config');
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || 'Failed to load admin configuration');
    renderAdminConfig(data);
    adminConfigLoaded = true;
    if (state) state.textContent = '';
    if (grid) grid.classList.remove('hidden');
  } catch (err) {
    if (state) state.textContent = 'Could not load admin settings: ' + err.message;
  }
}
// ── Space switching ────────────────────────────────────────
function applySpaceTabs() {
  spaceTabs.forEach((btn) => btn.classList.toggle('active', btn.dataset.space === currentSpace));

  const meta = SPACE_META[currentSpace];
  if (spaceDescription) {
    spaceDescription.innerHTML = `<span class="space-desc-icon">${meta.icon}</span>${meta.desc}`;
  }

  const folderToolbar = document.querySelector('.folder-toolbar');
  if (folderToolbar) folderToolbar.classList.remove('hidden');
  if (btnNewFolder) btnNewFolder.classList.toggle('hidden', currentSpace !== 'shared' || viewingTrash);
  updateTrashButton();
}

function switchSpace(space) {
  if (!SPACE_ROOTS[space]) return;
  currentSpace = space;
  viewingTrash = false;
  currentFolderId = SPACE_ROOTS[space];
  filesLoaded = false;
  applySpaceTabs();
  loadFiles(currentFolderId);
}

function renderBreadcrumb() {
  if (!folderPath.length) {
    folderBreadcrumb.innerHTML = `<span class="crumb current">${viewingTrash ? 'Bin' : 'Root'}</span>`;
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
    const isOwner = file.ownerEmail === currentUserEmail;
    const space = file.space || 'shared';
    const spaceBadgeHtml = space !== 'shared'
      ? `<span class="space-badge space-badge-${space}">${space === 'private' ? '🔒' : '📚'}</span>`
      : '';
    // Show delete only if owner (or for shared: anyone can delete)
    const inTrash = !!file.deletedAt;
    const canDelete = space === 'shared' || isOwner || (space === 'library' && isCurrentUserAdmin);
    return `
      <div class="file-card" data-id="${file.id}">
        <div class="file-card-thumb ${iconClass}">
          ${TYPE_ICON[file.type] || TYPE_ICON.other}
        </div>
        <span class="type-badge ${badgeClass}">${label}</span>
        ${spaceBadgeHtml}
        <div class="file-card-body">
          <div class="file-card-name" title="${escHtml(file.name)}">${escHtml(file.name)}</div>
          <div class="file-card-meta">${formatSize(file.size)} · ${formatDate(file.uploadedAt)}${inTrash ? ' · In bin' : ''}</div>
        </div>
        <div class="file-card-actions">
          <button class="btn btn-primary btn-open" data-id="${file.id}">Open</button>
          <button class="btn btn-ghost btn-download" data-id="${file.id}">Save</button>
          ${canDelete ? `<button class="btn ${inTrash ? 'btn-danger' : 'btn-ghost'} btn-manage" data-id="${file.id}" title="Manage">${inTrash ? 'Trash' : 'Manage'}</button>` : ''}
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
    viewingTrash = !!payload.folder?.isTrash;

    renderBreadcrumb();
    renderFiles();
    updateTrashButton();
  } catch (e) {
    showToast('Could not load folder: ' + e.message, 'error');
  }
}

async function createFolder() {
  if (viewingTrash) {
    showToast('Cannot create folders in the bin', 'error');
    return;
  }

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

  if (current === 'admin' && !isCurrentUserAdmin) {
    location.hash = '#/home';
    return;
  }

  homePanel.classList.toggle('hidden', current !== 'home');
  filesPanel.classList.toggle('hidden', current !== 'files');
  emailPanel.classList.toggle('hidden', current !== 'email');
  tasksPanel.classList.toggle('hidden', current !== 'tasks');
  calendarPanel.classList.toggle('hidden', current !== 'calendar');
  adminPanel.classList.toggle('hidden', current !== 'admin');

  fileActions.classList.toggle('hidden', current !== 'files');

  navBtns.forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.route === current);
  });

  if (current === 'files' && !filesLoaded) {
    filesLoaded = true;
    loadFiles(currentFolderId);
  }

  if (current === 'admin') {
    loadAdminConfig();
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
  const action = file.deletedAt ? 'Permanently delete' : 'Move to the bin';
  if (!confirm(`${action} "${file.name}"?`)) return;

  try {
    await apiJson(`/api/files/${id}`, { method: 'DELETE' });
    await loadFiles(currentFolderId);
    closeFileManageModal();
    showToast(file.deletedAt ? 'File deleted permanently' : 'File moved to bin', 'success');
  } catch (e) {
    showToast('Delete failed: ' + e.message, 'error');
  }
}

function downloadFile(id) {
  window.location.href = `/api/files/${encodeURIComponent(id)}/download`;
}

async function renameFile(id) {
  const file = allFiles.find((entry) => entry.id === id);
  if (!file) return;
  const trimmed = String(fileManageName?.value || file.name).trim();
  if (!trimmed) {
    showToast('File name is required', 'error');
    return;
  }

  await apiJson(`/api/files/${encodeURIComponent(id)}`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: trimmed }),
  });
  await loadFiles(currentFolderId);
  closeFileManageModal();
  showToast('File renamed', 'success');
}

async function moveFileToFolder(id, destinationId = fileManageDestination.value) {
  if (!destinationId) {
    showToast('Destination folder is required', 'error');
    return;
  }

  await apiJson(`/api/files/${encodeURIComponent(id)}/move`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ folderId: destinationId }),
  });
  await loadFiles(currentFolderId);
  closeFileManageModal();
  showToast('File moved', 'success');
}

async function copyFileToFolder(id, destinationId = fileManageDestination.value) {
  if (!destinationId) {
    showToast('Destination folder is required', 'error');
    return;
  }

  await apiJson(`/api/files/${encodeURIComponent(id)}/copy`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ folderId: destinationId }),
  });
  await loadFiles(currentFolderId);
  closeFileManageModal();
  showToast('File copied', 'success');
}

async function restoreFile(id) {
  await apiJson(`/api/files/${encodeURIComponent(id)}/restore`, { method: 'POST' });
  await loadFiles(currentFolderId);
  closeFileManageModal();
  showToast('File restored', 'success');
}

async function manageFile(id) {
  try {
    await openFileManageModal(id);
  } catch (e) {
    showToast(e.message || 'Could not open file manager', 'error');
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
    fd.append('space', currentSpace);

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
  if (newDocSpace) newDocSpace.value = currentSpace;
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
  const space = newDocSpace ? newDocSpace.value : currentSpace;

  try {
    const res = await fetch('/api/create', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, type: selectedDocType, folderId: currentFolderId, space }),
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
btnToggleTrash?.addEventListener('click', () => {
  viewingTrash = !viewingTrash;
  loadFiles(getCurrentRootFolderId());
});

document.getElementById('btnRefreshAdminConfig')?.addEventListener('click', () => {
  loadAdminConfig(true);
});

// Space tabs
spaceTabs.forEach((btn) => {
  btn.addEventListener('click', () => switchSpace(btn.dataset.space));
});

btnLogout?.addEventListener('click', async () => {
  await fetch('/auth/logout', { method: 'POST' });
  location.href = '/login';
});

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
document.getElementById('btnFileManageCancel').addEventListener('click', closeFileManageModal);
document.getElementById('btnFileRename').addEventListener('click', async () => {
  const file = getManagedFile();
  if (!file) return;
  try {
    await renameFile(file.id);
  } catch (e) {
    showToast(e.message || 'Rename failed', 'error');
  }
});
document.getElementById('btnFileMove').addEventListener('click', async () => {
  const file = getManagedFile();
  if (!file) return;
  try {
    await moveFileToFolder(file.id);
  } catch (e) {
    showToast(e.message || 'Move failed', 'error');
  }
});
document.getElementById('btnFileCopy').addEventListener('click', async () => {
  const file = getManagedFile();
  if (!file) return;
  try {
    await copyFileToFolder(file.id);
  } catch (e) {
    showToast(e.message || 'Copy failed', 'error');
  }
});
document.getElementById('btnFileDelete').addEventListener('click', async () => {
  const file = getManagedFile();
  if (!file) return;
  await deleteFile(file.id);
});
document.getElementById('btnFileDeletePermanent').addEventListener('click', async () => {
  const file = getManagedFile();
  if (!file) return;
  await deleteFile(file.id);
});
document.getElementById('btnFileRestore').addEventListener('click', async () => {
  const file = getManagedFile();
  if (!file) return;
  try {
    await restoreFile(file.id);
  } catch (e) {
    showToast(e.message || 'Restore failed', 'error');
  }
});

newDocModal.addEventListener('click', (event) => {
  if (event.target === newDocModal) closeNewDocModal();
});

fileManageModal.addEventListener('click', (event) => {
  if (event.target === fileManageModal) closeFileManageModal();
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

fileManageName.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    const file = getManagedFile();
    if (file && !file.deletedAt) {
      renameFile(file.id).catch((e) => {
        showToast(e.message || 'Rename failed', 'error');
      });
    }
  }
  if (event.key === 'Escape') closeFileManageModal();
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
  const downloadBtn = event.target.closest('.btn-download');
  const manageBtn = event.target.closest('.btn-manage');
  const folderCard = event.target.closest('.folder-card');
  const card = event.target.closest('.file-card');

  if (downloadBtn) {
    event.stopPropagation();
    downloadFile(downloadBtn.dataset.id);
    return;
  }

  if (manageBtn) {
    event.stopPropagation();
    manageFile(manageBtn.dataset.id);
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
syncHeaderAuthState();
applyRoute(getRouteFromHash());
applySpaceTabs();
updateTrashButton();
