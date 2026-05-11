/* ═══════════════════════════════════════════
   HSA Tracker – Application Logic
   Auth, Google Sheets, Google Drive, UI
   ═══════════════════════════════════════════ */

// ── State ──
let expenses = [];      // { row, name, date, amount, reimbursed, reimbDate, receipt }
let selected = new Set();
let currentFilter = 'all';
let currentYear = 'all';
let sortCol = 'date';
let sortAsc = false;

// ── Auth helpers ──
function getToken() {
  return sessionStorage.getItem('hsa_access_token');
}

function requireAuth() {
  if (!getToken()) {
    window.location.href = 'index.html';
    return false;
  }
  return true;
}

function signOut() {
  const token = getToken();
  if (token) {
    google.accounts.oauth2.revoke(token, () => {});
  }
  sessionStorage.clear();
  window.location.href = 'index.html';
}

// ── Formatting ──
function fmt(n) {
  return '$' + Number(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function fmtDateDisplay(isoDate) {
  if (!isoDate) return '';
  const parts = isoDate.split('-');
  if (parts.length !== 3) return isoDate;
  return parts[1] + '/' + parts[2] + '/' + parts[0];
}

function parseSheetDate(val) {
  if (!val) return '';
  // Handle MM/DD/YYYY format from sheets
  const parts = val.split('/');
  if (parts.length === 3) {
    const y = parts[2].length === 2 ? '20' + parts[2] : parts[2];
    return y + '-' + parts[0].padStart(2, '0') + '-' + parts[1].padStart(2, '0');
  }
  // Already ISO
  return val;
}

function toSheetDate(isoDate) {
  if (!isoDate) return '';
  const parts = isoDate.split('-');
  return parts[1] + '/' + parts[2] + '/' + parts[0];
}

// ── Google Sheets API ──
async function apiFetch(url, options = {}) {
  const token = getToken();
  if (!token) { requireAuth(); return null; }

  const headers = { Authorization: 'Bearer ' + token, ...(options.headers || {}) };
  const res = await fetch(url, { ...options, headers });

  if (res.status === 401) {
    // Token expired
    sessionStorage.clear();
    window.location.href = 'index.html';
    return null;
  }
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || 'API error ' + res.status);
  }
  return res.json();
}

async function loadExpenses() {
  const range = encodeURIComponent(CONFIG.SHEET_NAME + '!A:F');
  const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + CONFIG.SPREADSHEET_ID + '/values/' + range;
  const data = await apiFetch(url);
  if (!data) return;

  const rows = data.values || [];
  expenses = [];

  // Skip header row (index 0)
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue; // skip empty rows
    expenses.push({
      row: i + 1, // 1-based sheet row (header is row 1, first data is row 2)
      name: r[0] || '',
      date: parseSheetDate(r[1] || ''),
      amount: parseFloat((r[2] || '0').replace(/[$,]/g, '')) || 0,
      reimbursed: (r[3] || '').toLowerCase().trim() === 'yes',
      reimbDate: parseSheetDate(r[4] || ''),
      receipt: r[5] || '',
    });
  }
}

async function appendExpense(name, date, amount, receiptLink) {
  const range = encodeURIComponent(CONFIG.SHEET_NAME + '!A:F');
  const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + CONFIG.SPREADSHEET_ID
    + '/values/' + range + ':append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS';

  const values = [[name, toSheetDate(date), amount, 'No', '', receiptLink]];

  await apiFetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ values }),
  });
}

async function updateReimbursed(sheetRow, reimbDateISO) {
  const range = encodeURIComponent(CONFIG.SHEET_NAME + '!D' + sheetRow + ':E' + sheetRow);
  const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + CONFIG.SPREADSHEET_ID
    + '/values/' + range + '?valueInputOption=USER_ENTERED';

  await apiFetch(url, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: [['Yes', toSheetDate(reimbDateISO)]] }),
  });
}

async function updateExpenseRow(sheetRow, name, date, amount) {
  const range = encodeURIComponent(CONFIG.SHEET_NAME + '!A' + sheetRow + ':C' + sheetRow);
  const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + CONFIG.SPREADSHEET_ID
    + '/values/' + range + '?valueInputOption=USER_ENTERED';

  await apiFetch(url, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: [[name, toSheetDate(date), amount]] }),
  });
}

async function deleteExpenseRow(sheetRow) {
  // Get the sheet's grid ID
  const metaUrl = 'https://sheets.googleapis.com/v4/spreadsheets/' + CONFIG.SPREADSHEET_ID
    + '?fields=sheets(properties)';
  const meta = await apiFetch(metaUrl);
  const sheetMeta = meta.sheets.find(s => s.properties.title === CONFIG.SHEET_NAME);
  const sheetGid = sheetMeta ? sheetMeta.properties.sheetId : 0;

  const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + CONFIG.SPREADSHEET_ID
    + ':batchUpdate';

  await apiFetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      requests: [{
        deleteDimension: {
          range: {
            sheetId: sheetGid,
            dimension: 'ROWS',
            startIndex: sheetRow - 1,
            endIndex: sheetRow,
          }
        }
      }]
    }),
  });
}

// ── Google Drive API ──
async function uploadReceipt(file) {
  const token = getToken();
  if (!token) { requireAuth(); return ''; }

  const metadata = {
    name: file.name,
    parents: [CONFIG.DRIVE_FOLDER_ID],
  };

  const form = new FormData();
  form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
  form.append('file', file);

  // Upload file
  const uploadRes = await fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id',
    { method: 'POST', headers: { Authorization: 'Bearer ' + token }, body: form }
  );

  if (!uploadRes.ok) {
    const err = await uploadRes.json().catch(() => ({}));
    throw new Error(err.error?.message || 'Drive upload failed');
  }

  const uploaded = await uploadRes.json();
  const fileId = uploaded.id;

  // Make shareable (anyone with link)
  await apiFetch('https://www.googleapis.com/drive/v3/files/' + fileId + '/permissions', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ role: 'reader', type: 'anyone' }),
  });

  // Get web view link
  const fileData = await apiFetch(
    'https://www.googleapis.com/drive/v3/files/' + fileId + '?fields=webViewLink'
  );

  return fileData?.webViewLink || ('https://drive.google.com/file/d/' + fileId + '/view');
}

// ── UI: Views ──
function switchView(viewName) {
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.getElementById('view-' + viewName).classList.add('active');
  document.querySelectorAll('.nav-links button').forEach(b => {
    b.classList.toggle('active', b.dataset.view === viewName);
  });

  if (viewName === 'expenses') {
    refreshExpenses();
  }
}

// ── UI: Add Expense ──
let selectedFile = null;

function initDropzone() {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');

  dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('drag-over'); });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag-over'));
  dropzone.addEventListener('drop', e => {
    e.preventDefault();
    dropzone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) showFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', () => {
    if (fileInput.files.length) showFile(fileInput.files[0]);
  });
}

function showFile(file) {
  if (file.size > 10 * 1024 * 1024) {
    showToast('File too large. Max 10 MB.', 'error');
    return;
  }
  selectedFile = file;
  document.getElementById('fileName').textContent = file.name + ' (' + (file.size / 1024).toFixed(1) + ' KB)';
  document.getElementById('filePreview').classList.add('show');
  document.getElementById('dropzone').style.display = 'none';
}

function removeFile() {
  selectedFile = null;
  document.getElementById('fileInput').value = '';
  document.getElementById('filePreview').classList.remove('show');
  document.getElementById('dropzone').style.display = '';
}

async function submitExpense() {
  const name = document.getElementById('expenseName').value.trim();
  const date = document.getElementById('expenseDate').value;
  const amount = document.getElementById('expenseAmount').value;

  if (!name) { showToast('Please enter an expense name.', 'error'); return; }
  if (!date) { showToast('Please enter a date.', 'error'); return; }
  if (!amount || parseFloat(amount) <= 0) { showToast('Please enter a valid amount.', 'error'); return; }

  const btn = document.getElementById('submitBtn');
  const txtEl = document.getElementById('submitText');
  btn.disabled = true;
  txtEl.innerHTML = '<span class="spinner"></span> Submitting...';

  try {
    let receiptLink = '';

    // Upload receipt if provided
    if (selectedFile) {
      txtEl.innerHTML = '<span class="spinner"></span> Uploading receipt...';
      receiptLink = await uploadReceipt(selectedFile);
    }

    // Append to sheet
    txtEl.innerHTML = '<span class="spinner"></span> Saving to spreadsheet...';
    await appendExpense(name, date, parseFloat(amount).toFixed(2), receiptLink);

    showToast('Expense submitted successfully!', 'success');

    // Reset form
    document.getElementById('expenseName').value = '';
    document.getElementById('expenseAmount').value = '';
    document.getElementById('expenseDate').valueAsDate = new Date();
    removeFile();

  } catch (err) {
    showToast('Error: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
    txtEl.textContent = '✅ Submit Expense';
  }
}

// ── UI: Expenses List ──
async function refreshExpenses() {
  showLoading('Loading expenses...');
  try {
    await loadExpenses();
    selected.clear();
    populateYearFilter();
    renderExpenses();
  } catch (err) {
    showToast('Failed to load expenses: ' + err.message, 'error');
  } finally {
    hideLoading();
  }
}

function populateYearFilter() {
  const years = new Set();
  expenses.forEach(e => {
    if (e.date) years.add(e.date.substring(0, 4));
  });

  const sel = document.getElementById('yearFilter');
  const prev = sel.value;
  sel.innerHTML = '<option value="all">All Years</option>';
  Array.from(years).sort().reverse().forEach(y => {
    sel.innerHTML += '<option value="' + y + '">' + y + '</option>';
  });
  sel.value = years.has(prev) ? prev : 'all';
  currentYear = sel.value;
}

function getFilteredExpenses() {
  let list = expenses;

  // Year filter
  if (currentYear !== 'all') {
    list = list.filter(e => e.date.startsWith(currentYear));
  }

  // Status filter
  if (currentFilter === 'pending') list = list.filter(e => !e.reimbursed);
  if (currentFilter === 'done') list = list.filter(e => e.reimbursed);

  // Sort
  list = [...list].sort((a, b) => {
    let va = a[sortCol], vb = b[sortCol];
    if (sortCol === 'amount') {
      va = Number(va); vb = Number(vb);
    } else if (sortCol === 'reimbursed') {
      va = va ? 1 : 0; vb = vb ? 1 : 0;
    } else {
      va = String(va).toLowerCase(); vb = String(vb).toLowerCase();
    }
    if (va < vb) return sortAsc ? -1 : 1;
    if (va > vb) return sortAsc ? 1 : -1;
    return 0;
  });

  return list;
}

function renderExpenses() {
  const list = getFilteredExpenses();
  const tbody = document.getElementById('expenseBody');

  if (list.length === 0) {
    tbody.innerHTML = '<tr><td colspan="8"><div class="empty-state">'
      + '<div class="empty-icon">📋</div><p>No expenses found.</p></div></td></tr>';
  } else {
    tbody.innerHTML = list.map(e => {
      const isSelected = selected.has(e.row);
      return '<tr class="' + (isSelected ? 'selected' : '') + '" onclick="toggleSelect(' + e.row + ')">'
        + '<td><div class="cb"><input type="checkbox" ' + (isSelected ? 'checked' : '')
        + ' onclick="event.stopPropagation(); toggleSelect(' + e.row + ')"><div class="checkmark"></div></div></td>'
        + '<td>' + escHtml(e.name) + '</td>'
        + '<td>' + fmtDateDisplay(e.date) + '</td>'
        + '<td class="amount-cell">' + fmt(e.amount) + '</td>'
        + '<td><span class="badge ' + (e.reimbursed ? 'badge-yes' : 'badge-no') + '">'
        + (e.reimbursed ? 'Yes' : 'No') + '</span></td>'
        + '<td>' + fmtDateDisplay(e.reimbDate) + '</td>'
        + '<td>' + (e.receipt
          ? '<a class="receipt-link" href="' + escAttr(e.receipt) + '" target="_blank" rel="noopener" onclick="event.stopPropagation()">🔗 View</a>'
          : '<span style="color:var(--text-muted);font-size:12px">--</span>') + '</td>'
        + '<td class="row-actions">'
        + '<button class="action-edit" onclick="event.stopPropagation(); openEditModal(' + e.row + ')" title="Edit">✏️</button>'
        + '<button class="action-delete" onclick="event.stopPropagation(); confirmDelete(' + e.row + ')" title="Delete">🗑️</button>'
        + '</td>'
        + '</tr>';
    }).join('');
  }

  updateSummary(list);
  updateActionBar();
  updateSortHeaders();
}

function escHtml(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

function escAttr(s) {
  return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function updateSummary(filtered) {
  const data = currentYear !== 'all'
    ? expenses.filter(e => e.date.startsWith(currentYear))
    : expenses;

  const total = data.reduce((s, e) => s + e.amount, 0);
  const reimb = data.filter(e => e.reimbursed).reduce((s, e) => s + e.amount, 0);
  document.getElementById('totalVal').textContent = fmt(total);
  document.getElementById('reimbVal').textContent = fmt(reimb);
  document.getElementById('pendingVal').textContent = fmt(total - reimb);
}

function updateSortHeaders() {
  document.querySelectorAll('.data-table th[data-col]').forEach(th => {
    const col = th.dataset.col;
    const arrow = th.querySelector('.sort-arrow');
    th.classList.toggle('sorted', col === sortCol);
    if (col === sortCol) {
      arrow.innerHTML = sortAsc ? '&#9650;' : '&#9660;';
    } else {
      arrow.innerHTML = '&#9650;';
    }
  });
}

// ── Selection ──
function toggleSelect(row) {
  selected.has(row) ? selected.delete(row) : selected.add(row);
  renderExpenses();
}

function toggleSelectAll(checked) {
  const list = getFilteredExpenses();
  if (checked) {
    list.forEach(e => selected.add(e.row));
  } else {
    selected.clear();
  }
  renderExpenses();
}

function updateActionBar() {
  const items = expenses.filter(e => selected.has(e.row));
  document.getElementById('selCount').textContent = items.length;
  document.getElementById('selTotal').textContent = fmt(items.reduce((s, e) => s + e.amount, 0));
  document.getElementById('actionBar').classList.toggle('visible', items.length > 0);
}

// ── Reimbursement Modal ──
function openModal() {
  const items = expenses.filter(e => selected.has(e.row));
  if (!items.length) return;

  document.getElementById('reimbDate').valueAsDate = new Date();
  document.getElementById('modalItems').innerHTML = items.map(e =>
    '<div class="modal-item"><span>' + escHtml(e.name) + '</span><span class="mi-amount">' + fmt(e.amount) + '</span></div>'
  ).join('');
  document.getElementById('modalTotal').textContent = fmt(items.reduce((s, e) => s + e.amount, 0));
  document.getElementById('modal').classList.add('show');
}

function closeModal() {
  document.getElementById('modal').classList.remove('show');
}

async function confirmReimburse() {
  const dt = document.getElementById('reimbDate').value;
  if (!dt) { showToast('Please select a reimbursement date.', 'error'); return; }

  const items = expenses.filter(e => selected.has(e.row));
  const btn = document.getElementById('confirmBtn');
  btn.disabled = true;
  btn.textContent = 'Updating...';

  try {
    // Update each row in the sheet
    for (const e of items) {
      await updateReimbursed(e.row, dt);
      e.reimbursed = true;
      e.reimbDate = dt;
    }

    selected.clear();
    closeModal();
    renderExpenses();
    showToast(items.length + ' expense(s) marked as reimbursed!', 'success');
  } catch (err) {
    showToast('Error updating: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = '✅ Confirm';
  }
}

// ── Edit Expense ──
function openEditModal(row) {
  const e = expenses.find(x => x.row === row);
  if (!e) return;
  document.getElementById('editRow').value = row;
  document.getElementById('editName').value = e.name;
  document.getElementById('editDate').value = e.date;
  document.getElementById('editAmount').value = e.amount;
  document.getElementById('editModal').classList.add('show');
}

function closeEditModal() {
  document.getElementById('editModal').classList.remove('show');
}

async function saveEdit() {
  const row = parseInt(document.getElementById('editRow').value);
  const name = document.getElementById('editName').value.trim();
  const date = document.getElementById('editDate').value;
  const amount = document.getElementById('editAmount').value;

  if (!name) { showToast('Please enter an expense name.', 'error'); return; }
  if (!date) { showToast('Please enter a date.', 'error'); return; }
  if (!amount || parseFloat(amount) <= 0) { showToast('Please enter a valid amount.', 'error'); return; }

  const btn = document.getElementById('editSaveBtn');
  btn.disabled = true;
  btn.textContent = 'Saving...';

  try {
    await updateExpenseRow(row, name, date, parseFloat(amount).toFixed(2));
    const e = expenses.find(x => x.row === row);
    if (e) {
      e.name = name;
      e.date = date;
      e.amount = parseFloat(amount);
    }
    closeEditModal();
    renderExpenses();
    showToast('Expense updated!', 'success');
  } catch (err) {
    showToast('Error: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = '\ud83d\udcbe Save';
  }
}

// ── Delete Expense ──
function confirmDelete(row) {
  const e = expenses.find(x => x.row === row);
  if (!e) return;
  document.getElementById('deleteRow').value = row;
  document.getElementById('deleteName').textContent = e.name;
  document.getElementById('deleteModal').classList.add('show');
}

function closeDeleteModal() {
  document.getElementById('deleteModal').classList.remove('show');
}

async function confirmDeleteExpense() {
  const row = parseInt(document.getElementById('deleteRow').value);
  const btn = document.getElementById('deleteConfirmBtn');
  btn.disabled = true;
  btn.textContent = 'Deleting...';

  try {
    await deleteExpenseRow(row);
    closeDeleteModal();
    selected.delete(row);
    await refreshExpenses();
    showToast('Expense deleted!', 'success');
  } catch (err) {
    showToast('Error: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = '\ud83d\uddd1\ufe0f Delete';
  }
}

// ── Toast ──
function showToast(msg, type) {
  const el = document.getElementById('toast');
  el.textContent = (type === 'success' ? '✅ ' : '❌ ') + msg;
  el.className = 'toast toast-' + type + ' show';
  clearTimeout(el._timer);
  el._timer = setTimeout(() => el.classList.remove('show'), 4000);
}

// ── Loading ──
function showLoading(msg) {
  document.getElementById('loadingText').textContent = msg || 'Loading...';
  document.getElementById('loadingOverlay').classList.add('show');
}
function hideLoading() {
  document.getElementById('loadingOverlay').classList.remove('show');
}

// ── Init ──
(function init() {
  if (!requireAuth()) return;

  // Set user info
  const email = sessionStorage.getItem('hsa_user_email') || '';
  const name = sessionStorage.getItem('hsa_user_name') || '';
  const picture = sessionStorage.getItem('hsa_user_picture') || '';

  document.getElementById('userEmail').textContent = email;
  const avatarEl = document.getElementById('userAvatar');
  if (picture) {
    avatarEl.innerHTML = '<img src="' + escAttr(picture) + '" alt="avatar">';
  } else {
    const initials = name.split(' ').map(w => w[0]).join('').substring(0, 2).toUpperCase();
    avatarEl.textContent = initials || '??';
  }

  // Default date
  document.getElementById('expenseDate').valueAsDate = new Date();

  // Dropzone
  initDropzone();

  // Nav tabs
  document.querySelectorAll('.nav-links button').forEach(btn => {
    btn.addEventListener('click', () => switchView(btn.dataset.view));
  });

  // Filters
  document.querySelectorAll('.filter-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentFilter = btn.dataset.filter;
      selected.clear();
      renderExpenses();
    });
  });

  // Year filter
  document.getElementById('yearFilter').addEventListener('change', function () {
    currentYear = this.value;
    selected.clear();
    renderExpenses();
  });

  // Sort headers
  document.querySelectorAll('.data-table th[data-col]').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.col;
      if (sortCol === col) {
        sortAsc = !sortAsc;
      } else {
        sortCol = col;
        sortAsc = true;
      }
      renderExpenses();
    });
  });

  // Modal overlay click to close
  document.getElementById('modal').addEventListener('click', e => {
    if (e.target === e.currentTarget) closeModal();
  });
  document.getElementById('editModal').addEventListener('click', e => {
    if (e.target === e.currentTarget) closeEditModal();
  });
  document.getElementById('deleteModal').addEventListener('click', e => {
    if (e.target === e.currentTarget) closeDeleteModal();
  });
})();
