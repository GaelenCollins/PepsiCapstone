var isViewerMode = (typeof require === 'undefined');
var ipcRenderer = null;
var getErrorTypeLabel = function(x) { return x || 'Mismatch'; };

function viewerInvoke(channel, arg) {
  var opts = { headers: { 'Content-Type': 'application/json' } };
  if (channel === 'authenticate') {
    return fetch('/api/authenticate', { method: 'POST', ...opts, body: JSON.stringify({ password: arg }) })
      .then(function(r) { return r.json(); }).then(function(d) { return !!d.ok; });
  }
  if (channel === 'check-auth') {
    return fetch('/api/check-auth').then(function(r) { return r.json(); }).then(function(d) { return d.authenticated; });
  }
  if (channel === 'get-product-master') {
    return fetch('/api/product-master').then(function(r) { return r.json(); });
  }
  if (channel === 'save-product-master') {
    return fetch('/api/product-master', { method: 'POST', ...opts, body: JSON.stringify({ products: arg }) })
      .then(function(r) {
        if (r.ok) return true;
        return r.json().then(function(d) {
          var msg = (d && d.error) || 'Failed';
          var err = new Error(msg);
          err.status = r.status;
          throw err;
        });
      });
  }
  if (channel === 'override-mismatch') {
    return fetch('/api/override', { method: 'POST', ...opts, body: JSON.stringify({ id: arg }) })
      .then(function(r) { if (!r.ok) return r.json().then(function(d) { throw new Error(d.error || 'Failed'); }); return true; });
  }
  if (channel === 'get-email-config') {
    return fetch('/api/email-config').then(function(r) { return r.json(); });
  }
  if (channel === 'save-email-config') {
    return fetch('/api/email-config', { method: 'POST', ...opts, body: JSON.stringify(arg) })
      .then(function(r) { if (!r.ok) return r.json().then(function(d) { throw new Error(d.error || 'Failed'); }); return true; });
  }
  if (channel === 'export-data') {
    var start = (arg && arg.startDate) || '';
    var end = (arg && arg.endDate) || '';
    return fetch('/api/export?startDate=' + encodeURIComponent(start) + '&endDate=' + encodeURIComponent(end))
      .then(function(r) { return r.json(); });
  }
  if (channel === 'get-mismatches') {
    var q = arg || {};
    var params = new URLSearchParams();
    if (q.limit != null) params.set('limit', q.limit);
    if (q.startDate) params.set('startDate', q.startDate);
    if (q.endDate) params.set('endDate', q.endDate);
    return fetch('/api/mismatches?' + params.toString()).then(function(r) { return r.json(); });
  }
  if (channel === 'get-statistics') {
    var q = arg || {};
    var params = new URLSearchParams();
    if (q.startDate) params.set('startDate', q.startDate);
    if (q.endDate) params.set('endDate', q.endDate);
    return fetch('/api/statistics?' + params.toString()).then(function(r) { return r.json(); });
  }
  if (channel === 'clear-alarm') {
    return Promise.resolve(true);
  }
  if (channel === 'viewer-state' || channel === 'log-scan' || channel === 'log-mismatch' || channel === 'lookup-barcode') {
    return Promise.resolve(null);
  }
  return Promise.reject(new Error('Unknown channel: ' + channel));
}

if (!isViewerMode) {
  ipcRenderer = require('electron').ipcRenderer;
  getErrorTypeLabel = require('./services/error-types').getErrorTypeLabel;
} else {
  ipcRenderer = { invoke: viewerInvoke };
}

let isAuthenticated = false;
let productMaster = [];
let mismatches = [];
let pendingAction = null;

// Date range state (YYYY-MM-DD or null). Main view uses history range for stats + data.
let historyDateStart = null;
let historyDateEnd = null;
let statsDateStart = null;
let statsDateEnd = null;

function getDateRangeOptions() {
  return {
    startDate: historyDateStart || undefined,
    endDate: historyDateEnd || undefined
  };
}

// Barcode scanner state (continuous mode: scanner sends chars + Enter as HID keyboard)
let lastScannedLpn = '';
let lastScannedProduct = '';
let lastScanTime = 0;  // timestamp of last successful scan (for "No code for Xs")
const MISMATCH_DELAY_MS = 5000;  // wait 5s before showing mismatch (changeover buffer)
let pendingMismatchTimeoutId = null;
let pendingMismatchData = null;

/** Stack light clears via main process; mismatch path calls triggerAlarm but match never did until we added this. */
function clearPhysicalAlarmOnMatch() {
  ipcRenderer.invoke('clear-alarm').catch(function() {});
}

// PRODUCT SCANNER NEVER SCANS A LETTER. If the code has ANY letter (A-Z, a-z), it is an LPN — NEVER treat it as a product barcode.
function codeHasLetters(s) {
  return typeof s === 'string' && /[A-Za-z]/.test(s);
}
// LPN: the first 7 characters of the last 10 characters are the legacy item name (7 digits).
function getLegacyFromLpn(lpn) {
  if (typeof lpn !== 'string' || lpn.length < 10) return null;
  return lpn.slice(-10, -3);  // first 7 of last 10
}
// Normalize barcode for comparison (strip leading zeros so 012000811197 matches 0012000811197)
function normalizeBarcode(s) {
  var t = (s || '').toString().trim().replace(/^0+/, '') || '0';
  return t;
}
function findProductByBarcode(barcode) {
  var s = (barcode || '').toString().trim();
  if (!s) return null;
  var norm = normalizeBarcode(s);
  return productMaster.find(function(p) {
    var b = (p.barcode != null ? p.barcode : (p.sku != null ? p.sku : '')).toString().trim();
    if (b === s) return true;
    return normalizeBarcode(b) === norm;
  });
}
function findProductByLegacyItemName(legacyItemName) {
  var s = (legacyItemName || '').toString().trim();
  if (!s) return null;
  return productMaster.find(function(p) {
    var leg = (p.legacyItemName != null ? p.legacyItemName : (p.sku != null ? p.sku : '')).toString().trim();
    return leg === s;
  });
}
// Returns { result: 'match'|'lpn_invalid_sku'|'lpn_wrong_product' }. Only match when product list says they match.
function compareLpnToProduct(lpn, productBarcode) {
  if (!productBarcode || !(productBarcode + '').trim()) return { result: 'lpn_wrong_product', errorType: 'lpn_wrong_product' };
  var legacyFromLpn = getLegacyFromLpn(lpn);
  if (!legacyFromLpn) return { result: 'lpn_invalid_sku', errorType: 'lpn_invalid_sku' };
  var product = findProductByBarcode(productBarcode);
  if (!product) return { result: 'lpn_wrong_product', errorType: 'lpn_wrong_product' };
  var legacyItemName = (product.legacyItemName != null ? product.legacyItemName : (product.sku != null ? product.sku : '')).toString().trim();
  if (legacyFromLpn !== legacyItemName) return { result: 'lpn_wrong_product', errorType: 'lpn_wrong_product' };
  return { result: 'match' };
}
let lastLookedUpProductCode = '';  // only lookup when product changes (save API calls)
let barcodeBuffer = '';
let barcodeFirstKeyTime = 0;
const BARCODE_MAX_MS = 250;  // treat as barcode if Enter within this ms of first key
const SCANNING_IDLE_MS = 100; // "Currently scanning" for this long after last key

function formatSecondsSince(ts) {
  if (!ts || ts <= 0) return '—';
  const sec = Math.floor((Date.now() - ts) / 1000);
  if (sec < 60) return sec + 's';
  const m = Math.floor(sec / 60);
  const s = sec % 60;
  if (m < 60) return s ? m + 'm ' + s + 's' : m + 'm 0s';
  const h = Math.floor(m / 60);
  const mm = m % 60;
  return (mm || s) ? h + 'h ' + mm + 'm ' + s + 's' : h + 'h 0m 0s';
}

// Initialize app
document.addEventListener('DOMContentLoaded', async () => {
  if (isViewerMode) {
    runViewerMode();
    setupEventListeners();
    setDefaultDates();
    fetch('/api/product-master').then(function(r) { return r.json(); }).then(function(p) { productMaster = p || []; });
    fetch('/api/check-auth').then(function(r) { return r.json(); }).then(function(d) { isAuthenticated = !!d.authenticated; });
    return;
  }
  setupEventListeners();
  setupBarcodeCapture();
  setDefaultDates();
  await loadData();
  updateUI();
  updateComparisonStatusForSingleScanner();
  updateScannerStatusLines();
  pushViewerState();
  setInterval(pushViewerState, 2500);
  ipcRenderer.on('mismatches-changed', function() {
    loadData()
      .then(function() {
        updateUI();
        updateAdvancedStatistics();
      })
      .catch(function() {});
  });
});

function setDefaultDates() {
  const today = new Date();
  const lastMonth = new Date(today);
  lastMonth.setMonth(lastMonth.getMonth() - 1);
  const fmt = d => d.toISOString().slice(0, 10);
  const todayStr = fmt(today);
  const lastMonthStr = fmt(lastMonth);
  const historyStartEl = document.getElementById('history-date-start');
  const historyEndEl = document.getElementById('history-date-end');
  const statsEndEl = document.getElementById('stats-date-end');
  if (historyStartEl && !historyStartEl.value) {
    historyStartEl.value = lastMonthStr;
    historyDateStart = lastMonthStr;
  }
  if (historyEndEl && !historyEndEl.value) {
    historyEndEl.value = todayStr;
    historyDateEnd = todayStr;
  }
  const statsStartEl = document.getElementById('stats-date-start');
  if (statsStartEl && !statsStartEl.value) statsStartEl.value = lastMonthStr;
  if (statsEndEl && !statsEndEl.value) statsEndEl.value = todayStr;
}

// Load initial data
async function loadData() {
  try {
    productMaster = await ipcRenderer.invoke('get-product-master');
    if (productMaster.length === 0) {
      productMaster = [
        { name: 'PEPSI COL CAN 12OZ 12P2C FM', sku: '300011736/0083774' },
        { name: 'MTN DEW ORG CAN 12OZ 36P1C CB', sku: '300011737/0083775' }
      ];
    }
    const opts = getDateRangeOptions();
    mismatches = await ipcRenderer.invoke('get-mismatches', { limit: 500, ...opts });
    updateStatistics();
    updateAdvancedStatistics();
  } catch (error) {
    console.error('Error loading data:', error);
  }
}

// Setup event listeners
function setupEventListeners() {
  document.getElementById('edit-item-master-btn').addEventListener('click', showItemMasterModal);
  document.getElementById('password-submit').addEventListener('click', handlePasswordSubmit);
  document.getElementById('password-cancel').addEventListener('click', closePasswordModal);
  document.getElementById('password-input').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') handlePasswordSubmit();
  });
  document.getElementById('close-item-master').addEventListener('click', closeItemMasterModal);
  document.getElementById('add-product-btn').addEventListener('click', showAddProductModal);
  document.getElementById('save-product-btn').addEventListener('click', saveNewProduct);
  document.getElementById('cancel-add-product').addEventListener('click', closeAddProductModal);
  var refreshItemMasterBtn = document.getElementById('refresh-item-master-btn');
  if (refreshItemMasterBtn) refreshItemMasterBtn.addEventListener('click', refreshItemMaster);
  var exportCsvBtn = document.getElementById('export-item-master-csv');
  if (exportCsvBtn) exportCsvBtn.addEventListener('click', exportItemMasterCsv);
  var importCsvBtn = document.getElementById('import-item-master-csv');
  var importFileInput = document.getElementById('import-item-master-file');
  if (importCsvBtn && importFileInput) {
    importCsvBtn.addEventListener('click', function() { importFileInput.click(); });
    importFileInput.addEventListener('change', function() {
      var f = importFileInput.files && importFileInput.files[0];
      if (f) importItemMasterCsv(f);
      importFileInput.value = '';
    });
  }

  // Date range - history (drives main stats + table) and stats tab
  const historyApply = document.getElementById('history-date-apply');
  if (historyApply) historyApply.addEventListener('click', applyHistoryDateRange);
  const statsApply = document.getElementById('stats-date-apply');
  if (statsApply) statsApply.addEventListener('click', applyStatsDateRange);

  // Export
  document.getElementById('export-data-btn').addEventListener('click', openExportModal);
  document.getElementById('export-csv-btn').addEventListener('click', exportCsv);
  document.getElementById('export-cancel-btn').addEventListener('click', () => {
    document.getElementById('export-modal').classList.remove('active');
    if (window.refocusBarcodeInput) setTimeout(window.refocusBarcodeInput, 100);
  });

  // Settings (email Oscar on error)
  document.getElementById('settings-btn').addEventListener('click', openSettingsModal);
  document.getElementById('settings-save-btn').addEventListener('click', saveSettings);
  document.getElementById('settings-cancel-btn').addEventListener('click', () => {
    document.getElementById('settings-modal').classList.remove('active');
    if (window.refocusBarcodeInput) setTimeout(window.refocusBarcodeInput, 100);
  });

  // Mobile bottom tabs (Main View | Error History | Stats)
  const mobileTabs = document.getElementById('mobile-bottom-tabs');
  if (mobileTabs) {
    mobileTabs.querySelectorAll('.mobile-tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const tab = btn.getAttribute('data-mobile-tab');
        document.body.classList.remove('mobile-tab-main', 'mobile-tab-error-history', 'mobile-tab-stats');
        document.body.classList.add('mobile-tab-' + tab);
        mobileTabs.querySelectorAll('.mobile-tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        if (tab === 'stats') updateAdvancedStatistics();
      });
    });
    // Default mobile tab when in mobile view
    if (window.matchMedia('(max-width: 640px)').matches) {
      document.body.classList.add('mobile-tab-main');
    }
  }

  document.querySelectorAll('.modal').forEach(modal => {
    modal.addEventListener('click', (e) => {
      if (e.target === modal) modal.classList.remove('active');
    });
  });
}

// Capture barcode from WoneNice (or any HID keyboard-mode) scanner. One input = Product scanner for now.
function setupBarcodeCapture() {
  const input = document.getElementById('barcode-input-product');
  if (!input) return;

  const focusInput = () => {
    if (document.visibilityState === 'visible' && !document.getElementById('password-modal').classList.contains('active')) {
      input.focus();
    }
  };

  input.addEventListener('focus', () => input.setAttribute('data-focused', '1'));
  input.addEventListener('blur', () => input.removeAttribute('data-focused'));
  window.addEventListener('focus', focusInput);
  document.addEventListener('visibilitychange', focusInput);
  setTimeout(focusInput, 500);
  window.refocusBarcodeInput = focusInput;

  // Click anywhere in main view to focus barcode input (so next scan goes into app)
  const mainTab = document.getElementById('main-tab');
  if (mainTab) {
    mainTab.addEventListener('click', () => { focusInput(); });
  }

  input.addEventListener('keydown', (e) => {
    const now = Date.now();
    if (e.key === 'Enter') {
      e.preventDefault();
      if (barcodeBuffer.length >= 4 && (now - barcodeFirstKeyTime) < BARCODE_MAX_MS) {
        const newCode = barcodeBuffer.trim();
        // LETTERS IN CODE = LPN. PRODUCT SCANNER NEVER SCANS A LETTER. NEVER put a code with letters in the product box.
        const hasLetters = codeHasLetters(newCode);
        if (hasLetters) {
          // Code has letters → it is an LPN. Do not clear the 5s timer here; only clear when we get a match below.
          lastScannedLpn = newCode;
          const legacySku = getLegacyFromLpn(lastScannedLpn);
          const lpnProduct = legacySku ? findProductByLegacyItemName(legacySku) : null;
          const lpnResolvedName = (lpnProduct && (lpnProduct.name || '').trim()) ? (lpnProduct.name || '').trim() : (legacySku || lastScannedLpn);
          const resEl = document.getElementById('scanner1-product-resolved');
          const codeEl = document.getElementById('scanner1-product-code');
          if (resEl) resEl.textContent = lpnResolvedName;
          if (codeEl) codeEl.textContent = legacySku || lastScannedLpn;
          const comparison = compareLpnToProduct(lastScannedLpn, lastScannedProduct);
          if (comparison.result === 'match') {
            if (pendingMismatchTimeoutId) {
              clearTimeout(pendingMismatchTimeoutId);
              pendingMismatchTimeoutId = null;
              pendingMismatchData = null;
              clearPendingMismatchState();
            }
            updateScannerBoxes(lastScannedLpn, lastScannedProduct || '—', 'match');
            clearPhysicalAlarmOnMatch();
          } else {
            if (!pendingMismatchTimeoutId) {
              showPendingMismatchState();
              pendingMismatchTimeoutId = setTimeout(function() {
                pendingMismatchTimeoutId = null;
                pendingMismatchData = null;
                var lpn = lastScannedLpn;
                var prod = lastScannedProduct || '—';
                var comp = compareLpnToProduct(lpn, prod);
                if (comp.result !== 'match') {
                  updateScannerBoxes(lpn, prod, comp.result, comp.errorType);
                  logError(prod, lpn, comp.errorType);
                }
              }, MISMATCH_DELAY_MS);
            }
            pendingMismatchData = { lpnValue: lastScannedLpn, productValue: lastScannedProduct || '—', result: comparison.result, errorType: comparison.errorType };
            updatePendingMismatchDisplay(lastScannedLpn, lastScannedProduct || '—');
          }
        } else {
          // Digits only → product barcode. Do not reset the 5s mismatch timer on new scan; only clear it when we get a match.
          lastScannedProduct = newCode;
          lastLookedUpProductCode = newCode;
          lastScanTime = Date.now();
          const codeEl = document.getElementById('scanner2-product-code');
          const resolvedEl = document.getElementById('scanner2-product-resolved');
          if (codeEl) codeEl.textContent = lastScannedProduct;
          if (resolvedEl) resolvedEl.textContent = '…';
          const hintEl = document.getElementById('scanner-hint');
          if (hintEl) hintEl.classList.add('scanner-hint-hidden');
          updateScannerStatusLines();
          updateProductResolvedName(lastScannedProduct).then(function() {
            if (lastScannedLpn) {
              var comparison = compareLpnToProduct(lastScannedLpn, lastScannedProduct);
              if (comparison.result === 'match') {
                if (pendingMismatchTimeoutId) {
                  clearTimeout(pendingMismatchTimeoutId);
                  pendingMismatchTimeoutId = null;
                  pendingMismatchData = null;
                  clearPendingMismatchState();
                }
                updateScannerBoxes(lastScannedLpn, lastScannedProduct || '—', 'match');
                clearPhysicalAlarmOnMatch();
              } else {
                if (!pendingMismatchTimeoutId) {
                  showPendingMismatchState();
                  pendingMismatchTimeoutId = setTimeout(function() {
                    pendingMismatchTimeoutId = null;
                    pendingMismatchData = null;
                    var lpn = lastScannedLpn;
                    var prod = lastScannedProduct || '—';
                    var comp = compareLpnToProduct(lpn, prod);
                    if (comp.result !== 'match') {
                      updateScannerBoxes(lpn, prod, comp.result, comp.errorType);
                      logError(prod, lpn, comp.errorType);
                    }
                  }, MISMATCH_DELAY_MS);
                }
                pendingMismatchData = { lpnValue: lastScannedLpn, productValue: lastScannedProduct || '—', result: comparison.result, errorType: comparison.errorType };
                updatePendingMismatchDisplay(lastScannedLpn, lastScannedProduct || '—');
              }
            } else {
              document.getElementById('scanner1-status').textContent = '—';
              document.getElementById('scanner2-status').textContent = '—';
              document.getElementById('scanner1-box').classList.remove('match', 'mismatch', 'pending-mismatch');
              document.getElementById('scanner2-box').classList.remove('match', 'mismatch', 'pending-mismatch');
            }
            focusInput();
          });
        }
      }
      barcodeBuffer = '';
      barcodeFirstKeyTime = 0;
      return;
    }
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      if (barcodeBuffer.length === 0) barcodeFirstKeyTime = now;
      barcodeBuffer += e.key;
      updateScannerStatusLines();
    }
  });

  setInterval(() => {
    if (barcodeBuffer.length === 0) updateScannerStatusLines();
  }, 80);
  setInterval(() => {
    if (lastScannedProduct && barcodeBuffer.length === 0) updateScannerStatusLines();
  }, 1000);
}

function updateScannerStatusLines() {
  const now = Date.now();
  const productLine = document.getElementById('scanner2-status-line');
  const lpnLine = document.getElementById('scanner1-status-line');
  if (productLine) {
    if (barcodeBuffer.length > 0 || (barcodeFirstKeyTime && (now - barcodeFirstKeyTime) < SCANNING_IDLE_MS)) {
      productLine.textContent = 'Currently scanning';
      productLine.className = 'scanner-status-line scanning';
    } else {
      productLine.textContent = lastScannedProduct
        ? 'No code for ' + formatSecondsSince(lastScanTime)
        : 'No code currently';
      productLine.className = 'scanner-status-line no-code';
    }
  }
  if (lpnLine) {
    lpnLine.textContent = 'Disconnected';
    lpnLine.className = 'scanner-status-line disconnected';
  }
}

function applyHistoryDateRange() {
  historyDateStart = document.getElementById('history-date-start').value || null;
  historyDateEnd = document.getElementById('history-date-end').value || null;
  loadData();
}


function applyStatsDateRange() {
  statsDateStart = document.getElementById('stats-date-start').value || null;
  statsDateEnd = document.getElementById('stats-date-end').value || null;
  updateAdvancedStatistics();
}

// Show password modal
function showPasswordModal(actionCallback) {
    pendingAction = actionCallback; // Store the action to execute after auth
    document.getElementById('password-modal').classList.add('active');
    document.getElementById('password-input').focus();
}

// Close password modal
function closePasswordModal() {
  document.getElementById('password-modal').classList.remove('active');
  document.getElementById('password-input').value = '';
  document.getElementById('password-error').classList.remove('show');
  pendingAction = null;
  if (window.refocusBarcodeInput) setTimeout(window.refocusBarcodeInput, 100);
}

// Handle password submission
async function handlePasswordSubmit() {
    const password = document.getElementById('password-input').value;
    const errorDiv = document.getElementById('password-error');
    
    try {
        isAuthenticated = await ipcRenderer.invoke('authenticate', password);
        if (isAuthenticated) {
            var next = pendingAction;
            closePasswordModal();
            if (next) {
                next();
            }
        } else {
            errorDiv.textContent = 'Incorrect password. Please try again.';
            errorDiv.classList.add('show');
        }
    } catch (error) {
        errorDiv.textContent = 'Error authenticating. Please try again.';
        errorDiv.classList.add('show');
    }
}

// Show item master modal
async function showItemMasterModal() {
    if (!isAuthenticated) {
        showPasswordModal(async () => {
            await loadItemMaster();
            document.getElementById('item-master-modal').classList.add('active');
        });
        return;
    }
    
    await loadItemMaster();
    document.getElementById('item-master-modal').classList.add('active');
}

// Refresh item master from server and reload table
async function refreshItemMaster() {
  try {
    productMaster = await ipcRenderer.invoke('get-product-master');
    await loadItemMaster();
  } catch (e) {
    alert('Refresh failed: ' + (e.message || e));
  }
}

// Load item master table
async function loadItemMaster() {
    const tbody = document.getElementById('item-master-body');
    tbody.innerHTML = '';
    
    productMaster.forEach((product, index) => {
        const row = document.createElement('tr');
        const barcode = product.barcode != null ? product.barcode : '';
        const legacy = product.legacyItemName != null ? product.legacyItemName : (product.sku != null ? product.sku : '');
        const name = product.name != null ? product.name : '';
        row.innerHTML = `
            <td>${escapeHtml(barcode)}</td>
            <td>${escapeHtml(legacy)}</td>
            <td>${escapeHtml(name)}</td>
            <td>
                <button class="delete-product-btn" onclick="deleteProduct(${index})">Delete</button>
            </td>
        `;
        tbody.appendChild(row);
    });
}

// Delete product
async function deleteProduct(index) {
    if (confirm('Are you sure you want to delete this product?')) {
        productMaster.splice(index, 1);
        await saveProductMaster();
        await loadItemMaster();
    }
}

// Show add product modal
function showAddProductModal() {
    document.getElementById('add-product-modal').classList.add('active');
    (document.getElementById('new-product-barcode') || {}).value = '';
    (document.getElementById('new-product-legacy') || {}).value = '';
    (document.getElementById('new-product-name') || {}).value = '';
}

// Close add product modal
function closeAddProductModal() {
  document.getElementById('add-product-modal').classList.remove('active');
  if (window.refocusBarcodeInput) setTimeout(window.refocusBarcodeInput, 100);
}

function escapeHtml(s) {
    if (s == null) return '';
    var str = String(s);
    var div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// Save new product (barcode, legacyItemName, name = mfg description)
async function saveNewProduct() {
    const barcode = (document.getElementById('new-product-barcode') || {}).value.trim();
    const legacyItemName = (document.getElementById('new-product-legacy') || {}).value.trim();
    const name = (document.getElementById('new-product-name') || {}).value.trim();
    if (!barcode || !legacyItemName || !name) {
        alert('Please fill in Barcode, Legacy item name (8), and Product name.');
        return;
    }
    if (legacyItemName.length !== 7) {
        alert('Legacy item name must be exactly 7 characters.');
        return;
    }
    productMaster.push({ barcode, legacyItemName, name });
    await saveProductMaster();
    closeAddProductModal();
    await loadItemMaster();
}

// Save product master (in viewer mode, prompts for password if not authenticated)
async function saveProductMaster() {
    try {
        await ipcRenderer.invoke('save-product-master', productMaster);
    } catch (error) {
        if (isViewerMode && (error.status === 401 || (error.message && error.message.toLowerCase().indexOf('authenticated') !== -1))) {
            showPasswordModal(function() {
                saveProductMaster();
            });
            return;
        }
        alert('Error saving product master: ' + error.message);
    }
}

// Item master CSV: barcode, legacy_item_name, name (mfg description). Same source for app and viewer.
function exportItemMasterCsv() {
    var header = 'barcode,legacy_item_name,name';
    var rows = productMaster.map(function(p) {
        var barcode = p.barcode != null ? p.barcode : '';
        var legacy = p.legacyItemName != null ? p.legacyItemName : (p.sku != null ? p.sku : '');
        var name = p.name != null ? p.name : '';
        return escapeCsvField(barcode) + ',' + escapeCsvField(legacy) + ',' + escapeCsvField(name);
    });
    var csv = header + '\n' + rows.join('\n');
    var blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'item-master.csv';
    a.click();
    URL.revokeObjectURL(url);
}
function escapeCsvField(s) {
    if (s == null) return '""';
    var str = String(s);
    if (/[",\n\r]/.test(str)) return '"' + str.replace(/"/g, '""') + '"';
    return str;
}
function importItemMasterCsv(file) {
    var reader = new FileReader();
    reader.onload = function() {
        var text = String(reader.result).replace(/^\uFEFF/, ''); // strip UTF-8 BOM
        var lines = text.split(/\r?\n/);
        if (lines.length < 2) {
            alert('File is empty or has only a header row.');
            return;
        }
        // Row 0 is always the column titles row — do not import it
        var header = parseCsvLine(lines[0]);
        var colBarcode = -1, colLegacy = -1, colName = -1;
        for (var h = 0; h < header.length; h++) {
            var raw = (header[h] || '').replace(/\uFEFF/g, '').trim();
            var cell = raw.toLowerCase().replace(/\s/g, '').replace(/["\u201C\u201D\u201E\u201F]/g, '');
            // Barcode: Barcode Number, UPC, GTIN, etc.
            if (colBarcode < 0 && (cell === 'barcodenumber' || cell === 'upc' || cell === 'gtin' || cell === 'barcode' ||
                (cell.indexOf('barcode') !== -1 && cell.indexOf('number') !== -1))) colBarcode = h;
            // Legacy / SKU: Legacy Item Name, SKU, Item Name, Item Number, etc.
            if (colLegacy < 0 && (cell === 'legacyitemname' || cell === 'legacy_item_name' || cell === 'sku' ||
                cell === 'itemname' || cell === 'item_name' || cell === 'itemnumber' || cell === 'item_number' ||
                (cell.indexOf('legacy') !== -1 && cell.indexOf('item') !== -1))) colLegacy = h;
            // Product name/description: MFG Description, Description, Product Name, etc.
            if (colName < 0 && (cell === 'mfgdescription' || cell === 'mfg_description' || cell === 'description' ||
                cell === 'productname' || cell === 'product_name' || cell === 'productdescription' || cell === 'name' ||
                (cell.indexOf('mfg') !== -1 && cell.indexOf('description') !== -1) ||
                (cell.indexOf('product') !== -1 && (cell.indexOf('name') !== -1 || cell.indexOf('description') !== -1)))) colName = h;
        }
        // Fallback: "Have Pallet Tag?, Legacy Item Name, MFG Description, Barcode Number" = cols 1,2,3
        if ((colBarcode < 0 || colLegacy < 0 || colName < 0) && header.length >= 4) {
            var first = (header[0] || '').toLowerCase();
            if (first.indexOf('pallet') !== -1 || first.indexOf('tag') !== -1 || first.indexOf('have') !== -1) {
                if (colLegacy < 0) colLegacy = 1;
                if (colName < 0) colName = 2;
                if (colBarcode < 0) colBarcode = 3;
            }
        }
        var parsed = [];
        for (var i = 1; i < lines.length; i++) {
            var row = parseCsvLine(lines[i]);
            var barcode = colBarcode >= 0 && row[colBarcode] != null ? String(row[colBarcode]).trim() : '';
            var legacy = colLegacy >= 0 && row[colLegacy] != null ? String(row[colLegacy]).trim() : '';
            var name = colName >= 0 && row[colName] != null ? String(row[colName]).trim() : '';
            if (!barcode && !legacy && !name) continue;
            parsed.push({ barcode: barcode, legacyItemName: legacy, name: name });
        }
        parsed = parsed.filter(function(p) {
            var b = (p.barcode || '').trim();
            var l = (p.legacyItemName || '').trim();
            var n = (p.name || '').trim();
            return b !== '' || l !== '' || n !== '';
        });
        if (parsed.length === 0) {
            var found = [];
            if (colBarcode >= 0) found.push('Barcode');
            if (colLegacy >= 0) found.push('Legacy/SKU');
            if (colName >= 0) found.push('Product name');
            var msg = 'No valid rows found. First row is treated as headers.\n\n';
            msg += 'Detected columns: ' + (found.length ? found.join(', ') : 'none') + '.\n\n';
            msg += 'Expected at least one of these header names:\n';
            msg += '• Barcode: "Barcode Number", "UPC", "GTIN", or "Barcode"\n';
            msg += '• Legacy/SKU: "Legacy Item Name", "SKU", "Item Name", or "Item Number"\n';
            msg += '• Product name: "MFG Description", "Description", "Product Name", or "Name"\n\n';
            msg += 'Rename your CSV columns to match, or ensure data rows are not empty.';
            alert(msg);
            return;
        }
        productMaster = parsed;
        saveProductMaster().then(function() {
            loadItemMaster();
            alert('Imported ' + parsed.length + ' product(s). Column titles row and all blank rows were excluded.');
        }).catch(function(e) {
            alert('Import failed: ' + (e.message || e));
        });
    };
    reader.readAsText(file, 'UTF-8');
}
function parseCsvLine(line) {
    var out = [];
    var i = 0;
    while (i < line.length) {
        if (line[i] === '"') {
            var end = line.indexOf('""', i + 1);
            var s = '';
            i++;
            while (i < line.length) {
                if (line[i] === '"') {
                    if (line[i + 1] === '"') { s += '"'; i += 2; }
                    else { i++; break; }
                } else { s += line[i]; i++; }
            }
            out.push(s);
        } else {
            var j = line.indexOf(',', i);
            if (j === -1) j = line.length;
            out.push(line.slice(i, j));
            i = j + 1;
        }
    }
    return out;
}

// Close item master modal
function closeItemMasterModal() {
  document.getElementById('item-master-modal').classList.remove('active');
  if (window.refocusBarcodeInput) setTimeout(window.refocusBarcodeInput, 100);
}

// Settings (email Oscar on error + SMTP)
async function openSettingsModal() {
  try {
    const config = await ipcRenderer.invoke('get-email-config');
    document.getElementById('settings-email-on-error').checked = !!config.emailOscarOnError;
    document.getElementById('settings-oscar-email').value = config.oscarEmail || '';
    document.getElementById('settings-use-mailgun').checked = config.useMailgun !== false;
    document.getElementById('settings-mailgun-api-key').value = config.mailgunApiKey || '';
    document.getElementById('settings-mailgun-domain').value = config.mailgunDomain || '';
    document.getElementById('settings-smtp-host').value = config.smtpHost || '';
    document.getElementById('settings-smtp-port').value = config.smtpPort != null ? config.smtpPort : 587;
    document.getElementById('settings-smtp-secure').checked = !!config.smtpSecure;
    document.getElementById('settings-smtp-user').value = config.smtpUser || '';
    document.getElementById('settings-smtp-pass').value = config.smtpPass || '';
    document.getElementById('settings-from-email').value = config.fromEmail || '';
  } catch (e) {
    console.error('Load email config:', e);
  }
  document.getElementById('settings-modal').classList.add('active');
}

async function saveSettings() {
  try {
    await ipcRenderer.invoke('save-email-config', {
      emailOscarOnError: document.getElementById('settings-email-on-error').checked,
      oscarEmail: document.getElementById('settings-oscar-email').value.trim(),
      useMailgun: document.getElementById('settings-use-mailgun').checked,
      mailgunApiKey: document.getElementById('settings-mailgun-api-key').value,
      mailgunDomain: document.getElementById('settings-mailgun-domain').value.trim(),
      smtpHost: document.getElementById('settings-smtp-host').value.trim(),
      smtpPort: parseInt(document.getElementById('settings-smtp-port').value, 10) || 587,
      smtpSecure: document.getElementById('settings-smtp-secure').checked,
      smtpUser: document.getElementById('settings-smtp-user').value.trim(),
      smtpPass: document.getElementById('settings-smtp-pass').value,
      fromEmail: document.getElementById('settings-from-email').value.trim()
    });
    document.getElementById('settings-modal').classList.remove('active');
    if (window.refocusBarcodeInput) setTimeout(window.refocusBarcodeInput, 100);
  } catch (e) {
    alert('Failed to save settings: ' + (e.message || e));
  }
}

// Show only product title (stop after "bottle"), no company/category. Larger text; code stays below.
function formatProductDisplay(productInfo) {
  if (!productInfo || typeof productInfo !== 'string') return '';
  const firstPart = productInfo.split(' · ')[0].trim();
  const bottleIdx = firstPart.toLowerCase().indexOf(' bottle');
  return bottleIdx === -1 ? firstPart : firstPart.slice(0, bottleIdx + 7);
}

// Product name from product list only (no API).
async function updateProductResolvedName(code) {
  const el = document.getElementById('scanner2-product-resolved');
  const codeEl = document.getElementById('scanner2-product-code');
  if (!el) return;
  if (codeEl) codeEl.textContent = code || '—';
  if (!code) {
    el.textContent = '';
    pushViewerState();
    return;
  }
  var displayText = '';
  var product = findProductByBarcode(code);
  if (product && (product.name || '').trim()) {
    displayText = (product.name || '').trim();
  } else {
    displayText = 'Not in product list';
  }
  el.textContent = displayText;
  try {
    await ipcRenderer.invoke('log-scan', { code, product: displayText });
  } catch (logErr) {
    console.warn('Scan log failed', logErr);
  }
  pushViewerState();
}

// Export
function openExportModal() {
  const start = document.getElementById('export-date-start');
  const end = document.getElementById('export-date-end');
  const todayStr = new Date().toISOString().slice(0, 10);
  start.value = historyDateStart || '';
  end.value = historyDateEnd || todayStr;
  document.getElementById('export-modal').classList.add('active');
}

async function exportCsv() {
  const start = document.getElementById('export-date-start').value || null;
  const end = document.getElementById('export-date-end').value || null;
  try {
    const { rows } = await ipcRenderer.invoke('export-data', { startDate: start, endDate: end, format: 'csv' });
    const header = 'Timestamp,Product scanner,LPN scanner,Error type,Status,ID\n';
    const body = rows.map(r => {
      const ts = r.timestamp ? new Date(r.timestamp).toISOString() : '';
      const exp = (r.expected || '').replace(/"/g, '""');
      const act = (r.actual || '').replace(/"/g, '""');
      const errType = getErrorTypeLabel(r.errorType);
      return `"${ts}","${exp}","${act}","${errType}","${r.status || ''}","${r.id || ''}"`;
    }).join('\n');
    const csv = header + body;
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `scanner-export-${start || 'all'}-to-${end || 'all'}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    document.getElementById('export-modal').classList.remove('active');
    if (window.refocusBarcodeInput) setTimeout(window.refocusBarcodeInput, 100);
  } catch (err) {
    alert('Export failed: ' + (err.message || err));
  }
}

// Update UI
function updateUI() {
  var legacySku = lastScannedLpn ? getLegacyFromLpn(lastScannedLpn) : null;
  var lpnProduct = legacySku ? findProductByLegacyItemName(legacySku) : null;
  var lpnResolved = (lpnProduct && (lpnProduct.name || '').trim()) ? (lpnProduct.name || '').trim() : (legacySku || lastScannedLpn || '—');
  var s1Res = document.getElementById('scanner1-product-resolved');
  var s1Code = document.getElementById('scanner1-product-code');
  if (s1Res) s1Res.textContent = lpnResolved;
  if (s1Code) s1Code.textContent = legacySku || lastScannedLpn || '—';
  document.getElementById('scanner2-product-code').textContent = lastScannedProduct || '—';
  const hintEl = document.getElementById('scanner-hint');
  if (hintEl) hintEl.classList.toggle('scanner-hint-hidden', !!lastScannedProduct);
  updateProductResolvedName(lastScannedProduct || null);
  if (!lastScannedLpn && !lastScannedProduct) {
    document.getElementById('scanner1-status').textContent = '—';
    document.getElementById('scanner2-status').textContent = '—';
    updateComparisonStatusForSingleScanner();
  }
  updateScannerStatusLines();
  updateStatusTable();
  updateLastErrorLine();
  pushViewerState();
}

function updateLastErrorLine() {
  const el = document.getElementById('last-error-line');
  if (!el) return;
  if (mismatches.length === 0) {
    el.textContent = 'No errors recorded';
    el.classList.remove('last-error-line--recent');
  } else {
    const mostRecent = mismatches[0];
    const timeAgo = getTimeAgo(new Date(mostRecent.timestamp));
    el.textContent = 'Last error: ' + timeAgo;
    el.classList.add('last-error-line--recent');
  }
}

function pushViewerState() {
  if (isViewerMode || !ipcRenderer) return;
  const box1 = document.getElementById('scanner1-box');
  const box2 = document.getElementById('scanner2-box');
  const state = {
    scanner1ProductResolved: (document.getElementById('scanner1-product-resolved') || {}).textContent || '—',
    scanner1ProductCode: (document.getElementById('scanner1-product-code') || {}).textContent || '—',
    scanner2ProductCode: (document.getElementById('scanner2-product-code') || {}).textContent || '—',
    scanner2ProductResolved: (document.getElementById('scanner2-product-resolved') || {}).textContent || '',
    scanner1Status: (document.getElementById('scanner1-status') || {}).textContent || '—',
    scanner2Status: (document.getElementById('scanner2-status') || {}).textContent || '—',
    comparisonStatusText: (document.getElementById('comparison-status') || {}).textContent || '',
    comparisonStatusClass: (document.getElementById('comparison-status') || {}).className || '',
    scanner1StatusLine: (document.getElementById('scanner1-status-line') || {}).textContent || '',
    scanner2StatusLine: (document.getElementById('scanner2-status-line') || {}).textContent || '',
    lastErrorLineText: (document.getElementById('last-error-line') || {}).textContent || '',
    box1Match: box1 ? box1.classList.contains('match') : false,
    box1Mismatch: box1 ? box1.classList.contains('mismatch') : false,
    box2Match: box2 ? box2.classList.contains('match') : false,
    box2Mismatch: box2 ? box2.classList.contains('mismatch') : false
  };
  ipcRenderer.invoke('viewer-state', state);
}

function runViewerMode() {
  document.body.classList.add('viewer-mode');
  fetch('/api/host-info').then(function(r) { return r.json(); }).then(function(info) {
    var el = document.getElementById('join-from-phone');
    if (!el) return;
    var urls = (info.hostUrls && info.hostUrls.length) ? info.hostUrls : (info.hostUrl ? [info.hostUrl] : []);
    var mainUrl = info.hostUrl || urls[0] || 'http://localhost:' + (info.port || 3847);
    var urlList = urls.length > 1 ? ' Try: ' + urls.map(function(u) { return '<strong>' + u + '</strong>'; }).join(' or ') : '<strong>' + mainUrl + '</strong>';
    el.innerHTML = 'Join from your phone (same Wi&#8203;Fi): open ' + urlList + ' in the browser. Use <strong>http</strong> not https, port <strong>3847</strong>. <span class="join-from-phone-tip">Still not working? Phone and Mac must be on the same Wi&#8203;Fi (not guest network). On the router, turn off "AP isolation" or "client isolation" if enabled. Confirm the Mac\'s IP in System Settings &rarr; Network.</span>';
  }).catch(function() {});
  function poll() {
    fetch('/api/state').then(r => r.json()).then(applyViewerState).catch(() => {});
  }
  poll();
  setInterval(poll, 2000);
}

function applyViewerState(data) {
  if (!data) return;
  setEl('scanner1-product-resolved', data.scanner1ProductResolved != null ? data.scanner1ProductResolved : (data.scanner1ProductName || '—'));
  setEl('scanner1-product-code', data.scanner1ProductCode != null ? data.scanner1ProductCode : '—');
  setEl('scanner2-product-code', data.scanner2ProductCode);
  setEl('scanner2-product-resolved', data.scanner2ProductResolved);
  setEl('scanner1-status', data.scanner1Status);
  setEl('scanner2-status', data.scanner2Status);
  var statusEl = document.getElementById('comparison-status');
  if (statusEl) {
    statusEl.textContent = data.comparisonStatusText || '—';
    statusEl.className = data.comparisonStatusClass || 'comparison-status';
  }
  setEl('scanner1-status-line', data.scanner1StatusLine);
  setEl('scanner2-status-line', data.scanner2StatusLine);
  var lastErrEl = document.getElementById('last-error-line');
  if (lastErrEl) {
    lastErrEl.textContent = data.lastErrorLineText || (data.lastErrorAgo ? 'Last error: ' + data.lastErrorAgo : 'No errors recorded');
    lastErrEl.classList.toggle('last-error-line--recent', !!(data.mismatches && data.mismatches.length > 0));
  }
  var hintEl = document.getElementById('scanner-hint');
  if (hintEl) hintEl.classList.toggle('scanner-hint-hidden', !!(data.scanner2ProductCode && data.scanner2ProductCode !== '—'));
  var box1 = document.getElementById('scanner1-box');
  var box2 = document.getElementById('scanner2-box');
  if (box1) {
    box1.classList.toggle('match', !!data.box1Match);
    box1.classList.toggle('mismatch', !!data.box1Mismatch);
    box1.classList.remove('lpn-missing');
  }
  if (box2) {
    box2.classList.toggle('match', !!data.box2Match);
    box2.classList.toggle('mismatch', !!data.box2Mismatch);
    box2.classList.remove('lpn-missing');
  }
  if (data.stats) {
    setEl('total-mismatches', data.stats.total);
    setEl('today-mismatches', data.stats.today);
    setEl('resolved-count', data.stats.resolved);
    setEl('pending-count', data.stats.pending);
  }
  if (data.mismatches && Array.isArray(data.mismatches)) {
    var tbody = document.getElementById('status-table-body');
    if (tbody) {
      tbody.innerHTML = '';
      data.mismatches.slice(0, 100).forEach(function(m) {
        var row = document.createElement('tr');
        var timeAgo = getTimeAgo(new Date(m.timestamp));
        var statusClass = m.status === 'override' ? 'status-override' : 'status-pending';
        var label = m.errorTypeLabel || getErrorTypeLabel(m.errorType);
        var actionCell = (m.status === 'pending') ? '<button class="action-btn override-btn" onclick="overrideMismatch(\'' + (m.id || '') + '\')">Override</button>' : '';
        row.innerHTML = '<td>' + formatTimestamp(m.timestamp) + '</td><td>' + (m.expected || '') + '</td><td>' + (m.actual || '—') + '</td><td>' + label + '</td><td><span class="status-badge ' + statusClass + '">' + (m.status || 'pending').toUpperCase() + '</span></td><td>' + (m.id || '') + '</td><td>' + timeAgo + '</td><td>' + actionCell + '</td>';
        tbody.appendChild(row);
      });
    }
    applyViewerAdvancedStats(data.mismatches);
  }
}
function setEl(id, val) {
  var el = document.getElementById(id);
  if (el) el.textContent = val != null ? val : '—';
}

function applyViewerAdvancedStats(data) {
  var set = function(id, text) { var el = document.getElementById(id); if (el) el.textContent = text; };
  if (!data || data.length === 0) {
    set('worst-hour', '--');
    set('worst-hour-count', '0 errors');
    set('best-hour', '--');
    set('best-hour-count', '0 errors');
    set('worst-day', '--');
    set('worst-day-count', '0 errors');
    set('best-day', '--');
    set('best-day-count', '0 errors');
    var emptyMsg = '<p style="color:#888;text-align:center;padding:20px;">No data for selected range</p>';
    var hc = document.getElementById('hourly-chart'); if (hc) hc.innerHTML = emptyMsg;
    var dc = document.getElementById('daily-chart'); if (dc) dc.innerHTML = emptyMsg;
    var tbody = document.getElementById('worst-products-body'); if (tbody) tbody.innerHTML = '';
    return;
  }
  var hourCounts = {}, dayCounts = {}, productCounts = {};
  data.forEach(function(m) {
    var date = new Date(m.timestamp);
    var hour = date.getHours();
    var day = date.toLocaleDateString('en-US', { weekday: 'long' });
    var product = (m.expected || m.actual || '').trim();
    if (product) productCounts[product] = (productCounts[product] || 0) + 1;
    hourCounts[hour] = (hourCounts[hour] || 0) + 1;
    dayCounts[day] = (dayCounts[day] || 0) + 1;
  });
  var hours = Object.keys(hourCounts).map(function(k) { return [parseInt(k, 10), hourCounts[k]]; }).sort(function(a, b) { return b[1] - a[1]; });
  if (hours.length > 0) {
    set('worst-hour', hours[0][0] + ':00');
    set('worst-hour-count', hours[0][1] + ' error' + (hours[0][1] !== 1 ? 's' : ''));
    set('best-hour', hours[hours.length - 1][0] + ':00');
    set('best-hour-count', hours[hours.length - 1][1] + ' error' + (hours[hours.length - 1][1] !== 1 ? 's' : ''));
  }
  var days = Object.keys(dayCounts).map(function(k) { return [k, dayCounts[k]]; }).sort(function(a, b) { return b[1] - a[1]; });
  if (days.length > 0) {
    set('worst-day', days[0][0]);
    set('worst-day-count', days[0][1] + ' error' + (days[0][1] !== 1 ? 's' : ''));
    set('best-day', days[days.length - 1][0]);
    set('best-day-count', days[days.length - 1][1] + ' error' + (days[days.length - 1][1] !== 1 ? 's' : ''));
  }
  var products = Object.keys(productCounts).map(function(k) { return [k, productCounts[k]]; }).sort(function(a, b) { return b[1] - a[1]; }).slice(0, 10);
  var totalProductErrors = products.reduce(function(sum, p) { return sum + p[1]; }, 0);
  var tbody = document.getElementById('worst-products-body');
  if (tbody) {
    tbody.innerHTML = '';
    products.forEach(function(p, index) {
      var pct = totalProductErrors > 0 ? ((p[1] / totalProductErrors) * 100).toFixed(1) : '0';
      var row = document.createElement('tr');
      row.innerHTML = '<td>' + (index + 1) + '</td><td>' + p[0] + '</td><td>' + p[1] + '</td><td>' + pct + '%</td>';
      tbody.appendChild(row);
    });
  }
  var maxH = Math.max.apply(null, Object.keys(hourCounts).map(function(k) { return hourCounts[k]; }).concat([1]));
  var totalH = Object.keys(hourCounts).reduce(function(s, k) { return s + hourCounts[k]; }, 0);
  var hourlyChart = document.getElementById('hourly-chart');
  if (hourlyChart) {
    hourlyChart.innerHTML = '';
    for (var h = 0; h < 24; h++) {
      var count = hourCounts[h] || 0;
      var pct = totalH > 0 ? ((count / totalH) * 100).toFixed(1) : '0.0';
      var w = maxH > 0 ? (count / maxH) * 100 : 0;
      var barClass = count > maxH * 0.7 ? 'high' : (count > maxH * 0.4 ? 'medium' : '');
      var bar = document.createElement('div');
      bar.className = 'bar-item';
      bar.innerHTML = '<div class="bar-label">' + h + ':00</div><div class="bar-wrapper"><div class="bar-fill ' + barClass + '" style="width:' + w + '%">' + (count > 0 ? '<span class="bar-value">' + count + '</span>' : '') + '</div></div><div class="bar-details"><span class="count">' + count + '</span> errors <span class="percentage">(' + pct + '%)</span></div>';
      hourlyChart.appendChild(bar);
    }
  }
  var dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  var maxD = Math.max.apply(null, dayNames.map(function(d) { return dayCounts[d] || 0; }).concat([1]));
  var totalD = dayNames.reduce(function(s, d) { return s + (dayCounts[d] || 0); }, 0);
  var dailyChart = document.getElementById('daily-chart');
  if (dailyChart) {
    dailyChart.innerHTML = '';
    dayNames.forEach(function(day) {
      var count = dayCounts[day] || 0;
      var pct = totalD > 0 ? ((count / totalD) * 100).toFixed(1) : '0.0';
      var w = maxD > 0 ? (count / maxD) * 100 : 0;
      var barClass = count > maxD * 0.7 ? 'high' : (count > maxD * 0.4 ? 'medium' : '');
      var bar = document.createElement('div');
      bar.className = 'bar-item';
      bar.innerHTML = '<div class="bar-label">' + day + '</div><div class="bar-wrapper"><div class="bar-fill ' + barClass + '" style="width:' + w + '%">' + (count > 0 ? '<span class="bar-value">' + count + '</span>' : '') + '</div></div><div class="bar-details"><span class="count">' + count + '</span> errors <span class="percentage">(' + pct + '%)</span></div>';
      dailyChart.appendChild(bar);
    });
  }
}

// Update status table
function updateStatusTable() {
  const tbody = document.getElementById('status-table-body');
  tbody.innerHTML = '';
  const list = mismatches.slice(0, 100);
  list.forEach(mismatch => {
    const row = document.createElement('tr');
    const timeAgo = getTimeAgo(new Date(mismatch.timestamp));
    const statusClass = mismatch.status === 'override' ? 'status-override' : 'status-pending';
    const errorTypeLabel = getErrorTypeLabel(mismatch.errorType);
    row.innerHTML = `
      <td>${formatTimestamp(mismatch.timestamp)}</td>
      <td>${mismatch.expected}</td>
      <td>${mismatch.actual || '—'}</td>
      <td>${errorTypeLabel}</td>
      <td><span class="status-badge ${statusClass}">${(mismatch.status || 'pending').toUpperCase()}</span></td>
      <td>${mismatch.id}</td>
      <td>${timeAgo}</td>
      <td>${mismatch.status === 'pending' ? `<button class="action-btn override-btn" onclick="overrideMismatch('${mismatch.id}')">Override</button>` : ''}</td>
    `;
    tbody.appendChild(row);
  });
}

// Override mismatch
async function overrideMismatch(id) {
    if (!isAuthenticated) {
        showPasswordModal(async () => {
            try {
                await ipcRenderer.invoke('override-mismatch', id);
                await loadData();
                updateUI();
                updateAdvancedStatistics();
            } catch (error) {
                alert('Error overriding mismatch: ' + error.message);
            }
        });
        return;
    }
    
    try {
        await ipcRenderer.invoke('override-mismatch', id);
        await loadData();
        updateUI();
        updateAdvancedStatistics();
    } catch (error) {
        alert('Error overriding mismatch: ' + error.message);
    }
}

// Update statistics
async function updateStatistics() {
  try {
    const stats = await ipcRenderer.invoke('get-statistics', getDateRangeOptions());
    document.getElementById('total-mismatches').textContent = stats.total ?? 0;
    document.getElementById('today-mismatches').textContent = stats.today ?? 0;
    document.getElementById('resolved-count').textContent = stats.resolved ?? 0;
    document.getElementById('pending-count').textContent = stats.pending ?? 0;
  } catch (error) {
    console.error('Error loading statistics:', error);
  }
}

// Update comparison status when only one scanner is connected (no comparison yet)
function updateComparisonStatusForSingleScanner() {
  const statusEl = document.getElementById('comparison-status');
  if (!statusEl) return;
  statusEl.textContent = 'Product scanner active. Connect LPN scanner to compare.';
  statusEl.className = 'comparison-status';
}

function clearPendingMismatchState() {
  var box1 = document.getElementById('scanner1-box');
  var box2 = document.getElementById('scanner2-box');
  var statusEl = document.getElementById('comparison-status');
  if (box1) box1.classList.remove('pending-mismatch');
  if (box2) box2.classList.remove('pending-mismatch');
  if (statusEl) statusEl.classList.remove('pending-mismatch');
}

// Update the displayed codes in the grey pending boxes (timer is not reset).
function updatePendingMismatchDisplay(lpnValue, productValue) {
  var legacySku = lpnValue ? getLegacyFromLpn(lpnValue) : null;
  var lpnProduct = legacySku ? findProductByLegacyItemName(legacySku) : null;
  var lpnResolved = (lpnProduct && (lpnProduct.name || '').trim()) ? (lpnProduct.name || '').trim() : (legacySku || lpnValue || '—');
  var s1Res = document.getElementById('scanner1-product-resolved');
  var s1Code = document.getElementById('scanner1-product-code');
  if (s1Res) s1Res.textContent = lpnResolved;
  if (s1Code) s1Code.textContent = legacySku || lpnValue || '—';
  document.getElementById('scanner2-product-code').textContent = productValue;
  var prod = findProductByBarcode(productValue);
  var resolved = (prod && (prod.name || '').trim()) ? (prod.name || '').trim() : 'Not in product list';
  var s2Res = document.getElementById('scanner2-product-resolved');
  if (s2Res) s2Res.textContent = resolved;
  pushViewerState();
}

// Grey out both boxes during the 5-second wait (no error logged yet).
function showPendingMismatchState() {
  const box1 = document.getElementById('scanner1-box');
  const box2 = document.getElementById('scanner2-box');
  const statusEl = document.getElementById('comparison-status');
  if (box1) {
    box1.classList.remove('match', 'mismatch');
    box1.classList.add('pending-mismatch');
  }
  if (box2) {
    box2.classList.remove('match', 'mismatch');
    box2.classList.add('pending-mismatch');
  }
  document.getElementById('scanner1-status').textContent = 'Checking…';
  document.getElementById('scanner2-status').textContent = 'Checking…';
  if (statusEl) {
    statusEl.textContent = 'Waiting 5s — will show mismatch only if still different.';
    statusEl.className = 'comparison-status pending-mismatch';
  }
  pushViewerState();
}

// Update both scanner boxes. Only match when product list says product and LPN match; otherwise red mismatch.
// result: 'match' | 'lpn_invalid_sku' | 'lpn_wrong_product'
function updateScannerBoxes(lpnValue, productValue, result, errorType) {
  const box1 = document.getElementById('scanner1-box');
  const box2 = document.getElementById('scanner2-box');
  const statusEl = document.getElementById('comparison-status');
  box1.classList.remove('pending-mismatch');
  box2.classList.remove('pending-mismatch');
  var legacySku = lpnValue ? getLegacyFromLpn(lpnValue) : null;
  var lpnProduct = legacySku ? findProductByLegacyItemName(legacySku) : null;
  var lpnResolved = (lpnProduct && (lpnProduct.name || '').trim()) ? (lpnProduct.name || '').trim() : (legacySku || lpnValue || '—');
  var s1Res = document.getElementById('scanner1-product-resolved');
  var s1Code = document.getElementById('scanner1-product-code');
  if (s1Res) s1Res.textContent = lpnResolved;
  if (s1Code) s1Code.textContent = legacySku || lpnValue || '—';
  document.getElementById('scanner2-product-code').textContent = productValue;
  const isMatch = result === 'match';
  const isMismatch = !isMatch;
  const mismatchLabel = isMismatch ? getErrorTypeLabel(errorType || result) : '';
  document.getElementById('scanner1-status').textContent = isMatch ? 'Match' : 'MISMATCH!';
  document.getElementById('scanner2-status').textContent = isMatch ? 'Match' : 'MISMATCH!';
  box1.classList.toggle('match', isMatch);
  box1.classList.toggle('mismatch', isMismatch);
  box1.classList.remove('lpn-missing');
  box2.classList.toggle('match', isMatch);
  box2.classList.toggle('mismatch', isMismatch);
  if (statusEl) {
    statusEl.textContent = isMatch ? 'Match — LPN and Product agree (per product list).' : `${mismatchLabel || 'Mismatch'} — LPN and Product do not match in product list.`;
    statusEl.className = 'comparison-status ' + (isMatch ? 'match' : 'mismatch');
  }
  pushViewerState();
}

// Log error when product and LPN do not match (lpn_invalid_sku | lpn_wrong_product; never 'match')
async function logError(productValue, lpnValue, errorType) {
  try {
    var productDesc = '';
    var product = findProductByBarcode(productValue);
    if (product && (product.name || '').trim()) productDesc = (product.name || '').trim();
    else productDesc = productValue || '—';
    var legacySku = lpnValue ? getLegacyFromLpn(lpnValue) : null;
    var lpnProduct = legacySku ? findProductByLegacyItemName(legacySku) : null;
    var lpnDesc = (lpnProduct && (lpnProduct.name || '').trim()) ? (lpnProduct.name || '').trim() : (legacySku || lpnValue || '—');
    const payload = {
      expected: productValue,
      actual: lpnValue || '',
      expectedDescription: productDesc,
      actualDescription: lpnDesc,
      errorType: errorType || 'lpn_wrong_product',
      timestamp: new Date().toISOString(),
      status: 'pending'
    };
    await ipcRenderer.invoke('log-mismatch', payload);
    await loadData();
    updateUI();
    updateAdvancedStatistics();
  } catch (error) {
    console.error('Error logging:', error);
  }
}

// Helper functions
function formatTimestamp(timestamp) {
    const date = new Date(timestamp);
    return date.toLocaleString();
}

function getTimeAgo(date) {
    const now = new Date();
    const diff = now - date;
    const minutes = Math.floor(diff / 60000);
    const hours = Math.floor(minutes / 60);
    const days = Math.floor(hours / 24);
    
    if (days > 0) return `${days} day${days > 1 ? 's' : ''} ago`;
    if (hours > 0) return `${hours} hr${hours > 1 ? 's' : ''} ago`;
    if (minutes > 0) return `${minutes} min${minutes > 1 ? 's' : ''} ago`;
    return 'Just now';
}

// Tab switching
function switchTab(tabName) {
    // Hide all tabs
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });
    
    // Remove active class from all buttons
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // Show selected tab
    if (tabName === 'main') {
        document.getElementById('main-tab').classList.add('active');
        document.getElementById('main-tab-btn').classList.add('active');
    } else if (tabName === 'stats') {
        document.getElementById('stats-tab').classList.add('active');
        document.getElementById('stats-tab-btn').classList.add('active');
        updateAdvancedStatistics();
    }
}

// Update advanced statistics (optionally filtered by stats date range)
async function updateAdvancedStatistics() {
  let data = mismatches;
  if (statsDateStart || statsDateEnd) {
    try {
      data = await ipcRenderer.invoke('get-mismatches', {
        limit: 5000,
        startDate: statsDateStart,
        endDate: statsDateEnd
      });
    } catch (e) {
      console.error(e);
    }
  }
  if (data.length === 0) {
        // Set defaults
        document.getElementById('worst-hour').textContent = '--';
        document.getElementById('worst-hour-count').textContent = '0 errors';
        document.getElementById('best-hour').textContent = '--';
        document.getElementById('best-hour-count').textContent = '0 errors';
        document.getElementById('worst-day').textContent = '--';
        document.getElementById('worst-day-count').textContent = '0 errors';
        document.getElementById('best-day').textContent = '--';
        document.getElementById('best-day-count').textContent = '0 errors';
    renderEmptyCharts();
    return;
  }

  const hourCounts = {};
  const dayCounts = {};
  const productCounts = {};
  data.forEach(mismatch => {
        const date = new Date(mismatch.timestamp);
        const hour = date.getHours();
        const day = date.toLocaleDateString('en-US', { weekday: 'long' });
        const product = (mismatch.expected || mismatch.actual || '').trim();
        if (product) {
          productCounts[product] = (productCounts[product] || 0) + 1;
        }
        hourCounts[hour] = (hourCounts[hour] || 0) + 1;
        dayCounts[day] = (dayCounts[day] || 0) + 1;
    });
    
    // Find worst/best hours
    const hours = Object.entries(hourCounts).sort((a, b) => b[1] - a[1]);
    if (hours.length > 0) {
        const worstHour = hours[0];
        const bestHour = hours[hours.length - 1];
        document.getElementById('worst-hour').textContent = `${worstHour[0]}:00`;
        document.getElementById('worst-hour-count').textContent = `${worstHour[1]} error${worstHour[1] !== 1 ? 's' : ''}`;
        document.getElementById('best-hour').textContent = `${bestHour[0]}:00`;
        document.getElementById('best-hour-count').textContent = `${bestHour[1]} error${bestHour[1] !== 1 ? 's' : ''}`;
    }
    
    // Find worst/best days
    const days = Object.entries(dayCounts).sort((a, b) => b[1] - a[1]);
    if (days.length > 0) {
        const worstDay = days[0];
        const bestDay = days[days.length - 1];
        document.getElementById('worst-day').textContent = worstDay[0];
        document.getElementById('worst-day-count').textContent = `${worstDay[1]} error${worstDay[1] !== 1 ? 's' : ''}`;
        document.getElementById('best-day').textContent = bestDay[0];
        document.getElementById('best-day-count').textContent = `${bestDay[1]} error${bestDay[1] !== 1 ? 's' : ''}`;
    }
    
    // Worst products
    const products = Object.entries(productCounts).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const totalProductErrors = products.reduce((sum, p) => sum + p[1], 0);
    const tbody = document.getElementById('worst-products-body');
    tbody.innerHTML = '';
    products.forEach((product, index) => {
        const percentage = ((product[1] / totalProductErrors) * 100).toFixed(1);
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${index + 1}</td>
            <td>${product[0]}</td>
            <td>${product[1]}</td>
            <td>${percentage}%</td>
        `;
        tbody.appendChild(row);
    });
    
    // Hourly distribution - Bar Chart
    const hourlyChart = document.getElementById('hourly-chart');
    hourlyChart.innerHTML = '';
    const allHours = Array.from({length: 24}, (_, i) => i);
    const totalHourlyErrors = Object.values(hourCounts).reduce((a, b) => a + b, 0);
    const maxHourlyCount = Math.max(...Object.values(hourCounts), 1);
    
    allHours.forEach(hour => {
        const count = hourCounts[hour] || 0;
        const percentage = totalHourlyErrors > 0 ? ((count / totalHourlyErrors) * 100).toFixed(1) : '0.0';
        const widthPercent = maxHourlyCount > 0 ? (count / maxHourlyCount) * 100 : 0;
        
        const barItem = document.createElement('div');
        barItem.className = 'bar-item';
        
        // Determine bar color based on count
        let barClass = '';
        if (count > maxHourlyCount * 0.7) barClass = 'high';
        else if (count > maxHourlyCount * 0.4) barClass = 'medium';
        
        barItem.innerHTML = `
            <div class="bar-label">${hour}:00</div>
            <div class="bar-wrapper">
                <div class="bar-fill ${barClass}" style="width: ${widthPercent}%">
                    ${count > 0 ? `<span class="bar-value">${count}</span>` : ''}
                </div>
            </div>
            <div class="bar-details">
                <span class="count">${count}</span> errors
                <span class="percentage">(${percentage}%)</span>
            </div>
        `;
        hourlyChart.appendChild(barItem);
    });
    
    // Daily distribution - Bar Chart
    const dailyChart = document.getElementById('daily-chart');
    dailyChart.innerHTML = '';
    const dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
    const totalDailyErrors = Object.values(dayCounts).reduce((a, b) => a + b, 0);
    const maxDailyCount = Math.max(...Object.values(dayCounts), 1);
    
    dayNames.forEach(day => {
        const count = dayCounts[day] || 0;
        const percentage = totalDailyErrors > 0 ? ((count / totalDailyErrors) * 100).toFixed(1) : '0.0';
        const widthPercent = maxDailyCount > 0 ? (count / maxDailyCount) * 100 : 0;
        
        const barItem = document.createElement('div');
        barItem.className = 'bar-item';
        
        // Determine bar color based on count
        let barClass = '';
        if (count > maxDailyCount * 0.7) barClass = 'high';
        else if (count > maxDailyCount * 0.4) barClass = 'medium';
        
        barItem.innerHTML = `
            <div class="bar-label">${day}</div>
            <div class="bar-wrapper">
                <div class="bar-fill ${barClass}" style="width: ${widthPercent}%">
                    ${count > 0 ? `<span class="bar-value">${count}</span>` : ''}
                </div>
            </div>
            <div class="bar-details">
                <span class="count">${count}</span> errors
                <span class="percentage">(${percentage}%)</span>
            </div>
        `;
        dailyChart.appendChild(barItem);
    });
}

function renderEmptyCharts() {
  ['hourly-chart', 'daily-chart'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = '<p style="color:#888;text-align:center;padding:20px;">No data for selected range</p>';
  });
  const tbody = document.getElementById('worst-products-body');
  if (tbody) tbody.innerHTML = '';
}

// Make functions available globally
window.deleteProduct = deleteProduct;
window.overrideMismatch = overrideMismatch;
window.switchTab = switchTab;
