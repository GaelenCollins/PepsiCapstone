/**
 * Mismatch log — JSON file store (no native sqlite3; fast npm install on Pi / no electron-rebuild for DB).
 * File: data/mismatches-store.json (gitignored with data/)
 * If you have an old mismatches.db, it is not auto-imported; copy data manually if needed.
 */

const path = require('path');
const fs = require('fs');

const storePath = path.join(__dirname, '..', 'data', 'mismatches-store.json');

let mismatches = [];
let nextId = 1;

function ensureDataDir() {
  const dir = path.dirname(storePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

function loadStore() {
  try {
    const raw = fs.readFileSync(storePath, 'utf8');
    const data = JSON.parse(raw);
    if (Array.isArray(data)) {
      mismatches = data;
    } else if (data && Array.isArray(data.mismatches)) {
      mismatches = data.mismatches;
    } else {
      mismatches = [];
    }
  } catch (e) {
    mismatches = [];
  }
  nextId = mismatches.reduce((max, r) => {
    const n = parseInt(String(r.id), 10);
    return Number.isFinite(n) ? Math.max(max, n) : max;
  }, 0) + 1;
}

function saveStore() {
  ensureDataDir();
  fs.writeFileSync(storePath, JSON.stringify(mismatches, null, 2), 'utf8');
}

function datePart(iso) {
  if (!iso) return '';
  const s = String(iso);
  return s.length >= 10 ? s.slice(0, 10) : s;
}

function localTodayYmd() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function rowToApi(row) {
  return {
    id: String(row.id),
    expected: row.expected,
    actual: row.actual,
    errorType: row.errorType || row.error_type || 'mismatch',
    timestamp: row.timestamp,
    status: row.status,
    resolvedAt: row.resolvedAt || row.resolved_at,
    resolvedBy: row.resolvedBy || row.resolved_by
  };
}

function initDatabase() {
  ensureDataDir();
  loadStore();
  console.log('Mismatch store ready (JSON):', mismatches.length, 'records');
}

function logMismatch(mismatchData) {
  return new Promise((resolve, reject) => {
    try {
      const { expected, actual, errorType = 'mismatch', timestamp, status = 'pending' } = mismatchData;
      const row = {
        id: nextId++,
        expected: expected != null ? String(expected) : '',
        actual: actual != null ? String(actual) : '',
        errorType,
        timestamp: timestamp || new Date().toISOString(),
        status,
        resolvedAt: null,
        resolvedBy: null
      };
      mismatches.push(row);
      saveStore();
      resolve(row.id);
    } catch (e) {
      reject(e);
    }
  });
}

function getMismatches(limit = 500, startDate = null, endDate = null) {
  return new Promise((resolve) => {
    let list = [...mismatches];
    if (startDate) {
      list = list.filter((r) => datePart(r.timestamp) >= startDate);
    }
    if (endDate) {
      list = list.filter((r) => datePart(r.timestamp) <= endDate);
    }
    list.sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)));
    if (limit && list.length > limit) {
      list = list.slice(0, limit);
    }
    resolve(list.map(rowToApi));
  });
}

function getStatistics(startDate = null, endDate = null) {
  return new Promise((resolve) => {
    let list = [...mismatches];
    if (startDate) {
      list = list.filter((r) => datePart(r.timestamp) >= startDate);
    }
    if (endDate) {
      list = list.filter((r) => datePart(r.timestamp) <= endDate);
    }
    const today = localTodayYmd();
    let total = list.length;
    let todayCount = 0;
    let resolved = 0;
    let pending = 0;
    for (const r of list) {
      if (datePart(r.timestamp) === today) todayCount++;
      const st = r.status || 'pending';
      if (st === 'pending') pending++;
      if (st === 'sent' || st === 'override') resolved++;
    }
    resolve({ total, today: todayCount, resolved, pending });
  });
}

function overrideMismatch(id) {
  return new Promise((resolve, reject) => {
    try {
      const sid = String(id);
      const row = mismatches.find((r) => String(r.id) === sid);
      if (!row) {
        resolve(0);
        return;
      }
      row.status = 'override';
      row.resolvedAt = new Date().toISOString();
      row.resolvedBy = 'operator';
      saveStore();
      resolve(1);
    } catch (e) {
      reject(e);
    }
  });
}

/** Most recent pending row by timestamp (same status as UI Override). */
function overrideLatestPendingMismatch(resolvedBy = 'gpio_reset') {
  return new Promise((resolve, reject) => {
    try {
      let latest = null;
      for (const r of mismatches) {
        if ((r.status || 'pending') !== 'pending') continue;
        const ts = String(r.timestamp || '');
        if (!latest || ts.localeCompare(String(latest.timestamp || '')) > 0) {
          latest = r;
        }
      }
      if (!latest) {
        resolve(0);
        return;
      }
      latest.status = 'override';
      latest.resolvedAt = new Date().toISOString();
      latest.resolvedBy = resolvedBy;
      saveStore();
      resolve(1);
    } catch (e) {
      reject(e);
    }
  });
}

module.exports = {
  initDatabase,
  logMismatch,
  getMismatches,
  getStatistics,
  overrideMismatch,
  overrideLatestPendingMismatch
};
