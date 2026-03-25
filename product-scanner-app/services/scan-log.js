/**
 * Log every scanned barcode to data/scan-log.jsonl (one JSON object per line).
 * Each line: { "timestamp": "ISO date", "code": "096619295210", "product": "Safe Can Kirkland Sparkling Water 12Oz" }
 */

const path = require('path');
const fs = require('fs');

const SCAN_LOG_PATH = path.join(__dirname, '..', 'data', 'scan-log.jsonl');

function ensureDataDir() {
  const dir = path.dirname(SCAN_LOG_PATH);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

/**
 * Append one scan to the log. Call after each successful scan.
 * @param {string} code - Barcode (e.g. 096619295210)
 * @param {string} [product] - Product display name (optional)
 */
function logScan(code, product = '') {
  if (!code || typeof code !== 'string') return;
  ensureDataDir();
  const line = JSON.stringify({
    timestamp: new Date().toISOString(),
    code: code.trim(),
    product: (product || '').trim()
  }) + '\n';
  fs.appendFileSync(SCAN_LOG_PATH, line, 'utf8');
}

module.exports = {
  logScan,
  SCAN_LOG_PATH
};
