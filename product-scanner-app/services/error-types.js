/**
 * Scan outcomes: only match when product list says product and LPN match.
 * - match: LPN and product agree per product list (same legacy item name).
 * - lpn_invalid_sku: Mismatch — LPN doesn't contain a valid legacy SKU.
 * - lpn_wrong_product: Mismatch — product and LPN do not match in product list (or no product scanned).
 * Legacy 'lpn_missing' in DB is shown as 'Mismatch'. No LPN-missing logic; non-match = red mismatch.
 */

const LABELS = {
  match: 'Match',
  lpn_invalid_sku: 'LPN invalid SKU',
  lpn_wrong_product: 'LPN wrong product',
  mismatch: 'Mismatch',
  lpn_missing: 'Mismatch' // legacy DB; display as Mismatch
};

function getErrorTypeLabel(errorType) {
  if (!errorType) return 'Mismatch';
  return LABELS[errorType] || errorType;
}

function isError(errorType) {
  if (!errorType) return true;
  return errorType !== 'match';
}

module.exports = {
  getErrorTypeLabel,
  isError,
  LABELS
};
