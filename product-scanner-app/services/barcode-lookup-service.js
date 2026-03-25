/**
 * Barcode lookup for product identification.
 * Uses our local cache first (data/barcode-cache.json). Only calls UPCitemDB API on cache miss,
 * then stores the result so we don't burn free-tier calls. For testing, seed the cache and it
 * becomes the main source.
 *
 * UPCitemDB free: 100 req/day, 6/min. https://www.upcitemdb.com – 681M+ UPC/EAN/ISBN.
 * Free tier: /prod/trial/lookup (omit user_key/key_type).
 *
 * Uses Node https (not fetch) so it works in Electron's main process (Node 16 has no fetch).
 *
 * LPN barcodes: format TBD. Once LPN barcode structure is known we can:
 * - Parse LPN fields here (e.g. product code, location, batch)
 * - Link LPN to product lookup and show in UI
 */

const path = require('path');
const fs = require('fs');
const https = require('https');

const UPCITEMDB_HOST = 'api.upcitemdb.com';
const UPCITEMDB_PATH_PREFIX = '/prod/trial/lookup';
const OPENFOODFACTS_HOST = 'world.openfoodfacts.org';
const CACHE_PATH = path.join(__dirname, '..', 'data', 'barcode-cache.json');

function ensureDataDir() {
  const dir = path.dirname(CACHE_PATH);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function loadCache() {
  try {
    const raw = fs.readFileSync(CACHE_PATH, 'utf8');
    return JSON.parse(raw);
  } catch (e) {
    return {};
  }
}

function saveCache(cache) {
  ensureDataDir();
  fs.writeFileSync(CACHE_PATH, JSON.stringify(cache, null, 2), 'utf8');
}

function normalizeCode(code) {
  if (!code || typeof code !== 'string') return '';
  return String(code).replace(/\D/g, '');
}

/** Variants to try: with leading zero for UPC-A (11→12) and EAN-13 (12→13). Open Food Facts accepts many lengths. */
function barcodeVariants(upc) {
  if (!upc || typeof upc !== 'string') return [];
  const len = upc.length;
  const out = [upc];
  if (len === 11) out.unshift('0' + upc);
  if (len === 12) out.push('0' + upc);
  if (len === 14 && upc.startsWith('0')) out.push(upc.slice(1));
  return [...new Set(out)];
}

function isValidUpcFormat(upc) {
  if (!upc || typeof upc !== 'string') return false;
  const len = upc.length;
  return len === 8 || len === 12 || len === 13 || len === 11;
}

/** Take text and return only the part before INGREDIENTS (we don't store ingredients). */
function everythingUntilIngredients(text) {
  if (!text || typeof text !== 'string') return '';
  const i = text.search(/\bINGREDIENTS\s*:/i);
  return i === -1 ? text.trim() : text.slice(0, i).trim();
}

/** Build single product-info string from API item: everything until INGREDIENTS. */
function buildProductInfo(item) {
  const parts = [];
  if (item.title) parts.push(everythingUntilIngredients(item.title));
  if (item.brand && !parts.includes(item.brand)) parts.push(item.brand);
  if (item.category) parts.push(item.category);
  if (item.description) {
    const desc = everythingUntilIngredients(item.description);
    if (desc && !parts.includes(desc)) parts.push(desc);
  }
  return parts.filter(Boolean).join(' · ');
}

/** Call UPCitemDB trial API with Node https (works in Electron main process). */
function fetchUpcItemDb(upc) {
  return new Promise((resolve, reject) => {
    const pathQuery = `${UPCITEMDB_PATH_PREFIX}?upc=${encodeURIComponent(upc)}`;
    const opts = {
      hostname: UPCITEMDB_HOST,
      path: pathQuery,
      method: 'GET',
      headers: { 'Accept': 'application/json' }
    };
    const req = https.request(opts, (res) => {
      let body = '';
      res.on('data', (chunk) => { body += chunk; });
      res.on('end', () => {
        if (res.statusCode !== 200) {
          let apiCode = '';
          try {
            const j = JSON.parse(body);
            if (j && j.code) apiCode = j.code;
          } catch (_) {}
          const err = new Error(`UPCitemDB ${res.statusCode}: ${body.slice(0, 200)}`);
          err.statusCode = res.statusCode;
          err.apiCode = apiCode;
          reject(err);
          return;
        }
        try {
          resolve(JSON.parse(body));
        } catch (e) {
          reject(e);
        }
      });
    });
    req.on('error', reject);
    req.setTimeout(15000, () => { req.destroy(); reject(new Error('timeout')); });
    req.end();
  });
}

/** Open Food Facts: free, no key, 100/min. Good for food/grocery. */
function fetchOpenFoodFacts(upc) {
  return new Promise((resolve) => {
    const pathQuery = `/api/v2/product/${encodeURIComponent(upc)}.json?fields=product_name,brands,product_name_en`;
    const opts = {
      hostname: OPENFOODFACTS_HOST,
      path: pathQuery,
      method: 'GET',
      headers: { 'Accept': 'application/json' }
    };
    const req = https.request(opts, (res) => {
      let body = '';
      res.on('data', (chunk) => { body += chunk; });
      res.on('end', () => {
        if (res.statusCode !== 200) {
          resolve(null);
          return;
        }
        try {
          const j = JSON.parse(body);
          const p = j && j.product;
          if (!p || j.status !== 1) {
            resolve(null);
            return;
          }
          const name = (p.product_name_en || p.product_name || '').trim();
          const brands = (p.brands || '').trim();
          const productInfo = [name, brands].filter(Boolean).join(' · ') || null;
          resolve(productInfo ? { productInfo } : null);
        } catch (e) {
          resolve(null);
        }
      });
    });
    req.on('error', () => resolve(null));
    req.setTimeout(10000, () => { req.destroy(); resolve(null); });
    req.end();
  });
}

/**
 * Look up a barcode: check local cache first, then API only on miss. Saves API result to cache.
 * @param {string} code - Barcode string (digits or with leading zero)
 * @returns {Promise<{ title: string, brand: string, category?: string, raw?: object } | null>}
 */
async function lookupBarcode(code) {
  const upc = normalizeCode(code);
  if (!upc || upc.length < 8) return null;

  const cache = loadCache();
  const cached = cache[upc];
  if (cached) {
    if (cached.invalid && cached.apiError === 'TOO_MANY_REQUESTS') return null;
    const productInfo = cached.productInfo || [cached.title, cached.brand, cached.category, cached.description].filter(Boolean).join(' · ');
    const noProduct = !productInfo || /^no product found/i.test(productInfo.trim());
    if (!noProduct) {
      return { productInfo, code: upc };
    }
  }

  const variants = barcodeVariants(upc);
  if (variants.length === 0) return null;

  // Try Open Food Facts first (no key, lenient on format, 100/min)
  for (const variant of variants) {
    const off = await fetchOpenFoodFacts(variant);
    if (off && off.productInfo) {
      cache[upc] = {
        productInfo: off.productInfo,
        title: '',
        brand: '',
        category: '',
        description: '',
        cachedAt: new Date().toISOString(),
        source: 'openfoodfacts'
      };
      saveCache(cache);
      return { productInfo: off.productInfo, code: upc };
    }
  }

  // Then try UPCitemDB (strict format: 8, 12, or 13 digits only)
  const tryUpcItemDb = async (variant) => {
    if (variant.length !== 8 && variant.length !== 12 && variant.length !== 13) return null;
    try {
      const data = await fetchUpcItemDb(variant);
      const items = data.items;
      if (!Array.isArray(items) || items.length === 0) return null;
      const item = items[0];
      const productInfo = buildProductInfo(item);
      return { productInfo, title: item.title || '', brand: item.brand || '', category: item.category || '', description: everythingUntilIngredients(item.description || '') };
    } catch (e) {
      if (e.statusCode === 429) throw e;
      return null;
    }
  };

  for (const variant of variants) {
    try {
      const result = await tryUpcItemDb(variant);
      if (result) {
        cache[upc] = {
          productInfo: result.productInfo,
          title: result.title || '',
          brand: result.brand || '',
          category: result.category || '',
          description: result.description || '',
          cachedAt: new Date().toISOString(),
          source: 'upcitemdb'
        };
        saveCache(cache);
        return { productInfo: result.productInfo, code: upc };
      }
    } catch (e) {
      if (e.statusCode === 429) {
        cache[upc] = { invalid: true, cachedAt: new Date().toISOString(), apiError: 'TOO_MANY_REQUESTS' };
        saveCache(cache);
        console.warn('[Barcode lookup]', e.message);
        return null;
      }
    }
  }

  // Don't cache NOT_FOUND so next scan can retry (e.g. OFF might get the product later)
  return null;
}

module.exports = {
  lookupBarcode
};
