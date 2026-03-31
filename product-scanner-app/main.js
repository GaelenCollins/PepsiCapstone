const { app, BrowserWindow, ipcMain } = require('electron');
// Pi / Electron: avoids GpuControl / CreateCommandBuffer crashes on some ARM setups
if (process.platform === 'linux' && process.env.ELECTRON_DISABLE_GPU !== '0') {
  app.disableHardwareAcceleration();
}
const path = require('path');
const fs = require('fs');
const http = require('http');
const os = require('os');
const {
  triggerAlarm,
  clearAlarm,
  startAlarmResetButtonWatcher,
  initStackLightGpioAtStartup,
  setMismatchesChangedNotify
} = require('./services/alarm-service');
const { initDatabase, logMismatch, getMismatches, getStatistics, overrideMismatch } = require('./services/database');
const { getErrorTypeLabel } = require('./services/error-types');
const { getEmailConfig, setEmailConfig, sendErrorNotificationToOscar } = require('./services/email-service');
const { lookupBarcode } = require('./services/barcode-lookup-service');
const { logScan } = require('./services/scan-log');

let mainWindow;
let isAuthenticated = false;
let viewerState = {};
let viewerAuthenticated = false;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true
    },
    icon: path.join(__dirname, 'assets', 'icon.png')
  });

  mainWindow.loadFile('index.html');
}

const VIEWER_PORT = parseInt(process.env.VIEWER_PORT || '3847', 10);

function startViewerServer() {
  const appDir = __dirname;
  const mime = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'application/javascript',
    '.css': 'text/css',
    '.json': 'application/json',
    '.png': 'image/png',
    '.ico': 'image/x-icon'
  };
  const server = http.createServer(async (req, res) => {
    const url = new URL(req.url || '', 'http://localhost');
    if (req.method === 'OPTIONS') {
      res.writeHead(204, {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type'
      });
      res.end();
      return;
    }
    if (url.pathname === '/api/state') {
      try {
        const [rows, stats] = await Promise.all([
          getMismatches(500),
          getStatistics(null, null)
        ]);
        const mismatches = rows.map(r => ({ ...r, errorTypeLabel: getErrorTypeLabel(r.errorType) }));
        const lastErrorAt = mismatches.length > 0 ? mismatches[0].timestamp : null;
        const lastErrorAgo = lastErrorAt ? timeAgo(new Date(lastErrorAt)) : null;
        const body = JSON.stringify({
          ...viewerState,
          mismatches,
          stats: { total: stats.total, today: stats.today, resolved: stats.resolved, pending: stats.pending },
          lastErrorAt,
          lastErrorAgo
        });
        res.setHeader('Content-Type', 'application/json');
        res.setHeader('Access-Control-Allow-Origin', '*');
        res.end(body);
      } catch (e) {
        res.statusCode = 500;
        res.setHeader('Content-Type', 'application/json');
        res.end(JSON.stringify({ error: String(e.message) }));
      }
      return;
    }
    const readBody = () => new Promise((resolve, reject) => {
      let chunks = [];
      req.on('data', c => chunks.push(c));
      req.on('end', () => {
        try {
          const raw = Buffer.concat(chunks).toString('utf8');
          resolve(raw ? JSON.parse(raw) : {});
        } catch (e) {
          resolve({});
        }
      });
      req.on('error', reject);
    });
    const json = (obj) => {
      res.setHeader('Content-Type', 'application/json');
      res.setHeader('Access-Control-Allow-Origin', '*');
      res.end(JSON.stringify(obj));
    };
    const err = (code, msg) => {
      res.statusCode = code;
      res.setHeader('Content-Type', 'application/json');
      res.setHeader('Access-Control-Allow-Origin', '*');
      res.end(JSON.stringify({ error: msg || 'Error' }));
    };
    if (url.pathname === '/api/authenticate' && req.method === 'POST') {
      try {
        const body = await readBody();
        viewerAuthenticated = (body.password === 'admin123');
        json({ ok: viewerAuthenticated });
      } catch (e) {
        err(500, e.message);
      }
      return;
    }
    if (url.pathname === '/api/check-auth' && req.method === 'GET') {
      json({ authenticated: viewerAuthenticated });
      return;
    }
    if (url.pathname === '/api/product-master' && req.method === 'GET') {
      try {
        const data = fs.readFileSync(path.join(__dirname, 'data', 'product-master.json'), 'utf8');
        json(JSON.parse(data));
      } catch (e) {
        json([]);
      }
      return;
    }
    if (url.pathname === '/api/product-master' && req.method === 'POST') {
      try {
        const body = await readBody();
        if (body.password === 'admin123') viewerAuthenticated = true;
        if (!viewerAuthenticated) {
          err(401, 'Not authenticated');
          return;
        }
        const products = body.products != null ? body.products : body;
        if (Array.isArray(products)) {
          const dir = path.join(__dirname, 'data');
          if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
          fs.writeFileSync(path.join(dir, 'product-master.json'), JSON.stringify(products, null, 2));
        }
        json({ ok: true });
      } catch (e) {
        err(500, e.message);
      }
      return;
    }
    if (url.pathname === '/api/override' && req.method === 'POST') {
      try {
        const body = await readBody();
        if (body.password === 'admin123') viewerAuthenticated = true;
        if (!viewerAuthenticated) {
          err(401, 'Not authenticated');
          return;
        }
        await overrideMismatch(body.id);
        await clearAlarm();
        json({ ok: true });
      } catch (e) {
        err(500, e.message);
      }
      return;
    }
    if (url.pathname === '/api/email-config' && req.method === 'GET') {
      json(getEmailConfig());
      return;
    }
    if (url.pathname === '/api/email-config' && req.method === 'POST') {
      try {
        const body = await readBody();
        setEmailConfig(body);
        json({ ok: true });
      } catch (e) {
        err(500, e.message);
      }
      return;
    }
    if (url.pathname === '/api/export' && req.method === 'GET') {
      try {
        const start = url.searchParams.get('startDate') || null;
        const end = url.searchParams.get('endDate') || null;
        const rows = await getMismatches(10000, start, end);
        json({ rows, format: 'csv' });
      } catch (e) {
        err(500, e.message);
      }
      return;
    }
    if (url.pathname === '/api/mismatches' && req.method === 'GET') {
      try {
        const limit = parseInt(url.searchParams.get('limit') || '500', 10);
        const start = url.searchParams.get('startDate') || null;
        const end = url.searchParams.get('endDate') || null;
        const rows = await getMismatches(limit, start, end);
        json(rows.map(r => ({ ...r, errorTypeLabel: getErrorTypeLabel(r.errorType) })));
      } catch (e) {
        err(500, e.message);
      }
      return;
    }
    if (url.pathname === '/api/statistics' && req.method === 'GET') {
      try {
        const start = url.searchParams.get('startDate') || null;
        const end = url.searchParams.get('endDate') || null;
        const stats = await getStatistics(start, end);
        json(stats);
      } catch (e) {
        err(500, e.message);
      }
      return;
    }
    if (url.pathname === '/api/host-info' && req.method === 'GET') {
      const port = VIEWER_PORT;
      const hostUrls = [];
      try {
        const ifaces = os.networkInterfaces();
        for (const name of Object.keys(ifaces)) {
          for (const iface of ifaces[name]) {
            if (iface.family === 'IPv4' && !iface.internal) {
              hostUrls.push('http://' + iface.address + ':' + port);
            }
          }
        }
      } catch (e) {}
      json({ port, hostUrl: hostUrls[0] || 'http://localhost:' + port, hostUrls });
      return;
    }
    let filePath = url.pathname === '/' ? '/index.html' : url.pathname;
    if (filePath.startsWith('/assets/')) {
      filePath = path.join(appDir, filePath.slice(1));
    } else if (filePath === '/index.html' || filePath === '/renderer.js' || filePath === '/styles.css') {
      filePath = path.join(appDir, filePath.slice(1));
    } else {
      res.statusCode = 404;
      res.end('Not found');
      return;
    }
    const resolved = path.resolve(filePath);
    if (!resolved.startsWith(appDir)) {
      res.statusCode = 404;
      res.end('Not found');
      return;
    }
    try {
      const data = fs.readFileSync(resolved);
      const ext = path.extname(resolved);
      res.setHeader('Content-Type', mime[ext] || 'application/octet-stream');
      res.end(data);
    } catch (err) {
      if (err.code === 'ENOENT') {
        res.statusCode = 404;
        res.end('Not found');
      } else {
        res.statusCode = 500;
        res.end('Error');
      }
    }
  });
  server.listen(VIEWER_PORT, '0.0.0.0', () => {
    const urls = ['http://127.0.0.1:' + VIEWER_PORT];
    try {
      const ifaces = os.networkInterfaces();
      for (const name of Object.keys(ifaces)) {
        for (const iface of ifaces[name]) {
          if (iface.family === 'IPv4' && !iface.internal) {
            urls.push('http://' + iface.address + ':' + VIEWER_PORT);
          }
        }
      }
    } catch (e) {}
    console.log('Viewer listening on port ' + VIEWER_PORT + '. Try from phone (same Wi-Fi, use http not https):');
    urls.slice(1).forEach(u => console.log('  ' + u));
    console.log('If phone cannot connect: (1) Confirm Mac IP in System Settings > Network. (2) Turn off "Guest network" or "AP isolation" on router. (3) On Mac run: curl -s -o /dev/null -w "%{http_code}" ' + (urls[1] || urls[0]) + '/api/state  (expect 200)');
  });
  server.on('error', (err) => {
    if (err.code === 'EADDRINUSE') console.warn(`Viewer port ${VIEWER_PORT} in use; skip viewer.`);
    else console.warn('Viewer server error:', err.message);
  });
}

function timeAgo(date) {
  const sec = Math.floor((Date.now() - date.getTime()) / 1000);
  if (sec >= 86400) return `${Math.floor(sec / 86400)}d ago`;
  if (sec >= 3600) return `${Math.floor(sec / 3600)}h ago`;
  if (sec >= 60) return `${Math.floor(sec / 60)}m ago`;
  return sec > 0 ? `${sec}s ago` : 'just now';
}

app.whenReady().then(() => {
  initDatabase();
  createWindow();
  setMismatchesChangedNotify(() => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('mismatches-changed');
    }
  });
  startViewerServer();
  // Stack light first so alarm-service picks pigpio vs onoff before reset watcher runs
  try {
    initStackLightGpioAtStartup();
  } catch (e) {
    console.error('[Alarm] Stack light startup probe failed:', e && e.message);
  }
  try {
    startAlarmResetButtonWatcher();
  } catch (e) {
    console.error('[Alarm] Reset button setup failed:', e && e.message);
  }

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

ipcMain.handle('authenticate', async (event, password) => {
  const correctPassword = 'admin123';
  isAuthenticated = (password === correctPassword);
  return isAuthenticated;
});

ipcMain.handle('viewer-state', (event, state) => {
  viewerState = state || {};
  return true;
});

ipcMain.handle('check-auth', () => {
  return isAuthenticated;
});

ipcMain.handle('logout', () => {
  isAuthenticated = false;
  return true;
});

ipcMain.handle('get-product-master', async () => {
  try {
    const data = fs.readFileSync(path.join(__dirname, 'data', 'product-master.json'), 'utf8');
    return JSON.parse(data);
  } catch (error) {
    return [];
  }
});

ipcMain.handle('save-product-master', async (event, products) => {
  if (!isAuthenticated) {
    throw new Error('Not authenticated');
  }
  try {
    const dir = path.join(__dirname, 'data');
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    fs.writeFileSync(
      path.join(dir, 'product-master.json'),
      JSON.stringify(products, null, 2)
    );
    return true;
  } catch (error) {
    throw error;
  }
});

ipcMain.handle('log-mismatch', async (event, mismatchData) => {
  try {
    const id = await logMismatch(mismatchData);
    triggerAlarm();
    // Email Oscar only when there is an error (if enabled in Settings)
    sendErrorNotificationToOscar(mismatchData);
    return id;
  } catch (error) {
    console.error('Error logging mismatch:', error);
    throw error;
  }
});

ipcMain.handle('get-mismatches', async (event, options = {}) => {
  const { limit = 500, startDate, endDate } = options;
  return await getMismatches(limit, startDate, endDate);
});

ipcMain.handle('get-statistics', async (event, options = {}) => {
  const { startDate, endDate } = options;
  return await getStatistics(startDate, endDate);
});

ipcMain.handle('override-mismatch', async (event, mismatchId) => {
  if (!isAuthenticated) {
    throw new Error('Not authenticated');
  }
  try {
    await overrideMismatch(mismatchId);
    await clearAlarm();
    return true;
  } catch (error) {
    throw error;
  }
});

ipcMain.handle('clear-alarm', async () => {
  await clearAlarm();
  return true;
});

ipcMain.handle('export-data', async (event, options = {}) => {
  const { startDate, endDate, format = 'csv' } = options;
  const rows = await getMismatches(10000, startDate, endDate);
  return { rows, format };
});

// Email-on-error settings (Oscar)
ipcMain.handle('get-email-config', () => getEmailConfig());
ipcMain.handle('save-email-config', (event, config) => setEmailConfig(config));

// Barcode lookup for product identification (UPCitemDB)
ipcMain.handle('lookup-barcode', async (event, code) => {
  if (!code || typeof code !== 'string') return null;
  return await lookupBarcode(code.trim());
});

// Log every scan to data/scan-log.jsonl
ipcMain.handle('log-scan', (event, { code, product }) => {
  logScan(code || '', product || '');
});
