/**
 * Email-on-error service: notify Oscar when an error occurs.
 * Supports Mailgun API (if API key + domain set) or SMTP (nodemailer).
 */

const path = require('path');
const fs = require('fs');
const https = require('https');
const nodemailer = require('nodemailer');

const CONFIG_PATH = path.join(__dirname, '..', 'data', 'email-config.json');

const DEFAULT_CONFIG = {
  emailOscarOnError: false,
  oscarEmail: '',
  useMailgun: true,
  mailgunApiKey: '',
  mailgunDomain: '',
  smtpHost: '',
  smtpPort: 587,
  smtpSecure: false,
  smtpUser: '',
  smtpPass: '',
  fromEmail: ''
};

function loadConfig() {
  try {
    const data = fs.readFileSync(CONFIG_PATH, 'utf8');
    return { ...DEFAULT_CONFIG, ...JSON.parse(data) };
  } catch (e) {
    return { ...DEFAULT_CONFIG };
  }
}

function saveConfig(config) {
  const dir = path.dirname(CONFIG_PATH);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(config, null, 2));
}

/**
 * Get current email config (for Settings UI). Masks secrets in UI.
 */
function getEmailConfig() {
  const c = loadConfig();
  return {
    ...c,
    smtpPass: c.smtpPass ? '********' : '',
    mailgunApiKey: c.mailgunApiKey ? '********' : ''
  };
}

/**
 * Save email config (from Settings UI). Preserve existing password if UI sends mask.
 */
function setEmailConfig(config) {
  const current = loadConfig();
  const next = {
    emailOscarOnError: config.emailOscarOnError !== undefined ? config.emailOscarOnError : current.emailOscarOnError,
    oscarEmail: (config.oscarEmail !== undefined ? config.oscarEmail : current.oscarEmail) || '',
    useMailgun: config.useMailgun !== undefined ? !!config.useMailgun : current.useMailgun !== false,
    mailgunApiKey: (config.mailgunApiKey !== undefined && config.mailgunApiKey !== '********') ? config.mailgunApiKey : current.mailgunApiKey,
    mailgunDomain: (config.mailgunDomain !== undefined ? config.mailgunDomain : current.mailgunDomain) || '',
    smtpHost: (config.smtpHost !== undefined ? config.smtpHost : current.smtpHost) || '',
    smtpPort: config.smtpPort !== undefined ? Number(config.smtpPort) : (current.smtpPort || 587),
    smtpSecure: config.smtpSecure !== undefined ? !!config.smtpSecure : !!current.smtpSecure,
    smtpUser: (config.smtpUser !== undefined ? config.smtpUser : current.smtpUser) || '',
    smtpPass: (config.smtpPass !== undefined && config.smtpPass !== '********') ? config.smtpPass : current.smtpPass,
    fromEmail: (config.fromEmail !== undefined ? config.fromEmail : current.fromEmail) || ''
  };
  saveConfig(next);
  return getEmailConfig();
}

/**
 * Send via Mailgun API (POST to api.mailgun.net).
 */
function sendViaMailgun(config, from, to, subject, text) {
  return new Promise((resolve, reject) => {
    const domain = config.mailgunDomain.trim();
    const auth = Buffer.from('api:' + config.mailgunApiKey).toString('base64');
    const body = new URLSearchParams({
      from: from,
      to: to,
      subject: subject,
      text: text
    }).toString();

    const req = https.request({
      hostname: 'api.mailgun.net',
      path: '/v3/' + encodeURIComponent(domain) + '/messages',
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + auth,
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body)
      }
    }, (res) => {
      let data = '';
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve();
        } else {
          reject(new Error('Mailgun: ' + (res.statusCode + ' ' + (data || res.statusMessage))));
        }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

/**
 * Send error notification to Oscar. Uses Mailgun if configured, otherwise SMTP.
 */
async function sendErrorNotificationToOscar(mismatchData) {
  const config = loadConfig();
  if (!config.emailOscarOnError || !config.oscarEmail || !config.oscarEmail.trim()) {
    return;
  }

  const { expectedDescription, actualDescription, timestamp } = mismatchData;
  const productDesc = (expectedDescription != null && String(expectedDescription).trim()) ? String(expectedDescription).trim() : (mismatchData.expected || '—');
  const lpnDesc = (actualDescription != null && String(actualDescription).trim()) ? String(actualDescription).trim() : (mismatchData.actual || '—');
  const timeStr = timestamp ? new Date(timestamp).toLocaleString() : (timestamp || '—');
  const subject = 'EOL Scanner Error';
  const text = `The Product Scanned was ${productDesc} and the LPN scanned was ${lpnDesc} at this timestamp: ${timeStr}`;
  const to = config.oscarEmail.trim();

  const useMailgun = config.useMailgun !== false && config.mailgunApiKey && config.mailgunDomain && config.mailgunDomain.trim();

  if (useMailgun) {
    const domain = config.mailgunDomain.trim();
    const from = (config.fromEmail && config.fromEmail.trim()) || ('EOL Scanner <postmaster@' + domain + '>');
    try {
      await sendViaMailgun(config, from, to, subject, text);
      console.log('[Email] Sent via Mailgun to', to);
    } catch (err) {
      console.error('[Email] Mailgun send failed:', err.message);
    }
    return;
  }

  if (!config.smtpHost || !config.smtpHost.trim()) {
    console.warn('[Email] Neither Mailgun nor SMTP configured. Set Mailgun (API key + domain) or SMTP in Settings.');
    return;
  }

  const from = (config.fromEmail && config.fromEmail.trim()) || config.smtpUser || 'e80-scanner@local';
  let transporter;
  try {
    transporter = nodemailer.createTransport({
      host: config.smtpHost.trim(),
      port: config.smtpPort || 587,
      secure: !!config.smtpSecure,
      auth: (config.smtpUser && config.smtpPass) ? {
        user: config.smtpUser.trim(),
        pass: config.smtpPass
      } : undefined
    });
  } catch (err) {
    console.error('[Email] Failed to create SMTP transporter:', err.message);
    return;
  }

  try {
    await transporter.sendMail({ from, to, subject, text });
    console.log('[Email] Sent via SMTP to', to);
  } catch (err) {
    console.error('[Email] SMTP send failed:', err.message);
  }
}

module.exports = {
  getEmailConfig,
  setEmailConfig,
  sendErrorNotificationToOscar
};
