#!/usr/bin/env node
/**
 * Live read of a BCM GPIO pin via pigpio (pigs). Run on the Pi, not on your Mac.
 *
 * Usage:
 *   node scripts/poll-gpio-pin.js [BCM_PIN] [pull]
 *
 * BCM_PIN default: 9 (reset button default in the app)
 * pull: down | up | off  (default down — matches app active-high button: idle 0, pressed 3.3 V → 1)
 *
 * Stop the Electron app first so only one process owns the pin.
 *
 * Examples:
 *   node scripts/poll-gpio-pin.js 9 down
 *   node scripts/poll-gpio-pin.js 17 down   # if you wired the button to BCM 17 by mistake
 */

const { execFileSync } = require('child_process');

const pin = process.argv[2] || '9';
const pullArg = (process.argv[3] || 'down').toLowerCase();
// pigs pud: d = pull-down, u = pull-up, o = off (numeric codes are invalid)
const pudMap = { down: 'd', up: 'u', off: 'o' };
const pud = pudMap[pullArg];
if (!pud) {
  console.error('pull must be down, up, or off');
  process.exit(1);
}

function pigs(args) {
  return execFileSync('pigs', args, {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    timeout: 3000
  }).trim();
}

console.log('--- GPIO read test (pigpio pigs) ---');
try {
  pigs(['t']);
} catch (e) {
  console.error('Cannot talk to pigpiod. On the Pi run: sudo systemctl start pigpiod');
  console.error(e.message);
  process.exit(1);
}

console.log(`BCM ${pin}: mode input, pud ${pullArg} (${pud === 'd' ? 'pull-down' : pud === 'u' ? 'pull-up' : 'none'})`);
console.log('Expected for your wiring (idle 0 V, pressed 3.3 V): pull-down, idle read 0, pressed read 1');
console.log('Ctrl+C to stop.\n');

try {
  pigs(['m', String(pin), '0']);
  pigs(['pud', String(pin), pud]);
} catch (e) {
  console.error('Setup failed:', e.message);
  process.exit(1);
}

const interval = setInterval(() => {
  try {
    const v = pigs(['r', String(pin)]);
    const n = parseInt(String(v).trim(), 10);
    const ts = new Date().toISOString();
    console.log(ts, 'BCM' + pin, 'read =', String(v).trim(), Number.isFinite(n) ? `(parsed ${n})` : '');
  } catch (e) {
    console.error('read error:', e.message);
  }
}, 200);

process.on('SIGINT', () => {
  clearInterval(interval);
  console.log('\nExiting.');
  process.exit(0);
});
