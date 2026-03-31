/**
 * Physical stack light + reset button on Raspberry Pi.
 *
 * Backends (auto):
 *   1) onoff (sysfs) — older Pi OS
 *   2) pigpio pigs — Raspberry Pi OS Bookworm+ (no /sys/class/gpio); no Electron native rebuild needed
 *
 * Stack light default: BCM GPIO 17 (physical pin 11 on 40-pin header).
 * Reset default: BCM GPIO 9 (physical pin 21). Pin 11 is BCM 17 — do not use for reset if the light is on 17.
 * Reset wiring: default expects pressed = GND (pull-up). If pressed = 3.3 V (idle low), set GPIO_RESET_ACTIVE_HIGH=1
 * and use internal pull-down (pigpio); wire so released = low, pressed = high.
 *
 * Pigpio setup on Pi:
 *   sudo apt install pigpio
 *   sudo systemctl enable pigpiod
 *   sudo systemctl start pigpiod
 *   sudo usermod -a -G gpio $USER   # then re-login
 *
 * Env: ENABLE_GPIO=0, GPIO_PIN=17, STACK_LIGHT_ACTIVE_LOW=1, GPIO_RESET_PIN=9,
 *      GPIO_RESET_ACTIVE_HIGH=1 (pressed reads 1 / 3.3 V; idle 0), ENABLE_GPIO_RESET=0,
 *      GPIO_BACKEND=pigpio|onoff (force)
 */

const { execFileSync } = require('child_process');

let alarmActive = false;
let gpioOut = null;
let gpioResetIn = null;
/** @type {'onoff'|'pigpio'|null} */
let gpioBackend = null;
let pigpioOutPin = -1;
let pigpioResetPoll = null;
/** Consecutive polls reading "pressed" level (debounce). */
let pigpioResetPressStreak = 0;
let gpioOutputReady = false;
let gpioOutputSkip = false;
let gpioOutputFailed = false;
let resetWatcherStarted = false;
let shutdownRegistered = false;

function getGpioPinNumber() {
  const n = parseInt(process.env.GPIO_PIN || '17', 10);
  return Number.isFinite(n) && n >= 0 ? n : 17;
}

function getResetPinNumber() {
  if (process.env.ENABLE_GPIO_RESET === '0') return null;
  const raw = process.env.GPIO_RESET_PIN;
  if (raw === '' || raw === 'false' || raw === 'off') return null;
  const n = parseInt(raw != null && raw !== '' ? raw : '9', 10);
  if (!Number.isFinite(n) || n < 0) return 9;
  return n;
}

function activeLow() {
  return process.env.STACK_LIGHT_ACTIVE_LOW === '1' || process.env.STACK_LIGHT_ACTIVE_LOW === 'true';
}

/**
 * True: idle low, pressed connects 3.3 V (read 1), internal pull-down.
 * False: pull-up, pressed shorts to GND (read 0).
 * Override: GPIO_RESET_ACTIVE_LOW=1 for GND switch; GPIO_RESET_ACTIVE_HIGH=0 to force low-active.
 */
function resetActiveHigh() {
  const low = (process.env.GPIO_RESET_ACTIVE_LOW || '').toLowerCase();
  if (low === '1' || low === 'true' || low === 'yes') return false;
  const hi = (process.env.GPIO_RESET_ACTIVE_HIGH || '').toLowerCase();
  if (hi === '0' || hi === 'false' || hi === 'no') return false;
  if (hi === '1' || hi === 'true' || hi === 'yes') return true;
  return true;
}

function resetPinPressed(reading) {
  return resetActiveHigh() ? reading === 1 : reading === 0;
}

function pigs(args) {
  execFileSync('pigs', args, { stdio: ['ignore', 'pipe', 'pipe'], timeout: 3000, encoding: 'utf8' });
}

function pigpioDaemonOk() {
  try {
    execFileSync('pigs', ['t'], { stdio: ['ignore', 'pipe', 'pipe'], timeout: 2000 });
    return true;
  } catch {
    return false;
  }
}

function initPigpioOutput(pin) {
  if (!pigpioDaemonOk()) {
    console.warn('[Alarm] pigpio daemon not running. Install and start: sudo apt install pigpio && sudo systemctl start pigpiod');
    return false;
  }
  try {
    pigs(['m', String(pin), '1']);
    const startLow = activeLow() ? 1 : 0;
    pigs(['w', String(pin), String(startLow)]);
    gpioBackend = 'pigpio';
    pigpioOutPin = pin;
    gpioOutputReady = true;
    console.log('[Alarm] pigpio: BCM', pin, 'stack light (active', activeLow() ? 'LOW' : 'HIGH', ')');
    return true;
  } catch (e) {
    console.warn('[Alarm] pigpio stack light failed:', e.message);
    return false;
  }
}

function writePigpioOut(on) {
  if (pigpioOutPin < 0) return;
  const high = activeLow() ? !on : on;
  try {
    pigs(['w', String(pigpioOutPin), high ? '1' : '0']);
  } catch (e) {
    console.error('[Alarm] pigpio write failed:', e.message);
  }
}

function writeAlarmState(on) {
  if (gpioBackend === 'pigpio') {
    writePigpioOut(on);
    return;
  }
  if (!gpioOut) return;
  const high = activeLow() ? !on : on;
  try {
    gpioOut.writeSync(high ? 1 : 0);
  } catch (e) {
    console.error('[Alarm] GPIO write failed:', e.message);
  }
}

function registerShutdown() {
  if (shutdownRegistered) return;
  shutdownRegistered = true;
  function shutdown() {
    if (pigpioResetPoll) {
      clearInterval(pigpioResetPoll);
      pigpioResetPoll = null;
    }
    if (gpioResetIn) {
      try {
        gpioResetIn.unwatch();
        gpioResetIn.unexport();
      } catch (err) {
        console.warn('[Alarm] Reset GPIO shutdown:', err.message);
      }
      gpioResetIn = null;
    }
    if (gpioBackend === 'pigpio' && pigpioOutPin >= 0) {
      try {
        const off = activeLow() ? 1 : 0;
        pigs(['w', String(pigpioOutPin), String(off)]);
      } catch (_) {}
      pigpioOutPin = -1;
      gpioBackend = null;
    }
    if (gpioOut) {
      try {
        gpioOut.writeSync(activeLow() ? 1 : 0);
        gpioOut.unexport();
      } catch (err) {
        console.warn('[Alarm] Stack light GPIO shutdown:', err.message);
      }
      gpioOut = null;
    }
  }
  process.once('SIGINT', shutdown);
  process.once('SIGTERM', shutdown);
}

function initGpio() {
  if (gpioOutputSkip || gpioOutputReady) return;
  if (gpioOutputFailed) return;

  if (process.env.ENABLE_GPIO === '0') {
    gpioOutputSkip = true;
    console.log('[Alarm] GPIO disabled (ENABLE_GPIO=0)');
    return;
  }

  if (process.platform !== 'linux') {
    gpioOutputSkip = true;
    console.log('[Alarm] GPIO skipped (not Linux)');
    return;
  }

  const force = (process.env.GPIO_BACKEND || '').toLowerCase();
  const pin = getGpioPinNumber();

  if (force === 'pigpio') {
    if (initPigpioOutput(pin)) {
      registerShutdown();
    } else {
      gpioOutputFailed = true;
      registerShutdown();
    }
    return;
  }

  if (force !== 'onoff') {
    if (initPigpioOutput(pin)) {
      registerShutdown();
      return;
    }
  }

  try {
    const { Gpio } = require('onoff');
    gpioOut = new Gpio(pin, 'out');
    gpioOut.writeSync(activeLow() ? 1 : 0);
    gpioBackend = 'onoff';
    gpioOutputReady = true;
    console.log('[Alarm] onoff: BCM', pin, 'stack light (active', activeLow() ? 'LOW' : 'HIGH', ')');
  } catch (e) {
    gpioOut = null;
    console.warn('[Alarm] onoff stack light failed:', e.message);
    if (force === 'onoff') {
      gpioOutputFailed = true;
      registerShutdown();
      return;
    }
    if (initPigpioOutput(pin)) {
      registerShutdown();
      return;
    }
    gpioOutputFailed = true;
    console.warn('[Alarm] No GPIO backend available. For Bookworm Pi OS run: sudo apt install pigpio && sudo systemctl start pigpiod');
    registerShutdown();
    return;
  }

  registerShutdown();
}

function initStackLightGpioAtStartup() {
  initGpio();
  writeAlarmState(false);
}

function startPigpioResetPoll(resetPin) {
  if (pigpioResetPoll || !pigpioDaemonOk()) return;
  try {
    pigs(['m', String(resetPin), '0']);
    // Active-low button: pull up so open = 1, GND press = 0. Active-high: pull down so open = 0, 3.3 V press = 1.
    pigs(['pud', String(resetPin), resetActiveHigh() ? '1' : '2']);
  } catch (e) {
    console.warn('[Alarm] pigpio reset pin setup failed:', e.message);
    return;
  }
  pigpioResetPressStreak = 0;
  pigpioResetPoll = setInterval(() => {
    let v;
    try {
      const out = execFileSync('pigs', ['r', String(resetPin)], { encoding: 'utf8', timeout: 2000 }).trim();
      v = parseInt(out, 10);
      if (!Number.isFinite(v)) return;
    } catch {
      return;
    }
    if (alarmActive) {
      if (resetPinPressed(v)) {
        pigpioResetPressStreak += 1;
        if (pigpioResetPressStreak >= 3) {
          pigpioResetPressStreak = 0;
          clearAlarm();
          console.log('[Alarm] Cleared by GPIO reset button (BCM', resetPin + ', pigpio)');
        }
      } else {
        pigpioResetPressStreak = 0;
      }
    } else {
      pigpioResetPressStreak = 0;
    }
  }, 40);
  console.log(
    '[Alarm] pigpio: BCM',
    resetPin,
    'reset button (poll,',
    resetActiveHigh() ? 'active-HIGH pressed=1' : 'active-LOW pressed=Gnd',
    ')'
  );
  registerShutdown();
}

function startAlarmResetButtonWatcher() {
  if (resetWatcherStarted) return;
  resetWatcherStarted = true;

  if (process.env.ENABLE_GPIO === '0') return;
  if (process.env.ENABLE_GPIO_RESET === '0') {
    console.log('[Alarm] GPIO reset button disabled (ENABLE_GPIO_RESET=0)');
    return;
  }
  if (process.platform !== 'linux') return;

  const resetPin = getResetPinNumber();
  if (resetPin == null) return;

  const outPin = getGpioPinNumber();
  if (resetPin === outPin) {
    console.warn('[Alarm] GPIO_RESET_PIN cannot match GPIO_PIN');
    return;
  }

  const force = (process.env.GPIO_BACKEND || '').toLowerCase();

  if (gpioBackend === 'pigpio' || (force !== 'onoff' && pigpioDaemonOk())) {
    startPigpioResetPoll(resetPin);
    return;
  }

  try {
    const { Gpio } = require('onoff');
    const edge = resetActiveHigh() ? 'rising' : 'falling';
    gpioResetIn = new Gpio(resetPin, 'in', edge, { debounceTimeout: 120 });
    gpioResetIn.watch((err) => {
      if (err) {
        console.error('[Alarm] Reset button watch error:', err.message);
        return;
      }
      if (alarmActive) {
        clearAlarm();
        console.log('[Alarm] Cleared by GPIO reset button (BCM', resetPin + ')');
      }
    });
    console.log('[Alarm] onoff: BCM', resetPin, 'reset button');
    registerShutdown();
  } catch (e) {
    console.warn('[Alarm] onoff reset button failed:', e.message);
    if (pigpioDaemonOk()) {
      startPigpioResetPoll(resetPin);
    }
  }
}

function triggerAlarm() {
  alarmActive = true;
  initGpio();
  writeAlarmState(true);
  console.log('[Alarm] TRIGGERED - mismatch detected');
}

function clearAlarm() {
  alarmActive = false;
  initGpio();
  writeAlarmState(false);
  console.log('[Alarm] CLEARED');
}

function isAlarmActive() {
  return alarmActive;
}

module.exports = {
  triggerAlarm,
  clearAlarm,
  isAlarmActive,
  startAlarmResetButtonWatcher,
  initStackLightGpioAtStartup
};
