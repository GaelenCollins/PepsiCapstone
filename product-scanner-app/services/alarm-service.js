/**
 * Physical stack light / alarm on Raspberry Pi GPIO + optional reset button.
 *
 * Stack light (output):
 *   Default BCM GPIO 10 (physical pin 19 on 40-pin header).
 *
 * Reset button (input, optional):
 *   Default BCM GPIO 9 (physical pin 21). Wire: one side to GPIO, other to GND (internal pull-up).
 *   Press clears the stack light (same as clear-alarm in the app). Does not change DB / override.
 *
 * Env (optional):
 *   ENABLE_GPIO=0          — disable all GPIO (stack light + reset button)
 *   GPIO_PIN=10            — stack light output (BCM)
 *   STACK_LIGHT_ACTIVE_LOW=1 — alarm ON = GPIO LOW
 *   GPIO_RESET_PIN=9       — reset button input (BCM); set empty or ENABLE_GPIO_RESET=0 to disable
 *   ENABLE_GPIO_RESET=0    — disable only the reset button (keep stack light GPIO)
 */

let alarmActive = false;
let gpioOut = null;
let gpioResetIn = null;
let gpioInitAttempted = false;
let resetWatcherStarted = false;
let shutdownRegistered = false;

function getGpioPinNumber() {
  const n = parseInt(process.env.GPIO_PIN || '10', 10);
  return Number.isFinite(n) && n >= 0 ? n : 10;
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

function writeAlarmState(on) {
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
    if (gpioResetIn) {
      try {
        gpioResetIn.unwatch();
        gpioResetIn.unexport();
      } catch (err) {
        console.warn('[Alarm] Reset GPIO shutdown:', err.message);
      }
      gpioResetIn = null;
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
  if (gpioInitAttempted) return;
  gpioInitAttempted = true;

  if (process.env.ENABLE_GPIO === '0') {
    console.log('[Alarm] GPIO disabled (ENABLE_GPIO=0)');
    return;
  }

  if (process.platform !== 'linux') {
    console.log('[Alarm] GPIO skipped (not Linux — use Raspberry Pi for stack light)');
    return;
  }

  try {
    const { Gpio } = require('onoff');
    const pin = getGpioPinNumber();
    gpioOut = new Gpio(pin, 'out');
    gpioOut.writeSync(activeLow() ? 1 : 0);
    console.log('[Alarm] GPIO BCM', pin, 'ready for stack light (active', activeLow() ? 'LOW' : 'HIGH', ')');
  } catch (e) {
    console.warn('[Alarm] Stack light GPIO not available:', e.message);
    gpioOut = null;
  }

  registerShutdown();
}

/**
 * Watch GPIO for physical reset button (falling edge = pressed to GND).
 * Call once from main process when app is ready.
 */
function startAlarmResetButtonWatcher() {
  if (resetWatcherStarted) return;
  resetWatcherStarted = true;

  if (process.env.ENABLE_GPIO === '0') {
    return;
  }
  if (process.env.ENABLE_GPIO_RESET === '0') {
    console.log('[Alarm] GPIO reset button disabled (ENABLE_GPIO_RESET=0)');
    return;
  }

  if (process.platform !== 'linux') {
    return;
  }

  const resetPin = getResetPinNumber();
  if (resetPin == null) {
    return;
  }

  const outPin = getGpioPinNumber();
  if (resetPin === outPin) {
    console.warn('[Alarm] GPIO_RESET_PIN cannot match GPIO_PIN (stack light); reset button not started');
    return;
  }

  try {
    const { Gpio } = require('onoff');
    gpioResetIn = new Gpio(resetPin, 'in', 'falling', { debounceTimeout: 120 });
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
    console.log('[Alarm] GPIO BCM', resetPin, 'ready for alarm reset button (falling edge → GND)');
    registerShutdown();
  } catch (e) {
    console.warn('[Alarm] Reset button GPIO not available:', e.message);
    gpioResetIn = null;
  }
}

/**
 * Trigger the physical alarm (stack light on).
 */
function triggerAlarm() {
  alarmActive = true;
  initGpio();
  writeAlarmState(true);
  console.log('[Alarm] TRIGGERED - mismatch detected');
}

/**
 * Turn off the physical alarm (stack light off).
 */
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
  startAlarmResetButtonWatcher
};
