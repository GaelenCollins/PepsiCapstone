/**
 * Physical stack light / alarm on Raspberry Pi GPIO.
 * Default: BCM GPIO 10 (physical pin 19 on 40-pin header — matches common "GPIO 10" silkscreen).
 *
 * Wiring: connect your stack-light driver (relay/transistor module) to 3.3V/GND/GPIO as your board requires.
 * Many relay modules use VCC, GND, and IN — IN often goes to GPIO; check active-high vs active-low for your module.
 *
 * Env (optional):
 *   ENABLE_GPIO=0     — disable GPIO (logs only; use on dev machine)
 *   GPIO_PIN=10       — BCM pin number (default 10)
 *   STACK_LIGHT_ACTIVE_LOW=1 — if set, alarm ON = GPIO LOW (common for some relay boards)
 */

let alarmActive = false;
let gpioOut = null;
let gpioInitAttempted = false;

function getGpioPinNumber() {
  const n = parseInt(process.env.GPIO_PIN || '10', 10);
  return Number.isFinite(n) && n >= 0 ? n : 10;
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
    console.warn('[Alarm] GPIO not available:', e.message);
    gpioOut = null;
  }

  function shutdown() {
    if (gpioOut) {
      try {
        gpioOut.writeSync(activeLow() ? 1 : 0);
        gpioOut.unexport();
      } catch (err) {
        console.warn('[Alarm] GPIO shutdown:', err.message);
      }
      gpioOut = null;
    }
  }
  process.once('SIGINT', shutdown);
  process.once('SIGTERM', shutdown);
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
  isAlarmActive
};
