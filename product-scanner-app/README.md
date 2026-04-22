# E80 Product Scanner App

Compares **LPN scanner** vs **Product scanner** on the line. The product scanner should always have a read; the LPN can be missing. Possible outcomes: **Match** (no error; LPN and product agree), **LPN Missing** (no LPN read), **LPN invalid SKU** (LPN doesn’t contain a valid SKU), or **LPN wrong product** (LPN has a valid SKU that doesn’t match the current product). Alerts use a **physical alarm**; optionally **email Oscar** when there is an error (see Settings).

## Features

- **LPN vs Product only**: No “expected” product box. We only check that the two scanners agree and that LPN is present when product is present.
- **Error types**: Only errors are logged. Possible errors: **LPN Missing** (product scanned, no LPN); **LPN invalid SKU** (LPN doesn’t contain valid SKU); **LPN wrong product** (LPN has valid SKU but doesn’t match current product). No error = **Match**.
- **Physical alarm**: Ready for hardware (buzzer/relay/GPIO). See `alarm-service.js`. Triggered on mismatch or LPN missing; override clears it.
- **Email Oscar on error**: Optional. In **Settings**, enable "Email Oscar when there is an error" and set Oscar's email. When a mismatch or LPN-missing occurs, a notification is sent (placeholder: logs to console; wire to SMTP/mail API in `email-service.js` for production).
- **Barcode product identification**: Product scanner codes (UPC/EAN) are looked up via a **local cache first** (`data/barcode-cache.json`). Only on cache miss do we call the [UPCitemDB](https://www.upcitemdb.com) free API, then we store the result for next time. We also skip lookup when the scanned product hasn’t changed (same barcode scanned again), so the free tier isn’t used up. For testing, seed the cache and it becomes the main source. The resolved product name (and brand) is shown under the Product scanner box. LPN barcode format is TBD; once known, LPN can be parsed and linked here.
- **Export data**: Export error/mismatch history to CSV with optional date range.
- **Date range filters**: Main view, Error History, and Advanced Statistics each support optional date ranges.
- **Item list CSV = backbone**: The item list CSV (e.g. in `data/item-lists/`) is how everything is linked: barcode ↔ legacy item name ↔ product name. Product scanner and LPN scanner both look up against it; match/mismatch is decided from it. Import via **Edit Item Master → Import CSV**. Data is stored in `data/product-master.json`; re-import when you update the CSV.
- **Error history**: Scrollable list with Override action.
- **Statistics & advanced stats**: Totals, worst hours/days, worst products, and bar charts by hour/day.

## Installation

**Recommended:** Use **Node.js 20 LTS** for best compatibility with Electron. If you use [nvm](https://github.com/nvm-sh/nvm): `nvm install 20 && nvm use 20`.

```bash
npm install
npm start
```

### Troubleshooting

- **`Electron failed to install correctly`** or **`downloadArtifact` / `pathExists` errors**  
  Delete `node_modules/electron` and run `npm install` again. If it still fails, use Node 20 LTS (see above). The `package.json` may include an `overrides` entry for `fs-extra` to help with Node 24.

- **`Cannot read properties of undefined (reading 'whenReady')`**  
  This usually means `require('electron')` in the main process is resolving to the npm package (executable path) instead of the Electron API. Use **Node.js 20 LTS** and reinstall:  
  `rm -rf node_modules package-lock.json && npm install && npm start`.  
  If you don’t have nvm, install Node 20 from [nodejs.org](https://nodejs.org/).

## Joining from your phone

With the app running, a **viewer server** runs on port **3847**. Supervisors can open the same UI in a browser (view live scans, override, edit item master, export, settings).

- **On this computer:** open **http://localhost:3847**
- **On your phone (same Wi‑Fi):** when you open the viewer in a browser on this machine, a bar at the top shows the URL to use on your phone (e.g. `http://192.168.1.68:3847`). Open that **exact** URL in your phone’s browser. Use **http** (not https) and port **3847**. The terminal also prints this URL when the app starts.

**If your phone can’t reach the URL:**

1. **Full URL in the address bar:** Use **`http://10.10.54.55:3847`** (include **`http://`**, not `https://`). Typing `10.10.54.55:3847` alone can fail or use search instead of opening the page.
2. **Same subnet:** On the Capstone / showcase Wi‑Fi, your **laptop or phone** should have an address in the **same range** as the Pi (e.g. if the Pi is `10.10.54.55`, the client should be `10.10.54.x`). If the laptop is `192.168.x.x` while the Pi is `10.10.54.x`, you are on a different network path and need to connect to the same SSID/VLAN the Pi uses.
3. **Raspberry Pi (not Mac):** Use the **IP printed in the terminal when the app starts** (or `hostname -I` on the Pi). Test from the Pi: `curl -s -o /dev/null -w "%{http_code}" http://127.0.0.1:3847/api/state` → **200**. From your laptop: `curl` the Pi’s `http://IP:3847/api/state` → **200** if the network allows device-to-device traffic.
4. **Event / school “special” Wi‑Fi** often enables **AP isolation** or **client isolation** (devices can reach the internet but not each other). That blocks phones/laptops from reaching the Pi. Fix: router admin, or a different SSID, or a wired path — network staff has to allow client-to-client on that VLAN.
5. **Pi firewall (if enabled):** `sudo ufw allow 3847/tcp` then `sudo ufw reload` (skip if ufw is inactive).
6. **Confirm a computer’s IP:** **Mac:** System Settings → Network → Wi‑Fi → Details, or `ifconfig` / `ipconfig getifaddr en0` in Terminal. The IP can change (DHCP).
7. **Same Wi‑Fi as the Pi (not guest / not cellular** on the phone) for client devices that should load the viewer.
8. **Router “AP isolation” (repeat):** Turn off for the SSID the Pi uses if you control the router.

The item master and all data are the same for the host app and every viewer.

## Default password

Admin password: `admin123`. Change it in `main.js` for production.

## Building

```bash
npm run build
```

Output is in the `dist` folder.

## Project structure

```
product-scanner-app/
├── main.js                 # Electron main, IPC, viewer server
├── renderer.js             # UI, scans, item master, export
├── index.html
├── styles.css
├── package.json
├── services/               # Backend modules
│   ├── alarm-service.js    # Physical alarm (stub for hardware)
│   ├── barcode-lookup-service.js  # Cache + UPCitemDB / Open Food Facts
│   ├── database.js         # JSON file store for mismatches/stats (no native sqlite)
│   ├── email-service.js    # Email Oscar on error (placeholder)
│   ├── error-types.js      # Labels: match, lpn_missing, lpn_invalid_sku, lpn_wrong_product
│   └── scan-log.js         # Append scan log (JSONL)
├── data/
│   ├── product-master.json # Item master (barcode, legacy item name, mfg description)
│   ├── item-lists/         # Put your item list CSV file(s) here (see README inside)
│   ├── email-config.json
│   ├── barcode-cache.json
│   ├── scan-log.jsonl
│   └── mismatches-store.json  # error history (JSON; old mismatches.db not auto-imported)
├── assets/                 # Logos, icon
└── scripts/                # Helper scripts (e.g. fix-electron-sigbus.sh)
```

## Physical alarm (stack light on Raspberry Pi)

`services/alarm-service.js` drives a **stack light** (or relay) via **GPIO** on Linux (Raspberry Pi). It runs in the **main process** only.

- **Default pin**: **BCM GPIO 17** (physical pin **11**). **BCM 10** (pin 19) is **SPI MOSI** and often causes **`EINVAL` on write** if SPI is on — use `GPIO_PIN=10 npm start` only if SPI is disabled and you wired pin 19. Wire your driver to **3.3 V logic** (GPIO is not 5 V tolerant).
- **Electron on Pi:** hardware acceleration is disabled on Linux by default to reduce `GpuControl` / GPU crashes. Set `ELECTRON_DISABLE_GPU=0` before `npm start` to re-enable if needed.
- **When**: `triggerAlarm()` runs when a mismatch is **logged** (after the 5 s delay). `clearAlarm()` runs on **Override** in the app, on a **scan match**, or on the **physical reset button**. The physical button also marks the **newest pending** mismatch as **override** (same as the Override button) and refreshes the stats table.
- **Reset button (input, optional)**: default **BCM GPIO 9** (physical pin **21** on the 40-pin header — **not** pin 11; pin 11 is BCM **17**, the default stack light). **Default wiring**: idle **low** (0), pressed feeds **3.3 V** (read **1**); pigpio uses an internal **pull-down**. Pressing clears the stack light when the alarm is on. For a switch to **GND** instead (pull-up, press = 0), set **`GPIO_RESET_ACTIVE_LOW=1`** before `npm start`.
- **Env (optional)**:
  - `ENABLE_GPIO=0` — disable GPIO (e.g. development on a PC); alarm still logs to console.
  - `GPIO_PIN=17` — stack light output, BCM (default **17** / physical pin 11).
  - `STACK_LIGHT_ACTIVE_LOW=1` — set if your relay module turns **on** when GPIO is **LOW**.
  - `GPIO_RESET_PIN=9` — reset button input, BCM (default 9). Set `ENABLE_GPIO_RESET=0` to disable the button watcher only.
  - `GPIO_RESET_ACTIVE_LOW=1` — reset button shorts to **GND** (active-low). Omit for **3.3 V when pressed** (default).

**Permissions**: add the user running Electron to the `gpio` group, then reboot:  
`sudo usermod -a -G gpio $USER`

**Check that the Pi sees the button (same stack as the app — pigpio):**

1. `sudo systemctl start pigpiod` and `pigs t` — should print a number (daemon OK).
2. Quit the scanner app so nothing else is using the pin, then from `product-scanner-app/`:  
   `node scripts/poll-gpio-pin.js 9 down`  
   Hold the button: you should see `read = 1` (or `0` when released). Wrong pin? Try `17` if the wire is on physical pin 11 (BCM 17).
3. One-shot: `pigs m 9 0; pigs pud 9 d; pigs r 9` — prints `0` or `1` (`pud` uses **`d`**/**`u`**/**`o`**, not numbers).

## Raspberry Pi deployment

This app is intended to run on a **Raspberry Pi** with scanners and the stack light connected.

- **Electron on Pi**: Install Node.js on the Pi, clone the app, run `npm install` and `npm start`. `onoff` is optional on non-Linux; on the Pi it should install and `electron-rebuild` will rebuild native modules for Electron.
- **Two USB scanners**: Plug both into the Pi; the app uses one keyboard buffer and distinguishes LPN (letters) vs product (digits only).
- **Hardware in Node, not the UI**: GPIO and alarms belong in the main process (`alarm-service.js`), not the renderer.

## Barcode scanner (continuous mode) — one scanner for both product and LPN

You can use **a single USB scanner** for both product and LPN. The app detects the type by **letters vs digits** (they have many differences):

- **Contains any letter** (e.g. GY) → **LPN**. The 7-digit legacy item name (before the last 3 digits) is matched to the item master; product name is shown as the LPN box heading.
- **Digits only** → **product barcode** (UPC/EAN). Never mixed up with LPN.

Plug in the scanner (keyboard/HID mode), keep the app window focused, and scan: product barcodes and LPNs in any order. No need for two scanners or to switch inputs.

## Usage

1. **Main view**: Plug in the Product scanner; scan to see the last code. Set date range (optional). Use "Export Data" for CSV.
2. **Error history**: Optional date range, then scroll and use Override as needed.
3. **Advanced Statistics**: Set date range and click Apply to filter worst times/days and products.
4. **Edit Item Master**: Add/delete products (password required).

## Author

Gaelen Collins – UGA Capstone Project
