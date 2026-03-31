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

1. **Check the URL:** Use **http** (not https) and port **3847**. Example: `http://192.168.1.68:3847`.
2. **Confirm the Mac’s IP:** On the Mac, open **System Settings → Network → Wi‑Fi → Details** (or run `ifconfig | grep "inet "` in Terminal). The IP might have changed (DHCP).
3. **Test from the Mac first:** With the app running, on the Mac open Terminal and run:  
   `curl -s -o /dev/null -w "%{http_code}" http://YOUR_MAC_IP:3847/api/state`  
   (replace `YOUR_MAC_IP` with the same IP you use on the phone, e.g. `192.168.1.68`). If you see **200**, the server is reachable on that IP; the problem is then between the phone and the Mac.
4. **Same Wi‑Fi:** Phone and Mac must be on the **same** Wi‑Fi network (not guest network, not cellular on the phone).
5. **Router “AP isolation” / “client isolation”:** Many routers have a setting that blocks Wi‑Fi devices from talking to each other. In the router’s admin page, turn that off so the phone can reach the Mac.

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
│   ├── database.js         # SQLite (mismatches, stats)
│   ├── email-service.js    # Email Oscar on error (placeholder)
│   ├── error-types.js      # Labels: match, lpn_missing, lpn_invalid_sku, lpn_wrong_product
│   └── scan-log.js         # Append scan log (JSONL)
├── data/
│   ├── product-master.json # Item master (barcode, legacy item name, mfg description)
│   ├── item-lists/         # Put your item list CSV file(s) here (see README inside)
│   ├── email-config.json
│   ├── barcode-cache.json
│   ├── scan-log.jsonl
│   └── mismatches.db
├── assets/                 # Logos, icon
└── scripts/                # Helper scripts (e.g. fix-electron-sigbus.sh)
```

## Physical alarm (stack light on Raspberry Pi)

`services/alarm-service.js` drives a **stack light** (or relay) via **GPIO** on Linux (Raspberry Pi). It runs in the **main process** only.

- **Default pin**: **BCM GPIO 10** (often labeled “GPIO 10” on diagrams; physical pin **19** on the 40-pin header). Wire your driver circuit to **3.3 V logic** on the Pi (GPIO is not 5 V tolerant — use a level shifter or a 3.3 V–compatible relay module if your board expects 5 V on the control pin). **Note:** BCM 10 is **SPI MOSI**; if SPI is enabled in `raspi-config`, use another pin (`GPIO_PIN=17 npm start`) or disable SPI.
- **When**: `triggerAlarm()` runs when a mismatch is **logged** (after the 5 s delay). `clearAlarm()` runs on **Override** in the app.
- **Reset button (input, optional)**: default **BCM GPIO 9** (physical pin **21**). Wire the button between **GPIO 9** and **GND** (uses internal pull-up; press = connect to GND). Pressing clears the stack light if the alarm is on (same as software clear; does not override DB records).
- **Env (optional)**:
  - `ENABLE_GPIO=0` — disable GPIO (e.g. development on a PC); alarm still logs to console.
  - `GPIO_PIN=10` — stack light output, BCM (default 10).
  - `STACK_LIGHT_ACTIVE_LOW=1` — set if your relay module turns **on** when GPIO is **LOW**.
  - `GPIO_RESET_PIN=9` — reset button input, BCM (default 9). Set `ENABLE_GPIO_RESET=0` to disable the button watcher only.

**Permissions**: add the user running Electron to the `gpio` group, then reboot:  
`sudo usermod -a -G gpio $USER`

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
