#!/usr/bin/env bash
# Fix Electron SIGBUS on macOS: clean reinstall of Electron
set -e
cd "$(dirname "$0")/.."
echo "Removing Electron and clearing cache..."
rm -rf node_modules/electron
npm cache clean --force
echo "Reinstalling Electron..."
npm install electron@27.3.0 --save-dev
echo "Done. Run: npm start"
