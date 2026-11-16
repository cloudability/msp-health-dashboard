#!/bin/bash
# macOS: double-click to refresh the dashboard (same pipeline as Update-Dashboard.bat on Windows).
cd "$(dirname "$0")" || exit 1
if ! command -v node >/dev/null 2>&1; then
  echo "ERROR: Node.js is not installed. Install it from https://nodejs.org"
  read -r -p "Press Enter to close..."
  exit 1
fi
node build/update.js
status=$?
[ $status -ne 0 ] && read -r -p "Update failed. Press Enter to close..."
exit $status
