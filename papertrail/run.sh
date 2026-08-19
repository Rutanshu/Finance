#!/usr/bin/env bash
# PaperTrail launcher (macOS / Linux)
set -e
cd "$(dirname "$0")"
command -v python3 >/dev/null || { echo "Python 3 is required."; exit 1; }
python3 -c "import openpyxl" 2>/dev/null || python3 -m pip install -r requirements.txt
exec python3 -m papertrail "$@"
