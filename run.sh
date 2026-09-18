#!/usr/bin/env bash
# Launcher for macOS / Linux. Creates a venv on first run, then starts the app.
set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 was not found. Install it from https://www.python.org/downloads/ and run this again."
  exit 1
fi

if [ ! -x ".venv/bin/python" ]; then
  echo "Creating virtual environment..."
  python3 -m venv .venv
fi

echo "Installing / checking dependencies..."
.venv/bin/python -m pip install --quiet --upgrade pip
.venv/bin/python -m pip install --quiet -r requirements.txt

echo
echo "Starting TechWitch Zendesk Helper. Press Ctrl+C to stop it."
exec .venv/bin/python app.py
