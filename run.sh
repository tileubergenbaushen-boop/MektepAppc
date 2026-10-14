#!/usr/bin/env bash
# MektepApp AI - one-click launcher for macOS / Linux
set -e
cd "$(dirname "$0")"

if [ ! -x .venv/bin/python ]; then
  echo "[1/3] Creating virtual environment .venv ..."
  python3 -m venv .venv
fi

echo "[2/3] Installing dependencies - the first run takes a few minutes ..."
.venv/bin/python -m pip install --disable-pip-version-check -q -r requirements.txt

echo "[3/3] Starting MektepApp AI - the browser will open automatically."
exec .venv/bin/python -m streamlit run app.py
