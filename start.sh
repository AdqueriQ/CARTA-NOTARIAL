#!/bin/bash
set -e

# Make uno module (installed via python3-uno apt package) visible to Python 3.11
export PYTHONPATH="/usr/lib/python3/dist-packages:${PYTHONPATH}"

# Start unoserver — keeps LibreOffice loaded in memory between conversions
echo "[start.sh] Iniciando unoserver..."
unoserver &

# Wait for LibreOffice to fully load (~4-5 seconds on first start)
sleep 6
echo "[start.sh] unoserver listo"

# Start Flask app
exec python app.py
