#!/usr/bin/env bash
# Start gevin-metal-system on Linux (creates .venv on first run if needed).
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

# Drop macOS AppleDouble junk that can confuse Python tooling.
find "$ROOT_DIR" -maxdepth 2 -name '._*' -not -path '*/.git/*' -delete 2>/dev/null || true

mkdir -p data output/invoices output/reports templates logs

if [[ ! -x .venv/bin/python ]]; then
  echo "Virtual environment missing — bootstrapping..."
  if ! command -v python3 >/dev/null 2>&1; then
    echo "ERROR: python3 not found. Run: bash scripts/install-ubuntu.sh" >&2
    exit 1
  fi
  python3 -m venv .venv
  .venv/bin/pip install --upgrade pip
  .venv/bin/pip install -r requirements.txt
fi

PORT="${PORT:-7861}"
export PORT
export GRADIO_ANALYTICS_ENABLED="${GRADIO_ANALYTICS_ENABLED:-False}"
export GRADIO_SERVER_NAME="${GRADIO_SERVER_NAME:-0.0.0.0}"
export GRADIO_NODE_SERVER_NAME="${GRADIO_NODE_SERVER_NAME:-0.0.0.0}"
export PYTHONUNBUFFERED=1

echo "Starting gevin-metal-system on http://0.0.0.0:${PORT}"
echo "Open from this machine: http://127.0.0.1:${PORT}"
echo "Open from iPad/phone/PC on LAN: http://<server-ip>:${PORT}"
exec .venv/bin/python app.py
