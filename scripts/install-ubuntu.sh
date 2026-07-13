#!/usr/bin/env bash
# Install system packages + Python venv for Ubuntu/Debian Linux.
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

# Drop macOS AppleDouble junk that breaks tools / imports on Linux.
find "$ROOT_DIR" -name '._*' -not -path '*/.git/*' -delete 2>/dev/null || true
find "$ROOT_DIR" -name '.DS_Store' -not -path '*/.git/*' -delete 2>/dev/null || true

echo "==> Installing Ubuntu/Debian system packages..."
if command -v apt-get >/dev/null 2>&1; then
  sudo apt-get update
  sudo apt-get install -y python3 python3-venv python3-pip
else
  echo "WARNING: apt-get not found. Ensure python3 + venv + pip are installed." >&2
fi

if ! command -v python3 >/dev/null 2>&1; then
  echo "ERROR: python3 is required." >&2
  exit 1
fi

echo "==> Creating Python virtual environment at $ROOT_DIR/.venv ..."
python3 -m venv .venv

echo "==> Installing Python dependencies..."
.venv/bin/pip install --upgrade pip
.venv/bin/pip install -r requirements.txt

echo "==> Ensuring runtime directories exist..."
mkdir -p data output/invoices output/reports templates logs

if [[ ! -f templates/invoice_template.xlsx ]]; then
  echo "WARNING: templates/invoice_template.xlsx is missing."
  echo "Copy the Excel template into templates/ before generating invoices."
fi

chmod +x scripts/install-ubuntu.sh scripts/run.sh 2>/dev/null || true

echo ""
echo "Install complete."
echo "Start the app with:"
echo "  cd $ROOT_DIR && ./scripts/run.sh"
echo ""
echo "Or one-liner:"
echo "  cd $ROOT_DIR && bash scripts/run.sh"
