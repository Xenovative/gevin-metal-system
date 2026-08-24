#!/usr/bin/env bash
# Pull the latest code from GitHub and restart the Linux server.
# Keeps data/ (SQLite), output/, and logs/ — staff accounts and invoices stay.
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"
# shellcheck source=lan-urls.sh
source "$ROOT_DIR/scripts/lan-urls.sh"

if [[ ! -d .git ]]; then
  echo "ERROR: this folder is not a git clone. On the client server:" >&2
  echo "  git clone https://github.com/Xenovative/gevin-metal-system.git" >&2
  exit 1
fi

echo "==> Updating from GitHub (data/ and output/ are not overwritten)..."
git fetch origin
git pull --ff-only origin main

mkdir -p data output/invoices output/reports logs templates

if command -v docker >/dev/null 2>&1 && docker compose ps -q gevin-metal 2>/dev/null | grep -q .; then
  echo "==> Rebuilding Docker container..."
  docker compose up -d --build
elif command -v systemctl >/dev/null 2>&1 && systemctl is-active --quiet gevin-metal 2>/dev/null; then
  echo "==> Refreshing Python deps and restarting systemd service..."
  if [[ -x .venv/bin/pip ]]; then
    .venv/bin/pip install -r requirements.txt
  fi
  sudo systemctl restart gevin-metal
elif [[ -x .venv/bin/pip ]]; then
  echo "==> Refreshing Python deps..."
  .venv/bin/pip install -r requirements.txt
  echo "Restart the running app (Ctrl+C the old process, then): bash scripts/run.sh"
else
  echo "No Docker/systemd app detected. Start with: bash scripts/docker-run.sh"
  echo "  or: bash scripts/run.sh"
fi

echo ""
echo "Update complete. Same shared database for every device:"
print_access_urls
