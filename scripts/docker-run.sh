#!/usr/bin/env bash
# Build and start gevin-metal-system with Docker Compose (LAN-ready).
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

if ! command -v docker >/dev/null 2>&1; then
  echo "ERROR: docker is not installed. On Ubuntu: sudo apt-get install -y docker.io docker-compose-v2" >&2
  exit 1
fi

mkdir -p data output/invoices output/reports logs templates

if [[ ! -f templates/invoice_template.xlsx ]]; then
  echo "WARNING: templates/invoice_template.xlsx is missing."
fi

echo "==> Building and starting gevin-metal (port ${PORT:-7861})..."
docker compose up -d --build

echo ""
echo "Running."
echo "  This machine:  http://127.0.0.1:${PORT:-7861}"
echo "  LAN devices:   http://<server-ip>:${PORT:-7861}"
echo "  Login:         admin / admin123"
echo ""
echo "Logs:   docker compose logs -f"
echo "Stop:   docker compose down"
