#!/usr/bin/env bash
# Build and start gevin-metal-system with Docker Compose (LAN-ready).
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"
# shellcheck source=lan-urls.sh
source "$ROOT_DIR/scripts/lan-urls.sh"

if ! command -v docker >/dev/null 2>&1; then
  echo "ERROR: docker is not installed. On Ubuntu: sudo apt-get install -y docker.io docker-compose-v2" >&2
  exit 1
fi

mkdir -p data output/invoices output/reports logs templates

if [[ -f data/gevin.db ]]; then
  echo "NOTE: data/gevin.db already exists on this host — existing invoices will be kept."
  echo "      For a blank client install, remove data/ before first start (fresh clone has no DB)."
else
  echo "NOTE: no database yet — first start creates a clean SQLite with admin / admin123 only."
fi

if [[ ! -f templates/invoice_template.xlsx ]]; then
  echo "WARNING: templates/invoice_template.xlsx is missing."
fi

echo "==> Building and starting gevin-metal (port ${PORT:-7861})..."
docker compose up -d --build

PORT="${PORT:-7861}"
export PORT
echo ""
echo "Running. One Linux server, one database — open the URL on iPad / PC / phone:"
print_access_urls
echo ""
echo "Logs:   docker compose logs -f"
echo "Stop:   docker compose down"
echo "Update: bash scripts/update.sh"
