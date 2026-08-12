#!/usr/bin/env bash
# Pre-deploy checks for a clean Linux / Docker client install.
# Run from repo root: bash scripts/verify_linux_ready.sh
# Or: python3 scripts/verify_linux_ready.py
set -euo pipefail
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"
if [[ -x .venv/bin/python ]]; then
  exec .venv/bin/python scripts/verify_linux_ready.py
fi
exec python3 scripts/verify_linux_ready.py
