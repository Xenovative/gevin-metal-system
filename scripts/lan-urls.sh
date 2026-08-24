#!/usr/bin/env bash
# Print how iPad / PC browsers should open this Linux server.
# Source from other scripts after cd to the repo root.

print_access_urls() {
  local port="${PORT:-7861}"
  echo "  This Linux server:  http://127.0.0.1:${port}"
  local ip
  local ips=""
  if command -v hostname >/dev/null 2>&1; then
    ips="$(hostname -I 2>/dev/null || true)"
  fi
  if [[ -z "${ips// }" ]] && command -v ip >/dev/null 2>&1; then
    ips="$(ip -4 -o addr show scope global 2>/dev/null | awk '{print $4}' | cut -d/ -f1 | tr '\n' ' ')"
  fi
  for ip in $ips; do
    case "$ip" in
      127.*|::1|fe80:*|"") continue ;;
    esac
    echo "  iPad / PC / phone:  http://${ip}:${port}"
  done
  echo "  Shared database:    ${PWD}/data/gevin.db"
  echo "  Login:              admin / admin123"
  echo "  Staff accounts:     Admin → 🔐 Admin 管理 → 建立員工帳號"
  echo "  iPad and other PCs: browser only — do not install a second copy."
}
