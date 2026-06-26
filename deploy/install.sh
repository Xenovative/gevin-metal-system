#!/usr/bin/env bash
# Gevin Metal System — Hostinger VPS install script
# Run as root on Ubuntu/Debian: bash install.sh

set -euo pipefail

APP_DIR="/var/www/gevin-metal-system"
APP_USER="gevin"
APP_PORT="${APP_PORT:-7861}"
REPO_URL="${REPO_URL:-https://github.com/Xenovative/gevin-metal-system.git}"
DOMAIN="${DOMAIN:-}"   # optional, e.g. gevin.yourdomain.com

echo "==> Installing system packages..."
export DEBIAN_FRONTEND=noninteractive
apt-get update -qq
apt-get install -y -qq git python3 python3-venv python3-pip nginx ufw curl

echo "==> Creating app user..."
if ! id "$APP_USER" &>/dev/null; then
  useradd --system --shell /bin/bash "$APP_USER"
fi

echo "==> Cloning or updating app..."
mkdir -p "$(dirname "$APP_DIR")"
if [ -d "$APP_DIR/.git" ]; then
  cd "$APP_DIR"
  git pull
else
  rm -rf "$APP_DIR"
  git clone "$REPO_URL" "$APP_DIR"
fi
chown -R "$APP_USER:$APP_USER" "$APP_DIR"

echo "==> Creating Python virtual environment..."
cd "$APP_DIR"
sudo -u "$APP_USER" python3 -m venv .venv
sudo -u "$APP_USER" .venv/bin/pip install --upgrade pip
sudo -u "$APP_USER" .venv/bin/pip install -r requirements.txt

echo "==> Creating data/output directories..."
sudo -u "$APP_USER" mkdir -p data output/invoices output/reports

echo "==> Installing systemd service..."
cp deploy/gevin-metal.service /etc/systemd/system/gevin-metal.service
sed -i "s|__APP_DIR__|$APP_DIR|g" /etc/systemd/system/gevin-metal.service
sed -i "s|__APP_PORT__|$APP_PORT|g" /etc/systemd/system/gevin-metal.service
sed -i "s|__APP_USER__|$APP_USER|g" /etc/systemd/system/gevin-metal.service

systemctl daemon-reload
systemctl enable gevin-metal
systemctl restart gevin-metal

echo "==> Configuring nginx..."
if [ -n "$DOMAIN" ]; then
  cp deploy/nginx-domain.conf /etc/nginx/sites-available/gevin-metal
  sed -i "s|__DOMAIN__|$DOMAIN|g" /etc/nginx/sites-available/gevin-metal
  sed -i "s|__APP_PORT__|$APP_PORT|g" /etc/nginx/sites-available/gevin-metal
else
  cp deploy/nginx-ip.conf /etc/nginx/sites-available/gevin-metal
  sed -i "s|__APP_PORT__|$APP_PORT|g" /etc/nginx/sites-available/gevin-metal
fi

ln -sf /etc/nginx/sites-available/gevin-metal /etc/nginx/sites-enabled/gevin-metal
rm -f /etc/nginx/sites-enabled/default
nginx -t
systemctl reload nginx

echo "==> Configuring firewall..."
ufw allow OpenSSH
ufw allow 80/tcp
ufw allow 443/tcp
ufw allow "${APP_PORT}/tcp"
ufw --force enable || true

echo ""
echo "=============================================="
echo "  Install complete!"
echo "=============================================="
systemctl status gevin-metal --no-pager -l || true
echo ""
if [ -n "$DOMAIN" ]; then
  echo "  App URL:  http://$DOMAIN"
  echo "  (Add SSL: certbot --nginx -d $DOMAIN)"
else
  echo "  App URL:  http://$(curl -s ifconfig.me 2>/dev/null || hostname -I | awk '{print $1}'):${APP_PORT}"
  echo "  Or via nginx on port 80: http://$(curl -s ifconfig.me 2>/dev/null || hostname -I | awk '{print $1}')"
fi
echo "  Login:    admin / admin123  (change after first login)"
echo "=============================================="
