# Deploy Gevin Metal System on Hostinger VPS

Server IP: **151.106.125.6**

After deployment you can access the app at:

| Method | URL |
|--------|-----|
| **Direct port** (simplest) | `http://151.106.125.6:7861` |
| **Via nginx** (port 80) | `http://151.106.125.6` |
| **Custom subdomain** (recommended) | `http://gevin.yourdomain.com` |

Default login: `admin` / `admin123` — change the password after first login.

---

## What you need before starting

1. **Hostinger VPS** with IP `151.106.125.6`
2. **SSH access** (root or sudo user) — find credentials in Hostinger hPanel → VPS → SSH Access
3. **GitHub access** to the private repo `Xenovative/gevin-metal-system`
4. **(Optional)** A domain pointed to `151.106.125.6` for a clean link

---

## Step 1 — Connect to your server

From your computer (PowerShell or Terminal):

```bash
ssh root@151.106.125.6
```

If Hostinger uses a custom SSH port (e.g. 65002), use:

```bash
ssh root@151.106.125.6 -p 65002
```

---

## Step 2 — Open firewall ports in Hostinger

In **Hostinger hPanel**:

1. Go to **VPS** → your server → **Firewall**
2. Allow these inbound ports:
   - `22` (SSH)
   - `80` (HTTP — nginx)
   - `443` (HTTPS — optional, for SSL later)
   - `7861` (direct app access)

---

## Step 3 — Clone the repo (private)

Because the repo is private, authenticate first.

### Option A — GitHub personal access token (easiest)

1. GitHub → Settings → Developer settings → Personal access tokens → Generate token
2. Enable scope: `repo`
3. On the server:

```bash
apt update && apt install -y git
git clone https://github.com/Xenovative/gevin-metal-system.git /var/www/gevin-metal-system
# Username: Xenovative
# Password: paste your token (not your GitHub password)
```

### Option B — SSH deploy key

```bash
ssh-keygen -t ed25519 -C "hostinger-gevin" -f ~/.ssh/gevin_deploy -N ""
cat ~/.ssh/gevin_deploy.pub
```

Add the public key in GitHub → repo → Settings → Deploy keys → Add deploy key.

```bash
GIT_SSH_COMMAND='ssh -i ~/.ssh/gevin_deploy' git clone git@github.com:Xenovative/gevin-metal-system.git /var/www/gevin-metal-system
```

---

## Step 4 — Run the install script

```bash
cd /var/www/gevin-metal-system
chmod +x deploy/install.sh
bash deploy/install.sh
```

### With a custom subdomain (recommended)

First point your domain A record to `151.106.125.6` in Hostinger DNS, then:

```bash
DOMAIN=gevin.yourdomain.com bash deploy/install.sh
```

### With a custom port

```bash
APP_PORT=7861 bash deploy/install.sh
```

---

## Step 5 — Verify it is running

```bash
systemctl status gevin-metal
curl -I http://127.0.0.1:7861
curl -I http://127.0.0.1
```

Open in your browser:

- **http://151.106.125.6:7861** — direct access
- **http://151.106.125.6** — via nginx on port 80

---

## Step 6 — (Optional) Add HTTPS / SSL

If you set up a domain (e.g. `gevin.yourdomain.com`):

```bash
apt install -y certbot python3-certbot-nginx
certbot --nginx -d gevin.yourdomain.com
```

Your link becomes: **https://gevin.yourdomain.com**

---

## Useful commands after deployment

| Action | Command |
|--------|---------|
| View logs | `journalctl -u gevin-metal -f` |
| Restart app | `systemctl restart gevin-metal` |
| Stop app | `systemctl stop gevin-metal` |
| Update from GitHub | `cd /var/www/gevin-metal-system && git pull && systemctl restart gevin-metal` |
| Check nginx | `nginx -t && systemctl status nginx` |

---

## Manual install (without script)

If you prefer to run commands yourself:

```bash
# 1. System packages
apt update && apt install -y git python3 python3-venv nginx ufw

# 2. Clone repo (see Step 3 for auth)
git clone https://github.com/Xenovative/gevin-metal-system.git /var/www/gevin-metal-system
cd /var/www/gevin-metal-system

# 3. Python environment
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
mkdir -p data output/invoices output/reports

# 4. Test run (Ctrl+C to stop)
PORT=7861 .venv/bin/python app.py

# 5. Install as service + nginx (use deploy/ files)
bash deploy/install.sh
```

---

## Troubleshooting

**Cannot connect from browser**
- Check Hostinger firewall allows ports 80 and 7861
- Run `ufw status` on server
- Run `systemctl status gevin-metal`

**502 Bad Gateway**
- App not running: `systemctl restart gevin-metal`
- Check port: `ss -tlnp | grep 7861`

**Private repo clone fails**
- Use a GitHub token with `repo` scope, not your password

**Change default admin password**
- Log in as admin → Admin tab → update user profile
