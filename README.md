# 貴金屬加工廠 — 發票與倉存系統

為貴金屬加工廠銷售人員設計的逐步填寫系統，自動生成 Excel 發票並記錄金屬進出倉。

## 功能

### 發票系統
- 逐步引導填寫：交易性質 → 基本資料 → 貨品 → 備註 → 生成 Excel
- **單號自動產生**：`前綴 + 年份(2位) + 月份(2位) + 流水號(5位)`
  - 銷售 `S`、購入 `P`、兌料 `T`、交收 `D`
  - 例：2026年3月第一張銷售單 → `S260300001`，每月由 `00001` 重新計數
  - 同一前綴共用流水號（所有交收類型共用 D、所有兌料類型共用 T）
- 支援多種交易性質，對應 Excel 範本中的不同分頁
- 兌料類交易支援「對換貨品」區塊
- 重量單位：克（必填）、両（可選）、安士 oz（可選）

### 倉存系統
- 每張發票自動記錄金屬入倉/出倉
- 即時查看庫存結存與進出倉明細

### 報表
- 每日 / 每月 / 每年倉存報表與發票報表（Excel）

## Docker 部署（建議 — 客戶 Linux 伺服器 / LAN）

適用於要給區網多台裝置（iPad / 手機 / PC）使用的 Linux 主機。環境固定，較少 Python/venv 問題。

### 一鍵（需已安裝 Docker）

```bash
sudo apt-get update && sudo apt-get install -y git docker.io docker-compose-v2
sudo usermod -aG docker "$USER"   # 登出再登入後生效；或暫時用 sudo
git clone https://github.com/Xenovative/gevin-metal-system.git
cd gevin-metal-system
bash scripts/docker-run.sh
```

或手動：

```bash
cd gevin-metal-system
mkdir -p data output/invoices output/reports logs
docker compose up -d --build
```

瀏覽器：

- 本機：`http://127.0.0.1:7861`
- 區網：`http://<伺服器IP>:7861`

常用指令：

```bash
docker compose logs -f      # 看日誌
docker compose restart      # 重啟
docker compose down         # 停止
docker compose up -d --build   # 更新程式後重建
```

`data/`、`output/`、`logs/`、`templates/` 會掛載到主機，資料庫與發票不會因重建容器而遺失。

防火牆若有開：`sudo ufw allow 7861/tcp`

### 預設登入

- 帳號：`admin`
- 密碼：`admin123`

## Linux 一鍵部署（不用 Docker / venv）

```bash
sudo apt-get update && sudo apt-get install -y git && git clone https://github.com/Xenovative/gevin-metal-system.git && cd gevin-metal-system && bash scripts/install-ubuntu.sh && bash scripts/run.sh
```

若專案已在本機：

```bash
bash scripts/run.sh
```

第一次會自動建立 `.venv` 並安裝依賴。預設埠 **7861**。

瀏覽器：`http://127.0.0.1:7861` 或 `http://<伺服器IP>:7861`

自訂埠：`PORT=8080 bash scripts/run.sh`

除錯模式（詳細日誌）：`GEVIN_DEBUG=1 bash scripts/run.sh`

## 手動安裝

```bash
sudo apt-get update && sudo apt-get install -y python3 python3-venv python3-pip
cd /path/to/gevin-metal-system
python3 -m venv .venv
.venv/bin/pip install --upgrade pip
.venv/bin/pip install -r requirements.txt
PORT=7861 .venv/bin/python app.py
```

## VPS 部署（Hostinger 等）

見 `deploy/HOSTINGER_DEPLOY.md`，或：

```bash
cd /var/www/gevin-metal-system && bash deploy/install.sh
```

## 環境變數

| 變數 | 預設 | 說明 |
|------|------|------|
| `PORT` | `7861` | Gradio 服務埠 |
| `GRADIO_ANALYTICS_ENABLED` | `False` | 關閉 Gradio 遙測 |

## 目錄結構

```
gevin-metal-system/
├── app.py
├── Dockerfile
├── docker-compose.yml
├── scripts/
│   ├── docker-run.sh
│   ├── install-ubuntu.sh
│   └── run.sh
├── deploy/
├── templates/
│   └── invoice_template.xlsx
├── output/
└── data/
```

## 相依套件

見 `requirements.txt`（gradio / openpyxl / sqlalchemy / pandas）。請在 Linux 上建立新的 `.venv`，勿複製其他系統的虛擬環境。
