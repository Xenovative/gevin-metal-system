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

## Linux 一鍵啟動（建議）

在專案目錄執行（Ubuntu / Debian / 其他有 `python3` 的 Linux）：

```bash
bash scripts/run.sh
```

第一次執行會自動建立 `.venv` 並安裝依賴；之後同一指令即可啟動。預設埠 **7861**。

若系統尚未安裝 `python3-venv`（常見於全新 Ubuntu），先跑一次：

```bash
bash scripts/install-ubuntu.sh && bash scripts/run.sh
```

瀏覽器開啟：`http://127.0.0.1:7861`（遠端伺服器則用 `http://<伺服器IP>:7861`）

自訂埠號：

```bash
PORT=8080 bash scripts/run.sh
```

### 預設登入

- 帳號：`admin`
- 密碼：`admin123`

## 手動安裝（不用腳本）

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

systemd 範例：`deploy/gevin-metal.service`（安裝腳本會自動替換路徑佔位符）。

## 環境變數

| 變數 | 預設 | 說明 |
|------|------|------|
| `PORT` | `7861` | Gradio 服務埠 |
| `GRADIO_ANALYTICS_ENABLED` | `False` | 關閉 Gradio 遙測 |

## 目錄結構

```
gevin-metal-system/
├── app.py                  # Gradio 主程式
├── config.py               # 交易類型與選項設定
├── database.py             # SQLite 資料庫
├── invoice_generator.py    # Excel 發票生成
├── inventory.py            # 倉存邏輯
├── reports.py              # 報表生成
├── scripts/
│   ├── install-ubuntu.sh   # Ubuntu 系統套件 + venv 安裝
│   └── run.sh              # Linux 啟動（可自動建 venv）
├── deploy/
│   ├── install.sh          # VPS 一鍵安裝
│   └── gevin-metal.service # systemd 範例
├── templates/
│   └── invoice_template.xlsx
├── output/
│   ├── invoices/           # 生成的發票
│   └── reports/            # 生成的報表
└── data/
    └── gevin.db            # 資料庫檔案
```

## 相依套件

見 `requirements.txt`：

- gradio
- openpyxl
- sqlalchemy
- pandas

皆為跨平台 PyPI 套件，無 macOS / Windows 專用依賴。請使用 Linux 上的 `.venv`（勿複製其他系統的虛擬環境）。
