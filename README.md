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

**重要（客戶機乾淨安裝）：** GitHub 倉庫**不含**本機測試用的 SQLite／發票 mock 資料（`data/`、`output/` 已列入 `.gitignore` 與 `.dockerignore`）。在客戶 Linux 上 `git clone` 後第一次啟動會建立**空白**資料庫，只有預設 `admin` / `admin123`，不會帶入筆電上的假單據。

### 一鍵（需已安裝 Docker）

```bash
sudo apt-get update && sudo apt-get install -y git docker.io docker-compose-v2
sudo usermod -aG docker "$USER"   # 登出再登入後生效；或暫時用 sudo
git clone https://github.com/Xenovative/gevin-metal-system.git
cd gevin-metal-system
bash scripts/docker-run.sh
```

部署前可在伺服器上跑：

```bash
bash scripts/verify_linux_ready.sh
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

## A4 品牌紙列印檢查清單（操作員）

列印真相：**Excel → 預印 A4 品牌收據紙**（現場）。數位預覽／下載為 Perfect V2 PDF（由 Excel 儲存格產生）。

1. 在「發票覆核」選單號 → 預覽 → **下載 Excel**（系統會依資料庫重新生成，確保與畫面一致）。
2. 用 Excel / LibreOffice 開啟後選擇 **A4、直向、符合頁面（fit to page / 縮放至一頁）**。
3. 印表機設定：**實際大小 / 符合可列印區域**（關閉「適合邊距」以外的額外縮放），邊距維持範本預設。
4. 紙匣放入 **品牌 A4 收據紙**（上半客戶單、下半公司單；金滿堂／收據框已印好）。
5. 試印一張：單號、客戶、貨品、金額（J=貨幣、K=現金倉正負數）、備註／合計／付款應對齊框線；Logo／標題帶不可被 Excel 字蓋住。
6. 貨品過多時畫面會警告，Excel 只印得下的項次，以免蓋住備註／合計區。

自動檢查：
- `python scripts/print_alignment_test.py`（Excel 儲存格／A4）
- `python scripts/excel_pdf_parity_test.py`（Excel 儲存格＋Perfect V2 數位 PDF 內容一致）

### 數位 PDF（Perfect V2）vs 現場列印

- **數位 PDF**：以 Excel 儲存格為準（`receipt_model`），用 `assets/receipt_header.png` 作雙聯 A4 收據（客戶單／裁切線／公司單）。金額顯示 `HKD$ 0.00`；客戶單與公司單經手人均為實際經手人；庫存欄不併入貨品名稱。
- **現場列印**：仍用 Excel 印在預印品牌 A4 紙上（勿依賴數位 PDF 當紙本）。

若 Perfect V2 產生失敗，會退回 LibreOffice／Excel COM 把工作簿轉成 PDF（可選）。

## 相依套件

見 `requirements.txt`（gradio / openpyxl / sqlalchemy / pandas / reportlab / pypdf）。請在 Linux 上建立新的 `.venv`，勿複製其他系統的虛擬環境。
