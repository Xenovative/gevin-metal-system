from pathlib import Path

# Resolve to absolute paths so the app works regardless of cwd (systemd, nohup, etc.)
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates" / "invoice_template.xlsx"
OUTPUT_DIR = BASE_DIR / "output" / "invoices"
REPORT_DIR = BASE_DIR / "output" / "reports"
DB_PATH = BASE_DIR / "data" / "gevin.db"
LOG_DIR = BASE_DIR / "logs"


def ensure_runtime_dirs() -> None:
    """Create data/output directories on Linux/macOS/Windows before first use."""
    for path in (OUTPUT_DIR, REPORT_DIR, DB_PATH.parent, LOG_DIR, TEMPLATE_PATH.parent):
        path.mkdir(parents=True, exist_ok=True)

ITEM_TYPES = [
    "足金 Pure Gold",
    "飾金 Gold Products",
    "雜金 Gold Accessories",
    "K金 Karat Gold",
    "純銀 Silver",
    "其他 Other",
]

QUALITY_OPTIONS = ["足金", "24K", "24k", "18K", "9.997", "其他", "按資料"]

UNITS = ["克 Gram", "両 Teal"]

PAYMENT_METHODS = ["現金 Cash", "轉帳 Transfer", "支票 Cheque", "其他 Other"]
CASH_PAYMENT_METHOD = "現金 Cash"
CASH_CURRENCIES = ["HKD$", "USD$", "CNY¥", "EUR€", "GBP£", "JPY¥", "TWD$", "MOP$", "SGD$"]
DEFAULT_CASH_CURRENCY = "HKD$"

ROLE_ADMIN = "admin"
ROLE_STAFF = "staff"
INVOICE_STATUS_ACTIVE = "active"
INVOICE_STATUS_VOIDED = "voided"

ROLE_PERMISSIONS = {
    ROLE_ADMIN: {
        "create_invoice": True,
        "view_inventory": True,
        "download_reports": True,
        "void_invoice": True,
        "manage_users": True,
        "view_audit": True,
    },
    ROLE_STAFF: {
        "create_invoice": True,
        "view_inventory": True,
        "download_reports": True,
        "void_invoice": False,
        "manage_users": False,
        "view_audit": False,
    },
}

WAREHOUSES = ["A倉庫", "B倉庫", "C倉庫", "現金倉"]
METAL_WAREHOUSES = ["A倉庫", "B倉庫", "C倉庫"]
CASH_WAREHOUSE = "現金倉"
STORAGE_ACTIONS = ["存", "取"]
STORAGE_LOCATION_CHOICES = [
    f"{action} {warehouse}"
    for action in STORAGE_ACTIONS
    for warehouse in WAREHOUSES
]

# 舊版保險箱名稱對照（讀取歷史資料用）
WAREHOUSE_ALIASES = {
    "A保險箱": "A倉庫",
    "B保險箱": "B倉庫",
    "C保險箱": "C倉庫",
    "客戶": "客戶",
}

# 現金收入／支出交易類型
CASH_IN_TRANSACTION_TYPES = {
    "銷售", "兌料",
}
CASH_OUT_TRANSACTION_TYPES = {
    "購入", "交收去料",
}

SAFE_LOCATIONS = STORAGE_LOCATION_CHOICES
SAFE_BOXES = METAL_WAREHOUSES
SAFE_SUMMARY_CATEGORIES = ["金", "純銀"]

# 交易性質設定
# customer_notes_col: 客戶單備註欄（D=4, E=5）；公司單統一用 E 欄
# inventory_direction: in=入倉 / out=出倉 / exchange=兌換（來料入、對換出）
TRANSACTION_TYPES = {
    "銷售": {
        "sheet": "銷售(保護)",
        "prefix": "S",
        "number_label": "Invoice No.",
        "inventory_direction": "out",
        "has_exchange": False,
        "has_amount": True,
        "customer_notes_col": 4,
        "description": "銷售金屬給客戶，金屬出倉",
    },
    "購入": {
        "sheet": "購入單(保護)",
        "prefix": "P",
        "number_label": "Invoice No.",
        "inventory_direction": "in",
        "has_exchange": False,
        "has_amount": True,
        "customer_notes_col": 5,
        "description": "向供應商購入金屬，金屬入倉",
    },
    "兌料": {
        "sheet": "兌料單(保護)",
        "prefix": "T",
        "number_label": "Invoice No.",
        "inventory_direction": "exchange",
        "has_exchange": True,
        "has_amount": True,
        "customer_notes_col": 4,
        "description": "客戶來料兌換新貨，需填寫「對換貨品」",
    },
    "交收去料": {
        "sheet": "交收單(保護)去料",
        "prefix": "D",
        "number_label": "編號 No.",
        "inventory_direction": "out",
        "has_exchange": False,
        "has_amount": False,
        "customer_notes_col": 5,
        "description": "金屬送去提純等，出倉記錄",
    },
}

GRAMS_PER_TAEL = 37.5
