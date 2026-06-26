import re
from datetime import date, datetime

from config import TRANSACTION_TYPES


def parse_tx_date(tx_date) -> date:
    """解析 Gradio 日期元件回傳值（str / datetime / date / float 時間戳）。"""
    if tx_date is None or tx_date == "":
        return date.today()
    if isinstance(tx_date, date) and not isinstance(tx_date, datetime):
        return tx_date
    if isinstance(tx_date, datetime):
        return tx_date.date()
    if isinstance(tx_date, (int, float)):
        return datetime.fromtimestamp(tx_date).date()
    if isinstance(tx_date, str):
        for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d"):
            try:
                return datetime.strptime(tx_date[:19], fmt).date()
            except ValueError:
                continue
    return date.today()

# 格式：{前綴}{YY}{MM}{NNNNN}，例 S260300001
INVOICE_NO_PATTERN = re.compile(r"^([SPTD])(\d{2})(\d{2})(\d{5})$")
SEQUENCE_WIDTH = 5


def format_invoice_number(prefix: str, year: int, month: int, sequence: int) -> str:
    yy = year % 100
    return f"{prefix}{yy:02d}{month:02d}{sequence:0{SEQUENCE_WIDTH}d}"


def parse_invoice_number(invoice_no: str):
    match = INVOICE_NO_PATTERN.match(invoice_no.strip())
    if not match:
        return None
    prefix, yy, mm, seq = match.groups()
    return {
        "prefix": prefix,
        "year": 2000 + int(yy),
        "month": int(mm),
        "sequence": int(seq),
    }


def get_prefix(transaction_type: str) -> str:
    return TRANSACTION_TYPES[transaction_type]["prefix"]


def get_next_invoice_number(session, transaction_type: str, tx_date) -> str:
    """依前綴與交易日期，產生當月下一個單號。"""
    tx_date = parse_tx_date(tx_date)
    prefix = get_prefix(transaction_type)

    from database import Invoice

    invoices = session.query(Invoice.invoice_no).all()
    max_seq = 0
    for (invoice_no,) in invoices:
        parsed = parse_invoice_number(invoice_no)
        if parsed and parsed["prefix"] == prefix:
            if parsed["year"] == tx_date.year and parsed["month"] == tx_date.month:
                max_seq = max(max_seq, parsed["sequence"])

    return format_invoice_number(prefix, tx_date.year, tx_date.month, max_seq + 1)


def validate_invoice_number(invoice_no: str, transaction_type: str, tx_date) -> str | None:
    """驗證單號格式是否正確，回傳錯誤訊息或 None。"""
    parsed = parse_invoice_number(invoice_no)
    if not parsed:
        return "單號格式錯誤，應為：前綴 + 年份(2位) + 月份(2位) + 流水號(5位)，例 S260300001"

    expected_prefix = get_prefix(transaction_type)
    if parsed["prefix"] != expected_prefix:
        return f"單號前綴應為 {expected_prefix}（{transaction_type}）"

    tx_date = parse_tx_date(tx_date)

    if parsed["year"] != tx_date.year or parsed["month"] != tx_date.month:
        return (
            f"單號年月應與交易日期一致"
            f"（{tx_date.year % 100:02d}{tx_date.month:02d}）"
        )

    return None
