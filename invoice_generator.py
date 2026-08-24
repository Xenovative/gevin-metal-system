"""
Invoice Excel facade: money/payment formatters + template fill entry point.

Invoices are generated from templates/invoice_template.xlsx (branded dual-copy
A4, 100% scale). This module keeps helpers used by app/reports.
"""
import json
from pathlib import Path

from config import DEFAULT_CASH_CURRENCY, GRAMS_PER_TAEL, OUTPUT_DIR

from invoice_excel_generator import (  # noqa: F401 — re-export public print API
    COMPANY_COPY,
    COMPANY_COPY_OFFSET,
    CUSTOMER_COPY,
    PAYMENT_LINE_ROWS,
    PRINT_AREA,
    estimate_item_rows,
    estimate_items_block_rows,
    fit_items_for_print,
    generate_invoice_excel,
)


def _parse_payment_entries(payment_method):
    if not payment_method:
        return []
    try:
        payments = json.loads(payment_method)
    except (json.JSONDecodeError, TypeError):
        return [{"method": str(payment_method), "amount": None}]
    if not isinstance(payments, list):
        return [{"method": str(payment_method), "amount": None}]
    return payments


def format_money(value, currency=None):
    """Format numeric amount with currency label, e.g. HKD 1,234.56."""
    if value is None or value == "":
        return None
    currency = currency or DEFAULT_CASH_CURRENCY
    num = float(value)
    if num == int(num):
        return f"{currency} {int(num):,}"
    return f"{currency} {num:,.2f}"


def format_cash_warehouse_number(value):
    """Signed Cash Warehouse number only (no currency), default 0."""
    num = 0.0 if value is None or value == "" else float(value)
    body = f"{int(num):,}" if float(num) == int(num) else f"{num:,.2f}"
    if num > 0:
        body = f"+{body}"
    return body


def format_cash_warehouse_amount(value, currency=None):
    """
    Excel Amount column = Cash Warehouse (現金倉).
    Signed: positive = 入倉, negative = 出倉. Default 0.
    """
    currency = currency or DEFAULT_CASH_CURRENCY
    return f"{currency} {format_cash_warehouse_number(value)}"


def format_payment_excel_lines(payment_method):
    """每種付款方式一行，供 Excel 垂直列出。"""
    lines = []
    for entry in _parse_payment_entries(payment_method):
        method = entry.get("method", "")
        amount = entry.get("amount")
        currency = entry.get("currency") or DEFAULT_CASH_CURRENCY
        if method and amount is not None:
            lines.append(f"{method} {format_money(amount, currency)}")
        elif method:
            lines.append(method)
    return lines


def format_payment_method_display(payment_method):
    """將付款資料（JSON 或舊版純文字）轉為可讀字串。"""
    lines = format_payment_excel_lines(payment_method)
    if lines:
        return " / ".join(lines)
    if not payment_method:
        return ""
    try:
        json.loads(payment_method)
        return ""
    except (json.JSONDecodeError, TypeError):
        return str(payment_method)


def resolve_invoice_excel_path(excel_path, invoice_no=None):
    """Resolve stored excel_path (relative or absolute) to an existing file."""
    candidates = []
    if excel_path:
        p = Path(str(excel_path).strip())
        candidates.append(p)
        candidates.append(OUTPUT_DIR / p.name)
        if not p.is_absolute():
            candidates.append(OUTPUT_DIR.parent / p)
            candidates.append(OUTPUT_DIR.parent.parent / p)
    if invoice_no:
        candidates.append(OUTPUT_DIR / f"{invoice_no}.xlsx")
    seen = set()
    for cand in candidates:
        try:
            key = str(cand)
            if key in seen:
                continue
            seen.add(key)
            if cand.exists():
                return str(cand.resolve())
        except OSError:
            continue
    return None


def compute_tael_from_gram(grams):
    if grams is None or grams == "":
        return None
    return round(float(grams) / GRAMS_PER_TAEL, 3)
