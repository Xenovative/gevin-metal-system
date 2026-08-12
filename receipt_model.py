"""
Canonical receipt document shared by Excel fill verification and PDF rendering.

Zeros are printable (only None / blank string are missing). Invoice numbers are
never invented here — they come from Excel cells or invoice_data.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional

import openpyxl

# Keep in sync with invoice_generator.CUSTOMER_COPY
CUSTOMER_COPY = {
    "invoice_no_row": 4,
    "info_row": 6,
    "items_start": 11,
    "notes_row": 19,
    "total_row": 22,
    "payment_row": 23,
}
COMPANY_COPY_OFFSET = 27
CUSTOMER_HANDLER = "XXXX"


def _is_blank(val: Any) -> bool:
    if val is None:
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    return False


def _strip_invoice_no(val: Any) -> str:
    if _is_blank(val):
        return ""
    text = str(val).strip()
    for prefix in ("Invoice No.", "Invoice No", "編號 No.", "編號 No"):
        if text.startswith(prefix):
            return text[len(prefix) :].strip()
    return text


def _fmt_date(val: Any) -> str:
    if _is_blank(val):
        return ""
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, date):
        return val.isoformat()
    text = str(val).strip()
    return text[:10] if len(text) >= 10 and text[4] == "-" else text


def _normalize_unit(unit: Any) -> str:
    if _is_blank(unit):
        return ""
    text = str(unit).strip()
    text = text.replace("Teal", "Tael")
    return text


def _strip_stock_label(text: Any) -> str:
    """Remove legacy Excel prefixes 倉存存取 / 倉存位置; keep action+location only."""
    if _is_blank(text):
        return ""
    t = str(text).strip()
    for prefix in ("倉存存取", "倉存位置"):
        if t.startswith(prefix):
            t = t[len(prefix) :].strip()
            break
    return t


def _fmt_weight_line(weight: Any, unit: Any) -> str:
    if _is_blank(weight) and _is_blank(unit):
        return ""
    w = weight
    if isinstance(w, float) and w == int(w):
        w = int(w)
    unit_s = _normalize_unit(unit)
    if _is_blank(w):
        return unit_s
    if unit_s:
        return f"{w} {unit_s}".strip()
    return str(w)


def _fmt_money(currency: Any, amount: Any) -> str:
    """Format as 'HKD$ 0.00' — numeric 0 is valid."""
    curr = "" if _is_blank(currency) else str(currency).strip()
    if _is_blank(amount):
        return curr
    try:
        num = float(amount)
        money = f"{num:,.2f}"
    except (TypeError, ValueError):
        money = str(amount).strip()
    if curr:
        return f"{curr} {money}".strip()
    return money


def _fmt_price(val: Any) -> str:
    if _is_blank(val):
        return ""
    try:
        num = float(val)
        if num == int(num):
            return str(int(num))
        return f"{num:g}"
    except (TypeError, ValueError):
        return str(val).strip()


@dataclass
class ReceiptLine:
    item_type: str = ""
    quality: str = ""
    weight_lines: list[str] = field(default_factory=list)  # e.g. "100 克 Gram"
    unit_price: str = ""
    currency: str = ""
    amount: Any = None  # raw numeric for zeros
    amount_display: str = ""  # "HKD$ 0.00"
    stock_action: str = ""  # e.g. 取 A倉庫
    stock_location: str = ""  # e.g. 存 客戶


@dataclass
class ReceiptCopy:
    kind: str  # "customer" | "company"
    invoice_no: str = ""
    date: str = ""
    customer: str = ""
    phone: str = ""
    lines: list[ReceiptLine] = field(default_factory=list)
    notes: str = ""
    note_amount_display: str = ""
    total_display: str = ""
    handler: str = ""
    payment_lines: list[str] = field(default_factory=list)

    @property
    def footer_label(self) -> str:
        if self.kind == "company":
            return "(公司單 Company Copy)"
        return "(客戶單 Customer Copy)"


@dataclass
class ReceiptDocument:
    invoice_no: str
    customer: ReceiptCopy
    company: ReceiptCopy


def _read_item_lines(ws, items_start: int, notes_row: int, is_company: bool) -> list[ReceiptLine]:
    lines: list[ReceiptLine] = []
    r = items_start
    while r < notes_row:
        item = ws.cell(r, 3).value
        item_s = "" if _is_blank(item) else str(item).strip()
        if item_s and "對換" in item_s and _is_blank(ws.cell(r, 5).value):
            r += 1
            continue

        quality = ws.cell(r, 5).value
        weight = ws.cell(r, 6).value
        unit = ws.cell(r, 7).value
        price = ws.cell(r, 9).value
        curr = ws.cell(r, 10).value
        amt = ws.cell(r, 11).value
        stock = ws.cell(r, 8).value if is_company else None

        has_any = not all(
            _is_blank(x) for x in (item, quality, weight, unit, price, curr, amt, stock)
        )
        if not has_any:
            r += 1
            continue

        # Start of a new item block when col C has item name, or continue weights
        if not _is_blank(item) or not _is_blank(quality) or (
            not _is_blank(price) or not _is_blank(amt)
        ):
            line = ReceiptLine(
                item_type=item_s,
                quality="" if _is_blank(quality) else str(quality).strip(),
                unit_price=_fmt_price(price),
                currency="" if _is_blank(curr) else str(curr).strip(),
                amount=amt if not _is_blank(amt) else None,
                amount_display=_fmt_money(curr, amt) if not _is_blank(amt) or not _is_blank(curr) else "",
            )
            wl = _fmt_weight_line(weight, unit)
            if wl:
                line.weight_lines.append(wl)
            if is_company and not _is_blank(stock):
                st = _strip_stock_label(stock)
                if st.startswith("存 ") or (st.startswith("存") and "取" not in st[:2]):
                    line.stock_location = st
                else:
                    line.stock_action = st

            # Peek following rows for tael/oz / stock location without new item name
            peek = r + 1
            while peek < notes_row:
                p_item = ws.cell(peek, 3).value
                p_qual = ws.cell(peek, 5).value
                p_w = ws.cell(peek, 6).value
                p_u = ws.cell(peek, 7).value
                p_price = ws.cell(peek, 9).value
                p_amt = ws.cell(peek, 11).value
                p_stock = ws.cell(peek, 8).value if is_company else None
                new_item = not _is_blank(p_item) and "對換" not in str(p_item)
                if new_item:
                    break
                if not _is_blank(p_qual) and _is_blank(p_item):
                    # unlikely; treat as new
                    break
                if all(_is_blank(x) for x in (p_w, p_u, p_stock, p_price, p_amt)):
                    break
                # weight continuation
                if not _is_blank(p_w) or not _is_blank(p_u):
                    wl2 = _fmt_weight_line(p_w, p_u)
                    if wl2:
                        line.weight_lines.append(wl2)
                if is_company and not _is_blank(p_stock):
                    st = _strip_stock_label(p_stock)
                    if st.startswith("存 ") or (st.startswith("存") and not line.stock_location):
                        line.stock_location = st
                    elif not line.stock_action:
                        line.stock_action = st
                peek += 1
                # stop after oz row typically; allow up to 2 continuations
                if peek > r + 3:
                    break
            lines.append(line)
            r = peek
            continue
        r += 1
    return lines


def _read_copy(ws, base: dict, kind: str) -> ReceiptCopy:
    offset = 0 if kind == "customer" else COMPANY_COPY_OFFSET
    inv_row = base["invoice_no_row"] + offset
    info_row = base["info_row"] + offset
    items_start = base["items_start"] + offset
    notes_row = base["notes_row"] + offset
    total_row = base["total_row"] + offset
    payment_row = base["payment_row"] + offset
    is_company = kind == "company"

    invoice_no = _strip_invoice_no(ws.cell(inv_row, 11).value)
    customer = "" if _is_blank(ws.cell(info_row, 5).value) else str(ws.cell(info_row, 5).value).strip()
    phone = ""
    if is_company:
        phone = "" if _is_blank(ws.cell(info_row, 7).value) else str(ws.cell(info_row, 7).value).strip()
    date_s = _fmt_date(ws.cell(info_row, 11).value)

    lines = _read_item_lines(ws, items_start, notes_row, is_company)

    notes = ""
    for ncol in (4, 5):
        nv = ws.cell(notes_row, ncol).value
        if not _is_blank(nv):
            t = str(nv).strip()
            if "備註" not in t:
                notes = t
                break
    note_amt = _fmt_money(ws.cell(notes_row, 10).value, ws.cell(notes_row, 11).value)
    total = _fmt_money(ws.cell(total_row, 10).value, ws.cell(total_row, 11).value)

    handler_raw = ws.cell(payment_row, 6).value
    handler = "" if _is_blank(handler_raw) else str(handler_raw).strip()
    if kind == "customer":
        # Canonical: customer copy always shows XXXX
        handler = CUSTOMER_HANDLER
    elif not handler:
        handler = "Admin"

    payments: list[str] = []
    for i in range(4):
        pv = ws.cell(payment_row + i, 3).value
        if _is_blank(pv):
            continue
        t = str(pv).strip()
        if "付款" in t or "Payment" in t:
            continue
        payments.append(t)

    return ReceiptCopy(
        kind=kind,
        invoice_no=invoice_no,
        date=date_s,
        customer=customer,
        phone=phone,
        lines=lines,
        notes=notes,
        note_amount_display=note_amt,
        total_display=total,
        handler=handler,
        payment_lines=payments,
    )


def from_excel(excel_path) -> ReceiptDocument:
    path = Path(excel_path)
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb.active
    customer = _read_copy(ws, CUSTOMER_COPY, "customer")
    company = _read_copy(ws, CUSTOMER_COPY, "company")
    invoice_no = customer.invoice_no or company.invoice_no
    # Prefer company phone/customer if customer blank
    if not customer.customer:
        customer.customer = company.customer
    return ReceiptDocument(invoice_no=invoice_no, customer=customer, company=company)


def from_invoice_data(
    invoice_data: dict,
    main_items: list[dict],
    exchange_items: Optional[list[dict]] = None,
) -> ReceiptDocument:
    """Build the same document shape from the live Gradio/DB payload."""
    inv = _strip_invoice_no(invoice_data.get("invoice_no", ""))
    date_s = _fmt_date(invoice_data.get("transaction_date"))
    customer_name = (invoice_data.get("customer_name") or "").strip()
    phone = (invoice_data.get("customer_phone") or "").strip()
    currency = invoice_data.get("invoice_currency") or "HKD$"
    handler_co = (invoice_data.get("handler") or "").strip() or "Admin"

    def build_lines(with_stock: bool) -> list[ReceiptLine]:
        out: list[ReceiptLine] = []
        for item in main_items or []:
            weights = []
            if item.get("weight_gram") is not None:
                weights.append(_fmt_weight_line(item.get("weight_gram"), "克 Gram"))
            if item.get("weight_tael") is not None:
                weights.append(_fmt_weight_line(item.get("weight_tael"), "両 Tael"))
            if item.get("weight_oz") is not None:
                weights.append(_fmt_weight_line(item.get("weight_oz"), "安士 oz"))
            amt = item.get("amount")
            line = ReceiptLine(
                item_type=(item.get("item_type") or "").strip(),
                quality=(item.get("quality") or "").strip(),
                weight_lines=weights,
                unit_price=_fmt_price(item.get("unit_price")),
                currency=currency,
                amount=amt,
                amount_display=_fmt_money(currency, amt if amt is not None else 0),
            )
            if with_stock:
                src = (invoice_data.get("source_location") or "").strip()
                dst = (invoice_data.get("destination_location") or "").strip()
                if src:
                    line.stock_action = _strip_stock_label(src)
                if dst:
                    line.stock_location = _strip_stock_label(dst)
            out.append(line)
        return out

    total = invoice_data.get("cash_warehouse_amount")
    if total is None or total == "":
        total = invoice_data.get("total_amount")
    if total is None or total == "":
        total = 0
    total_disp = _fmt_money(currency, total)

    customer = ReceiptCopy(
        kind="customer",
        invoice_no=inv,
        date=date_s,
        customer=customer_name,
        phone="",
        lines=build_lines(False),
        notes=(invoice_data.get("notes") or "").strip(),
        total_display=total_disp,
        handler=CUSTOMER_HANDLER,
    )
    company = ReceiptCopy(
        kind="company",
        invoice_no=inv,
        date=date_s,
        customer=customer_name,
        phone=phone,
        lines=build_lines(True),
        notes=(invoice_data.get("notes") or "").strip(),
        total_display=total_disp,
        handler=handler_co,
    )
    return ReceiptDocument(invoice_no=inv, customer=customer, company=company)
