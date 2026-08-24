"""
Fill templates/invoice_template.xlsx with live invoice data.

Copies the branded dual-copy workbook (銷售 / 購入單 / 兌料單 / 交收單),
clears sample cells, then writes this invoice into the matching sheet.
Print at 100% A4 — never Fit-to-Page.
"""
from __future__ import annotations

import logging
import shutil
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

from cash import signed_cash_warehouse_amount
from config import (
    DEFAULT_CASH_CURRENCY,
    EXTERNAL_PARTY_LOCATION,
    INVENTORY_ACTION_DEPOSIT,
    INVENTORY_ACTION_WITHDRAW,
    OUTPUT_DIR,
    TEMPLATE_PATH,
    TRANSACTION_TYPES,
)

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Geometry — S260800018v4.xlsx blueprint (客戶單 top / 公司單 bottom).
# Footer is not a uniform +27: notes/total/handler sit 3 rows higher on 公司單.
# ---------------------------------------------------------------------------
CUSTOMER_COPY = {
    "invoice_no_row": 3,   # J3 type, K3 serial
    "info_row": 5,         # E5 customer, G5 phone, K5 date
    "items_start": 10,     # C10 item / E quality / F–G weights / I price / J–K money
    "notes_row": 18,       # notes; do not let items cross this row
    "total_row": 21,       # 付款方式 / 合計 / checkboxes (height 27.75)
    "payment_row": 22,     # C22+ used payment lines only
    "handler_row": 24,     # F24 name; F23 = 經手人 Handler:
}
COMPANY_COPY = {
    "invoice_no_row": 30,  # J30 type, K30 serial
    "info_row": 32,
    "items_start": 37,     # taller 24.75 pt rows; ~5 item lines
    "notes_row": 42,
    "total_row": 45,
    "payment_row": 46,
    "handler_row": 48,     # F48 name; F47 = 經手人 Handler:
}
# Header/items still +27 (3→30, 10→37); notes/payment/handler are +24.
COMPANY_COPY_OFFSET = 27
PAYMENT_LINE_ROWS = 4
PRINT_AREA = "A1:K50"
LAST_ROW = 50
LAST_COL = 11  # K

# Column widths from S260800018v4 (A/B/D spacers; H stock company-copy only).
COL_WIDTHS = {
    "A": 2.5,
    "B": 2.12,
    "C": 15.25,
    "D": 1.25,
    "E": 7.25,
    "F": 11.25,
    "G": 9.5,
    "H": 17.25,
    "I": 12.75,
    "J": 5.62,
    "K": 13.51,
}

# Row heights (pt) from S260800018v4 rows 1–50. Unlisted rows use DEFAULT_ROW_HEIGHT.
DEFAULT_ROW_HEIGHT = 13.5
ROW_HEIGHTS = {
    1: 43.25, 2: 9.75, 3: 14.9, 4: 7.5, 5: 18.0, 6: 9.75, 7: 13.5,
    8: 13.5, 9: 6.0,
    10: 13.5, 11: 13.5, 12: 13.5, 13: 13.5, 14: 13.5, 15: 13.5, 16: 12.0,
    17: 65.65, 18: 13.5, 19: 12.0, 20: 9.75, 21: 27.75,
    22: 13.5, 23: 15.65, 24: 15.75, 25: 9.75, 26: 18.0,
    27: 24.6, 28: 23.1, 29: 9.75, 30: 18.0, 31: 7.5, 32: 18.0, 33: 9.75,
    34: 13.5, 35: 13.5, 36: 6.0,
    37: 24.75, 38: 24.75, 39: 24.75, 40: 24.75, 41: 24.75,
    42: 13.5, 43: 12.0, 44: 9.75, 45: 27.75,
    46: 13.5, 47: 13.5, 48: 15.75, 49: 9.75, 50: 7.5,
}

# Print margins (inches) from the v4 blueprint. Generate preserves template margins.
MARGIN_LEFT = 0.15
MARGIN_RIGHT = 0.15
MARGIN_TOP = 0.2
MARGIN_BOTTOM = 0.2
MARGIN_HEADER = 0.511811023622047
MARGIN_FOOTER = 0.511811023622047

# Font: Windows print PCs resolve Microsoft JhengHei; Excel stores the name.
FONT_NAME = "Microsoft JhengHei"  # fallbacks: PMingLiU, Arial Unicode MS
FONT_DATA = Font(name=FONT_NAME, size=9)
FONT_SMALL = Font(name=FONT_NAME, size=8)
FONT_TOTAL = Font(name=FONT_NAME, size=11, bold=True)
FONT_HANDLER = Font(name=FONT_NAME, size=10)

ALIGN_LEFT_CLIP = Alignment(
    horizontal="left", vertical="center", wrap_text=True, shrink_to_fit=True,
)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_RIGHT_TOTAL = Alignment(horizontal="right", vertical="center")

FMT_MONEY = "#,##0.00"
FMT_WEIGHT = "0.000"
NAME_MAX_CHARS = 40
NOTES_MAX_CHARS = 80
EXCHANGE_MARKER = "對換"
COL_HEADER_CN = {
    3: "貨品",
    5: "成色",
    6: "重量",
    8: "庫存",
    9: "單價",
}
COL_HEADER_EN = {
    3: "Item",
    5: "Quality",
    6: "Weight",
    8: "Stock",
    9: "Unit ($)",
}
# Match header alignment to the values underneath (item left, money right).
COL_HEADER_ALIGN = {
    3: ALIGN_LEFT_CLIP,
    5: ALIGN_CENTER,
    6: ALIGN_CENTER,
    8: ALIGN_LEFT_CLIP,
    9: ALIGN_RIGHT,
}
STOCK_HEADER_COLS = (8,)
AMOUNT_HEADER_TEXTS = ("現金倉", "Cash Warehouse")
HANDLER_LABEL = "經手人 Handler:"
CUSTOMER_LABEL = "客戶 Customer :"
PAYMENT_LABEL = "付款方式 Payment Method :"

# Import formatters from the facade module (no circular import: this module
# does not import invoice_generator). Local copies of the tiny numeric helpers.


def _cash_warehouse_value(value):
    if value is None or value == "":
        return 0.0
    return float(value)


def _weight_used(value):
    """True when a 両/安士 value should occupy its own overlay row."""
    if value is None or value == "":
        return False
    try:
        return float(value) != 0.0
    except (TypeError, ValueError):
        return bool(str(value).strip())


def _clip(text, max_chars):
    if text is None:
        return ""
    s = str(text).strip()
    if len(s) <= max_chars:
        return s
    return s[: max_chars - 1] + "…"


def _company_layout(base=None):
    return dict(COMPANY_COPY)


def _is_company_copy(layout):
    return layout["items_start"] == COMPANY_COPY["items_start"]


def estimate_item_rows(item):
    """Rows for one line item: name/gram row + optional 両 + optional 安士."""
    rows = 1
    if _weight_used(item.get("weight_tael")):
        rows += 1
    if _weight_used(item.get("weight_oz")):
        rows += 1
    return rows


def estimate_items_block_rows(main_items, exchange_items=None):
    total = sum(estimate_item_rows(it) for it in (main_items or []))
    if exchange_items:
        total += 1  # 對換 marker
        total += sum(estimate_item_rows(it) for it in exchange_items)
    return total


def fit_items_for_print(main_items, exchange_items=None, max_rows=None):
    """Clamp items so they fit above the notes row. Returns (main, exchange, warning)."""
    if max_rows is None:
        max_rows = CUSTOMER_COPY["notes_row"] - CUSTOMER_COPY["items_start"]
    main_items = list(main_items or [])
    exchange_items = list(exchange_items or []) if exchange_items else []

    def fits(m, e):
        return estimate_items_block_rows(m, e or None) <= max_rows

    if fits(main_items, exchange_items):
        return main_items, exchange_items or None, None

    warning_parts = []
    trimmed_ex = list(exchange_items)
    while trimmed_ex and not fits(main_items, trimmed_ex):
        trimmed_ex.pop()
    if len(trimmed_ex) < len(exchange_items):
        warning_parts.append(
            f"對換貨品過多，列印僅顯示前 {len(trimmed_ex)} 項（其餘略過以免蓋住備註／合計）"
        )

    trimmed_main = list(main_items)
    while trimmed_main and not fits(trimmed_main, trimmed_ex):
        trimmed_main.pop()
    if len(trimmed_main) < len(main_items):
        warning_parts.append(
            f"貨品過多，列印僅顯示前 {len(trimmed_main)} 項（其餘略過以免蓋住備註／合計）"
        )

    if not trimmed_main and main_items:
        first = dict(main_items[0])
        first["weight_tael"] = None
        first["weight_oz"] = None
        trimmed_main = [first]
        warning_parts.append("貨品過多，列印僅保留第 1 項重量(克)")

    warning = "；".join(warning_parts) if warning_parts else None
    return trimmed_main, (trimmed_ex or None), warning


def _anchor_cell(ws, row, col):
    """Return the writable cell; merged ranges must be written at the origin."""
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        return cell
    for rng in ws.merged_cells.ranges:
        if cell.coordinate in rng:
            return ws.cell(row=rng.min_row, column=rng.min_col)
    return cell


def _clear_cell(ws, row, col):
    _anchor_cell(ws, row, col).value = None


def _clear_copy_variables(ws, layout):
    """Remove template sample/placeholder values; keep labels and checkboxes."""
    _clear_cell(ws, layout["invoice_no_row"], 10)
    _clear_cell(ws, layout["invoice_no_row"], 11)
    _clear_cell(ws, layout["info_row"], 5)
    _clear_cell(ws, layout["info_row"], 7)
    _clear_cell(ws, layout["info_row"], 11)
    for row in range(layout["items_start"], layout["notes_row"]):
        for col in (3, 5, 6, 7, 8, 9, 10, 11):
            _clear_cell(ws, row, col)
    notes_row = layout["notes_row"]
    for row in (notes_row, notes_row + 1):
        for col in (4, 5, 10, 11):
            _clear_cell(ws, row, col)
    _clear_cell(ws, layout["total_row"], 10)
    _clear_cell(ws, layout["total_row"], 11)
    for offset in range(PAYMENT_LINE_ROWS):
        _clear_cell(ws, layout["payment_row"] + offset, 3)
    _clear_cell(ws, layout["handler_row"], 6)
    _clear_amount_column_headers(ws, layout)


def _clear_amount_column_headers(ws, layout):
    """K-column (and company J/K merge) must not print 現金倉 / Cash Warehouse."""
    for row in (layout["items_start"] - 3, layout["items_start"] - 2):
        for col in (10, 11):
            cell = _anchor_cell(ws, row, col)
            text = str(cell.value or "").strip()
            if any(label in text for label in AMOUNT_HEADER_TEXTS):
                cell.value = None


def _style_cell(cell, font=None, align=None, num_fmt=None):
    cell.font = font or FONT_DATA
    if align:
        cell.alignment = align
    if num_fmt:
        cell.number_format = num_fmt


def _put_text(ws, row, col, value, font=None, align=None):
    if value is None or value == "":
        return
    cell = _anchor_cell(ws, row, col)
    cell.value = value
    _style_cell(cell, font=font, align=align or ALIGN_LEFT_CLIP)


def _put_number(ws, row, col, value, font=None, align=None, num_fmt=None):
    if value is None or value == "":
        return
    cell = _anchor_cell(ws, row, col)
    cell.value = float(value)
    _style_cell(cell, font=font, align=align, num_fmt=num_fmt)


def _parse_tx_date(tx_date):
    if isinstance(tx_date, str):
        return datetime.strptime(tx_date[:10], "%Y-%m-%d").date()
    if isinstance(tx_date, datetime):
        return tx_date.date()
    if isinstance(tx_date, date):
        return tx_date
    return tx_date


def _apply_geometry(ws):
    for letter, width in COL_WIDTHS.items():
        ws.column_dimensions[letter].width = width
    for row in range(1, LAST_ROW + 1):
        ws.row_dimensions[row].height = ROW_HEIGHTS.get(row, DEFAULT_ROW_HEIGHT)
    # Hide unused columns beyond K (none printed).
    for col in range(LAST_COL + 1, 16):
        ws.column_dimensions[get_column_letter(col)].hidden = True


def _apply_print_settings(ws, *, preserve_margins=False):
    """A4 portrait, 100% scale — never Fit-to-Page (that caused shrink/overflow)."""
    ws.page_setup.paperSize = int(ws.PAPERSIZE_A4)
    ws.page_setup.orientation = "portrait"
    ws.page_setup.scale = 100
    ws.page_setup.fitToWidth = 0
    ws.page_setup.fitToHeight = 0
    ws.page_setup.horizontalCentered = True
    if ws.sheet_properties.pageSetUpPr is None:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=False)
    else:
        ws.sheet_properties.pageSetUpPr.fitToPage = False
    if not preserve_margins:
        ws.page_margins = PageMargins(
            left=MARGIN_LEFT,
            right=MARGIN_RIGHT,
            top=MARGIN_TOP,
            bottom=MARGIN_BOTTOM,
            header=MARGIN_HEADER,
            footer=MARGIN_FOOTER,
        )
    ws.sheet_view.showGridLines = False
    ws.print_area = PRINT_AREA
    ws.page_setup.printGridLines = False
    ws.sheet_properties.pageSetUpPr.fitToPage = False


WEIGHT_HEADER_MERGES = ("F8:G8", "F9:G9", "F35:G35", "F36:G36")


def _unmerge_weight_header_boxes(ws):
    """重量 / Weight sit in column F only — do not span into the unit column G."""
    targets = {ref.replace("$", "") for ref in WEIGHT_HEADER_MERGES}
    for rng in list(ws.merged_cells.ranges):
        if str(rng).replace("$", "") in targets:
            ws.unmerge_cells(str(rng))


def _excel_stock_label(location):
    """Warehouse 存/取 only — never print 存 客戶 or 取 客戶."""
    text = (location or "").strip()
    if not text:
        return ""
    banned = {
        f"{INVENTORY_ACTION_DEPOSIT} {EXTERNAL_PARTY_LOCATION}",
        f"{INVENTORY_ACTION_WITHDRAW} {EXTERNAL_PARTY_LOCATION}",
    }
    if text in banned:
        return ""
    return text


def _write_item_block(
    ws, start_row, items, *, has_amount, write_stock, source, destination,
    currency, transaction_type, stop_before_row,
):
    currency = currency or DEFAULT_CASH_CURRENCY
    source = _excel_stock_label(source)
    destination = _excel_stock_label(destination)
    row = start_row
    for item in items:
        need = estimate_item_rows(item)
        if stop_before_row is not None and row + need > stop_before_row:
            break

        _put_text(ws, row, 3, item.get("item_type") or "", align=ALIGN_LEFT_CLIP)
        _put_text(ws, row, 5, item.get("quality") or "", align=ALIGN_CENTER)

        gram = item.get("weight_gram")
        if gram is not None and gram != "":
            _put_number(ws, row, 6, gram, align=ALIGN_CENTER, num_fmt=FMT_WEIGHT)
            _put_text(ws, row, 7, "克 Gram", font=FONT_SMALL, align=ALIGN_CENTER)

        if item.get("unit_price") is not None and item.get("unit_price") != "":
            _put_number(
                ws, row, 9, item.get("unit_price"),
                align=ALIGN_RIGHT, num_fmt=FMT_MONEY,
            )
        if has_amount:
            _put_text(ws, row, 10, currency, font=FONT_SMALL, align=ALIGN_RIGHT)
            signed = signed_cash_warehouse_amount(
                _cash_warehouse_value(item.get("amount")), transaction_type,
            )
            _put_number(
                ws, row, 11, signed, align=ALIGN_RIGHT, num_fmt=FMT_MONEY,
            )
        if write_stock and source:
            _put_text(ws, row, 8, source, font=FONT_SMALL, align=ALIGN_LEFT_CLIP)

        row += 1
        if _weight_used(item.get("weight_tael")):
            if stop_before_row is not None and row >= stop_before_row:
                break
            _put_number(
                ws, row, 6, item.get("weight_tael"),
                align=ALIGN_CENTER, num_fmt=FMT_WEIGHT,
            )
            _put_text(ws, row, 7, "両 Tael", font=FONT_SMALL, align=ALIGN_CENTER)
            if write_stock and destination:
                _put_text(ws, row, 8, destination, font=FONT_SMALL, align=ALIGN_LEFT_CLIP)
            row += 1
        elif write_stock and destination:
            # No 両 row: put 存/destination on the same H cell (newline) to avoid extra rows.
            if source:
                cell = _anchor_cell(ws, row - 1, 8)
                cell.value = f"{source}\n{destination}"
                _style_cell(cell, font=FONT_SMALL, align=ALIGN_LEFT_CLIP)
            else:
                _put_text(
                    ws, row - 1, 8, destination,
                    font=FONT_SMALL, align=ALIGN_LEFT_CLIP,
                )

        if _weight_used(item.get("weight_oz")):
            if stop_before_row is not None and row >= stop_before_row:
                break
            _put_number(
                ws, row, 6, item.get("weight_oz"),
                align=ALIGN_CENTER, num_fmt=FMT_WEIGHT,
            )
            _put_text(ws, row, 7, "安士 oz", font=FONT_SMALL, align=ALIGN_CENTER)
            row += 1
    return row


def _write_column_headers(ws, layout, include_stock=False):
    """Bilingual item-table headers. 庫存/Stock is company-copy only."""
    cn_row = layout["items_start"] - 3
    en_row = layout["items_start"] - 2
    skip = set() if include_stock else set(STOCK_HEADER_COLS)
    for col, text in COL_HEADER_CN.items():
        if col in skip:
            continue
        _put_text(
            ws, cn_row, col, text, font=FONT_SMALL,
            align=COL_HEADER_ALIGN.get(col, ALIGN_CENTER),
        )
    for col, text in COL_HEADER_EN.items():
        if col in skip:
            continue
        _put_text(
            ws, en_row, col, text, font=FONT_SMALL,
            align=COL_HEADER_ALIGN.get(col, ALIGN_CENTER),
        )


def _write_copy(ws, layout, invoice_data, main_items, exchange_items, tx_config, tx_date):
    has_amount = tx_config.get("has_amount", True)
    has_exchange = tx_config.get("has_exchange", False)
    is_company = _is_company_copy(layout)
    currency = invoice_data.get("invoice_currency") or DEFAULT_CASH_CURRENCY
    tx_type = invoice_data.get("transaction_type")
    invoice_label = (tx_config.get("invoice_label") or "").strip()
    if invoice_label:
        _put_text(
            ws, layout["invoice_no_row"], 10, invoice_label,
            font=FONT_SMALL, align=ALIGN_RIGHT,
        )
    _put_text(
        ws, layout["invoice_no_row"], 11,
        invoice_data.get("invoice_no") or "",
        font=FONT_DATA, align=ALIGN_RIGHT,
    )
    _put_text(
        ws, layout["info_row"], 3, CUSTOMER_LABEL,
        font=FONT_SMALL, align=ALIGN_LEFT_CLIP,
    )
    _put_text(
        ws, layout["info_row"], 5,
        _clip(invoice_data.get("customer_name"), NAME_MAX_CHARS),
        align=ALIGN_LEFT_CLIP,
    )
    date_s = tx_date.strftime("%Y-%m-%d") if hasattr(tx_date, "strftime") else str(tx_date)[:10]
    _put_text(ws, layout["info_row"], 11, date_s, align=ALIGN_RIGHT)

    phone = (invoice_data.get("customer_phone") or "").strip()
    if phone:
        _put_text(ws, layout["info_row"], 7, phone, font=FONT_SMALL, align=ALIGN_CENTER)

    _write_column_headers(ws, layout, include_stock=is_company)

    stock_source = invoice_data.get("source_location") if is_company else None
    stock_dest = invoice_data.get("destination_location") if is_company else None
    next_row = _write_item_block(
        ws, layout["items_start"], main_items,
        has_amount=has_amount,
        write_stock=is_company,
        source=stock_source,
        destination=stock_dest,
        currency=currency,
        transaction_type=tx_type,
        stop_before_row=layout["notes_row"],
    )

    if exchange_items and has_exchange:
        label_row = min(next_row, layout["notes_row"] - 2)
        if label_row < layout["notes_row"]:
            _put_text(ws, label_row, 3, EXCHANGE_MARKER, font=FONT_SMALL)
            ex_source = (
                invoice_data.get("exchange_source_location") if is_company else None
            ) or stock_source
            ex_dest = (
                invoice_data.get("exchange_destination_location") if is_company else None
            ) or stock_dest
            _write_item_block(
                ws, label_row + 1, exchange_items,
                has_amount=has_amount,
                write_stock=is_company,
                source=ex_source,
                destination=ex_dest,
                currency=currency,
                transaction_type=tx_type,
                stop_before_row=layout["notes_row"],
            )

    notes = _clip(invoice_data.get("notes") or "", NOTES_MAX_CHARS)
    if notes:
        notes_col = 5 if is_company else tx_config.get("customer_notes_col", 4)
        _put_text(ws, layout["notes_row"], notes_col, notes, font=FONT_SMALL, align=ALIGN_LEFT_CLIP)

    if has_amount:
        note_amount = invoice_data.get("note_amount")
        _put_text(ws, layout["notes_row"], 10, currency, font=FONT_SMALL, align=ALIGN_RIGHT)
        _put_number(
            ws, layout["notes_row"], 11,
            signed_cash_warehouse_amount(_cash_warehouse_value(note_amount), tx_type),
            align=ALIGN_RIGHT, num_fmt=FMT_MONEY,
        )
        signed_total = invoice_data.get("cash_warehouse_amount")
        if signed_total is None or signed_total == "":
            signed_total = signed_cash_warehouse_amount(
                _cash_warehouse_value(invoice_data.get("total_amount")), tx_type,
            )
        _put_text(
            ws, layout["total_row"], 10, currency,
            font=FONT_TOTAL, align=ALIGN_RIGHT,
        )
        _put_number(
            ws, layout["total_row"], 11, float(signed_total),
            font=FONT_TOTAL, align=ALIGN_RIGHT_TOTAL, num_fmt=FMT_MONEY,
        )

    from invoice_generator import format_payment_excel_lines

    _put_text(
        ws, layout["total_row"], 3, PAYMENT_LABEL,
        font=FONT_SMALL, align=ALIGN_LEFT_CLIP,
    )
    pay_row = layout["payment_row"]
    for idx, line in enumerate(format_payment_excel_lines(invoice_data.get("payment_method", ""))):
        if idx >= PAYMENT_LINE_ROWS:
            break
        if line:
            _put_text(ws, pay_row + idx, 3, line, font=FONT_SMALL, align=ALIGN_LEFT_CLIP)

    handler = (invoice_data.get("handler") or "").strip() or "Admin"
    _put_text(
        ws, layout["handler_row"] - 1, 6, HANDLER_LABEL,
        font=FONT_SMALL, align=ALIGN_LEFT_CLIP,
    )
    _put_text(
        ws, layout["handler_row"], 6, handler,
        font=FONT_HANDLER, align=ALIGN_LEFT_CLIP,
    )


def generate_invoice_excel(invoice_data, main_items, exchange_items=None):
    """
    Copy invoice_template.xlsx and fill the matching sheet with this invoice.

    Returns the absolute path of output/invoices/{invoice_no}.xlsx.
    """
    tx_type = invoice_data["transaction_type"]
    tx_config = TRANSACTION_TYPES[tx_type]
    sheet_name = tx_config["sheet"]

    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"找不到發票範本：{TEMPLATE_PATH}")

    main_items, exchange_items, overflow_warning = fit_items_for_print(
        main_items, exchange_items,
    )
    if overflow_warning:
        invoice_data["print_warning"] = overflow_warning

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / f"{invoice_data['invoice_no']}.xlsx"
    shutil.copy(TEMPLATE_PATH, output_path)

    wb = load_workbook(output_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"範本中找不到分頁：{sheet_name}（現有：{', '.join(wb.sheetnames)}）"
        )
    for name in list(wb.sheetnames):
        if name != sheet_name:
            del wb[name]
    ws = wb[sheet_name]
    wb.active = ws

    _clear_copy_variables(ws, CUSTOMER_COPY)
    _clear_copy_variables(ws, _company_layout())
    _apply_print_settings(ws, preserve_margins=True)
    _unmerge_weight_header_boxes(ws)

    tx_date = _parse_tx_date(invoice_data["transaction_date"])
    _write_copy(
        ws, CUSTOMER_COPY, invoice_data, main_items, exchange_items, tx_config, tx_date,
    )
    _write_copy(
        ws, _company_layout(), invoice_data, main_items, exchange_items, tx_config, tx_date,
    )

    wb.save(output_path)
    excel_abs = str(output_path.resolve())
    logger.info("Template invoice Excel saved: %s", excel_abs)
    invoice_data.pop("pdf_path", None)
    return excel_abs
