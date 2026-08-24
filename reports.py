from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from sqlalchemy import and_

from config import CASH_WAREHOUSE, METAL_WAREHOUSES, REPORT_DIR
from database import CashMovement, InventoryMovement, Invoice
from invoice_generator import format_money, format_payment_method_display
from cash import extract_cash_amount, extract_cash_currency, get_cash_balances
from warehouse import resolve_movement_warehouse

PERIOD_ALL = "全部"
PERIOD_DAY = "每日"
PERIOD_MONTH = "每月"
PERIOD_YEAR = "每年"
PERIOD_CHOICES = [PERIOD_ALL, PERIOD_DAY, PERIOD_MONTH, PERIOD_YEAR]
WAREHOUSE_FILTER_ALL = "全部"
WAREHOUSE_FILTER_CHOICES = [WAREHOUSE_FILTER_ALL, *METAL_WAREHOUSES, CASH_WAREHOUSE]
UNASSIGNED_WAREHOUSE = "未歸倉庫"

SUMMARY_METAL_COLUMNS = [
    "倉庫", "期初(克)", "入倉(克)", "出倉(克)", "淨額(克)", "期末(克)", "單據數",
]
LEDGER_METAL_COLUMNS = [
    "日期", "單號", "倉庫", "方向", "重量(+/-克)", "累計(克)",
    "貨品", "成色", "交易性質", "客戶",
]
SUMMARY_CASH_COLUMNS = [
    "倉庫", "貨幣", "期初", "存入", "支出", "淨額", "期末", "單據數",
]
LEDGER_CASH_COLUMNS = [
    "日期", "單號", "方向", "金額(+/-)", "累計", "貨幣", "交易性質", "客戶",
]


def empty_summary_metal_df():
    return pd.DataFrame(columns=SUMMARY_METAL_COLUMNS)


def empty_ledger_metal_df():
    return pd.DataFrame(columns=LEDGER_METAL_COLUMNS)


def empty_summary_cash_df():
    return pd.DataFrame(columns=SUMMARY_CASH_COLUMNS)


def empty_ledger_cash_df():
    return pd.DataFrame(columns=LEDGER_CASH_COLUMNS)


def resolve_report_period(period, report_date=None, year=None, month=None):
    """Return (start_date, end_date, label). 全部 → (None, None, 全部)."""
    from invoice_number import parse_tx_date

    period = (period or PERIOD_ALL).strip()
    today = date.today()
    if period == PERIOD_ALL:
        return None, None, PERIOD_ALL
    if period == PERIOD_DAY:
        day = parse_tx_date(report_date) if report_date not in (None, "") else today
        return day, day, f"每日_{day}"
    if period == PERIOD_MONTH:
        year = int(year or today.year)
        month = int(month or today.month)
        start = date(year, month, 1)
        if month == 12:
            end = date(year + 1, 1, 1) - timedelta(days=1)
        else:
            end = date(year, month + 1, 1) - timedelta(days=1)
        return start, end, f"{year}年{month:02d}月"
    if period == PERIOD_YEAR:
        year = int(year or today.year)
        return date(year, 1, 1), date(year, 12, 31), f"{year}年"
    return None, None, PERIOD_ALL


def _fmt_day(value):
    if value is None:
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    return str(value)[:10]


def _signed_grams(movement):
    grams = float(movement.weight_gram or 0)
    return grams if movement.direction == "in" else -grams


def _signed_cash(movement):
    amount = float(movement.amount or 0)
    return amount if movement.direction == "in" else -amount


def build_metal_inventory_view(movements, start_date=None, end_date=None, warehouse="全部"):
    """Ledger + summary for A/B/C. Running total starts from the first DB record."""
    warehouse = (warehouse or WAREHOUSE_FILTER_ALL).strip()
    if warehouse == CASH_WAREHOUSE:
        return empty_summary_metal_df(), empty_ledger_metal_df()

    tracked = list(METAL_WAREHOUSES) + [UNASSIGNED_WAREHOUSE]
    if warehouse in METAL_WAREHOUSES:
        tracked = [warehouse]

    running = {wh: 0.0 for wh in list(METAL_WAREHOUSES) + [UNASSIGNED_WAREHOUSE]}
    inflow = {wh: 0.0 for wh in running}
    outflow = {wh: 0.0 for wh in running}
    invoices = {wh: set() for wh in running}
    opening = {wh: 0.0 for wh in running}
    opening_ready = start_date is None
    ledger_rows = []

    def emit_opening():
        for wh in tracked:
            ledger_rows.append({
                "日期": _fmt_day(start_date) if start_date else "",
                "單號": "",
                "倉庫": wh,
                "方向": "期初",
                "重量(+/-克)": 0.0,
                "累計(克)": round(opening[wh], 3),
                "貨品": "",
                "成色": "",
                "交易性質": "",
                "客戶": "",
            })

    ordered = sorted(
        movements,
        key=lambda m: (m.movement_date or date.min, getattr(m, "id", 0) or 0),
    )
    for m in ordered:
        wh = resolve_movement_warehouse(m) or UNASSIGNED_WAREHOUSE
        day = m.movement_date
        if end_date and day and day > end_date:
            break
        if warehouse in METAL_WAREHOUSES and wh != warehouse:
            continue
        signed = _signed_grams(m)
        if start_date and day and day < start_date:
            running[wh] += signed
            continue
        if not opening_ready:
            opening = dict(running)
            opening_ready = True
            if start_date:
                emit_opening()
        running[wh] += signed
        if signed >= 0:
            inflow[wh] += signed
        else:
            outflow[wh] += abs(signed)
        if m.invoice_no:
            invoices[wh].add(m.invoice_no)
        ledger_rows.append({
            "日期": _fmt_day(day),
            "單號": m.invoice_no or "",
            "倉庫": wh,
            "方向": "入倉" if m.direction == "in" else "出倉",
            "重量(+/-克)": round(signed, 3),
            "累計(克)": round(running[wh], 3),
            "貨品": m.item_type or "",
            "成色": m.quality or "",
            "交易性質": m.transaction_type or "",
            "客戶": m.customer_name or "",
        })

    if start_date and not opening_ready:
        opening = dict(running)
        emit_opening()

    if start_date is None:
        opening = {wh: 0.0 for wh in running}

    summary_rows = []
    for wh in tracked:
        net = inflow[wh] - outflow[wh]
        close = opening[wh] + net
        summary_rows.append({
            "倉庫": wh,
            "期初(克)": round(opening[wh], 3),
            "入倉(克)": round(inflow[wh], 3),
            "出倉(克)": round(outflow[wh], 3),
            "淨額(克)": round(net, 3),
            "期末(克)": round(close, 3),
            "單據數": len(invoices[wh]),
        })
    summary_df = pd.DataFrame(summary_rows, columns=SUMMARY_METAL_COLUMNS)
    ledger_df = pd.DataFrame(ledger_rows, columns=LEDGER_METAL_COLUMNS) if ledger_rows else empty_ledger_metal_df()
    return summary_df, ledger_df


def build_cash_inventory_view(movements, start_date=None, end_date=None, warehouse="全部"):
    """Cash-warehouse ledger with running total from the first DB record."""
    warehouse = (warehouse or WAREHOUSE_FILTER_ALL).strip()
    if warehouse in METAL_WAREHOUSES:
        return empty_summary_cash_df(), empty_ledger_cash_df()

    running = {}
    inflow = {}
    outflow = {}
    invoices = {}
    opening = {}
    opening_ready = start_date is None
    ledger_rows = []

    def ensure(currency):
        running.setdefault(currency, 0.0)
        inflow.setdefault(currency, 0.0)
        outflow.setdefault(currency, 0.0)
        invoices.setdefault(currency, set())
        opening.setdefault(currency, 0.0)

    def emit_opening():
        for currency, bal in sorted(opening.items()):
            ledger_rows.append({
                "日期": _fmt_day(start_date) if start_date else "",
                "單號": "",
                "方向": "期初",
                "金額(+/-)": 0.0,
                "累計": round(bal, 2),
                "貨幣": currency,
                "交易性質": "",
                "客戶": "",
            })

    ordered = sorted(
        movements,
        key=lambda m: (m.movement_date or date.min, getattr(m, "id", 0) or 0),
    )
    for m in ordered:
        currency = m.currency or "HKD"
        ensure(currency)
        signed = _signed_cash(m)
        day = m.movement_date
        if start_date and day and day < start_date:
            running[currency] += signed
            continue
        if end_date and day and day > end_date:
            break
        if not opening_ready:
            opening = dict(running)
            opening_ready = True
            if start_date:
                emit_opening()
        running[currency] += signed
        if signed >= 0:
            inflow[currency] += signed
        else:
            outflow[currency] += abs(signed)
        if m.invoice_no:
            invoices[currency].add(m.invoice_no)
        ledger_rows.append({
            "日期": _fmt_day(day),
            "單號": m.invoice_no or "",
            "方向": "存入" if m.direction == "in" else "支出",
            "金額(+/-)": round(signed, 2),
            "累計": round(running[currency], 2),
            "貨幣": currency,
            "交易性質": m.transaction_type or "",
            "客戶": m.customer_name or "",
        })

    if start_date and not opening_ready:
        opening = dict(running)
        emit_opening()
    if start_date is None:
        opening = {c: 0.0 for c in running}

    currencies = sorted(set(running) | set(opening))
    summary_rows = []
    for currency in currencies:
        ensure(currency)
        net = inflow[currency] - outflow[currency]
        close = opening.get(currency, 0.0) + net
        summary_rows.append({
            "倉庫": CASH_WAREHOUSE,
            "貨幣": currency,
            "期初": round(opening.get(currency, 0.0), 2),
            "存入": round(inflow[currency], 2),
            "支出": round(outflow[currency], 2),
            "淨額": round(net, 2),
            "期末": round(close, 2),
            "單據數": len(invoices[currency]),
        })
    if not summary_rows:
        summary_rows.append({
            "倉庫": CASH_WAREHOUSE,
            "貨幣": "HKD",
            "期初": 0.0, "存入": 0.0, "支出": 0.0, "淨額": 0.0, "期末": 0.0, "單據數": 0,
        })
    summary_df = pd.DataFrame(summary_rows, columns=SUMMARY_CASH_COLUMNS)
    ledger_df = pd.DataFrame(ledger_rows, columns=LEDGER_CASH_COLUMNS) if ledger_rows else empty_ledger_cash_df()
    return summary_df, ledger_df


def build_inventory_view(session, start_date=None, end_date=None, warehouse="全部"):
    """Load DB movements and build period/warehouse summary + cumulative ledgers."""
    warehouse = (warehouse or WAREHOUSE_FILTER_ALL).strip()
    metal_q = session.query(InventoryMovement).order_by(
        InventoryMovement.movement_date, InventoryMovement.id,
    )
    cash_q = session.query(CashMovement).order_by(
        CashMovement.movement_date, CashMovement.id,
    )
    summary_metal, ledger_metal = build_metal_inventory_view(
        metal_q.all(), start_date, end_date, warehouse,
    )
    summary_cash, ledger_cash = build_cash_inventory_view(
        cash_q.all(), start_date, end_date, warehouse,
    )
    return {
        "summary_metal": summary_metal,
        "ledger_metal": ledger_metal,
        "summary_cash": summary_cash,
        "ledger_cash": ledger_cash,
    }


def export_inventory_view(session, start_date, end_date, period_label, warehouse="全部"):
    """Write the on-screen inventory view to Excel for printing."""
    view = build_inventory_view(session, start_date, end_date, warehouse)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    wh_tag = (warehouse or WAREHOUSE_FILTER_ALL).replace(" ", "")
    if start_date and end_date:
        filename = f"倉存報表_{period_label}_{wh_tag}_{start_date}_{end_date}.xlsx"
    else:
        filename = f"倉存報表_{period_label}_{wh_tag}.xlsx"
    output_path = REPORT_DIR / filename

    item_df = _build_item_detail_df(session, start_date, end_date, warehouse)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        view["summary_metal"].to_excel(writer, sheet_name="倉庫進出匯總", index=False)
        view["ledger_metal"].to_excel(writer, sheet_name="倉庫累計明細", index=False)
        view["summary_cash"].to_excel(writer, sheet_name="現金倉進出匯總", index=False)
        view["ledger_cash"].to_excel(writer, sheet_name="現金倉累計明細", index=False)
        item_df.to_excel(writer, sheet_name="倉存明細", index=False)
        _build_warehouse_summary_df(session).to_excel(writer, sheet_name="倉庫總覽", index=False)
    return str(output_path), view


def _build_item_detail_df(session, start_date, end_date, warehouse="全部"):
    warehouse = (warehouse or WAREHOUSE_FILTER_ALL).strip()
    q = session.query(InventoryMovement)
    if start_date:
        q = q.filter(InventoryMovement.movement_date >= start_date)
    if end_date:
        q = q.filter(InventoryMovement.movement_date <= end_date)
    movements = q.order_by(InventoryMovement.movement_date, InventoryMovement.id).all()
    status_map = _invoice_status_map(session, {m.invoice_no for m in movements if m.invoice_no})
    rows = []
    for m in movements:
        wh = resolve_movement_warehouse(m) or UNASSIGNED_WAREHOUSE
        if warehouse in METAL_WAREHOUSES and wh != warehouse:
            continue
        if warehouse == CASH_WAREHOUSE:
            continue
        rows.append({
            "日期": _fmt_day(m.movement_date),
            "單號": m.invoice_no or "",
            "單據狀態": _invoice_status_label(status_map.get(m.invoice_no, "active")),
            "記錄類型": _movement_kind_label(getattr(m, "movement_kind", "normal")),
            "交易性質": m.transaction_type,
            "倉庫": wh,
            "方向": "入倉" if m.direction == "in" else "出倉",
            "重量(+/-克)": round(_signed_grams(m), 3),
            "貨品": m.item_type,
            "成色": m.quality,
            "重量(克)": m.weight_gram,
            "重量(両)": m.weight_tael,
            "重量(安士)": getattr(m, "weight_oz", None) or "",
            "倉存存取": m.source_location or "",
            "倉存位置": m.destination_location or "",
            "客戶": m.customer_name or "",
            "經手人": m.handler or "",
            "備註": m.notes or "",
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=[
        "日期", "單號", "單據狀態", "記錄類型", "交易性質", "倉庫", "方向",
        "重量(+/-克)", "貨品", "成色", "重量(克)", "重量(両)", "重量(安士)",
        "倉存存取", "倉存位置", "客戶", "經手人", "備註",
    ])


def _invoice_status_map(session, invoice_nos):
    if not invoice_nos:
        return {}
    return {
        inv.invoice_no: inv.status
        for inv in session.query(Invoice).filter(Invoice.invoice_no.in_(invoice_nos)).all()
    }


def _invoice_status_label(status):
    return "已作廢" if status == "voided" else "正常"


def _movement_kind_label(kind):
    return "作廢沖銷" if kind == "reversal" else "正常"


def _date_filter(query, model, start_date, end_date):
    return query.filter(
        and_(
            model.movement_date >= start_date,
            model.movement_date <= end_date,
        )
    )


def generate_inventory_report(session, start_date, end_date, period_label, warehouse="全部"):
    """Generate inventory Excel for a date range (optional warehouse filter)."""
    path, view = export_inventory_view(
        session, start_date, end_date, period_label, warehouse,
    )
    return path, view["ledger_metal"]


def _build_cash_report_df(session, start_date, end_date):
    rows = []
    movements = (
        session.query(CashMovement)
        .filter(
            CashMovement.movement_date >= start_date,
            CashMovement.movement_date <= end_date,
        )
        .order_by(CashMovement.movement_date, CashMovement.id)
        .all()
    )
    status_map = _invoice_status_map(session, {m.invoice_no for m in movements if m.invoice_no})
    for m in movements:
        rows.append({
            "日期": m.movement_date.strftime("%Y-%m-%d"),
            "單號": m.invoice_no or "",
            "單據狀態": _invoice_status_label(status_map.get(m.invoice_no, "active")),
            "記錄類型": _movement_kind_label(getattr(m, "movement_kind", "normal")),
            "交易性質": m.transaction_type,
            "方向": "存入" if m.direction == "in" else "支出",
            # Amount = Cash Warehouse signed value (default 0)
            "金額": (m.amount or 0) if m.direction == "in" else -(m.amount or 0),
            "貨幣": m.currency or "HKD",
            "倉庫": m.warehouse or "現金倉",
            "客戶": m.customer_name or "",
            "經手人": m.handler or "",
            "備註": m.notes or "",
        })
    if not movements:
        rows.append({
            "日期": "", "單號": "", "交易性質": "", "方向": "",
            "金額": 0, "貨幣": "", "倉庫": "現金倉", "客戶": "", "經手人": "", "備註": "本期間無現金記錄",
        })
    balances = get_cash_balances(session)
    for currency, balance in balances.items():
        rows.append({
            "日期": "", "單號": "", "交易性質": "", "方向": f"現金倉結存 ({currency})",
            "金額": balance, "貨幣": currency, "倉庫": "現金倉",
            "客戶": "", "經手人": "", "備註": "",
        })
    if not balances:
        rows.append({
            "日期": "", "單號": "", "交易性質": "", "方向": "現金倉結存",
            "金額": 0, "貨幣": "HKD", "倉庫": "現金倉",
            "客戶": "", "經手人": "", "備註": "",
        })
    return pd.DataFrame(rows)


def _build_warehouse_summary_df(session):
    from config import METAL_WAREHOUSES, SAFE_SUMMARY_CATEGORIES
    from inventory import get_safe_totals, get_unassigned_metal_totals

    totals = get_safe_totals(session)
    unassigned = get_unassigned_metal_totals(session)
    rows = []
    for wh in METAL_WAREHOUSES:
        for cat in SAFE_SUMMARY_CATEGORIES:
            rows.append({
                "倉庫": wh,
                "品種": cat,
                "結存(克)": totals[wh][cat],
                "結存(金額)": "",
                "備註": "即時結存",
            })
    for cat, grams in unassigned.items():
        rows.append({
            "倉庫": "未歸倉庫",
            "品種": cat,
            "結存(克)": grams,
            "結存(金額)": "",
            "備註": "缺少倉存欄位的歷史記錄",
        })
    balances = get_cash_balances(session)
    for currency, balance in balances.items():
        rows.append({
            "倉庫": "現金倉",
            "品種": f"現金 ({currency})",
            "結存(克)": "",
            "結存(金額)": balance,
            "備註": "即時結存",
        })
    if not balances:
        rows.append({
            "倉庫": "現金倉",
            "品種": "現金 (HKD)",
            "結存(克)": "",
            "結存(金額)": 0,
            "備註": "即時結存",
        })
    return pd.DataFrame(rows)


def generate_invoice_report(session, start_date, end_date, period_label):
    """Generate invoice summary report for a date range (None = all records)."""
    q = session.query(Invoice)
    if start_date:
        q = q.filter(Invoice.transaction_date >= start_date)
    if end_date:
        q = q.filter(Invoice.transaction_date <= end_date)
    invoices = q.order_by(Invoice.transaction_date, Invoice.id).all()

    rows = []
    for inv in invoices:
        rows.append({
            "日期": inv.transaction_date.strftime("%Y-%m-%d"),
            "單號": inv.invoice_no,
            "狀態": "已作廢" if inv.status == "voided" else "正常",
            "交易性質": inv.transaction_type,
            "客戶": inv.customer_name,
            "客戶電話": inv.customer_phone or "",
            "金額": inv.total_amount,
            "貨幣": getattr(inv, "invoice_currency", None) or "HKD",
            "金額(含貨幣)": format_money(
                inv.total_amount,
                getattr(inv, "invoice_currency", None) or "HKD",
            ),
            "現金金額": extract_cash_amount(inv.payment_method or ""),
            "現金貨幣": extract_cash_currency(inv.payment_method or ""),
            "倉存存取": inv.source_location or "",
            "倉存位置": inv.destination_location or "",
            "經手人": inv.handler or "",
            "付款方式": format_payment_method_display(inv.payment_method or ""),
            "備註": inv.notes or "",
        })

    df = pd.DataFrame(rows)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    filename = (
        f"發票報表_{period_label}_{start_date}_{end_date}.xlsx"
        if start_date and end_date
        else f"發票報表_{period_label}.xlsx"
    )
    output_path = REPORT_DIR / filename

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="發票明細", index=False)
        if not df.empty:
            by_type = (
                df.groupby("交易性質")
                .agg({"單號": "count", "金額": "sum"})
                .rename(columns={"單號": "數量"})
                .reset_index()
            )
            by_type.to_excel(writer, sheet_name="按性質匯總", index=False)

    return str(output_path), df


def daily_report(session, report_date=None):
    report_date = report_date or date.today()
    return (
        generate_inventory_report(session, report_date, report_date, "每日"),
        generate_invoice_report(session, report_date, report_date, "每日"),
    )


def monthly_report(session, year=None, month=None):
    today = date.today()
    year = year or today.year
    month = month or today.month
    start = date(year, month, 1)
    if month == 12:
        end = date(year + 1, 1, 1)
    else:
        end = date(year, month + 1, 1)
    from datetime import timedelta
    end = end - timedelta(days=1)
    label = f"{year}年{month:02d}月"
    return (
        generate_inventory_report(session, start, end, label),
        generate_invoice_report(session, start, end, label),
    )


def yearly_report(session, year=None):
    year = year or date.today().year
    start = date(year, 1, 1)
    end = date(year, 12, 31)
    label = f"{year}年"
    return (
        generate_inventory_report(session, start, end, label),
        generate_invoice_report(session, start, end, label),
    )
