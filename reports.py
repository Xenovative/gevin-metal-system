from datetime import date
from pathlib import Path

import pandas as pd
from sqlalchemy import and_

from config import REPORT_DIR
from database import CashMovement, InventoryMovement, Invoice
from invoice_generator import format_payment_method_display
from cash import extract_cash_amount, extract_cash_currency, get_cash_balances


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


def generate_inventory_report(session, start_date, end_date, period_label):
    """Generate inventory movement report for a date range."""
    movements = (
        session.query(InventoryMovement)
        .filter(
            InventoryMovement.movement_date >= start_date,
            InventoryMovement.movement_date <= end_date,
        )
        .order_by(InventoryMovement.movement_date, InventoryMovement.id)
        .all()
    )

    rows = []
    status_map = _invoice_status_map(session, {m.invoice_no for m in movements if m.invoice_no})
    for m in movements:
        rows.append({
            "日期": m.movement_date.strftime("%Y-%m-%d"),
            "單號": m.invoice_no or "",
            "單據狀態": _invoice_status_label(status_map.get(m.invoice_no, "active")),
            "記錄類型": _movement_kind_label(getattr(m, "movement_kind", "normal")),
            "交易性質": m.transaction_type,
            "方向": "入倉" if m.direction == "in" else "出倉",
            "貨品": m.item_type,
            "成色": m.quality,
            "重量(克)": m.weight_gram,
            "倉存存取": m.source_location or "",
            "倉存位置": m.destination_location or "",
            "客戶": m.customer_name or "",
            "經手人": m.handler or "",
            "備註": m.notes or "",
        })

    df = pd.DataFrame(rows)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    filename = f"倉存報表_{period_label}_{start_date}_{end_date}.xlsx"
    output_path = REPORT_DIR / filename

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="倉存明細", index=False)

        if not df.empty:
            summary = (
                df.groupby(["交易性質", "方向", "貨品", "成色"])
                .agg({"重量(克)": "sum"})
                .reset_index()
            )
            summary.to_excel(writer, sheet_name="匯總", index=False)

        cash_df = _build_cash_report_df(session, start_date, end_date)
        cash_df.to_excel(writer, sheet_name="現金倉明細", index=False)

        warehouse_df = _build_warehouse_summary_df(session)
        warehouse_df.to_excel(writer, sheet_name="倉庫總覽", index=False)

    return str(output_path), df


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
            "金額": m.amount,
            "貨幣": m.currency or "HKD$",
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
            "金額": 0, "貨幣": "HKD$", "倉庫": "現金倉",
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
            "品種": "現金 (HKD$)",
            "結存(克)": "",
            "結存(金額)": 0,
            "備註": "即時結存",
        })
    return pd.DataFrame(rows)


def generate_invoice_report(session, start_date, end_date, period_label):
    """Generate invoice summary report for a date range."""
    invoices = (
        session.query(Invoice)
        .filter(
            Invoice.transaction_date >= start_date,
            Invoice.transaction_date <= end_date,
        )
        .order_by(Invoice.transaction_date, Invoice.id)
        .all()
    )

    rows = []
    for inv in invoices:
        rows.append({
            "日期": inv.transaction_date.strftime("%Y-%m-%d"),
            "單號": inv.invoice_no,
            "狀態": "已作廢" if inv.status == "voided" else "正常",
            "交易性質": inv.transaction_type,
            "客戶": inv.customer_name,
            "金額": inv.total_amount,
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
    filename = f"發票報表_{period_label}_{start_date}_{end_date}.xlsx"
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
