"""
金滿堂發票系統
逐步引導銷售人員填寫資料，自動生成 Excel 發票並記錄倉存。
"""

import json
import logging
import os
from datetime import date, datetime
from pathlib import Path

# Quiet Gradio analytics on headless Linux servers
os.environ.setdefault("GRADIO_ANALYTICS_ENABLED", "False")
# Debug: verbose Gradio + app logs when GEVIN_DEBUG=1
if os.environ.get("GEVIN_DEBUG", "").strip() in ("1", "true", "True", "yes"):
    os.environ.setdefault("GRADIO_DEBUG", "1")
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
else:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

logger = logging.getLogger("gevin")

import gradio as gr
import pandas as pd

from config import (
    BASE_DIR,
    CASH_CURRENCIES,
    DEFAULT_CASH_CURRENCY,
    INVENTORY_ACTION_CHOICES,
    INVENTORY_ACTION_DEPOSIT,
    INVENTORY_ACTION_WITHDRAW,
    ITEM_TYPES,
    OUTPUT_DIR,
    PAYMENT_METHODS,
    REPORT_DIR,
    WAREHOUSE_LOCATION_CHOICES,
    compose_receipt_storage,
    TRANSACTION_TYPES,
    UNITS,
    ensure_runtime_dirs,
)
from auth import (
    authenticate,
    can_download_reports,
    can_view_inventory,
    create_staff_user,
    has_permission,
    is_admin,
    is_logged_in,
    list_users,
    load_audit_logs,
    log_audit,
    require_download_reports,
    require_login,
    require_permission,
    require_view_inventory,
    update_user_profile,
)
from invoice_void import (
    ARCHIVE_HEADERS,
    cancel_invoice,
    cancelled_archive_file,
    load_cancelled_archive_rows,
    search_cancelled_archive,
)
from cash import build_cash_movement, extract_cash_payment, get_cash_balances, signed_cash_warehouse_amount
from database import init_db, save_invoice
from inventory import (
    build_inventory_movements,
    build_safe_summary_html,
    get_current_stock,
    get_safe_totals,
    get_unassigned_metal_totals,
)
from invoice_generator import (
    format_money,
    format_payment_method_display,
    generate_invoice_excel,
    resolve_invoice_excel_path,
)
from invoice_number import get_next_invoice_number, validate_invoice_number
from reports import (
    PERIOD_CHOICES,
    WAREHOUSE_FILTER_CHOICES,
    build_inventory_view,
    empty_ledger_cash_df,
    empty_ledger_metal_df,
    empty_summary_cash_df,
    empty_summary_metal_df,
    export_inventory_view,
    generate_invoice_report,
    resolve_report_period,
)

session = init_db()


def _invoice_status_label(status):
    return "已作廢" if status == "voided" else "正常"


def _movement_kind_label(kind):
    return "作廢沖銷" if kind == "reversal" else "正常"


def _handler_field_update(user):
    if not is_logged_in(user):
        return gr.update(value="", interactive=True)
    return gr.update(
        value=user.get("display_name", ""),
        interactive=is_admin(user),
    )


def _tab_visibility_for_user(user):
    if not is_logged_in(user):
        return (
            gr.update(visible=False),
            gr.update(visible=False),
            gr.update(visible=False),
            gr.update(visible=False),
        )
    return (
        gr.update(visible=True),
        gr.update(visible=can_view_inventory(user)),
        gr.update(visible=can_download_reports(user)),
        gr.update(visible=is_admin(user)),
    )


def do_login(username, password):
    fail_login = (
        None,
        "❌ 帳號或密碼錯誤",
        gr.update(visible=True),
        gr.update(visible=False),
        _handler_field_update(None),
        *_tab_visibility_for_user(None),
        "",
    )
    fail_review = _empty_review_payload("請先登入")
    try:
        user = authenticate(session, username, password)
        if not user:
            return (*fail_login, *fail_review)
        log_audit(session, user, "login")
        session.commit()
    except Exception as exc:
        session.rollback()
        fail_exc = (
            None,
            f"❌ 登入失敗：{exc}",
            gr.update(visible=True),
            gr.update(visible=False),
            _handler_field_update(None),
            *_tab_visibility_for_user(None),
            "",
        )
        return (*fail_exc, *fail_review)
    role_label = "Admin" if is_admin(user) else "員工"
    login_ok = (
        user,
        f"✅ 已登入：{user['display_name']}（{role_label}）",
        gr.update(visible=False),
        gr.update(visible=True),
        _handler_field_update(user),
        *_tab_visibility_for_user(user),
        _show_user_info_text(user),
    )
    # Pre-load Invoice Review so dashboard data is visible immediately
    if can_view_inventory(user):
        review = load_review_page()
    else:
        review = _empty_review_payload("❌ 您沒有權限執行此操作")
    return (*login_ok, *review)


def _show_user_info_text(user):
    if not is_logged_in(user):
        return ""
    role_label = "Admin" if is_admin(user) else "員工"
    return f"**目前使用者：** {user['display_name']}（{role_label}）"


def do_logout(current_user):
    if is_logged_in(current_user):
        log_audit(session, current_user, "logout")
        session.commit()
    return (
        None,
        "",
        gr.update(visible=True),
        gr.update(visible=False),
        _handler_field_update(None),
        *_tab_visibility_for_user(None),
    )


def run_cancel_invoice(invoice_no, current_user):
    result = cancel_invoice(session, invoice_no, current_user)
    archive = cancelled_archive_file()
    if result.startswith("✅"):
        return result, load_movements(), archive
    return result, gr.update(), gr.update()


def run_create_user(username, display_name, password, current_user):
    msg = create_staff_user(session, current_user, username, display_name, password)
    users_df = run_load_users(current_user)
    if msg.startswith("✅"):
        return (
            msg,
            users_df,
            gr.update(value=""),
            gr.update(value=""),
            gr.update(value=""),
        )
    return msg, users_df, gr.update(), gr.update(), gr.update()


def run_update_user(username, new_display_name, new_password, current_user):
    msg = update_user_profile(session, current_user, username, new_display_name, new_password)
    users_df = run_load_users(current_user)
    return msg, users_df


def run_load_users(current_user):
    from auth import require_admin
    err = require_admin(current_user)
    if err:
        return pd.DataFrame(columns=["帳號", "姓名", "權限", "狀態"])
    users = list_users(session)
    return pd.DataFrame(users) if users else pd.DataFrame(columns=["帳號", "姓名", "權限", "狀態"])


def run_load_audit(current_user):
    from auth import require_admin
    err = require_admin(current_user)
    if err:
        return pd.DataFrame(columns=["時間", "操作者", "動作", "對象", "詳情"])
    logs = load_audit_logs(session)
    return pd.DataFrame(logs) if logs else pd.DataFrame(columns=["時間", "操作者", "動作", "對象", "詳情"])


def parse_line_items_json(items_json):
    if not items_json:
        return []
    try:
        items = json.loads(items_json) if isinstance(items_json, str) else items_json
        return items if isinstance(items, list) else []
    except json.JSONDecodeError:
        return []


def _is_other_item(value):
    return value and "其他" in str(value)


def _resolve_item_type(item_type, custom_item_type):
    if _is_other_item(item_type):
        return (custom_item_type or "").strip()
    return item_type


def toggle_other_item_field(item_type):
    return gr.update(visible=_is_other_item(item_type))


def _is_empty_number(value):
    return value is None or value == ""


def add_line_item(items_json, item_type, custom_item_type, quality,
                  weight_gram, weight_tael, weight_oz, unit_price, amount, unit):
    items = parse_line_items_json(items_json)
    resolved_item = _resolve_item_type(item_type, custom_item_type)
    resolved_quality = (quality or "").strip()

    if _is_other_item(item_type) and not resolved_item:
        return items_json, _items_to_table(items), "❌ 選擇「其他」貨品時，請填寫貨品名稱"
    if _is_empty_number(weight_gram):
        return items_json, _items_to_table(items), "❌ 請填寫重量(克)"

    gram = float(weight_gram)
    if gram < 0:
        return items_json, _items_to_table(items), "❌ 重量(克)不可為負數"

    tael = None
    if not _is_empty_number(weight_tael):
        tael = float(weight_tael)
        if tael < 0:
            return items_json, _items_to_table(items), "❌ 重量(両)不可為負數"

    oz = None
    if not _is_empty_number(weight_oz):
        oz = float(weight_oz)
        if oz < 0:
            return items_json, _items_to_table(items), "❌ 重量(安士)不可為負數"

    unit_price_val = float(unit_price) if not _is_empty_number(unit_price) else None
    # Amount = Cash Warehouse: keep signed value; if left at 0 but unit price set, auto total
    if _is_empty_number(amount):
        amount_val = 0.0
    else:
        amount_val = float(amount)
    if abs(amount_val) < 0.000001 and unit_price_val is not None and abs(unit_price_val) > 0:
        amount_val = round(unit_price_val * gram, 2)

    items.append({
        "item_type": resolved_item,
        "quality": resolved_quality,
        "weight_gram": gram,
        "weight_tael": tael,
        "weight_oz": oz,
        "unit_price": unit_price_val,
        "amount": amount_val,
        "unit": unit,
    })
    return json.dumps(items, ensure_ascii=False), _items_to_table(items), ""


def remove_last_item(items_json):
    items = parse_line_items_json(items_json)
    if items:
        items.pop()
    return json.dumps(items, ensure_ascii=False), _items_to_table(items)


def clear_items():
    return "[]", _items_to_table([])


def _items_to_table(items):
    columns = ["貨品", "成色", "重量(克)", "重量(両)", "重量(安士)", "單價", "金額"]
    if not items:
        return pd.DataFrame(columns=columns)
    rows = []
    for it in items:
        rows.append({
            "貨品": it.get("item_type", ""),
            "成色": it.get("quality", ""),
            "重量(克)": it.get("weight_gram", ""),
            "重量(両)": "" if it.get("weight_tael") is None else it.get("weight_tael"),
            "重量(安士)": "" if it.get("weight_oz") is None else it.get("weight_oz"),
            "單價": it.get("unit_price", ""),
            "金額": it.get("amount", ""),
        })
    return pd.DataFrame(rows)


def show_exchange_section(tx_type):
    if not tx_type:
        return gr.update(visible=False)
    cfg = TRANSACTION_TYPES.get(tx_type, {})
    return gr.update(visible=cfg.get("has_exchange", False))


def auto_generate_invoice_no(tx_type, tx_date):
    if not tx_type:
        return gr.update(value=""), _invoice_no_banner("")
    invoice_no = get_next_invoice_number(session, tx_type, tx_date)
    return gr.update(value=invoice_no), _invoice_no_banner(invoice_no, tx_type)


def _invoice_no_banner(invoice_no, tx_type=None):
    if not invoice_no:
        return (
            "### 📋 請先選擇交易性質\n"
            "系統將依規則自動分配單號：`前綴 + 年份 + 月份 + 流水號`\n"
            "例：`S260300001`（2026年3月第1張銷售單）"
        )
    prefix = TRANSACTION_TYPES.get(tx_type, {}).get("prefix", invoice_no[0])
    type_name = tx_type or ""
    return (
        f"## ✅ 系統已為您分配單號\n\n"
        f"# `{invoice_no}`\n\n"
        f"類型：**{type_name}**　｜　格式：`{prefix}` + 年份(2位) + 月份(2位) + 流水號(5位)"
    )


def _amount_fields_visibility(tx_type):
    has_amount = TRANSACTION_TYPES.get(tx_type, {}).get("has_amount", True) if tx_type else True
    vis = gr.update(visible=has_amount)
    return vis, vis


def _calculate_invoice_total(main_items, note_amount, has_amount):
    """Sum Cash Warehouse (Amount) values; default 0. May be positive or negative."""
    if not has_amount:
        return 0.0
    total = 0.0
    for it in main_items:
        total += float(it.get("amount") or 0)
    if not _is_empty_number(note_amount):
        total += float(note_amount)
    return total


def build_payments_json(selected_methods, amounts, invoice_currency=None):
    """現金貨幣已併入交易貨幣：所有付款方式統一使用 invoice_currency。"""
    invoice_currency = invoice_currency or DEFAULT_CASH_CURRENCY
    payments = []
    for method, amount in zip(PAYMENT_METHODS, amounts):
        if method in (selected_methods or []):
            entry = {
                "method": method,
                "amount": float(amount) if amount else 0,
                "currency": invoice_currency,
            }
            payments.append(entry)
    return json.dumps(payments, ensure_ascii=False)


def validate_payments(selected_methods, amounts, total_amount, invoice_currency):
    if total_amount is None:
        return None
    # Amount/Cash Warehouse total may be +/- ; skip payment check when net is 0
    if abs(float(total_amount)) < 0.01:
        return None
    if not selected_methods:
        return "請至少選擇一種付款方式"
    invoice_currency = invoice_currency or DEFAULT_CASH_CURRENCY
    for method in selected_methods:
        idx = PAYMENT_METHODS.index(method)
        amount = amounts[idx]
        if amount is None or float(amount) <= 0:
            return f"請為「{method}」填寫付款金額"
    payment_total = sum(float(amounts[PAYMENT_METHODS.index(m)]) for m in selected_methods)
    # Compare payment split to absolute Cash Warehouse total
    if abs(payment_total - abs(float(total_amount))) > 0.01:
        return (
            f"付款金額合計 ({format_money(payment_total, invoice_currency)}) "
            f"與現金倉金額合計 ({format_money(abs(float(total_amount)), invoice_currency)}) 不符"
        )
    return None


def _reset_payment_fields():
    return (
        [gr.update(value=[])]
        + [gr.update(value=None) for _ in PAYMENT_METHODS]
        + [gr.update(value=DEFAULT_CASH_CURRENCY)]
    )


def on_basic_info_change(tx_type, tx_date):
    if not tx_type:
        return (
            show_exchange_section(tx_type),
            gr.update(value=""),
            _invoice_no_banner(""),
            "",
            *_amount_fields_visibility(None),
        )
    invoice_no = get_next_invoice_number(session, tx_type, tx_date)
    desc = TRANSACTION_TYPES[tx_type].get("description", "")
    return (
        show_exchange_section(tx_type),
        gr.update(value=invoice_no),
        _invoice_no_banner(invoice_no, tx_type),
        f"**{tx_type}**：{desc}" if desc else "",
        *_amount_fields_visibility(tx_type),
    )


def default_inventory_action(tx_type):
    """Suggest Deposit for inbound/exchange, Withdraw for outbound."""
    direction = (TRANSACTION_TYPES.get(tx_type) or {}).get("inventory_direction")
    if direction in ("in", "exchange"):
        return INVENTORY_ACTION_DEPOSIT
    if direction == "out":
        return INVENTORY_ACTION_WITHDRAW
    return None


def submit_invoice(
    tx_type, invoice_no, customer_name, customer_phone, tx_date, handler,
    selected_payment_methods,
    pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
    invoice_currency,
    inventory_action, warehouse_location,
    exchange_warehouse_location,
    notes, note_amount,
    main_items_json, exchange_items_json,
    current_user,
):
    login_err = require_permission(current_user, "create_invoice")
    if login_err:
        return login_err, None, _items_to_table([])

    if not is_admin(current_user):
        handler = current_user.get("display_name", "")

    payment_amounts = [pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other]
    if not tx_type:
        return "❌ 請選擇交易性質", None, _items_to_table([])

    customer_name = (customer_name or "").strip()
    if not customer_name:
        return "❌ 請填寫客戶姓名", None, _items_to_table([])

    customer_phone = (customer_phone or "").strip()

    tx_config = TRANSACTION_TYPES[tx_type]
    is_exchange = tx_config.get("has_exchange", False)

    if not inventory_action:
        return "❌ 請選擇 Inventory Deposit/Withdrawal（存入／取出）", None, _items_to_table([])
    if not warehouse_location:
        return "❌ 請選擇 Warehouse Location（A/B/C 倉庫）", None, _items_to_table([])
    if is_exchange and not exchange_warehouse_location:
        return "❌ 兌料單請選擇對換出倉的 Warehouse Location（B 貨品出倉）", None, _items_to_table([])

    source_location, destination_location = compose_receipt_storage(
        inventory_action, warehouse_location, tx_type,
    )
    if not source_location and not destination_location:
        return "❌ 倉存存取設定無效", None, _items_to_table([])

    # 兌料單：A（來料）入 warehouse_location；B（對換）出 exchange_warehouse_location
    exchange_source_location = ""
    exchange_destination_location = ""
    if is_exchange:
        exchange_source_location, exchange_destination_location = compose_receipt_storage(
            INVENTORY_ACTION_WITHDRAW, exchange_warehouse_location, tx_type,
        )

    invoice_no = (invoice_no or "").strip()
    if not invoice_no:
        invoice_no = get_next_invoice_number(session, tx_type, tx_date)

    validation_error = validate_invoice_number(session, invoice_no, tx_type, tx_date)
    if validation_error:
        return f"❌ {validation_error}", None, _items_to_table([])

    main_items = parse_line_items_json(main_items_json)
    exchange_items = parse_line_items_json(exchange_items_json)
    if not main_items:
        return "❌ 請至少新增一項貨品", None, _items_to_table(main_items)

    for idx, it in enumerate(main_items, start=1):
        if it.get("weight_gram") is None:
            return f"❌ 第 {idx} 項貨品缺少重量(克)", None, _items_to_table(main_items)

    if tx_config.get("has_exchange") and not exchange_items:
        return "❌ 兌料類交易請至少新增一項對換貨品", None, _items_to_table(main_items)

    for idx, it in enumerate(exchange_items, start=1):
        if it.get("weight_gram") is None:
            return f"❌ 第 {idx} 項對換貨品缺少重量(克)", None, _items_to_table(main_items)

    from invoice_number import parse_tx_date
    tx_date_parsed = parse_tx_date(tx_date)

    has_amount = tx_config.get("has_amount", True)
    total = _calculate_invoice_total(main_items, note_amount, has_amount)
    invoice_currency = invoice_currency or DEFAULT_CASH_CURRENCY

    payment_error = validate_payments(
        selected_payment_methods, payment_amounts, total, invoice_currency,
    )
    if payment_error:
        return f"❌ {payment_error}", None, _items_to_table(main_items)

    payment_json = (
        build_payments_json(
            selected_payment_methods, payment_amounts, invoice_currency,
        )
        if abs(float(total or 0)) > 0.01 else ""
    )

    note_amt = float(note_amount) if (has_amount and not _is_empty_number(note_amount)) else 0.0
    invoice_data = {
        "invoice_no": invoice_no.strip(),
        "transaction_type": tx_type,
        "customer_name": customer_name.strip(),
        "customer_phone": customer_phone,
        "transaction_date": tx_date_parsed,
        "handler": handler or "",
        "payment_method": payment_json,
        "invoice_currency": invoice_currency,
        "source_location": source_location,
        "destination_location": destination_location,
        "exchange_source_location": exchange_source_location,
        "exchange_destination_location": exchange_destination_location,
        "notes": notes or "",
        "note_amount": note_amt,
        "total_amount": total if total is not None else 0.0,
        # Cash Warehouse signed by transaction type (購入單=出倉負數, 銷售單=入倉正數等)
        "cash_warehouse_amount": signed_cash_warehouse_amount(
            total if total is not None else 0.0, tx_type
        ),
    }

    try:
        excel_path = generate_invoice_excel(
            invoice_data, main_items,
            exchange_items if exchange_items else None,
        )
        # Persist portable relative name; keep absolute path for Gradio download
        invoice_data["excel_path"] = Path(excel_path).name
        movements = build_inventory_movements(
            tx_type, main_items, exchange_items or None,
            main_source_location=source_location,
            main_destination_location=destination_location,
            exchange_source_location=exchange_source_location,
            exchange_destination_location=exchange_destination_location,
        )
        cash_movement = build_cash_movement(invoice_data)
        save_invoice(
            session, invoice_data, main_items + [
                {**it, "section": "exchange"} for it in (exchange_items or [])
            ], movements, cash_movement=cash_movement,
            created_by_user_id=current_user.get("id"),
        )
        log_audit(session, current_user, "create_invoice", "invoice", invoice_no)
        session.commit()
        session.expire_all()

        from invoice_generator import format_cash_warehouse_amount

        cash_line = ""
        if cash_movement:
            signed = cash_movement.get("signed_amount")
            if signed is None:
                signed = cash_movement["amount"] if cash_movement["direction"] == "in" else -cash_movement["amount"]
            cash_line = (
                f"現金倉：{format_cash_warehouse_amount(signed, cash_movement.get('currency') or invoice_currency)}\n"
            )
        try:
            snapshot = _inventory_snapshot_text()
        except Exception:
            logger.exception("Inventory snapshot failed after create")
            snapshot = "【倉存已寫入資料庫；請開啟發票查閱查看各倉庫總額】"
        msg = (
            f"✅ 發票已成功生成！\n"
            f"單號：{invoice_no}\n"
            f"客戶：{customer_name}\n"
            + (f"電話：{customer_phone}\n" if customer_phone else "")
            + f"Inventory：{'Deposit 存入' if inventory_action == INVENTORY_ACTION_DEPOSIT else 'Withdraw 取出'}"
            + f" Gold/Silver @ {warehouse_location}\n"
            + (
                f"倉存存取：{source_location}\n"
                if source_location else ""
            )
            + (
                f"倉存位置：{destination_location}\n"
                if destination_location else ""
            )
            + (
                (
                    (f"對換出倉：{exchange_source_location}\n" if exchange_source_location else "")
                    + (f"對換入倉：{exchange_destination_location}\n" if exchange_destination_location else "")
                )
                if is_exchange else ""
            )
            + f"Total Foreign Currency：{format_cash_warehouse_amount(invoice_data.get('cash_warehouse_amount'), invoice_currency)}\n"
            + (
                f"付款：{format_payment_method_display(payment_json)}\n"
                if payment_json else ""
            )
            + cash_line
            + f"檔案：{excel_path}\n"
            + snapshot
        )
        if invoice_data.get("print_warning"):
            msg += f"\n⚠️ {invoice_data['print_warning']}"
        return msg, excel_path, _items_to_table(main_items)
    except Exception as e:
        session.rollback()
        logger.exception("Invoice create failed for %s", invoice_no)
        err = str(e)
        if "UNIQUE" in err.upper() or "unique" in err.lower():
            return f"❌ 單號重複：{invoice_no}，請重新整理後再試", None, _items_to_table(main_items)
        return f"❌ 生成失敗：{e}", None, _items_to_table(main_items)


def submit_and_refresh(
    tx_type, invoice_no, customer_name, customer_phone, tx_date, handler,
    selected_payment_methods,
    pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
    invoice_currency,
    inventory_action, warehouse_location,
    exchange_warehouse_location,
    notes, note_amount,
    main_items_json, exchange_items_json,
    current_user,
):
    msg, excel_path, table = submit_invoice(
        tx_type, invoice_no, customer_name, customer_phone, tx_date, handler,
        selected_payment_methods,
        pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
        invoice_currency,
        inventory_action, warehouse_location,
        exchange_warehouse_location,
        notes, note_amount,
        main_items_json, exchange_items_json,
        current_user,
    )
    if msg.startswith("✅"):
        next_no = get_next_invoice_number(session, tx_type, tx_date)
        banner = _invoice_no_banner(next_no, tx_type)
        empty = _items_to_table([])
        return (
            msg, excel_path,
            gr.update(value=next_no), banner,
            "", "", "[]", empty, "[]", empty, "", None,
            *_reset_payment_fields(),
        )
    return (
        msg, excel_path,
        gr.update(value=invoice_no), _invoice_no_banner(invoice_no or "", tx_type),
        customer_name, customer_phone, main_items_json, table,
        exchange_items_json, _items_to_table(parse_line_items_json(exchange_items_json)),
        notes, note_amount,
        gr.update(), gr.update(), gr.update(), gr.update(), gr.update(), gr.update(),
    )


def start_new_invoice(tx_type, tx_date):
    """開立新發票：重新分配單號並清空表單。"""
    invoice_no = get_next_invoice_number(session, tx_type, tx_date) if tx_type else ""
    empty = _items_to_table([])
    return (
        gr.update(value=invoice_no),
        _invoice_no_banner(invoice_no, tx_type),
        "", "", empty, "[]", empty, "[]", "", None, "",
        *_reset_payment_fields(),
    )


MOVEMENT_TABLE_COLUMNS = [
    "日期", "單號", "單據狀態", "記錄類型", "性質", "方向",
    "貨品", "成色", "重量(克)", "重量(両)", "重量(安士)",
    "客戶", "倉存存取", "倉存位置",
]
STOCK_TABLE_COLUMNS = ["貨品", "成色", "結存(克)"]
REVIEW_ITEMS_COLUMNS = ["區塊", "貨品", "成色", "重量(克)", "重量(両)", "單價", "金額"]
LOOKUP_HEADER_COLUMNS = [
    "單號", "狀態", "交易性質", "日期", "客戶", "電話",
    "經手人", "貨幣", "總額", "付款方式", "倉存存取", "倉存位置", "備註",
    "取消時間", "取消操作者",
]
LOOKUP_ITEMS_COLUMNS = [
    "區塊", "貨品", "成色", "重量(克)", "重量(両)", "重量(安士)", "單價", "金額",
]
LOOKUP_MOVE_COLUMNS = [
    "類型", "方向", "倉庫", "重量(+/-克)", "金額(+/-)", "貨品", "倉存存取", "倉存位置",
]


def _empty_movements_df():
    return pd.DataFrame(columns=MOVEMENT_TABLE_COLUMNS)


def _empty_stock_df():
    return pd.DataFrame(columns=STOCK_TABLE_COLUMNS)


def _empty_review_items_df():
    return pd.DataFrame(columns=REVIEW_ITEMS_COLUMNS)


def _empty_lookup_header_df():
    return pd.DataFrame(columns=LOOKUP_HEADER_COLUMNS)


def _empty_lookup_items_df():
    return pd.DataFrame(columns=LOOKUP_ITEMS_COLUMNS)


def _empty_lookup_moves_df():
    return pd.DataFrame(columns=LOOKUP_MOVE_COLUMNS)


def _empty_cancelled_search_df():
    return pd.DataFrame(columns=ARCHIVE_HEADERS)


def _fmt_cell(value, with_time=False):
    if value is None or value == "":
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d %H:%M") if with_time else value.strftime("%Y-%m-%d")
    return value


def _refresh_db_view():
    """Force a clean read of SQLite after writes from other Gradio threads."""
    try:
        session.commit()
    except Exception:
        session.rollback()
    try:
        session.expire_all()
    except Exception:
        pass
    # Drop thread-local session so the next query opens a fresh connection
    try:
        session.remove()
    except Exception:
        pass


def _inventory_snapshot_text():
    """Short SQL-backed warehouse totals for create-success feedback."""
    from config import METAL_WAREHOUSES, SAFE_SUMMARY_CATEGORIES
    from inventory import format_gram_display, get_safe_totals, get_unassigned_metal_totals

    _refresh_db_view()
    totals = get_safe_totals(session)
    cash = get_cash_balances(session)
    lines = ["【倉存已更新 Inventory Management】"]
    for wh in METAL_WAREHOUSES:
        parts = [
            f"{cat} {format_gram_display(totals[wh][cat])}"
            for cat in SAFE_SUMMARY_CATEGORIES
            if abs(totals[wh][cat]) > 0.001
        ]
        if parts:
            lines.append(f"  {wh}：{'｜'.join(parts)}")
    if len(lines) == 1:
        lines.append("  （各倉庫金屬合計目前為 0）")
    unassigned = get_unassigned_metal_totals(session)
    if unassigned:
        ua = "｜".join(
            f"{cat} {format_gram_display(g)}" for cat, g in unassigned.items()
        )
        lines.append(f"  未指定倉庫：{ua}")
    if cash:
        cash_bits = "｜".join(f"{cur} {amt:,.2f}" for cur, amt in cash.items())
        lines.append(f"  現金倉：{cash_bits}")
    return "\n".join(lines)


def refresh_inventory_panels(current_user):
    """Display-only refresh of Inventory Management (no invoice write)."""
    err = require_view_inventory(current_user)
    if err:
        return f"<p>{err}</p>", gr.update(), gr.update()
    _refresh_db_view()
    now_text = datetime.now().strftime("目前時間：%Y年%m月%d日 %H:%M")
    totals = get_safe_totals(session)
    cash_balances = get_cash_balances(session)
    unassigned = get_unassigned_metal_totals(session)
    return (
        build_safe_summary_html(totals, now_text, cash_balances, unassigned),
        gr.update(value=load_stock()),
        gr.update(value=load_movements()),
    )


def _empty_review_payload(message=""):
    return (
        f"<p>{message}</p>" if message else "<p></p>",
        gr.update(value=_empty_stock_df()),
        gr.update(value=_empty_movements_df()),
        gr.update(choices=[], value=None),
        "",
        gr.update(value=_empty_review_items_df()),
        None,  # Excel
        message or "",
    )


def run_load_inventory_page(current_user):
    err = require_view_inventory(current_user)
    if err:
        return _empty_review_payload(err)
    return load_review_page()


def load_review_page():
    """Invoice Review steps 1–2 data + recent order numbers for step 3."""
    _refresh_db_view()
    now_text = datetime.now().strftime("目前時間：%Y年%m月%d日 %H:%M")
    totals = get_safe_totals(session)
    cash_balances = get_cash_balances(session)
    unassigned = get_unassigned_metal_totals(session)
    nos = list_recent_invoice_nos()
    stock_df = load_stock()
    move_df = load_movements()
    latest = nos[0] if nos else ""
    count_msg = (
        f"已載入 {len(nos)} 張近期發票、{len(move_df)} 筆進出倉記錄。"
        + (f" 最新單號：{latest}" if latest else "")
    )
    # Use gr.update(value=...) so Gradio always refreshes Dataframes / Dropdown
    return (
        build_safe_summary_html(totals, now_text, cash_balances, unassigned),
        gr.update(value=stock_df),
        gr.update(value=move_df),
        gr.update(choices=nos, value=(latest or None)),
        latest,
        gr.update(value=_empty_review_items_df()),
        None,  # Excel
        f"✅ {count_msg}" if nos else "⚠️ 尚無發票資料。請先在「開立發票」建立單據後再刷新。",
    )


def list_recent_invoice_nos(limit=80):
    from database import Invoice

    _refresh_db_view()
    rows = (
        session.query(Invoice.invoice_no)
        .order_by(Invoice.id.desc())
        .limit(limit)
        .all()
    )
    return [r[0] for r in rows]


def load_inventory_page():
    """Backward-compatible inventory-only payload (summary/stock/movements). """
    _refresh_db_view()
    now_text = datetime.now().strftime("目前時間：%Y年%m月%d日 %H:%M")
    totals = get_safe_totals(session)
    cash_balances = get_cash_balances(session)
    unassigned = get_unassigned_metal_totals(session)
    return (
        build_safe_summary_html(totals, now_text, cash_balances, unassigned),
        load_stock(),
        load_movements(),
    )


def load_stock():
    stock = get_current_stock(session)
    if not stock:
        return _empty_stock_df()
    return pd.DataFrame([
        {
            "貨品": s["item_type"],
            "成色": s["quality"],
            "結存(克)": s["weight_gram"],
        }
        for s in stock
    ])


def load_movements():
    from database import InventoryMovement, Invoice

    records = (
        session.query(InventoryMovement)
        .order_by(InventoryMovement.movement_date.desc(), InventoryMovement.id.desc())
        .limit(200)
        .all()
    )
    if not records:
        return _empty_movements_df()

    invoice_nos = {r.invoice_no for r in records if r.invoice_no}
    status_map = {}
    if invoice_nos:
        for inv in session.query(Invoice).filter(Invoice.invoice_no.in_(invoice_nos)).all():
            status_map[inv.invoice_no] = inv.status

    return pd.DataFrame([
        {
            "日期": r.movement_date.strftime("%Y-%m-%d") if r.movement_date else "",
            "單號": r.invoice_no,
            "單據狀態": _invoice_status_label(status_map.get(r.invoice_no, "active")),
            "記錄類型": _movement_kind_label(getattr(r, "movement_kind", "normal")),
            "性質": r.transaction_type,
            "方向": "入倉" if r.direction == "in" else "出倉",
            "貨品": r.item_type,
            "成色": r.quality,
            "重量(克)": r.weight_gram,
            "重量(両)": r.weight_tael,
            "重量(安士)": getattr(r, "weight_oz", None) or "",
            "客戶": r.customer_name,
            "倉存存取": r.source_location or "",
            "倉存位置": r.destination_location or "",
        }
        for r in records
    ])


def refresh_review_after_submit(current_user):
    """Always reload Invoice Review after create-invoice (tab may be hidden)."""
    _refresh_db_view()
    return run_load_inventory_page(current_user)


def _fill_order_no_from_movement(evt: gr.SelectData, movement_df):
    """Step 2 → Step 3: clicking a movement row fills the order number."""
    if movement_df is None or getattr(movement_df, "empty", True):
        return gr.update(), gr.update()
    try:
        row_idx = evt.index[0] if isinstance(evt.index, (list, tuple)) else evt.index
        invoice_no = str(movement_df.iloc[int(row_idx)]["單號"]).strip()
        return gr.update(), invoice_no
    except Exception:
        return gr.update(), gr.update()


def _pick_recent_invoice(selected_no):
    """Dropdown pick → fill order-number textbox."""
    return (selected_no or "").strip()


def _on_main_tabs_select(evt: gr.SelectData, current_user):
    """Reload Invoice Review whenever that tab is opened."""
    label = str(getattr(evt, "value", "") or "")
    idx = getattr(evt, "index", None)
    if (
        "Invoice Review" in label
        or "發票查閱" in label
        or idx == 1
    ):
        logger.info("Invoice Review tab selected — reloading from DB")
        return run_load_inventory_page(current_user)
    return (
        gr.update(),
        gr.update(),
        gr.update(),
        gr.update(),
        gr.update(),
        gr.update(),
        gr.update(),
        gr.update(),
    )

def run_preview_invoice(invoice_no, current_user):
    """Load invoice lines and regenerate Excel for A4 brand-paper print."""
    from database import Invoice, InvoiceLineItem
    from invoice_generator import generate_invoice_excel, resolve_invoice_excel_path
    from invoice_preview import lines_from_orm
    from cash import signed_cash_warehouse_amount
    from config import TRANSACTION_TYPES

    empty_items = _empty_review_items_df()
    err = require_view_inventory(current_user)
    if err:
        return empty_items, None, err

    invoice_no = (invoice_no or "").strip()
    if not invoice_no:
        return empty_items, None, "❌ 請先輸入或選擇單號"

    _refresh_db_view()
    invoice = session.query(Invoice).filter(Invoice.invoice_no == invoice_no).first()
    if not invoice:
        return empty_items, None, f"❌ 找不到單號「{invoice_no}」"

    lines = (
        session.query(InvoiceLineItem)
        .filter(InvoiceLineItem.invoice_id == invoice.id)
        .order_by(InvoiceLineItem.sort_order, InvoiceLineItem.id)
        .all()
    )
    line_dicts = lines_from_orm(lines)
    currency = invoice.invoice_currency or "HKD"

    exchange_source_location = ""
    exchange_destination_location = ""
    tx_cfg = TRANSACTION_TYPES.get(invoice.transaction_type or "", {})
    if tx_cfg.get("has_exchange"):
        from database import InventoryMovement
        movs = (
            session.query(InventoryMovement)
            .filter(
                InventoryMovement.invoice_no == invoice.invoice_no,
                InventoryMovement.direction == "out",
                InventoryMovement.movement_kind == "normal",
            )
            .all()
        )
        ex_line_types = {d.get("item_type", "") for d in line_dicts if d.get("section") == "exchange"}
        ex_mov = next((m for m in movs if m.item_type in ex_line_types), None) or (movs[0] if movs else None)
        if ex_mov:
            exchange_source_location = ex_mov.source_location or ""
            exchange_destination_location = ex_mov.destination_location or ""
    items_df = pd.DataFrame([
        {
            "區塊": "對換" if (line.section or "") == "exchange" else "主項",
            "貨品": line.item_type or "",
            "成色": line.quality or "",
            "重量(克)": line.weight_gram,
            "重量(両)": line.weight_tael,
            "單價": line.unit_price,
            "金額": line.amount,
        }
        for line in lines
    ]) if lines else empty_items

    print_file = None
    excel_error = None
    invoice_data = {
        "invoice_no": invoice.invoice_no,
        "transaction_type": invoice.transaction_type,
        "customer_name": invoice.customer_name or "",
        "customer_phone": invoice.customer_phone or "",
        "transaction_date": invoice.transaction_date,
        "handler": invoice.handler or "",
        "payment_method": invoice.payment_method or "",
        "invoice_currency": currency,
        "source_location": invoice.source_location or "",
        "destination_location": invoice.destination_location or "",
        "exchange_source_location": exchange_source_location,
        "exchange_destination_location": exchange_destination_location,
        "notes": invoice.notes or "",
        "note_amount": invoice.note_amount if invoice.note_amount is not None else 0,
        "total_amount": invoice.total_amount if invoice.total_amount is not None else 0,
        "cash_warehouse_amount": signed_cash_warehouse_amount(
            invoice.total_amount if invoice.total_amount is not None else 0,
            invoice.transaction_type,
        ),
    }
    try:
        main_items = [d for d in line_dicts if d.get("section") != "exchange"]
        exchange_items = [d for d in line_dicts if d.get("section") == "exchange"]
        abs_path = generate_invoice_excel(
            invoice_data, main_items, exchange_items or None
        )
        invoice.excel_path = Path(abs_path).name
        session.commit()
        print_file = abs_path
    except Exception as exc:
        logger.exception("Failed to regenerate Excel for %s", invoice_no)
        excel_error = exc
        print_file = resolve_invoice_excel_path(
            invoice.excel_path or "", invoice.invoice_no
        )

    excel_name = Path(print_file).name if print_file else ""
    if print_file and Path(print_file).exists():
        msg = f"✅ 已載入 {invoice_no} 發票 Excel（{excel_name}）。請下載後以 A4 品牌紙直接列印。"
        if invoice_data.get("print_warning"):
            msg += f"\n⚠️ {invoice_data['print_warning']}"
        if excel_error:
            msg = f"⚠️ 重新生成失敗（{excel_error}）；已提供既有 Excel：{excel_name}"
    else:
        msg = "⚠️ 找不到發票 Excel。" + (f" 生成錯誤：{excel_error}" if excel_error else "")

    log_audit(
        session,
        current_user,
        "preview_invoice",
        target_type="invoice",
        target_id=invoice_no,
        details="Invoice Review Excel for print",
    )
    try:
        session.commit()
    except Exception:
        session.rollback()

    return items_df, print_file, msg


def run_daily_report(report_date, current_user):
    return run_print_inventory_report("每日", report_date, None, None, "全部", current_user)


def run_monthly_report(year, month, current_user):
    return run_print_inventory_report("每月", None, year, month, "全部", current_user)


def run_yearly_report(year, current_user):
    return run_print_inventory_report("每年", None, year, None, "全部", current_user)


def _inventory_view_payload(period, report_date, year, month, warehouse, current_user, *, write_excel):
    empty = (
        empty_summary_metal_df(),
        empty_ledger_metal_df(),
        empty_summary_cash_df(),
        empty_ledger_cash_df(),
        None,
        None,
    )
    err = require_download_reports(current_user)
    if err:
        return err, *empty
    start, end, label = resolve_report_period(period, report_date, year, month)
    warehouse = (warehouse or "全部").strip()
    _refresh_db_view()
    if write_excel:
        inv_path, view = export_inventory_view(session, start, end, label, warehouse)
        invoice_path, _ = generate_invoice_report(session, start, end, label)
    else:
        view = build_inventory_view(session, start, end, warehouse)
        inv_path = None
        invoice_path = None
    n_metal = len(view["ledger_metal"])
    n_cash = len(view["ledger_cash"])
    range_txt = "全部紀錄（由資料庫第一筆起）" if start is None else f"{start} ～ {end}"
    action = "已匯出 Excel" if write_excel else "已載入"
    msg = (
        f"✅ {action}｜{label}｜倉庫 {warehouse}｜{range_txt}\n"
        f"金屬明細 {n_metal} 列、現金明細 {n_cash} 列"
    )
    if write_excel:
        msg += f"\n倉存：{inv_path}\n發票：{invoice_path}"
    return (
        msg,
        view["summary_metal"],
        view["ledger_metal"],
        view["summary_cash"],
        view["ledger_cash"],
        inv_path,
        invoice_path,
    )


def run_view_inventory_report(period, report_date, year, month, warehouse, current_user):
    return _inventory_view_payload(
        period, report_date, year, month, warehouse, current_user, write_excel=False,
    )


def run_print_inventory_report(period, report_date, year, month, warehouse, current_user):
    return _inventory_view_payload(
        period, report_date, year, month, warehouse, current_user, write_excel=True,
    )


def _lookup_order_movements_df(session, invoice_no):
    from database import CashMovement, InventoryMovement
    from warehouse import resolve_movement_warehouse

    rows = []
    for m in (
        session.query(InventoryMovement)
        .filter_by(invoice_no=invoice_no)
        .order_by(InventoryMovement.id)
        .all()
    ):
        signed = float(m.weight_gram or 0)
        if m.direction != "in":
            signed = -signed
        rows.append({
            "類型": "倉存",
            "方向": "入倉" if m.direction == "in" else "出倉",
            "倉庫": resolve_movement_warehouse(m) or "",
            "重量(+/-克)": round(signed, 3),
            "金額(+/-)": "",
            "貨品": m.item_type or "",
            "倉存存取": m.source_location or "",
            "倉存位置": m.destination_location or "",
        })
    for m in (
        session.query(CashMovement)
        .filter_by(invoice_no=invoice_no)
        .order_by(CashMovement.id)
        .all()
    ):
        signed = float(m.amount or 0)
        if m.direction != "in":
            signed = -signed
        rows.append({
            "類型": "現金倉",
            "方向": "存入" if m.direction == "in" else "支出",
            "倉庫": m.warehouse or "現金倉",
            "重量(+/-克)": "",
            "金額(+/-)": round(signed, 2),
            "貨品": "",
            "倉存存取": "",
            "倉存位置": "",
        })
    return pd.DataFrame(rows, columns=LOOKUP_MOVE_COLUMNS) if rows else _empty_lookup_moves_df()


def _lookup_invoice_excel(invoice):
    """Return the generated invoice Excel path; rebuild from DB if the file is missing."""
    from invoice_generator import generate_invoice_excel, resolve_invoice_excel_path
    from invoice_preview import lines_from_orm
    from database import InvoiceLineItem
    from cash import signed_cash_warehouse_amount
    from config import TRANSACTION_TYPES

    existing = resolve_invoice_excel_path(invoice.excel_path or "", invoice.invoice_no)
    if existing:
        return existing
    lines = (
        session.query(InvoiceLineItem)
        .filter(InvoiceLineItem.invoice_id == invoice.id)
        .order_by(InvoiceLineItem.sort_order, InvoiceLineItem.id)
        .all()
    )
    line_dicts = lines_from_orm(lines)
    tx_cfg = TRANSACTION_TYPES.get(invoice.transaction_type or "", {})
    exchange_source_location = ""
    exchange_destination_location = ""
    if tx_cfg.get("has_exchange"):
        from database import InventoryMovement
        movs = (
            session.query(InventoryMovement)
            .filter_by(invoice_no=invoice.invoice_no, direction="out", movement_kind="normal")
            .all()
        )
        if movs:
            exchange_source_location = movs[0].source_location or ""
            exchange_destination_location = movs[0].destination_location or ""
    invoice_data = {
        "invoice_no": invoice.invoice_no,
        "transaction_type": invoice.transaction_type,
        "customer_name": invoice.customer_name or "",
        "customer_phone": invoice.customer_phone or "",
        "transaction_date": invoice.transaction_date,
        "handler": invoice.handler or "",
        "payment_method": invoice.payment_method or "",
        "invoice_currency": invoice.invoice_currency or "HKD",
        "source_location": invoice.source_location or "",
        "destination_location": invoice.destination_location or "",
        "exchange_source_location": exchange_source_location,
        "exchange_destination_location": exchange_destination_location,
        "notes": invoice.notes or "",
        "note_amount": invoice.note_amount if invoice.note_amount is not None else 0,
        "total_amount": invoice.total_amount if invoice.total_amount is not None else 0,
        "cash_warehouse_amount": signed_cash_warehouse_amount(
            invoice.total_amount if invoice.total_amount is not None else 0,
            invoice.transaction_type,
        ),
    }
    main_items = [d for d in line_dicts if d.get("section") != "exchange"]
    exchange_items = [d for d in line_dicts if d.get("section") == "exchange"]
    abs_path = generate_invoice_excel(invoice_data, main_items, exchange_items or None)
    invoice.excel_path = Path(abs_path).name
    try:
        session.commit()
    except Exception:
        session.rollback()
    return abs_path


def run_lookup_order(invoice_no, current_user):
    """Look up a live or cancelled invoice by serial: full content + generated Excel."""
    empty_h = _empty_lookup_header_df()
    empty_i = _empty_lookup_items_df()
    empty_m = _empty_lookup_moves_df()
    err = require_download_reports(current_user)
    if err:
        return err, empty_h, empty_i, empty_m, None

    invoice_no = (invoice_no or "").strip()
    if not invoice_no:
        return "❌ 請填寫單號", empty_h, empty_i, empty_m, None

    from database import (
        CancelledInvoice,
        CashMovement,
        InventoryMovement,
        Invoice,
        InvoiceLineItem,
    )

    _refresh_db_view()
    invoice = (
        session.query(Invoice)
        .filter(Invoice.invoice_no.ilike(invoice_no))
        .first()
    )
    if invoice:
        lines = (
            session.query(InvoiceLineItem)
            .filter(InvoiceLineItem.invoice_id == invoice.id)
            .order_by(InvoiceLineItem.sort_order, InvoiceLineItem.id)
            .all()
        )
        header = pd.DataFrame([{
            "單號": invoice.invoice_no,
            "狀態": "有效",
            "交易性質": invoice.transaction_type or "",
            "日期": _fmt_cell(invoice.transaction_date),
            "客戶": invoice.customer_name or "",
            "電話": invoice.customer_phone or "",
            "經手人": invoice.handler or "",
            "貨幣": invoice.invoice_currency or "",
            "總額": invoice.total_amount if invoice.total_amount is not None else "",
            "付款方式": format_payment_method_display(invoice.payment_method or ""),
            "倉存存取": invoice.source_location or "",
            "倉存位置": invoice.destination_location or "",
            "備註": invoice.notes or "",
            "取消時間": "",
            "取消操作者": "",
        }])
        items = pd.DataFrame([
            {
                "區塊": "對換" if (line.section or "") == "exchange" else "主項",
                "貨品": line.item_type or "",
                "成色": line.quality or "",
                "重量(克)": line.weight_gram if line.weight_gram is not None else "",
                "重量(両)": line.weight_tael if line.weight_tael is not None else "",
                "重量(安士)": line.weight_oz if getattr(line, "weight_oz", None) is not None else "",
                "單價": line.unit_price if line.unit_price is not None else "",
                "金額": line.amount if line.amount is not None else "",
            }
            for line in lines
        ]) if lines else empty_i
        moves = _lookup_order_movements_df(session, invoice.invoice_no)
        cash_n = session.query(CashMovement).filter_by(invoice_no=invoice.invoice_no).count()
        inv_n = session.query(InventoryMovement).filter_by(invoice_no=invoice.invoice_no).count()
        excel_path = None
        excel_note = ""
        try:
            excel_path = _lookup_invoice_excel(invoice)
            excel_note = f"、發票 Excel：{Path(excel_path).name}"
        except Exception as exc:
            excel_note = f"、Excel 未能載入（{exc}）"
        msg = (
            f"✅ 找到有效單據 {invoice.invoice_no}\n"
            f"貨品 {len(lines)} 項、倉存記錄 {inv_n} 筆、現金記錄 {cash_n} 筆{excel_note}"
        )
        return msg, header, items, moves, excel_path

    cancelled = (
        session.query(CancelledInvoice)
        .filter(CancelledInvoice.invoice_no.ilike(invoice_no))
        .first()
    )
    archive_rows = []
    target_no = cancelled.invoice_no if cancelled else invoice_no
    for rec in load_cancelled_archive_rows():
        if str(rec.get("單號", "")).strip().lower() == str(target_no).strip().lower():
            archive_rows.append(rec)

    if not cancelled and not archive_rows:
        return f"❌ 找不到單號「{invoice_no}」（有效或已取消）", empty_h, empty_i, empty_m, None

    first = archive_rows[0] if archive_rows else {}
    shown_no = cancelled.invoice_no if cancelled else invoice_no
    header = pd.DataFrame([{
        "單號": shown_no,
        "狀態": "已取消",
        "交易性質": (
            cancelled.transaction_type if cancelled else first.get("交易性質", "")
        ) or "",
        "日期": (
            _fmt_cell(cancelled.transaction_date) if cancelled else first.get("日期", "")
        ),
        "客戶": (
            cancelled.customer_name if cancelled else first.get("客戶", "")
        ) or "",
        "電話": first.get("客戶電話", ""),
        "經手人": first.get("經手人", ""),
        "貨幣": (
            cancelled.invoice_currency if cancelled else first.get("貨幣", "")
        ) or "",
        "總額": (
            cancelled.total_amount if cancelled and cancelled.total_amount is not None
            else first.get("金額", "")
        ),
        "付款方式": first.get("付款方式", ""),
        "倉存存取": first.get("倉存存取", ""),
        "倉存位置": first.get("倉存位置", ""),
        "備註": first.get("原備註", ""),
        "取消時間": (
            _fmt_cell(cancelled.cancelled_at, with_time=True)
            if cancelled else first.get("取消時間", "")
        ),
        "取消操作者": (
            cancelled.cancelled_by if cancelled else first.get("取消操作者", "")
        ) or "",
    }])
    items = pd.DataFrame([
        {
            "區塊": rec.get("區塊", ""),
            "貨品": rec.get("貨品", ""),
            "成色": rec.get("成色", ""),
            "重量(克)": rec.get("重量(克)", ""),
            "重量(両)": rec.get("重量(両)", ""),
            "重量(安士)": rec.get("重量(安士)", ""),
            "單價": rec.get("單價", ""),
            "金額": rec.get("金額", ""),
        }
        for rec in archive_rows
    ]) if archive_rows else empty_i
    excel_path = resolve_invoice_excel_path("", shown_no)
    when = ""
    if cancelled and cancelled.cancelled_at:
        when = f"（{cancelled.cancelled_at:%Y-%m-%d %H:%M}，{cancelled.cancelled_by or ''}）"
    excel_note = f"、發票 Excel：{Path(excel_path).name}" if excel_path else "、資料庫已刪除，若檔案仍在磁碟則可下載"
    msg = f"✅ 單號 {shown_no} 已取消{when}\n存檔列數：{len(archive_rows)}{excel_note}"
    return msg, header, items, empty_m, excel_path


def run_search_cancelled_orders(query, current_user):
    empty = _empty_cancelled_search_df()
    err = require_download_reports(current_user)
    if err:
        return err, empty
    rows = search_cancelled_archive(query)
    if not rows:
        hint = ""
        if str(query or "").strip():
            hint = f"（關鍵字：{query}）"
        elif not cancelled_archive_file():
            hint = "（尚無 cancelled_orders.xlsx）"
        return f"沒有符合的取消訂單{hint}", empty
    df = pd.DataFrame(rows, columns=ARCHIVE_HEADERS)
    label = f"關鍵字「{query}」" if str(query or "").strip() else "全部"
    return f"✅ {label}：{len(rows)} 列", df


def run_download_cancelled_list(current_user):
    err = require_download_reports(current_user)
    if err:
        return err, None
    path = cancelled_archive_file()
    if not path:
        return "尚未有取消訂單記錄（cancelled_orders.xlsx 不存在）", None
    return f"✅ 取消訂單清單：{Path(path).name}", path


def build_app():
    tx_choices = list(TRANSACTION_TYPES.keys())
    empty_items = "[]"
    empty_table = _items_to_table([])

    with gr.Blocks(title="金滿堂發票系統") as demo:
        current_user = gr.State(None)

        gr.Markdown("# 金滿堂發票系統")

        with gr.Column(visible=True) as login_panel:
            gr.Markdown("### 請登入系統")
            gr.Markdown(
                "首次使用預設 Admin 帳號：`admin` / 密碼：`admin123`（登入後請立即修改密碼）"
            )
            with gr.Row():
                login_username = gr.Textbox(label="帳號", placeholder="輸入帳號")
                login_password = gr.Textbox(label="密碼", type="password", placeholder="輸入密碼")
            login_btn = gr.Button("登入", variant="primary")
            login_msg = gr.Textbox(label="登入狀態", interactive=False, lines=1)

        with gr.Column(visible=False) as main_app:
            with gr.Row():
                user_info = gr.Markdown("")
                logout_btn = gr.Button("登出", variant="secondary", scale=0, min_width=100)

            gr.Markdown("銷售人員請依照步驟填寫資料，系統將自動生成對應的 Excel 發票並更新倉存記錄。")

            with gr.Tabs() as main_tabs:
                # ── Tab 1: 開立發票 ──
                with gr.Tab("📋 開立發票") as invoice_tab:
                    invoice_no_banner = gr.Markdown(
                        value=_invoice_no_banner(""),
                        elem_id="invoice-no-banner",
                    )
                    gr.Markdown("### 步驟 1：基本資料")
                    with gr.Row():
                        tx_type = gr.Dropdown(
                            choices=tx_choices, label="交易性質 *",
                            info="選擇後系統會自動分配單號並使用對應的 Excel 分頁格式",
                        )
                        invoice_no = gr.Textbox(
                            label="單號（系統自動填寫）",
                            interactive=False,
                            info="由系統自動產生，無法手動修改。S=銷售單 P=購入單 T=兌料單 D=交收單",
                        )
                        customer_name = gr.Textbox(label="客戶姓名 *", placeholder="例：陳太")
                        customer_phone = gr.Textbox(
                            label="客戶電話",
                            placeholder="例：9123 4567",
                            info="選填；會顯示於公司單客戶姓名右側",
                        )
                    tx_desc = gr.Markdown("")
                    with gr.Row():
                        tx_date = gr.DateTime(
                            label="交易日期 *",
                            include_time=False,
                            value=datetime.now().strftime("%Y-%m-%d"),
                        )
                        handler = gr.Textbox(
                            label="經手人",
                            placeholder="經手人姓名",
                            info="員工登入後自動填入姓名且不可修改",
                        )
                    payment_section = gr.Column()
                    with payment_section:
                        invoice_currency = gr.Dropdown(
                            choices=CASH_CURRENCIES,
                            label="Currency / Total Foreign Currency *",
                            value=DEFAULT_CASH_CURRENCY,
                            info="合計 Total 外幣：寫入發票合計旁貨幣欄（HKD / USD / CNY …）",
                        )
                        payment_methods = gr.CheckboxGroup(
                            choices=PAYMENT_METHODS,
                            label="付款方式（可複選）",
                            info="可同時選擇多種付款方式，並分別填寫金額",
                        )
                        with gr.Row():
                            pay_amt_cash = gr.Number(label="現金 Cash 金額", precision=2)
                            pay_amt_transfer = gr.Number(label="轉帳 Transfer 金額", precision=2)
                        with gr.Row():
                            pay_amt_cheque = gr.Number(label="支票 Cheque 金額", precision=2)
                            pay_amt_other = gr.Number(label="其他 Other 金額", precision=2)
                    gr.Markdown(
                        "#### Receipt Inventory Deposit/Withdrawal\n"
                        "Deposit Gold/Silver → Warehouse Location　｜　"
                        "Withdraw Gold/Silver from Warehouse"
                    )
                    with gr.Row():
                        inventory_action = gr.Radio(
                            choices=INVENTORY_ACTION_CHOICES,
                            label="Inventory Deposit/Withdrawal *",
                            info="存入／取出 Gold & Silver",
                            value=INVENTORY_ACTION_DEPOSIT,
                        )
                        warehouse_location = gr.Dropdown(
                            choices=WAREHOUSE_LOCATION_CHOICES,
                            label="Warehouse Location *",
                            info="A / B / C 倉庫（Gold & Silver）；兌料單時為貨品 A（來料入倉）之倉庫",
                            value=WAREHOUSE_LOCATION_CHOICES[0],
                        )

                    gr.Markdown("### 步驟 2：新增貨品（可新增多項）")
                    with gr.Row():
                        item_type = gr.Dropdown(choices=ITEM_TYPES, label="貨品")
                        quality = gr.Textbox(
                            label="成色",
                            placeholder="例：足金 / 24K / 999.9",
                            info="手動輸入成色",
                        )
                        weight_gram = gr.Number(label="重量(克) *", precision=3, value=0)
                        weight_tael = gr.Number(label="重量(両)", precision=3, value=0)
                        weight_oz = gr.Number(label="重量(安士 oz)", precision=3, value=0)
                    with gr.Row():
                        custom_item_type = gr.Textbox(
                            label="其他貨品名稱",
                            placeholder="選擇「其他 Other」時請填寫",
                            visible=False,
                        )
                    item_add_status = gr.Textbox(label="提示", interactive=False, lines=1)
                    with gr.Row():
                        unit_price = gr.Number(label="單價（選填）", precision=2, value=0)
                        amount = gr.Number(
                            label="金額／現金倉 Amount (Cash Warehouse)",
                            precision=2,
                            value=0,
                            info="可為正數或負數；預設 0。正=入現金倉，負=出現金倉",
                        )
                        unit = gr.Dropdown(choices=UNITS, label="單位", value="克 Gram")

                    main_items_json = gr.State(empty_items)
                    with gr.Row():
                        btn_add = gr.Button("➕ 新增貨品", variant="secondary")
                        btn_remove = gr.Button("➖ 移除最後一項")
                        btn_clear = gr.Button("🗑️ 清空貨品")

                    main_items_table = gr.Dataframe(
                        label="已新增的貨品", value=empty_table, interactive=False,
                    )

                    gr.Markdown("### 步驟 3：對換貨品（僅兌料類交易）")
                    exchange_section = gr.Column(visible=False)
                    with exchange_section:
                        with gr.Row():
                            ex_item_type = gr.Dropdown(choices=ITEM_TYPES, label="對換貨品")
                            ex_quality = gr.Textbox(
                                label="對換成色",
                                placeholder="例：足金 / 24K / 999.9",
                            )
                            ex_weight_gram = gr.Number(label="對換重量(克) *", precision=3, value=0)
                            ex_weight_tael = gr.Number(label="對換重量(両)", precision=3, value=0)
                            ex_weight_oz = gr.Number(label="對換重量(安士 oz)", precision=3, value=0)
                        with gr.Row():
                            ex_custom_item_type = gr.Textbox(
                                label="其他對換貨品名稱",
                                placeholder="選擇「其他 Other」時請填寫",
                                visible=False,
                            )
                        ex_item_add_status = gr.Textbox(label="提示", interactive=False, lines=1)
                        with gr.Row():
                            ex_unit_price = gr.Number(label="對換單價（選填）", precision=2, value=0)
                            ex_amount = gr.Number(
                                label="對換金額／現金倉（可正負，預設 0）",
                                precision=2,
                                value=0,
                            )
                        exchange_items_json = gr.State(empty_items)
                        with gr.Row():
                            exchange_warehouse_location = gr.Dropdown(
                                choices=WAREHOUSE_LOCATION_CHOICES,
                                label="對換出倉 Warehouse Location（貨品 B 取出）*",
                                info="兌料單：貨品 B（對換／給客戶）由此倉庫取出",
                                value=WAREHOUSE_LOCATION_CHOICES[0],
                            )
                        with gr.Row():
                            btn_ex_add = gr.Button("➕ 新增對換貨品", variant="secondary")
                            btn_ex_clear = gr.Button("🗑️ 清空對換貨品")
                        exchange_items_table = gr.Dataframe(
                            label="已新增的對換貨品", value=empty_table, interactive=False,
                        )

                    gr.Markdown("### 步驟 4：備註與確認")
                    with gr.Row():
                        notes = gr.Textbox(
                            label="備註",
                            placeholder="例：補水費/加工費/訂金、代提純費、D26060001 待提純後回料",
                        )
                        note_amount = gr.Number(
                            label="備註金額／現金倉（可正負，預設 0）",
                            precision=2,
                            value=0,
                            visible=True,
                        )

                    with gr.Row():
                        submit_btn = gr.Button("✅ 生成發票 Excel", variant="primary", size="lg")
                        new_invoice_btn = gr.Button("📝 開立下一張發票", variant="secondary")

                    result_msg = gr.Textbox(label="結果", interactive=False, lines=5)
                    excel_download = gr.File(label="下載發票")

                    def on_tx_type_change(tx_type, tx_date):
                        base = on_basic_info_change(tx_type, tx_date)
                        action = default_inventory_action(tx_type)
                        return (*base, gr.update(value=action) if action else gr.update())

                    tx_type.input(
                        on_tx_type_change, [tx_type, tx_date],
                        [exchange_section, invoice_no, invoice_no_banner, tx_desc,
                         note_amount, payment_section, inventory_action],
                    )
                    tx_type.change(
                        on_tx_type_change, [tx_type, tx_date],
                        [exchange_section, invoice_no, invoice_no_banner, tx_desc,
                         note_amount, payment_section, inventory_action],
                    )
                    tx_date.change(
                        auto_generate_invoice_no, [tx_type, tx_date],
                        [invoice_no, invoice_no_banner],
                    )

                    item_type.change(toggle_other_item_field, item_type, custom_item_type)
                    ex_item_type.change(toggle_other_item_field, ex_item_type, ex_custom_item_type)

                    btn_add.click(
                        add_line_item,
                        [main_items_json, item_type, custom_item_type, quality,
                         weight_gram, weight_tael, weight_oz, unit_price, amount, unit],
                        [main_items_json, main_items_table, item_add_status],
                    )
                    btn_remove.click(remove_last_item, main_items_json, [main_items_json, main_items_table])
                    btn_clear.click(clear_items, outputs=[main_items_json, main_items_table])

                    btn_ex_add.click(
                        add_line_item,
                        [exchange_items_json, ex_item_type, ex_custom_item_type, ex_quality,
                         ex_weight_gram, ex_weight_tael, ex_weight_oz, ex_unit_price, ex_amount, unit],
                        [exchange_items_json, exchange_items_table, ex_item_add_status],
                    )
                    btn_ex_clear.click(clear_items, outputs=[exchange_items_json, exchange_items_table])

                    submit_event = submit_btn.click(
                        submit_and_refresh,
                        [tx_type, invoice_no, customer_name, customer_phone, tx_date, handler,
                         payment_methods, pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
                         invoice_currency,
                         inventory_action, warehouse_location,
                         exchange_warehouse_location,
                         notes, note_amount,
                         main_items_json, exchange_items_json, current_user],
                        [
                            result_msg, excel_download,
                            invoice_no, invoice_no_banner,
                            customer_name, customer_phone, main_items_json, main_items_table,
                            exchange_items_json, exchange_items_table,
                            notes, note_amount,
                            payment_methods, pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
                            invoice_currency,
                        ],
                    )
                    new_invoice_btn.click(
                        start_new_invoice, [tx_type, tx_date],
                        [
                            invoice_no, invoice_no_banner,
                            customer_name, customer_phone, main_items_table, main_items_json,
                            exchange_items_table, exchange_items_json,
                            notes, note_amount, result_msg,
                            payment_methods, pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
                            invoice_currency,
                        ],
                    )

                # ── Tab 2: Invoice Review（固定順序）──
                with gr.Tab("🔎 發票查閱 Invoice Review", visible=False) as inventory_tab:
                    gr.Markdown(
                        "依序操作：**1 倉存管理 → 2 進出倉記錄 → 3 單號 → 4 載入 Excel 列印**\n\n"
                        "在「開立發票」按 **生成發票 Excel** 後，本頁 **各倉庫總額／進出倉** 會自動從資料庫更新；"
                        "亦可按 **🔄 刷新** 或重新點開本分頁。"
                    )

                    gr.Markdown("### 1. Inventory Management（倉存管理）")
                    gr.Markdown("#### 各倉庫總額")
                    safe_summary = gr.HTML()
                    gr.Markdown("#### 目前庫存結存")
                    stock_table = gr.Dataframe(
                        label="庫存結存",
                        headers=STOCK_TABLE_COLUMNS,
                        value=_empty_stock_df(),
                        interactive=False,
                    )

                    gr.Markdown("### 2. Recent Inbound/Outbound Records（最近進出倉記錄）")
                    gr.Markdown("點選一列可自動帶入下方單號。")
                    movement_table = gr.Dataframe(
                        label="進出倉明細",
                        headers=MOVEMENT_TABLE_COLUMNS,
                        value=_empty_movements_df(),
                        interactive=True,
                    )
                    refresh_btn = gr.Button("🔄 刷新倉存與記錄（載入最新）", variant="primary")

                    gr.Markdown("### 3. Order Number（單號）")
                    with gr.Row():
                        review_invoice_pick = gr.Dropdown(
                            label="近期單號（下拉選擇）",
                            choices=[],
                            filterable=True,
                            info="選擇後會填入右側單號欄",
                        )
                        review_invoice_no = gr.Textbox(
                            label="單號 Order Number *",
                            placeholder="例：S260700001",
                            info="可手動輸入，或由上方記錄／下拉帶入",
                        )
                        preview_btn = gr.Button("4. 載入 Excel", variant="primary")

                    review_items_table = gr.Dataframe(
                        label="貨品明細",
                        interactive=False,
                        headers=REVIEW_ITEMS_COLUMNS,
                        value=_empty_review_items_df(),
                    )
                    review_msg = gr.Textbox(label="查閱狀態", interactive=False, lines=2)

                    gr.Markdown("### 4. Print（列印／下載 Excel）")
                    gr.Markdown(
                        "現場列印請下載 **Excel**，以 **A4 品牌收據紙** 直接列印（100% 比例，勿 Fit to Page）。"
                    )
                    review_print_file = gr.File(label="下載／列印發票 Excel")

                    review_preview_outputs = [
                        review_items_table,
                        review_print_file,
                        review_msg,
                    ]
                    review_outputs = [
                        safe_summary,
                        stock_table,
                        movement_table,
                        review_invoice_pick,
                        review_invoice_no,
                        review_items_table,
                        review_print_file,
                        review_msg,
                    ]
                    inventory_panel_outputs = [
                        safe_summary,
                        stock_table,
                        movement_table,
                    ]

                    # Immediately after create+Excel: refresh Inventory Management from SQL
                    submit_event.then(
                        refresh_review_after_submit,
                        [current_user],
                        outputs=review_outputs,
                    )

                    refresh_btn.click(
                        run_load_inventory_page,
                        current_user,
                        outputs=review_outputs,
                    ).then(
                        run_preview_invoice,
                        [review_invoice_no, current_user],
                        review_preview_outputs,
                    ).then(
                        refresh_inventory_panels,
                        current_user,
                        inventory_panel_outputs,
                    )
                    # Opening this tab always reloads latest DB + shows print preview of latest invoice
                    inventory_tab.select(
                        run_load_inventory_page,
                        current_user,
                        outputs=review_outputs,
                    ).then(
                        run_preview_invoice,
                        [review_invoice_no, current_user],
                        review_preview_outputs,
                    ).then(
                        refresh_inventory_panels,
                        current_user,
                        inventory_panel_outputs,
                    )
                    main_tabs.select(
                        _on_main_tabs_select,
                        current_user,
                        outputs=review_outputs,
                    )
                    movement_table.select(
                        _fill_order_no_from_movement,
                        [movement_table],
                        [review_invoice_pick, review_invoice_no],
                    )
                    review_invoice_pick.change(
                        _pick_recent_invoice,
                        review_invoice_pick,
                        review_invoice_no,
                    ).then(
                        run_preview_invoice,
                        [review_invoice_no, current_user],
                        review_preview_outputs,
                    ).then(
                        refresh_inventory_panels,
                        current_user,
                        inventory_panel_outputs,
                    )
                    preview_btn.click(
                        run_preview_invoice,
                        [review_invoice_no, current_user],
                        review_preview_outputs,
                    ).then(
                        refresh_inventory_panels,
                        current_user,
                        inventory_panel_outputs,
                    )

                # ── Tab 3: 報表 ──
                with gr.Tab("📊 報表", visible=False) as reports_tab:
                    gr.Markdown("### 倉存查閱 Inventory")
                    gr.Markdown(
                        "預設顯示**資料庫全部紀錄**（由第一筆起，含累計）。"
                        "可改為每日／每月／每年，並只看 **A／B／C 倉庫** 或現金倉。"
                        "查閱後可下載 Excel 列印。"
                    )
                    with gr.Row():
                        inv_period = gr.Radio(
                            choices=PERIOD_CHOICES,
                            value="全部",
                            label="期間 Period",
                        )
                        inv_warehouse = gr.Dropdown(
                            choices=WAREHOUSE_FILTER_CHOICES,
                            value="全部",
                            label="倉庫 Warehouse",
                            info="全部，或只看 A／B／C／現金倉",
                        )
                    with gr.Row():
                        report_date = gr.DateTime(
                            label="日期 Date（每日）",
                            include_time=False,
                            value=date.today().strftime("%Y-%m-%d"),
                            info="期間選「每日」時使用",
                        )
                        report_year = gr.Number(label="年份（每月／每年）", value=date.today().year, precision=0)
                        report_month = gr.Number(label="月份（每月）", value=date.today().month, precision=0)
                    with gr.Row():
                        view_inv_btn = gr.Button("🔎 查閱 View", variant="primary")
                        print_inv_btn = gr.Button("🖨 列印／下載 Excel", variant="primary")
                    with gr.Row():
                        daily_btn = gr.Button("📅 每日報表")
                        monthly_btn = gr.Button("📆 每月報表")
                        yearly_btn = gr.Button("📈 每年報表")
                    report_msg = gr.Textbox(label="查閱／報表結果", interactive=False, lines=4)
                    inv_summary_metal = gr.Dataframe(
                        label="倉庫進出匯總（期初／入／出／期末）",
                        value=empty_summary_metal_df(),
                        interactive=False,
                    )
                    inv_ledger_metal = gr.Dataframe(
                        label="倉庫累計明細（日期、單號、重量 +/-、累計）",
                        value=empty_ledger_metal_df(),
                        interactive=False,
                    )
                    inv_summary_cash = gr.Dataframe(
                        label="現金倉進出匯總",
                        value=empty_summary_cash_df(),
                        interactive=False,
                    )
                    inv_ledger_cash = gr.Dataframe(
                        label="現金倉累計明細（日期、單號、金額 +/-、累計）",
                        value=empty_ledger_cash_df(),
                        interactive=False,
                    )
                    report_inv_file = gr.File(label="倉存報表下載／列印")
                    report_invoice_file = gr.File(label="發票報表下載")

                    gr.Markdown("### 查詢單據 Order Lookup")
                    gr.Markdown("輸入任何單號（有效或已取消）即可查看該訂單詳情。")
                    with gr.Row():
                        lookup_invoice_no = gr.Textbox(
                            label="單號", placeholder="例：S260800001"
                        )
                        lookup_btn = gr.Button("查詢此單", variant="primary")
                    lookup_msg = gr.Textbox(label="查詢結果", interactive=False, lines=3)
                    lookup_header_table = gr.Dataframe(label="單據資料", interactive=False)
                    lookup_items_table = gr.Dataframe(label="貨品明細", interactive=False)
                    lookup_moves_table = gr.Dataframe(label="此單倉存／現金進出", interactive=False)
                    lookup_excel_file = gr.File(label="此單發票 Excel（系統生成）")

                    gr.Markdown("### 取消訂單 Cancel Order（僅 Admin）")
                    gr.Markdown(
                        "輸入單號即取消該訂單：單據會**從資料庫刪除**，完整內容存檔至 "
                        "`cancelled_orders.xlsx`，且該單號**永不重用**。"
                    )
                    with gr.Row():
                        cancel_invoice_no = gr.Textbox(
                            label="單號", placeholder="例：S260800001"
                        )
                        cancel_btn = gr.Button("取消此單", variant="stop")
                    cancel_msg = gr.Textbox(label="取消結果", interactive=False, lines=3)

                    gr.Markdown("### 取消訂單紀錄 Cancelled Orders")
                    with gr.Row():
                        cancelled_search_q = gr.Textbox(
                            label="搜尋取消訂單",
                            placeholder="單號、客戶、貨品、經手人…（空白=全部）",
                        )
                        cancelled_search_btn = gr.Button("搜尋取消訂單")
                        cancelled_download_btn = gr.Button("下載取消清單", variant="primary")
                    cancelled_search_msg = gr.Textbox(label="搜尋結果", interactive=False, lines=2)
                    cancelled_search_table = gr.Dataframe(label="取消訂單明細", interactive=False)
                    cancelled_list_file = gr.File(label="取消訂單清單下載")

                    inv_view_outputs = [
                        report_msg,
                        inv_summary_metal,
                        inv_ledger_metal,
                        inv_summary_cash,
                        inv_ledger_cash,
                        report_inv_file,
                        report_invoice_file,
                    ]
                    inv_view_inputs = [
                        inv_period, report_date, report_year, report_month,
                        inv_warehouse, current_user,
                    ]
                    view_inv_btn.click(
                        run_view_inventory_report, inv_view_inputs, inv_view_outputs,
                    )
                    print_inv_btn.click(
                        run_print_inventory_report, inv_view_inputs, inv_view_outputs,
                    )
                    daily_btn.click(
                        run_daily_report, [report_date, current_user],
                        outputs=inv_view_outputs,
                    )
                    monthly_btn.click(
                        run_monthly_report, [report_year, report_month, current_user],
                        outputs=inv_view_outputs,
                    )
                    yearly_btn.click(
                        run_yearly_report, [report_year, current_user],
                        outputs=inv_view_outputs,
                    )
                    lookup_btn.click(
                        run_lookup_order,
                        [lookup_invoice_no, current_user],
                        [
                            lookup_msg, lookup_header_table, lookup_items_table,
                            lookup_moves_table, lookup_excel_file,
                        ],
                    )
                    cancel_btn.click(
                        run_cancel_invoice,
                        [cancel_invoice_no, current_user],
                        [cancel_msg, movement_table, cancelled_list_file],
                    )
                    cancelled_search_btn.click(
                        run_search_cancelled_orders,
                        [cancelled_search_q, current_user],
                        [cancelled_search_msg, cancelled_search_table],
                    )
                    cancelled_download_btn.click(
                        run_download_cancelled_list,
                        [current_user],
                        [cancelled_search_msg, cancelled_list_file],
                    )

                # ── Tab 4: Admin ──
                with gr.Tab("🔐 Admin 管理", visible=False) as admin_tab:
                    gr.Markdown("### 員工帳號管理")
                    gr.Markdown("請填寫 **帳號、姓名、初始密碼** 三項後按建立。建立成功後下方列表會自動更新。")
                    with gr.Row():
                        new_username = gr.Textbox(label="新帳號 *", placeholder="例：staff01")
                        new_display_name = gr.Textbox(label="姓名 *", placeholder="例：陳大明")
                        new_password = gr.Textbox(
                            label="初始密碼 *",
                            type="password",
                            placeholder="至少 4 個字元",
                        )
                        create_user_btn = gr.Button("建立員工帳號", variant="primary")
                    create_user_msg = gr.Textbox(label="帳號管理訊息", interactive=False, lines=2)

                    gr.Markdown("### 重設密碼 / 修改姓名")
                    with gr.Row():
                        edit_username = gr.Textbox(label="帳號")
                        edit_display_name = gr.Textbox(label="新姓名（選填）")
                        edit_password = gr.Textbox(label="新密碼（選填）", type="password")
                        update_user_btn = gr.Button("更新帳號", variant="secondary")

                    users_table = gr.Dataframe(label="使用者列表", interactive=False)
                    refresh_users_btn = gr.Button("🔄 刷新使用者列表")

                    gr.Markdown("### 稽核日誌（僅 Admin）")
                    audit_table = gr.Dataframe(label="操作記錄", interactive=False)
                    refresh_audit_btn = gr.Button("🔄 刷新稽核日誌")

                    create_user_btn.click(
                        run_create_user,
                        [new_username, new_display_name, new_password, current_user],
                        [create_user_msg, users_table, new_username, new_display_name, new_password],
                    )
                    update_user_btn.click(
                        run_update_user,
                        [edit_username, edit_display_name, edit_password, current_user],
                        [create_user_msg, users_table],
                    )
                    refresh_users_btn.click(
                        run_load_users, current_user, users_table,
                    )
                    refresh_audit_btn.click(
                        run_load_audit, current_user, audit_table,
                    )
                    admin_tab.select(
                        run_load_users, current_user, users_table,
                    )

        def _show_user_info(user):
            return _show_user_info_text(user)

        login_btn.click(
            do_login,
            [login_username, login_password],
            [
                current_user, login_msg, login_panel, main_app, handler,
                invoice_tab, inventory_tab, reports_tab, admin_tab, user_info,
                *review_outputs,
            ],
        )
        logout_btn.click(
            do_logout,
            current_user,
            [
                current_user, user_info, login_panel, main_app, handler,
                invoice_tab, inventory_tab, reports_tab, admin_tab,
            ],
        )

    return demo


if __name__ == "__main__":
    ensure_runtime_dirs()
    debug = os.environ.get("GEVIN_DEBUG", "").strip() in ("1", "true", "True", "yes")
    logger.info("Starting Gevin Metal System (debug=%s)", debug)
    app = build_app()
    port = int(os.environ.get("PORT", "7861"))
    # Bind all interfaces so LAN devices (iPad/phone/PC) can reach the Linux server.
    # strict_cors=False is required for access via http://<lan-ip>:port (not only localhost).
    # ssr_mode=False avoids Gradio Node SSR binding only to 127.0.0.1 on some hosts.
    launch_kwargs = {
        "server_name": "0.0.0.0",
        "server_port": port,
        "share": False,
        "allowed_paths": [
            str(OUTPUT_DIR),
            str(REPORT_DIR),
            str(BASE_DIR / "assets"),
            str(BASE_DIR / "templates"),
        ],
        "show_error": True,
        "strict_cors": False,
        "ssr_mode": False,
        "theme": gr.themes.Soft(),
    }
    if debug:
        launch_kwargs["debug"] = True
    try:
        app.launch(**launch_kwargs)
    except TypeError:
        # Older Gradio without newer launch kwargs
        for key in ("allowed_paths", "show_error", "strict_cors", "ssr_mode", "theme", "debug"):
            launch_kwargs.pop(key, None)
        app.launch(**launch_kwargs)
