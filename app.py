"""
貴金屬加工廠 — 發票與倉存系統
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
    CASH_CURRENCIES,
    CASH_PAYMENT_METHOD,
    DEFAULT_CASH_CURRENCY,
    INVENTORY_ACTION_CHOICES,
    INVENTORY_ACTION_DEPOSIT,
    INVENTORY_ACTION_WITHDRAW,
    ITEM_TYPES,
    OUTPUT_DIR,
    PAYMENT_METHODS,
    QUALITY_OPTIONS,
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
from invoice_void import void_invoice
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
from reports import daily_report, monthly_report, yearly_report

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


def run_void_invoice(invoice_no, current_user):
    result = void_invoice(session, invoice_no, current_user)
    if result.startswith("✅"):
        return result, load_movements()
    return result, gr.update()


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


def _is_other_quality(value):
    return value and str(value).strip() == "其他"


def _resolve_item_type(item_type, custom_item_type):
    if _is_other_item(item_type):
        return (custom_item_type or "").strip()
    return item_type


def _resolve_quality(quality, custom_quality):
    if _is_other_quality(quality):
        return (custom_quality or "").strip()
    return quality


def toggle_other_item_field(item_type):
    return gr.update(visible=_is_other_item(item_type))


def toggle_other_quality_field(quality):
    return gr.update(visible=_is_other_quality(quality))


def _is_empty_number(value):
    return value is None or value == ""


def add_line_item(items_json, item_type, custom_item_type, quality, custom_quality,
                  weight_gram, weight_tael, weight_oz, unit_price, amount, unit):
    items = parse_line_items_json(items_json)
    resolved_item = _resolve_item_type(item_type, custom_item_type)
    resolved_quality = _resolve_quality(quality, custom_quality)

    if _is_other_item(item_type) and not resolved_item:
        return items_json, _items_to_table(items), "❌ 選擇「其他」貨品時，請填寫貨品名稱"
    if _is_other_quality(quality) and not resolved_quality:
        return items_json, _items_to_table(items), "❌ 選擇「其他」成色時，請填寫成色內容"
    if _is_empty_number(weight_gram):
        return items_json, _items_to_table(items), "❌ 請填寫重量(克)"

    gram = float(weight_gram)
    if gram <= 0:
        return items_json, _items_to_table(items), "❌ 重量(克)必須大於 0"

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
    prefix_names = {"S": "銷售", "P": "購入", "T": "兌料", "D": "交收"}
    type_name = prefix_names.get(prefix, "")
    return (
        f"## ✅ 系統已為您分配單號\n\n"
        f"# `{invoice_no}`\n\n"
        f"類型：**{type_name}單**　｜　格式：`{prefix}` + 年份(2位) + 月份(2位) + 流水號(5位)"
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


def build_payments_json(
    selected_methods, amounts, cash_currency=DEFAULT_CASH_CURRENCY,
    invoice_currency=None,
):
    invoice_currency = invoice_currency or DEFAULT_CASH_CURRENCY
    payments = []
    for method, amount in zip(PAYMENT_METHODS, amounts):
        if method in (selected_methods or []):
            entry = {
                "method": method,
                "amount": float(amount) if amount else 0,
                "currency": (
                    cash_currency or DEFAULT_CASH_CURRENCY
                    if method == CASH_PAYMENT_METHOD
                    else invoice_currency
                ),
            }
            payments.append(entry)
    return json.dumps(payments, ensure_ascii=False)


def validate_payments(selected_methods, amounts, total_amount, cash_currency, invoice_currency):
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
    if CASH_PAYMENT_METHOD in (selected_methods or []) and not cash_currency:
        return "請選擇現金貨幣"
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
        + [
            gr.update(value=DEFAULT_CASH_CURRENCY, visible=False),
            gr.update(value=DEFAULT_CASH_CURRENCY),
        ]
    )


def toggle_cash_currency(selected_methods):
    visible = CASH_PAYMENT_METHOD in (selected_methods or [])
    return gr.update(visible=visible)


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
    cash_currency, invoice_currency,
    inventory_action, warehouse_location,
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

    if not inventory_action:
        return "❌ 請選擇 Inventory Deposit/Withdrawal（存入／取出）", None, _items_to_table([])
    if not warehouse_location:
        return "❌ 請選擇 Warehouse Location（A/B/C 倉庫）", None, _items_to_table([])

    source_location, destination_location = compose_receipt_storage(
        inventory_action, warehouse_location, tx_type,
    )
    if not source_location or not destination_location:
        return "❌ 倉存存取設定無效", None, _items_to_table([])

    invoice_no = (invoice_no or "").strip()
    if not invoice_no:
        invoice_no = get_next_invoice_number(session, tx_type, tx_date)

    validation_error = validate_invoice_number(invoice_no, tx_type, tx_date)
    if validation_error:
        return f"❌ {validation_error}", None, _items_to_table([])

    main_items = parse_line_items_json(main_items_json)
    exchange_items = parse_line_items_json(exchange_items_json)
    if not main_items:
        return "❌ 請至少新增一項貨品", None, _items_to_table(main_items)

    for idx, it in enumerate(main_items, start=1):
        if it.get("weight_gram") is None:
            return f"❌ 第 {idx} 項貨品缺少重量(克)", None, _items_to_table(main_items)

    tx_config = TRANSACTION_TYPES[tx_type]
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
        selected_payment_methods, payment_amounts, total, cash_currency, invoice_currency,
    )
    if payment_error:
        return f"❌ {payment_error}", None, _items_to_table(main_items)

    payment_json = (
        build_payments_json(
            selected_payment_methods, payment_amounts, cash_currency, invoice_currency,
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
        "notes": notes or "",
        "note_amount": note_amt,
        "total_amount": total if total is not None else 0.0,
        # Cash Warehouse signed by transaction type (購入=出倉負數, 銷售=入倉正數等)
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
        movements = build_inventory_movements(tx_type, main_items, exchange_items or None)
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
        msg = (
            f"✅ 發票已成功生成！\n"
            f"單號：{invoice_no}\n"
            f"客戶：{customer_name}\n"
            + (f"電話：{customer_phone}\n" if customer_phone else "")
            + f"Inventory：{'Deposit 存入' if inventory_action == INVENTORY_ACTION_DEPOSIT else 'Withdraw 取出'}"
            + f" Gold/Silver @ {warehouse_location}\n"
            + f"倉存存取：{source_location} → 倉存位置：{destination_location}\n"
            + f"Total Foreign Currency：{format_cash_warehouse_amount(invoice_data.get('cash_warehouse_amount'), invoice_currency)}\n"
            + (
                f"付款：{format_payment_method_display(payment_json)}\n"
                if payment_json else ""
            )
            + cash_line
            + f"檔案：{excel_path}"
        )
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
    cash_currency, invoice_currency,
    inventory_action, warehouse_location,
    notes, note_amount,
    main_items_json, exchange_items_json,
    current_user,
):
    msg, excel_path, table = submit_invoice(
        tx_type, invoice_no, customer_name, customer_phone, tx_date, handler,
        selected_payment_methods,
        pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
        cash_currency, invoice_currency,
        inventory_action, warehouse_location,
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
        gr.update(), gr.update(), gr.update(), gr.update(), gr.update(), gr.update(), gr.update(),
    )


def refresh_review_after_submit(result_msg, current_user):
    """After a successful create-invoice, reload Invoice Review tables."""
    if isinstance(result_msg, str) and result_msg.startswith("✅"):
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
        gr.update(),
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


def _empty_movements_df():
    return pd.DataFrame(columns=MOVEMENT_TABLE_COLUMNS)


def _empty_stock_df():
    return pd.DataFrame(columns=STOCK_TABLE_COLUMNS)


def _empty_review_items_df():
    return pd.DataFrame(columns=REVIEW_ITEMS_COLUMNS)


def _empty_review_payload(message=""):
    return (
        f"<p>{message}</p>" if message else "<p></p>",
        _empty_stock_df(),
        _empty_movements_df(),
        gr.update(choices=[], value=None),
        "",
        "<p>請在步驟 3 輸入或選擇單號，再按「預覽」。</p>",
        _empty_review_items_df(),
        None,
        message or "",
    )


def run_load_inventory_page(current_user):
    err = require_view_inventory(current_user)
    if err:
        return _empty_review_payload(err)
    return load_review_page()


def load_review_page():
    """Invoice Review steps 1–2 data + recent order numbers for step 3."""
    # Avoid stale SQLAlchemy identity-map reads across Gradio worker threads
    session.expire_all()
    now_text = datetime.now().strftime("目前時間：%Y年%m月%d日 %H:%M")
    totals = get_safe_totals(session)
    cash_balances = get_cash_balances(session)
    unassigned = get_unassigned_metal_totals(session)
    nos = list_recent_invoice_nos()
    count_msg = f"已載入 {len(nos)} 張近期發票、進出倉記錄已更新。"
    return (
        build_safe_summary_html(totals, now_text, cash_balances, unassigned),
        load_stock(),
        load_movements(),
        gr.update(choices=nos, value=(nos[0] if nos else None)),
        nos[0] if nos else "",
        "<p>請在步驟 3 輸入或選擇單號，再按「預覽」。</p>",
        _empty_review_items_df(),
        None,
        f"✅ {count_msg}" if nos else "⚠️ 尚無發票資料。請先在「開立發票」建立單據後再刷新。",
    )


def list_recent_invoice_nos(limit=80):
    from database import Invoice

    rows = (
        session.query(Invoice.invoice_no)
        .order_by(Invoice.id.desc())
        .limit(limit)
        .all()
    )
    return [r[0] for r in rows]


def load_inventory_page():
    """Backward-compatible inventory-only payload (summary/stock/movements). """
    session.expire_all()
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

    session.expire_all()
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


def _fill_order_no_from_movement(evt: gr.SelectData, movement_df):
    """Step 2 → Step 3: clicking a movement row fills the order number."""
    if movement_df is None or getattr(movement_df, "empty", True):
        return gr.update(), gr.update()
    try:
        row_idx = evt.index[0] if isinstance(evt.index, (list, tuple)) else evt.index
        invoice_no = str(movement_df.iloc[int(row_idx)]["單號"]).strip()
        # Update textbox for preview; dropdown value only if already in choices
        return gr.update(), invoice_no
    except Exception:
        return gr.update(), gr.update()


def _pick_recent_invoice(selected_no):
    """Dropdown pick → fill order-number textbox."""
    selected_no = (selected_no or "").strip()
    return selected_no


def _on_main_tabs_select(evt: gr.SelectData, current_user):
    """Reload Invoice Review whenever that tab is opened."""
    label = str(getattr(evt, "value", "") or "")
    if "Invoice Review" in label or "發票查閱" in label:
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
        gr.update(),
    )

def run_preview_invoice(invoice_no, current_user):
    """Steps 4–5: preview invoice details and prepare Excel for print/download."""
    from database import Invoice, InvoiceLineItem
    from invoice_generator import format_money, format_payment_method_display

    empty_items = _empty_review_items_df()
    err = require_view_inventory(current_user)
    if err:
        return f"<p>{err}</p>", empty_items, None, err

    invoice_no = (invoice_no or "").strip()
    if not invoice_no:
        return (
            "<p>❌ 請先輸入或選擇單號</p>",
            empty_items,
            None,
            "❌ 請先輸入或選擇單號",
        )

    session.expire_all()
    invoice = session.query(Invoice).filter(Invoice.invoice_no == invoice_no).first()
    if not invoice:
        return (
            f"<p>❌ 找不到單號「{invoice_no}」</p>",
            empty_items,
            None,
            f"❌ 找不到單號「{invoice_no}」",
        )

    lines = (
        session.query(InvoiceLineItem)
        .filter(InvoiceLineItem.invoice_id == invoice.id)
        .order_by(InvoiceLineItem.sort_order, InvoiceLineItem.id)
        .all()
    )
    currency = invoice.invoice_currency or "HKD$"
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

    status = _invoice_status_label(invoice.status or "active")
    payment_text = format_payment_method_display(invoice.payment_method or "")
    preview_html = f"""
    <div style="line-height:1.6">
      <h3>步驟 4 — 發票預覽</h3>
      <p><b>單號：</b>{invoice.invoice_no} &nbsp; <b>狀態：</b>{status}</p>
      <p><b>交易性質：</b>{invoice.transaction_type}
         &nbsp; <b>日期：</b>{invoice.transaction_date.strftime('%Y-%m-%d') if invoice.transaction_date else ''}</p>
      <p><b>客戶：</b>{invoice.customer_name or ''}
         &nbsp; <b>電話：</b>{invoice.customer_phone or ''}</p>
      <p><b>經手人：</b>{invoice.handler or ''}</p>
      <p><b>倉存存取：</b>{invoice.source_location or ''}
         → <b>倉存位置：</b>{invoice.destination_location or ''}</p>
      <p><b>付款方式：</b>{payment_text or '—'}</p>
      <p><b>合計：</b>{format_money(invoice.total_amount, currency) if invoice.total_amount is not None else '—'}</p>
      <p><b>備註：</b>{(invoice.notes or '—').replace(chr(10), '<br>')}</p>
    </div>
    """

    excel_path = (invoice.excel_path or "").strip()
    print_file = resolve_invoice_excel_path(excel_path, invoice.invoice_no)
    if print_file:
        msg = f"✅ 已載入 {invoice_no}。請於步驟 5 下載 Excel，用 A4 品牌收據紙列印。"
    else:
        msg = (
            f"⚠️ 已載入 {invoice_no} 預覽，但找不到 Excel 檔案"
            f"{f'（{excel_path}）' if excel_path else ''}，無法列印下載。"
        )

    log_audit(
        session,
        current_user,
        "preview_invoice",
        target_type="invoice",
        target_id=invoice_no,
        details="Invoice Review preview",
    )
    try:
        session.commit()
    except Exception:
        session.rollback()

    return preview_html, items_df, print_file, msg


def run_daily_report(current_user):
    err = require_download_reports(current_user)
    if err:
        return err, None, None
    (inv_path, inv_df), (invoice_path, invoice_df) = daily_report(session)
    return (
        f"✅ 每日報表已生成\n倉存：{inv_path}\n發票：{invoice_path}",
        inv_path,
        invoice_path,
    )


def run_monthly_report(year, month, current_user):
    err = require_download_reports(current_user)
    if err:
        return err, None, None
    year = int(year)
    month = int(month)
    (inv_path, _), (invoice_path, _) = monthly_report(session, year, month)
    return f"✅ {year}年{month:02d}月報表已生成\n倉存：{inv_path}\n發票：{invoice_path}", inv_path, invoice_path


def run_yearly_report(year, current_user):
    err = require_download_reports(current_user)
    if err:
        return err, None, None
    year = int(year)
    (inv_path, _), (invoice_path, _) = yearly_report(session, year)
    return f"✅ {year}年報表已生成\n倉存：{inv_path}\n發票：{invoice_path}", inv_path, invoice_path


def build_app():
    tx_choices = list(TRANSACTION_TYPES.keys())
    empty_items = "[]"
    empty_table = _items_to_table([])

    with gr.Blocks(title="貴金屬加工廠 — 發票與倉存系統") as demo:
        current_user = gr.State(None)

        gr.Markdown("# 貴金屬加工廠 — 發票與倉存系統")

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
                            info="由系統自動產生，無法手動修改。S=銷售 P=購入 T=兌料 D=交收",
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
                            info="合計 Total 外幣：寫入發票合計旁貨幣欄（HKD$ / USD$ / CNY¥ …）",
                        )
                        payment_methods = gr.CheckboxGroup(
                            choices=PAYMENT_METHODS,
                            label="付款方式（可複選）",
                            info="可同時選擇多種付款方式，並分別填寫金額",
                        )
                        with gr.Row():
                            pay_amt_cash = gr.Number(label="現金 Cash 金額", precision=2)
                            cash_currency = gr.Dropdown(
                                choices=CASH_CURRENCIES,
                                label="現金貨幣",
                                value=DEFAULT_CASH_CURRENCY,
                                visible=False,
                                info="選擇現金付款的交易貨幣，例如 HKD$",
                            )
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
                            info="A / B / C 倉庫（Gold & Silver）",
                            value=WAREHOUSE_LOCATION_CHOICES[0],
                        )

                    gr.Markdown("### 步驟 2：新增貨品（可新增多項）")
                    with gr.Row():
                        item_type = gr.Dropdown(choices=ITEM_TYPES, label="貨品")
                        quality = gr.Dropdown(choices=QUALITY_OPTIONS, label="成色")
                        weight_gram = gr.Number(label="重量(克) *", precision=3, value=0)
                        weight_tael = gr.Number(label="重量(両)", precision=3, value=0)
                        weight_oz = gr.Number(label="重量(安士 oz)", precision=3, value=0)
                    with gr.Row():
                        custom_item_type = gr.Textbox(
                            label="其他貨品名稱",
                            placeholder="選擇「其他 Other」時請填寫",
                            visible=False,
                        )
                        custom_quality = gr.Textbox(
                            label="其他成色",
                            placeholder="選擇「其他」時請填寫",
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
                            ex_quality = gr.Dropdown(choices=QUALITY_OPTIONS, label="對換成色")
                            ex_weight_gram = gr.Number(label="對換重量(克) *", precision=3, value=0)
                            ex_weight_tael = gr.Number(label="對換重量(両)", precision=3, value=0)
                            ex_weight_oz = gr.Number(label="對換重量(安士 oz)", precision=3, value=0)
                        with gr.Row():
                            ex_custom_item_type = gr.Textbox(
                                label="其他對換貨品名稱",
                                placeholder="選擇「其他 Other」時請填寫",
                                visible=False,
                            )
                            ex_custom_quality = gr.Textbox(
                                label="其他對換成色",
                                placeholder="選擇「其他」時請填寫",
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
                    quality.change(toggle_other_quality_field, quality, custom_quality)
                    ex_item_type.change(toggle_other_item_field, ex_item_type, ex_custom_item_type)
                    ex_quality.change(toggle_other_quality_field, ex_quality, ex_custom_quality)

                    btn_add.click(
                        add_line_item,
                        [main_items_json, item_type, custom_item_type, quality, custom_quality,
                         weight_gram, weight_tael, weight_oz, unit_price, amount, unit],
                        [main_items_json, main_items_table, item_add_status],
                    )
                    btn_remove.click(remove_last_item, main_items_json, [main_items_json, main_items_table])
                    btn_clear.click(clear_items, outputs=[main_items_json, main_items_table])

                    btn_ex_add.click(
                        add_line_item,
                        [exchange_items_json, ex_item_type, ex_custom_item_type, ex_quality, ex_custom_quality,
                         ex_weight_gram, ex_weight_tael, ex_weight_oz, ex_unit_price, ex_amount, unit],
                        [exchange_items_json, exchange_items_table, ex_item_add_status],
                    )
                    btn_ex_clear.click(clear_items, outputs=[exchange_items_json, exchange_items_table])

                    submit_event = submit_btn.click(
                        submit_and_refresh,
                        [tx_type, invoice_no, customer_name, customer_phone, tx_date, handler,
                         payment_methods, pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
                         cash_currency, invoice_currency,
                         inventory_action, warehouse_location,
                         notes, note_amount,
                         main_items_json, exchange_items_json, current_user],
                        [
                            result_msg, excel_download,
                            invoice_no, invoice_no_banner,
                            customer_name, customer_phone, main_items_json, main_items_table,
                            exchange_items_json, exchange_items_table,
                            notes, note_amount,
                            payment_methods, pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
                            cash_currency, invoice_currency,
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
                            cash_currency, invoice_currency,
                        ],
                    )
                    payment_methods.change(
                        toggle_cash_currency, payment_methods, cash_currency,
                    )

                # ── Tab 2: Invoice Review（固定順序）──
                with gr.Tab("🔎 發票查閱 Invoice Review", visible=False) as inventory_tab:
                    gr.Markdown(
                        "依序操作：**1 倉存管理 → 2 進出倉記錄 → 3 單號 → 4 預覽 → 5 列印**\n\n"
                        "開立發票後請按 **🔄 刷新**，或重新點開本分頁以載入最新資料。"
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
                    refresh_btn = gr.Button("🔄 刷新倉存與記錄", variant="secondary")

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
                        preview_btn = gr.Button("4. 預覽 Preview", variant="primary")

                    gr.Markdown("### 4. Preview（預覽）")
                    review_preview = gr.HTML(
                        value="<p>請在步驟 3 輸入或選擇單號，再按「預覽」。</p>"
                    )
                    review_items_table = gr.Dataframe(
                        label="貨品明細",
                        interactive=False,
                        headers=REVIEW_ITEMS_COLUMNS,
                        value=_empty_review_items_df(),
                    )
                    review_msg = gr.Textbox(label="查閱狀態", interactive=False, lines=2)

                    gr.Markdown("### 5. Print（列印）")
                    gr.Markdown(
                        "下載 Excel 後，以 **A4 品牌收據紙** 列印（客戶單 + 公司單）。"
                    )
                    review_print_file = gr.File(label="下載／列印發票 Excel")

                    review_outputs = [
                        safe_summary,
                        stock_table,
                        movement_table,
                        review_invoice_pick,
                        review_invoice_no,
                        review_preview,
                        review_items_table,
                        review_print_file,
                        review_msg,
                    ]
                    refresh_btn.click(
                        run_load_inventory_page,
                        current_user,
                        outputs=review_outputs,
                    )
                    # Single tab-select reload (avoid double-fire with inventory_tab.select)
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
                    )
                    preview_btn.click(
                        run_preview_invoice,
                        [review_invoice_no, current_user],
                        [review_preview, review_items_table, review_print_file, review_msg],
                    )

                    # After invoice create succeeds, refresh Invoice Review from DB
                    submit_event.then(
                        refresh_review_after_submit,
                        [result_msg, current_user],
                        outputs=review_outputs,
                    )

                # ── Tab 3: 報表 ──
                with gr.Tab("📊 報表", visible=False) as reports_tab:
                    gr.Markdown("### 生成倉存與發票報表（Excel）")
                    with gr.Row():
                        daily_btn = gr.Button("📅 每日報表", variant="primary")
                    with gr.Row():
                        report_year = gr.Number(label="年份", value=date.today().year, precision=0)
                        report_month = gr.Number(label="月份", value=date.today().month, precision=0)
                        monthly_btn = gr.Button("📆 每月報表", variant="primary")
                        yearly_btn = gr.Button("📈 每年報表", variant="primary")
                    report_msg = gr.Textbox(label="報表結果", interactive=False, lines=3)
                    report_inv_file = gr.File(label="倉存報表下載")
                    report_invoice_file = gr.File(label="發票報表下載")

                    daily_btn.click(run_daily_report, current_user, outputs=[report_msg, report_inv_file, report_invoice_file])
                    monthly_btn.click(
                        run_monthly_report, [report_year, report_month, current_user],
                        outputs=[report_msg, report_inv_file, report_invoice_file],
                    )
                    yearly_btn.click(
                        run_yearly_report, [report_year, current_user],
                        outputs=[report_msg, report_inv_file, report_invoice_file],
                    )

                # ── Tab 4: Admin ──
                with gr.Tab("🔐 Admin 管理", visible=False) as admin_tab:
                    gr.Markdown("### 作廢發票（僅 Admin）")
                    gr.Markdown("作廢後系統會自動沖銷倉存與現金倉記錄，並在備註追加作廢時間。")
                    with gr.Row():
                        void_invoice_no = gr.Textbox(label="單號", placeholder="例：S260600001")
                        void_btn = gr.Button("作廢此單", variant="stop")
                    void_msg = gr.Textbox(label="作廢結果", interactive=False, lines=2)

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

                    void_btn.click(
                        run_void_invoice,
                        [void_invoice_no, current_user],
                        [void_msg, movement_table],
                    )
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
        "allowed_paths": [str(OUTPUT_DIR), str(REPORT_DIR)],
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
