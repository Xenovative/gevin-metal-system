"""
貴金屬加工廠 — 發票與倉存系統
逐步引導銷售人員填寫資料，自動生成 Excel 發票並記錄倉存。
"""

import json
import os
from datetime import date, datetime

import gradio as gr
import pandas as pd

from config import (
    CASH_CURRENCIES,
    CASH_PAYMENT_METHOD,
    DEFAULT_CASH_CURRENCY,
    ITEM_TYPES,
    PAYMENT_METHODS,
    QUALITY_OPTIONS,
    STORAGE_LOCATION_CHOICES,
    TRANSACTION_TYPES,
    UNITS,
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
from cash import build_cash_movement, extract_cash_payment, get_cash_balances
from database import init_db, save_invoice
from inventory import (
    build_inventory_movements,
    build_safe_summary_html,
    get_current_stock,
    get_safe_totals,
    get_unassigned_metal_totals,
)
from invoice_generator import format_payment_method_display, generate_invoice_excel
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
    user = authenticate(session, username, password)
    if not user:
        return (
            None,
            "❌ 帳號或密碼錯誤",
            gr.update(visible=True),
            gr.update(visible=False),
            _handler_field_update(None),
            *_tab_visibility_for_user(None),
            "",
        )
    log_audit(session, user, "login")
    session.commit()
    role_label = "Admin" if is_admin(user) else "員工"
    return (
        user,
        f"✅ 已登入：{user['display_name']}（{role_label}）",
        gr.update(visible=False),
        gr.update(visible=True),
        _handler_field_update(user),
        *_tab_visibility_for_user(user),
        _show_user_info_text(user),
    )


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
                  weight_gram, weight_tael, unit_price, amount, unit):
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
    if gram < 0:
        return items_json, _items_to_table(items), "❌ 重量(克)不可為負數"

    tael = None
    if not _is_empty_number(weight_tael):
        tael = float(weight_tael)
        if tael < 0:
            return items_json, _items_to_table(items), "❌ 重量(両)不可為負數"

    items.append({
        "item_type": resolved_item,
        "quality": resolved_quality,
        "weight_gram": gram,
        "weight_tael": tael,
        "unit_price": float(unit_price) if not _is_empty_number(unit_price) else None,
        "amount": float(amount) if not _is_empty_number(amount) else None,
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
    if not items:
        return pd.DataFrame(columns=["貨品", "成色", "重量(克)", "單價", "金額"])
    rows = []
    for it in items:
        rows.append({
            "貨品": it.get("item_type", ""),
            "成色": it.get("quality", ""),
            "重量(克)": it.get("weight_gram", ""),
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
    if not has_amount:
        return None
    item_amounts = [it.get("amount") for it in main_items if it.get("amount") is not None]
    has_note_amount = not _is_empty_number(note_amount)
    if not item_amounts and not has_note_amount:
        return None
    total = sum(float(a) for a in item_amounts)
    if has_note_amount:
        total += float(note_amount)
    return total


def build_payments_json(selected_methods, amounts, cash_currency=DEFAULT_CASH_CURRENCY):
    payments = []
    for method, amount in zip(PAYMENT_METHODS, amounts):
        if method in (selected_methods or []):
            entry = {
                "method": method,
                "amount": float(amount) if amount else 0,
            }
            if method == CASH_PAYMENT_METHOD:
                entry["currency"] = cash_currency or DEFAULT_CASH_CURRENCY
            payments.append(entry)
    return json.dumps(payments, ensure_ascii=False)


def validate_payments(selected_methods, amounts, total_amount, cash_currency):
    if total_amount is None:
        return None
    if not selected_methods:
        return "請至少選擇一種付款方式"
    for method in selected_methods:
        idx = PAYMENT_METHODS.index(method)
        amount = amounts[idx]
        if amount is None or float(amount) <= 0:
            return f"請為「{method}」填寫付款金額"
    if CASH_PAYMENT_METHOD in (selected_methods or []) and not cash_currency:
        return "請選擇現金貨幣"
    payment_total = sum(float(amounts[PAYMENT_METHODS.index(m)]) for m in selected_methods)
    if total_amount and abs(payment_total - float(total_amount)) > 0.01:
        return (
            f"付款金額合計 ({payment_total:,.2f}) 與發票合計 ({float(total_amount):,.2f}) 不符"
        )
    return None


def _reset_payment_fields():
    return (
        [gr.update(value=[])]
        + [gr.update(value=None) for _ in PAYMENT_METHODS]
        + [gr.update(value=DEFAULT_CASH_CURRENCY, visible=False)]
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


def submit_invoice(
    tx_type, invoice_no, customer_name, tx_date, handler,
    selected_payment_methods,
    pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
    cash_currency,
    source_location, destination_location,
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

    if not source_location:
        return "❌ 請選擇倉存存取", None, _items_to_table([])
    if not destination_location:
        return "❌ 請選擇倉存位置", None, _items_to_table([])

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

    payment_error = validate_payments(
        selected_payment_methods, payment_amounts, total, cash_currency,
    )
    if payment_error:
        return f"❌ {payment_error}", None, _items_to_table(main_items)

    payment_json = (
        build_payments_json(selected_payment_methods, payment_amounts, cash_currency)
        if total is not None else ""
    )

    invoice_data = {
        "invoice_no": invoice_no.strip(),
        "transaction_type": tx_type,
        "customer_name": customer_name.strip(),
        "transaction_date": tx_date_parsed,
        "handler": handler or "",
        "payment_method": payment_json,
        "source_location": source_location,
        "destination_location": destination_location,
        "notes": notes or "",
        "note_amount": float(note_amount) if (has_amount and not _is_empty_number(note_amount)) else 0,
        "total_amount": total,
    }

    try:
        excel_path = generate_invoice_excel(
            invoice_data, main_items,
            exchange_items if exchange_items else None,
        )
        invoice_data["excel_path"] = excel_path
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

        cash_info = extract_cash_payment(payment_json)
        msg = (
            f"✅ 發票已成功生成！\n"
            f"單號：{invoice_no}\n"
            f"客戶：{customer_name}\n"
            f"倉存存取：{source_location} → 倉存位置：{destination_location}\n"
            + (f"合計：{total:,.2f}\n" if total is not None else "")
            + (
                f"付款：{format_payment_method_display(payment_json)}\n"
                if payment_json else ""
            )
            + (
                f"現金倉：{'+' if cash_movement['direction'] == 'in' else '-'}"
                f"{cash_info['currency']} {cash_info['amount']:,.2f}\n"
                if cash_movement else ""
            )
            + f"檔案：{excel_path}"
        )
        return msg, excel_path, _items_to_table(main_items)
    except Exception as e:
        session.rollback()
        return f"❌ 生成失敗：{e}", None, _items_to_table(main_items)


def submit_and_refresh(
    tx_type, invoice_no, customer_name, tx_date, handler,
    selected_payment_methods,
    pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
    cash_currency,
    source_location, destination_location,
    notes, note_amount,
    main_items_json, exchange_items_json,
    current_user,
):
    msg, excel_path, table = submit_invoice(
        tx_type, invoice_no, customer_name, tx_date, handler,
        selected_payment_methods,
        pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
        cash_currency,
        source_location, destination_location,
        notes, note_amount,
        main_items_json, exchange_items_json,
        current_user,
    )
    if msg.startswith("✅"):
        next_no = get_next_invoice_number(session, tx_type, tx_date)
        banner = _invoice_no_banner(next_no, tx_type)
        empty = _items_to_table([])
        return (
            msg, excel_path, table,
            gr.update(value=next_no), banner,
            "", "[]", empty, "[]", empty, "", None,
            *_reset_payment_fields(),
        )
    return (
        msg, excel_path, table,
        gr.update(value=invoice_no), _invoice_no_banner(invoice_no or "", tx_type),
        customer_name, main_items_json, table,
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
        "", empty, "[]", empty, "[]", "", None, "",
        *_reset_payment_fields(),
    )


def run_load_inventory_page(current_user):
    err = require_view_inventory(current_user)
    if err:
        empty_stock = pd.DataFrame(columns=["貨品", "成色", "結存(克)"])
        return f"<p>{err}</p>", empty_stock, pd.DataFrame()
    return load_inventory_page()


def load_inventory_page():
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
        return pd.DataFrame(columns=["貨品", "成色", "結存(克)"])
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
        return pd.DataFrame()

    invoice_nos = {r.invoice_no for r in records if r.invoice_no}
    status_map = {}
    if invoice_nos:
        for inv in session.query(Invoice).filter(Invoice.invoice_no.in_(invoice_nos)).all():
            status_map[inv.invoice_no] = inv.status

    return pd.DataFrame([
        {
            "日期": r.movement_date.strftime("%Y-%m-%d"),
            "單號": r.invoice_no,
            "單據狀態": _invoice_status_label(status_map.get(r.invoice_no, "active")),
            "記錄類型": _movement_kind_label(getattr(r, "movement_kind", "normal")),
            "性質": r.transaction_type,
            "方向": "入倉" if r.direction == "in" else "出倉",
            "貨品": r.item_type,
            "成色": r.quality,
            "重量(克)": r.weight_gram,
            "客戶": r.customer_name,
            "倉存存取": r.source_location or "",
            "倉存位置": r.destination_location or "",
        }
        for r in records
    ])


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

    with gr.Blocks(title="貴金屬加工廠 — 發票與倉存系統", theme=gr.themes.Soft()) as demo:
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

            with gr.Tabs():
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
                    with gr.Row():
                        source_location = gr.Dropdown(
                            choices=STORAGE_LOCATION_CHOICES,
                            label="倉存存取 *",
                            info="存/取 A倉庫、B倉庫、C倉庫或現金倉",
                        )
                        destination_location = gr.Dropdown(
                            choices=STORAGE_LOCATION_CHOICES,
                            label="倉存位置 *",
                            info="存/取 A倉庫、B倉庫、C倉庫或現金倉（寫入公司單庫存欄）",
                        )

                    gr.Markdown("### 步驟 2：新增貨品（可新增多項）")
                    with gr.Row():
                        item_type = gr.Dropdown(choices=ITEM_TYPES, label="貨品")
                        quality = gr.Dropdown(choices=QUALITY_OPTIONS, label="成色")
                        weight_gram = gr.Number(label="重量(克) *", precision=3, value=0)
                        weight_tael = gr.Number(label="重量(両)", precision=3, value=0)
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
                        unit_price = gr.Number(label="單價（選填）", precision=2)
                        amount = gr.Number(label="金額（選填）", precision=2)
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
                            ex_unit_price = gr.Number(label="對換單價（選填）", precision=2)
                            ex_amount = gr.Number(label="對換金額（選填）", precision=2)
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
                        note_amount = gr.Number(label="備註金額", precision=2, visible=True)

                    with gr.Row():
                        submit_btn = gr.Button("✅ 生成發票 Excel", variant="primary", size="lg")
                        new_invoice_btn = gr.Button("📝 開立下一張發票", variant="secondary")

                    result_msg = gr.Textbox(label="結果", interactive=False, lines=5)
                    excel_download = gr.File(label="下載發票")

                    tx_type.input(
                        on_basic_info_change, [tx_type, tx_date],
                        [exchange_section, invoice_no, invoice_no_banner, tx_desc,
                         note_amount, payment_section],
                    )
                    tx_type.change(
                        on_basic_info_change, [tx_type, tx_date],
                        [exchange_section, invoice_no, invoice_no_banner, tx_desc,
                         note_amount, payment_section],
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
                         weight_gram, weight_tael, unit_price, amount, unit],
                        [main_items_json, main_items_table, item_add_status],
                    )
                    btn_remove.click(remove_last_item, main_items_json, [main_items_json, main_items_table])
                    btn_clear.click(clear_items, outputs=[main_items_json, main_items_table])

                    btn_ex_add.click(
                        add_line_item,
                        [exchange_items_json, ex_item_type, ex_custom_item_type, ex_quality, ex_custom_quality,
                         ex_weight_gram, ex_weight_tael, ex_unit_price, ex_amount, unit],
                        [exchange_items_json, exchange_items_table, ex_item_add_status],
                    )
                    btn_ex_clear.click(clear_items, outputs=[exchange_items_json, exchange_items_table])

                    submit_btn.click(
                        submit_and_refresh,
                        [tx_type, invoice_no, customer_name, tx_date, handler,
                         payment_methods, pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
                         cash_currency,
                         source_location, destination_location,
                         notes, note_amount,
                         main_items_json, exchange_items_json, current_user],
                        [
                            result_msg, excel_download, main_items_table,
                            invoice_no, invoice_no_banner,
                            customer_name, main_items_json, main_items_table,
                            exchange_items_json, exchange_items_table,
                            notes, note_amount,
                            payment_methods, pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
                            cash_currency,
                        ],
                    )
                    new_invoice_btn.click(
                        start_new_invoice, [tx_type, tx_date],
                        [
                            invoice_no, invoice_no_banner,
                            customer_name, main_items_table, main_items_json,
                            exchange_items_table, exchange_items_json,
                            notes, note_amount, result_msg,
                            payment_methods, pay_amt_cash, pay_amt_transfer, pay_amt_cheque, pay_amt_other,
                            cash_currency,
                        ],
                    )
                    payment_methods.change(
                        toggle_cash_currency, payment_methods, cash_currency,
                    )

                # ── Tab 2: 倉存 ──
                with gr.Tab("📦 倉存管理", visible=False) as inventory_tab:
                    gr.Markdown("### 各倉庫總額")
                    safe_summary = gr.HTML()
                    gr.Markdown("### 目前庫存結存")
                    stock_table = gr.Dataframe(label="庫存結存", interactive=False)
                    gr.Markdown("### 最近進出倉記錄")
                    movement_table = gr.Dataframe(label="進出倉明細", interactive=False)
                    refresh_btn = gr.Button("🔄 刷新倉存資料")
                    refresh_btn.click(
                        run_load_inventory_page,
                        current_user,
                        outputs=[safe_summary, stock_table, movement_table],
                    )
                    inventory_tab.select(
                        run_load_inventory_page,
                        current_user,
                        outputs=[safe_summary, stock_table, movement_table],
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
    app = build_app()
    port = int(os.environ.get("PORT", "7861"))
    app.launch(server_name="0.0.0.0", server_port=port, share=False)
