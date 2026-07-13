from config import (
    CASH_WAREHOUSE,
    METAL_WAREHOUSES,
    SAFE_SUMMARY_CATEGORIES,
    TRANSACTION_TYPES,
)
from warehouse import resolve_movement_warehouse


def normalize_item_category(item_type: str):
    """將貨品名稱歸類為倉庫總額表的品種。足金/飾金/雜金/K金 一律視為「金」。"""
    if not item_type:
        return None
    name = str(item_type)
    if any(keyword in name for keyword in ("足金", "飾金", "雜金", "K金")):
        return "金"
    if "純銀" in name or "Silver" in name:
        return "純銀"
    return None


def build_inventory_movements(transaction_type, main_items, exchange_items=None):
    """Build inventory movement records from invoice line items."""
    tx_config = TRANSACTION_TYPES[transaction_type]
    direction = tx_config["inventory_direction"]
    movements = []

    def add_items(items, dir_override=None):
        for item in items:
            gram = item.get("weight_gram")
            if gram is None:
                gram = 0
            tael = item.get("weight_tael")
            if tael is None:
                tael = 0
            oz = item.get("weight_oz")
            if oz is None:
                oz = 0
            movements.append({
                "direction": dir_override or direction,
                "item_type": item.get("item_type", ""),
                "quality": item.get("quality", ""),
                "weight_gram": gram,
                "weight_tael": tael,
                "weight_oz": oz,
            })

    if direction == "exchange":
        add_items(main_items, "in")
        if exchange_items:
            add_items(exchange_items, "out")
    elif direction == "in":
        add_items(main_items, "in")
    elif direction == "out":
        add_items(main_items, "out")

    return movements


def get_current_stock(session):
    """Calculate current stock balance from all movements."""
    from database import InventoryMovement

    movements = session.query(InventoryMovement).all()
    stock = {}

    for m in movements:
        key = (m.item_type, m.quality or "")
        if key not in stock:
            stock[key] = {"weight_gram": 0.0, "weight_tael": 0.0}

        sign = 1 if m.direction == "in" else -1
        stock[key]["weight_gram"] += sign * (m.weight_gram or 0)
        stock[key]["weight_tael"] += sign * (m.weight_tael or 0)

    return [
        {
            "item_type": k[0],
            "quality": k[1],
            "weight_gram": round(v["weight_gram"], 3),
            "weight_tael": round(v["weight_tael"], 3),
        }
        for k, v in sorted(stock.items())
        if abs(v["weight_gram"]) > 0.001 or abs(v["weight_tael"]) > 0.001
    ]


def get_safe_totals(session):
    """依倉庫與品種匯總目前庫存重量（克）。"""
    from database import InventoryMovement

    totals = {wh: {cat: 0.0 for cat in SAFE_SUMMARY_CATEGORIES} for wh in METAL_WAREHOUSES}
    movements = session.query(InventoryMovement).all()

    for m in movements:
        category = normalize_item_category(m.item_type)
        if not category:
            continue
        gram = m.weight_gram or 0
        if abs(gram) < 0.001:
            continue

        warehouse = resolve_movement_warehouse(m)
        if not warehouse:
            continue

        sign = 1 if m.direction == "in" else -1
        totals[warehouse][category] += sign * gram

    for wh in METAL_WAREHOUSES:
        for cat in SAFE_SUMMARY_CATEGORIES:
            totals[wh][cat] = round(totals[wh][cat], 3)

    return totals


def get_unassigned_metal_totals(session):
    """未指定倉庫的進出倉重量（舊發票或缺少倉存欄位）。"""
    from database import InventoryMovement

    totals = {cat: 0.0 for cat in SAFE_SUMMARY_CATEGORIES}
    for m in session.query(InventoryMovement).all():
        category = normalize_item_category(m.item_type)
        if not category:
            continue
        gram = m.weight_gram or 0
        if abs(gram) < 0.001:
            continue
        if resolve_movement_warehouse(m):
            continue
        sign = 1 if m.direction == "in" else -1
        totals[category] += sign * gram

    return {
        cat: round(amount, 3)
        for cat, amount in totals.items()
        if abs(amount) > 0.001
    }


def format_gram_display(grams):
    value = round(grams)
    return f"{value:,}克"


def build_safe_summary_html(totals, timestamp_text, cash_balances=None, unassigned_totals=None):
    """生成各倉庫總額列表 HTML，並顯示現金倉結存（按貨幣）。"""
    cash_balances = cash_balances or {}
    unassigned_totals = unassigned_totals or {}
    header_cells = "".join(
        f'<th colspan="2" style="border:1px solid #fff;padding:8px;text-align:center;">{wh}</th>'
        for wh in METAL_WAREHOUSES
    )
    body_rows = []
    for cat in SAFE_SUMMARY_CATEGORIES:
        cells = []
        for wh in METAL_WAREHOUSES:
            cells.append(
                f'<td style="border:1px solid #fff;padding:6px 10px;">{cat}</td>'
                f'<td style="border:1px solid #fff;padding:6px 10px;text-align:right;">'
                f'{format_gram_display(totals[wh][cat])}</td>'
            )
        body_rows.append(f"<tr>{''.join(cells)}</tr>")

    if cash_balances:
        cash_lines = "".join(
            f'<div style="margin-top:6px;"><strong>{currency}</strong> '
            f'<span style="font-size:18px;margin-left:8px;">{amount:,.2f}</span></div>'
            for currency, amount in cash_balances.items()
        )
    else:
        cash_lines = '<div style="color:#ccc;">暫無現金結存</div>'

    if unassigned_totals:
        unassigned_lines = "".join(
            f'<div style="margin-top:4px;">{cat}：{format_gram_display(grams)}</div>'
            for cat, grams in unassigned_totals.items()
        )
        unassigned_block = f"""
        <div style="margin-top:12px;padding:10px;border:1px dashed #888;border-radius:6px;font-size:13px;">
            <strong>未歸倉庫記錄：</strong>
            {unassigned_lines}
            <div style="color:#ccc;margin-top:6px;">
                早期發票未填寫倉存存取／倉存位置時，只會計入「庫存結存」，不計入 A/B/C 倉庫。
            </div>
        </div>
        """
    else:
        unassigned_block = ""

    return f"""
    <div style="background:#1a1a1a;color:#fff;padding:16px;border-radius:8px;font-family:sans-serif;">
        <div style="margin-bottom:12px;font-size:14px;">{timestamp_text}</div>
        <table style="width:100%;border-collapse:collapse;border:1px solid #fff;">
            <thead><tr>{header_cells}</tr></thead>
            <tbody>{''.join(body_rows)}</tbody>
        </table>
        <div style="margin-top:16px;padding:12px;border:1px solid #fff;border-radius:6px;">
            <strong>{CASH_WAREHOUSE}結存（按貨幣）：</strong>
            {cash_lines}
            <div style="font-size:12px;color:#ccc;margin-top:8px;">
                現金付款時請選擇貨幣，金額會自動記入{CASH_WAREHOUSE}
            </div>
        </div>
        {unassigned_block}
    </div>
    """
