from config import METAL_WAREHOUSES, WAREHOUSE_ALIASES


def normalize_warehouse_name(name):
    if not name:
        return None
    name = str(name).strip()
    return WAREHOUSE_ALIASES.get(name, name)


def parse_storage_location(value):
    """
    解析「存 A倉庫」／「取 B倉庫」格式。
    回傳 (action, warehouse)；舊版純倉庫名回傳 (None, warehouse)。
    """
    if not value:
        return None, None
    text = str(value).strip()
    for action in ("存", "取"):
        prefix = f"{action} "
        if text.startswith(prefix):
            warehouse = normalize_warehouse_name(text[len(prefix):].strip())
            return action, warehouse
    return None, normalize_warehouse_name(text)


def metal_warehouse_from_location(location):
    """從倉存欄位解析金屬倉庫名稱（A/B/C 倉庫）。"""
    _, warehouse = parse_storage_location(location)
    if warehouse in METAL_WAREHOUSES:
        return warehouse
    return None


def resolve_movement_warehouse(movement):
    """
    依進出倉方向決定影響哪個金屬倉庫：
    - 入倉：以「倉存位置」(destination) 為準
    - 出倉：以「倉存存取」(source) 為準；若為客戶等非倉庫，則改看去向（舊資料相容）
    """
    if movement.direction == "in":
        return metal_warehouse_from_location(movement.destination_location)
    if movement.direction == "out":
        warehouse = metal_warehouse_from_location(movement.source_location)
        if warehouse:
            return warehouse
        return metal_warehouse_from_location(movement.destination_location)
    return None
