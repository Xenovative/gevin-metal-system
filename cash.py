import json

from config import (
    CASH_IN_TRANSACTION_TYPES,
    CASH_OUT_TRANSACTION_TYPES,
    CASH_PAYMENT_METHOD,
    CASH_WAREHOUSE,
    DEFAULT_CASH_CURRENCY,
)


def _parse_payments(payment_method_json):
    if not payment_method_json:
        return []
    try:
        payments = json.loads(payment_method_json)
    except (json.JSONDecodeError, TypeError):
        return []
    return payments if isinstance(payments, list) else []


def extract_cash_payment(payment_method_json):
    """從付款 JSON 提取現金金額與貨幣。"""
    for entry in _parse_payments(payment_method_json):
        if entry.get("method") == CASH_PAYMENT_METHOD:
            return {
                "amount": float(entry.get("amount") or 0),
                "currency": entry.get("currency") or DEFAULT_CASH_CURRENCY,
            }
    return {"amount": 0, "currency": DEFAULT_CASH_CURRENCY}


def extract_cash_amount(payment_method_json):
    return extract_cash_payment(payment_method_json)["amount"]


def extract_cash_currency(payment_method_json):
    cash = extract_cash_payment(payment_method_json)
    return cash["currency"] if cash["amount"] > 0 else ""


def cash_movement_direction(transaction_type):
    """現金流向：收入入現金倉，支出出現金倉。"""
    if transaction_type in CASH_IN_TRANSACTION_TYPES:
        return "in"
    if transaction_type in CASH_OUT_TRANSACTION_TYPES:
        return "out"
    return None


def build_cash_movement(invoice_data):
    """若發票含現金付款，建立現金倉進出記錄。"""
    cash = extract_cash_payment(invoice_data.get("payment_method"))
    if cash["amount"] <= 0:
        return None
    direction = cash_movement_direction(invoice_data["transaction_type"])
    if not direction:
        return None
    return {
        "invoice_no": invoice_data["invoice_no"],
        "transaction_type": invoice_data["transaction_type"],
        "direction": direction,
        "amount": cash["amount"],
        "currency": cash["currency"],
        "movement_date": invoice_data["transaction_date"],
        "customer_name": invoice_data.get("customer_name", ""),
        "handler": invoice_data.get("handler", ""),
        "warehouse": CASH_WAREHOUSE,
        "notes": invoice_data.get("notes", ""),
    }


def get_cash_balances(session):
    """依貨幣計算現金倉目前結存。"""
    from database import CashMovement

    balances = {}
    for movement in session.query(CashMovement).all():
        currency = movement.currency or DEFAULT_CASH_CURRENCY
        sign = 1 if movement.direction == "in" else -1
        balances[currency] = balances.get(currency, 0.0) + sign * (movement.amount or 0)
    return {currency: round(amount, 2) for currency, amount in sorted(balances.items())}


def get_cash_balance(session, currency=DEFAULT_CASH_CURRENCY):
    """單一貨幣現金倉結存（預設港幣）。"""
    return get_cash_balances(session).get(currency, 0.0)
