"""Print-alignment smoke tests: Excel cells + A4 setup for all tx sheets."""
from datetime import date
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from openpyxl import load_workbook

from cash import signed_cash_warehouse_amount
from invoice_generator import (
    BRAND_CLEAR_ROWS,
    CUSTOMER_COPY,
    PRINT_AREA,
    fit_items_for_print,
    generate_invoice_excel,
)


def _base_data(invoice_no, tx_type, total, currency="HKD$"):
    return {
        "invoice_no": invoice_no,
        "transaction_type": tx_type,
        "customer_name": "對齊測試客戶",
        "customer_phone": "91234567",
        "transaction_date": date(2026, 8, 11),
        "handler": "倉管",
        "payment_method": (
            '[{"method":"現金","amount":100,"currency":"HKD$"},'
            '{"method":"轉帳","amount":50,"currency":"HKD$"}]'
        ),
        "invoice_currency": currency,
        "source_location": "取 A倉庫",
        "destination_location": "存 客戶",
        "notes": "備註測試",
        "note_amount": 0,
        "total_amount": total,
        "cash_warehouse_amount": signed_cash_warehouse_amount(total, tx_type),
    }


def _sample_item(amount=150.0):
    return {
        "item_type": "黃金 Gold",
        "quality": "999.9",
        "weight_gram": 10,
        "weight_tael": 0.267,
        "weight_oz": None,
        "unit_price": 15,
        "amount": amount,
    }


def _assert_a4(ws):
    # openpyxl may expose PAPERSIZE_A4 as str "9" while paperSize is int 9
    assert int(ws.page_setup.paperSize or 0) == int(ws.PAPERSIZE_A4)
    assert int(ws.page_setup.fitToWidth or 0) == 1
    assert int(ws.page_setup.fitToHeight or 0) == 1
    area = (ws.print_area or "").replace("$", "")
    assert "A1:K54" in area


def _assert_brand_clear(ws):
    for row in BRAND_CLEAR_ROWS:
        for col in (9, 10, 11):
            val = ws.cell(row=row, column=col).value
            if not val:
                continue
            text = str(val)
            for bad in ("Sales Invoice", "Purchase Invoice", "收據", "RECEIPT"):
                assert bad not in text, f"brand row {row} col {col}: {text}"


def _assert_copy_cells(ws, layout, invoice_no, customer, signed_total, currency, has_amount=True):
    inv = str(ws.cell(row=layout["invoice_no_row"], column=11).value or "")
    assert invoice_no in inv, inv
    assert ws.cell(row=layout["info_row"], column=5).value == customer
    assert ws.cell(row=layout["items_start"], column=3).value  # item type
    if has_amount:
        assert ws.cell(row=layout["total_row"], column=10).value == currency
        assert float(ws.cell(row=layout["total_row"], column=11).value) == float(signed_total)
        assert float(ws.cell(row=layout["items_start"], column=11).value) == float(signed_total)


def _company_layout():
    return {k: v + 27 for k, v in CUSTOMER_COPY.items()}


def test_sales():
    data = _base_data("ALIGN_S1", "銷售", 150.0)
    items = [_sample_item(150.0)]
    path = generate_invoice_excel(data, items)
    ws = load_workbook(path)["銷售"]
    _assert_a4(ws)
    _assert_brand_clear(ws)
    signed = signed_cash_warehouse_amount(150.0, "銷售")
    _assert_copy_cells(ws, CUSTOMER_COPY, "ALIGN_S1", "對齊測試客戶", signed, "HKD$")
    _assert_copy_cells(ws, _company_layout(), "ALIGN_S1", "對齊測試客戶", signed, "HKD$")
    # company stock labels
    assert ws.cell(row=_company_layout()["items_start"], column=8).value
    # payment lines capped
    assert ws.cell(row=CUSTOMER_COPY["payment_row"], column=3).value
    print("OK 銷售", path)


def test_purchase():
    data = _base_data("ALIGN_P1", "購入", 200.0)
    data["source_location"] = "取 客戶"
    data["destination_location"] = "存 A倉庫"
    items = [{**_sample_item(200.0), "amount": 200.0, "unit_price": 20}]
    path = generate_invoice_excel(data, items)
    ws = load_workbook(path)["購入單"]
    _assert_a4(ws)
    _assert_brand_clear(ws)
    signed = signed_cash_warehouse_amount(200.0, "購入")
    assert signed == -200.0
    _assert_copy_cells(ws, CUSTOMER_COPY, "ALIGN_P1", "對齊測試客戶", signed, "HKD$")
    _assert_copy_cells(ws, _company_layout(), "ALIGN_P1", "對齊測試客戶", signed, "HKD$")
    print("OK 購入", path)


def test_exchange():
    data = _base_data("ALIGN_T1", "兌料", 80.0)
    main = [_sample_item(80.0)]
    exchange = [{
        "item_type": "純銀 Silver",
        "quality": "999",
        "weight_gram": 30,
        "weight_tael": None,
        "weight_oz": None,
        "unit_price": None,
        "amount": None,
    }]
    path = generate_invoice_excel(data, main, exchange)
    ws = load_workbook(path)["兌料單"]
    _assert_a4(ws)
    _assert_brand_clear(ws)
    assert "ALIGN_T1" in str(ws.cell(row=4, column=11).value)
    # 兌料 customer notes use column E
    notes_val = ws.cell(row=CUSTOMER_COPY["notes_row"], column=5).value
    assert notes_val == "備註測試", notes_val
    print("OK 兌料", path)


def test_delivery():
    data = _base_data("ALIGN_D1", "交收去料", 0.0)
    data["payment_method"] = ""
    items = [{
        "item_type": "黃金 Gold",
        "quality": "999.9",
        "weight_gram": 5,
        "weight_tael": None,
        "weight_oz": None,
        "unit_price": None,
        "amount": None,
    }]
    path = generate_invoice_excel(data, items)
    ws = load_workbook(path)["交收單"]
    _assert_a4(ws)
    _assert_brand_clear(ws)
    assert "ALIGN_D1" in str(ws.cell(row=4, column=11).value)
    assert ws.cell(row=CUSTOMER_COPY["items_start"], column=3).value == "黃金 Gold"
    # no amount overwrite required for 交收
    print("OK 交收", path)


def test_overflow_clamp():
    """Many items must not write past notes_row."""
    data = _base_data("ALIGN_OV1", "銷售", 10.0)
    items = []
    for i in range(12):
        items.append({
            "item_type": f"Item{i}",
            "quality": "999",
            "weight_gram": 1,
            "weight_tael": 0.1,
            "weight_oz": 0.03,
            "unit_price": 1,
            "amount": 1,
        })
    from invoice_generator import estimate_items_block_rows

    fitted, _, warning = fit_items_for_print(items)
    assert warning
    max_rows = CUSTOMER_COPY["notes_row"] - CUSTOMER_COPY["items_start"]
    assert estimate_items_block_rows(fitted) <= max_rows

    path = generate_invoice_excel(data, items)
    ws = load_workbook(path)["銷售"]
    # notes row must still be notes, not an item type from overflow
    notes_cell = ws.cell(row=CUSTOMER_COPY["notes_row"], column=4).value
    assert notes_cell == "備註測試", notes_cell
    assert data.get("print_warning")
    print("OK overflow", path, "warning=", data["print_warning"])


def main():
    test_sales()
    test_purchase()
    test_exchange()
    test_delivery()
    test_overflow_clamp()
    print("ALL PRINT ALIGNMENT TESTS PASSED")


if __name__ == "__main__":
    main()
