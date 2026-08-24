"""Quick smoke test for create/excel/cash/review before deploy."""
from datetime import date
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from cash import build_cash_movement, signed_cash_warehouse_amount
from config import compose_receipt_storage
from invoice_generator import generate_invoice_excel, resolve_invoice_excel_path
from openpyxl import load_workbook
import database as db
import app as m


def main():
    assert signed_cash_warehouse_amount(1000, "購入單") == -1000
    assert signed_cash_warehouse_amount(1000, "銷售單") == 1000

    src, dst = compose_receipt_storage("存", "A倉庫")
    assert src == "" and dst == "存 A倉庫"
    assert "客戶" not in src + dst
    src, dst = compose_receipt_storage("取", "B倉庫")
    assert src == "取 B倉庫" and dst == ""
    assert "客戶" not in src + dst

    data = {
        "invoice_no": "TESTFIX_P1",
        "transaction_type": "購入單",
        "customer_name": "FixTest",
        "customer_phone": "91234567",
        "transaction_date": date(2026, 7, 30),
        "handler": "Admin",
        "payment_method": "",
        "invoice_currency": "HKD",
        "source_location": "",
        "destination_location": "存 A倉庫",
        "notes": "n1",
        "note_amount": 0,
        "total_amount": 500.0,
        "cash_warehouse_amount": signed_cash_warehouse_amount(500, "購入單"),
    }
    items = [{
        "item_type": "純銀 Silver",
        "quality": "足金",
        "weight_gram": 5,
        "weight_tael": None,
        "weight_oz": None,
        "unit_price": 100,
        "amount": 500,
    }]
    path = generate_invoice_excel(data, items)
    ws = load_workbook(path)["購入單"]
    assert ws.cell(10, 11).value == -500, ws.cell(10, 11).value
    assert ws.cell(37, 11).value == -500, ws.cell(37, 11).value
    assert ws.cell(21, 11).value == -500, ws.cell(21, 11).value
    stock = str(ws.cell(37, 8).value or "")
    assert "存 A倉庫" in stock
    assert "取 客戶" not in stock
    assert "存 客戶" not in stock

    cm = build_cash_movement(data)
    assert cm["direction"] == "out" and cm["signed_amount"] == -500

    assert resolve_invoice_excel_path(Path(path).name, "TESTFIX_P1")
    assert resolve_invoice_excel_path(path, "TESTFIX_P1")

    payload = m.load_review_page()
    # safe_summary, stock, movements, pick, invoice_no, items, excel, msg
    assert len(payload) == 8, len(payload)

    text = Path(db.__file__).read_text(encoding="utf-8")
    assert text.count('"inventory_movements": [') == 1

    # ── Cancel-order flow: delete from DB, archive to Excel, block serial reuse ──
    from config import REPORT_DIR
    from inventory import build_inventory_movements
    from invoice_number import get_next_invoice_number, validate_invoice_number
    from invoice_void import cancel_invoice

    admin = {"id": 1, "username": "admin", "display_name": "Admin", "role": "admin"}
    tx_day = date(2026, 8, 19)
    cxl_no = get_next_invoice_number(m.session, "購入單", tx_day)
    blocked = validate_invoice_number(m.session, cxl_no, "購入單", tx_day)
    assert blocked is None, blocked
    cxl_data = dict(data, invoice_no=cxl_no, transaction_date=tx_day)
    db.save_invoice(
        m.session, cxl_data, items,
        build_inventory_movements("購入單", items),
        cash_movement=None, created_by_user_id=1,
    )
    m.session.commit()
    assert m.session.query(db.Invoice).filter_by(invoice_no=cxl_no).first()

    next_before = get_next_invoice_number(m.session, "購入單", date(2026, 8, 19))
    msg = cancel_invoice(m.session, cxl_no, admin)
    assert msg.startswith("✅"), msg

    # Deleted from main DB
    assert m.session.query(db.Invoice).filter_by(invoice_no=cxl_no).first() is None
    assert m.session.query(db.InventoryMovement).filter_by(invoice_no=cxl_no).count() == 0
    # Serial blocked forever
    assert m.session.query(db.CancelledInvoice).filter_by(invoice_no=cxl_no).first()
    # Archived to cancelled_orders.xlsx
    archive = REPORT_DIR / "cancelled_orders.xlsx"
    assert archive.exists()
    archived = load_workbook(archive)["取消訂單"]
    archived_nos = [row[0] for row in archived.iter_rows(min_row=2, values_only=True)]
    assert cxl_no in archived_nos
    # Next serial does not reuse the cancelled number
    next_after = get_next_invoice_number(m.session, "購入單", date(2026, 8, 19))
    assert next_after == next_before, (next_before, next_after)
    assert next_after != cxl_no
    # Double cancel rejected
    msg2 = cancel_invoice(m.session, cxl_no, admin)
    assert msg2.startswith("❌"), msg2
    print("OK cancel-order flow")

    # ── Lookup / search / download cancelled list ──
    from invoice_void import cancelled_archive_file, search_cancelled_archive

    live_no = get_next_invoice_number(m.session, "購入單", tx_day)
    live_blocked = validate_invoice_number(m.session, live_no, "購入單", tx_day)
    assert live_blocked is None, live_blocked
    live_data = dict(data, invoice_no=live_no, transaction_date=tx_day, customer_name="LiveLookup")
    db.save_invoice(
        m.session, live_data, items,
        build_inventory_movements(
            "購入單", items,
            main_source_location=live_data["source_location"],
            main_destination_location=live_data["destination_location"],
        ),
        cash_movement=None, created_by_user_id=1,
    )
    m.session.commit()
    msg, header, items, moves, excel = m.run_lookup_order(live_no, admin)
    assert msg.startswith("✅"), msg
    assert header.iloc[0]["狀態"] == "有效"
    assert header.iloc[0]["單號"] == live_no
    assert not items.empty
    assert "付款方式" in header.columns
    assert excel and Path(excel).exists(), excel
    assert not moves.empty
    assert "A倉庫" in set(moves["倉庫"].astype(str))

    msg, header, items, moves, excel = m.run_lookup_order(cxl_no, admin)
    assert "已取消" in msg and msg.startswith("✅"), msg
    assert header.iloc[0]["狀態"] == "已取消"
    assert not items.empty

    msg, header, items, moves, excel = m.run_lookup_order("NO_SUCH_SERIAL", admin)
    assert msg.startswith("❌")

    msg, df = m.run_search_cancelled_orders("", admin)
    assert msg.startswith("✅"), msg
    assert cxl_no in df["單號"].astype(str).tolist()
    msg, df = m.run_search_cancelled_orders("FixTest", admin)
    assert msg.startswith("✅"), msg
    assert cxl_no in df["單號"].astype(str).tolist()
    msg, df = m.run_search_cancelled_orders("zzzz-no-match", admin)
    assert "沒有符合" in msg

    msg, path = m.run_download_cancelled_list(admin)
    assert msg.startswith("✅"), msg
    assert path == cancelled_archive_file()
    assert search_cancelled_archive(cxl_no)
    print("OK lookup + cancelled search/download")

    from reports import build_inventory_view
    from warehouse import resolve_movement_warehouse
    from config import METAL_WAREHOUSES

    all_moves = m.session.query(db.InventoryMovement).all()
    ledger = build_inventory_view(m.session, None, None, "全部")["ledger_metal"]
    real_rows = ledger[ledger["方向"].isin(["入倉", "出倉"])]
    assert len(real_rows) == len(all_moves), (len(real_rows), len(all_moves))
    for wh in METAL_WAREHOUSES:
        db_n = sum(1 for mv in all_moves if resolve_movement_warehouse(mv) == wh)
        v = build_inventory_view(m.session, None, None, wh)
        n = int((v["ledger_metal"]["方向"].isin(["入倉", "出倉"])).sum())
        assert n == db_n, (wh, n, db_n)
    print("OK warehouse A/B/C synced with DB")

    shop = (
        m.session.query(db.Invoice)
        .filter(db.Invoice.invoice_no != live_no)
        .order_by(db.Invoice.id.desc())
        .first()
    )
    if shop:
        shop_no = shop.invoice_no
        msg, header, items, moves, excel = m.run_lookup_order(shop_no, admin)
        assert msg.startswith("✅"), msg
        assert header.iloc[0]["單號"] == shop_no
        assert excel and Path(excel).exists(), excel
        print("OK shop serial lookup", shop_no, Path(excel).name)

    live_inv = m.session.query(db.Invoice).filter_by(invoice_no=live_no).first()
    if live_inv:
        m.session.query(db.InvoiceLineItem).filter_by(invoice_id=live_inv.id).delete()
        m.session.query(db.InventoryMovement).filter_by(invoice_no=live_no).delete()
        m.session.query(db.CashMovement).filter_by(invoice_no=live_no).delete()
        m.session.delete(live_inv)
        m.session.commit()

    m.build_app()
    print("ALL SMOKE PASSED")


if __name__ == "__main__":
    main()
