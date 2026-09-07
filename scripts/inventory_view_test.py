"""Warehouse period ledger: running totals, A/B/C filter, day range."""
from datetime import date
from types import SimpleNamespace
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from reports import (
    build_cash_inventory_view,
    build_metal_inventory_view,
    resolve_report_period,
)


def _m(**kwargs):
    defaults = dict(
        id=1,
        movement_date=date(2026, 8, 1),
        direction="in",
        weight_gram=10,
        invoice_no="S1",
        item_type="足金 Pure Gold",
        quality="999",
        transaction_type="購入單",
        customer_name="A",
        source_location="取 客戶",
        destination_location="存 A倉庫",
    )
    defaults.update(kwargs)
    return SimpleNamespace(**defaults)


def test_period_resolve():
    start, end, label = resolve_report_period("全部")
    assert start is None and end is None and label == "全部"
    start, end, label = resolve_report_period("每日", date(2026, 8, 19), 2026, 8)
    assert start == end == date(2026, 8, 19)
    start, end, _ = resolve_report_period("每月", None, 2026, 8)
    assert start == date(2026, 8, 1) and end == date(2026, 8, 31)
    start, end, _ = resolve_report_period("每年", None, 2026, 1)
    assert start == date(2026, 1, 1) and end == date(2026, 12, 31)
    print("OK period resolve")


def test_running_and_warehouse_filter():
    moves = [
        _m(id=1, movement_date=date(2026, 7, 31), weight_gram=10,
           destination_location="存 A倉庫", invoice_no="P1"),
        _m(id=2, movement_date=date(2026, 8, 19), direction="out", weight_gram=3,
           source_location="取 A倉庫", destination_location="存 客戶", invoice_no="S1"),
        _m(id=3, movement_date=date(2026, 8, 19), weight_gram=5,
           destination_location="存 B倉庫", invoice_no="P2"),
        _m(id=4, movement_date=date(2026, 8, 20), weight_gram=1,
           destination_location="存 A倉庫", invoice_no="P3"),
    ]
    # All-time
    summary, ledger = build_metal_inventory_view(moves, None, None, "全部")
    a = summary.loc[summary["倉庫"] == "A倉庫"].iloc[0]
    b = summary.loc[summary["倉庫"] == "B倉庫"].iloc[0]
    assert a["入倉(克)"] == 11 and a["出倉(克)"] == 3 and a["期末(克)"] == 8
    assert b["入倉(克)"] == 5 and b["期末(克)"] == 5
    a_rows = ledger[ledger["單號"] != ""]
    last_a = a_rows[a_rows["倉庫"] == "A倉庫"].iloc[-1]
    assert last_a["累計(克)"] == 8

    # Day 8/19, warehouse A only: opening 10, out 3, close 7; must ignore P3 next day
    summary, ledger = build_metal_inventory_view(
        moves, date(2026, 8, 19), date(2026, 8, 19), "A倉庫",
    )
    assert list(summary["倉庫"]) == ["A倉庫"]
    row = summary.iloc[0]
    assert row["期初(克)"] == 10
    assert row["出倉(克)"] == 3
    assert row["入倉(克)"] == 0
    assert row["期末(克)"] == 7
    assert "P3" not in set(ledger["單號"])
    assert "S1" in set(ledger["單號"])
    print("OK metal running + A filter")


def test_delivery_deposit_and_withdraw_direction():
    from inventory import build_inventory_movements, get_safe_totals
    from sqlalchemy import create_engine
    from sqlalchemy.orm import sessionmaker
    import tempfile
    import database as dbmod

    items = [{
        "item_type": "足金 Pure Gold",
        "quality": "999",
        "weight_gram": 10,
        "weight_tael": 0,
        "weight_oz": 0,
    }]
    deposit = build_inventory_movements(
        "交收單", items, main_destination_location="存 A倉庫",
    )
    assert len(deposit) == 1
    assert deposit[0]["direction"] == "in", deposit[0]
    assert deposit[0]["destination_location"] == "存 A倉庫"

    withdraw = build_inventory_movements(
        "交收單", items, main_source_location="取 A倉庫",
    )
    assert len(withdraw) == 1
    assert withdraw[0]["direction"] == "out", withdraw[0]
    assert withdraw[0]["source_location"] == "取 A倉庫"

    with tempfile.TemporaryDirectory() as tmp:
        engine = create_engine(
            f"sqlite:///{Path(tmp).as_posix()}/t.db",
            connect_args={"check_same_thread": False},
        )
        dbmod.Base.metadata.create_all(engine)
        session = sessionmaker(bind=engine)()
        before = get_safe_totals(session)["A倉庫"]["金"]
        session.add(
            dbmod.InventoryMovement(
                invoice_no="D_TEST_IN",
                transaction_type="交收單",
                direction=deposit[0]["direction"],
                item_type=deposit[0]["item_type"],
                quality=deposit[0]["quality"],
                weight_gram=deposit[0]["weight_gram"],
                weight_tael=0,
                weight_oz=0,
                movement_date=date(2026, 9, 7),
                source_location=deposit[0]["source_location"],
                destination_location=deposit[0]["destination_location"],
            )
        )
        session.commit()
        after = get_safe_totals(session)["A倉庫"]["金"]
        assert after == before + 10, (before, after)
        session.close()
    print("OK 交收單 Deposit in / Withdraw out + A倉庫 金 +10")


def test_cash_running():
    moves = [
        SimpleNamespace(
            id=1, movement_date=date(2026, 8, 1), direction="in", amount=100,
            currency="HKD", invoice_no="S1", transaction_type="銷售單", customer_name="x",
        ),
        SimpleNamespace(
            id=2, movement_date=date(2026, 8, 19), direction="out", amount=40,
            currency="HKD", invoice_no="P1", transaction_type="購入單", customer_name="y",
        ),
    ]
    summary, ledger = build_cash_inventory_view(
        moves, date(2026, 8, 19), date(2026, 8, 19), "現金倉",
    )
    row = summary.iloc[0]
    assert row["期初"] == 100
    assert row["支出"] == 40
    assert row["期末"] == 60
    assert ledger.iloc[-1]["累計"] == 60
    print("OK cash running")


def main():
    test_period_resolve()
    test_running_and_warehouse_filter()
    test_delivery_deposit_and_withdraw_direction()
    test_cash_running()
    print("ALL INVENTORY VIEW TESTS PASSED")


if __name__ == "__main__":
    main()
