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
    # All-time — Gold vs Silver are separate totals
    summary, ledger = build_metal_inventory_view(moves, None, None, "全部")
    a = summary[(summary["倉庫"] == "A倉庫") & (summary["品種"] == "金")].iloc[0]
    b = summary[(summary["倉庫"] == "B倉庫") & (summary["品種"] == "金")].iloc[0]
    a_silver = summary[(summary["倉庫"] == "A倉庫") & (summary["品種"] == "純銀")].iloc[0]
    assert a["入倉(克)"] == 11 and a["出倉(克)"] == 3 and a["期末(克)"] == 8
    assert b["入倉(克)"] == 5 and b["期末(克)"] == 5
    assert a_silver["期末(克)"] == 0
    a_rows = ledger[(ledger["單號"] != "") & (ledger["倉庫"] == "A倉庫") & (ledger["品種"] == "金")]
    last_a = a_rows.iloc[-1]
    assert last_a["累計(克)"] == 8

    # Day 8/19, warehouse A only: opening 10, out 3, close 7; must ignore P3 next day
    summary, ledger = build_metal_inventory_view(
        moves, date(2026, 8, 19), date(2026, 8, 19), "A倉庫",
    )
    assert set(summary["倉庫"]) == {"A倉庫"}
    assert list(summary["品種"]) == ["金", "純銀"]
    row = summary[summary["品種"] == "金"].iloc[0]
    assert row["期初(克)"] == 10
    assert row["出倉(克)"] == 3
    assert row["入倉(克)"] == 0
    assert row["期末(克)"] == 7
    assert "P3" not in set(ledger["單號"])
    assert "S1" in set(ledger["單號"])
    print("OK metal running + A filter")


def test_gold_silver_split_on_warehouse_a():
    moves = [
        _m(id=1, movement_date=date(2026, 9, 1), weight_gram=10,
           item_type="足金 Pure Gold", destination_location="存 A倉庫", invoice_no="P_GOLD"),
        _m(id=2, movement_date=date(2026, 9, 1), weight_gram=4,
           item_type="純銀 Silver", destination_location="存 A倉庫", invoice_no="P_SILVER"),
        _m(id=3, movement_date=date(2026, 9, 2), direction="out", weight_gram=1,
           item_type="純銀 Silver", source_location="取 A倉庫", invoice_no="S_SILVER"),
    ]
    summary, ledger = build_metal_inventory_view(moves, None, None, "A倉庫")
    gold = summary[summary["品種"] == "金"].iloc[0]
    silver = summary[summary["品種"] == "純銀"].iloc[0]
    assert gold["期末(克)"] == 10
    assert silver["入倉(克)"] == 4 and silver["出倉(克)"] == 1 and silver["期末(克)"] == 3
    gold_run = ledger[(ledger["品種"] == "金") & (ledger["單號"] != "")].iloc[-1]["累計(克)"]
    silver_run = ledger[(ledger["品種"] == "純銀") & (ledger["單號"] != "")].iloc[-1]["累計(克)"]
    assert gold_run == 10
    assert silver_run == 3
    from reports import filter_metal_ledger_by_category
    gold_only = filter_metal_ledger_by_category(ledger, "金")
    silver_only = filter_metal_ledger_by_category(ledger, "純銀")
    assert set(gold_only["品種"]) == {"金"}
    assert set(silver_only["品種"]) == {"純銀"}
    assert "P_GOLD" in set(gold_only["單號"])
    assert "P_SILVER" not in set(gold_only["單號"])
    assert "P_SILVER" in set(silver_only["單號"])
    assert "P_GOLD" not in set(silver_only["單號"])
    print("OK A倉庫 Gold/Silver split")


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


def test_invoice_cash_sign_and_total():
    import pandas as pd
    from reports import append_invoice_detail_totals, signed_invoice_cash_amount

    pay = '[{"method":"現金 Cash","amount":1000,"currency":"HKD"}]'
    assert signed_invoice_cash_amount(pay, "銷售單") == 1000
    assert signed_invoice_cash_amount(pay, "購入單") == -1000
    assert signed_invoice_cash_amount(pay, "交收單") == -1000
    assert signed_invoice_cash_amount("", "銷售單") == 0

    df = pd.DataFrame([
        {"單號": "S1", "現金金額": 1000, "現金貨幣": "HKD"},
        {"單號": "P1", "現金金額": -300, "現金貨幣": "HKD"},
    ])
    out = append_invoice_detail_totals(df)
    assert out.iloc[-1]["單號"] == "合計"
    assert out.iloc[-1]["現金金額"] == 700
    assert out.iloc[-1]["現金貨幣"] == "HKD"

    mixed = pd.DataFrame([
        {"單號": "S1", "現金金額": 100, "現金貨幣": "HKD"},
        {"單號": "S2", "現金金額": 50, "現金貨幣": "USD"},
    ])
    mixed_out = append_invoice_detail_totals(mixed)
    totals = mixed_out[mixed_out["單號"] == "合計"]
    assert list(totals["現金貨幣"]) == ["HKD", "USD"]
    assert list(totals["現金金額"]) == [100, 50]
    print("OK invoice cash sign + 合計")


def main():
    test_period_resolve()
    test_running_and_warehouse_filter()
    test_gold_silver_split_on_warehouse_a()
    test_delivery_deposit_and_withdraw_direction()
    test_cash_running()
    test_invoice_cash_sign_and_total()
    print("ALL INVENTORY VIEW TESTS PASSED")


if __name__ == "__main__":
    main()
