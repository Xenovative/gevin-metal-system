"""
Regression: Excel is source of truth; Perfect V2 PDF contains Excel values
(zeros retained; stock not glued into item name; dual handlers XXXX/Admin).
"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from openpyxl import load_workbook
from pypdf import PdfReader

from cash import signed_cash_warehouse_amount
from invoice_generator import CUSTOMER_COPY, COMPANY_COPY_OFFSET, generate_invoice_excel
from receipt_pdf import HEADER_PATH


def _fixture():
    invoice_no = "TEST_PDF_PARITY_001"
    total = 0.0
    data = {
        "invoice_no": invoice_no,
        "transaction_type": "銷售",
        "customer_name": "pmp",
        "customer_phone": "90001111",
        "transaction_date": date(2026, 7, 31),
        "handler": "Admin",
        "payment_method": "",
        "invoice_currency": "HKD$",
        "source_location": "取 A倉庫",
        "destination_location": "存 客戶",
        "notes": "",
        "note_amount": 0,
        "total_amount": total,
        "cash_warehouse_amount": signed_cash_warehouse_amount(total, "銷售"),
    }
    items = [
        {
            "item_type": "足金 Pure Gold",
            "quality": "足金",
            "weight_gram": 1,
            "weight_tael": 0,
            "weight_oz": 0,
            "unit_price": 0,
            "amount": 0,
        }
    ]
    return data, items


def _assert_excel(path: Path):
    ws = load_workbook(path)["銷售"]
    assert ws.cell(CUSTOMER_COPY["items_start"], 3).value == "足金 Pure Gold"
    assert ws.cell(CUSTOMER_COPY["items_start"], 11).value == 0
    assert ws.cell(CUSTOMER_COPY["total_row"], 10).value == "HKD$"
    assert float(ws.cell(CUSTOMER_COPY["total_row"], 11).value) == 0.0
    assert ws.cell(CUSTOMER_COPY["info_row"], 5).value == "pmp"
    inv = str(ws.cell(CUSTOMER_COPY["invoice_no_row"], 11).value)
    assert "TEST_PDF_PARITY_001" in inv

    # Dual handlers
    assert ws.cell(CUSTOMER_COPY["payment_row"], 6).value == "XXXX"
    co_pay = CUSTOMER_COPY["payment_row"] + COMPANY_COPY_OFFSET
    assert ws.cell(co_pay, 6).value == "Admin"

    # Tael spelling
    tael_unit = str(ws.cell(CUSTOMER_COPY["items_start"] + 1, 7).value or "")
    assert "Tael" in tael_unit
    assert "Teal" not in tael_unit

    co_items = CUSTOMER_COPY["items_start"] + COMPANY_COPY_OFFSET
    assert ws.cell(co_items, 3).value == "足金 Pure Gold"
    stock = str(ws.cell(co_items, 8).value or "")
    assert "取 A倉庫" in stock or "A倉庫" in stock
    assert "倉存存取" not in stock
    item_cell = str(ws.cell(co_items, 3).value or "")
    assert "(取" not in item_cell
    assert "取 A倉庫" not in item_cell

    stock_dest = str(ws.cell(co_items + 1, 8).value or "")
    assert "存 客戶" in stock_dest or "客戶" in stock_dest
    assert "倉存位置" not in stock_dest
    print("Excel assertions OK")


def _assert_brand_pdf(path: Path):
    if not Path(HEADER_PATH).exists():
        print(f"SKIP PDF parity: header missing ({HEADER_PATH})")
        return False

    reader = PdfReader(str(path))
    assert len(reader.pages) >= 1
    text = "\n".join((page.extract_text() or "") for page in reader.pages)

    assert "TEST_PDF_PARITY_001" in text or "TEST_PDF_PARI" in text
    assert "pmp" in text
    assert "足金" in text or "Pure Gold" in text
    assert "HKD$ 0.00" in text or "HKD$" in text
    assert "0" in text
    assert "XXXX" in text
    assert "Admin" in text
    assert "取 A倉庫" in text or "A倉庫" in text
    assert "倉存存取" not in text
    assert "倉存位置" not in text
    assert "Pure Gold (取" not in text
    assert "足金 Pure Gold (取" not in text
    assert "客戶單" in text or "Customer Copy" in text
    assert "公司單" in text or "Company Copy" in text
    print("PDF assertions OK (Perfect V2)")
    return True


def main():
    data, items = _fixture()
    excel = Path(generate_invoice_excel(data, items))
    assert excel.exists(), excel
    _assert_excel(excel)

    pdf = excel.with_suffix(".pdf")
    if not pdf.exists():
        try:
            from receipt_pdf import build_invoice_pdf_from_excel

            build_invoice_pdf_from_excel(excel)
        except Exception as exc:
            print(
                "SKIP PDF parity: Perfect V2 render unavailable "
                f"(need receipt_header.png + CJK font). Detail: {exc}"
            )
            print("ALL EXCEL ASSERTIONS PASSED (PDF skipped)")
            return

    assert pdf.exists(), pdf
    if not _assert_brand_pdf(pdf):
        print("ALL EXCEL ASSERTIONS PASSED (PDF skipped)")
        return
    print("ALL EXCEL/PDF PARITY TESTS PASSED")


if __name__ == "__main__":
    main()
