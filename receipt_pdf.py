"""
Perfect V2 style receipt PDF: header strip + dual A4 copies from Excel cells.

Does not merge onto the full-page scanned Invoice_Template.pdf (that caused
overlap with printed address/borders). Data comes from receipt_model.
"""
from __future__ import annotations

import logging
from pathlib import Path

from reportlab.lib.colors import Color, black, white
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as pdf_canvas

from config import BASE_DIR, OUTPUT_DIR
from receipt_model import ReceiptCopy, ReceiptDocument, from_excel, from_invoice_data

logger = logging.getLogger(__name__)

PAGE_W, PAGE_H = A4
HEADER_PATH = BASE_DIR / "assets" / "receipt_header.png"

# Layout calibrated to Perfect V2 dual-copy A4
MARGIN_X = 18.0
HEADER_MAX_H = 72.0
COPY_GAP = 8.0  # space around cut line
CUT_Y = PAGE_H / 2.0

FONT_NAME = "GevinReceiptCJK"
INK = Color(0.12, 0.10, 0.08)
RULE = Color(0.55, 0.45, 0.30)
LIGHT_RULE = Color(0.75, 0.70, 0.60)


def _register_font() -> str:
    if FONT_NAME in pdfmetrics.getRegisteredFontNames():
        return FONT_NAME
    # TrueType CJK first. Noto CJK TTC on Debian is often CFF/PostScript
    # outlines, which ReportLab cannot embed.
    candidates = [
        Path("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"),
        Path("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"),
        Path("/usr/share/fonts/truetype/arphic/uming.ttc"),
        Path("/usr/share/fonts/truetype/arphic/ukai.ttc"),
        Path(r"C:\Windows\Fonts\kaiu.ttf"),
        Path(r"C:\Windows\Fonts\msyh.ttc"),
        Path(r"C:\Windows\Fonts\msjh.ttc"),
        Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"),
        Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
    ]
    for path in candidates:
        if not path.exists():
            continue
        try:
            if path.suffix.lower() == ".ttc":
                pdfmetrics.registerFont(TTFont(FONT_NAME, str(path), subfontIndex=0))
            else:
                pdfmetrics.registerFont(TTFont(FONT_NAME, str(path)))
            return FONT_NAME
        except Exception:
            continue
    raise RuntimeError("No Chinese-capable font for receipt PDF")


def _fit_string(c, text: str, font: str, size: float, max_w: float, min_size: float = 6.0) -> float:
    s = size
    while s > min_size and c.stringWidth(text, font, s) > max_w:
        s -= 0.5
    return s


def _draw_text(
    c,
    text: str,
    x: float,
    y: float,
    *,
    font: str,
    size: float = 9.0,
    align: str = "left",
    max_w: float | None = None,
    color=INK,
):
    if not text:
        return
    c.setFillColor(color)
    sz = size
    if max_w is not None:
        sz = _fit_string(c, text, font, size, max_w)
    c.setFont(font, sz)
    if align == "right":
        c.drawRightString(x, y, text)
    elif align == "center":
        c.drawCentredString(x, y, text)
    else:
        c.drawString(x, y, text)


def _draw_header(c, y_top: float) -> float:
    """Draw header image; return y (from bottom) just below header + gap."""
    content_w = PAGE_W - 2 * MARGIN_X
    if not HEADER_PATH.exists():
        raise FileNotFoundError(f"Receipt header missing: {HEADER_PATH}")
    # Preserve aspect; cap height so data never sits in logo/address band
    aspect = 1492 / 228
    h = min(HEADER_MAX_H, content_w / aspect)
    w = h * aspect
    if w > content_w:
        w = content_w
        h = w / aspect
    x = MARGIN_X + (content_w - w) / 2.0
    y_img = y_top - h
    c.drawImage(
        str(HEADER_PATH),
        x,
        y_img,
        width=w,
        height=h,
        preserveAspectRatio=True,
        mask="auto",
    )
    # Clear gap under header before any data labels
    return y_img - 14.0


def _draw_hline(c, y: float, x0: float = None, x1: float = None, color=LIGHT_RULE, stroke=0.6):
    c.setStrokeColor(color)
    c.setLineWidth(stroke)
    c.line(x0 if x0 is not None else MARGIN_X, y, x1 if x1 is not None else PAGE_W - MARGIN_X, y)


def _draw_copy(c, font: str, copy: ReceiptCopy, y_top: float, y_floor: float) -> None:
    """
    Draw one receipt copy in the vertical band [y_floor, y_top] (PDF coords, origin bottom-left).
    """
    y = _draw_header(c, y_top)
    left = MARGIN_X
    right = PAGE_W - MARGIN_X
    mid = PAGE_W / 2.0
    usable_bottom = y_floor + 10.0

    # Invoice type + number (right) + Date
    if getattr(copy, "invoice_label", ""):
        _draw_text(c, copy.invoice_label, right - 155, y, font=font, size=8.0, align="right", max_w=155)
    _draw_text(c, copy.invoice_no, right, y, font=font, size=10.0, align="right", max_w=160)
    y -= 16
    _draw_text(c, "日期 Date:", right - 110, y, font=font, size=8.0, align="left")
    _draw_text(c, copy.date, right, y, font=font, size=9.0, align="right", max_w=100)
    y -= 18

    # Customer / Tel — both copies
    _draw_text(c, "客戶 Customer:", left, y, font=font, size=8.0)
    _draw_text(c, copy.customer, left + 78, y, font=font, size=10.0, max_w=200)
    if copy.phone:
        _draw_text(c, "電話 Tel:", mid - 20, y, font=font, size=8.0)
        _draw_text(c, copy.phone, mid + 40, y, font=font, size=10.0, max_w=120)
    y -= 6
    _draw_hline(c, y - 2, color=RULE, stroke=0.9)
    y -= 14

    # Column headers (match old Excel bilingual titles)
    col_item = left
    col_qual = left + 155
    col_wt = left + 210
    col_stock = left + 300
    col_price = right - 150
    col_amt = right
    _draw_text(c, "貨品 Item", col_item, y, font=font, size=7.0, color=RULE)
    _draw_text(c, "成色 Quality", col_qual, y, font=font, size=7.0, color=RULE)
    _draw_text(c, "重量 Weight", col_wt, y, font=font, size=7.0, color=RULE)
    if copy.kind == "company":
        _draw_text(c, "庫存 Stock", col_stock, y, font=font, size=7.0, color=RULE)
    _draw_text(c, "單價 Unit ($)", col_price, y, font=font, size=7.0, align="right", color=RULE)
    _draw_text(c, "現金倉 Cash Warehouse", col_amt, y, font=font, size=7.0, align="right", color=RULE)
    y -= 4
    _draw_hline(c, y - 2)
    y -= 14

    for line in copy.lines:
        if y < usable_bottom + 80:
            break
        row_top = y
        _draw_text(c, line.item_type, col_item, y, font=font, size=9.0, max_w=148)
        _draw_text(c, line.quality, col_qual, y, font=font, size=9.0, max_w=50)
        if line.unit_price != "":
            _draw_text(
                c, line.unit_price,
                col_price, y, font=font, size=9.0, align="right", max_w=50,
            )
        if line.amount_display:
            _draw_text(
                c, line.amount_display, col_amt, y, font=font, size=9.0, align="right", max_w=120,
            )
        elif line.amount is not None:
            _draw_text(
                c, f"{line.currency} {float(line.amount):.2f}".strip(),
                col_amt, y, font=font, size=9.0, align="right", max_w=120,
            )

        # Weights stacked
        wy = y
        for wl in line.weight_lines:
            _draw_text(c, wl, col_wt, wy, font=font, size=8.5, max_w=130)
            wy -= 12
        # Stock (company) under weights / to the right of weights band
        stock_y = min(wy, y) if line.weight_lines else y - 12
        if copy.kind == "company":
            sy = y - 12
            if line.stock_action:
                _draw_text(c, line.stock_action, col_wt + 95, sy, font=font, size=7.5, max_w=140)
                sy -= 11
            if line.stock_location:
                _draw_text(c, line.stock_location, col_wt + 95, sy, font=font, size=7.5, max_w=140)
                stock_y = min(stock_y, sy - 2)

        y = min(wy, stock_y if copy.kind == "company" else wy) - 6
        _draw_hline(c, y + 2, color=LIGHT_RULE, stroke=0.4)
        y -= 10
        _ = row_top  # keep structure clear

    # Notes
    if copy.notes and y > usable_bottom + 50:
        _draw_text(c, f"備註 Notes: {copy.notes}", left, y, font=font, size=8.0, max_w=right - left)
        y -= 14

    # Total
    y = max(y - 4, usable_bottom + 46)
    _draw_hline(c, y + 10, color=RULE, stroke=1.0)
    _draw_text(c, "合計 Total:", right - 130, y, font=font, size=9.0)
    _draw_text(
        c, copy.total_display or "HKD 0.00",
        right, y, font=font, size=11.0, align="right", max_w=120,
    )
    y -= 18

    # Payment methods
    _draw_text(c, "付款方式 Payment Method :", left, y, font=font, size=8.0)
    y -= 12
    for pl in copy.payment_lines or []:
        _draw_text(c, pl, left + 8, y, font=font, size=8.5, max_w=right - left - 8)
        y -= 11
    y -= 4

    # Handler
    _draw_text(c, "經手人 Handled by:", left, y, font=font, size=8.0)
    _draw_text(c, copy.handler, left + 95, y, font=font, size=10.0, max_w=120)
    y -= 16

    # Footer copy marker
    _draw_text(
        c, copy.footer_label, PAGE_W / 2.0, max(y, usable_bottom + 4),
        font=font, size=8.0, align="center", color=RULE,
    )


def _draw_cut_line(c, font: str):
    y = CUT_Y
    c.setStrokeColor(RULE)
    c.setDash(3, 3)
    c.setLineWidth(0.8)
    c.line(MARGIN_X, y, PAGE_W - MARGIN_X, y)
    c.setDash()
    _draw_text(
        c, "✂ 裁切線 Cut here", PAGE_W / 2.0, y + 3,
        font=font, size=7.0, align="center", color=RULE,
    )


def render_receipt_pdf(doc: ReceiptDocument, out_path: Path) -> str:
    font = _register_font()
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    c = pdf_canvas.Canvas(str(out_path), pagesize=A4)
    c.setFillColor(white)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)

    # Top = customer, bottom = company
    customer_top = PAGE_H - 10.0
    customer_floor = CUT_Y + COPY_GAP
    company_top = CUT_Y - COPY_GAP
    company_floor = 10.0

    _draw_copy(c, font, doc.customer, customer_top, customer_floor)
    _draw_cut_line(c, font)
    _draw_copy(c, font, doc.company, company_top, company_floor)

    c.save()
    logger.info("Perfect V2 receipt PDF written: %s (%s bytes)", out_path, out_path.stat().st_size)
    return str(out_path.resolve())


def build_invoice_pdf_from_data(invoice_data, main_items, exchange_items=None, out_path=None) -> str:
    """Perfect V2 PDF from the live invoice payload (not overlay Excel cells)."""
    if not HEADER_PATH.exists():
        raise FileNotFoundError(f"Receipt header missing: {HEADER_PATH}")
    doc = from_invoice_data(invoice_data, main_items, exchange_items)
    logger.info(
        "Building Perfect V2 PDF invoice_no=%s customer=%s lines=%s",
        doc.invoice_no,
        doc.customer.customer,
        len(doc.customer.lines),
    )
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if out_path is None:
        stem = invoice_data.get("invoice_no") or "invoice"
        out_path = OUTPUT_DIR / f"{stem}.pdf"
    return render_receipt_pdf(doc, Path(out_path))


def build_invoice_pdf_from_excel(excel_path) -> str:
    """Legacy: read old template-filled workbooks. New overlays use from_invoice_data."""
    excel_path = Path(excel_path)
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel not found: {excel_path}")
    if not HEADER_PATH.exists():
        raise FileNotFoundError(f"Receipt header missing: {HEADER_PATH}")

    doc = from_excel(excel_path)
    logger.info(
        "Building Perfect V2 PDF invoice_no=%s customer=%s lines=%s",
        doc.invoice_no,
        doc.customer.customer,
        len(doc.customer.lines),
    )
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUTPUT_DIR / f"{excel_path.stem}.pdf"
    return render_receipt_pdf(doc, out)


def build_invoice_pdf(excel_path, *args, **kwargs) -> str:
    """Compatibility entry: Excel path → Perfect V2 PDF."""
    return build_invoice_pdf_from_excel(excel_path)
