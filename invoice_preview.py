"""
Print-aligned Invoice Review preview.

Primary: embed the generated PDF (from Excel). Fallback: HTML from Excel cells
over the branded A4 background.
"""

from __future__ import annotations

import base64
import html
from pathlib import Path

from config import BASE_DIR
from receipt_model import ReceiptCopy, ReceiptDocument, from_excel

RECEIPT_BG_PATH = BASE_DIR / "assets" / "receipt_a4_bg.png"


def _esc(value):
    if value is None:
        return ""
    return html.escape(str(value))


def build_pdf_preview_html(
    pdf_path,
    *,
    invoice_no: str = "",
    tx_type: str = "",
    status: str = "正常",
    excel_name: str = "",
) -> str:
    """
    Embed the real receipt PDF in Invoice Review step 4.
    Uses a base64 data-URI iframe so Gradio can show it without a public file URL.
    """
    path = Path(pdf_path)
    if not path.exists():
        return (
            "<div class='preview-wrap'>"
            f"<p>❌ 找不到 PDF：{_esc(path.name)}。請再按「預覽」重新生成。</p>"
            "</div>"
        )
    try:
        raw = path.read_bytes()
    except OSError as exc:
        return (
            "<div class='preview-wrap'>"
            f"<p>❌ 無法讀取 PDF：{_esc(exc)}</p>"
            "</div>"
        )

    b64 = base64.b64encode(raw).decode("ascii")
    data_uri = f"data:application/pdf;base64,{b64}"
    inv = invoice_no or path.stem
    excel_bit = f"｜Excel {_esc(excel_name)}" if excel_name else ""
    pdf_bit = f"｜PDF {_esc(path.name)}"

    return f"""
<div class="preview-wrap">
  <div class="preview-banner">
    <strong>列印預覽 Print Preview (PDF)</strong>
    — 由剛生成的 Excel 產生｜單號 {_esc(inv)}｜{_esc(tx_type) or "—"}｜狀態 {_esc(status)}{excel_bit}{pdf_bit}
  </div>
  <div class="pdf-frame">
    <iframe
      title="Receipt PDF {_esc(path.name)}"
      src="{data_uri}#toolbar=1&navpanes=0&view=FitH"
      type="application/pdf"
    ></iframe>
  </div>
  <p class="pdf-hint">若瀏覽器無法內嵌 PDF，請於步驟 5 下載 PDF 後開啟。</p>
</div>
<style>
.preview-wrap {{ font-family: "Microsoft JhengHei", "Segoe UI", sans-serif; }}
.preview-banner {{
  margin: 0 0 10px; padding: 8px 12px; background:#2f2a24; color:#f3e6c8;
  border-radius: 6px; font-size: 13px; line-height: 1.4;
}}
.pdf-frame {{
  width: min(100%, 820px);
  margin: 0 auto;
  border: 1px solid #c9b89a;
  box-shadow: 0 8px 28px rgba(60,40,10,.18);
  background: #525659;
  border-radius: 4px;
  overflow: hidden;
}}
.pdf-frame iframe {{
  display: block;
  width: 100%;
  height: 1100px;
  border: 0;
  background: #525659;
}}
.pdf-hint {{
  width: min(100%, 820px);
  margin: 8px auto 0;
  font-size: 12px;
  color: #6a5a48;
}}
@media (max-width: 700px) {{
  .pdf-frame iframe {{ height: 900px; }}
}}
</style>
"""


def _bg_data_uri():
    if not RECEIPT_BG_PATH.exists():
        return None
    raw = RECEIPT_BG_PATH.read_bytes()
    b64 = base64.b64encode(raw).decode("ascii")
    return f"data:image/png;base64,{b64}"


def _val(text, fallback="—"):
    """Visible data chip so values never wash out against the brand bg."""
    t = (text or "").strip()
    if not t:
        t = fallback
    return f'<span class="val">{_esc(t)}</span>'


def _items_rows_html(copy: ReceiptCopy) -> str:
    rows = []
    for ln in copy.lines:
        weights = "<br>".join(_esc(w) for w in ln.weight_lines) if ln.weight_lines else "—"
        stock_bits = []
        if copy.kind == "company":
            if ln.stock_action:
                stock_bits.append(_esc(ln.stock_action))
            if ln.stock_location:
                stock_bits.append(_esc(ln.stock_location))
        stock = "<br>".join(stock_bits) if stock_bits else ""
        price = ln.unit_price if ln.unit_price != "" else "—"
        amt = ln.amount_display or "—"
        rows.append(
            "<tr>"
            f"<td><span class='val'>{_esc(ln.item_type) or '—'}</span></td>"
            f"<td><span class='val'>{_esc(ln.quality) or '—'}</span></td>"
            f"<td><span class='val'>{weights}</span></td>"
            f"<td class='stock'><span class='val'>{stock or ('—' if copy.kind == 'company' else '')}</span></td>"
            f"<td class='num'><span class='val'>{_esc(price)}</span></td>"
            f"<td class='num'><span class='val'>{_esc(amt)}</span></td>"
            "</tr>"
        )
    if not rows:
        rows.append('<tr><td colspan="6" class="muted">（Excel 無貨品列）</td></tr>')
    return "\n".join(rows)


def _copy_panel_html(copy: ReceiptCopy, number_label: str = "Invoice No.") -> str:
    phone_html = ""
    if copy.kind == "company" and copy.phone:
        phone_html = f' <span class="phone">電話 {_val(copy.phone, "")}</span>'
    pay_html = (
        "<br>".join(_esc(p) for p in copy.payment_lines) if copy.payment_lines else "—"
    )
    stock_th = "庫存 Stock" if copy.kind == "company" else ""
    title = (
        "公司單 Company Copy" if copy.kind == "company" else "客戶單 Customer Copy"
    )

    return f"""
    <section class="copy-panel" data-copy="{_esc(copy.kind)}">
      <div class="brand-spacer" aria-hidden="true"></div>
      <div class="form-frame">
        <div class="meta-row">
          <div class="meta-left">
            <div><span class="lbl">客戶 Customer</span> {_val(copy.customer)}{phone_html}</div>
          </div>
          <div class="meta-right">
            <div><span class="lbl">{_esc(number_label)}</span> {_val(copy.invoice_no)}</div>
            <div><span class="lbl">Date</span> {_val(copy.date)}</div>
          </div>
        </div>

        <table class="items">
          <thead>
            <tr>
              <th>貨品 Item</th>
              <th>成色 Quality</th>
              <th>重量 Gross</th>
              <th>{stock_th}</th>
              <th>單價 Unit</th>
              <th>現金倉 Amount</th>
            </tr>
          </thead>
          <tbody>
            {_items_rows_html(copy)}
          </tbody>
        </table>

        <div class="footer-grid">
          <div class="pay-box">
            <div class="lbl">付款方式 Payment</div>
            <div class="val">{pay_html}</div>
            <div class="handler"><span class="lbl">經手人</span> {_val(copy.handler)}</div>
          </div>
          <div class="notes-box">
            <div class="lbl">備註 Notes</div>
            <div class="val">{_esc(copy.notes) or "—"}</div>
            {f"<div class='num val'>{_esc(copy.note_amount_display)}</div>" if copy.note_amount_display else ""}
          </div>
          <div class="total-box">
            <div class="lbl">合計 Total / Foreign Currency</div>
            <div class="total-amt">{_val(copy.total_display, "HKD$ 0.00")}</div>
            <div class="copy-tag">{_esc(title)}</div>
          </div>
        </div>
      </div>
    </section>
    """


def build_print_preview_html_from_document(
    doc: ReceiptDocument,
    *,
    tx_type: str = "",
    status: str = "正常",
    number_label: str = "Invoice No.",
    excel_name: str = "",
) -> str:
    """Render dual-copy A4 preview from canonical receipt document (Excel cells)."""
    bg = _bg_data_uri()
    bg_css = (
        f"background-image:url('{bg}'); background-size:100% 100%; background-repeat:no-repeat;"
        if bg
        else "background:#f7f3ea;"
    )
    src_note = f"｜Excel {_esc(excel_name)}" if excel_name else "｜來源 Excel 列印檔"
    customer_panel = _copy_panel_html(doc.customer, number_label=number_label)
    company_panel = _copy_panel_html(doc.company, number_label=number_label)

    return f"""
<div class="preview-wrap">
  <div class="preview-banner">
    <strong>列印預覽 Print Preview</strong>
    — 與即將列印的 Excel 相同（客戶單 + 公司單）｜單號 {_esc(doc.invoice_no)}｜{_esc(tx_type) or "—"}｜狀態 {_esc(status)}{src_note}
  </div>
  <div class="a4-sheet" style="{bg_css}">
    {customer_panel}
    {company_panel}
  </div>
</div>
<style>
.preview-wrap {{ font-family: "Microsoft JhengHei", "Segoe UI", "Noto Sans TC", sans-serif; color:#1a1510; }}
.preview-banner {{
  margin: 0 0 10px; padding: 8px 12px; background:#2f2a24; color:#f3e6c8;
  border-radius: 6px; font-size: 13px; line-height: 1.4;
}}
.a4-sheet {{
  width: min(100%, 820px);
  aspect-ratio: 210 / 297;
  margin: 0 auto;
  border: 1px solid #c9b89a;
  box-shadow: 0 8px 28px rgba(60,40,10,.18);
  position: relative;
  overflow: hidden;
  background-color: #fffdf8;
}}
.copy-panel {{
  height: 50%;
  position: relative;
  box-sizing: border-box;
  padding: 0 3.2% 1.2% 3.2%;
  display: flex;
  flex-direction: column;
}}
.brand-spacer {{
  height: 17.5%;
  flex: 0 0 auto;
}}
.form-frame {{
  flex: 1 1 auto;
  min-height: 0;
  border: 1.5px solid rgba(90,70,40,.45);
  background: rgba(255,253,248,.92);
  padding: 1.6% 1.8% 1.2%;
  display: flex;
  flex-direction: column;
  gap: 4px;
}}
.meta-row {{ display:flex; justify-content:space-between; gap:8px; font-size:12px; }}
.meta-right {{ text-align:right; }}
.lbl {{ color:#6a5538; font-size:10px; margin-right:4px; font-weight:600; }}
.val {{
  color:#111; font-weight:700; font-size:12px;
  background: rgba(255,255,255,.92);
  padding: 0 3px; border-radius: 2px;
  display: inline-block; line-height: 1.35;
}}
.phone {{ margin-left:6px; }}
table.items {{
  width:100%; border-collapse:collapse; font-size:11px; margin-top:2px;
  flex: 1 1 auto;
}}
table.items th, table.items td {{
  border-bottom: 1px solid rgba(120,100,70,.3);
  padding: 3px 3px; vertical-align: top;
}}
table.items th {{
  color:#5a4530; font-weight:700; text-align:left; font-size:10px;
  border-bottom: 1.5px solid rgba(120,100,70,.5);
}}
table.items td.num, table.items th:nth-child(5), table.items th:nth-child(6) {{ text-align:right; }}
table.items td.stock {{ font-size:10px; max-width:110px; }}
.footer-grid {{
  display:grid; grid-template-columns: 1.2fr 1fr 0.9fr; gap:6px; margin-top:4px; font-size:11px;
}}
.pay-box, .notes-box, .total-box {{
  border: 1px solid rgba(120,100,70,.4); padding: 5px 7px; min-height: 54px;
  background:rgba(255,255,255,.85);
}}
.total-box {{ text-align:right; }}
.total-amt {{ font-size:15px; margin-top:2px; }}
.total-amt .val {{ font-size:15px; color:#4a2e0a; }}
.copy-tag {{ margin-top:4px; font-size:9px; color:#7a6548; font-weight:600; }}
.handler {{ margin-top:6px; }}
.muted {{ color:#8a7a66; }}
@media (max-width: 700px) {{
  .a4-sheet {{ width:100%; aspect-ratio: auto; min-height: 1100px; }}
  .footer-grid {{ grid-template-columns: 1fr; }}
}}
</style>
"""


def build_print_preview_html_from_excel(
    excel_path,
    *,
    tx_type: str = "",
    status: str = "正常",
    number_label: str = "Invoice No.",
) -> str:
    """Load filled workbook and build preview from its cells (print source of truth)."""
    path = Path(excel_path)
    doc = from_excel(path)
    return build_print_preview_html_from_document(
        doc,
        tx_type=tx_type,
        status=status,
        number_label=number_label,
        excel_name=path.name,
    )


def build_print_preview_html(invoice, line_dicts):
    """
    Legacy entry: prefer Excel path via build_print_preview_html_from_excel.
    Falls back to reconstructing a document from ORM-ish fields (less accurate).
    """
    from receipt_model import ReceiptLine, ReceiptCopy, ReceiptDocument, CUSTOMER_HANDLER
    from invoice_generator import format_payment_excel_lines
    from config import TRANSACTION_TYPES, DEFAULT_CASH_CURRENCY
    from receipt_model import _fmt_money, _fmt_weight_line, _fmt_price, _fmt_date

    tx_type = getattr(invoice, "transaction_type", "") or ""
    tx_config = TRANSACTION_TYPES.get(tx_type, {})
    number_label = tx_config.get("number_label", "Invoice No.")
    currency = getattr(invoice, "invoice_currency", None) or DEFAULT_CASH_CURRENCY
    tx_date = _fmt_date(getattr(invoice, "transaction_date", None))
    inv = getattr(invoice, "invoice_no", "") or ""
    customer = getattr(invoice, "customer_name", "") or ""
    phone = getattr(invoice, "customer_phone", "") or ""
    handler = getattr(invoice, "handler", "") or "Admin"
    payments = format_payment_excel_lines(getattr(invoice, "payment_method", "") or "")
    total = getattr(invoice, "total_amount", 0)
    total_disp = _fmt_money(currency, total if total is not None else 0)
    notes = getattr(invoice, "notes", "") or ""
    source = getattr(invoice, "source_location", "") or ""
    destination = getattr(invoice, "destination_location", "") or ""

    def lines(with_stock: bool):
        out = []
        for d in line_dicts or []:
            w = []
            if d.get("weight_gram") is not None:
                w.append(_fmt_weight_line(d.get("weight_gram"), "克 Gram"))
            if d.get("weight_tael") is not None:
                w.append(_fmt_weight_line(d.get("weight_tael"), "両 Tael"))
            if d.get("weight_oz") is not None:
                w.append(_fmt_weight_line(d.get("weight_oz"), "安士 oz"))
            amt = d.get("amount")
            ln = ReceiptLine(
                item_type=d.get("item_type") or "",
                quality=d.get("quality") or "",
                weight_lines=w,
                unit_price=_fmt_price(d.get("unit_price")),
                currency=currency,
                amount=amt,
                amount_display=_fmt_money(currency, amt if amt is not None else 0),
            )
            if with_stock:
                if source:
                    ln.stock_action = str(source).strip()
                if destination:
                    ln.stock_location = str(destination).strip()
            out.append(ln)
        return out

    doc = ReceiptDocument(
        invoice_no=inv,
        customer=ReceiptCopy(
            kind="customer",
            invoice_no=inv,
            date=tx_date,
            customer=customer,
            lines=lines(False),
            notes=notes,
            total_display=total_disp,
            handler=CUSTOMER_HANDLER,
            payment_lines=payments,
        ),
        company=ReceiptCopy(
            kind="company",
            invoice_no=inv,
            date=tx_date,
            customer=customer,
            phone=phone,
            lines=lines(True),
            notes=notes,
            total_display=total_disp,
            handler=handler or "Admin",
            payment_lines=payments,
        ),
    )
    status = "已作廢" if (getattr(invoice, "status", "") or "") == "voided" else "正常"
    return build_print_preview_html_from_document(
        doc, tx_type=tx_type, status=status, number_label=number_label
    )


def lines_from_orm(line_items):
    """Convert InvoiceLineItem ORM rows to preview dicts."""
    out = []
    for line in line_items:
        out.append({
            "section": line.section or "main",
            "item_type": line.item_type or "",
            "quality": line.quality or "",
            "weight_gram": line.weight_gram,
            "weight_tael": line.weight_tael,
            "weight_oz": getattr(line, "weight_oz", None),
            "unit_price": line.unit_price,
            "amount": line.amount,
        })
    return out
