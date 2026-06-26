"""發票作廢與倉存／現金倉反向沖銷。"""

from datetime import datetime

from auth import append_admin_note, log_audit, require_admin


def void_invoice(session, invoice_no, admin_user):
    from database import CashMovement, InventoryMovement, Invoice

    err = require_admin(admin_user)
    if err:
        return err

    invoice_no = (invoice_no or "").strip()
    if not invoice_no:
        return "❌ 請填寫單號"

    invoice = session.query(Invoice).filter(Invoice.invoice_no == invoice_no).first()
    if not invoice:
        return f"❌ 找不到單號「{invoice_no}」"
    if invoice.status == "voided":
        return f"❌ 單號「{invoice_no}」已作廢"

    invoice.status = "voided"
    invoice.voided_at = datetime.now()
    invoice.voided_by_user_id = admin_user.get("id")
    invoice.notes = append_admin_note(invoice.notes, "作廢本單")

    inventory_rows = (
        session.query(InventoryMovement)
        .filter(
            InventoryMovement.invoice_no == invoice_no,
            InventoryMovement.movement_kind == "normal",
        )
        .all()
    )
    for movement in inventory_rows:
        reverse_direction = "out" if movement.direction == "in" else "in"
        session.add(
            InventoryMovement(
                invoice_no=invoice_no,
                transaction_type=movement.transaction_type,
                direction=reverse_direction,
                item_type=movement.item_type,
                quality=movement.quality,
                weight_gram=movement.weight_gram,
                weight_tael=movement.weight_tael,
                movement_date=datetime.now().date(),
                customer_name=movement.customer_name,
                handler=admin_user.get("display_name", ""),
                notes="作廢沖銷",
                source_location=movement.source_location,
                destination_location=movement.destination_location,
                movement_kind="reversal",
            )
        )

    cash_rows = (
        session.query(CashMovement)
        .filter(
            CashMovement.invoice_no == invoice_no,
            CashMovement.movement_kind == "normal",
        )
        .all()
    )
    for movement in cash_rows:
        reverse_direction = "out" if movement.direction == "in" else "in"
        session.add(
            CashMovement(
                invoice_no=invoice_no,
                transaction_type=movement.transaction_type,
                direction=reverse_direction,
                amount=movement.amount,
                currency=movement.currency,
                warehouse=movement.warehouse,
                movement_date=datetime.now().date(),
                customer_name=movement.customer_name,
                handler=admin_user.get("display_name", ""),
                notes="作廢沖銷",
                movement_kind="reversal",
            )
        )

    log_audit(session, admin_user, "void_invoice", "invoice", invoice_no, {
        "inventory_reversals": len(inventory_rows),
        "cash_reversals": len(cash_rows),
    })
    session.commit()
    return (
        f"✅ 已作廢單號 {invoice_no}\n"
        f"已沖銷倉存記錄 {len(inventory_rows)} 筆、現金記錄 {len(cash_rows)} 筆"
    )
