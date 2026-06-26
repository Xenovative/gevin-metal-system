from datetime import date, datetime
from pathlib import Path

from sqlalchemy import (
    Boolean,
    Column,
    Date,
    DateTime,
    Float,
    Integer,
    String,
    Text,
    create_engine,
)
from sqlalchemy.orm import declarative_base, sessionmaker

from auth import ensure_default_admin
from config import DB_PATH, INVOICE_STATUS_ACTIVE

Base = declarative_base()


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True)
    username = Column(String(50), unique=True, nullable=False)
    display_name = Column(String(100), nullable=False)
    password_hash = Column(String(200), nullable=False)
    role = Column(String(20), nullable=False)  # admin / staff
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.now)


class AuditLog(Base):
    __tablename__ = "audit_logs"

    id = Column(Integer, primary_key=True)
    user_id = Column(Integer)
    username = Column(String(50))
    display_name = Column(String(100))
    action = Column(String(50), nullable=False)
    target_type = Column(String(50))
    target_id = Column(String(100))
    details = Column(Text)
    created_at = Column(DateTime, default=datetime.now)


class Invoice(Base):
    __tablename__ = "invoices"

    id = Column(Integer, primary_key=True)
    invoice_no = Column(String(50), unique=True, nullable=False)
    transaction_type = Column(String(50), nullable=False)
    customer_name = Column(String(100), nullable=False)
    transaction_date = Column(Date, nullable=False)
    handler = Column(String(100))
    payment_method = Column(Text)
    notes = Column(Text)
    note_amount = Column(Float, default=0)
    total_amount = Column(Float, default=0)
    excel_path = Column(String(500))
    source_location = Column(String(50))
    destination_location = Column(String(50))
    status = Column(String(20), default=INVOICE_STATUS_ACTIVE)
    created_by_user_id = Column(Integer)
    voided_at = Column(DateTime)
    voided_by_user_id = Column(Integer)
    created_at = Column(DateTime, default=datetime.now)


class InvoiceLineItem(Base):
    __tablename__ = "invoice_line_items"

    id = Column(Integer, primary_key=True)
    invoice_id = Column(Integer, nullable=False)
    section = Column(String(20), default="main")  # main / exchange
    item_type = Column(String(100))
    quality = Column(String(50))
    weight_gram = Column(Float)
    weight_tael = Column(Float)
    unit_price = Column(Float)
    amount = Column(Float)
    sort_order = Column(Integer, default=0)


class CashMovement(Base):
    __tablename__ = "cash_movements"

    id = Column(Integer, primary_key=True)
    invoice_no = Column(String(50))
    transaction_type = Column(String(50), nullable=False)
    direction = Column(String(10), nullable=False)  # in / out
    amount = Column(Float, default=0)
    currency = Column(String(20), default="HKD$")
    warehouse = Column(String(50), default="現金倉")
    movement_date = Column(Date, nullable=False)
    customer_name = Column(String(100))
    handler = Column(String(100))
    notes = Column(Text)
    movement_kind = Column(String(20), default="normal")  # normal / reversal
    created_at = Column(DateTime, default=datetime.now)


class InventoryMovement(Base):
    __tablename__ = "inventory_movements"

    id = Column(Integer, primary_key=True)
    invoice_no = Column(String(50))
    transaction_type = Column(String(50), nullable=False)
    direction = Column(String(10), nullable=False)  # in / out
    item_type = Column(String(100), nullable=False)
    quality = Column(String(50))
    weight_gram = Column(Float, default=0)
    weight_tael = Column(Float, default=0)
    movement_date = Column(Date, nullable=False)
    customer_name = Column(String(100))
    handler = Column(String(100))
    notes = Column(Text)
    source_location = Column(String(50))
    destination_location = Column(String(50))
    movement_kind = Column(String(20), default="normal")  # normal / reversal
    created_at = Column(DateTime, default=datetime.now)


def _migrate_db(engine):
    """為既有資料庫補上新欄位。"""
    from sqlalchemy import inspect, text

    migrations = {
        "invoices": [
            "source_location VARCHAR(50)",
            "destination_location VARCHAR(50)",
            "status VARCHAR(20)",
            "created_by_user_id INTEGER",
            "voided_at DATETIME",
            "voided_by_user_id INTEGER",
        ],
        "inventory_movements": [
            "source_location VARCHAR(50)",
            "destination_location VARCHAR(50)",
            "movement_kind VARCHAR(20)",
        ],
        "cash_movements": [
            "currency VARCHAR(20)",
            "movement_kind VARCHAR(20)",
        ],
    }
    inspector = inspect(engine)
    with engine.begin() as conn:
        for table, columns in migrations.items():
            if table not in inspector.get_table_names():
                continue
            existing = {col["name"] for col in inspector.get_columns(table)}
            for col_def in columns:
                col_name = col_def.split()[0]
                if col_name not in existing:
                    conn.execute(text(f"ALTER TABLE {table} ADD COLUMN {col_def}"))

        if "invoices" in inspector.get_table_names():
            conn.execute(text(
                "UPDATE invoices SET status = 'active' WHERE status IS NULL OR status = ''"
            ))
        if "inventory_movements" in inspector.get_table_names():
            conn.execute(text(
                "UPDATE inventory_movements SET movement_kind = 'normal' "
                "WHERE movement_kind IS NULL OR movement_kind = ''"
            ))
        if "cash_movements" in inspector.get_table_names():
            conn.execute(text(
                "UPDATE cash_movements SET movement_kind = 'normal' "
                "WHERE movement_kind IS NULL OR movement_kind = ''"
            ))


def init_db():
    Path(DB_PATH).parent.mkdir(parents=True, exist_ok=True)
    engine = create_engine(f"sqlite:///{DB_PATH}")
    Base.metadata.create_all(engine)
    _migrate_db(engine)
    db_session = sessionmaker(bind=engine)()
    ensure_default_admin(db_session)
    return db_session


def save_invoice(session, invoice_data, line_items, movements, cash_movement=None, created_by_user_id=None):
    invoice = Invoice(
        invoice_no=invoice_data["invoice_no"],
        transaction_type=invoice_data["transaction_type"],
        customer_name=invoice_data["customer_name"],
        transaction_date=invoice_data["transaction_date"],
        handler=invoice_data.get("handler", ""),
        payment_method=invoice_data.get("payment_method", ""),
        notes=invoice_data.get("notes", ""),
        note_amount=invoice_data.get("note_amount", 0),
        total_amount=invoice_data.get("total_amount") or 0,
        excel_path=invoice_data.get("excel_path", ""),
        source_location=invoice_data.get("source_location", ""),
        destination_location=invoice_data.get("destination_location", ""),
        status=INVOICE_STATUS_ACTIVE,
        created_by_user_id=created_by_user_id,
    )
    session.add(invoice)
    session.flush()

    for idx, item in enumerate(line_items):
        session.add(
            InvoiceLineItem(
                invoice_id=invoice.id,
                section=item.get("section", "main"),
                item_type=item.get("item_type", ""),
                quality=item.get("quality", ""),
                weight_gram=item.get("weight_gram"),
                weight_tael=item.get("weight_tael"),
                unit_price=item.get("unit_price"),
                amount=item.get("amount"),
                sort_order=idx,
            )
        )

    for movement in movements:
        session.add(
            InventoryMovement(
                invoice_no=invoice_data["invoice_no"],
                transaction_type=invoice_data["transaction_type"],
                direction=movement["direction"],
                item_type=movement["item_type"],
                quality=movement.get("quality", ""),
                weight_gram=movement.get("weight_gram", 0) or 0,
                weight_tael=movement.get("weight_tael", 0) or 0,
                movement_date=invoice_data["transaction_date"],
                customer_name=invoice_data["customer_name"],
                handler=invoice_data.get("handler", ""),
                notes=movement.get("notes", ""),
                source_location=invoice_data.get("source_location", ""),
                destination_location=invoice_data.get("destination_location", ""),
                movement_kind="normal",
            )
        )

    if cash_movement:
        session.add(
            CashMovement(
                invoice_no=cash_movement["invoice_no"],
                transaction_type=cash_movement["transaction_type"],
                direction=cash_movement["direction"],
                amount=cash_movement["amount"],
                currency=cash_movement.get("currency", "HKD$"),
                warehouse=cash_movement.get("warehouse", "現金倉"),
                movement_date=cash_movement["movement_date"],
                customer_name=cash_movement.get("customer_name", ""),
                handler=cash_movement.get("handler", ""),
                notes=cash_movement.get("notes", ""),
                movement_kind="normal",
            )
        )

    session.commit()
    return invoice
