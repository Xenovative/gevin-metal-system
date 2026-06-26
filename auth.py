"""使用者登入、權限與稽核日誌。"""

import hashlib
import json
import secrets
from datetime import datetime

from config import ROLE_ADMIN, ROLE_STAFF, ROLE_PERMISSIONS

_PBKDF2_ITERATIONS = 120_000


def format_admin_timestamp():
    """Admin 備註時間戳：YYYY/MM/DD 時間：Hour : Minutes AM/PM"""
    now = datetime.now()
    return (
        f"Admin於系統時間：{now.strftime('%Y/%m/%d')} "
        f"時間：{now.strftime('%I').lstrip('0') or '12'} : {now.strftime('%M %p')}"
    )


def append_admin_note(existing_notes, action_text):
    return f"{existing_notes or ''}\n{format_admin_timestamp()} {action_text}".strip()


def hash_password(password: str) -> str:
    salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt.encode("utf-8"), _PBKDF2_ITERATIONS,
    )
    return f"{salt}${digest.hex()}"


def verify_password(password: str, password_hash: str) -> bool:
    if not password_hash or "$" not in password_hash:
        return False
    salt, expected = password_hash.split("$", 1)
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt.encode("utf-8"), _PBKDF2_ITERATIONS,
    )
    return secrets.compare_digest(digest.hex(), expected)


def user_to_dict(user):
    return {
        "id": user.id,
        "username": user.username,
        "display_name": user.display_name,
        "role": user.role,
    }


def authenticate(session, username, password):
    from database import User

    username = (username or "").strip()
    if not username or not password:
        return None
    user = (
        session.query(User)
        .filter(User.username == username, User.is_active.is_(True))
        .first()
    )
    if user and verify_password(password, user.password_hash):
        return user_to_dict(user)
    return None


def is_admin(user):
    return user and user.get("role") == ROLE_ADMIN


def has_permission(user, permission):
    if not is_logged_in(user):
        return False
    role_perms = ROLE_PERMISSIONS.get(user.get("role"), {})
    return bool(role_perms.get(permission))


def can_view_inventory(user):
    return has_permission(user, "view_inventory")


def can_download_reports(user):
    return has_permission(user, "download_reports")


def require_permission(user, permission):
    err = require_login(user)
    if err:
        return err
    if not has_permission(user, permission):
        return "❌ 您沒有權限執行此操作"
    return None


def require_view_inventory(user):
    return require_permission(user, "view_inventory")


def require_download_reports(user):
    return require_permission(user, "download_reports")


def is_logged_in(user):
    return bool(user and user.get("id"))


def require_login(user):
    if not is_logged_in(user):
        return "❌ 請先登入"
    return None


def require_admin(user):
    err = require_login(user)
    if err:
        return err
    if not is_admin(user):
        return "❌ 僅 Admin 可執行此操作"
    return None


def log_audit(session, user, action, target_type=None, target_id=None, details=None):
    from database import AuditLog

    session.add(
        AuditLog(
            user_id=user.get("id") if user else None,
            username=user.get("username", "") if user else "",
            display_name=user.get("display_name", "") if user else "",
            action=action,
            target_type=target_type or "",
            target_id=target_id or "",
            details=json.dumps(details, ensure_ascii=False) if details else "",
        )
    )


def ensure_default_admin(session):
    from database import User

    if session.query(User).count() > 0:
        return
    session.add(
        User(
            username="admin",
            display_name="Admin",
            password_hash=hash_password("admin123"),
            role=ROLE_ADMIN,
            is_active=True,
        )
    )
    session.commit()


def list_users(session):
    from database import User

    users = session.query(User).order_by(User.role, User.username).all()
    return [
        {
            "帳號": u.username,
            "姓名": u.display_name,
            "權限": "Admin" if u.role == ROLE_ADMIN else "員工",
            "狀態": "啟用" if u.is_active else "停用",
        }
        for u in users
    ]


def create_staff_user(session, admin_user, username, display_name, password):
    from database import User

    err = require_admin(admin_user)
    if err:
        return err
    username = (username or "").strip()
    display_name = (display_name or "").strip()
    if not username or not display_name or not password:
        return "❌ 請填寫帳號、姓名及密碼"
    password = password.strip()
    if len(password) < 4:
        return "❌ 密碼至少需要 4 個字元"
    if session.query(User).filter(User.username == username).first():
        return f"❌ 帳號「{username}」已存在"
    try:
        session.add(
            User(
                username=username,
                display_name=display_name,
                password_hash=hash_password(password),
                role=ROLE_STAFF,
                is_active=True,
            )
        )
        log_audit(session, admin_user, "create_user", "user", username, {
            "display_name": display_name, "role": ROLE_STAFF,
        })
        session.commit()
    except Exception as exc:
        session.rollback()
        return f"❌ 建立失敗：{exc}"
    return f"✅ 已建立員工帳號：{username}（姓名：{display_name}）"


def update_user_profile(session, admin_user, username, new_display_name, new_password):
    from database import User

    err = require_admin(admin_user)
    if err:
        return err
    user = session.query(User).filter(User.username == username).first()
    if not user:
        return f"❌ 找不到帳號「{username}」"
    if user.role == ROLE_ADMIN and user.username != admin_user.get("username"):
        return "❌ 不可修改其他 Admin 帳號"
    changes = {}
    if new_display_name and new_display_name.strip():
        changes["display_name"] = {"old": user.display_name, "new": new_display_name.strip()}
        user.display_name = new_display_name.strip()
    if new_password and new_password.strip():
        user.password_hash = hash_password(new_password.strip())
        changes["password"] = "reset"
    if not changes:
        return "❌ 請填寫新姓名或新密碼"
    log_audit(session, admin_user, "update_user", "user", username, changes)
    session.commit()
    return f"✅ 已更新帳號「{username}」"


def load_audit_logs(session, limit=200):
    from database import AuditLog

    rows = (
        session.query(AuditLog)
        .order_by(AuditLog.created_at.desc(), AuditLog.id.desc())
        .limit(limit)
        .all()
    )
    return [
        {
            "時間": r.created_at.strftime("%Y-%m-%d %H:%M:%S"),
            "操作者": r.display_name or r.username,
            "動作": r.action,
            "對象": r.target_id or "",
            "詳情": r.details or "",
        }
        for r in rows
    ]
