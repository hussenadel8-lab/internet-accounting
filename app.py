"""
برنامج محاسبة الإنترنت
-----------------------
موقع ويب بسيط لإدارة مشتركي الإنترنت والمحفظة والتفعيل/الترحيل والتقارير.
البيانات تُخزَّن في قاعدة بيانات SQLite داخل ملف واحد.

التشغيل:
    pip install -r requirements.txt
    python app.py
ثم افتح المتصفح على: http://localhost:5000
"""

import os
import re
import secrets
import sqlite3
import tempfile
from datetime import datetime, date, timedelta
from functools import wraps
from io import BytesIO

from flask import (
    Flask, request, redirect, url_for, session,
    render_template_string, send_file, flash, abort
)
from markupsafe import escape
from openpyxl import Workbook
from openpyxl.styles import Font
from werkzeug.security import generate_password_hash, check_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)
# مفتاح الجلسات: يُفضَّل ضبطه كمتغير بيئة على السيرفر.
app.secret_key = os.environ.get("SECRET_KEY") or "change-this-secret-key"
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=12)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # أقصى حجم للملف المرفوع

DB_NAME = os.environ.get("DB_PATH", os.path.join(BASE_DIR, "internet_accounting.db"))
PROFIT = 5000              # الربح الافتراضي للباقات الجديدة
SUBSCRIPTION_DAYS = 30     # مدة الاشتراك بالأيام
EXPIRY_WARNING_DAYS = 3    # التنبيه قبل انتهاء الاشتراك بهذه الأيام

PLANS = [
    ("فايبر", "فايبر 35", 36000),
    ("فايبر", "فايبر 50", 45000),
    ("فايبر", "فايبر 75", 65000),
    ("فايبر", "فايبر 150", 100000),
    ("وايرليس", "إيكو", 36000),
    ("وايرليس", "بلص", 38000),
    ("وايرليس", "استاندر", 46000),
    ("وايرليس", "اكتف", 66000),
]


# ---------------------------------------------------------------------------
# أدوات مساعدة
# ---------------------------------------------------------------------------

def esc(value):
    """تنظيف النصوص المدخلة من المستخدم قبل عرضها في HTML."""
    return escape(value if value is not None else "")


def money(value):
    return f"{int(value or 0):,}"


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_str():
    return date.today().strftime("%Y-%m-%d")


def parse_date(value):
    """يرجع التاريخ بصيغة YYYY-MM-DD أو None إذا كان غير صالح."""
    try:
        return datetime.strptime((value or "").strip(), "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        return None


def parse_amount(value):
    """يحوّل المبلغ لرقم صحيح (يقبل الفواصل) أو None."""
    try:
        return int(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def form_text(name):
    return (request.form.get(name) or "").strip()


def db():
    connection = sqlite3.connect(DB_NAME, timeout=10)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def column_exists(cur, table, column):
    return any(r["name"] == column for r in cur.execute(f"PRAGMA table_info({table})"))


def init_db():
    con = db()
    cur = con.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT NOT NULL,
            name TEXT NOT NULL,
            retail_price INTEGER NOT NULL,
            profit INTEGER NOT NULL,
            wholesale_price INTEGER NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS subscribers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            username TEXT NOT NULL,
            phone TEXT,
            plan_id INTEGER NOT NULL,
            activation_date TEXT NOT NULL,
            notes TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY(plan_id) REFERENCES plans(id)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS wallet_transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            transaction_type TEXT NOT NULL,
            amount INTEGER NOT NULL,
            notes TEXT,
            created_at TEXT NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS activations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            subscriber_id INTEGER NOT NULL,
            plan_id INTEGER NOT NULL,
            operation_type TEXT NOT NULL,
            amount INTEGER NOT NULL,
            notes TEXT,
            operation_date TEXT NOT NULL,
            FOREIGN KEY(subscriber_id) REFERENCES subscribers(id),
            FOREIGN KEY(plan_id) REFERENCES plans(id)
        )
    """)

    # --- ترقية قواعد البيانات القديمة ---
    # حفظ سعر البيع والربح وقت العملية حتى لا تتغير التقارير القديمة
    # عند تعديل أسعار الباقات لاحقاً.
    if not column_exists(cur, "activations", "retail_price"):
        cur.execute("ALTER TABLE activations ADD COLUMN retail_price INTEGER")
        cur.execute("ALTER TABLE activations ADD COLUMN profit INTEGER")
        cur.execute("""
            UPDATE activations SET
                retail_price = (SELECT retail_price FROM plans WHERE plans.id = activations.plan_id),
                profit = (SELECT profit FROM plans WHERE plans.id = activations.plan_id)
        """)

    if not column_exists(cur, "subscribers", "expiry_date"):
        cur.execute("ALTER TABLE subscribers ADD COLUMN expiry_date TEXT")
        for row in cur.execute("SELECT id, activation_date FROM subscribers").fetchall():
            start = parse_date(row["activation_date"]) or today_str()
            expiry = (datetime.strptime(start, "%Y-%m-%d")
                      + timedelta(days=SUBSCRIPTION_DAYS)).strftime("%Y-%m-%d")
            cur.execute("UPDATE subscribers SET expiry_date = ? WHERE id = ?",
                        (expiry, row["id"]))

    # الديون: كل دين أو تسديد سطر مستقل، والرصيد = مجموع الديون - مجموع التسديدات
    cur.execute("""
        CREATE TABLE IF NOT EXISTS debts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            subscriber_id INTEGER NOT NULL,
            kind TEXT NOT NULL,
            amount INTEGER NOT NULL,
            notes TEXT,
            activation_id INTEGER,
            created_at TEXT NOT NULL
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_debts_sub ON debts(subscriber_id)")

    if not column_exists(cur, "activations", "payment_type"):
        cur.execute("ALTER TABLE activations ADD COLUMN payment_type TEXT NOT NULL DEFAULT 'نقد'")
        cur.execute("ALTER TABLE activations ADD COLUMN paid_amount INTEGER")
        cur.execute("UPDATE activations SET paid_amount = retail_price")

    cur.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)")

    cur.execute("CREATE INDEX IF NOT EXISTS idx_act_date ON activations(operation_date)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_wallet_date ON wallet_transactions(created_at)")

    user = cur.execute("SELECT id FROM users WHERE username = ?", ("admin",)).fetchone()
    if not user:
        cur.execute(
            "INSERT INTO users(username, password) VALUES (?, ?)",
            ("admin", generate_password_hash("admin123"))
        )

    count = cur.execute("SELECT COUNT(*) AS c FROM plans").fetchone()["c"]
    if count == 0:
        for category, name, retail_price in PLANS:
            cur.execute("""
                INSERT INTO plans (category, name, retail_price, profit, wholesale_price)
                VALUES (?, ?, ?, ?, ?)
            """, (category, name, retail_price, PROFIT, retail_price - PROFIT))

    con.commit()
    con.close()


def expiry_from(start_date):
    return (datetime.strptime(start_date, "%Y-%m-%d")
            + timedelta(days=SUBSCRIPTION_DAYS)).strftime("%Y-%m-%d")


def status_badge(expiry_date):
    """شارة حالة الاشتراك حسب تاريخ الانتهاء."""
    exp = parse_date(expiry_date)
    if not exp:
        return '<span class="badge bg-secondary">غير معروف</span>'
    days_left = (datetime.strptime(exp, "%Y-%m-%d").date() - date.today()).days
    if days_left < 0:
        return f'<span class="badge bg-danger">منتهي منذ {-days_left} يوم</span>'
    if days_left <= EXPIRY_WARNING_DAYS:
        return f'<span class="badge bg-warning text-dark">باقي {days_left} يوم</span>'
    return f'<span class="badge bg-success">فعّال ({days_left} يوم)</span>'


# رصيد دين المشترك (يُستخدم داخل استعلامات فيها s = جدول المشتركين)
DEBT_SUBQUERY = """
    SELECT COALESCE(SUM(CASE WHEN d.kind = 'دين' THEN d.amount ELSE -d.amount END), 0)
    FROM debts d WHERE d.subscriber_id = s.id
"""


def get_setting(con, key, default=""):
    row = con.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
    return row["value"] if row and row["value"] is not None else default


def set_setting(con, key, value):
    con.execute("INSERT INTO settings(key, value) VALUES (?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value = excluded.value", (key, value))


def payment_badge(kind):
    if kind == "دين":
        return '<span class="badge bg-danger-subtle text-danger-emphasis">دين</span>'
    return '<span class="badge bg-success-subtle text-success-emphasis">نقد</span>'


# ---------------------------------------------------------------------------
# واتساب
# ---------------------------------------------------------------------------

def wa_phone(phone):
    """تحويل رقم الهاتف لصيغة واتساب الدولية (العراق 964 افتراضياً)."""
    digits = "".join(ch for ch in str(phone or "") if ch.isdigit())
    if digits.startswith("00"):
        digits = digits[2:]
    if digits.startswith("0") and len(digits) == 11:
        digits = "964" + digits[1:]
    elif digits.startswith("7") and len(digits) == 10:
        digits = "964" + digits
    return digits if len(digits) >= 10 else ""


def wa_link(phone, text):
    from urllib.parse import quote
    num = wa_phone(phone)
    return f"https://wa.me/{num}?text={quote(text)}" if num else ""


def wa_button(phone, text, label="واتساب", size="btn-sm"):
    link = wa_link(phone, text)
    content = f'<i class="bi bi-whatsapp"></i>{" " + label if label else ""}'
    if not link:
        return f'<button type="button" class="btn {size} btn-wa" disabled title="لا يوجد رقم هاتف">{content}</button>'
    return f'<a href="{esc(link)}" target="_blank" rel="noopener" class="btn {size} btn-wa" title="مراسلة على واتساب">{content}</a>'


def _sign(shop):
    return f"\n\n{shop}" if shop else ""


def msg_reminder(s, shop=""):
    exp = parse_date(s["expiry_date"]) or ""
    if exp and exp < today_str():
        line = f"نود إعلامك بأن اشتراكك ({s['plan_name']}) انتهى بتاريخ {exp}."
    else:
        line = f"نود تذكيرك بأن اشتراكك ({s['plan_name']}) ينتهي بتاريخ {exp}."
    return (f"مرحباً {s['name']} 🌹\n{line}\n"
            f"اليوزر: {s['username']}\nللتجديد يرجى التواصل معنا." + _sign(shop))


def msg_debt(s, debt, shop=""):
    if not debt or debt <= 0:
        return f"مرحباً {s['name']} 🌹\nلا يوجد عليك أي مبلغ متبقي، شكراً لك." + _sign(shop)
    return (f"مرحباً {s['name']} 🌹\nنود تذكيرك بوجود مبلغ متبقي بذمتك قدره {money(debt)} د.ع "
            f"عن اشتراك الإنترنت (اليوزر: {s['username']}).\nيرجى التسديد بأقرب وقت، وشكراً لك." + _sign(shop))


def msg_activated(s, shop=""):
    return (f"مرحباً {s['name']} 🌹\nتم تفعيل اشتراكك ({s['plan_name']}) بنجاح ✅\n"
            f"اليوزر: {s['username']}\nينتهي بتاريخ: {s['expiry_date']}" + _sign(shop))


def msg_confirm(s, op, debt, shop=""):
    text = (f"مرحباً {s['name']} 🌹\nتم {op['operation_type']} اشتراكك ({op['plan_name']}) بنجاح ✅\n"
            f"اليوزر: {s['username']}\nينتهي بتاريخ: {s['expiry_date']}\n"
            f"المبلغ: {money(op['retail_price'])} د.ع")
    if op["payment_type"] == "دين":
        text += f"\nالمدفوع: {money(op['paid_amount'])} د.ع\nالمتبقي بذمتك: {money(debt)} د.ع"
    return text + _sign(shop)


# ---------------------------------------------------------------------------
# الحماية: تسجيل الدخول + رمز CSRF
# ---------------------------------------------------------------------------

def login_required(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return func(*args, **kwargs)
    return wrapper


def csrf_token():
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_hex(16)
    return session["csrf_token"]


app.jinja_env.globals["csrf_token"] = csrf_token


@app.before_request
def check_csrf():
    if request.method == "POST":
        sent = request.form.get("csrf_token", "")
        if not sent or not secrets.compare_digest(sent, session.get("csrf_token", "")):
            abort(400, "انتهت صلاحية الصفحة، أعد تحميلها وحاول مرة ثانية.")


def wallet_balance(con=None):
    own = con is None
    con = con or db()
    result = con.execute("""
        SELECT COALESCE(SUM(
            CASE
                WHEN transaction_type = 'شحن' THEN amount
                WHEN transaction_type = 'خصم' THEN -amount
                ELSE 0
            END
        ), 0) AS balance
        FROM wallet_transactions
    """).fetchone()
    if own:
        con.close()
    return result["balance"]


# ---------------------------------------------------------------------------
# القالب العام
# ---------------------------------------------------------------------------

BASE_HTML = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
    <meta name="csrf-token" content="{{ csrf_token() }}">
    <meta name="theme-color" content="#4f46e5">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <title>محاسبة الإنترنت</title>
    <script>
        // تطبيق الوضع الليلي قبل عرض الصفحة
        (function () {
            let t = null;
            try { t = localStorage.getItem('theme'); } catch (e) {}
            if (!t) t = matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
            document.documentElement.setAttribute('data-bs-theme', t);
        })();
    </script>

    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.rtl.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css" rel="stylesheet">

    <style>
        :root, [data-bs-theme="light"] {
            --bg: #f4f6fb;
            --card: #ffffff;
            --soft: #f1f3f9;
            --text: #0f172a;
            --muted: #64748b;
            --line: #e5e7ef;
            --primary: #4f46e5;
            --primary-2: #7c3aed;
            --primary-soft: #eef2ff;
            --shadow: 0 1px 2px rgba(15,23,42,.04), 0 8px 24px rgba(15,23,42,.06);
            --bs-body-bg: var(--bg);
            --bs-body-color: var(--text);
            --bs-primary: #4f46e5;
            --bs-primary-rgb: 79,70,229;
        }
        [data-bs-theme="dark"] {
            --bg: #0b1120;
            --card: #131c2e;
            --soft: #1a2438;
            --text: #e2e8f0;
            --muted: #94a3b8;
            --line: #243049;
            --primary: #818cf8;
            --primary-2: #a78bfa;
            --primary-soft: rgba(129,140,248,.14);
            --shadow: 0 1px 2px rgba(0,0,0,.3), 0 8px 24px rgba(0,0,0,.25);
            --bs-body-bg: var(--bg);
            --bs-body-color: var(--text);
            --bs-border-color: var(--line);
        }

        * { -webkit-tap-highlight-color: transparent; }
        body {
            background: var(--bg);
            color: var(--text);
            font-family: 'Tajawal', Tahoma, sans-serif;
            font-size: 15.5px;
        }
        h2 { font-weight: 800; font-size: 1.6rem; }
        h3, h4, h5 { font-weight: 700; }
        h5 { font-size: 1.08rem; margin-bottom: 1rem; }
        .text-muted { color: var(--muted) !important; }
        a { color: var(--primary); }

        /* ---------- الهيكل ---------- */
        .layout { display: flex; min-height: 100vh; }
        .sidebar {
            width: 260px; flex-shrink: 0;
            background: var(--card);
            border-left: 1px solid var(--line);
            padding: 22px 16px;
            position: sticky; top: 0; height: 100vh; overflow-y: auto;
        }
        .brand {
            display: flex; align-items: center; gap: 10px;
            font-weight: 800; font-size: 1.2rem; color: var(--text);
            padding: 4px 8px 22px;
        }
        .brand .logo {
            width: 40px; height: 40px; border-radius: 12px;
            display: grid; place-items: center; color: #fff; font-size: 1.2rem;
            background: linear-gradient(135deg, var(--primary), var(--primary-2));
        }
        .nav-side a {
            display: flex; align-items: center; gap: 12px;
            color: var(--muted); text-decoration: none; font-weight: 500;
            padding: 11px 14px; border-radius: 12px; margin-bottom: 4px;
            transition: .15s;
        }
        .nav-side a i { font-size: 1.15rem; }
        .nav-side a:hover { background: var(--soft); color: var(--text); }
        .nav-side a.active { background: var(--primary-soft); color: var(--primary); font-weight: 700; }
        .side-foot { border-top: 1px solid var(--line); margin-top: 18px; padding-top: 14px; }
        .who { display: flex; align-items: center; gap: 10px; padding: 6px 8px 10px; color: var(--muted); font-size: .92rem; }
        .avatar {
            width: 34px; height: 34px; border-radius: 50%;
            background: var(--primary-soft); color: var(--primary);
            display: grid; place-items: center; font-weight: 800;
        }

        .main { flex: 1; min-width: 0; padding: 28px 32px 40px; }
        .main.full { padding: 0; }

        .topbar {
            display: none;
            position: sticky; top: 0; z-index: 1020;
            background: color-mix(in srgb, var(--card) 88%, transparent);
            backdrop-filter: blur(12px);
            border-bottom: 1px solid var(--line);
            padding: 12px 16px; padding-top: calc(12px + env(safe-area-inset-top));
            align-items: center; justify-content: space-between;
        }
        .topbar .title { font-weight: 800; font-size: 1.1rem; }
        .icon-btn {
            width: 40px; height: 40px; border-radius: 12px; border: 1px solid var(--line);
            background: var(--card); color: var(--text);
            display: inline-grid; place-items: center; font-size: 1.1rem;
        }

        .bottom-nav {
            display: none;
            position: fixed; bottom: 0; left: 0; right: 0; z-index: 1030;
            background: var(--card); border-top: 1px solid var(--line);
            padding: 6px 4px calc(6px + env(safe-area-inset-bottom));
            box-shadow: 0 -6px 20px rgba(15,23,42,.06);
        }
        .bottom-nav a, .bottom-nav button {
            flex: 1; display: flex; flex-direction: column; align-items: center; gap: 2px;
            color: var(--muted); text-decoration: none; font-size: .72rem; font-weight: 600;
            padding: 6px 0; background: none; border: 0;
        }
        .bottom-nav i { font-size: 1.35rem; }
        .bottom-nav .active { color: var(--primary); }

        /* ---------- البطاقات ---------- */
        .card {
            background: var(--card); color: var(--text);
            border: 1px solid var(--line); border-radius: 18px;
            box-shadow: var(--shadow);
        }
        .stat-card { min-height: 118px; position: relative; overflow: hidden; }
        .stat-card h6 { color: var(--muted); font-weight: 600; font-size: .88rem; margin-bottom: .6rem; }
        .stat-card h2, .stat-card h3 { font-size: 1.55rem; font-weight: 800; margin: 0; letter-spacing: -.3px; }
        .stat-card small { display: block; margin-top: 6px; }
        .stat-card .stat-icon {
            position: absolute; top: 16px; left: 16px;
            width: 40px; height: 40px; border-radius: 12px;
            display: grid; place-items: center; font-size: 1.15rem;
            background: var(--primary-soft); color: var(--primary);
        }
        .balance-card {
            background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%) !important;
            color: #fff !important; border: 0;
        }
        .balance-card h6, .balance-card small { color: rgba(255,255,255,.8) !important; }
        .balance-card .stat-icon { background: rgba(255,255,255,.18); color: #fff; }

        /* ---------- النماذج والأزرار ---------- */
        .form-label { font-weight: 600; font-size: .9rem; color: var(--muted); margin-bottom: 6px; }
        .form-control, .form-select {
            border-radius: 12px; padding: 10px 14px; min-height: 46px;
            background-color: var(--soft); border: 1px solid transparent; color: var(--text);
        }
        .form-control:focus, .form-select:focus {
            background-color: var(--card); border-color: var(--primary);
            box-shadow: 0 0 0 4px var(--primary-soft);
        }
        .form-control-sm { min-height: 38px; padding: 6px 10px; }
        .btn { border-radius: 12px; font-weight: 600; padding: 10px 18px; }
        .btn-sm { border-radius: 10px; padding: 6px 12px; font-size: .85rem; }
        .btn-primary { background: var(--primary); border-color: var(--primary); }
        .btn-primary:hover { background: #4338ca; border-color: #4338ca; }
        .btn-wa { background: #25D366; border-color: #25D366; color: #fff; }
        .btn-wa:hover, .btn-wa:focus { background: #1ebe5b; border-color: #1ebe5b; color: #fff; }
        .btn-wa:disabled { background: #a7d7b8; border-color: #a7d7b8; color: #fff; }
        .pay-options { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
        .pay-opt {
            display: flex; flex-direction: column; align-items: center; gap: 2px; cursor: pointer;
            padding: 16px 10px; border-radius: 16px; border: 2px solid var(--line); background: var(--soft);
            text-align: center; transition: .15s;
        }
        .pay-opt i { font-size: 1.7rem; color: #10b981; }
        .pay-opt.debt i { color: #ef4444; }
        .pay-opt b { font-size: 1.05rem; }
        .pay-opt small { color: var(--muted); }
        .btn-check:checked + .pay-opt { border-color: #10b981; background: rgba(16,185,129,.08); }
        .btn-check:checked + .pay-opt.debt { border-color: #ef4444; background: rgba(239,68,68,.08); }
        .badge { font-weight: 600; padding: .45em .8em; border-radius: 999px; }
        .alert { border-radius: 14px; border: 0; }

        /* ---------- الجداول ---------- */
        .table { --bs-table-bg: transparent; color: var(--text); vertical-align: middle; margin: 0; }
        .table > :not(caption) > * > * { border-color: var(--line); padding: 12px 10px; color: var(--text); }
        .table thead th {
            color: var(--muted); font-weight: 700; font-size: .82rem;
            background: var(--soft) !important; border: 0; white-space: nowrap;
        }
        .table thead th:first-child { border-radius: 0 10px 10px 0; }
        .table thead th:last-child { border-radius: 10px 0 0 10px; }
        .table-striped > tbody > tr:nth-of-type(odd) > * { --bs-table-bg-type: transparent; }
        .table-hover > tbody > tr:hover > * { --bs-table-bg-state: var(--soft); }

        /* ---------- الموبايل ---------- */
        @media (max-width: 991px) {
            .sidebar { display: none; }
            .topbar { display: flex; }
            .bottom-nav { display: flex; }
            .main { padding: 16px 14px calc(96px + env(safe-area-inset-bottom)); }
            .main.full { padding: 0; }
            h2 { font-size: 1.3rem; }
            .main > h2:first-child, .page-head h2 { display: none; }
            .card.p-4 { padding: 18px !important; }
            .stat-card { min-height: 104px; padding: 16px !important; }
            .stat-card h2, .stat-card h3 { font-size: 1.2rem; }
            .stat-card .stat-icon { width: 34px; height: 34px; top: 12px; left: 12px; font-size: 1rem; }
            .filter-bar > * { max-width: none !important; flex: 1 1 100%; }
            .page-head { justify-content: flex-end !important; }
            .page-head > div, .page-head > div .btn { flex: 1; }
            .page-head > div { display: flex; gap: 8px; }

            /* تحويل الجداول إلى بطاقات */
            .table thead { display: none; }
            .table, .table tbody, .table tr, .table td { display: block; width: 100%; }
            .table tr {
                background: var(--soft); border-radius: 14px;
                padding: 10px 14px; margin-bottom: 10px;
            }
            .table > :not(caption) > * > * {
                border: 0; padding: 6px 0; background: transparent !important; box-shadow: none !important;
            }
            .table td[data-label] {
                display: flex; justify-content: space-between; align-items: center; gap: 12px;
                text-align: left;
            }
            .table td[data-label]::before {
                content: attr(data-label); color: var(--muted);
                font-weight: 600; font-size: .85rem; text-align: right; flex-shrink: 0;
            }
            .table td[data-label=""]::before { content: none; }
            .table td[data-label=""] { justify-content: flex-start; gap: 6px; flex-wrap: wrap; padding-top: 10px; }
            .table td[colspan] { text-align: center; background: transparent; }
            .table td .form-control { max-width: 60%; background-color: var(--card); border-color: var(--line); }
            .filter-bar > .btn { flex: 1 1 40%; }
            .table td[data-label]:has(.form-select) { flex-direction: column; align-items: stretch; gap: 6px; }
            .table td .form-select { width: 100%; max-width: none; background-color: var(--card); }
            .table-responsive { overflow: visible; }
        }

        @media print {
            .sidebar, .topbar, .bottom-nav, .no-print, form { display: none !important; }
            .main { padding: 0; }
            .card { box-shadow: none; border: 1px solid #ddd; }
        }

        /* صفحة الدخول */
        .login-wrap {
            min-height: 100vh; display: grid; place-items: center; padding: 20px;
            background:
                radial-gradient(1000px 500px at 100% 0%, rgba(124,58,237,.18), transparent 60%),
                radial-gradient(800px 500px at 0% 100%, rgba(79,70,229,.18), transparent 60%),
                var(--bg);
        }
        .login-card { width: 100%; max-width: 410px; padding: 34px 28px !important; }
    </style>
</head>

<body>
{% set titles = {
    'dashboard': 'الرئيسية', 'subscribers': 'المشتركون', 'subscriber_detail': 'تفاصيل المشترك',
    'wallet': 'المحفظة', 'plans': 'الباقات والأسعار', 'reports': 'التقارير',
    'settings': 'الإعدادات', 'debts': 'حساب الديون', 'import_users': 'استيراد اليوزرات', 'activate': 'تفعيل اشتراك', 'transfer': 'ترحيل اشتراك'
} %}
{% set ep = request.endpoint %}

{% if session.get('user_id') %}
<div class="layout">
    <aside class="sidebar">
        <div class="brand"><span class="logo"><i class="bi bi-router"></i></span> محاسبة الإنترنت</div>
        <nav class="nav-side">
            <a href="{{ url_for('dashboard') }}" class="{{ 'active' if ep == 'dashboard' }}"><i class="bi bi-grid-1x2"></i> الرئيسية</a>
            <a href="{{ url_for('subscribers') }}" class="{{ 'active' if ep in ['subscribers','subscriber_detail','activate','transfer','import_users'] }}"><i class="bi bi-people"></i> المشتركون</a>
            <a href="{{ url_for('debts') }}" class="{{ 'active' if ep == 'debts' }}"><i class="bi bi-journal-text"></i> الديون</a>
            <a href="{{ url_for('wallet') }}" class="{{ 'active' if ep == 'wallet' }}"><i class="bi bi-wallet2"></i> المحفظة</a>
            <a href="{{ url_for('plans') }}" class="{{ 'active' if ep == 'plans' }}"><i class="bi bi-box-seam"></i> الباقات والأسعار</a>
            <a href="{{ url_for('reports') }}" class="{{ 'active' if ep == 'reports' }}"><i class="bi bi-bar-chart-line"></i> التقارير</a>
            <a href="{{ url_for('settings') }}" class="{{ 'active' if ep == 'settings' }}"><i class="bi bi-gear"></i> الإعدادات</a>
        </nav>
        <div class="side-foot nav-side">
            <div class="who"><span class="avatar">{{ (session.get('username') or '?')[:1]|upper }}</span> {{ session.get('username') }}</div>
            <a href="#" onclick="toggleTheme();return false;"><i class="bi bi-moon-stars"></i> الوضع الليلي</a>
            <a href="{{ url_for('logout') }}"><i class="bi bi-box-arrow-left"></i> تسجيل الخروج</a>
        </div>
    </aside>

    <div class="flex-grow-1" style="min-width:0">
        <header class="topbar">
            <div class="d-flex align-items-center gap-2">
                <span class="brand p-0"><span class="logo" style="width:34px;height:34px;font-size:1rem"><i class="bi bi-router"></i></span></span>
                <span class="title">{{ titles.get(ep, 'محاسبة الإنترنت') }}</span>
            </div>
            <div class="d-flex gap-2">
                <button class="icon-btn" onclick="toggleTheme()" aria-label="الوضع الليلي"><i class="bi bi-moon-stars"></i></button>
            </div>
        </header>

        <main class="main">
            {% with messages = get_flashed_messages(with_categories=true) %}
                {% for category, message in messages %}
                    <div class="alert alert-{{ 'info' if category == 'message' else category }} alert-dismissible fade show">
                        {{ message }}
                        <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                    </div>
                {% endfor %}
            {% endwith %}

            {{ body|safe }}
        </main>
    </div>
</div>

<nav class="bottom-nav">
    <a href="{{ url_for('dashboard') }}" class="{{ 'active' if ep == 'dashboard' }}"><i class="bi bi-grid-1x2{{ '-fill' if ep == 'dashboard' }}"></i>الرئيسية</a>
    <a href="{{ url_for('subscribers') }}" class="{{ 'active' if ep in ['subscribers','subscriber_detail','activate','transfer'] }}"><i class="bi bi-people{{ '-fill' if ep in ['subscribers','subscriber_detail','activate','transfer'] }}"></i>المشتركون</a>
    <a href="{{ url_for('wallet') }}" class="{{ 'active' if ep == 'wallet' }}"><i class="bi bi-wallet{{ '-fill' if ep == 'wallet' else '2' }}"></i>المحفظة</a>
    <a href="{{ url_for('debts') }}" class="{{ 'active' if ep == 'debts' }}"><i class="bi bi-journal-text"></i>الديون</a>
    <button data-bs-toggle="offcanvas" data-bs-target="#more" class="{{ 'active' if ep in ['plans','settings','reports'] }}"><i class="bi bi-three-dots"></i>المزيد</button>
</nav>

<div class="offcanvas offcanvas-bottom" tabindex="-1" id="more" style="height:auto;border-radius:22px 22px 0 0;background:var(--card)">
    <div class="offcanvas-body nav-side pb-4">
        <div class="who mb-2"><span class="avatar">{{ (session.get('username') or '?')[:1]|upper }}</span> {{ session.get('username') }}</div>
        <a href="{{ url_for('reports') }}"><i class="bi bi-bar-chart-line"></i> التقارير</a>
        <a href="{{ url_for('plans') }}"><i class="bi bi-box-seam"></i> الباقات والأسعار</a>
        <a href="{{ url_for('settings') }}"><i class="bi bi-gear"></i> الإعدادات والنسخ الاحتياطي</a>
        <a href="#" onclick="toggleTheme();return false;"><i class="bi bi-moon-stars"></i> الوضع الليلي</a>
        <a href="{{ url_for('logout') }}" class="text-danger"><i class="bi bi-box-arrow-left"></i> تسجيل الخروج</a>
    </div>
</div>
{% else %}
<main class="main full">
    {% with messages = get_flashed_messages(with_categories=true) %}
        {% for category, message in messages %}
            <div class="alert alert-{{ 'info' if category == 'message' else category }} m-3">{{ message }}</div>
        {% endfor %}
    {% endwith %}
    {{ body|safe }}
</main>
{% endif %}

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
<script>
    function toggleTheme() {
        const el = document.documentElement;
        const t = el.getAttribute('data-bs-theme') === 'dark' ? 'light' : 'dark';
        el.setAttribute('data-bs-theme', t);
        try { localStorage.setItem('theme', t); } catch (e) {}
    }
    // إضافة رمز الحماية لكل نموذج POST تلقائياً
    const token = document.querySelector('meta[name="csrf-token"]').content;
    document.querySelectorAll('form[method="post"]').forEach(f => {
        const i = document.createElement('input');
        i.type = 'hidden'; i.name = 'csrf_token'; i.value = token;
        f.appendChild(i);
    });
    // تأكيد قبل الحذف
    document.querySelectorAll('[data-confirm]').forEach(el => {
        el.addEventListener('submit', e => { if (!confirm(el.dataset.confirm)) e.preventDefault(); });
    });
    // عناوين الأعمدة لعرض الجداول كبطاقات على الموبايل
    document.querySelectorAll('table.table').forEach(t => {
        const heads = [...t.querySelectorAll('thead th')].map(th => th.textContent.trim());
        if (!heads.length) return;
        t.querySelectorAll('tbody tr').forEach(tr => {
            [...tr.children].forEach((td, i) => {
                if (td.tagName === 'TD' && !td.hasAttribute('colspan')) td.dataset.label = heads[i] || '';
            });
        });
    });
</script>
</body>
</html>
"""


def page(body):
    return render_template_string(BASE_HTML, body=body)


LOGIN_PAGE = """
<div class="login-wrap">
    <div class="card login-card">
        <div class="text-center mb-4">
            <div class="brand justify-content-center p-0 mb-3">
                <span class="logo" style="width:60px;height:60px;font-size:1.7rem;border-radius:18px"><i class="bi bi-router"></i></span>
            </div>
            <h3 class="mb-1">محاسبة الإنترنت</h3>
            <div class="text-muted">سجّل دخولك للمتابعة</div>
        </div>

        <form method="post">
            <label class="form-label">اسم المستخدم</label>
            <input name="username" class="form-control mb-3" required autofocus autocomplete="username">

            <label class="form-label">كلمة المرور</label>
            <input name="password" type="password" class="form-control mb-4" required autocomplete="current-password">

            <button class="btn btn-primary w-100 py-3">دخول <i class="bi bi-arrow-left ms-1"></i></button>
        </form>

        {% if show_default %}
        <div class="text-muted mt-4 text-center small">
            الدخول الافتراضي: admin / admin123 — غيّرها من الإعدادات بعد الدخول
        </div>
        {% endif %}
    </div>
</div>
"""


# ---------------------------------------------------------------------------
# تسجيل الدخول
# ---------------------------------------------------------------------------

@app.route("/", methods=["GET", "POST"])
def login():
    if "user_id" in session and request.method == "GET":
        return redirect(url_for("dashboard"))

    con = db()
    if request.method == "POST":
        username = form_text("username")
        password = request.form.get("password", "")

        user = con.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()

        if user and check_password_hash(user["password"], password):
            con.close()
            token = session.get("csrf_token")
            session.clear()
            session.permanent = True
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            if token:
                session["csrf_token"] = token
            return redirect(url_for("dashboard"))

        flash("اسم المستخدم أو كلمة المرور غير صحيحة", "danger")

    # إظهار بيانات الدخول الافتراضية فقط إذا لم يتم تغييرها
    admin = con.execute("SELECT password FROM users WHERE username = 'admin'").fetchone()
    con.close()
    show_default = bool(admin and check_password_hash(admin["password"], "admin123"))

    return page(render_template_string(LOGIN_PAGE, show_default=show_default))


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# لوحة التحكم
# ---------------------------------------------------------------------------

@app.route("/dashboard")
@login_required
def dashboard():
    con = db()
    month_start = date.today().replace(day=1).strftime("%Y-%m-%d")

    subscribers_count = con.execute("SELECT COUNT(*) AS c FROM subscribers").fetchone()["c"]

    stats = con.execute("""
        SELECT
            SUM(CASE WHEN operation_type = 'تفعيل' THEN 1 ELSE 0 END) AS activations,
            SUM(CASE WHEN operation_type = 'ترحيل' THEN 1 ELSE 0 END) AS transfers,
            COALESCE(SUM(profit), 0) AS total_profit
        FROM activations
    """).fetchone()

    month = con.execute("""
        SELECT COUNT(*) AS ops, COALESCE(SUM(profit), 0) AS profit,
               COALESCE(SUM(retail_price), 0) AS sales
        FROM activations WHERE operation_date >= ?
    """, (month_start,)).fetchone()

    warn_until = (date.today() + timedelta(days=EXPIRY_WARNING_DAYS)).strftime("%Y-%m-%d")
    expiring = con.execute("""
        SELECT s.*, p.name AS plan_name
        FROM subscribers s JOIN plans p ON p.id = s.plan_id
        WHERE s.expiry_date <= ?
        ORDER BY s.expiry_date
        LIMIT 50
    """, (warn_until,)).fetchall()

    recent = con.execute("""
        SELECT a.*, s.username AS subscriber_username, p.name AS plan_name
        FROM activations a
        JOIN subscribers s ON s.id = a.subscriber_id
        JOIN plans p ON p.id = a.plan_id
        ORDER BY a.id DESC LIMIT 8
    """).fetchall()

    balance = wallet_balance(con)
    total_debt = con.execute(
        "SELECT COALESCE(SUM(CASE WHEN kind = 'دين' THEN amount ELSE -amount END), 0) AS t FROM debts"
    ).fetchone()["t"]
    shop = get_setting(con, "shop_name")
    con.close()

    expiring_rows = "".join(f"""
        <tr>
            <td>{esc(s['name'])}</td>
            <td>{esc(s['username'])}</td>
            <td>{esc(s['phone'])}</td>
            <td>{esc(s['plan_name'])}</td>
            <td>{esc(s['expiry_date'])}</td>
            <td>{status_badge(s['expiry_date'])}</td>
            <td><div class="d-flex gap-1">
                <a href="{url_for('transfer', subscriber_id=s['id'])}" class="btn btn-sm btn-warning">ترحيل</a>
                {wa_button(s['phone'], msg_reminder(s, shop), 'تذكير')}
            </div></td>
        </tr>""" for s in expiring) or '<tr><td colspan="7" class="text-muted text-center">لا توجد اشتراكات منتهية أو قريبة الانتهاء 👍</td></tr>'

    recent_rows = "".join(f"""
        <tr>
            <td>{esc(a['operation_type'])}</td>
            <td>{esc(a['subscriber_username'])}</td>
            <td>{esc(a['plan_name'])}</td>
            <td>{money(a['amount'])}</td>
            <td>{esc(a['operation_date'])}</td>
        </tr>""" for a in recent) or '<tr><td colspan="5" class="text-muted text-center">لا توجد عمليات بعد</td></tr>'

    body = f"""
    <h2 class="mb-4">لوحة التحكم</h2>

    <div class="row g-3 mb-4">
        <div class="col-6 col-md-3">
            <div class="card balance-card stat-card p-4">
                <span class="stat-icon"><i class="bi bi-wallet2"></i></span>
                <h6>رصيد المحفظة</h6>
                <h2>{money(balance)} د.ع</h2>
            </div>
        </div>
        <div class="col-6 col-md-3">
            <div class="card stat-card p-4">
                <span class="stat-icon"><i class="bi bi-people"></i></span>
                <h6>عدد المشتركين</h6>
                <h2>{subscribers_count}</h2>
            </div>
        </div>
        <div class="col-6 col-md-3">
            <div class="card stat-card p-4">
                <span class="stat-icon" style="background:rgba(16,185,129,.14);color:#10b981"><i class="bi bi-graph-up-arrow"></i></span>
                <h6>أرباح هذا الشهر</h6>
                <h2 class="text-success">{money(month['profit'])} د.ع</h2>
                <small class="text-muted">{month['ops']} عملية — مبيعات {money(month['sales'])} د.ع</small>
            </div>
        </div>
        <div class="col-6 col-md-3">
            <a href="{url_for('debts')}" class="text-decoration-none">
            <div class="card stat-card p-4">
                <span class="stat-icon" style="background:rgba(239,68,68,.14);color:#ef4444"><i class="bi bi-journal-text"></i></span>
                <h6>ديون على المشتركين</h6>
                <h2 class="text-danger">{money(total_debt)} د.ع</h2>
                <small class="text-muted">إجمالي الأرباح: {money(stats['total_profit'])} د.ع</small>
            </div>
            </a>
        </div>
    </div>

    <div class="card p-4 mb-4">
        <h5><i class="bi bi-alarm text-warning"></i> اشتراكات منتهية أو تنتهي خلال {EXPIRY_WARNING_DAYS} أيام</h5>
        <div class="table-responsive">
            <table class="table table-hover">
                <thead><tr><th>الاسم</th><th>اليوزر</th><th>الهاتف</th><th>الاشتراك</th>
                <th>تاريخ الانتهاء</th><th>الحالة</th><th></th></tr></thead>
                <tbody>{expiring_rows}</tbody>
            </table>
        </div>
    </div>

    <div class="card p-4">
        <h5><i class="bi bi-clock-history text-primary"></i> آخر العمليات</h5>
        <div class="table-responsive">
            <table class="table table-striped">
                <thead><tr><th>العملية</th><th>اليوزر</th><th>الاشتراك</th><th>المبلغ</th><th>التاريخ</th></tr></thead>
                <tbody>{recent_rows}</tbody>
            </table>
        </div>
    </div>
    """
    return page(body)


# ---------------------------------------------------------------------------
# المحفظة
# ---------------------------------------------------------------------------

@app.route("/wallet", methods=["GET", "POST"])
@login_required
def wallet():
    if request.method == "POST":
        amount = parse_amount(request.form.get("amount"))
        kind = request.form.get("kind", "شحن")
        notes = form_text("notes")

        if kind not in ("شحن", "خصم"):
            kind = "شحن"
        if not amount or amount <= 0:
            flash("أدخل مبلغاً صحيحاً أكبر من صفر", "danger")
            return redirect(url_for("wallet"))

        con = db()
        if kind == "خصم" and wallet_balance(con) < amount:
            con.close()
            flash("الرصيد غير كافٍ لهذا الخصم", "danger")
            return redirect(url_for("wallet"))

        con.execute("""
            INSERT INTO wallet_transactions (transaction_type, amount, notes, created_at)
            VALUES (?, ?, ?, ?)
        """, (kind, amount, notes or ("خصم يدوي" if kind == "خصم" else ""), now_str()))
        con.commit()
        con.close()

        flash("تم شحن المحفظة بنجاح" if kind == "شحن" else "تم خصم المبلغ من المحفظة", "success")
        return redirect(url_for("wallet"))

    con = db()
    transactions = con.execute("SELECT * FROM wallet_transactions ORDER BY id DESC LIMIT 500").fetchall()
    balance = wallet_balance(con)
    con.close()

    rows = ""
    for item in transactions:
        charge = item["transaction_type"] == "شحن"
        rows += f"""
        <tr>
            <td>{esc(item['transaction_type'])}</td>
            <td class="{'text-success' if charge else 'text-danger'}">{'+' if charge else '-'}{money(item['amount'])}</td>
            <td>{esc(item['notes'])}</td>
            <td>{esc(item['created_at'])}</td>
        </tr>"""

    body = f"""
    <h2 class="mb-4">المحفظة</h2>

    <div class="card p-4 mb-4">
        <h4>الرصيد الحالي: <span class="text-success">{money(balance)} د.ع</span></h4>

        <form method="post" class="row g-3 mt-2">
            <div class="col-md-2">
                <label class="form-label">نوع الحركة</label>
                <select name="kind" class="form-select">
                    <option value="شحن">شحن (إضافة)</option>
                    <option value="خصم">خصم يدوي</option>
                </select>
            </div>
            <div class="col-md-3">
                <label class="form-label">المبلغ</label>
                <input type="number" name="amount" min="1" class="form-control" required>
            </div>
            <div class="col-md-5">
                <label class="form-label">ملاحظات</label>
                <input name="notes" class="form-control" placeholder="مثال: شحن من الإدارة">
            </div>
            <div class="col-md-2 d-flex align-items-end">
                <button class="btn btn-success w-100">حفظ</button>
            </div>
        </form>
    </div>

    <div class="card p-4">
        <h5>حركات المحفظة <small class="text-muted">(آخر 500 حركة)</small></h5>
        <div class="table-responsive">
            <table class="table table-striped">
                <thead><tr><th>النوع</th><th>المبلغ</th><th>الملاحظات</th><th>التاريخ</th></tr></thead>
                <tbody>{rows or '<tr><td colspan="4" class="text-muted text-center">لا توجد حركات</td></tr>'}</tbody>
            </table>
        </div>
    </div>
    """
    return page(body)


# ---------------------------------------------------------------------------
# المشتركون
# ---------------------------------------------------------------------------

def plan_options(plans, selected_id=None):
    html = ""
    for plan in plans:
        sel = "selected" if selected_id is not None and plan["id"] == selected_id else ""
        html += f"""
        <option value="{plan['id']}" data-retail="{plan['retail_price']}" {sel}>
            {esc(plan['category'])} - {esc(plan['name'])}
            | البيع {money(plan['retail_price'])}
            | الجملة {money(plan['wholesale_price'])}
        </option>"""
    return html


def read_subscriber_form(con):
    """قراءة بيانات المشترك من النموذج والتحقق منها. ترجع (data, error)."""
    data = {
        "name": form_text("name"),
        "username": form_text("username"),
        "phone": form_text("phone"),
        "notes": form_text("notes"),
        "activation_date": parse_date(request.form.get("activation_date")),
        "expiry_date": parse_date(request.form.get("expiry_date")),
    }
    try:
        data["plan_id"] = int(request.form.get("plan_id", ""))
    except ValueError:
        return None, "اختر نوع الاشتراك"

    if not data["name"] or not data["username"]:
        return None, "اسم المشترك واليوزر مطلوبان"
    if not data["activation_date"]:
        return None, "تاريخ التفعيل غير صحيح"
    if not con.execute("SELECT 1 FROM plans WHERE id = ?", (data["plan_id"],)).fetchone():
        return None, "الباقة غير موجودة"
    if not data["expiry_date"]:
        data["expiry_date"] = expiry_from(data["activation_date"])
    return data, None


def subscriber_form_fields(plans_list, s=None, with_expiry=False):
    """حقول نموذج المشترك (للإضافة والتعديل)."""
    v = (lambda k: esc(s[k]) if s else "")
    expiry = f"""
            <div class="col-6 col-md-2">
                <label class="form-label">تاريخ الانتهاء</label>
                <input type="date" name="expiry_date" value="{v('expiry_date')}" class="form-control">
            </div>""" if with_expiry else ""
    return f"""
            <div class="col-md-4">
                <label class="form-label">اسم المشترك</label>
                <input name="name" value="{v('name')}" class="form-control" required>
            </div>
            <div class="col-md-4">
                <label class="form-label">يوزر المشترك</label>
                <input name="username" value="{v('username')}" class="form-control" required>
            </div>
            <div class="col-md-4">
                <label class="form-label">رقم الهاتف (واتساب)</label>
                <input name="phone" value="{v('phone')}" type="tel" inputmode="tel" class="form-control" placeholder="07XXXXXXXXX">
            </div>
            <div class="col-md-4">
                <label class="form-label">نوع الاشتراك</label>
                <select name="plan_id" class="form-select" required>
                    {'' if s else '<option value="">اختر الاشتراك</option>'}
                    {plan_options(plans_list, s['plan_id'] if s else None)}
                </select>
            </div>
            <div class="{'col-6 col-md-2' if with_expiry else 'col-md-4'}">
                <label class="form-label">تاريخ التفعيل</label>
                <input type="date" name="activation_date" value="{v('activation_date') if s else today_str()}" class="form-control" required>
            </div>{expiry}
            <div class="col-md-4">
                <label class="form-label">ملاحظات</label>
                <input name="notes" value="{v('notes')}" class="form-control">
            </div>"""


def delete_form(subscriber_id, label=True, cls="btn-sm"):
    return f"""
        <form method="post" action="{url_for('delete_subscriber', subscriber_id=subscriber_id)}" class="d-inline"
              data-confirm="سيتم حذف المشترك وكل عملياته وديونه نهائياً. متأكد؟">
            <button class="btn {cls} btn-outline-danger"><i class="bi bi-trash3"></i>{' حذف' if label else ''}</button>
        </form>"""


@app.route("/subscribers", methods=["GET", "POST"])
@login_required
def subscribers():
    con = db()

    if request.method == "POST":
        data, error = read_subscriber_form(con)
        if error:
            con.close()
            flash(error, "danger")
            return redirect(url_for("subscribers"))

        dup = con.execute("SELECT 1 FROM subscribers WHERE username = ?", (data["username"],)).fetchone()

        con.execute("""
            INSERT INTO subscribers
            (name, username, phone, plan_id, activation_date, expiry_date, notes, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (data["name"], data["username"], data["phone"], data["plan_id"],
              data["activation_date"], data["expiry_date"], data["notes"], now_str()))
        con.commit()
        con.close()

        flash("تمت إضافة المشترك بنجاح", "success")
        if dup:
            flash("تنبيه: يوجد مشترك آخر بنفس اليوزر", "warning")
        return redirect(url_for("subscribers"))

    q = (request.args.get("q") or "").strip()
    status = request.args.get("status", "")
    plan_filter = request.args.get("plan", "")

    where, params = [], []
    if q:
        where.append("(s.name LIKE ? OR s.username LIKE ? OR s.phone LIKE ?)")
        params += [f"%{q}%"] * 3
    today = today_str()
    if status == "expired":
        where.append("s.expiry_date < ?")
        params.append(today)
    elif status == "active":
        where.append("s.expiry_date >= ?")
        params.append(today)
    elif status == "debt":
        where.append(f"({DEBT_SUBQUERY}) > 0")
    if plan_filter.isdigit():
        where.append("s.plan_id = ?")
        params.append(int(plan_filter))

    plans_list = con.execute("SELECT * FROM plans ORDER BY category, id").fetchall()
    subscriber_rows = con.execute(f"""
        SELECT s.*, p.name AS plan_name, p.retail_price, p.wholesale_price, p.profit,
               ({DEBT_SUBQUERY}) AS debt
        FROM subscribers s
        JOIN plans p ON p.id = s.plan_id
        {"WHERE " + " AND ".join(where) if where else ""}
        ORDER BY s.id DESC
    """, params).fetchall()
    shop = get_setting(con, "shop_name")
    con.close()

    subscriber_html = ""
    for s in subscriber_rows:
        debt = s["debt"] or 0
        debt_html = f'<span class="text-danger fw-bold">{money(debt)}</span>' if debt > 0 else '<span class="text-muted">—</span>'
        wa_text = msg_debt(s, debt, shop) if debt > 0 else msg_reminder(s, shop)
        subscriber_html += f"""
        <tr>
            <td><a href="{url_for('subscriber_detail', subscriber_id=s['id'])}" class="fw-bold text-decoration-none">{esc(s['name'])}</a></td>
            <td>{esc(s['username'])}</td>
            <td dir="ltr">{esc(s['phone'])}</td>
            <td>{esc(s['plan_name'])}</td>
            <td>{money(s['retail_price'])}</td>
            <td>{status_badge(s['expiry_date'])}</td>
            <td>{debt_html}</td>
            <td>
                <div class="d-flex flex-wrap gap-1">
                    <a href="{url_for('activate', subscriber_id=s['id'])}" class="btn btn-sm btn-success">تفعيل</a>
                    <a href="{url_for('transfer', subscriber_id=s['id'])}" class="btn btn-sm btn-warning">ترحيل</a>
                    {wa_button(s['phone'], wa_text, '')}
                    <a href="{url_for('subscriber_detail', subscriber_id=s['id'])}#edit" class="btn btn-sm btn-outline-primary" title="تعديل"><i class="bi bi-pencil"></i></a>
                    {delete_form(s['id'], label=False)}
                </div>
            </td>
        </tr>"""

    plan_filter_opts = "".join(
        f'<option value="{p["id"]}" {"selected" if str(p["id"]) == plan_filter else ""}>'
        f'{esc(p["category"])} - {esc(p["name"])}</option>' for p in plans_list)

    body = f"""
    <h2 class="mb-4">المشتركون</h2>

    <div class="card p-4 mb-4">
        <button class="btn btn-primary w-100 d-flex align-items-center justify-content-center gap-2" type="button"
                data-bs-toggle="collapse" data-bs-target="#addSub">
            <i class="bi bi-person-plus"></i> إضافة مشترك جديد
        </button>
        <a href="{url_for('import_users')}" class="btn btn-outline-primary w-100 mt-2 d-flex align-items-center justify-content-center gap-2">
            <i class="bi bi-file-earmark-arrow-up"></i> استيراد يوزرات من ملف الشركة (PDF)
        </a>
        <div class="collapse" id="addSub">
        <form method="post" class="row g-3 mt-2">
            {subscriber_form_fields(plans_list)}
            <div class="col-12">
                <button class="btn btn-primary">حفظ المشترك</button>
                <small class="text-muted ms-2">الإضافة لا تخصم من المحفظة — استخدم زر "تفعيل" للخصم.</small>
            </div>
        </form>
        </div>
    </div>

    <div class="card p-4">
        <div class="d-flex flex-wrap justify-content-between align-items-center mb-3 gap-2">
            <h5 class="mb-0">قائمة المشتركين ({len(subscriber_rows)})</h5>
            <form method="get" class="d-flex flex-wrap gap-2 filter-bar">
                <input name="q" value="{esc(q)}" class="form-control" placeholder="بحث بالاسم/اليوزر/الهاتف" style="max-width:240px">
                <select name="status" class="form-select" style="max-width:160px">
                    <option value="">كل الحالات</option>
                    <option value="active" {"selected" if status == "active" else ""}>فعّال</option>
                    <option value="expired" {"selected" if status == "expired" else ""}>منتهي</option>
                    <option value="debt" {"selected" if status == "debt" else ""}>عليه دين</option>
                </select>
                <select name="plan" class="form-select" style="max-width:180px">
                    <option value="">كل الباقات</option>{plan_filter_opts}
                </select>
                <button class="btn btn-outline-primary">بحث</button>
                <a href="{url_for('subscribers')}" class="btn btn-outline-secondary">مسح</a>
            </form>
        </div>

        <div class="table-responsive">
            <table class="table table-hover">
                <thead>
                    <tr>
                        <th>الاسم</th><th>اليوزر</th><th>الهاتف</th><th>الاشتراك</th>
                        <th>سعر البيع</th><th>الحالة</th><th>الدين</th><th></th>
                    </tr>
                </thead>
                <tbody>{subscriber_html or '<tr><td colspan="8" class="text-muted text-center">لا يوجد مشتركون</td></tr>'}</tbody>
            </table>
        </div>
    </div>
    """
    return page(body)


def get_subscriber(con, subscriber_id):
    return con.execute(f"""
        SELECT s.*, p.name AS plan_name, p.category,
               p.retail_price, p.wholesale_price, p.profit,
               ({DEBT_SUBQUERY}) AS debt
        FROM subscribers s
        JOIN plans p ON p.id = s.plan_id
        WHERE s.id = ?
    """, (subscriber_id,)).fetchone()


@app.route("/subscribers/<int:subscriber_id>", methods=["GET", "POST"])
@login_required
def subscriber_detail(subscriber_id):
    con = db()
    s = get_subscriber(con, subscriber_id)
    if not s:
        con.close()
        flash("المشترك غير موجود", "danger")
        return redirect(url_for("subscribers"))

    if request.method == "POST":
        data, error = read_subscriber_form(con)
        if error:
            con.close()
            flash(error, "danger")
            return redirect(url_for("subscriber_detail", subscriber_id=subscriber_id))

        con.execute("""
            UPDATE subscribers SET name = ?, username = ?, phone = ?, plan_id = ?,
                   activation_date = ?, expiry_date = ?, notes = ?
            WHERE id = ?
        """, (data["name"], data["username"], data["phone"], data["plan_id"],
              data["activation_date"], data["expiry_date"], data["notes"], subscriber_id))
        con.commit()
        con.close()
        flash("تم حفظ التعديلات", "success")
        return redirect(url_for("subscriber_detail", subscriber_id=subscriber_id))

    plans_list = con.execute("SELECT * FROM plans ORDER BY category, id").fetchall()
    history = con.execute("""
        SELECT a.*, p.name AS plan_name FROM activations a
        JOIN plans p ON p.id = a.plan_id
        WHERE a.subscriber_id = ? ORDER BY a.id DESC
    """, (subscriber_id,)).fetchall()
    ledger = con.execute("""
        SELECT * FROM debts WHERE subscriber_id = ? ORDER BY id DESC
    """, (subscriber_id,)).fetchall()
    shop = get_setting(con, "shop_name")

    # رسالة تأكيد بعد عملية تفعيل/ترحيل مباشرة
    done_html = ""
    done_id = request.args.get("done", "")
    if done_id.isdigit():
        op = con.execute("""
            SELECT a.*, p.name AS plan_name FROM activations a JOIN plans p ON p.id = a.plan_id
            WHERE a.id = ? AND a.subscriber_id = ?
        """, (int(done_id), subscriber_id)).fetchone()
        if op:
            done_html = f"""
            <div class="card p-4 mb-4 border-success" style="border-width:2px">
                <div class="d-flex flex-wrap align-items-center justify-content-between gap-3">
                    <div>
                        <h5 class="mb-1 text-success"><i class="bi bi-check-circle-fill"></i> تمت عملية {esc(op['operation_type'])} بنجاح</h5>
                        <div class="text-muted">الدفع: <b>{esc(op['payment_type'])}</b>
                        {f" — دين جديد: <b class='text-danger'>{money((op['retail_price'] or 0) - (op['paid_amount'] or 0))} د.ع</b>" if op['payment_type'] == 'دين' else ''}</div>
                    </div>
                    {wa_button(s['phone'], msg_confirm(s, op, s['debt'] or 0, shop), 'أرسل تأكيد للمشترك', 'btn-lg')}
                </div>
            </div>"""
    con.close()

    debt = s["debt"] or 0
    total_paid = sum(h["retail_price"] or 0 for h in history)
    total_profit = sum(h["profit"] or 0 for h in history)

    history_rows = "".join(f"""
        <tr>
            <td>{esc(h['operation_type'])}</td>
            <td>{esc(h['plan_name'])}</td>
            <td>{money(h['retail_price'])}</td>
            <td>{payment_badge(h['payment_type'])}</td>
            <td>{money(h['amount'])}</td>
            <td>{money(h['profit'])}</td>
            <td>{esc(h['notes'])}</td>
            <td>{esc(h['operation_date'])}</td>
        </tr>""" for h in history) or '<tr><td colspan="8" class="text-muted text-center">لا توجد عمليات لهذا المشترك</td></tr>'

    ledger_rows = "".join(f"""
        <tr>
            <td>{'<span class="badge bg-danger">دين</span>' if d['kind'] == 'دين' else '<span class="badge bg-success">تسديد</span>'}</td>
            <td class="{'text-danger' if d['kind'] == 'دين' else 'text-success'} fw-bold">{'+' if d['kind'] == 'دين' else '-'}{money(d['amount'])}</td>
            <td>{esc(d['notes'])}</td>
            <td>{esc(d['created_at'])}</td>
        </tr>""" for d in ledger) or '<tr><td colspan="4" class="text-muted text-center">لا توجد حركات ديون</td></tr>'

    pay_form = f"""
        <form method="post" action="{url_for('pay_debt', subscriber_id=subscriber_id)}" class="row g-2 mb-3">
            <div class="col-6 col-md-3">
                <label class="form-label">مبلغ التسديد</label>
                <input type="number" name="amount" min="1" max="{debt}" value="{debt}" class="form-control" required>
            </div>
            <div class="col-6 col-md-5">
                <label class="form-label">ملاحظات</label>
                <input name="notes" class="form-control" placeholder="مثال: قبض نقداً">
            </div>
            <div class="col-12 col-md-4 d-flex align-items-end gap-2">
                <button class="btn btn-success flex-fill"><i class="bi bi-cash-coin"></i> تسجيل تسديد</button>
                {wa_button(s['phone'], msg_debt(s, debt, shop), 'مطالبة')}
            </div>
        </form>""" if debt > 0 else '<p class="text-success mb-3"><i class="bi bi-check2-circle"></i> لا يوجد دين على هذا المشترك.</p>'

    body = f"""
    <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-4 page-head">
        <h2 class="mb-0">المشترك: {esc(s['name'])}</h2>
        <div>
            <a href="{url_for('activate', subscriber_id=subscriber_id)}" class="btn btn-success">تفعيل</a>
            <a href="{url_for('transfer', subscriber_id=subscriber_id)}" class="btn btn-warning">ترحيل</a>
            <a href="{url_for('subscribers')}" class="btn btn-outline-secondary">رجوع</a>
        </div>
    </div>

    {done_html}

    <div class="row g-3 mb-4">
        <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>الحالة</h6><div>{status_badge(s['expiry_date'])}</div>
            <small class="text-muted">ينتهي: {esc(s['expiry_date'])}</small></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>الدين الحالي</h6>
            <h4 class="fw-bold {'text-danger' if debt > 0 else 'text-success'}">{money(debt)} د.ع</h4></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>مجموع المبيعات له</h6><h4 class="fw-bold">{money(total_paid)} د.ع</h4></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>مجموع الأرباح منه</h6><h4 class="text-success fw-bold">{money(total_profit)} د.ع</h4></div></div>
    </div>

    <div class="card p-4 mb-4">
        <h5><i class="bi bi-whatsapp" style="color:#25D366"></i> مراسلة المشترك على واتساب</h5>
        {"<p class='text-muted mb-0'>أضف رقم هاتف للمشترك حتى تكدر تراسله.</p>" if not wa_phone(s['phone']) else f'''
        <div class="d-flex flex-wrap gap-2 mb-3">
            <button type="button" class="btn btn-outline-secondary btn-sm" onclick="setMsg(this)" data-msg="{esc(msg_reminder(s, shop))}">تذكير بانتهاء الاشتراك</button>
            <button type="button" class="btn btn-outline-secondary btn-sm" onclick="setMsg(this)" data-msg="{esc(msg_debt(s, debt, shop))}">مطالبة بالدين</button>
            <button type="button" class="btn btn-outline-secondary btn-sm" onclick="setMsg(this)" data-msg="{esc(msg_activated(s, shop))}">تأكيد التفعيل</button>
        </div>
        <textarea id="waText" class="form-control mb-3" rows="4" placeholder="اكتب رسالتك هنا...">{esc(msg_reminder(s, shop))}</textarea>
        <div><button type="button" class="btn btn-wa" onclick="sendWa()"><i class="bi bi-whatsapp"></i> إرسال على واتساب</button></div>
        <script>
            function setMsg(b) {{ document.getElementById('waText').value = b.dataset.msg; }}
            function sendWa() {{
                const t = document.getElementById('waText').value.trim();
                if (!t) return;
                window.open('https://wa.me/{wa_phone(s['phone'])}?text=' + encodeURIComponent(t), '_blank');
            }}
        </script>'''}
    </div>

    <div class="card p-4 mb-4" id="debts">
        <h5><i class="bi bi-journal-text text-danger"></i> حساب الديون — الرصيد: <span class="{'text-danger' if debt > 0 else 'text-success'}">{money(debt)} د.ع</span></h5>
        {pay_form}
        <div class="table-responsive">
            <table class="table">
                <thead><tr><th>النوع</th><th>المبلغ</th><th>الملاحظات</th><th>التاريخ</th></tr></thead>
                <tbody>{ledger_rows}</tbody>
            </table>
        </div>
    </div>

    <div class="card p-4 mb-4" id="edit">
        <h5><i class="bi bi-pencil-square text-primary"></i> تعديل البيانات</h5>
        <form method="post" class="row g-3 mt-1">
            {subscriber_form_fields(plans_list, s, with_expiry=True)}
            <div class="col-12 d-flex flex-wrap gap-2 justify-content-between">
                <button class="btn btn-primary"><i class="bi bi-check2"></i> حفظ التعديلات</button>
            </div>
        </form>
        <hr>
        <div class="d-flex flex-wrap align-items-center justify-content-between gap-2">
            <small class="text-muted">حذف المشترك يحذف كل عملياته وديونه من التقارير نهائياً (حركات المحفظة تبقى).</small>
            {delete_form(subscriber_id, cls='')}
        </div>
    </div>

    <div class="card p-4">
        <h5><i class="bi bi-clock-history text-primary"></i> سجل العمليات</h5>
        <div class="table-responsive">
            <table class="table table-striped">
                <thead><tr><th>العملية</th><th>الاشتراك</th><th>سعر البيع</th><th>الدفع</th><th>سعر الجملة</th>
                <th>الربح</th><th>الملاحظات</th><th>التاريخ</th></tr></thead>
                <tbody>{history_rows}</tbody>
            </table>
        </div>
    </div>
    """
    return page(body)


@app.route("/subscribers/<int:subscriber_id>/delete", methods=["POST"])
@login_required
def delete_subscriber(subscriber_id):
    con = db()
    s = con.execute("SELECT username FROM subscribers WHERE id = ?", (subscriber_id,)).fetchone()
    if not s:
        con.close()
        flash("المشترك غير موجود", "danger")
        return redirect(url_for("subscribers"))
    con.execute("DELETE FROM debts WHERE subscriber_id = ?", (subscriber_id,))
    con.execute("DELETE FROM activations WHERE subscriber_id = ?", (subscriber_id,))
    con.execute("DELETE FROM subscribers WHERE id = ?", (subscriber_id,))
    con.commit()
    con.close()
    flash(f"تم حذف المشترك {s['username']}", "success")
    return redirect(url_for("subscribers"))


# ---------------------------------------------------------------------------
# الديون
# ---------------------------------------------------------------------------

@app.route("/subscribers/<int:subscriber_id>/pay", methods=["POST"])
@login_required
def pay_debt(subscriber_id):
    back = request.form.get("back") or url_for("subscriber_detail", subscriber_id=subscriber_id) + "#debts"
    if not back.startswith("/"):
        back = url_for("debts")
    amount = parse_amount(request.form.get("amount"))
    con = db()
    s = get_subscriber(con, subscriber_id)
    if not s:
        con.close()
        flash("المشترك غير موجود", "danger")
        return redirect(url_for("debts"))
    debt = s["debt"] or 0
    if not amount or amount <= 0:
        flash("أدخل مبلغاً صحيحاً", "danger")
    elif amount > debt:
        flash(f"المبلغ أكبر من الدين ({money(debt)} د.ع)", "danger")
    else:
        notes = form_text("notes") or "تسديد"
        con.execute("""
            INSERT INTO debts (subscriber_id, kind, amount, notes, created_at)
            VALUES (?, 'تسديد', ?, ?, ?)
        """, (subscriber_id, amount, notes, now_str()))
        con.commit()
        left = debt - amount
        flash(f"تم تسجيل تسديد {money(amount)} د.ع من {s['name']} — المتبقي: {money(left)} د.ع", "success")
    con.close()
    return redirect(back)


@app.route("/debts")
@login_required
def debts():
    con = db()
    rows = con.execute(f"""
        SELECT s.*, p.name AS plan_name, ({DEBT_SUBQUERY}) AS debt,
               (SELECT MAX(created_at) FROM debts d WHERE d.subscriber_id = s.id AND d.kind = 'دين') AS last_debt
        FROM subscribers s JOIN plans p ON p.id = s.plan_id
        WHERE ({DEBT_SUBQUERY}) > 0
        ORDER BY debt DESC
    """).fetchall()
    month_start = date.today().replace(day=1).strftime("%Y-%m-%d")
    collected = con.execute("""
        SELECT COALESCE(SUM(amount), 0) AS c FROM debts WHERE kind = 'تسديد' AND created_at >= ?
    """, (month_start,)).fetchone()["c"]
    recent = con.execute("""
        SELECT d.*, s.name, s.username FROM debts d JOIN subscribers s ON s.id = d.subscriber_id
        ORDER BY d.id DESC LIMIT 30
    """).fetchall()
    shop = get_setting(con, "shop_name")
    con.close()

    total = sum(r["debt"] for r in rows)
    back = url_for("debts")

    debt_rows = "".join(f"""
        <tr>
            <td><a href="{url_for('subscriber_detail', subscriber_id=r['id'])}#debts" class="fw-bold text-decoration-none">{esc(r['name'])}</a></td>
            <td>{esc(r['username'])}</td>
            <td dir="ltr">{esc(r['phone'])}</td>
            <td class="text-danger fw-bold">{money(r['debt'])} د.ع</td>
            <td>{esc((r['last_debt'] or '')[:10])}</td>
            <td>
                <form method="post" action="{url_for('pay_debt', subscriber_id=r['id'])}" class="d-flex gap-1 flex-wrap">
                    <input type="hidden" name="back" value="{back}">
                    <input type="number" name="amount" min="1" max="{r['debt']}" value="{r['debt']}" class="form-control form-control-sm" style="max-width:120px" required>
                    <button class="btn btn-sm btn-success"><i class="bi bi-cash-coin"></i> تسديد</button>
                    {wa_button(r['phone'], msg_debt(r, r['debt'], shop), '')}
                </form>
            </td>
        </tr>""" for r in rows)

    recent_rows = "".join(f"""
        <tr>
            <td>{'<span class="badge bg-danger">دين</span>' if d['kind'] == 'دين' else '<span class="badge bg-success">تسديد</span>'}</td>
            <td>{esc(d['name'])}</td>
            <td class="{'text-danger' if d['kind'] == 'دين' else 'text-success'} fw-bold">{money(d['amount'])}</td>
            <td>{esc(d['notes'])}</td>
            <td>{esc(d['created_at'])}</td>
        </tr>""" for d in recent)

    body = f"""
    <h2 class="mb-4">حساب الديون</h2>

    <div class="row g-3 mb-4">
        <div class="col-6 col-md-4">
            <div class="card stat-card p-4" style="background:linear-gradient(135deg,#dc2626,#f43f5e);color:#fff;border:0">
                <span class="stat-icon" style="background:rgba(255,255,255,.18);color:#fff"><i class="bi bi-journal-text"></i></span>
                <h6 style="color:rgba(255,255,255,.85)">مجموع الديون</h6>
                <h2>{money(total)} د.ع</h2>
            </div>
        </div>
        <div class="col-6 col-md-4">
            <div class="card stat-card p-4">
                <span class="stat-icon"><i class="bi bi-people"></i></span>
                <h6>عدد المدينين</h6>
                <h2>{len(rows)}</h2>
            </div>
        </div>
        <div class="col-12 col-md-4">
            <div class="card stat-card p-4">
                <span class="stat-icon" style="background:rgba(16,185,129,.14);color:#10b981"><i class="bi bi-cash-stack"></i></span>
                <h6>المقبوض من الديون هذا الشهر</h6>
                <h2 class="text-success">{money(collected)} د.ع</h2>
            </div>
        </div>
    </div>

    <div class="card p-4 mb-4">
        <h5>المشتركين الي عليهم دين</h5>
        <div class="table-responsive">
            <table class="table table-hover">
                <thead><tr><th>الاسم</th><th>اليوزر</th><th>الهاتف</th><th>الدين</th><th>آخر دين</th><th>تسديد / مطالبة</th></tr></thead>
                <tbody>{debt_rows or '<tr><td colspan="6" class="text-muted text-center">ماكو ديون 👍</td></tr>'}</tbody>
            </table>
        </div>
    </div>

    <div class="card p-4">
        <h5>آخر حركات الديون</h5>
        <div class="table-responsive">
            <table class="table table-striped">
                <thead><tr><th>النوع</th><th>المشترك</th><th>المبلغ</th><th>الملاحظات</th><th>التاريخ</th></tr></thead>
                <tbody>{recent_rows or '<tr><td colspan="5" class="text-muted text-center">لا توجد حركات</td></tr>'}</tbody>
            </table>
        </div>
    </div>
    """
    return page(body)


# ---------------------------------------------------------------------------
# التفعيل والترحيل
# ---------------------------------------------------------------------------

def perform_operation(subscriber_id, operation_type):
    con = db()
    # قفل قاعدة البيانات أثناء العملية حتى لا يُخصم الرصيد مرتين بنفس الوقت
    con.execute("BEGIN IMMEDIATE")

    subscriber = get_subscriber(con, subscriber_id)
    if not subscriber:
        con.rollback()
        con.close()
        flash("المشترك غير موجود", "danger")
        return redirect(url_for("subscribers"))

    # يمكن تغيير الباقة أثناء التفعيل/الترحيل
    try:
        plan_id = int(request.form.get("plan_id") or subscriber["plan_id"])
    except ValueError:
        plan_id = subscriber["plan_id"]
    plan = con.execute("SELECT * FROM plans WHERE id = ?", (plan_id,)).fetchone()
    if not plan:
        con.rollback()
        con.close()
        flash("الباقة غير موجودة", "danger")
        return redirect(url_for("subscribers"))

    amount = plan["wholesale_price"]
    balance = wallet_balance(con)
    back = url_for("transfer" if operation_type == "ترحيل" else "activate", subscriber_id=subscriber_id)

    if balance < amount:
        con.rollback()
        con.close()
        flash(f"الرصيد غير كافٍ. المطلوب {money(amount)} د.ع، والرصيد الحالي {money(balance)} د.ع", "danger")
        return redirect(back)

    # طريقة الدفع: نقد (مقبوض) أو دين (يتحول لحساب الديون)
    payment_type = request.form.get("payment_type", "نقد")
    if payment_type not in ("نقد", "دين"):
        payment_type = "نقد"
    retail = plan["retail_price"]
    if payment_type == "نقد":
        paid = retail
    else:
        paid = parse_amount(request.form.get("paid_amount")) or 0
        if paid < 0 or paid >= retail:
            con.rollback()
            con.close()
            flash("المبلغ المدفوع لازم يكون أقل من سعر البيع (إذا دفع كامل اختار نقد)", "danger")
            return redirect(back)

    notes = form_text("notes")
    now = now_str()
    start = today_str()

    # الترحيل: إذا الاشتراك بعده فعّال، تبدأ المدة الجديدة من تاريخ انتهاء الحالي
    if operation_type == "ترحيل":
        current_exp = parse_date(subscriber["expiry_date"])
        if current_exp and current_exp > start:
            start = current_exp
    new_expiry = expiry_from(start)

    con.execute("""
        INSERT INTO wallet_transactions (transaction_type, amount, notes, created_at)
        VALUES (?, ?, ?, ?)
    """, ("خصم", amount, f"{operation_type} - {subscriber['username']} - {plan['name']}"
                         + (f" - {notes}" if notes else ""), now))

    cur = con.execute("""
        INSERT INTO activations
        (subscriber_id, plan_id, operation_type, amount, retail_price, profit,
         payment_type, paid_amount, notes, operation_date)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (subscriber_id, plan_id, operation_type, amount, retail, plan["profit"],
          payment_type, paid, notes, now))
    activation_id = cur.lastrowid

    if payment_type == "دين":
        con.execute("""
            INSERT INTO debts (subscriber_id, kind, amount, notes, activation_id, created_at)
            VALUES (?, 'دين', ?, ?, ?, ?)
        """, (subscriber_id, retail - paid,
              f"{operation_type} {plan['name']}" + (f" (دفع {money(paid)})" if paid else ""),
              activation_id, now))

    con.execute("""
        UPDATE subscribers
        SET activation_date = ?, expiry_date = ?, plan_id = ?,
            notes = CASE WHEN ? <> '' THEN ? ELSE notes END
        WHERE id = ?
    """, (today_str(), new_expiry, plan_id, notes, notes, subscriber_id))

    con.commit()
    con.close()

    msg = (f"تمت عملية {operation_type} للمشترك {subscriber['username']} وخصم {money(amount)} د.ع "
           f"من المحفظة — ينتهي الاشتراك بتاريخ {new_expiry}")
    if payment_type == "دين":
        msg += f" — أُضيف {money(retail - paid)} د.ع لحساب الديون"
    flash(msg, "success")
    return redirect(url_for("subscriber_detail", subscriber_id=subscriber_id, done=activation_id))


def operation_page(subscriber_id, operation_type):
    con = db()
    subscriber = get_subscriber(con, subscriber_id)
    plans_list = con.execute("SELECT * FROM plans ORDER BY category, id").fetchall()
    balance = wallet_balance(con)
    con.close()

    if not subscriber:
        flash("المشترك غير موجود", "danger")
        return redirect(url_for("subscribers"))

    low = balance < subscriber["wholesale_price"]
    debt = subscriber["debt"] or 0

    body = f"""
    <h2 class="mb-4">{operation_type} الاشتراك</h2>
    <div class="row g-4">
        <div class="col-lg-4 order-lg-2">
            <div class="card p-4">
                <h5><i class="bi bi-person-circle text-primary"></i> {esc(subscriber['name'])}</h5>
                <div class="d-flex justify-content-between py-1"><span class="text-muted">اليوزر</span><b>{esc(subscriber['username'])}</b></div>
                <div class="d-flex justify-content-between py-1"><span class="text-muted">الحالة</span>{status_badge(subscriber['expiry_date'])}</div>
                <div class="d-flex justify-content-between py-1"><span class="text-muted">ينتهي</span><span>{esc(subscriber['expiry_date'])}</span></div>
                <div class="d-flex justify-content-between py-1"><span class="text-muted">دين سابق</span>
                    <b class="{'text-danger' if debt > 0 else 'text-success'}">{money(debt)} د.ع</b></div>
                <div class="d-flex justify-content-between py-1"><span class="text-muted">رصيد المحفظة</span>
                    <b class="{'text-danger' if low else 'text-success'}">{money(balance)} د.ع</b></div>
            </div>
        </div>

        <div class="col-lg-8 order-lg-1">
            <div class="card p-4">
                <form method="post">
                    <label class="form-label">الباقة</label>
                    <select name="plan_id" id="planSel" class="form-select mb-2" onchange="updPay()">{plan_options(plans_list, subscriber['plan_id'])}</select>
                    <small class="text-muted d-block mb-4">المبلغ المخصوم من المحفظة = سعر الجملة للباقة المختارة.
                    {"عند الترحيل تُضاف المدة على تاريخ الانتهاء الحالي إذا كان الاشتراك بعده فعّال." if operation_type == "ترحيل" else ""}</small>

                    <label class="form-label">طريقة الدفع</label>
                    <div class="pay-options mb-3">
                        <input type="radio" class="btn-check" name="payment_type" id="payCash" value="نقد" checked onchange="updPay()">
                        <label class="pay-opt" for="payCash">
                            <i class="bi bi-cash-coin"></i>
                            <b>نقداً</b><small>قبضت المبلغ كامل</small>
                        </label>
                        <input type="radio" class="btn-check" name="payment_type" id="payDebt" value="دين" onchange="updPay()">
                        <label class="pay-opt debt" for="payDebt">
                            <i class="bi bi-journal-text"></i>
                            <b>دين</b><small>يتحول لحساب الديون</small>
                        </label>
                    </div>

                    <div id="debtBox" class="mb-3" style="display:none">
                        <label class="form-label">المبلغ المقبوض الآن (اتركه 0 إذا ما دفع شي)</label>
                        <input type="number" name="paid_amount" id="paidAmt" min="0" value="0" class="form-control" oninput="updPay()">
                    </div>

                    <div class="alert mb-3" id="paySummary" style="background:var(--soft)"></div>

                    <label class="form-label">ملاحظات {operation_type}</label>
                    <textarea name="notes" class="form-control mb-3" rows="2" placeholder="اكتب الملاحظات هنا"></textarea>

                    <div class="d-flex gap-2">
                        <button class="btn btn-success flex-fill py-3"><i class="bi bi-check2-circle"></i> تأكيد {operation_type}</button>
                        <a href="{url_for('subscribers')}" class="btn btn-outline-secondary py-3">رجوع</a>
                    </div>
                </form>
            </div>
        </div>
    </div>
    <script>
        function fmt(n) {{ return Number(n).toLocaleString('en-US'); }}
        function updPay() {{
            const retail = +document.getElementById('planSel').selectedOptions[0].dataset.retail;
            const isDebt = document.getElementById('payDebt').checked;
            document.getElementById('debtBox').style.display = isDebt ? '' : 'none';
            const paid = isDebt ? Math.max(0, +document.getElementById('paidAmt').value || 0) : retail;
            const s = document.getElementById('paySummary');
            if (!isDebt) {{
                s.innerHTML = 'سعر البيع: <b>' + fmt(retail) + ' د.ع</b> — مقبوض نقداً بالكامل ✅';
            }} else if (paid >= retail) {{
                s.innerHTML = '<span class="text-danger">المبلغ المقبوض يساوي أو أكثر من السعر — اختار نقداً</span>';
            }} else {{
                s.innerHTML = 'سعر البيع: <b>' + fmt(retail) + '</b> — المقبوض: <b>' + fmt(paid) +
                    '</b> — يُضاف للديون: <b class="text-danger">' + fmt(retail - paid) + ' د.ع</b>';
            }}
        }}
        updPay();
    </script>
    """
    return page(body)


@app.route("/activate/<int:subscriber_id>", methods=["GET", "POST"])
@login_required
def activate(subscriber_id):
    if request.method == "POST":
        return perform_operation(subscriber_id, "تفعيل")
    return operation_page(subscriber_id, "تفعيل")


@app.route("/transfer/<int:subscriber_id>", methods=["GET", "POST"])
@login_required
def transfer(subscriber_id):
    if request.method == "POST":
        return perform_operation(subscriber_id, "ترحيل")
    return operation_page(subscriber_id, "ترحيل")


# ---------------------------------------------------------------------------
# استيراد اليوزرات من ملف PDF الشركة
# ---------------------------------------------------------------------------

IMPORT_DIR = os.path.join(os.path.dirname(DB_NAME) or BASE_DIR, "imports_tmp")

# عناوين أعمدة تقرير الشركة (كما تظهر بعد قراءة النص)
PDF_HEADERS = {
    "ت": "seq",
    "نوع الكارت": "card_type",
    "اسم الحساب": "username",
    "اسم الخساب": "username",
    "اسم العميل": "name",
    "رقم الكارت": "card_no",
    "المبلغ": "amount",
    "مدخل البيانات": "entered_by",
    "التاريخ": "date",
}
# مواقع الأعمدة الافتراضية إذا الصفحة ما بيها عناوين
PDF_DEFAULT_COLS = {"seq": 819, "card_type": 733, "username": 614, "name": 481,
                    "card_no": 344, "amount": 253, "entered_by": 170, "date": 77}

_ARABIC = re.compile(r"[؀-ۿ]")
_LTR_RUN = re.compile(r"[A-Za-z0-9@._:/,\-]+")


def _pdf_fonts(page):
    """خريطة كل رمز بالخط إلى الحرف الحقيقي، مع إصلاح حرف (لا/لأ) المدموج."""
    fonts = {}
    resources = page.get("/Resources") or {}
    font_dict = resources.get("/Font") or {}
    for key, ref in font_dict.items():
        font = ref.get_object()
        mapping = {}
        tu = font.get("/ToUnicode")
        if tu is not None:
            cmap = tu.get_object().get_data().decode("latin1")
            for a, b, u in re.findall(r"<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]*)>", cmap):
                for code in range(int(a, 16), int(b, 16) + 1):
                    try:
                        mapping[code] = chr(int(u, 16)) if u else ""
                    except ValueError:
                        mapping[code] = ""
        # الحروف المدموجة (لا، لأ) تنقرأ "ل" فقط، نميزها من عرض الحرف
        widths = font.get("/Widths")
        first = font.get("/FirstChar", 0)
        if widths is not None:
            widths = [float(w) for w in widths]
            for c, ch in mapping.items():
                if ch == "ل" and 0 <= c - first < len(widths):
                    w = widths[c - first]
                    if w >= 520:
                        # النص بالترتيب البصري (معكوس)، فنكتب الألف قبل اللام
                        mapping[c] = "أل" if 540 <= w <= 548 else "ال"
        fonts[key] = mapping
    return fonts


def _visual_to_logical(text):
    """النص العربي بملف الشركة مخزون بالمقلوب: نعدله ونرجع الأرقام والإنكليزي لاتجاهها."""
    if not _ARABIC.search(text):
        return " ".join(text.split())
    rev = text[::-1]
    rev = _LTR_RUN.sub(lambda m: m.group(0)[::-1], rev)
    return " ".join(rev.split())


def _pdf_items(page, reader):
    from pypdf.generic import ContentStream
    fonts = _pdf_fonts(page)
    stream = ContentStream(page.get_contents(), reader)
    font = None
    x = y = lx = ly = 0.0
    items = []
    for operands, op in stream.operations:
        if op == b"Tf":
            font = operands[0]
        elif op == b"Tm":
            x = lx = float(operands[4])
            y = ly = float(operands[5])
        elif op in (b"Td", b"TD"):
            lx += float(operands[0])
            ly += float(operands[1])
            x, y = lx, ly
        elif op in (b"Tj", b"TJ"):
            parts = operands[0] if op == b"TJ" else [operands[0]]
            raw = b"".join(p.get_original_bytes() for p in parts if hasattr(p, "get_original_bytes"))
            mapping = fonts.get(font, {})
            text = "".join(mapping.get(c, "") for c in raw)
            if text.strip():
                items.append((y, x, text))
    return items


def parse_company_pdf(data):
    """قراءة تقرير الشركة. ترجع قائمة صفوف: يوزر، اسم، نوع الكارت، المبلغ، التاريخ."""
    import pypdf
    reader = pypdf.PdfReader(BytesIO(data))
    rows = []
    cols = dict(PDF_DEFAULT_COLS)
    for page in reader.pages:
        items = _pdf_items(page, reader)
        # تجميع العناصر حسب السطر
        lines = []
        for y, x, text in sorted(items, key=lambda t: -t[0]):
            if lines and abs(lines[-1][0] - y) <= 1.5:
                lines[-1][1].append((x, text))
            else:
                lines.append([y, [(x, text)]])

        # مواقع العناوين بهذي الصفحة
        found = {}
        for _, parts in lines:
            for x, text in parts:
                key = PDF_HEADERS.get(_visual_to_logical(text))
                if key:
                    found[key] = x
        if len(found) >= 5:
            cols.update(found)

        for _, parts in lines:
            fields = {}
            for x, text in sorted(parts, key=lambda p: p[0]):
                key = min(cols, key=lambda k: abs(cols[k] - x))
                fields.setdefault(key, []).append(text)
            date_txt = re.sub(r"/\s+", "/", " ".join(" ".join(fields.get("date", [])).split()))
            m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", date_txt)
            if not m:
                continue
            day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
            try:
                op_date = date(year, month, day).strftime("%Y-%m-%d")
            except ValueError:
                continue
            username = "".join(" ".join(fields.get("username", [])).split())
            if not username:
                continue
            name_parts = [_visual_to_logical(t) for t in fields.get("name", []) + fields.get("card_no", [])
                          if _ARABIC.search(t)]
            amount = parse_amount(" ".join(fields.get("amount", [])).replace(" ", "")) or 0
            rows.append({
                "username": username,
                "name": " ".join(name_parts).strip(),
                "card_type": " ".join(_visual_to_logical(t) for t in fields.get("card_type", [])).strip(),
                "amount": amount,
                "entered_by": " ".join(_visual_to_logical(t) for t in fields.get("entered_by", [])).strip(),
                "date": op_date,
                "time": date_txt[m.end():].strip(),
            })
    return rows


PLAN_ALIASES = {
    "economy": ["إيكو", "ايكو", "economy"],
    "plus": ["بلص", "بلس", "plus"],
    "standard": ["استاندر", "ستاندر", "standard"],
    "active": ["اكتف", "active"],
    "basic": ["بيسك", "basic"],
}


def guess_plan(card_type, plans_list):
    """تخمين الباقة المطابقة لنوع الكارت."""
    ct = card_type.lower().strip()
    for p in plans_list:
        if p["name"].lower().strip() == ct:
            return p["id"]
    num = re.search(r"\d+", ct)
    if "fiber" in ct and num:
        for p in plans_list:
            if "فايبر" in p["name"] and re.search(rf"\b{num.group(0)}\b", p["name"]):
                return p["id"]
    if not ct.endswith("-w"):
        for key, names in PLAN_ALIASES.items():
            if ct == key:
                for p in plans_list:
                    if p["name"].strip().lower() in names:
                        return p["id"]
    return "new"


def _import_path(token):
    if not re.fullmatch(r"[0-9a-f]{32}", token or ""):
        return None
    return os.path.join(IMPORT_DIR, token + ".json")


def _cleanup_imports():
    try:
        cutoff = datetime.now().timestamp() - 86400
        for f in os.listdir(IMPORT_DIR):
            path = os.path.join(IMPORT_DIR, f)
            if os.path.getmtime(path) < cutoff:
                os.remove(path)
    except OSError:
        pass


@app.route("/import", methods=["GET", "POST"])
@login_required
def import_users():
    import json

    if request.method == "POST" and request.form.get("step") == "upload":
        file = request.files.get("file")
        if not file or not file.filename:
            flash("اختار ملف PDF", "danger")
            return redirect(url_for("import_users"))
        data = file.read()
        if not data.startswith(b"%PDF"):
            flash("الملف مو PDF. نزّل التقرير من لوحة الشركة بصيغة PDF", "danger")
            return redirect(url_for("import_users"))
        try:
            rows = parse_company_pdf(data)
        except ImportError:
            flash("مكتبة قراءة PDF غير منصّبة على السيرفر. نفّذ الأمر: pip3 install pypdf", "danger")
            return redirect(url_for("import_users"))
        except Exception:
            app.logger.exception("PDF import failed")
            flash("ما گدرت أقرأ الملف. تأكد إنه تقرير اليوزرات من لوحة الشركة", "danger")
            return redirect(url_for("import_users"))
        if not rows:
            flash("ما لگيت يوزرات بالملف. تأكد إنه تقرير اليوزرات من لوحة الشركة", "warning")
            return redirect(url_for("import_users"))

        # إذا اليوزر مكرر بالملف (جدد أكثر من مرة) نأخذ آخر عملية
        latest = {}
        for r in rows:
            key = r["username"].lower()
            if key not in latest or (r["date"], r["time"]) >= (latest[key]["date"], latest[key]["time"]):
                if key in latest and not r["name"]:
                    r["name"] = latest[key]["name"]
                latest[key] = r
        rows = sorted(latest.values(), key=lambda r: (r["date"], r["time"]))

        os.makedirs(IMPORT_DIR, exist_ok=True)
        _cleanup_imports()
        token = secrets.token_hex(16)
        with open(_import_path(token), "w", encoding="utf-8") as f:
            json.dump({"rows": rows, "file": file.filename, "total": len(rows)}, f, ensure_ascii=False)
        return redirect(url_for("import_users", t=token))

    if request.method == "POST" and request.form.get("step") == "confirm":
        path = _import_path(request.form.get("token"))
        if not path or not os.path.exists(path):
            flash("انتهت صلاحية الاستيراد، ارفع الملف مرة ثانية", "danger")
            return redirect(url_for("import_users"))
        with open(path, encoding="utf-8") as f:
            rows = json.load(f)["rows"]

        update_existing = request.form.get("update_existing") == "1"
        profit = parse_amount(request.form.get("profit"))
        profit = PROFIT if profit is None or profit < 0 else profit

        con = db()
        con.execute("BEGIN IMMEDIATE")
        plans_by_type = {}
        created_plans = 0
        types = sorted({r["card_type"] for r in rows})
        for i, ct in enumerate(types):
            choice = request.form.get(f"plan_{i}", "new")
            if choice.isdigit() and con.execute("SELECT 1 FROM plans WHERE id = ?", (int(choice),)).fetchone():
                plans_by_type[ct] = int(choice)
            else:
                same = [r for r in rows if r["card_type"] == ct]
                wholesale = max(same, key=lambda r: (r["date"], r["time"]))["amount"]
                category = "فايبر" if "fiber" in ct.lower() else "وايرليس"
                cur = con.execute("""
                    INSERT INTO plans (category, name, retail_price, profit, wholesale_price)
                    VALUES (?, ?, ?, ?, ?)
                """, (category, ct or "غير معروف", wholesale + profit, profit, wholesale))
                plans_by_type[ct] = cur.lastrowid
                created_plans += 1

        added = updated = skipped = 0
        now = now_str()
        for i, r in enumerate(rows):
            if request.form.get(f"use_{i}") != "1":
                continue
            name = (request.form.get(f"name_{i}") or "").strip() or r["name"] or r["username"]
            plan_id = plans_by_type[r["card_type"]]
            start = r["date"]
            expiry = expiry_from(start)
            existing = con.execute(
                "SELECT * FROM subscribers WHERE lower(username) = lower(?)", (r["username"],)
            ).fetchone()
            if existing:
                if not update_existing or (existing["activation_date"] or "") > start:
                    skipped += 1
                    continue
                con.execute("""
                    UPDATE subscribers SET plan_id = ?, activation_date = ?, expiry_date = ?,
                           name = CASE WHEN name = '' OR name = username THEN ? ELSE name END
                    WHERE id = ?
                """, (plan_id, start, expiry, name, existing["id"]))
                updated += 1
            else:
                con.execute("""
                    INSERT INTO subscribers
                    (name, username, phone, plan_id, activation_date, expiry_date, notes, created_at)
                    VALUES (?, ?, '', ?, ?, ?, ?, ?)
                """, (name, r["username"], plan_id, start, expiry, "", now))
                added += 1
        con.commit()
        con.close()
        try:
            os.remove(path)
        except OSError:
            pass

        msg = f"تم الاستيراد: {added} مشترك جديد"
        if updated:
            msg += f"، تحديث {updated}"
        if skipped:
            msg += f"، تخطي {skipped}"
        if created_plans:
            msg += f" — وأُضيفت {created_plans} باقة جديدة (راجع أسعارها من صفحة الباقات)"
        flash(msg, "success")
        return redirect(url_for("subscribers"))

    # --- عرض المعاينة ---
    token = request.args.get("t")
    path = _import_path(token) if token else None
    if path and os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            payload = json.load(f)
        rows = payload["rows"]
        con = db()
        plans_list = con.execute("SELECT * FROM plans ORDER BY category, id").fetchall()
        existing = {r["u"]: r["activation_date"] for r in con.execute(
            "SELECT lower(username) AS u, activation_date FROM subscribers")}
        con.close()

        types = sorted({r["card_type"] for r in rows})
        type_rows = ""
        for i, ct in enumerate(types):
            same = [r for r in rows if r["card_type"] == ct]
            amounts = sorted({r["amount"] for r in same})
            guess = guess_plan(ct, plans_list)
            opts = f'<option value="new" {"selected" if guess == "new" else ""}>➕ باقة جديدة باسم "{esc(ct)}"</option>'
            opts += "".join(
                f'<option value="{p["id"]}" {"selected" if guess == p["id"] else ""}>'
                f'{esc(p["category"])} - {esc(p["name"])} (جملة {money(p["wholesale_price"])})</option>'
                for p in plans_list)
            type_rows += f"""
            <tr>
                <td><b dir="ltr">{esc(ct) or '—'}</b></td>
                <td>{len(same)}</td>
                <td>{" / ".join(money(a) for a in amounts)}</td>
                <td><select name="plan_{i}" class="form-select form-select-sm">{opts}</select></td>
            </tr>"""

        new_count = sum(1 for r in rows if r["username"].lower() not in existing)
        exist_count = len(rows) - new_count
        user_rows = ""
        for i, r in enumerate(rows):
            is_existing = r["username"].lower() in existing
            badge = ('<span class="badge bg-warning text-dark">موجود</span>' if is_existing
                     else '<span class="badge bg-success">جديد</span>')
            user_rows += f"""
            <tr>
                <td><input type="checkbox" name="use_{i}" value="1" class="form-check-input row-check" checked></td>
                <td><input name="name_{i}" value="{esc(r['name'])}" class="form-control form-control-sm" placeholder="{esc(r['username'])}"></td>
                <td dir="ltr">{esc(r['username'])}</td>
                <td dir="ltr">{esc(r['card_type'])}</td>
                <td>{money(r['amount'])}</td>
                <td>{esc(r['date'])}</td>
                <td>{badge}</td>
            </tr>"""

        body = f"""
        <h2 class="mb-4">استيراد اليوزرات</h2>
        <form method="post">
            <input type="hidden" name="step" value="confirm">
            <input type="hidden" name="token" value="{esc(token)}">

            <div class="row g-3 mb-4">
                <div class="col-4"><div class="card stat-card p-3"><h6>يوزرات بالملف</h6><h3>{len(rows)}</h3></div></div>
                <div class="col-4"><div class="card stat-card p-3"><h6>جديد</h6><h3 class="text-success">{new_count}</h3></div></div>
                <div class="col-4"><div class="card stat-card p-3"><h6>موجود سابقاً</h6><h3 class="text-warning">{exist_count}</h3></div></div>
            </div>

            <div class="card p-4 mb-4">
                <h5><i class="bi bi-box-seam text-primary"></i> ربط نوع الكارت بالباقة</h5>
                <p class="text-muted small">لكل نوع كارت اختار الباقة المقابلة ببرنامجك، أو خليه "باقة جديدة" وتنضاف تلقائياً:
                سعر الجملة = آخر مبلغ بالملف، وسعر البيع = الجملة + الربح.</p>
                <div class="table-responsive">
                    <table class="table">
                        <thead><tr><th>نوع الكارت</th><th>العدد</th><th>المبلغ بالملف</th><th>الباقة ببرنامجك</th></tr></thead>
                        <tbody>{type_rows}</tbody>
                    </table>
                </div>
                <div class="row g-3 mt-1">
                    <div class="col-md-4">
                        <label class="form-label">الربح للباقات الجديدة</label>
                        <input type="number" name="profit" min="0" value="{PROFIT}" class="form-control">
                    </div>
                    <div class="col-md-8 d-flex align-items-end">
                        <div class="form-check">
                            <input class="form-check-input" type="checkbox" name="update_existing" value="1" id="upd" checked>
                            <label class="form-check-label" for="upd">
                                اليوزرات الموجودة: حدّث الباقة وتاريخ التفعيل والانتهاء (إذا تاريخ الملف أحدث)
                            </label>
                        </div>
                    </div>
                </div>
            </div>

            <div class="card p-4 mb-4">
                <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-2">
                    <h5 class="mb-0"><i class="bi bi-people text-primary"></i> اليوزرات</h5>
                    <div class="form-check">
                        <input class="form-check-input" type="checkbox" id="allChk" checked
                               onchange="document.querySelectorAll('.row-check').forEach(c => c.checked = this.checked)">
                        <label class="form-check-label" for="allChk">تحديد الكل</label>
                    </div>
                </div>
                <p class="text-muted small">تكدر تعدّل أي اسم قبل الحفظ. إذا الاسم فارغ ينحفظ باسم اليوزر.
                تاريخ الانتهاء = تاريخ العملية + {SUBSCRIPTION_DAYS} يوم. أرقام الهواتف ضيفها بعدين من التعديل.</p>
                <div class="table-responsive">
                    <table class="table table-hover">
                        <thead><tr><th></th><th>الاسم</th><th>اليوزر</th><th>نوع الكارت</th><th>المبلغ</th><th>التاريخ</th><th>الحالة</th></tr></thead>
                        <tbody>{user_rows}</tbody>
                    </table>
                </div>
            </div>

            <div class="d-flex gap-2 import-actions">
                <button class="btn btn-success flex-fill py-3"><i class="bi bi-cloud-arrow-down"></i> استيراد اليوزرات المحددة</button>
                <a href="{url_for('import_users')}" class="btn btn-outline-secondary py-3">إلغاء</a>
            </div>
        </form>
        """
        return page(body)

    body = f"""
    <h2 class="mb-4">استيراد اليوزرات</h2>
    <div class="card p-4" style="max-width:640px">
        <h5><i class="bi bi-file-earmark-pdf text-danger"></i> رفع تقرير اليوزرات من لوحة الشركة</h5>
        <p class="text-muted">ارفع ملف PDF الي تنزله من لوحة الشركة (تقرير الكارتات). البرنامج يقرأ
        اليوزر واسم العميل ونوع الكارت والمبلغ والتاريخ، ويعرضلك معاينة قبل الحفظ.</p>
        <form method="post" enctype="multipart/form-data">
            <input type="hidden" name="step" value="upload">
            <input type="file" name="file" accept="application/pdf,.pdf" class="form-control mb-3" required>
            <button class="btn btn-primary w-100 py-3"><i class="bi bi-upload"></i> قراءة الملف</button>
        </form>
        <ul class="text-muted small mt-3 mb-0">
            <li>ما ينضاف شي قبل ما تأكد بالمعاينة.</li>
            <li>اليوزر المكرر بالملف ينحسب مرة وحدة بآخر تاريخ.</li>
            <li>الاستيراد ما يخصم من المحفظة وما يسجل عمليات بالتقارير، بس يضيف المشتركين.</li>
        </ul>
    </div>
    """
    return page(body)


# ---------------------------------------------------------------------------
# الباقات والأسعار
# ---------------------------------------------------------------------------

def read_plan_form():
    category = form_text("category")
    name = form_text("name")
    retail = parse_amount(request.form.get("retail_price"))
    wholesale = parse_amount(request.form.get("wholesale_price"))
    if not category or not name:
        return None, "الفئة واسم الباقة مطلوبان"
    if retail is None or wholesale is None or retail < 0 or wholesale < 0:
        return None, "أدخل أسعاراً صحيحة"
    if wholesale > retail:
        return None, "سعر الجملة لا يمكن أن يكون أكبر من سعر البيع"
    return (category, name, retail, retail - wholesale, wholesale), None


@app.route("/plans", methods=["GET", "POST"])
@login_required
def plans():
    con = db()

    if request.method == "POST":
        values, error = read_plan_form()
        if error:
            flash(error, "danger")
        else:
            plan_id = request.form.get("plan_id")
            if plan_id:
                con.execute("""
                    UPDATE plans SET category = ?, name = ?, retail_price = ?, profit = ?, wholesale_price = ?
                    WHERE id = ?
                """, (*values, plan_id))
                flash("تم تعديل الباقة (العمليات القديمة تبقى بأسعارها الأصلية)", "success")
            else:
                con.execute("""
                    INSERT INTO plans (category, name, retail_price, profit, wholesale_price)
                    VALUES (?, ?, ?, ?, ?)
                """, values)
                flash("تمت إضافة الباقة", "success")
            con.commit()
        con.close()
        return redirect(url_for("plans"))

    rows = con.execute("""
        SELECT p.*,
               (SELECT COUNT(*) FROM subscribers s WHERE s.plan_id = p.id) AS subs,
               (SELECT COUNT(*) FROM activations a WHERE a.plan_id = p.id) AS ops
        FROM plans p ORDER BY category, id
    """).fetchall()
    con.close()

    plan_rows = ""
    for p in rows:
        in_use = p["subs"] or p["ops"]
        delete_btn = "" if in_use else f"""
            <form method="post" action="{url_for('delete_plan', plan_id=p['id'])}" class="d-inline"
                  data-confirm="حذف الباقة؟"><button class="btn btn-sm btn-outline-danger">حذف</button></form>"""
        fid = f"plan-{p['id']}"
        plan_rows += f"""
        <tr>
            <td><input form="{fid}" name="category" value="{esc(p['category'])}" class="form-control form-control-sm" required></td>
            <td><input form="{fid}" name="name" value="{esc(p['name'])}" class="form-control form-control-sm" required></td>
            <td><input form="{fid}" name="retail_price" type="number" min="0" value="{p['retail_price']}" class="form-control form-control-sm" required></td>
            <td><input form="{fid}" name="wholesale_price" type="number" min="0" value="{p['wholesale_price']}" class="form-control form-control-sm" required></td>
            <td class="text-success">{money(p['profit'])}</td>
            <td>{p['subs']}</td>
            <td class="text-nowrap">
                <form id="{fid}" method="post" class="d-inline">
                    <input type="hidden" name="plan_id" value="{p['id']}">
                    <button class="btn btn-sm btn-primary">حفظ</button>
                </form>
                {delete_btn}</td>
        </tr>"""

    body = f"""
    <h2 class="mb-4">الباقات والأسعار</h2>

    <div class="card p-4 mb-4">
        <h5>إضافة باقة جديدة</h5>
        <form method="post" class="row g-3 mt-1">
            <div class="col-md-3"><label class="form-label">الفئة</label>
                <input name="category" class="form-control" placeholder="فايبر / وايرليس" required></div>
            <div class="col-md-3"><label class="form-label">اسم الباقة</label>
                <input name="name" class="form-control" required></div>
            <div class="col-md-2"><label class="form-label">سعر البيع</label>
                <input name="retail_price" type="number" min="0" class="form-control" required></div>
            <div class="col-md-2"><label class="form-label">سعر الجملة</label>
                <input name="wholesale_price" type="number" min="0" class="form-control" required></div>
            <div class="col-md-2 d-flex align-items-end"><button class="btn btn-success w-100">إضافة</button></div>
        </form>
        <small class="text-muted mt-2">الربح = سعر البيع − سعر الجملة (يُحسب تلقائياً).</small>
    </div>

    <div class="card p-4">
        <div class="table-responsive">
            <table class="table">
                <thead><tr><th>الفئة</th><th>الباقة</th><th>سعر البيع</th><th>سعر الجملة</th>
                <th>الربح</th><th>المشتركين</th><th></th></tr></thead>
                <tbody>{plan_rows}</tbody>
            </table>
        </div>
    </div>
    """
    return page(body)


@app.route("/plans/<int:plan_id>/delete", methods=["POST"])
@login_required
def delete_plan(plan_id):
    con = db()
    used = con.execute("""
        SELECT (SELECT COUNT(*) FROM subscribers WHERE plan_id = ?)
             + (SELECT COUNT(*) FROM activations WHERE plan_id = ?) AS c
    """, (plan_id, plan_id)).fetchone()["c"]
    if used:
        flash("لا يمكن حذف باقة مستخدمة", "danger")
    else:
        con.execute("DELETE FROM plans WHERE id = ?", (plan_id,))
        con.commit()
        flash("تم حذف الباقة", "success")
    con.close()
    return redirect(url_for("plans"))


# ---------------------------------------------------------------------------
# التقارير
# ---------------------------------------------------------------------------

def report_range():
    """فترة التقرير من رابط الصفحة (افتراضياً: الشهر الحالي)."""
    start = parse_date(request.args.get("from")) or date.today().replace(day=1).strftime("%Y-%m-%d")
    end = parse_date(request.args.get("to")) or today_str()
    if start > end:
        start, end = end, start
    # نهاية اليوم تشمل كل الساعات
    return start, end, end + " 23:59:59"


def fetch_activations(con, start, end_full):
    return con.execute("""
        SELECT a.operation_type,
               s.name AS subscriber_name,
               s.username AS subscriber_username,
               s.phone,
               p.category,
               p.name AS plan_name,
               a.retail_price,
               a.payment_type,
               a.paid_amount,
               a.amount,
               a.profit,
               a.notes,
               a.operation_date
        FROM activations a
        JOIN subscribers s ON s.id = a.subscriber_id
        JOIN plans p ON p.id = a.plan_id
        WHERE a.operation_date BETWEEN ? AND ?
        ORDER BY a.id DESC
    """, (start, end_full)).fetchall()


def fetch_wallet(con, start, end_full, only_charges=False):
    return con.execute(f"""
        SELECT transaction_type, amount, notes, created_at
        FROM wallet_transactions
        WHERE created_at BETWEEN ? AND ?
        {"AND transaction_type = 'شحن'" if only_charges else ""}
        ORDER BY id DESC
    """, (start, end_full)).fetchall()


@app.route("/reports")
@login_required
def reports():
    start, end, end_full = report_range()
    con = db()
    activations = fetch_activations(con, start, end_full)
    charges = fetch_wallet(con, start, end_full, only_charges=True)
    by_plan = con.execute("""
        SELECT p.category, p.name, COUNT(*) AS ops,
               SUM(a.retail_price) AS sales, SUM(a.amount) AS cost, SUM(a.profit) AS profit
        FROM activations a JOIN plans p ON p.id = a.plan_id
        WHERE a.operation_date BETWEEN ? AND ?
        GROUP BY p.id ORDER BY profit DESC
    """, (start, end_full)).fetchall()
    con.close()

    total_sales = sum(a["retail_price"] or 0 for a in activations)
    total_cost = sum(a["amount"] or 0 for a in activations)
    total_profit = sum(a["profit"] or 0 for a in activations)
    total_charges = sum(c["amount"] for c in charges)

    activation_rows = "".join(f"""
        <tr>
            <td>{esc(a['operation_type'])}</td>
            <td>{esc(a['subscriber_name'])}</td>
            <td>{esc(a['subscriber_username'])}</td>
            <td>{esc(a['plan_name'])}</td>
            <td>{money(a['retail_price'])}</td>
            <td>{payment_badge(a['payment_type'])}</td>
            <td>{money(a['amount'])}</td>
            <td>{money(a['profit'])}</td>
            <td>{esc(a['notes'])}</td>
            <td>{esc(a['operation_date'])}</td>
        </tr>""" for a in activations)

    charge_rows = "".join(f"""
        <tr>
            <td>{money(c['amount'])}</td>
            <td>{esc(c['notes'])}</td>
            <td>{esc(c['created_at'])}</td>
        </tr>""" for c in charges)

    plan_rows = "".join(f"""
        <tr>
            <td>{esc(p['category'])} - {esc(p['name'])}</td>
            <td>{p['ops']}</td>
            <td>{money(p['sales'])}</td>
            <td>{money(p['cost'])}</td>
            <td class="text-success">{money(p['profit'])}</td>
        </tr>""" for p in by_plan)

    qs = f"?from={start}&to={end}"

    body = f"""
    <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-4 page-head">
        <h2 class="mb-0">التقارير</h2>
        <button onclick="window.print()" class="btn btn-outline-secondary no-print"><i class="bi bi-printer"></i> طباعة</button>
    </div>

    <div class="card p-4 mb-4 no-print">
        <form method="get" class="row g-3 align-items-end">
            <div class="col-md-3"><label class="form-label">من تاريخ</label>
                <input type="date" name="from" value="{start}" class="form-control"></div>
            <div class="col-md-3"><label class="form-label">إلى تاريخ</label>
                <input type="date" name="to" value="{end}" class="form-control"></div>
            <div class="col-md-2"><button class="btn btn-primary w-100">عرض</button></div>
            <div class="col-md-4 d-flex flex-wrap gap-2 justify-content-md-end">
                <a href="{url_for('export_activations')}{qs}" class="btn btn-success"><i class="bi bi-file-earmark-excel"></i> Excel العمليات</a>
                <a href="{url_for('export_charges')}{qs}" class="btn btn-primary"><i class="bi bi-file-earmark-excel"></i> Excel المحفظة</a>
            </div>
        </form>
    </div>

    <p class="text-muted">الفترة: من <b>{start}</b> إلى <b>{end}</b></p>

    <div class="row g-3 mb-4">
        <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>عدد العمليات</h6><h3>{len(activations)}</h3></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>المبيعات</h6><h3>{money(total_sales)}</h3></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>الخصم من المحفظة</h6><h3 class="text-danger">{money(total_cost)}</h3></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>صافي الربح</h6><h3 class="text-success">{money(total_profit)}</h3></div></div>
    </div>

    <div class="card p-4 mb-4">
        <h5>الأرباح حسب الباقة</h5>
        <div class="table-responsive">
            <table class="table table-striped">
                <thead><tr><th>الباقة</th><th>العمليات</th><th>المبيعات</th><th>الجملة</th><th>الربح</th></tr></thead>
                <tbody>{plan_rows or '<tr><td colspan="5" class="text-muted text-center">لا توجد بيانات</td></tr>'}</tbody>
            </table>
        </div>
    </div>

    <div class="card p-4 mb-4">
        <h5>تقرير التفعيل والترحيل</h5>
        <div class="table-responsive">
            <table class="table table-striped">
                <thead><tr><th>العملية</th><th>المشترك</th><th>اليوزر</th><th>الاشتراك</th><th>سعر البيع</th><th>الدفع</th>
                <th>سعر الجملة</th><th>الربح</th><th>الملاحظات</th><th>التاريخ</th></tr></thead>
                <tbody>{activation_rows or '<tr><td colspan="10" class="text-muted text-center">لا توجد عمليات بهذه الفترة</td></tr>'}</tbody>
            </table>
        </div>
    </div>

    <div class="card p-4">
        <h5>تقرير شحن المحفظة — المجموع: {money(total_charges)} د.ع</h5>
        <div class="table-responsive">
            <table class="table table-striped">
                <thead><tr><th>المبلغ</th><th>الملاحظات</th><th>التاريخ</th></tr></thead>
                <tbody>{charge_rows or '<tr><td colspan="3" class="text-muted text-center">لا يوجد شحن بهذه الفترة</td></tr>'}</tbody>
            </table>
        </div>
    </div>
    """
    return page(body)


def excel_response(title, headers, rows, filename, totals=None):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = title
    sheet.sheet_view.rightToLeft = True
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        sheet.append(list(row))
    if totals:
        sheet.append(totals)
        for cell in sheet[sheet.max_row]:
            cell.font = Font(bold=True)
    for col in sheet.columns:
        width = max(len(str(c.value or "")) for c in col) + 3
        sheet.column_dimensions[col[0].column_letter].width = min(width, 45)

    file = BytesIO()
    workbook.save(file)
    file.seek(0)
    return send_file(
        file, as_attachment=True, download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/export/activations")
@login_required
def export_activations():
    start, end, end_full = report_range()
    con = db()
    rows = fetch_activations(con, start, end_full)
    con.close()
    headers = ["نوع العملية", "اسم المشترك", "اليوزر", "رقم الهاتف", "الفئة", "نوع الاشتراك",
               "سعر البيع", "طريقة الدفع", "المقبوض", "سعر الجملة (المخصوم)", "الربح", "الملاحظات", "التاريخ"]
    totals = ["المجموع", "", "", "", "", "",
              sum(r["retail_price"] or 0 for r in rows), "",
              sum(r["paid_amount"] or 0 for r in rows),
              sum(r["amount"] or 0 for r in rows),
              sum(r["profit"] or 0 for r in rows), "", ""]
    return excel_response("تقرير اليوزرات", headers, rows,
                          f"users_report_{start}_{end}.xlsx", totals)


@app.route("/export/charges")
@login_required
def export_charges():
    start, end, end_full = report_range()
    con = db()
    rows = fetch_wallet(con, start, end_full)
    con.close()
    headers = ["نوع الحركة", "المبلغ", "الملاحظات", "التاريخ"]
    return excel_response("حركات المحفظة", headers, rows, f"wallet_{start}_{end}.xlsx")


# ---------------------------------------------------------------------------
# الإعدادات: كلمة المرور، المستخدمين، النسخ الاحتياطي
# ---------------------------------------------------------------------------

@app.route("/settings", methods=["GET", "POST"])
@login_required
def settings():
    con = db()

    if request.method == "POST":
        action = request.form.get("action")

        if action == "password":
            user = con.execute("SELECT * FROM users WHERE id = ?", (session["user_id"],)).fetchone()
            current = request.form.get("current", "")
            new = request.form.get("new", "")
            if not user or not check_password_hash(user["password"], current):
                flash("كلمة المرور الحالية غير صحيحة", "danger")
            elif len(new) < 6:
                flash("كلمة المرور الجديدة يجب أن تكون 6 أحرف على الأقل", "danger")
            elif new != request.form.get("confirm", ""):
                flash("تأكيد كلمة المرور غير مطابق", "danger")
            else:
                con.execute("UPDATE users SET password = ? WHERE id = ?",
                            (generate_password_hash(new), user["id"]))
                con.commit()
                flash("تم تغيير كلمة المرور", "success")

        elif action == "shop":
            set_setting(con, "shop_name", form_text("shop_name"))
            con.commit()
            flash("تم حفظ اسم المحل", "success")

        elif action == "add_user":
            username = form_text("username")
            password = request.form.get("password", "")
            if not username or len(password) < 6:
                flash("أدخل اسم مستخدم وكلمة مرور (6 أحرف على الأقل)", "danger")
            elif con.execute("SELECT 1 FROM users WHERE username = ?", (username,)).fetchone():
                flash("اسم المستخدم موجود مسبقاً", "danger")
            else:
                con.execute("INSERT INTO users(username, password) VALUES (?, ?)",
                            (username, generate_password_hash(password)))
                con.commit()
                flash("تمت إضافة المستخدم", "success")

        elif action == "delete_user":
            uid = parse_amount(request.form.get("user_id"))
            if uid == session["user_id"]:
                flash("لا يمكنك حذف حسابك الحالي", "danger")
            else:
                con.execute("DELETE FROM users WHERE id = ?", (uid,))
                con.commit()
                flash("تم حذف المستخدم", "success")

        con.close()
        return redirect(url_for("settings"))

    users = con.execute("SELECT id, username FROM users ORDER BY id").fetchall()
    shop = get_setting(con, "shop_name")
    con.close()

    user_rows = "".join(f"""
        <tr>
            <td>{esc(u['username'])}</td>
            <td>{'<span class="badge bg-primary">أنت</span>' if u['id'] == session['user_id'] else f'''
                <form method="post" data-confirm="حذف المستخدم؟">
                    <input type="hidden" name="action" value="delete_user">
                    <input type="hidden" name="user_id" value="{u['id']}">
                    <button class="btn btn-sm btn-outline-danger">حذف</button>
                </form>'''}</td>
        </tr>""" for u in users)

    body = f"""
    <h2 class="mb-4">الإعدادات</h2>

    <div class="row g-4">
        <div class="col-md-6">
            <div class="card p-4 h-100">
                <h5><i class="bi bi-key text-primary"></i> تغيير كلمة المرور</h5>
                <form method="post">
                    <input type="hidden" name="action" value="password">
                    <label class="form-label">كلمة المرور الحالية</label>
                    <input type="password" name="current" class="form-control mb-2" required>
                    <label class="form-label">كلمة المرور الجديدة</label>
                    <input type="password" name="new" minlength="6" class="form-control mb-2" required>
                    <label class="form-label">تأكيد كلمة المرور</label>
                    <input type="password" name="confirm" minlength="6" class="form-control mb-3" required>
                    <button class="btn btn-primary">تغيير</button>
                </form>
            </div>
        </div>

        <div class="col-md-6">
            <div class="card p-4 h-100">
                <h5><i class="bi bi-person-badge text-primary"></i> المستخدمون</h5>
                <table class="table"><tbody>{user_rows}</tbody></table>
                <form method="post" class="row g-2">
                    <input type="hidden" name="action" value="add_user">
                    <div class="col-5"><input name="username" class="form-control" placeholder="اسم المستخدم" required></div>
                    <div class="col-4"><input name="password" type="password" minlength="6" class="form-control" placeholder="كلمة المرور" required></div>
                    <div class="col-3"><button class="btn btn-success w-100">إضافة</button></div>
                </form>
            </div>
        </div>

        <div class="col-12">
            <div class="card p-4">
                <h5><i class="bi bi-whatsapp" style="color:#25D366"></i> اسم المحل (توقيع رسائل واتساب)</h5>
                <form method="post" class="row g-2">
                    <input type="hidden" name="action" value="shop">
                    <div class="col-md-8"><input name="shop_name" value="{esc(shop)}" class="form-control" placeholder="مثال: مكتب الأمل للإنترنت - 0770xxxxxxx"></div>
                    <div class="col-md-4"><button class="btn btn-primary w-100">حفظ</button></div>
                </form>
            </div>
        </div>

        <div class="col-12">
            <div class="card p-4">
                <h5><i class="bi bi-cloud-arrow-down text-primary"></i> النسخ الاحتياطي</h5>
                <p class="text-muted mb-3">نزّل نسخة كاملة من قاعدة البيانات واحتفظ بيها بمكان آمن (يُنصح بعمل نسخة كل أسبوع).</p>
                <div><a href="{url_for('backup')}" class="btn btn-dark"><i class="bi bi-download"></i> تحميل نسخة احتياطية</a></div>
            </div>
        </div>
    </div>
    """
    return page(body)


@app.route("/backup")
@login_required
def backup():
    fd, path = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    try:
        src = db()
        dst = sqlite3.connect(path)
        src.backup(dst)
        dst.close()
        src.close()
        with open(path, "rb") as f:
            data = BytesIO(f.read())
    finally:
        os.remove(path)
    return send_file(
        data, as_attachment=True,
        download_name=f"backup_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.db",
        mimetype="application/octet-stream"
    )


@app.errorhandler(400)
def bad_request(e):
    return page(f"""
        <div class="card p-4 mt-5 text-center">
            <h4>⚠️ {esc(getattr(e, 'description', 'طلب غير صالح'))}</h4>
            <div class="mt-3"><a href="javascript:history.back()" class="btn btn-primary">رجوع</a></div>
        </div>"""), 400


# تهيئة قاعدة البيانات عند التشغيل (سواء بـ python app.py أو gunicorn)
init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
