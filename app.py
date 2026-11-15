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

# (الفئة، الباقة، سعر البيع، سعر الجملة "يخصم من المحفظة"، سعر "من قبلي"، أسماء الكارت بملف الشركة)
PLANS = [
    ("فايبر", "Basic", 38000, 32000, 40000, ["basic"]),
    ("فايبر", "Plus", 45000, 39000, 45000, ["plus"]),
    ("فايبر", "Turbo", 65000, 58000, 65000, ["turbo"]),
    ("فايبر", "Pro Max", 100000, 90000, 100000, ["pro max", "promax", "pro-max"]),
    ("وايرليس", "ايكو", 36000, 31000, 36000, ["economy", "eco", "eco-w", "economy-w"]),
    ("وايرليس", "بلص", 38000, 33000, 38000, ["plus-w"]),
    ("وايرليس", "استاندر", 45000, 40000, 45000, ["standard", "standard-w"]),
    ("وايرليس", "تيربو", 65000, 60000, 65000, ["turbo-w"]),
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


def precise_now():
    """وقت دقيق (بالأجزاء من الثانية) لترتيب التحصيل والشحن."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")


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

    # مبالغ قديمة بذمة المشتركين (من قبل البرنامج) بحساب الدفع من قبلي
    cur.execute("""
        CREATE TABLE IF NOT EXISTS self_old (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            subscriber_id INTEGER NOT NULL,
            amount INTEGER NOT NULL,
            notes TEXT,
            created_at TEXT NOT NULL,
            collected_at TEXT
        )
    """)

    # محفظة حساب "يوزرات الدفع من قبلي": شحن (رصيدك) / خصم (جملة كل يوزر) / استلام (لما يرجّعلك)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS self_wallet (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kind TEXT NOT NULL,
            amount INTEGER NOT NULL,
            notes TEXT,
            activation_id INTEGER,
            created_at TEXT NOT NULL
        )
    """)

    # وقت ما رجّع المشترك فلوس "من قبلي" (فارغ = بعده ما مسدد إلك)
    if not column_exists(cur, "activations", "self_collected_at"):
        cur.execute("ALTER TABLE activations ADD COLUMN self_collected_at TEXT")

    # سعر البيع عند "من قبلي" لكل باقة
    if not column_exists(cur, "plans", "self_price"):
        cur.execute("ALTER TABLE plans ADD COLUMN self_price INTEGER")

    # نوع المستخدم: admin = البرنامج الرئيسي، self = حساب الدفع من قبلي بس
    if not column_exists(cur, "users", "role"):
        cur.execute("ALTER TABLE users ADD COLUMN role TEXT NOT NULL DEFAULT 'admin'")

    # المبالغ السابقة تنخصم من محفظة حساب الدفع من قبلي (مربوطة بـ old_id)
    if not column_exists(cur, "self_wallet", "old_id"):
        cur.execute("ALTER TABLE self_wallet ADD COLUMN old_id INTEGER")
        for o in cur.execute("""
            SELECT o.id, o.amount, s.username FROM self_old o LEFT JOIN subscribers s ON s.id = o.subscriber_id
        """).fetchall():
            cur.execute("INSERT INTO self_wallet (kind, amount, notes, old_id, created_at) VALUES ('خصم', ?, ?, ?, ?)",
                        (o["amount"], f"مبلغ سابق - {o['username'] or ''}", o["id"], now_str()))
        cur.execute("UPDATE self_wallet SET old_id = (SELECT o.id FROM self_old o WHERE o.collected_at IS NOT NULL "
                    "AND self_wallet.notes LIKE 'مبلغ سابق من %' AND o.amount = self_wallet.amount LIMIT 1) "
                    "WHERE kind = 'استلام' AND notes LIKE 'مبلغ سابق من %'")

    # حالة الدفع لكل تفعيل: نقد (محصّل) / دين (معلّق) / من قبلي (دفعته من جيبي)
    # settled_at = وقت التحصيل نقداً أو التحويل لحسابي الخاص
    if not column_exists(cur, "activations", "settled_at"):
        cur.execute("ALTER TABLE activations ADD COLUMN settled_at TEXT")
        cur.execute("UPDATE activations SET settled_at = operation_date WHERE payment_type = 'نقد'")
        # الديون الي تسددت كاملة بالنسخة السابقة تصير نقد
        for row in cur.execute("""
            SELECT subscriber_id,
                   SUM(CASE WHEN kind = 'دين' THEN amount ELSE -amount END) AS bal,
                   MAX(created_at) AS last
            FROM debts GROUP BY subscriber_id
        """).fetchall():
            if row["bal"] <= 0:
                cur.execute("""
                    UPDATE activations SET payment_type = 'نقد', settled_at = ?
                    WHERE subscriber_id = ? AND payment_type = 'دين'
                """, (row["last"], row["subscriber_id"]))

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
        seed_plans(cur)

    con.commit()
    con.close()


def seed_plans(cur):
    """إضافة قائمة الباقات الافتراضية."""
    for category, name, retail, wholesale, self_price, _ in PLANS:
        cur.execute("""
            INSERT INTO plans (category, name, retail_price, profit, wholesale_price, self_price)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (category, name, retail, retail - wholesale, wholesale, self_price))


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


PAY_CASH, PAY_DEBT, PAY_SELF = "نقد", "دين", "من قبلي"
PAY_TYPES = (PAY_CASH, PAY_DEBT, PAY_SELF)

def self_price_of(plan):
    """سعر البيع عند الدفع "من قبلي" (إذا ما محدد = سعر البيع)."""
    return plan["self_price"] or plan["retail_price"]


# دين المشترك = مجموع أسعار الجملة (المخصومة من المحفظة) للتفعيلات الي حالتها دين
# (داخل استعلام فيه s = المشتركين)
DEBT_SUBQUERY = """
    SELECT COALESCE(SUM(a.amount), 0) FROM activations a
    WHERE a.subscriber_id = s.id AND a.payment_type = 'دين'
"""
# مبلغ الدين بسعر البيع (للرسائل الي تنرسل للمشترك)
DEBT_SALE_SUBQUERY = """
    SELECT COALESCE(SUM(a.retail_price), 0) FROM activations a
    WHERE a.subscriber_id = s.id AND a.payment_type = 'دين'
"""
# المبلغ الي بذمة المشترك إلك (دفعته عنه "من قبلي" وبعده ما رجّعه)
SELF_SUBQUERY = """
    (SELECT COALESCE(SUM(a.retail_price), 0) FROM activations a
     WHERE a.subscriber_id = s.id AND a.payment_type = 'من قبلي' AND a.self_collected_at IS NULL)
    + (SELECT COALESCE(SUM(o.amount), 0) FROM self_old o
       WHERE o.subscriber_id = s.id AND o.collected_at IS NULL)
"""


PREV_OP = "اشتراك سابق"

SW_TOPUP, SW_OUT, SW_BACK, SW_WITHDRAW = "شحن", "خصم", "استلام", "سحب"


def self_wallet_balance(con):
    return con.execute("""
        SELECT COALESCE(SUM(CASE WHEN kind IN ('شحن', 'استلام') THEN amount ELSE -amount END), 0) AS b
        FROM self_wallet
    """).fetchone()["b"]


def self_wallet_on_self(con, activation_id):
    """لما تتحول عملية لـ"من قبلي": ينخصم سعر جملتها من محفظة حساب الدفع من قبلي."""
    op = con.execute("""
        SELECT a.amount, s.username, p.name AS plan_name FROM activations a
        JOIN subscribers s ON s.id = a.subscriber_id JOIN plans p ON p.id = a.plan_id WHERE a.id = ?
    """, (activation_id,)).fetchone()
    if not op or con.execute("SELECT 1 FROM self_wallet WHERE activation_id = ? AND kind = 'خصم'",
                             (activation_id,)).fetchone():
        return
    con.execute("INSERT INTO self_wallet (kind, amount, notes, activation_id, created_at) VALUES ('خصم', ?, ?, ?, ?)",
                (op["amount"], f"{op['username']} - {op['plan_name']}", activation_id, now_str()))


def self_wallet_clear(con, activation_id, kinds=("خصم", "استلام")):
    con.execute(f"DELETE FROM self_wallet WHERE activation_id = ? AND kind IN ({','.join('?' * len(kinds))})",
                (activation_id, *kinds))


def record_previous(con, subscriber_id, plan_id, start, retail, payment_type, wholesale=None, notes=""):
    """تسجيل اشتراك سابق (مفعّل من قبل) بدون خصم من المحفظة.
    يرجع False إذا عنده عملية بنفس التاريخ (حتى ما يتكرر)."""
    plan = con.execute("SELECT * FROM plans WHERE id = ?", (plan_id,)).fetchone()
    if not plan or payment_type not in PAY_TYPES:
        return False
    if con.execute("SELECT 1 FROM activations WHERE subscriber_id = ? AND substr(operation_date, 1, 10) = ?",
                   (subscriber_id, start)).fetchone():
        return False
    wholesale = plan["wholesale_price"] if wholesale is None else wholesale
    if retail is None:
        retail = self_price_of(plan) if payment_type == PAY_SELF else plan["retail_price"]
    when = f"{start} 00:00:00"
    con.execute("""
        INSERT INTO activations
        (subscriber_id, plan_id, operation_type, amount, retail_price, profit,
         payment_type, paid_amount, settled_at, notes, operation_date)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (subscriber_id, plan_id, PREV_OP, wholesale, retail, retail - wholesale, payment_type,
          retail if payment_type == PAY_CASH else 0,
          None if payment_type == PAY_DEBT else when, notes, when))
    if payment_type == PAY_SELF:
        self_wallet_on_self(con, con.execute("SELECT last_insert_rowid()").fetchone()[0])
    con.execute("""
        UPDATE subscribers SET plan_id = ?, activation_date = ?, expiry_date = ?
        WHERE id = ? AND (activation_date IS NULL OR activation_date <= ?)
    """, (plan_id, start, expiry_from(start), subscriber_id, start))
    return True


def quick_prices(con, *extra):
    """أسعار البيع السريعة (من الإعدادات) + سعر الباقة وسعر من قبلي."""
    raw = get_setting(con, "quick_prices", "")
    prices = {parse_amount(x) for x in raw.split(",")}
    prices.update(e for e in extra if e)
    return sorted(p for p in prices if p and p > 0)


def get_setting(con, key, default=""):
    row = con.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
    return row["value"] if row and row["value"] is not None else default


def set_setting(con, key, value):
    con.execute("INSERT INTO settings(key, value) VALUES (?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value = excluded.value", (key, value))


def payment_badge(kind):
    if kind == PAY_DEBT:
        return '<span class="badge bg-danger-subtle text-danger-emphasis">دين</span>'
    if kind == PAY_SELF:
        return '<span class="badge bg-info-subtle text-info-emphasis">من قبلي</span>'
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
    if op["payment_type"] == PAY_DEBT:
        text += f"\nالمبلغ المتبقي بذمتك: {money(debt)} د.ع"
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


ROLE_ADMIN, ROLE_SELF = "admin", "self"
# صفحات حساب الدفع من قبلي (تفتح بس بيوزر هذا الحساب)
SELF_ENDPOINTS = {"self_account", "self_collect", "self_old_add", "self_old_delete",
                  "self_wallet_move", "self_uncollect"}
PUBLIC_ENDPOINTS = {"login", "logout", "static", "manifest", "service_worker", "offline"}


def home_url():
    return url_for("self_account") if session.get("role") == ROLE_SELF else url_for("dashboard")


@app.before_request
def role_access():
    """كل يوزر يشوف حسابه بس: الرئيسي ما يفتح حساب الدفع من قبلي، وبالعكس."""
    ep = request.endpoint
    if not session.get("user_id") or ep is None or ep in PUBLIC_ENDPOINTS:
        return None
    is_self = session.get("role") == ROLE_SELF
    if is_self and ep not in SELF_ENDPOINTS:
        return redirect(url_for("self_account"))
    if not is_self and ep in SELF_ENDPOINTS:
        flash("حساب الدفع من قبلي له دخول منفصل — سجّل خروج وادخل بيوزره", "warning")
        return redirect(url_for("dashboard"))
    return None


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
    <meta name="mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-title" content="محاسبة الإنترنت">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
    <link rel="manifest" href="/manifest.webmanifest">
    <link rel="icon" type="image/png" href="/static/favicon.png">
    <link rel="apple-touch-icon" href="/static/icons/apple-touch-icon.png">
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
        .pay-options { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; }
        .pay-opt {
            display: flex; flex-direction: column; align-items: center; gap: 2px; cursor: pointer;
            padding: 16px 10px; border-radius: 16px; border: 2px solid var(--line); background: var(--soft);
            text-align: center; transition: .15s;
        }
        .pay-opt i { font-size: 1.7rem; color: #10b981; }
        .pay-opt.debt i { color: #ef4444; }
        .pay-opt.self i { color: #0ea5e9; }
        .pay-opt b { font-size: 1.05rem; }
        .pay-opt small { color: var(--muted); }
        .btn-check:checked + .pay-opt { border-color: #10b981; background: rgba(16,185,129,.08); }
        .btn-check:checked + .pay-opt.debt { border-color: #ef4444; background: rgba(239,68,68,.08); }
        .btn-check:checked + .pay-opt.self { border-color: #0ea5e9; background: rgba(14,165,233,.08); }
        @media (max-width: 400px) { .pay-opt { padding: 12px 4px; } .pay-opt small { font-size: .72rem; } }
        .bulk-bar { display: none; position: sticky; bottom: 12px; z-index: 1010; margin-top: 12px;
                    border: 2px solid #f59e0b; }
        .bulk-bar.show { display: block; }
        @media (max-width: 991px) { .bulk-bar { bottom: calc(80px + env(safe-area-inset-bottom)); } }
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
    'settings': 'الإعدادات', 'debts': 'الديون', 'collection': 'تحصيل المبالغ', 'self_account': 'الدفع من قبلي', 'settle_operation': 'ترحيل', 'import_users': 'استيراد اليوزرات', 'wa_import': 'تفعيلات واتساب', 'activate': 'تفعيل اشتراك', 'transfer': 'ترحيل'
} %}
{% set ep = request.endpoint %}

{% if session.get('user_id') and session.get('role') == 'self' %}
<div class="layout">
    <aside class="sidebar">
        <div class="brand"><span class="logo" style="background:linear-gradient(135deg,#0284c7,#0ea5e9)"><i class="bi bi-piggy-bank"></i></span> الدفع من قبلي</div>
        <nav class="nav-side">
            <a href="{{ url_for('self_account') }}" class="active"><i class="bi bi-piggy-bank"></i> يوزرات الدفع من قبلي</a>
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
                <span class="brand p-0"><span class="logo" style="width:34px;height:34px;font-size:1rem;background:linear-gradient(135deg,#0284c7,#0ea5e9)"><i class="bi bi-piggy-bank"></i></span></span>
                <span class="title">الدفع من قبلي</span>
            </div>
            <div class="d-flex gap-2">
                <button class="icon-btn" onclick="toggleTheme()" aria-label="الوضع الليلي"><i class="bi bi-moon-stars"></i></button>
                <a class="icon-btn" href="{{ url_for('logout') }}" aria-label="تسجيل الخروج"><i class="bi bi-box-arrow-left"></i></a>
            </div>
        </header>
        <main class="main" style="padding-bottom:40px">
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
{% elif session.get('user_id') %}
<div class="layout">
    <aside class="sidebar">
        <div class="brand"><span class="logo"><i class="bi bi-router"></i></span> محاسبة الإنترنت</div>
        <nav class="nav-side">
            <a href="{{ url_for('dashboard') }}" class="{{ 'active' if ep == 'dashboard' }}"><i class="bi bi-grid-1x2"></i> الرئيسية</a>
            <a href="{{ url_for('subscribers') }}" class="{{ 'active' if ep in ['subscribers','subscriber_detail','activate','transfer','import_users','wa_import'] }}"><i class="bi bi-people"></i> المشتركون</a>
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
    <button data-bs-toggle="offcanvas" data-bs-target="#more" class="{{ 'active' if ep in ['plans','settings','reports','self_account','wa_import'] }}"><i class="bi bi-three-dots"></i>المزيد</button>
</nav>

<div class="offcanvas offcanvas-bottom" tabindex="-1" id="more" style="height:auto;border-radius:22px 22px 0 0;background:var(--card)">
    <div class="offcanvas-body nav-side pb-4">
        <div class="who mb-2"><span class="avatar">{{ (session.get('username') or '?')[:1]|upper }}</span> {{ session.get('username') }}</div>
        <a href="{{ url_for('wa_import') }}"><i class="bi bi-whatsapp"></i> تفعيلات من واتساب الشركة</a>
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
    // تسجيل التطبيق (PWA) حتى ينثبت مثل تطبيق
    if ('serviceWorker' in navigator) {
        window.addEventListener('load', () => navigator.serviceWorker.register('/sw.js').catch(() => {}));
    }
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
        return redirect(home_url())

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
            session["role"] = user["role"] if "role" in user.keys() and user["role"] else ROLE_ADMIN
            if token:
                session["csrf_token"] = token
            return redirect(home_url())

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

def money_totals(con, since=None, until=None):
    """مجاميع الحسابات حسب تاريخ الاستلام/الترحيل:
    - المحصّل = مبالغ النقد + مبالغ "من قبلي"
    - الأرباح = من النقد فقط
    - الديون = الحالية (كل الفترات)"""
    cond, params = "", []
    if since:
        cond += " AND settled_at >= ?"
        params.append(since)
    if until:
        cond += " AND settled_at <= ?"
        params.append(until)
    cash = con.execute(f"""
        SELECT COUNT(*) AS n, COALESCE(SUM(retail_price), 0) AS total, COALESCE(SUM(profit), 0) AS profit
        FROM activations WHERE payment_type = 'نقد' {cond}
    """, params).fetchone()
    mine = con.execute(f"""
        SELECT COUNT(*) AS n, COALESCE(SUM(retail_price), 0) AS total
        FROM activations WHERE payment_type = 'من قبلي' {cond}
    """, params).fetchone()
    debt = con.execute("""
        SELECT COUNT(DISTINCT subscriber_id) AS n, COALESCE(SUM(amount), 0) AS total
        FROM activations WHERE payment_type = 'دين'
    """).fetchone()
    return {"cash_n": cash["n"], "cash": cash["total"], "profit": cash["profit"],
            "self_n": mine["n"], "self": mine["total"],
            "collected": cash["total"] + mine["total"], "collected_n": cash["n"] + mine["n"],
            "debt_n": debt["n"], "debt": debt["total"]}


def last_charge_time(con):
    """وقت آخر شحن (يتسجل بدقة وقت الشحن حتى يتصفر التحصيل)."""
    marker = get_setting(con, "collection_reset_at")
    if marker:
        return marker
    row = con.execute("SELECT MAX(created_at) AS t FROM wallet_transactions WHERE transaction_type = 'شحن'").fetchone()
    return row["t"]


def collection_since_charge(con):
    """تحصيل المبالغ منذ آخر شحن للمحفظة (يتصفر مع كل شحن)."""
    last = last_charge_time(con)
    return money_totals(con, since=last), last


def stat_card(title, value, icon, color, sub="", href=None, cls="col-6 col-md-4"):
    inner = f"""
            <div class="card stat-card p-4 h-100">
                <span class="stat-icon" style="background:{color}22;color:{color}"><i class="bi {icon}"></i></span>
                <h6>{title}</h6>
                <h2 style="color:{color}">{value}</h2>
                {f'<small class="text-muted">{sub}</small>' if sub else ''}
            </div>"""
    if href:
        inner = f'<a href="{href}" class="text-decoration-none">{inner}</a>'
    return f'<div class="{cls}">{inner}</div>'


@app.route("/dashboard")
@login_required
def dashboard():
    con = db()
    month_start = date.today().replace(day=1).strftime("%Y-%m-%d")

    subscribers_count = con.execute("SELECT COUNT(*) AS c FROM subscribers").fetchone()["c"]
    month = money_totals(con, since=month_start)
    total = money_totals(con)
    month_ops = con.execute("SELECT COUNT(*) AS c FROM activations WHERE operation_date >= ?",
                            (month_start,)).fetchone()["c"]

    warn_until = (date.today() + timedelta(days=EXPIRY_WARNING_DAYS)).strftime("%Y-%m-%d")
    expiring = con.execute("""
        SELECT s.*, p.name AS plan_name
        FROM subscribers s JOIN plans p ON p.id = s.plan_id
        WHERE s.expiry_date <= ?
        ORDER BY s.expiry_date DESC
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
    shop = get_setting(con, "shop_name")
    since, last_charge = collection_since_charge(con)
    con.close()

    expiring_rows = "".join(f"""
        <tr>
            <td>{esc(s['name'])}</td>
            <td>{esc(s['username'])}</td>
            <td dir="ltr">{esc(s['phone'])}</td>
            <td>{esc(s['plan_name'])}</td>
            <td>{esc(s['expiry_date'])}</td>
            <td>{status_badge(s['expiry_date'])}</td>
            <td><div class="d-flex gap-1">
                <a href="{url_for('activate', subscriber_id=s['id'])}" class="btn btn-sm btn-success">تفعيل</a>
                {wa_button(s['phone'], msg_reminder(s, shop), 'تذكير')}
            </div></td>
        </tr>""" for s in expiring) or '<tr><td colspan="7" class="text-muted text-center">لا توجد اشتراكات منتهية أو قريبة الانتهاء 👍</td></tr>'

    recent_rows = "".join(f"""
        <tr>
            <td>{esc(a['operation_type'])}</td>
            <td>{esc(a['subscriber_username'])}</td>
            <td>{esc(a['plan_name'])}</td>
            <td>{money(a['retail_price'])}</td>
            <td>{payment_badge(a['payment_type'])}</td>
            <td>{esc(a['operation_date'])}</td>
        </tr>""" for a in recent) or '<tr><td colspan="6" class="text-muted text-center">لا توجد عمليات بعد</td></tr>'

    cards = "".join([
        f"""<div class="col-6 col-md-4">
            <div class="card balance-card stat-card p-4 h-100">
                <span class="stat-icon"><i class="bi bi-wallet2"></i></span>
                <h6>رصيد المحفظة</h6>
                <h2>{money(balance)} د.ع</h2>
                <small>{subscribers_count} مشترك — {month_ops} تفعيل هذا الشهر</small>
            </div>
        </div>""",
        stat_card("تحصيل المبالغ", f"{money(since['collected'])} د.ع", "bi-cash-stack", "#4f46e5",
                  f"منذ آخر شحن ({esc((last_charge or 'البداية')[:16])}) — نقد {money(since['cash'])} + من قبلي {money(since['self'])} — اضغط للتفاصيل",
                  url_for("collection")),
        stat_card("الأرباح هذا الشهر", f"{money(month['profit'])} د.ع", "bi-graph-up-arrow", "#10b981",
                  f"من النقد فقط — الإجمالي: {money(total['profit'])} د.ع"),
        stat_card("الديون", f"{money(total['debt'])} د.ع", "bi-journal-text", "#ef4444",
                  f"على {total['debt_n']} مشترك", url_for("debts"), cls="col-12 col-md-4"),
    ])

    body = f"""
    <h2 class="mb-4">لوحة التحكم</h2>

    <div class="row g-3 mb-4">{cards}</div>

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
                <thead><tr><th>العملية</th><th>اليوزر</th><th>الاشتراك</th><th>سعر البيع</th><th>الدفع</th><th>التاريخ</th></tr></thead>
                <tbody>{recent_rows}</tbody>
            </table>
        </div>
    </div>
    """
    return page(body)


# ---------------------------------------------------------------------------
# المحفظة
# ---------------------------------------------------------------------------

@app.route("/collection")
@login_required
def collection():
    """تفاصيل تحصيل المبالغ: اليوزرات الي استلمت منهم من آخر شحن (يتصفر مع كل شحن)."""
    con = db()
    since, last = collection_since_charge(con)
    cond, params = "", []
    if last:
        cond, params = "AND a.settled_at >= ?", [last]
    rows = con.execute(f"""
        SELECT a.*, s.name, s.username, p.name AS plan_name
        FROM activations a JOIN subscribers s ON s.id = a.subscriber_id JOIN plans p ON p.id = a.plan_id
        WHERE a.payment_type IN ('نقد', 'من قبلي') AND a.settled_at IS NOT NULL {cond}
        ORDER BY a.settled_at DESC
    """, params).fetchall()
    charges = con.execute("""
        SELECT amount, notes, created_at FROM wallet_transactions
        WHERE transaction_type = 'شحن' ORDER BY id DESC LIMIT 20
    """).fetchall()
    balance = wallet_balance(con)
    con.close()

    table = "".join(f"""
        <tr>
            <td><a href="{url_for('subscriber_detail', subscriber_id=r['subscriber_id'])}" class="fw-bold text-decoration-none">{esc(r['name'])}</a></td>
            <td dir="ltr">{esc(r['username'])}</td>
            <td>{esc(r['plan_name'])}</td>
            <td class="fw-bold">{money(r['retail_price'])} د.ع</td>
            <td>{payment_badge(r['payment_type'])}</td>
            <td class="text-success">{money(r['profit']) if r['payment_type'] == PAY_CASH else '—'}</td>
            <td>{esc((r['settled_at'] or '')[:16])}</td>
        </tr>""" for r in rows)

    def before(note):
        m = re.search(r"التحصيل قبل الشحن: ([\d,]+)", note or "")
        return m.group(1) if m else "—"
    charge_rows = "".join(f"""
        <tr>
            <td>{esc(c['created_at'][:16])}</td>
            <td class="text-success fw-bold">+{money(c['amount'])}</td>
            <td>{before(c['notes'])}</td>
        </tr>""" for c in charges)

    body = f"""
    <h2 class="mb-4">تحصيل المبالغ</h2>

    <div class="row g-3 mb-4">
        <div class="col-12 col-md-4">
            <div class="card stat-card p-4" style="background:linear-gradient(135deg,#4f46e5,#7c3aed);color:#fff;border:0">
                <span class="stat-icon" style="background:rgba(255,255,255,.18);color:#fff"><i class="bi bi-cash-stack"></i></span>
                <h6 style="color:rgba(255,255,255,.85)">المحصّل منذ آخر شحن</h6>
                <h2>{money(since['collected'])} د.ع</h2>
                <small style="color:rgba(255,255,255,.85)">آخر شحن: {esc((last or 'ماكو')[:16])} — {len(rows)} يوزر</small>
            </div>
        </div>
        <div class="col-6 col-md-4"><div class="card stat-card p-4"><h6>نقد</h6>
            <h3 class="text-success">{money(since['cash'])} د.ع</h3><small class="text-muted">ربح {money(since['profit'])} د.ع</small></div></div>
        <div class="col-6 col-md-4"><div class="card stat-card p-4"><h6>من قبلي</h6>
            <h3 class="text-info">{money(since['self'])} د.ع</h3><small class="text-muted">بدون ربح هنا</small></div></div>
    </div>

    <div class="alert alert-info small d-flex flex-wrap justify-content-between align-items-center gap-2">
        <span>هذا المبلغ تتصفر تلقائياً لما تشحن المحفظة. رصيد المحفظة الحالي: <b>{money(balance)} د.ع</b></span>
        <a href="{url_for('wallet')}" class="btn btn-sm btn-primary">شحن المحفظة</a>
    </div>

    <div class="card p-4 mb-4">
        <h5><i class="bi bi-people text-primary"></i> اليوزرات الي استلمت منهم ({len(rows)})</h5>
        <div class="table-responsive">
            <table class="table table-hover">
                <thead><tr><th>الاسم</th><th>اليوزر</th><th>الباقة</th><th>المبلغ</th><th>الدفع</th><th>الربح</th><th>وقت الاستلام</th></tr></thead>
                <tbody>{table or '<tr><td colspan="7" class="text-muted text-center">ما استلمت من أحد من آخر شحن</td></tr>'}</tbody>
            </table>
        </div>
    </div>

    <div class="card p-4">
        <h5><i class="bi bi-clock-history text-secondary"></i> الشحنات السابقة</h5>
        <div class="table-responsive">
            <table class="table table-striped">
                <thead><tr><th>وقت الشحن</th><th>مبلغ الشحن</th><th>التحصيل قبل الشحن</th></tr></thead>
                <tbody>{charge_rows or '<tr><td colspan="3" class="text-muted text-center">ماكو شحنات</td></tr>'}</tbody>
            </table>
        </div>
    </div>
    """
    return page(body)


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

        before = collection_since_charge(con)[0]["collected"] if kind == "شحن" else 0
        if kind == "شحن":
            notes = (notes + " — " if notes else "") + f"التحصيل قبل الشحن: {money(before)}"
        con.execute("""
            INSERT INTO wallet_transactions (transaction_type, amount, notes, created_at)
            VALUES (?, ?, ?, ?)
        """, (kind, amount, notes or "خصم يدوي", now_str()))
        if kind == "شحن":
            set_setting(con, "collection_reset_at", precise_now())
        con.commit()
        con.close()

        if kind == "شحن":
            flash(f"تم شحن المحفظة بنجاح — وتصفّر تحصيل المبالغ (جان {money(before)} د.ع)", "success")
        else:
            flash("تم خصم المبلغ من المحفظة", "success")
        return redirect(url_for("wallet"))

    con = db()
    transactions = con.execute("SELECT * FROM wallet_transactions ORDER BY id DESC LIMIT 500").fetchall()
    balance = wallet_balance(con)
    since, last_charge = collection_since_charge(con)
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
        <div class="alert mt-2 mb-0" style="background:var(--primary-soft)">
            <i class="bi bi-cash-stack"></i> <a href="{url_for('collection')}" class="alert-link">تحصيل المبالغ</a> منذ آخر شحن ({esc((last_charge or 'البداية')[:16])}):
            <b>{money(since['collected'])} د.ع</b>
            <small class="text-muted">(نقد {money(since['cash'])} + من قبلي {money(since['self'])})</small>
            — يتصفر تلقائياً لما تشحن.
        </div>

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
        <option value="{plan['id']}" data-retail="{plan['retail_price']}" data-wholesale="{plan['wholesale_price']}" data-self="{self_price_of(plan)}" {sel}>
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
              data-confirm="سيتم حذف المشترك وكل عملياته نهائياً. متأكد؟">
            <button class="btn {cls} btn-outline-danger"><i class="bi bi-trash3"></i>{' حذف' if label else ''}</button>
        </form>"""


def account_badges(debt, self_paid):
    """شارات حساب المشترك بالحساب الرئيسي: دين، أو مسدد (والي دافع عنه "من قبلي" يبين جنبه)."""
    out = [f'<span class="badge bg-danger">دين {money(debt)}</span>' if debt
           else '<span class="badge bg-success-subtle text-success-emphasis">مسدد</span>']
    if self_paid:
        out.append('<span class="badge bg-info text-dark" title="محوّل لحساب الدفع من قبلي">من قبلي</span>')
    return " ".join(out)


def pay_choice_html(selected=PAY_CASH, prefix=""):
    """أزرار اختيار حالة الدفع الثلاثة."""
    opts = [
        (PAY_CASH, "cash", "bi-cash-coin", "نقداً", "استلمت المبلغ"),
        (PAY_DEBT, "debt", "bi-journal-text", "دين", "بذمة المشترك"),
        (PAY_SELF, "self", "bi-piggy-bank", "من قبلي", "سددته عنه"),
    ]
    html = '<div class="pay-options">'
    for value, cls, icon, title, sub in opts:
        pid = f"pay{prefix}{cls}"
        html += f"""
            <input type="radio" class="btn-check" name="payment_type" id="{pid}" value="{value}"
                   {'checked' if value == selected else ''} onchange="payCalc(this.form)">
            <label class="pay-opt {cls}" for="{pid}">
                <i class="bi {icon}"></i><b>{title}</b><small>{sub}</small>
            </label>"""
    return html + "</div>"


def price_field_html(price, prices):
    chips = "".join(
        f'<button type="button" class="btn btn-sm btn-outline-primary" '
        f'onclick="this.form.retail_price.value={p};payCalc(this.form)">{money(p)}</button>'
        for p in prices)
    return f"""
        <label class="form-label">سعر البيع للمشترك</label>
        <input type="number" name="retail_price" min="0" value="{price}" class="form-control mb-2"
               inputmode="numeric" oninput="payCalc(this.form)">
        <div class="d-flex flex-wrap gap-2 mb-3">{chips}</div>"""


PAY_CALC_JS = """
<script>
    function fmt(n) { return Number(n || 0).toLocaleString('en-US'); }
    function payCalc(f) {
        const sel = f.querySelector('select[name=plan_id]');
        let wholesale = +f.dataset.wholesale || 0;
        const o = sel && sel.selectedOptions[0];
        if (o && o.dataset.wholesale) {
            wholesale = +o.dataset.wholesale;
            if (f.dataset.lastPlan !== sel.value) {
                if (f.dataset.lastPlan !== undefined || !f.retail_price.value) f.retail_price.value = o.dataset.retail;
                f.dataset.lastPlan = sel.value;
            }
        }
        const t = (f.querySelector('input[name=payment_type]:checked') || {}).value;
        // "من قبلي" له سعر خاص (مثلاً 40,000)، والنقد/الدين بسعر البيع العادي
        const src = (o && o.dataset.wholesale) ? o.dataset : f.dataset;
        if (t && f.dataset.lastType !== undefined && f.dataset.lastType !== t && src.retail) {
            if (t === 'من قبلي') f.retail_price.value = src.self || src.retail;
            else if (f.dataset.lastType === 'من قبلي') f.retail_price.value = src.retail;
        }
        f.dataset.lastType = t || '';
        const price = +f.retail_price.value || 0;
        const out = f.querySelector('.pay-summary');
        const profit = price - wholesale;
        if (t === 'نقد') {
            out.innerHTML = '✅ يُضاف <b>' + fmt(price) + '</b> لتحصيل المبالغ، وربح <b class="' +
                (profit < 0 ? 'text-danger' : 'text-success') + '">' + fmt(profit) + ' د.ع</b> للأرباح';
        } else if (t === 'دين') {
            out.innerHTML = '📒 يُضاف <b class="text-danger">' + fmt(wholesale) +
                ' د.ع</b> (سعر الجملة) لحساب الديون — الربح ينحسب لما تستلم المبلغ (ترحيل)';
        } else {
            out.innerHTML = '🐷 من قبلي: مسدد بالحساب الرئيسي، ويُضاف <b>' + fmt(price) +
                ' د.ع</b> لتحصيل المبالغ — وينتقل لحساب يوزرات الدفع من قبلي (ربحه ينحسب هناك لما يرجّعلك المبلغ)';
        }
    }
    document.querySelectorAll('form.pay-form').forEach(payCalc);
</script>
"""


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

        cur = con.execute("""
            INSERT INTO subscribers
            (name, username, phone, plan_id, activation_date, expiry_date, notes, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (data["name"], data["username"], data["phone"], data["plan_id"],
              data["activation_date"], data["expiry_date"], data["notes"], now_str()))
        new_id = cur.lastrowid

        after = request.form.get("after", "save")
        msg = "تم حفظ بيانات المشترك (بدون أي خصم)"
        if after == "previous":
            ptype = request.form.get("payment_type", PAY_DEBT)
            retail = parse_amount(request.form.get("retail_price"))
            if record_previous(con, new_id, data["plan_id"], data["activation_date"], retail, ptype):
                msg = f"تم حفظ المشترك وتسجيل اشتراكه السابق كـ ({ptype}) بدون خصم من المحفظة"
        con.commit()
        con.close()

        if dup:
            flash("تنبيه: يوجد مشترك آخر بنفس اليوزر", "warning")
        if after == "activate":
            flash("تم حفظ المشترك — كمّل التفعيل (يخصم من المحفظة)", "success")
            return redirect(url_for("activate", subscriber_id=new_id))
        flash(msg, "success")
        return redirect(url_for("subscriber_detail", subscriber_id=new_id) if after == "previous" else url_for("subscribers"))

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
    elif status == "self":
        where.append(f"({SELF_SUBQUERY}) > 0")
    if plan_filter.isdigit():
        where.append("s.plan_id = ?")
        params.append(int(plan_filter))

    plans_list = con.execute("SELECT * FROM plans ORDER BY category, id").fetchall()
    subscriber_rows = con.execute(f"""
        SELECT s.*, p.name AS plan_name, p.retail_price, p.wholesale_price, p.profit,
               ({DEBT_SUBQUERY}) AS debt, ({DEBT_SALE_SUBQUERY}) AS debt_sale, ({SELF_SUBQUERY}) AS self_paid
        FROM subscribers s
        JOIN plans p ON p.id = s.plan_id
        {"WHERE " + " AND ".join(where) if where else ""}
        ORDER BY s.id DESC
    """, params).fetchall()
    shop = get_setting(con, "shop_name")
    quick_prices_list = quick_prices(con)
    con.close()

    subscriber_html = ""
    for s in subscriber_rows:
        debt = s["debt"] or 0
        pending = debt or s["self_paid"]
        wa_text = msg_debt(s, s["debt_sale"], shop) if debt > 0 else msg_reminder(s, shop)
        subscriber_html += f"""
        <tr>
            <td><input type="checkbox" name="ids" value="{s['id']}" form="bulkForm" class="form-check-input bulk-check" onchange="bulkCount()"></td>
            <td><a href="{url_for('subscriber_detail', subscriber_id=s['id'])}" class="fw-bold text-decoration-none">{esc(s['name'])}</a></td>
            <td>{esc(s['username'])}</td>
            <td dir="ltr">{esc(s['phone'])}</td>
            <td>{esc(s['plan_name'])}</td>
            <td>{status_badge(s['expiry_date'])}</td>
            <td>{account_badges(debt, s['self_paid'])}</td>
            <td>
                <div class="d-flex flex-wrap gap-1">
                    <a href="{url_for('activate', subscriber_id=s['id'])}" class="btn btn-sm btn-success">تفعيل</a>
                    <a href="{url_for('transfer', subscriber_id=s['id'])}" class="btn btn-sm {'btn-warning' if pending else 'btn-outline-warning'}">ترحيل</a>
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
        <a href="{url_for('wa_import')}" class="btn btn-outline-success w-100 mt-2 d-flex align-items-center justify-content-center gap-2">
            <i class="bi bi-whatsapp"></i> تسجيل تفعيلات من رسائل واتساب الشركة
        </a>
        <div class="collapse" id="addSub">
        <form method="post" class="row g-3 mt-2 pay-form">
            {subscriber_form_fields(plans_list)}
            <div class="col-12">
                <label class="form-label">بعد الحفظ</label>
                <div class="pay-options">
                    <input type="radio" class="btn-check" name="after" id="afSave" value="save" checked onchange="afterMode(this.form)">
                    <label class="pay-opt" for="afSave"><i class="bi bi-floppy" style="color:#64748b"></i><b>حفظ فقط</b><small>بيانات بدون أي خصم</small></label>
                    <input type="radio" class="btn-check" name="after" id="afPrev" value="previous" onchange="afterMode(this.form)">
                    <label class="pay-opt debt" for="afPrev"><i class="bi bi-clock-history"></i><b>اشتراك سابق</b><small>مفعّل من قبل — بدون خصم</small></label>
                    <input type="radio" class="btn-check" name="after" id="afAct" value="activate" onchange="afterMode(this.form)">
                    <label class="pay-opt" for="afAct"><i class="bi bi-lightning-charge"></i><b>تفعيل الآن</b><small>يخصم من المحفظة</small></label>
                </div>
            </div>
            <div class="col-12 prev-box" style="display:none">
                <div class="alert alert-warning small mb-3">يتسجل اشتراكه من <b>تاريخ التفعيل</b> المكتوب فوك، وما ينخصم شي من المحفظة.
                اختار حالة الدفع: إذا بعده ما دافع اختار <b>دين</b>.</div>
                {price_field_html('', quick_prices_list)}
                <div class="mb-3">{pay_choice_html(PAY_DEBT, prefix='add')}</div>
                <div class="alert pay-summary mb-0" style="background:var(--soft)"></div>
            </div>
            <div class="col-12">
                <button class="btn btn-primary px-4">حفظ المشترك</button>
            </div>
        </form>
        <script>
            function afterMode(f) {{
                f.querySelector('.prev-box').style.display = f.after.value === 'previous' ? '' : 'none';
                payCalc(f);
            }}
        </script>
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
                    <option value="self" {"selected" if status == "self" else ""}>محوّل للدفع من قبلي</option>
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
                        <th></th><th>الاسم</th><th>اليوزر</th><th>الهاتف</th><th>الاشتراك</th>
                        <th>الحالة</th><th>الحساب</th><th></th>
                    </tr>
                </thead>
                <tbody>{subscriber_html or '<tr><td colspan="8" class="text-muted text-center">لا يوجد مشتركون</td></tr>'}</tbody>
            </table>
        </div>
    </div>

    <form method="post" action="{url_for('bulk_previous')}" id="bulkForm" class="bulk-bar card p-3"
          data-confirm="راح يتسجل الاشتراك الحالي للمشتركين المحددين بدون خصم من المحفظة. متأكد؟">
        <div class="d-flex flex-wrap align-items-center gap-2">
            <div class="form-check me-2">
                <input class="form-check-input" type="checkbox" id="bulkAll"
                       onchange="document.querySelectorAll('.bulk-check').forEach(c => c.checked = this.checked); bulkCount()">
                <label class="form-check-label" for="bulkAll">الكل</label>
            </div>
            <b><span id="bulkN">0</span> محدد</b>
            <span class="text-muted small">سجّل اشتراكهم الحالي (بدون خصم) كـ</span>
            <select name="payment_type" class="form-select form-select-sm" style="width:auto">
                <option value="دين">دين</option>
                <option value="نقد">نقد</option>
                <option value="من قبلي">من قبلي</option>
            </select>
            <button class="btn btn-sm btn-warning"><i class="bi bi-check2-all"></i> تسجيل</button>
        </div>
    </form>
    <script>
        function bulkCount() {{
            const n = document.querySelectorAll('.bulk-check:checked').length;
            document.getElementById('bulkN').textContent = n;
            document.getElementById('bulkForm').classList.toggle('show', n > 0);
        }}
    </script>
    {PAY_CALC_JS}
    """
    return page(body)


@app.route("/subscribers/bulk-previous", methods=["POST"])
@login_required
def bulk_previous():
    ids = [int(i) for i in request.form.getlist("ids") if i.isdigit()]
    ptype = request.form.get("payment_type", PAY_DEBT)
    if not ids or ptype not in PAY_TYPES:
        flash("حدد مشتركين أولاً", "warning")
        return redirect(url_for("subscribers"))
    con = db()
    done = skipped = 0
    for sid in ids:
        s = con.execute("SELECT * FROM subscribers WHERE id = ?", (sid,)).fetchone()
        start = parse_date(s["activation_date"]) if s else None
        if s and start and record_previous(con, sid, s["plan_id"], start, None, ptype):
            done += 1
        else:
            skipped += 1
    con.commit()
    con.close()
    msg = f"تم تسجيل اشتراك {done} مشترك كـ ({ptype}) بدون خصم من المحفظة"
    if skipped:
        msg += f" — تخطي {skipped} (عندهم عملية مسجلة بنفس تاريخ التفعيل)"
    flash(msg, "success" if done else "warning")
    return redirect(request.referrer if (request.referrer or "").startswith(request.host_url) else url_for("subscribers"))


def get_subscriber(con, subscriber_id):
    return con.execute(f"""
        SELECT s.*, p.name AS plan_name, p.category,
               p.retail_price, p.wholesale_price, p.profit, p.self_price,
               ({DEBT_SUBQUERY}) AS debt, ({DEBT_SALE_SUBQUERY}) AS debt_sale, ({SELF_SUBQUERY}) AS self_paid
        FROM subscribers s
        JOIN plans p ON p.id = s.plan_id
        WHERE s.id = ?
    """, (subscriber_id,)).fetchone()


def self_status(h):
    if h["payment_type"] != PAY_SELF:
        return ""
    return ' <small class="text-info">بحساب الدفع من قبلي</small>'


def op_action(h, subscriber_id):
    """زر العملية بصفحة المشترك: ترحيل للدين، استلمت لـ"من قبلي"، تعديل للباقي."""
    if h["payment_type"] == PAY_DEBT:
        return f'<a href="{url_for("settle_operation", op_id=h["id"])}" class="btn btn-sm btn-warning">ترحيل</a>'
    return f'<a href="{url_for("settle_operation", op_id=h["id"])}" class="btn btn-sm btn-outline-secondary">تعديل</a>'


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
    shop = get_setting(con, "shop_name")
    prev_prices = quick_prices(con, s["retail_price"], s["self_price"])

    # رسالة تأكيد بعد عملية تفعيل مباشرة
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
                        <div class="text-muted">حالة الدفع: {payment_badge(op['payment_type'])} — سعر البيع {money(op['retail_price'])} د.ع</div>
                    </div>
                    {wa_button(s['phone'], msg_confirm(s, op, s['debt_sale'] or 0, shop), 'أرسل تأكيد للمشترك', 'btn-lg')}
                </div>
            </div>"""
    con.close()

    debt = s["debt"] or 0
    self_paid = s["self_paid"] or 0
    collected = sum(h["retail_price"] or 0 for h in history if h["payment_type"] == PAY_CASH)
    profit = sum(h["profit"] or 0 for h in history if h["payment_type"] == PAY_CASH)

    history_rows = "".join(f"""
        <tr>
            <td>{esc(h['operation_type'])}</td>
            <td>{esc(h['plan_name'])}</td>
            <td>{money(h['retail_price'])}</td>
            <td>{money(h['amount'])}</td>
            <td>{payment_badge(h['payment_type'])}{self_status(h)}</td>
            <td>{money(h['profit']) if h['payment_type'] == PAY_CASH or h['self_collected_at'] else '—'}</td>
            <td>{esc(h['operation_date'])}</td>
            <td>{op_action(h, subscriber_id)}</td>
        </tr>""" for h in history) or '<tr><td colspan="8" class="text-muted text-center">لا توجد عمليات لهذا المشترك</td></tr>'

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
        <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>الاشتراك</h6><div>{status_badge(s['expiry_date'])}</div>
            <small class="text-muted">ينتهي: {esc(s['expiry_date'])}</small></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>دين عليه</h6>
            <h4 class="fw-bold {'text-danger' if debt else 'text-success'}">{money(debt)} د.ع</h4></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>الدفع من قبلي</h6>
            <h4 class="fw-bold text-info">{'محوّل ✓' if self_paid else '—'}</h4>
            <small class="text-muted">{'بحساب الدفع من قبلي' if self_paid else 'لا'}</small></div></div>
        <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>المحصّل / الربح</h6>
            <h4 class="fw-bold">{money(collected)}</h4><small class="text-success">ربح {money(profit)} د.ع</small></div></div>
    </div>

    <div class="card p-4 mb-4">
        <h5><i class="bi bi-clock-history text-primary"></i> العمليات وحالة الدفع</h5>
        <p class="text-muted small">لتغيير حالة الدفع (استلام المبلغ، أو من قبلي) اضغط <b>ترحيل</b> جنب العملية — الترحيل ما يخصم من المحفظة.</p>
        <div class="table-responsive">
            <table class="table">
                <thead><tr><th>العملية</th><th>الاشتراك</th><th>سعر البيع</th><th>سعر الجملة</th>
                <th>الدفع</th><th>الربح</th><th>التاريخ</th><th></th></tr></thead>
                <tbody>{history_rows}</tbody>
            </table>
        </div>
    </div>

    <div class="card p-4 mb-4">
        <h5><i class="bi bi-clock-history text-warning"></i> تسجيل اشتراك سابق (بدون خصم من المحفظة)</h5>
        <p class="text-muted small">للمشترك الي مفعّل من قبل وما سجلته — مثلاً بعده ما دافع، سجّله <b>دين</b> حتى ينحسب بالديون.</p>
        <form method="post" action="{url_for('add_previous', subscriber_id=subscriber_id)}" class="row g-3 pay-form">
            <div class="col-md-6">
                <label class="form-label">الباقة</label>
                <select name="plan_id" class="form-select" onchange="payCalc(this.form)">{plan_options(plans_list, s['plan_id'])}</select>
            </div>
            <div class="col-md-6">
                <label class="form-label">تاريخ التفعيل</label>
                <input type="date" name="start" value="{esc(s['activation_date'])}" class="form-control" required>
            </div>
            <div class="col-12">{price_field_html(s['retail_price'], prev_prices)}
                <div class="mb-3">{pay_choice_html(PAY_DEBT, prefix='prev')}</div>
                <div class="alert pay-summary mb-3" style="background:var(--soft)"></div>
                <button class="btn btn-warning"><i class="bi bi-check2"></i> تسجيل الاشتراك السابق</button>
            </div>
        </form>
    </div>

    <div class="card p-4 mb-4">
        <h5><i class="bi bi-whatsapp" style="color:#25D366"></i> مراسلة المشترك على واتساب</h5>
        {"<p class='text-muted mb-0'>أضف رقم هاتف للمشترك حتى تكدر تراسله.</p>" if not wa_phone(s['phone']) else f'''
        <div class="d-flex flex-wrap gap-2 mb-3">
            <button type="button" class="btn btn-outline-secondary btn-sm" onclick="setMsg(this)" data-msg="{esc(msg_reminder(s, shop))}">تذكير بانتهاء الاشتراك</button>
            <button type="button" class="btn btn-outline-secondary btn-sm" onclick="setMsg(this)" data-msg="{esc(msg_debt(s, s['debt_sale'], shop))}">مطالبة بالدين</button>
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

    <div class="card p-4 mb-4" id="edit">
        <h5><i class="bi bi-pencil-square text-primary"></i> تعديل البيانات</h5>
        <form method="post" class="row g-3 mt-1">
            {subscriber_form_fields(plans_list, s, with_expiry=True)}
            <div class="col-12">
                <button class="btn btn-primary"><i class="bi bi-check2"></i> حفظ التعديلات</button>
            </div>
        </form>
        <hr>
        <div class="d-flex flex-wrap align-items-center justify-content-between gap-2">
            <small class="text-muted">حذف المشترك يحذف كل عملياته من التقارير نهائياً (حركات المحفظة تبقى).</small>
            {delete_form(subscriber_id, cls='')}
        </div>
    </div>
    {PAY_CALC_JS}
    """
    return page(body)


@app.route("/subscribers/<int:subscriber_id>/previous", methods=["POST"])
@login_required
def add_previous(subscriber_id):
    start = parse_date(request.form.get("start"))
    retail = parse_amount(request.form.get("retail_price"))
    ptype = request.form.get("payment_type", PAY_DEBT)
    try:
        plan_id = int(request.form.get("plan_id", ""))
    except ValueError:
        plan_id = None
    back = url_for("subscriber_detail", subscriber_id=subscriber_id)
    if not start or plan_id is None or retail is None or retail < 0:
        flash("أدخل التاريخ والسعر بشكل صحيح", "danger")
        return redirect(back)
    con = db()
    if not con.execute("SELECT 1 FROM subscribers WHERE id = ?", (subscriber_id,)).fetchone():
        con.close()
        flash("المشترك غير موجود", "danger")
        return redirect(url_for("subscribers"))
    ok = record_previous(con, subscriber_id, plan_id, start, retail, ptype)
    con.commit()
    con.close()
    if ok:
        flash(f"تم تسجيل الاشتراك السابق ({start}) كـ ({ptype}) بدون خصم من المحفظة", "success")
    else:
        flash("عنده عملية مسجلة بنفس التاريخ أو البيانات غير صحيحة — ما تسجل شي", "warning")
    return redirect(back)


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
    con.execute("DELETE FROM self_wallet WHERE activation_id IN (SELECT id FROM activations WHERE subscriber_id = ?)",
                (subscriber_id,))
    con.execute("DELETE FROM self_wallet WHERE old_id IN (SELECT id FROM self_old WHERE subscriber_id = ?)", (subscriber_id,))
    con.execute("DELETE FROM self_old WHERE subscriber_id = ?", (subscriber_id,))
    con.execute("DELETE FROM activations WHERE subscriber_id = ?", (subscriber_id,))
    con.execute("DELETE FROM subscribers WHERE id = ?", (subscriber_id,))
    con.commit()
    con.close()
    flash(f"تم حذف المشترك {s['username']}", "success")
    return redirect(url_for("subscribers"))


# ---------------------------------------------------------------------------
# التفعيل: يضيف شهر ويخصم الجملة من المحفظة
# ---------------------------------------------------------------------------

def perform_activation(subscriber_id):
    con = db()
    # قفل قاعدة البيانات أثناء العملية حتى لا يُخصم الرصيد مرتين بنفس الوقت
    con.execute("BEGIN IMMEDIATE")
    back = url_for("activate", subscriber_id=subscriber_id)

    subscriber = get_subscriber(con, subscriber_id)
    if not subscriber:
        con.rollback()
        con.close()
        flash("المشترك غير موجود", "danger")
        return redirect(url_for("subscribers"))

    try:
        plan_id = int(request.form.get("plan_id") or subscriber["plan_id"])
    except ValueError:
        plan_id = subscriber["plan_id"]
    plan = con.execute("SELECT * FROM plans WHERE id = ?", (plan_id,)).fetchone()
    if not plan:
        con.rollback()
        con.close()
        flash("الباقة غير موجودة", "danger")
        return redirect(back)

    amount = plan["wholesale_price"]
    balance = wallet_balance(con)
    if balance < amount:
        con.rollback()
        con.close()
        flash(f"الرصيد غير كافٍ. المطلوب {money(amount)} د.ع، والرصيد الحالي {money(balance)} د.ع", "danger")
        return redirect(back)

    payment_type = request.form.get("payment_type", PAY_CASH)
    if payment_type not in PAY_TYPES:
        payment_type = PAY_CASH
    retail = parse_amount(request.form.get("retail_price"))
    if retail is None or retail < 0:
        con.rollback()
        con.close()
        flash("أدخل سعر بيع صحيح", "danger")
        return redirect(back)

    notes = form_text("notes")
    now = now_str()
    start = today_str()
    # إذا الاشتراك بعده فعّال، الشهر الجديد يبدأ من تاريخ انتهاء الحالي
    current_exp = parse_date(subscriber["expiry_date"])
    if current_exp and current_exp > start:
        start = current_exp
    new_expiry = expiry_from(start)

    con.execute("""
        INSERT INTO wallet_transactions (transaction_type, amount, notes, created_at)
        VALUES (?, ?, ?, ?)
    """, ("خصم", amount, f"تفعيل - {subscriber['username']} - {plan['name']}"
                         + (f" - {notes}" if notes else ""), now))

    cur = con.execute("""
        INSERT INTO activations
        (subscriber_id, plan_id, operation_type, amount, retail_price, profit,
         payment_type, paid_amount, settled_at, notes, operation_date)
        VALUES (?, ?, 'تفعيل', ?, ?, ?, ?, ?, ?, ?, ?)
    """, (subscriber_id, plan_id, amount, retail, retail - amount, payment_type,
          retail if payment_type == PAY_CASH else 0,
          None if payment_type == PAY_DEBT else precise_now(), notes, now))
    activation_id = cur.lastrowid
    if payment_type == PAY_SELF:
        self_wallet_on_self(con, activation_id)

    con.execute("""
        UPDATE subscribers
        SET activation_date = ?, expiry_date = ?, plan_id = ?,
            notes = CASE WHEN ? <> '' THEN ? ELSE notes END
        WHERE id = ?
    """, (today_str(), new_expiry, plan_id, notes, notes, subscriber_id))

    con.commit()
    con.close()

    msg = (f"تم تفعيل {subscriber['username']} لغاية {new_expiry} وخصم {money(amount)} د.ع من المحفظة — ")
    msg += {PAY_CASH: f"نقداً: ربح {money(retail - amount)} د.ع",
            PAY_DEBT: f"دين: أُضيف {money(amount)} د.ع (سعر الجملة) لحساب الديون",
            PAY_SELF: f"من قبلي: مسدد وأُضيف {money(retail)} د.ع لتحصيل المبالغ، وانتقل لحساب يوزرات الدفع من قبلي"}[payment_type]
    flash(msg, "success")
    return redirect(url_for("subscriber_detail", subscriber_id=subscriber_id, done=activation_id))


@app.route("/activate/<int:subscriber_id>", methods=["GET", "POST"])
@login_required
def activate(subscriber_id):
    if request.method == "POST":
        return perform_activation(subscriber_id)

    con = db()
    subscriber = get_subscriber(con, subscriber_id)
    if not subscriber:
        con.close()
        flash("المشترك غير موجود", "danger")
        return redirect(url_for("subscribers"))
    plans_list = con.execute("SELECT * FROM plans ORDER BY category, id").fetchall()
    balance = wallet_balance(con)
    prices = quick_prices(con, subscriber["retail_price"], subscriber["self_price"])
    con.close()

    low = balance < subscriber["wholesale_price"]
    debt = subscriber["debt"] or 0

    body = f"""
    <h2 class="mb-4">تفعيل الاشتراك</h2>
    <div class="row g-4">
        <div class="col-lg-4 order-lg-2">
            <div class="card p-4">
                <h5><i class="bi bi-person-circle text-primary"></i> {esc(subscriber['name'])}</h5>
                <div class="d-flex justify-content-between py-1"><span class="text-muted">اليوزر</span><b>{esc(subscriber['username'])}</b></div>
                <div class="d-flex justify-content-between py-1"><span class="text-muted">الحالة</span>{status_badge(subscriber['expiry_date'])}</div>
                <div class="d-flex justify-content-between py-1"><span class="text-muted">ينتهي</span><span>{esc(subscriber['expiry_date'])}</span></div>
                <div class="d-flex justify-content-between py-1"><span class="text-muted">دين سابق</span>
                    <b class="{'text-danger' if debt else 'text-success'}">{money(debt)} د.ع</b></div>
                <div class="d-flex justify-content-between py-1"><span class="text-muted">رصيد المحفظة</span>
                    <b class="{'text-danger' if low else 'text-success'}">{money(balance)} د.ع</b></div>
            </div>
        </div>

        <div class="col-lg-8 order-lg-1">
            <div class="card p-4">
                <form method="post" class="pay-form">
                    <label class="form-label">الباقة</label>
                    <select name="plan_id" class="form-select mb-2" onchange="payCalc(this.form)">{plan_options(plans_list, subscriber['plan_id'])}</select>
                    <small class="text-muted d-block mb-3">يُضاف {SUBSCRIPTION_DAYS} يوم ويُخصم سعر الجملة من المحفظة.
                    إذا الاشتراك بعده فعّال، المدة تنضاف على تاريخ الانتهاء الحالي.</small>

                    {price_field_html(subscriber['retail_price'], prices)}

                    <label class="form-label">حالة الدفع</label>
                    <div class="mb-3">{pay_choice_html()}</div>

                    <div class="alert pay-summary mb-3" style="background:var(--soft)"></div>

                    <label class="form-label">ملاحظات</label>
                    <textarea name="notes" class="form-control mb-3" rows="2" placeholder="اكتب الملاحظات هنا"></textarea>

                    <div class="d-flex gap-2">
                        <button class="btn btn-success flex-fill py-3"><i class="bi bi-check2-circle"></i> تأكيد التفعيل</button>
                        <a href="{url_for('subscribers')}" class="btn btn-outline-secondary py-3">رجوع</a>
                    </div>
                </form>
            </div>
        </div>
    </div>
    {PAY_CALC_JS}
    """
    return page(body)


# ---------------------------------------------------------------------------
# الترحيل: تغيير حالة الدفع فقط (بدون خصم وبدون إضافة مدة)
# ---------------------------------------------------------------------------

def settle_card(op, prices, next_url):
    return f"""
    <div class="card p-4 mb-3">
        <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-3">
            <h5 class="mb-0">{esc(op['operation_type'])} {esc(op['plan_name'])}
                <small class="text-muted fs-6">— {esc(op['operation_date'][:10])}</small></h5>
            <div>الحالة الحالية: {payment_badge(op['payment_type'])}</div>
        </div>
        <form method="post" action="{url_for('settle_operation', op_id=op['id'])}" class="pay-form"
              data-wholesale="{op['amount']}" data-retail="{op['plan_retail']}" data-self="{op['plan_self'] or op['plan_retail']}">
            <input type="hidden" name="next" value="{esc(next_url)}">
            <div class="text-muted small mb-2">سعر الجملة: <b>{money(op['amount'])} د.ع</b></div>
            {price_field_html(op['retail_price'], prices)}
            <label class="form-label">حالة الدفع الجديدة</label>
            <div class="mb-3">{pay_choice_html(PAY_CASH, prefix=str(op['id']))}</div>
            <div class="alert pay-summary mb-3" style="background:var(--soft)"></div>
            <button class="btn btn-warning w-100 py-3"><i class="bi bi-arrow-left-right"></i> حفظ الترحيل</button>
        </form>
    </div>"""


@app.route("/transfer/<int:subscriber_id>")
@login_required
def transfer(subscriber_id):
    con = db()
    s = get_subscriber(con, subscriber_id)
    if not s:
        con.close()
        flash("المشترك غير موجود", "danger")
        return redirect(url_for("subscribers"))
    ops = con.execute("""
        SELECT a.*, p.name AS plan_name, p.retail_price AS plan_retail, p.self_price AS plan_self FROM activations a JOIN plans p ON p.id = a.plan_id
        WHERE a.subscriber_id = ? AND a.payment_type = 'دين' ORDER BY a.id
    """, (subscriber_id,)).fetchall()
    last = con.execute("""
        SELECT a.*, p.name AS plan_name, p.retail_price AS plan_retail, p.self_price AS plan_self FROM activations a JOIN plans p ON p.id = a.plan_id
        WHERE a.subscriber_id = ? ORDER BY a.id DESC LIMIT 1
    """, (subscriber_id,)).fetchone()
    prices = quick_prices(con, s["retail_price"], s["self_price"])
    con.close()

    here = url_for("transfer", subscriber_id=subscriber_id)
    if ops:
        cards = "".join(settle_card(op, prices, here) for op in ops)
        intro = f'<p class="text-muted">عنده <b>{len(ops)}</b> عملية دين. اختار الحالة الجديدة لكل وحدة — الترحيل ما يخصم من المحفظة وما يضيف مدة.</p>'
    elif last:
        cards = settle_card(last, prices, here)
        intro = '<div class="alert alert-success">ماكو دين على هذا المشترك ✅. إذا تريد تغيّر حالة آخر عملية تكدر من هنا.</div>'
    else:
        cards = ""
        intro = f"""<div class="alert alert-info">هذا المشترك ما عنده أي تفعيل بعد. لتجديد الاشتراك استخدم
            <a href="{url_for('activate', subscriber_id=subscriber_id)}" class="alert-link">تفعيل</a>.</div>"""

    body = f"""
    <h2 class="mb-4">ترحيل: {esc(s['name'])}</h2>
    <div class="card p-3 mb-3">
        <div class="d-flex flex-wrap justify-content-between gap-2">
            <span><b>{esc(s['name'])}</b> — <span dir="ltr">{esc(s['username'])}</span></span>
            <span>{account_badges(s['debt'], s['self_paid'])}</span>
        </div>
    </div>
    {intro}
    <div style="max-width:720px">{cards}</div>
    <a href="{url_for('subscriber_detail', subscriber_id=subscriber_id)}" class="btn btn-outline-secondary">رجوع للمشترك</a>
    {PAY_CALC_JS}
    """
    return page(body)


@app.route("/operations/<int:op_id>", methods=["GET", "POST"])
@login_required
def settle_operation(op_id):
    con = db()
    op = con.execute("""
        SELECT a.*, p.name AS plan_name, p.retail_price AS plan_retail, p.self_price AS plan_self, s.name AS sub_name, s.username
        FROM activations a JOIN plans p ON p.id = a.plan_id
        JOIN subscribers s ON s.id = a.subscriber_id WHERE a.id = ?
    """, (op_id,)).fetchone()
    if not op:
        con.close()
        flash("العملية غير موجودة", "danger")
        return redirect(url_for("subscribers"))

    if request.method == "POST":
        new_type = request.form.get("payment_type")
        retail = parse_amount(request.form.get("retail_price"))
        nxt = request.form.get("next") or ""
        if not nxt.startswith("/") or nxt.startswith("//"):
            nxt = url_for("subscriber_detail", subscriber_id=op["subscriber_id"])
        if new_type not in PAY_TYPES or retail is None or retail < 0:
            con.close()
            flash("اختار حالة الدفع وأدخل سعر صحيح", "danger")
            return redirect(nxt)
        if new_type == PAY_DEBT:
            settled = None
        elif new_type == op["payment_type"] and op["settled_at"]:
            settled = op["settled_at"]
        else:
            settled = precise_now()
        con.execute("""
            UPDATE activations SET payment_type = ?, retail_price = ?, profit = ?,
                   paid_amount = ?, settled_at = ?,
                   self_collected_at = CASE WHEN ? = 'من قبلي' AND payment_type = 'من قبلي' THEN self_collected_at ELSE NULL END
            WHERE id = ?
        """, (new_type, retail, retail - op["amount"],
              retail if new_type == PAY_CASH else 0, settled, new_type, op_id))
        if new_type == PAY_SELF and op["payment_type"] != PAY_SELF:
            self_wallet_on_self(con, op_id)
        elif new_type != PAY_SELF and op["payment_type"] == PAY_SELF:
            self_wallet_clear(con, op_id)
        con.commit()
        con.close()
        msg = {PAY_CASH: f"تم الترحيل: استلام {money(retail)} د.ع نقداً — ربح {money(retail - op['amount'])} د.ع",
               PAY_DEBT: f"تم الترحيل: {money(op['amount'])} د.ع (سعر الجملة) صارت دين على {op['sub_name']}",
               PAY_SELF: f"تم الترحيل: من قبلي {money(retail)} د.ع — مسدد بالحساب الرئيسي وانتقل لحساب يوزرات الدفع من قبلي"}[new_type]
        flash(msg, "success")
        return redirect(nxt)

    prices = quick_prices(con, op["retail_price"], op["plan_retail"], op["plan_self"])
    con.close()
    body = f"""
    <h2 class="mb-4">ترحيل / تعديل حالة الدفع</h2>
    <div class="card p-3 mb-3" style="max-width:720px">
        <b>{esc(op['sub_name'])}</b> <span dir="ltr" class="text-muted">{esc(op['username'])}</span>
    </div>
    <div style="max-width:720px">{settle_card(op, prices, url_for('subscriber_detail', subscriber_id=op['subscriber_id']))}</div>
    <a href="{url_for('subscriber_detail', subscriber_id=op['subscriber_id'])}" class="btn btn-outline-secondary">رجوع</a>
    {PAY_CALC_JS}
    """
    return page(body)


# ---------------------------------------------------------------------------
# الديون
# ---------------------------------------------------------------------------

@app.route("/debts")
@login_required
def debts():
    con = db()
    debt_ops = con.execute("""
        SELECT a.*, s.name, s.username, s.phone, s.expiry_date, p.name AS plan_name
        FROM activations a JOIN subscribers s ON s.id = a.subscriber_id JOIN plans p ON p.id = a.plan_id
        WHERE a.payment_type = 'دين'
        ORDER BY a.operation_date DESC
    """).fetchall()
    since, _ = collection_since_charge(con)
    shop = get_setting(con, "shop_name")
    con.close()

    total_debt = sum(o["amount"] or 0 for o in debt_ops)
    debt_by_sub, sale_by_sub = {}, {}
    for o in debt_ops:
        debt_by_sub[o["subscriber_id"]] = debt_by_sub.get(o["subscriber_id"], 0) + (o["amount"] or 0)
        sale_by_sub[o["subscriber_id"]] = sale_by_sub.get(o["subscriber_id"], 0) + (o["retail_price"] or 0)

    rows = "".join(f"""
        <tr>
            <td><a href="{url_for('subscriber_detail', subscriber_id=o['subscriber_id'])}" class="fw-bold text-decoration-none">{esc(o['name'])}</a></td>
            <td dir="ltr">{esc(o['username'])}</td>
            <td>{esc(o['plan_name'])}</td>
            <td class="fw-bold text-danger">{money(o['amount'])} د.ع</td>
            <td>{esc(o['operation_date'][:10])}</td>
            <td><div class="d-flex gap-1">
                <a href="{url_for('settle_operation', op_id=o['id'])}" class="btn btn-sm btn-warning"><i class="bi bi-arrow-left-right"></i> ترحيل</a>
                {wa_button(o['phone'], msg_debt(o, sale_by_sub.get(o['subscriber_id'], 0), shop), '')}
            </div></td>
        </tr>""" for o in debt_ops)

    body = f"""
    <h2 class="mb-4">الديون</h2>

    <div class="row g-3 mb-4">
        <div class="col-6 col-md-4">
            <div class="card stat-card p-4" style="background:linear-gradient(135deg,#dc2626,#f43f5e);color:#fff;border:0">
                <span class="stat-icon" style="background:rgba(255,255,255,.18);color:#fff"><i class="bi bi-journal-text"></i></span>
                <h6 style="color:rgba(255,255,255,.85)">مجموع الديون</h6>
                <h2>{money(total_debt)} د.ع</h2>
                <small style="color:rgba(255,255,255,.85)">{len(debt_by_sub)} مشترك</small>
            </div>
        </div>
        <div class="col-6 col-md-4">
            <a href="{url_for('collection')}" class="text-decoration-none">
            <div class="card stat-card p-4">
                <span class="stat-icon" style="background:rgba(16,185,129,.14);color:#10b981"><i class="bi bi-cash-stack"></i></span>
                <h6>تحصيل المبالغ منذ آخر شحن</h6>
                <h2 class="text-success">{money(since['collected'])} د.ع</h2>
            </div>
            </a>
        </div>
    </div>

    <div class="card p-4">
        <h5><i class="bi bi-journal-text text-danger"></i> الديون (بذمة المشتركين)</h5>
        <p class="text-muted small">لما تستلم المبلغ اضغط <b>ترحيل</b> واختار <b>نقداً</b> — ينضاف الربح للأرباح والمبلغ لتحصيل المبالغ.
        أو اختار <b>من قبلي</b> إذا دفعته انت عنه، فينتقل لحساب يوزرات الدفع من قبلي.</p>
        <div class="table-responsive">
            <table class="table table-hover">
                <thead><tr><th>الاسم</th><th>اليوزر</th><th>الباقة</th><th>الدين (الجملة)</th><th>تاريخ التفعيل</th><th></th></tr></thead>
                <tbody>{rows or '<tr><td colspan="6" class="text-muted text-center">ماكو ديون 👍</td></tr>'}</tbody>
            </table>
        </div>
    </div>
    """
    return page(body)


# ---------------------------------------------------------------------------
# حساب يوزرات الدفع من قبلي
# بالحساب الرئيسي: المشترك مسدد ومبلغه بتحصيل المبالغ (بدون ربح).
# هنا: المبلغ بذمته إلك، ولما يرجّعه تضغط "استلمت" وينحسب ربحه بهذا الحساب.
# ---------------------------------------------------------------------------

def self_account_totals(con, since=None):
    unpaid = con.execute("""
        SELECT COUNT(*) AS n, COALESCE(SUM(retail_price), 0) AS total
        FROM activations WHERE payment_type = 'من قبلي' AND self_collected_at IS NULL
    """).fetchone()
    old = con.execute("SELECT COUNT(*) AS n, COALESCE(SUM(amount), 0) AS total FROM self_old WHERE collected_at IS NULL").fetchone()
    subs = con.execute("""
        SELECT COUNT(*) AS c FROM (
            SELECT subscriber_id FROM activations WHERE payment_type = 'من قبلي' AND self_collected_at IS NULL
            UNION SELECT subscriber_id FROM self_old WHERE collected_at IS NULL)
    """).fetchone()["c"]
    cond, params = "", []
    if since:
        cond, params = " AND self_collected_at >= ?", [since]
    got = con.execute(f"""
        SELECT COUNT(*) AS n, COALESCE(SUM(retail_price), 0) AS total, COALESCE(SUM(profit), 0) AS profit
        FROM activations WHERE payment_type = 'من قبلي' AND self_collected_at IS NOT NULL {cond}
    """, params).fetchone()
    got_old = con.execute(f"""
        SELECT COALESCE(SUM(amount), 0) AS total FROM self_old
        WHERE collected_at IS NOT NULL {cond.replace('self_collected_at', 'collected_at')}
    """, params).fetchone()["total"]
    return {"unpaid": unpaid["total"] + old["total"], "unpaid_n": unpaid["n"], "unpaid_subs": subs,
            "old": old["total"], "old_n": old["n"],
            "got": got["total"] + got_old, "got_n": got["n"], "profit": got["profit"]}


@app.route("/self-account")
@login_required
def self_account():
    con = db()
    ops = con.execute("""
        SELECT a.*, s.name, s.username, s.phone, s.expiry_date, p.name AS plan_name
        FROM activations a JOIN subscribers s ON s.id = a.subscriber_id JOIN plans p ON p.id = a.plan_id
        WHERE a.payment_type = 'من قبلي'
        ORDER BY a.operation_date
    """).fetchall()
    month_start = date.today().replace(day=1).strftime("%Y-%m-%d")
    month = self_account_totals(con, since=month_start)
    total = self_account_totals(con)
    shop = get_setting(con, "shop_name")
    sw_balance = self_wallet_balance(con)
    sw_moves = con.execute("SELECT * FROM self_wallet ORDER BY id DESC LIMIT 100").fetchall()
    olds = con.execute("""
        SELECT o.*, s.name, s.username, s.phone, s.id AS subscriber_id FROM self_old o
        JOIN subscribers s ON s.id = o.subscriber_id ORDER BY o.id
    """).fetchall()
    all_subs = con.execute("SELECT id, name, username FROM subscribers ORDER BY name").fetchall()
    con.close()
    sub_options = "".join(f'<option value="{x["id"]}">{esc(x["name"])} — {esc(x["username"])}</option>' for x in all_subs)
    old_pending = sum(x["amount"] for x in olds if not x["collected_at"])
    old_got = sum(x["amount"] for x in olds if x["collected_at"])
    old_rows = "".join(f"""
        <tr>
            <td>{esc(x['name'])}</td>
            <td dir="ltr">{esc(x['username'])}</td>
            <td class="fw-bold">{money(x['amount'])} د.ع</td>
            <td>{esc(x['notes'])}</td>
            <td>{esc(x['created_at'][:10])}</td>
            <td>{'<span class="badge bg-success">مستلم ✓</span>' if x['collected_at'] else '<span class="badge bg-secondary">غير مستلم</span>'}</td>
        </tr>""" for x in reversed(olds))
    move_rows = "".join(f"""
        <tr>
            <td>{esc(m['kind'])}</td>
            <td class="fw-bold {'text-success' if m['kind'] in ('شحن', 'استلام') else 'text-danger'}">{'+' if m['kind'] in ('شحن', 'استلام') else '-'}{money(m['amount'])}</td>
            <td>{esc(m['notes'])}</td>
            <td>{esc(m['created_at'])}</td>
        </tr>""" for m in sw_moves)

    # تجميع حسب المشترك: كم اشتراك ما مسدد إلك
    subs = {}
    for o in ops:
        d = subs.setdefault(o["subscriber_id"], {"o": o, "unpaid": [], "old": [], "paid": 0, "profit": 0})
        if o["self_collected_at"]:
            d["paid"] += 1
            d["profit"] += o["profit"] or 0
        else:
            d["unpaid"].append(o)
    for x in olds:
        d = subs.setdefault(x["subscriber_id"], {"o": x, "unpaid": [], "old": [], "paid": 0, "profit": 0})
        if not x["collected_at"]:
            d["old"].append(x)
    owed = lambda d: sum(x["retail_price"] or 0 for x in d["unpaid"]) + sum(x["amount"] for x in d["old"])
    owing = sorted((d for d in subs.values() if d["unpaid"] or d["old"]), key=lambda d: -owed(d))
    settled = [d for d in subs.values() if not d["unpaid"] and not d["old"] and d["paid"]]

    def owing_row(d):
        o = d["o"]
        amount = owed(d)
        months = len(d["unpaid"])
        dates = "، ".join(x["operation_date"][:10] for x in d["unpaid"])
        old_html = "".join(f"""<div><span class="badge bg-secondary">مبلغ سابق {money(x['amount'])}</span>
            <small class="text-muted">{esc(x['notes'])}</small>
            <form method="post" action="{url_for('self_old_delete')}" class="d-inline" data-confirm="حذف هذا المبلغ السابق؟">
                <input type="hidden" name="id" value="{x['id']}">
                <button class="btn btn-link btn-sm p-0 text-danger" title="حذف"><i class="bi bi-x-circle"></i></button></form></div>"""
            for x in d["old"])
        return f"""
        <tr>
            <td class="fw-bold">{esc(o['name'])}</td>
            <td dir="ltr">{esc(o['username'])}</td>
            <td>{f'<span class="badge {"bg-danger" if months > 1 else "bg-warning text-dark"}">{months} اشتراك</span>' if months else ''}
                <div class="small text-muted">{esc(dates)}</div>{old_html}</td>
            <td class="fw-bold text-info">{money(amount)} د.ع</td>
            <td><div class="d-flex gap-1 flex-wrap">
                <form method="post" action="{url_for('self_collect')}" data-confirm="استلمت {money(amount)} د.ع من {esc(o['name'])}؟">
                    <input type="hidden" name="subscriber_id" value="{o['subscriber_id']}">
                    <button class="btn btn-sm btn-success"><i class="bi bi-check2-circle"></i> استلمت الكل</button>
                </form>
                {f'''<form method="post" action="{url_for('self_collect')}">
                    <input type="hidden" name="op_id" value="{d['unpaid'][0]['id']}">
                    <button class="btn btn-sm btn-outline-success">استلمت اشتراك واحد</button>
                </form>''' if months > 1 else ''}
                {wa_button(o['phone'], msg_debt(o, amount, shop), '')}
            </div></td>
        </tr>"""

    owing_rows = "".join(owing_row(d) for d in owing)
    settled_rows = "".join(f"""
        <tr>
            <td>{esc(d['o']['name'])}</td>
            <td dir="ltr">{esc(d['o']['username'])}</td>
            <td><span class="badge bg-success">مسدد ✓</span></td>
            <td>{d['paid']}</td>
            <td class="text-success">{money(d['profit'])} د.ع</td>
        </tr>""" for d in settled)

    body = f"""
    <h2 class="mb-4">يوزرات الدفع من قبلي</h2>

    <div class="row g-3 mb-4">
        <div class="col-6 col-md-4">
            <div class="card stat-card p-4" style="background:linear-gradient(135deg,#0284c7,#0ea5e9);color:#fff;border:0">
                <span class="stat-icon" style="background:rgba(255,255,255,.18);color:#fff"><i class="bi bi-piggy-bank"></i></span>
                <h6 style="color:rgba(255,255,255,.85)">بذمتهم إلك</h6>
                <h2>{money(total['unpaid'])} د.ع</h2>
                <small style="color:rgba(255,255,255,.85)">{total['unpaid_n']} اشتراك على {total['unpaid_subs']} مشترك</small>
            </div>
        </div>
        <div class="col-6 col-md-4">
            <div class="card stat-card p-4">
                <span class="stat-icon" style="background:rgba(16,185,129,.14);color:#10b981"><i class="bi bi-graph-up-arrow"></i></span>
                <h6>أرباح هذا الحساب (الشهر)</h6>
                <h2 class="text-success">{money(month['profit'])} د.ع</h2>
                <small class="text-muted">الإجمالي: {money(total['profit'])} د.ع</small>
            </div>
        </div>
        <div class="col-12 col-md-4">
            <div class="card stat-card p-4">
                <span class="stat-icon" style="background:rgba(79,70,229,.14);color:#4f46e5"><i class="bi bi-cash-stack"></i></span>
                <h6>استلمت منهم هذا الشهر</h6>
                <h2 class="text-primary">{money(month['got'])} د.ع</h2>
                <small class="text-muted">{month['got_n']} اشتراك</small>
            </div>
        </div>
    </div>

    <div class="card p-4 mb-4" style="border:2px solid #0ea5e9">
        <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-3">
            <h5 class="mb-0"><i class="bi bi-wallet2 text-info"></i> رصيد هذا الحساب</h5>
            <h3 class="mb-0 fw-bold {'text-danger' if sw_balance < 0 else 'text-info'}">{money(sw_balance)} د.ع</h3>
        </div>
        <p class="text-muted small">كل يوزر يتحول لـ"من قبلي" <b>ينخصم سعر جملته</b> من هذا الرصيد، ولما يرجّعلك فلوسه
        <b>يرجع المبلغ</b> للرصيد وينحسب ربحه. {'<b class="text-danger">الرصيد بالسالب — اشحن.</b>' if sw_balance < 0 else ''}</p>
        <form method="post" action="{url_for('self_wallet_move')}" class="row g-2">
            <div class="col-6 col-md-3"><select name="kind" class="form-select">
                <option value="شحن">شحن رصيد</option><option value="سحب">سحب</option></select></div>
            <div class="col-6 col-md-3"><input type="number" name="amount" min="1" class="form-control" placeholder="المبلغ" required></div>
            <div class="col-8 col-md-4"><input name="notes" class="form-control" placeholder="ملاحظات"></div>
            <div class="col-4 col-md-2"><button class="btn btn-info w-100 text-white">حفظ</button></div>
        </form>
    </div>

    <div class="card p-4 mb-4" style="border:2px solid #64748b">
        <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-2">
            <h5 class="mb-0"><i class="bi bi-clock-history text-secondary"></i> المطلوبين سابقاً</h5>
            <div><span class="badge bg-secondary fs-6">غير مستلم: {money(old_pending)} د.ع</span>
                 <span class="badge bg-success-subtle text-success-emphasis fs-6">مستلم: {money(old_got)} د.ع</span></div>
        </div>
        <p class="text-muted small">مبالغ قديمة دافعها من قبلك (قبل البرنامج). لما تضيفها <b>تنخصم من رصيد هذا الحساب</b>،
        ولما المشترك يرجّعها (<b>استلمت الكل</b>) ترجع للرصيد.</p>
        <form method="post" action="{url_for('self_old_add')}" class="row g-2 mb-3">
            <div class="col-md-5"><select name="subscriber_id" class="form-select" required>
                <option value="">اختار المشترك</option>{sub_options}</select></div>
            <div class="col-6 col-md-2"><input type="number" name="amount" min="1" class="form-control" placeholder="المبلغ" required></div>
            <div class="col-6 col-md-3"><input name="notes" class="form-control" placeholder="ملاحظة (مثلاً: دين قديم)"></div>
            <div class="col-md-2"><button class="btn btn-secondary w-100">إضافة</button></div>
        </form>
        <div class="table-responsive">
            <table class="table">
                <thead><tr><th>الاسم</th><th>اليوزر</th><th>المبلغ</th><th>الملاحظة</th><th>تاريخ الإضافة</th><th>الحالة</th></tr></thead>
                <tbody>{old_rows or '<tr><td colspan="6" class="text-muted text-center">ماكو مطلوبين سابقاً</td></tr>'}</tbody>
            </table>
        </div>
    </div>

    <div class="alert alert-info small">
        المشترك هنا <b>مسدد بالحساب الرئيسي</b> ومبلغه محسوب بتحصيل المبالغ، بس <b>انت دافع عنه</b>.
        كل شهر ما يدفعلك ينضاف عليه اشتراك. لما يدفعلك اضغط <b>استلمت</b> — وربحه ينحسب بهذا الحساب بس.
    </div>

    <div class="card p-4 mb-4">
        <h5><i class="bi bi-hourglass-split text-warning"></i> ما مسددين إلك ({len(owing)})</h5>
        <div class="table-responsive">
            <table class="table table-hover">
                <thead><tr><th>الاسم</th><th>اليوزر</th><th>الاشتراكات</th><th>المبلغ</th><th></th></tr></thead>
                <tbody>{owing_rows or '<tr><td colspan="5" class="text-muted text-center">الكل مسددين 👍</td></tr>'}</tbody>
            </table>
        </div>
    </div>

    <div class="card p-4">
        <h5><i class="bi bi-check2-all text-success"></i> مسددين ({len(settled)})</h5>
        <div class="table-responsive">
            <table class="table">
                <thead><tr><th>الاسم</th><th>اليوزر</th><th>الحالة</th><th>اشتراكات مسددة</th><th>الربح</th></tr></thead>
                <tbody>{settled_rows or '<tr><td colspan="5" class="text-muted text-center">لا يوجد</td></tr>'}</tbody>
            </table>
        </div>
    </div>

    <div class="card p-4 mt-4">
        <h5><i class="bi bi-list-ul text-info"></i> حركات رصيد الحساب</h5>
        <div class="table-responsive">
            <table class="table table-striped">
                <thead><tr><th>النوع</th><th>المبلغ</th><th>التفاصيل</th><th>التاريخ</th></tr></thead>
                <tbody>{move_rows or '<tr><td colspan="4" class="text-muted text-center">لا توجد حركات — اشحن رصيدك أول</td></tr>'}</tbody>
            </table>
        </div>
    </div>
    """
    return page(body)


@app.route("/self-account/wallet", methods=["POST"])
@login_required
def self_wallet_move():
    """شحن رصيد حساب الدفع من قبلي (أو سحب منه)."""
    amount = parse_amount(request.form.get("amount"))
    kind = SW_WITHDRAW if request.form.get("kind") == SW_WITHDRAW else SW_TOPUP
    if not amount or amount <= 0:
        flash("أدخل مبلغاً صحيحاً", "danger")
        return redirect(url_for("self_account"))
    con = db()
    con.execute("INSERT INTO self_wallet (kind, amount, notes, created_at) VALUES (?, ?, ?, ?)",
                (kind, amount, form_text("notes"), now_str()))
    con.commit()
    bal = self_wallet_balance(con)
    con.close()
    flash(f"تم {'شحن' if kind == SW_TOPUP else 'سحب'} {money(amount)} د.ع — رصيد حساب الدفع من قبلي: {money(bal)} د.ع", "success")
    return redirect(url_for("self_account"))


@app.route("/self-account/old", methods=["POST"])
@login_required
def self_old_add():
    """إضافة مبلغ قديم بذمة مشترك (بدون خصم من الرصيد)."""
    sid = request.form.get("subscriber_id", "")
    amount = parse_amount(request.form.get("amount"))
    back = request.form.get("back") or ""
    back = back if back.startswith("/") and not back.startswith("//") else url_for("self_account")
    con = db()
    s = con.execute("SELECT name FROM subscribers WHERE id = ?", (sid,)).fetchone() if sid.isdigit() else None
    if not s or not amount or amount <= 0:
        con.close()
        flash("اختار المشترك واكتب مبلغ صحيح", "danger")
        return redirect(back)
    cur = con.execute("INSERT INTO self_old (subscriber_id, amount, notes, created_at) VALUES (?, ?, ?, ?)",
                      (int(sid), amount, form_text("notes") or "مبلغ سابق", now_str()))
    # طالع من جيبك سابقاً: ينخصم من رصيد حساب الدفع من قبلي
    con.execute("INSERT INTO self_wallet (kind, amount, notes, old_id, created_at) VALUES ('خصم', ?, ?, ?, ?)",
                (amount, f"مبلغ سابق - {s['name']}", cur.lastrowid, now_str()))
    con.commit()
    bal = self_wallet_balance(con)
    con.close()
    flash(f"أُضيف {money(amount)} د.ع على حساب {s['name']} وانخصم من رصيد الحساب (الرصيد: {money(bal)} د.ع) ✅", "success")
    return redirect(back)


@app.route("/self-account/old/delete", methods=["POST"])
@login_required
def self_old_delete():
    oid = request.form.get("id", "")
    con = db()
    row = con.execute("SELECT subscriber_id FROM self_old WHERE id = ? AND collected_at IS NULL", (oid,)).fetchone() if oid.isdigit() else None
    if row:
        con.execute("DELETE FROM self_old WHERE id = ?", (oid,))
        con.execute("DELETE FROM self_wallet WHERE old_id = ?", (oid,))
        con.commit()
        flash("انحذف المبلغ السابق ورجع لرصيد الحساب", "success")
    con.close()
    back = request.form.get("back") or ""
    return redirect(back if back.startswith("/") and not back.startswith("//") else url_for("self_account"))


@app.route("/self-account/collect", methods=["POST"])
@login_required
def self_collect():
    """استلام المبلغ من مشترك "من قبلي": اشتراك واحد أو كل الي بذمته."""
    con = db()
    now = precise_now()
    base = """SELECT a.id, a.retail_price, s.username FROM activations a JOIN subscribers s ON s.id = a.subscriber_id
              WHERE a.payment_type = 'من قبلي' AND a.self_collected_at IS NULL"""
    if request.form.get("op_id", "").isdigit():
        ops = con.execute(base + " AND a.id = ?", (int(request.form["op_id"]),)).fetchall()
    elif request.form.get("subscriber_id", "").isdigit():
        ops = con.execute(base + " AND a.subscriber_id = ?", (int(request.form["subscriber_id"]),)).fetchall()
    else:
        ops = []
    for o in ops:
        con.execute("UPDATE activations SET self_collected_at = ? WHERE id = ?", (now, o["id"]))
        # المبلغ الي رجّعه المشترك يرجع لمحفظة حساب الدفع من قبلي
        con.execute("INSERT INTO self_wallet (kind, amount, notes, activation_id, created_at) VALUES ('استلام', ?, ?, ?, ?)",
                    (o["retail_price"], f"استلام من {o['username']}", o["id"], now_str()))
    n = len(ops)
    if request.form.get("subscriber_id", "").isdigit():
        olds = con.execute("""
            SELECT o.id, o.amount, s.username FROM self_old o JOIN subscribers s ON s.id = o.subscriber_id
            WHERE o.subscriber_id = ? AND o.collected_at IS NULL
        """, (int(request.form["subscriber_id"]),)).fetchall()
        for x in olds:
            con.execute("UPDATE self_old SET collected_at = ? WHERE id = ?", (now, x["id"]))
            con.execute("INSERT INTO self_wallet (kind, amount, notes, old_id, created_at) VALUES ('استلام', ?, ?, ?, ?)",
                        (x["amount"], f"مبلغ سابق من {x['username']}", x["id"], now_str()))
        n += len(olds)
    con.commit()
    con.close()
    flash(f"تم تسجيل استلام {n} مبلغ ✅ — ورجع لرصيد حساب الدفع من قبلي" if n else "ما تغيّر شي",
          "success" if n else "warning")
    back = request.form.get("back") or ""
    return redirect(back if back.startswith("/") and not back.startswith("//") else url_for("self_account"))


@app.route("/self-account/undo", methods=["POST"])
@login_required
def self_uncollect():
    op_id = request.form.get("op_id", "")
    con = db()
    s = con.execute("SELECT subscriber_id FROM activations WHERE id = ?", (op_id,)).fetchone() if op_id.isdigit() else None
    if s:
        con.execute("UPDATE activations SET self_collected_at = NULL WHERE id = ? AND payment_type = 'من قبلي'", (op_id,))
        self_wallet_clear(con, int(op_id), ("استلام",))
        con.commit()
        flash("رجعت العملية غير مسددة", "success")
    con.close()
    return redirect(url_for("subscriber_detail", subscriber_id=s["subscriber_id"]) if s else url_for("self_account"))


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


def guess_plan(card_type, plans_list):
    """تخمين الباقة المطابقة لنوع الكارت بملف الشركة."""
    ct = " ".join(card_type.lower().split())
    for p in plans_list:
        if p["name"].lower().strip() == ct:
            return p["id"]
    for category, name, *_rest, aliases in PLANS:
        if ct in aliases:
            for p in plans_list:
                if p["name"] == name and p["category"] == category:
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
        prev_type = request.form.get("prev_type", "")
        recorded = 0
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
                sid = existing["id"]
            else:
                cur = con.execute("""
                    INSERT INTO subscribers
                    (name, username, phone, plan_id, activation_date, expiry_date, notes, created_at)
                    VALUES (?, ?, '', ?, ?, ?, ?, ?)
                """, (name, r["username"], plan_id, start, expiry, "", now))
                added += 1
                sid = cur.lastrowid
            if prev_type in PAY_TYPES and record_previous(
                    con, sid, plan_id, start, None, prev_type, wholesale=r["amount"] or None):
                recorded += 1
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
        if recorded:
            msg += f" — وتسجيل {recorded} اشتراك كـ ({prev_type}) بدون خصم من المحفظة"
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
                    <div class="col-md-8">
                        <label class="form-label">تسجيل اشتراكاتهم بالحسابات (بدون خصم من المحفظة)</label>
                        <select name="prev_type" class="form-select">
                            <option value="دين" selected>سجّلها دين (ما استلمت منهم بعد)</option>
                            <option value="نقد">سجّلها نقد (استلمت منهم)</option>
                            <option value="من قبلي">سجّلها من قبلي</option>
                            <option value="">لا تسجّل — بس أضيف البيانات</option>
                        </select>
                    </div>
                    <div class="col-12">
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
            <li>الاستيراد ما يخصم من المحفظة. تكدر تسجّل اشتراكاتهم <b>دين</b> حتى تنحسب بالديون.</li>
        </ul>
    </div>
    """
    return page(body)


# ---------------------------------------------------------------------------
# تسجيل التفعيلات من رسائل واتساب الشركة
# ---------------------------------------------------------------------------

_AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
_INVISIBLE = re.compile(r"[‎‏‪-‮⁦-⁩﻿ ]")
# بداية رسالة بملف تصدير الدردشة: "14/11/2026, 10:23 - " أو "[14/11/2026, 10:23:45] " أو "[10:23, 14/11/2026] "
_MSG_START = re.compile(
    r"^\[?\s*(?:(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})[,،]?\s+(\d{1,2}:\d{2})"
    r"|(\d{1,2}:\d{2})(?::\d{2})?\s*(?:[AaPp][Mm]|[صم])?[,،]\s*(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4}))")


def _msg_date(d1, d2, y):
    """تحويل التاريخ (يوم/شهر أو شهر/يوم) إلى YYYY-MM-DD."""
    d1, d2, y = int(d1), int(d2), int(y)
    if y < 100:
        y += 2000
    day, month = (d2, d1) if d2 > 12 >= d1 else (d1, d2)
    try:
        return date(y, month, day).strftime("%Y-%m-%d")
    except ValueError:
        return None


def split_whatsapp_messages(text):
    """تقسيم نص الدردشة إلى رسائل، كل رسالة: (التاريخ، الوقت، النص)."""
    text = _INVISIBLE.sub("", text.translate(_AR_DIGITS)).replace("\r\n", "\n")
    messages, current = [], None
    for line in text.split("\n"):
        m = _MSG_START.match(line.strip())
        if m:
            if m.group(1):
                day, tm = _msg_date(m.group(1), m.group(2), m.group(3)), m.group(4)
            else:
                day, tm = _msg_date(m.group(6), m.group(7), m.group(8)), m.group(5)
            if day:
                current = [day, tm, line]
                messages.append(current)
                continue
        if current is not None:
            current[2] += "\n" + line
        elif line.strip():
            # نص ملصوق بدون تواريخ: كل فقرة رسالة بتاريخ اليوم
            current = [today_str(), "", line]
            messages.append(current)
    return messages


# رسالة الشركة: "تم تفعيل اشتراك انترنيت mu3f13p12 / نوع الاشتراك BASIC / المتبقي في حسابك 472000"
_CO_USER = re.compile(r"تفعيل\s+اشتراك[^\n]*?([A-Za-z0-9][A-Za-z0-9@._\-]{2,})")
_CO_TYPE = re.compile(r"نوع\s+الاشتراك\s*[:：]?\s*([A-Za-z0-9][A-Za-z0-9 +\-]*)")
_CO_LEFT = re.compile(r"المتبقي\s+في\s+حسابك\s*[:：]?\s*([\d,]+)")


def find_activations_in_chat(con, text):
    """يدور على التفعيلات داخل الرسائل.
    يرجع (موجودين، جدد، عدد الرسائل، رسائل بدون يوزر، آخر رصيد بحساب الشركة)."""
    subs = {}
    for s in con.execute("""
        SELECT s.id, s.name, s.username, s.plan_id, s.expiry_date, p.name AS plan_name,
               p.wholesale_price, p.retail_price
        FROM subscribers s JOIN plans p ON p.id = s.plan_id
    """):
        subs.setdefault(s["username"].strip().lower(), dict(s))
    plans_list = con.execute("SELECT * FROM plans").fetchall()
    plans_by_id = {p["id"]: p for p in plans_list}

    found, unknown, seen, unmatched = [], [], set(), 0
    company_left = None  # (التاريخ، الوقت، المبلغ)
    messages = []
    for day, tm, body in split_whatsapp_messages(text):
        # أكثر من رسالة ملصوقة ورا بعض بدون تواريخ: نقسمها عند "تم تفعيل"
        parts = [x for x in re.split(r"(?=تم\s+تفعيل)", body) if x.strip()] or [body]
        messages += [(day, tm, part) for part in parts]

    for day, tm, body in messages:
        card = _CO_TYPE.search(body)
        card_type = " ".join(card.group(1).split()) if card else ""
        left = _CO_LEFT.search(body)
        if left and (company_left is None or (day, tm) >= company_left[:2]):
            company_left = (day, tm, parse_amount(left.group(1)))
        plan_id = guess_plan(card_type, plans_list) if card_type else "new"

        co_user = _CO_USER.search(body)
        tokens = {t.lower().strip("._-") for t in re.findall(r"[A-Za-z0-9@._\-]{3,}", body)}
        if co_user:
            tokens.add(co_user.group(1).lower().strip("._-"))
        hits = [subs[t] for t in tokens if t in subs]

        if not hits and co_user:
            uname = co_user.group(1).strip("._-")
            key = (uname.lower(), day)
            if key not in seen:
                seen.add(key)
                plan = plans_by_id.get(plan_id) if plan_id != "new" else None
                unknown.append({"username": uname, "day": day, "time": tm, "card_type": card_type,
                                "plan_id": plan["id"] if plan else None,
                                "plan_name": plan["name"] if plan else "",
                                "wholesale_price": plan["wholesale_price"] if plan else 0,
                                "snippet": " ".join(body.split())[:140]})
            continue
        if not hits:
            unmatched += 1
            continue
        for s in hits:
            key = (s["id"], day)
            if key in seen:
                continue
            seen.add(key)
            row = dict(s)
            # إذا نوع الاشتراك بالرسالة يختلف عن باقة المشترك ناخذ الي بالرسالة
            if plan_id != "new" and plan_id != s["plan_id"]:
                p = plans_by_id[plan_id]
                row.update(plan_id=p["id"], plan_name=p["name"], wholesale_price=p["wholesale_price"],
                           retail_price=p["retail_price"], plan_changed=True)
            already = con.execute("""
                SELECT 1 FROM activations WHERE subscriber_id = ? AND substr(operation_date, 1, 10) = ?
            """, (s["id"], day)).fetchone()
            row.update(day=day, time=tm, already=bool(already), card_type=card_type,
                       snippet=" ".join(body.split())[:140])
            found.append(row)
    found.sort(key=lambda r: (r["day"], r["time"]))
    unknown.sort(key=lambda r: (r["day"], r["time"]))
    return found, unknown, len(messages), unmatched, company_left


def _read_chat_upload(file):
    """قراءة ملف تصدير الدردشة (txt أو zip من الآيفون)."""
    import zipfile
    data = file.read()
    if data[:2] == b"PK":
        with zipfile.ZipFile(BytesIO(data)) as z:
            names = [n for n in z.namelist() if n.lower().endswith(".txt")]
            if not names:
                return ""
            data = z.read(names[0])
    for enc in ("utf-8-sig", "utf-16", "cp1256"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", "ignore")


@app.route("/whatsapp-import", methods=["GET", "POST"])
@login_required
def wa_import():
    import json

    if request.method == "POST" and request.form.get("step") == "read":
        text = request.form.get("text") or ""
        file = request.files.get("file")
        if file and file.filename:
            text += "\n" + _read_chat_upload(file)
        if not text.strip():
            flash("الصق الرسائل أو ارفع ملف تصدير الدردشة", "danger")
            return redirect(url_for("wa_import"))
        con = db()
        found, unknown, total, unmatched, company_left = find_activations_in_chat(con, text)
        con.close()
        if not found and not unknown:
            flash(f"قريت {total} رسالة بس ما لگيت بيها أي رسالة تفعيل أو يوزر من مشتركينك.", "warning")
            return redirect(url_for("wa_import"))
        os.makedirs(IMPORT_DIR, exist_ok=True)
        _cleanup_imports()
        token = secrets.token_hex(16)
        with open(_import_path(token), "w", encoding="utf-8") as f:
            json.dump({"rows": found, "unknown": unknown, "total": total, "unmatched": unmatched,
                       "company_left": company_left}, f, ensure_ascii=False)
        return redirect(url_for("wa_import", t=token))

    if request.method == "POST" and request.form.get("step") == "confirm":
        path = _import_path(request.form.get("token"))
        if not path or not os.path.exists(path):
            flash("انتهت صلاحية العملية، اقرأ الرسائل مرة ثانية", "danger")
            return redirect(url_for("wa_import"))
        with open(path, encoding="utf-8") as f:
            payload = json.load(f)
        rows = list(payload["rows"])
        n_listed = len(rows)

        con = db()
        con.execute("BEGIN IMMEDIATE")
        done = skipped = no_balance = added = 0
        cost = 0
        # يوزرات جديدة: نضيفهم كمشتركين أول
        for j, u in enumerate(payload.get("unknown", [])):
            if request.form.get(f"new_{j}") != "1" or not u["plan_id"]:
                continue
            if con.execute("SELECT 1 FROM subscribers WHERE lower(username) = lower(?)", (u["username"],)).fetchone():
                continue
            cur = con.execute("""
                INSERT INTO subscribers (name, username, phone, plan_id, activation_date, expiry_date, notes, created_at)
                VALUES (?, ?, '', ?, ?, ?, '', ?)
            """, (u["username"], u["username"], u["plan_id"], u["day"], u["day"], now_str()))
            added += 1
            rows.append({**u, "id": cur.lastrowid})
        for i, r in enumerate(rows):
            if i < n_listed and request.form.get(f"use_{i}") != "1":
                continue
            s = con.execute("SELECT * FROM subscribers WHERE id = ?", (r["id"],)).fetchone()
            plan = con.execute("SELECT * FROM plans WHERE id = ?", (r.get("plan_id") or s["plan_id"],)).fetchone() if s else None
            if not s or not plan:
                skipped += 1
                continue
            # نفس اليوم مسجل من قبل؟
            if con.execute("SELECT 1 FROM activations WHERE subscriber_id = ? AND substr(operation_date, 1, 10) = ?",
                           (s["id"], r["day"])).fetchone():
                skipped += 1
                continue
            amount = plan["wholesale_price"]
            if wallet_balance(con) < amount:
                no_balance += 1
                continue
            when = f"{r['day']} {(r['time'] or '00:00')[:5]}:00"
            start = r["day"]
            current_exp = parse_date(s["expiry_date"])
            if current_exp and current_exp > start:
                start = current_exp
            new_expiry = expiry_from(start)
            con.execute("""
                INSERT INTO wallet_transactions (transaction_type, amount, notes, created_at)
                VALUES ('خصم', ?, ?, ?)
            """, (amount, f"تفعيل (واتساب الشركة) - {s['username']} - {plan['name']}", when))
            con.execute("""
                INSERT INTO activations
                (subscriber_id, plan_id, operation_type, amount, retail_price, profit,
                 payment_type, paid_amount, settled_at, notes, operation_date)
                VALUES (?, ?, 'تفعيل', ?, ?, ?, 'دين', 0, NULL, 'من رسائل واتساب الشركة', ?)
            """, (s["id"], plan["id"], amount, plan["retail_price"], plan["retail_price"] - amount, when))
            con.execute("UPDATE subscribers SET activation_date = ?, expiry_date = ?, plan_id = ? WHERE id = ?",
                        (r["day"], new_expiry, plan["id"], s["id"]))
            done += 1
            cost += amount
        con.commit()
        con.close()
        try:
            os.remove(path)
        except OSError:
            pass
        msg = f"تم تسجيل {done} تفعيل كـ دين — أُضيف شهر لكل واحد وخُصم {money(cost)} د.ع من المحفظة"
        if added:
            msg += f" — وأُضيف {added} مشترك جديد (عدّل أسماءهم وأرقامهم بعدين)"
        if skipped:
            msg += f" — تخطي {skipped} (مسجل من قبل بنفس اليوم)"
        flash(msg, "success" if done else "warning")
        if no_balance:
            flash(f"{no_balance} تفعيل ما تسجل لأن رصيد المحفظة ما يكفي — اشحن وأعد المحاولة", "danger")
        return redirect(url_for("debts"))

    # --- المعاينة ---
    token = request.args.get("t")
    path = _import_path(token) if token else None
    if path and os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            payload = json.load(f)
        rows = payload["rows"]
        con = db()
        balance = wallet_balance(con)
        con.close()
        new_rows = [r for r in rows if not r["already"]]
        unknown = payload.get("unknown", [])
        need = sum(r["wholesale_price"] for r in new_rows)
        left = payload.get("company_left")
        company_html = ""
        if left:
            after = balance - need
            diff = after - (left[2] or 0)
            unknown_cost = sum(u["wholesale_price"] for u in unknown)
            if diff == 0:
                note = "✅ متطابق"
            elif unknown and diff == unknown_cost:
                note = "✅ يتطابق إذا أضفت اليوزرات الجديدة الي تحت"
            else:
                note = f"— الفرق: <b>{money(diff)}</b> (تأكد من الشحنات أو التفعيلات الناقصة)"
            company_html = f"""
            <div class="alert {'alert-success' if note.startswith('✅') else 'alert-warning'}">
                <i class="bi bi-bank"></i> آخر رصيد بحساب الشركة (رسالة {esc(left[0])} {esc(left[1])}): <b>{money(left[2])} د.ع</b>
                — رصيد البرنامج بعد التسجيل: <b>{money(after)} د.ع</b> {note}
            </div>"""
        unknown_html = ""
        if unknown:
            urows = "".join(f"""
                <tr>
                    <td><input type="checkbox" name="new_{j}" value="1" class="form-check-input" {'disabled' if not u['plan_id'] else ''}></td>
                    <td>{esc(u['day'])} <small class="text-muted">{esc(u['time'])}</small></td>
                    <td dir="ltr"><b>{esc(u['username'])}</b></td>
                    <td dir="ltr">{esc(u['card_type'])}</td>
                    <td>{esc(u['plan_name']) or '<span class="text-danger">باقة غير معروفة</span>'}</td>
                    <td>{money(u['wholesale_price'])}</td>
                </tr>""" for j, u in enumerate(unknown))
            unknown_html = f"""
            <div class="card p-4 mb-4 border-warning" style="border-width:2px">
                <h5><i class="bi bi-person-plus text-warning"></i> يوزرات مو موجودة بالبرنامج ({len(unknown)})</h5>
                <p class="text-muted small">إذا تريد، أشّر عليها حتى تنضاف كمشتركين جدد (الاسم = اليوزر) وينسجل تفعيلها دين مثل الباقي.</p>
                <div class="table-responsive">
                    <table class="table">
                        <thead><tr><th></th><th>التاريخ</th><th>اليوزر</th><th>نوع الاشتراك</th><th>الباقة</th><th>الجملة</th></tr></thead>
                        <tbody>{urows}</tbody>
                    </table>
                </div>
            </div>"""
        table = "".join(f"""
            <tr class="{'opacity-50' if r['already'] else ''}">
                <td><input type="checkbox" name="use_{i}" value="1" class="form-check-input row-check" {'' if r['already'] else 'checked'}></td>
                <td>{esc(r['day'])} <small class="text-muted">{esc(r['time'])}</small></td>
                <td dir="ltr"><b>{esc(r['username'])}</b></td>
                <td>{esc(r['name'])}</td>
                <td>{esc(r['plan_name'])}{' <span class="badge bg-info text-dark">تغيّرت</span>' if r.get('plan_changed') else ''}</td>
                <td>{money(r['wholesale_price'])}</td>
                <td>{'<span class="badge bg-secondary">مسجل من قبل</span>' if r['already'] else '<span class="badge bg-success">جديد</span>'}</td>
                <td><small class="text-muted">{esc(r['snippet'])}</small></td>
            </tr>""" for i, r in enumerate(rows))
        body = f"""
        <h2 class="mb-4">تفعيلات من واتساب الشركة</h2>
        <form method="post">
            <input type="hidden" name="step" value="confirm">
            <input type="hidden" name="token" value="{esc(token)}">
            <div class="row g-3 mb-4">
                <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>رسائل مقروءة</h6><h3>{payload['total']}</h3>
                    <small class="text-muted">{payload['unmatched']} بدون يوزر معروف</small></div></div>
                <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>تفعيلات جديدة</h6><h3 class="text-success">{len(new_rows)}</h3>
                    <small class="text-muted">{len(rows) - len(new_rows)} مسجلة من قبل</small></div></div>
                <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>راح ينخصم</h6><h3 class="text-danger">{money(need)}</h3></div></div>
                <div class="col-6 col-md-3"><div class="card stat-card p-3"><h6>رصيد المحفظة</h6>
                    <h3 class="{'text-danger' if balance < need else 'text-success'}">{money(balance)}</h3></div></div>
            </div>
            <div class="alert alert-info">كل تفعيل محدد: <b>يضيف شهر</b> للمشترك، <b>يخصم سعر الجملة</b> من المحفظة،
            ويتسجل <b>دين</b>. لما تستلم المبلغ رحّله نقداً.</div>
            {company_html}
            <div class="card p-4 mb-4">
                <div class="form-check mb-2">
                    <input class="form-check-input" type="checkbox" id="allChk"
                           onchange="document.querySelectorAll('.row-check').forEach(c => c.checked = this.checked)">
                    <label class="form-check-label" for="allChk">تحديد الكل</label>
                </div>
                <div class="table-responsive">
                    <table class="table table-hover">
                        <thead><tr><th></th><th>التاريخ</th><th>اليوزر</th><th>الاسم</th><th>الباقة</th><th>الجملة</th><th>الحالة</th><th>الرسالة</th></tr></thead>
                        <tbody>{table or '<tr><td colspan="8" class="text-muted text-center">كل اليوزرات بالرسائل مو موجودة بالبرنامج — شوف القسم الجاي</td></tr>'}</tbody>
                    </table>
                </div>
            </div>
            {unknown_html}
            <div class="d-flex gap-2 import-actions">
                <button class="btn btn-success flex-fill py-3"><i class="bi bi-check2-all"></i> تسجيل التفعيلات المحددة</button>
                <a href="{url_for('wa_import')}" class="btn btn-outline-secondary py-3">إلغاء</a>
            </div>
        </form>
        """
        return page(body)

    body = f"""
    <h2 class="mb-4">تفعيلات من واتساب الشركة</h2>
    <div class="card p-4" style="max-width:720px">
        <h5><i class="bi bi-whatsapp" style="color:#25D366"></i> الصق رسائل التفعيل أو ارفع ملف الدردشة</h5>
        <p class="text-muted">البرنامج يدور بالرسائل على <b>يوزرات مشتركينك</b>. كل يوزر يلگاه:
        يضيف شهر، يخصم سعر الجملة من المحفظة، ويسجله <b>دين</b>. قبل الحفظ تشوف معاينة.</p>
        <form method="post" enctype="multipart/form-data">
            <input type="hidden" name="step" value="read">
            <label class="form-label">الصق الرسائل هنا</label>
            <textarea name="text" class="form-control mb-3" rows="7" dir="auto"
                      placeholder="انسخ رسائل التفعيل من محادثة الشركة والصقها هنا..."></textarea>
            <label class="form-label">أو ارفع ملف تصدير الدردشة (txt أو zip)</label>
            <input type="file" name="file" accept=".txt,.zip,text/plain,application/zip" class="form-control mb-3">
            <button class="btn btn-primary w-100 py-3"><i class="bi bi-search"></i> قراءة الرسائل</button>
        </form>
        <div class="text-muted small mt-3">
            <b>تصدير الدردشة من واتساب:</b> افتح محادثة الشركة ← ⋮ (أو اسم المحادثة بالآيفون) ←
            المزيد ← <b>تصدير الدردشة</b> ← <b>بدون وسائط</b> ← احفظ الملف وارفعه هنا.
            <br>التفعيل المسجل بنفس اليوم لنفس المشترك ما يتكرر، فتكدر ترفع نفس الدردشة كل يوم بدون خوف.
        </div>
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
    self_price = parse_amount(request.form.get("self_price"))
    if self_price is None or self_price <= 0:
        self_price = retail
    return (category, name, retail, retail - wholesale, wholesale, self_price), None


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
                    UPDATE plans SET category = ?, name = ?, retail_price = ?, profit = ?, wholesale_price = ?,
                           self_price = ?
                    WHERE id = ?
                """, (*values, plan_id))
                flash("تم تعديل الباقة (العمليات القديمة تبقى بأسعارها الأصلية)", "success")
            else:
                con.execute("""
                    INSERT INTO plans (category, name, retail_price, profit, wholesale_price, self_price)
                    VALUES (?, ?, ?, ?, ?, ?)
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
            <td><input form="{fid}" name="self_price" type="number" min="0" value="{self_price_of(p)}" class="form-control form-control-sm"></td>
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
            <div class="col-md-2"><label class="form-label">الفئة</label>
                <input name="category" class="form-control" placeholder="فايبر / وايرليس" required></div>
            <div class="col-md-2"><label class="form-label">اسم الباقة</label>
                <input name="name" class="form-control" required></div>
            <div class="col-md-2"><label class="form-label">سعر البيع</label>
                <input name="retail_price" type="number" min="0" class="form-control" required></div>
            <div class="col-md-2"><label class="form-label">يخصم من المحفظة</label>
                <input name="wholesale_price" type="number" min="0" class="form-control" required></div>
            <div class="col-md-2"><label class="form-label">سعر من قبلي</label>
                <input name="self_price" type="number" min="0" class="form-control" placeholder="= سعر البيع"></div>
            <div class="col-12"><button class="btn btn-success px-4">إضافة</button></div>
        </form>
        <small class="text-muted mt-2">الربح = سعر البيع − المخصوم من المحفظة (يُحسب تلقائياً). سعر من قبلي = السعر لما تسدد عن المشترك (بدون أرباح).</small>
    </div>

    <div class="card p-4">
        <div class="table-responsive">
            <table class="table">
                <thead><tr><th>الفئة</th><th>الباقة</th><th>سعر البيع</th><th>يخصم من المحفظة</th><th>سعر من قبلي</th>
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
               substr(a.settled_at, 1, 19) AS settled_at,
               a.amount,
               CASE WHEN a.payment_type = 'نقد' THEN a.profit ELSE 0 END AS profit,
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
        SELECT p.category, p.name, COUNT(*) AS ops, SUM(a.amount) AS cost,
               SUM(CASE WHEN a.payment_type = 'نقد' THEN a.retail_price ELSE 0 END) AS sales,
               SUM(CASE WHEN a.payment_type = 'نقد' THEN a.profit ELSE 0 END) AS profit
        FROM activations a JOIN plans p ON p.id = a.plan_id
        WHERE a.operation_date BETWEEN ? AND ?
        GROUP BY p.id ORDER BY profit DESC
    """, (start, end_full)).fetchall()
    totals = money_totals(con, since=start, until=end_full)
    con.close()

    total_cost = sum(a["amount"] or 0 for a in activations if a["operation_type"] != PREV_OP)
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
            <td>{money(a['profit']) if a['payment_type'] == PAY_CASH else '—'}</td>
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

    <div class="row g-3 mb-2">
        <div class="col-6 col-md-4"><div class="card stat-card p-3"><h6>التفعيلات</h6><h3>{len(activations)}</h3>
            <small class="text-muted">خصم من المحفظة: {money(total_cost)}</small></div></div>
        <div class="col-6 col-md-4"><div class="card stat-card p-3"><h6>مجموع مبالغ الاشتراكات</h6><h3 class="text-primary">{money(totals['collected'])}</h3>
            <small class="text-muted">نقد {money(totals['cash'])} + من قبلي {money(totals['self'])}</small></div></div>
        <div class="col-6 col-md-4"><div class="card stat-card p-3"><h6>الأرباح</h6><h3 class="text-success">{money(totals['profit'])}</h3>
            <small class="text-muted">من المحصّل نقداً</small></div></div>
        <div class="col-6 col-md-4"><div class="card stat-card p-3"><h6>مسددين من قبلي</h6><h3 class="text-info">{money(totals['self'])}</h3>
            <small class="text-muted">{totals['self_n']} عملية — بدون أرباح</small></div></div>
        <div class="col-12 col-md-4"><div class="card stat-card p-3"><h6>الديون الحالية</h6><h3 class="text-danger">{money(totals['debt'])}</h3>
            <small class="text-muted">على {totals['debt_n']} مشترك (كل الفترات)</small></div></div>
    </div>
    <p class="text-muted small mb-4">المحصّل والأرباح ومسددين من قبلي تنحسب حسب تاريخ الاستلام/الترحيل، والتفعيلات حسب تاريخ التفعيل.</p>

    <div class="card p-4 mb-4">
        <h5>حسب الباقة (الأرباح من النقد)</h5>
        <div class="table-responsive">
            <table class="table table-striped">
                <thead><tr><th>الباقة</th><th>التفعيلات</th><th>المحصّل</th><th>الجملة</th><th>الربح</th></tr></thead>
                <tbody>{plan_rows or '<tr><td colspan="5" class="text-muted text-center">لا توجد بيانات</td></tr>'}</tbody>
            </table>
        </div>
    </div>

    <div class="card p-4 mb-4">
        <h5>تقرير التفعيلات</h5>
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
               "سعر البيع", "حالة الدفع", "تاريخ الاستلام/الترحيل", "سعر الجملة (المخصوم)", "الربح المحصّل",
               "الملاحظات", "تاريخ التفعيل"]
    totals = ["المجموع", "", "", "", "", "",
              sum(r["retail_price"] or 0 for r in rows if r["payment_type"] == PAY_CASH), "نقد فقط", "",
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
            prices = [parse_amount(x) for x in re.split(r"[\s،,]+", form_text("quick_prices")) if x]
            set_setting(con, "quick_prices", ",".join(str(p) for p in prices if p and p > 0))
            con.commit()
            flash("تم حفظ الإعدادات", "success")

        elif action == "reset":
            if form_text("confirm_word") != "تصفير":
                flash("ما تصفّر شي — لازم تكتب كلمة (تصفير) بالضبط", "danger")
            else:
                # نسخة احتياطية تلقائية قبل التصفير
                backup_dir = os.path.join(os.path.dirname(DB_NAME) or BASE_DIR, "backups")
                os.makedirs(backup_dir, exist_ok=True)
                bpath = os.path.join(backup_dir, f"before_reset_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.db")
                dst = sqlite3.connect(bpath)
                con.backup(dst)
                dst.close()
                for table in ("debts", "activations", "wallet_transactions", "subscribers", "self_wallet", "self_old"):
                    con.execute(f"DELETE FROM {table}")
                if request.form.get("reset_plans") == "1":
                    con.execute("DELETE FROM plans")
                    seed_plans(con)
                con.execute("DELETE FROM settings WHERE key = 'collection_reset_at'")
                con.execute("DELETE FROM sqlite_sequence WHERE name IN "
                            "('debts', 'activations', 'wallet_transactions', 'subscribers', 'plans')")
                con.commit()
                flash("تم تصفير البرنامج: المشتركين والعمليات وحركات المحفظة صارت صفر"
                      + (" والباقات رجعت للأسعار الجديدة" if request.form.get("reset_plans") == "1" else "")
                      + f" — انحفظت نسخة احتياطية بالسيرفر: backups/{os.path.basename(bpath)}", "success")

        elif action == "tg_token":
            token = form_text("tg_token").replace(" ", "")
            if not token:
                for k in ("tg_token", "tg_chat", "tg_bot"):
                    con.execute("DELETE FROM settings WHERE key = ?", (k,))
                con.commit()
                flash("تم فصل التليكرام", "success")
            else:
                ok, res = telegram_api(token, "getMe")
                if ok:
                    set_setting(con, "tg_token", token)
                    set_setting(con, "tg_bot", res.get("username", ""))
                    con.commit()
                    flash(f"تم حفظ البوت @{res.get('username', '')} ✅ — هسه افتحه بالتليكرام واضغط Start، وبعدها اضغط (ربط المحادثة)", "success")
                else:
                    flash(f"التوكن غير صحيح: {res}", "danger")

        elif action == "tg_link":
            token = get_setting(con, "tg_token")
            ok, res = telegram_api(token, "getUpdates") if token else (False, "احفظ توكن البوت أول")
            chat = None
            if ok:
                for upd in reversed(res):
                    msg = upd.get("message") or upd.get("edited_message") or {}
                    if msg.get("chat", {}).get("id"):
                        chat = msg["chat"]
                        break
            if chat:
                set_setting(con, "tg_chat", str(chat["id"]))
                con.commit()
                telegram_api(token, "sendMessage", {"chat_id": chat["id"],
                             "text": "✅ تم ربط برنامج محاسبة الإنترنت. راح توصلك النسخ الاحتياطية هنا."})
                flash(f"تم ربط المحادثة ({chat.get('first_name') or chat.get('title') or chat['id']}) ✅ — وصلتك رسالة بالتليكرام", "success")
            elif ok:
                flash("ما لگيت رسالة — افتح البوت بالتليكرام واضغط Start (أو ابعثله أي رسالة) وبعدها اضغط ربط مرة ثانية", "warning")
            else:
                flash(f"ما گدرت أربط: {res}", "danger")

        elif action == "tg_send":
            con.close()
            result = run_backup()
            flash(result, "success" if "✅" in result else "warning")
            return redirect(url_for("settings"))

        elif action == "add_user":
            username = form_text("username")
            password = request.form.get("password", "")
            if not username or len(password) < 6:
                flash("أدخل اسم مستخدم وكلمة مرور (6 أحرف على الأقل)", "danger")
            elif con.execute("SELECT 1 FROM users WHERE username = ?", (username,)).fetchone():
                flash("اسم المستخدم موجود مسبقاً", "danger")
            else:
                role = ROLE_SELF if request.form.get("role") == ROLE_SELF else ROLE_ADMIN
                con.execute("INSERT INTO users(username, password, role) VALUES (?, ?, ?)",
                            (username, generate_password_hash(password), role))
                con.commit()
                flash("تمت إضافة المستخدم" + (" — يدخل على حساب الدفع من قبلي بس" if role == ROLE_SELF else ""), "success")

        elif action == "set_password":
            uid = parse_amount(request.form.get("user_id"))
            new = request.form.get("new_password", "")
            if len(new) < 6:
                flash("كلمة السر لازم تكون 6 أحرف أو أكثر", "danger")
            elif con.execute("UPDATE users SET password = ? WHERE id = ?",
                             (generate_password_hash(new), uid)).rowcount:
                con.commit()
                flash("تم تغيير كلمة سر المستخدم ✅", "success")

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

    users = con.execute("SELECT id, username, role FROM users ORDER BY id").fetchall()
    shop = get_setting(con, "shop_name")
    qp = get_setting(con, "quick_prices", "")
    tg_token = get_setting(con, "tg_token")
    tg_chat = get_setting(con, "tg_chat")
    tg_bot = get_setting(con, "tg_bot")
    last_backup = get_setting(con, "last_backup")
    con.close()
    job_cmd = f"python3 {os.path.join(BASE_DIR, 'backup_job.py')}"

    user_rows = "".join(f"""
        <tr>
            <td>{esc(u['username'])}</td>
            <td>{'<span class="badge bg-info text-dark">الدفع من قبلي</span>' if u['role'] == 'self' else '<span class="badge bg-secondary">البرنامج الرئيسي</span>'}</td>
            <td>{'<span class="badge bg-primary">أنت</span>' if u['id'] == session['user_id'] else f'''
                <div class="d-flex flex-wrap gap-1">
                    <form method="post" class="d-flex gap-1">
                        <input type="hidden" name="action" value="set_password">
                        <input type="hidden" name="user_id" value="{u['id']}">
                        <input type="password" name="new_password" minlength="6" class="form-control form-control-sm" style="max-width:130px" placeholder="كلمة سر جديدة" required>
                        <button class="btn btn-sm btn-outline-primary">تغيير</button>
                    </form>
                    <form method="post" data-confirm="حذف المستخدم؟">
                        <input type="hidden" name="action" value="delete_user">
                        <input type="hidden" name="user_id" value="{u['id']}">
                        <button class="btn btn-sm btn-outline-danger">حذف</button>
                    </form>
                </div>'''}</td>
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
                    <div class="col-6"><input name="username" class="form-control" placeholder="اسم المستخدم" required></div>
                    <div class="col-6"><input name="password" type="password" minlength="6" class="form-control" placeholder="كلمة المرور" required></div>
                    <div class="col-8"><select name="role" class="form-select">
                        <option value="admin">البرنامج الرئيسي</option>
                        <option value="self">حساب الدفع من قبلي بس</option></select></div>
                    <div class="col-4"><button class="btn btn-success w-100">إضافة</button></div>
                </form>
                <small class="text-muted d-block mt-2">يوزر "حساب الدفع من قبلي" لما يدخل يشوف بس يوزرات الدفع من قبلي ورصيدها.</small>
            </div>
        </div>

        <div class="col-12">
            <div class="card p-4">
                <h5><i class="bi bi-sliders text-primary"></i> إعدادات البيع والرسائل</h5>
                <form method="post" class="row g-3">
                    <input type="hidden" name="action" value="shop">
                    <div class="col-md-6"><label class="form-label">اسم المحل (توقيع رسائل واتساب)</label>
                        <input name="shop_name" value="{esc(shop)}" class="form-control" placeholder="مثال: مكتب الأمل للإنترنت - 0770xxxxxxx"></div>
                    <div class="col-md-4"><label class="form-label">أسعار بيع سريعة (أزرار بالتفعيل والترحيل)</label>
                        <input name="quick_prices" value="{esc(qp)}" class="form-control" dir="ltr" placeholder="مثال: 42000,50000"></div>
                    <div class="col-md-2 d-flex align-items-end"><button class="btn btn-primary w-100">حفظ</button></div>
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

        <div class="col-12">
            <div class="card p-4">
                <h5><i class="bi bi-telegram" style="color:#229ED9"></i> النسخ الاحتياطي التلقائي على التليكرام</h5>
                <div class="mb-3">
                    {'<span class="badge bg-success">مربوط ✅</span> البوت @' + esc(tg_bot) if tg_token and tg_chat else
                     ('<span class="badge bg-warning text-dark">البوت محفوظ — باقي ربط المحادثة</span>' if tg_token else
                      '<span class="badge bg-secondary">مو مربوط</span>')}
                    <div class="small text-muted mt-1">آخر نسخة: {esc(last_backup) or 'ماكو بعد'}</div>
                </div>

                <ol class="small text-muted mb-3">
                    <li>بالتليكرام ابحث عن <b dir="ltr">@BotFather</b> ← اضغط Start ← اكتب <b dir="ltr">/newbot</b> ← اختار اسم للبوت ← ينطيك <b>توكن</b> (رقم طويل).</li>
                    <li>الصق التوكن هنا واضغط حفظ.</li>
                    <li>افتح البوت الجديد بالتليكرام واضغط <b>Start</b>، وارجع اضغط <b>ربط المحادثة</b>.</li>
                    <li>اضغط <b>إرسال نسخة الآن</b> للتجربة.</li>
                </ol>

                <form method="post" class="row g-2 mb-2">
                    <input type="hidden" name="action" value="tg_token">
                    <div class="col-md-8"><input name="tg_token" class="form-control" dir="ltr" autocomplete="off"
                        placeholder="{'••••••••' + esc(tg_token[-4:]) + ' (محفوظ — اتركه فارغ واضغط حفظ للفصل)' if tg_token else '123456789:ABC...'}"></div>
                    <div class="col-md-4"><button class="btn btn-primary w-100">حفظ التوكن</button></div>
                </form>
                <div class="d-flex flex-wrap gap-2 mb-3">
                    <form method="post"><input type="hidden" name="action" value="tg_link">
                        <button class="btn btn-outline-primary" {'disabled' if not tg_token else ''}><i class="bi bi-link-45deg"></i> ربط المحادثة</button></form>
                    <form method="post"><input type="hidden" name="action" value="tg_send">
                        <button class="btn btn-success"><i class="bi bi-send"></i> إرسال نسخة الآن</button></form>
                </div>

                <div class="alert alert-success small mb-2">
                    <i class="bi bi-check-circle"></i> <b>النسخة اليومية تلقائية:</b> أول ما تفتح البرنامج كل يوم تتسوى نسخة
                    وتنبعث للتليكرام بالخلفية — ما تحتاج تسوي شي.
                </div>
                <div class="alert alert-light border small mb-0">
                    <b>اختياري (للحسابات المدفوعة بـ PythonAnywhere):</b> تكدر تخليها بوقت ثابت من صفحة <b>Tasks</b> ←
                    <b>Scheduled tasks</b> بهذا الأمر:
                    <div class="d-flex gap-2 mt-2">
                        <input class="form-control form-control-sm" dir="ltr" readonly value="{esc(job_cmd)}" id="jobCmd">
                        <button type="button" class="btn btn-sm btn-outline-secondary"
                                onclick="navigator.clipboard.writeText(document.getElementById('jobCmd').value);this.textContent='تم النسخ'">نسخ</button>
                    </div>
                    <div class="mt-2">الوقت بـ PythonAnywhere بتوقيت UTC (بغداد = UTC + 3)، يعني 20:00 هناك = 11 بالليل عندك.</div>
                </div>
            </div>
        </div>

        <div class="col-12">
            <div class="card p-4 border-danger" style="border-width:2px">
                <h5 class="text-danger"><i class="bi bi-exclamation-octagon"></i> تصفير البرنامج</h5>
                <p class="text-muted mb-3">يمسح <b>كل المشتركين والعمليات والديون وحركات المحفظة</b> حتى تبدي من جديد.
                المستخدمين والإعدادات يبقون. قبل التصفير تنحفظ نسخة احتياطية تلقائياً بالسيرفر — ومع ذلك نزّل نسخة لنفسك أول.</p>
                <form method="post" class="row g-2" data-confirm="آخر تأكيد: راح ينمسح كلشي. متأكد؟">
                    <input type="hidden" name="action" value="reset">
                    <div class="col-12">
                        <div class="form-check">
                            <input class="form-check-input" type="checkbox" name="reset_plans" value="1" id="rp" checked>
                            <label class="form-check-label" for="rp">ورجّع الباقات للأسعار الجديدة (Basic، Plus، Turbo، Pro Max، ايكو، بلص، استاندر، تيربو)</label>
                        </div>
                    </div>
                    <div class="col-md-6"><input name="confirm_word" class="form-control" placeholder="اكتب كلمة: تصفير" autocomplete="off" required></div>
                    <div class="col-md-6"><button class="btn btn-danger w-100"><i class="bi bi-trash3"></i> تصفير كل البيانات</button></div>
                </form>
            </div>
        </div>
    </div>
    """
    return page(body)


# ---------------------------------------------------------------------------
# النسخ الاحتياطي التلقائي + الإرسال للتليكرام
# ---------------------------------------------------------------------------

BACKUP_DIR = os.path.join(os.path.dirname(DB_NAME) or BASE_DIR, "backups")
BACKUP_KEEP_DAYS = 30


def make_backup_file(path):
    """نسخة كاملة وآمنة من قاعدة البيانات (تشتغل حتى والبرنامج مفتوح)."""
    src = db()
    dst = sqlite3.connect(path)
    src.backup(dst)
    dst.close()
    src.close()


def telegram_api(token, method, fields=None, file_path=None, file_name=None):
    """طلب لـ Telegram Bot API. يرجع (نجح؟، النتيجة أو رسالة الخطأ)."""
    import json
    import urllib.request
    import urllib.parse
    import urllib.error
    url = f"https://api.telegram.org/bot{token}/{method}"
    fields = {k: str(v) for k, v in (fields or {}).items()}
    if file_path:
        boundary = "----accounting" + secrets.token_hex(8)
        body = BytesIO()
        for k, v in fields.items():
            body.write(f'--{boundary}\r\nContent-Disposition: form-data; name="{k}"\r\n\r\n{v}\r\n'.encode())
        body.write(f'--{boundary}\r\nContent-Disposition: form-data; name="document"; '
                   f'filename="{file_name or os.path.basename(file_path)}"\r\n'
                   f'Content-Type: application/octet-stream\r\n\r\n'.encode())
        with open(file_path, "rb") as f:
            body.write(f.read())
        body.write(f"\r\n--{boundary}--\r\n".encode())
        req = urllib.request.Request(url, data=body.getvalue(),
                                     headers={"Content-Type": f"multipart/form-data; boundary={boundary}"})
    else:
        req = urllib.request.Request(url, data=urllib.parse.urlencode(fields).encode())
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        try:
            data = json.loads(e.read().decode())
        except Exception:
            return False, f"خطأ من تليكرام ({e.code})"
    except Exception as e:
        return False, f"ما گدر يتصل بتليكرام: {e}"
    if not data.get("ok"):
        return False, data.get("description", "خطأ غير معروف")
    return True, data.get("result")


def backup_summary(con):
    subs = con.execute("SELECT COUNT(*) AS c FROM subscribers").fetchone()["c"]
    totals = money_totals(con)
    return (f"💾 نسخة احتياطية — {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
            f"👥 المشتركين: {subs}\n"
            f"💰 رصيد المحفظة: {money(wallet_balance(con))} د.ع\n"
            f"📒 الديون: {money(totals['debt'])} د.ع على {totals['debt_n']} مشترك\n"
            f"📈 الأرباح الكلية: {money(totals['profit'])} د.ع")


def run_backup(send_telegram=True):
    """يسوي نسخة يومية بمجلد backups ويحذف القديمة، ويبعثها للتليكرام إذا مربوط.
    يرجع رسالة بالنتيجة."""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d")
    path = os.path.join(BACKUP_DIR, f"daily_{stamp}.db")
    make_backup_file(path)

    # الاحتفاظ بآخر 30 نسخة يومية فقط
    daily = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith("daily_") and f.endswith(".db"))
    for old in daily[:-BACKUP_KEEP_DAYS]:
        try:
            os.remove(os.path.join(BACKUP_DIR, old))
        except OSError:
            pass

    result = f"انحفظت النسخة: backups/{os.path.basename(path)}"
    con = db()
    token = get_setting(con, "tg_token")
    chat = get_setting(con, "tg_chat")
    if send_telegram and token and chat:
        caption = backup_summary(con)
        ok, res = telegram_api(token, "sendDocument", {"chat_id": chat, "caption": caption},
                               file_path=path, file_name=f"backup_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.db")
        result += " — وانبعثت للتليكرام ✅" if ok else f" — ⚠️ ما انبعثت للتليكرام: {res}"
    elif send_telegram:
        result += " — (التليكرام مو مربوط)"
    set_setting(con, "last_backup", f"{now_str()} — {result}")
    con.commit()
    con.close()
    return result


def _safe_backup():
    try:
        app.logger.info("Daily backup: %s", run_backup())
    except Exception:
        app.logger.exception("Daily backup failed")


@app.after_request
def auto_daily_backup(response):
    """نسخة احتياطية تلقائية مرة باليوم: أول ما يفتح المستخدم البرنامج.
    تشتغل بعد ما توصل الصفحة للمتصفح حتى ما تبطّئ."""
    if request.method == "GET" and session.get("user_id") and request.endpoint not in (None, "static", "backup"):
        try:
            con = db()
            today = today_str()
            if get_setting(con, "auto_backup_day") != today:
                set_setting(con, "auto_backup_day", today)
                con.commit()
                response.call_on_close(_safe_backup)
            con.close()
        except Exception:
            app.logger.exception("auto backup check failed")
    return response


@app.route("/backup")
@login_required
def backup():
    fd, path = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    try:
        make_backup_file(path)
        with open(path, "rb") as f:
            data = BytesIO(f.read())
    finally:
        os.remove(path)
    return send_file(
        data, as_attachment=True,
        download_name=f"backup_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.db",
        mimetype="application/octet-stream"
    )


# ---------------------------------------------------------------------------
# تطبيق الموبايل (PWA): يتثبت على الشاشة ويتحول APK عن طريق PWABuilder
# ---------------------------------------------------------------------------

@app.route("/manifest.webmanifest")
def manifest():
    import json
    data = {
        "id": "/",
        "name": "محاسبة الإنترنت",
        "short_name": "محاسبة الإنترنت",
        "description": "إدارة مشتركي الإنترنت: التفعيل والترحيل والديون والمحفظة والتقارير",
        "lang": "ar",
        "dir": "rtl",
        "start_url": "/dashboard",
        "scope": "/",
        "display": "standalone",
        "orientation": "portrait",
        "background_color": "#f4f6fb",
        "theme_color": "#4f46e5",
        "categories": ["business", "finance", "productivity"],
        "icons": [
            {"src": "/static/icons/icon-192.png", "sizes": "192x192", "type": "image/png", "purpose": "any"},
            {"src": "/static/icons/icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "any"},
            {"src": "/static/icons/maskable-192.png", "sizes": "192x192", "type": "image/png", "purpose": "maskable"},
            {"src": "/static/icons/maskable-512.png", "sizes": "512x512", "type": "image/png", "purpose": "maskable"},
        ],
        "shortcuts": [
            {"name": "المشتركون", "url": "/subscribers", "icons": [{"src": "/static/icons/icon-192.png", "sizes": "192x192"}]},
            {"name": "الديون", "url": "/debts", "icons": [{"src": "/static/icons/icon-192.png", "sizes": "192x192"}]},
            {"name": "تفعيلات واتساب", "url": "/whatsapp-import", "icons": [{"src": "/static/icons/icon-192.png", "sizes": "192x192"}]},
        ],
    }
    resp = app.response_class(json.dumps(data, ensure_ascii=False), mimetype="application/manifest+json")
    resp.headers["Cache-Control"] = "no-cache"
    return resp


SERVICE_WORKER_JS = """
// يحفظ صفحة "ماكو إنترنت" والأيقونات، والباقي دائماً من السيرفر (حتى البيانات تبقى محدثة)
const CACHE = 'acc-v1';
const OFFLINE = '/offline';
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll([OFFLINE, '/static/icons/icon-192.png'])));
  self.skipWaiting();
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  if (e.request.mode === 'navigate') {
    e.respondWith(fetch(e.request).catch(() => caches.match(OFFLINE)));
  }
});
"""


@app.route("/sw.js")
def service_worker():
    resp = app.response_class(SERVICE_WORKER_JS, mimetype="application/javascript")
    resp.headers["Cache-Control"] = "no-cache"
    resp.headers["Service-Worker-Allowed"] = "/"
    return resp


@app.route("/offline")
def offline():
    return """<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>ماكو إنترنت</title>
<style>body{margin:0;min-height:100vh;display:grid;place-items:center;font-family:Tahoma,sans-serif;
background:#f4f6fb;color:#0f172a;text-align:center;padding:20px}
@media (prefers-color-scheme:dark){body{background:#0b1120;color:#e2e8f0}}
img{width:96px;border-radius:22px}button{margin-top:16px;padding:12px 28px;border:0;border-radius:12px;
background:#4f46e5;color:#fff;font-size:16px}</style></head><body><div>
<img src="/static/icons/icon-192.png" alt=""><h2>ماكو اتصال بالإنترنت</h2>
<p>البرنامج يحتاج إنترنت حتى يجيب بياناتك. تأكد من الاتصال وحاول مرة ثانية.</p>
<button onclick="location.reload()">إعادة المحاولة</button></div></body></html>"""


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
