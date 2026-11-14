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
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="csrf-token" content="{{ csrf_token() }}">
    <title>برنامج محاسبة الإنترنت</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.rtl.min.css" rel="stylesheet">

    <style>
        body { background: #f1f5f9; font-family: Tahoma, Arial, sans-serif; }

        .sidebar {
            min-height: 100vh;
            background: linear-gradient(180deg, #172554, #1e3a8a);
            color: white;
            padding: 20px 12px;
        }
        .sidebar h4 {
            padding: 15px 8px;
            border-bottom: 1px solid rgba(255,255,255,.2);
            margin-bottom: 20px;
        }
        .sidebar a {
            display: block; color: white; text-decoration: none;
            padding: 12px; border-radius: 10px; margin-bottom: 7px;
        }
        .sidebar a:hover, .sidebar a.active { background: rgba(255,255,255,.18); }
        .sidebar .who { font-size: 13px; opacity: .75; padding: 0 12px 10px; }

        .content { padding: 25px; }
        .card { border: none; border-radius: 16px; box-shadow: 0 5px 18px rgba(15,23,42,.08); }
        .balance-card { background: linear-gradient(135deg, #0f766e, #14b8a6); color: white; }
        .stat-card { min-height: 125px; }
        .table { vertical-align: middle; }
        .btn { border-radius: 9px; }
        .form-control, .form-select { border-radius: 9px; padding: 10px; }
        .brand { font-weight: bold; font-size: 20px; }

        @media (max-width: 767px) {
            .sidebar { min-height: auto; }
            .sidebar a { display: inline-block; padding: 8px 10px; }
            .content { padding: 15px 10px; }
        }
        @media print {
            .sidebar, .no-print, form { display: none !important; }
            .col-md-10 { width: 100% !important; }
            .card { box-shadow: none; }
        }
    </style>
</head>

<body>
<div class="container-fluid">
    <div class="row">

        {% if session.get('user_id') %}
        <aside class="col-md-2 sidebar">
            <h4 class="brand">📡 محاسبة الإنترنت</h4>
            <div class="who">👤 {{ session.get('username') }}</div>

            <a href="{{ url_for('dashboard') }}" class="{{ 'active' if request.endpoint == 'dashboard' }}">🏠 الرئيسية</a>
            <a href="{{ url_for('subscribers') }}" class="{{ 'active' if request.endpoint == 'subscribers' }}">👥 المشتركون</a>
            <a href="{{ url_for('wallet') }}" class="{{ 'active' if request.endpoint == 'wallet' }}">💰 المحفظة</a>
            <a href="{{ url_for('plans') }}" class="{{ 'active' if request.endpoint == 'plans' }}">📦 الباقات والأسعار</a>
            <a href="{{ url_for('reports') }}" class="{{ 'active' if request.endpoint == 'reports' }}">📊 التقارير</a>
            <a href="{{ url_for('settings') }}" class="{{ 'active' if request.endpoint == 'settings' }}">⚙️ الإعدادات</a>
            <a href="{{ url_for('logout') }}">🚪 تسجيل الخروج</a>
        </aside>
        {% endif %}

        <main class="{% if session.get('user_id') %}col-md-10{% else %}col-md-12{% endif %} content">

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

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
<script>
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
</script>
</body>
</html>
"""


def page(body):
    return render_template_string(BASE_HTML, body=body)


LOGIN_PAGE = """
<div class="row justify-content-center mt-5">
    <div class="col-md-5">
        <div class="card p-4">
            <h3 class="text-center mb-4">📡 تسجيل الدخول</h3>

            <form method="post">
                <label class="form-label">اسم المستخدم</label>
                <input name="username" class="form-control mb-3" required autofocus>

                <label class="form-label">كلمة المرور</label>
                <input name="password" type="password" class="form-control mb-3" required>

                <button class="btn btn-primary w-100">دخول</button>
            </form>

            {% if show_default %}
            <div class="text-muted mt-3 text-center small">
                الدخول الافتراضي: admin / admin123 — غيّرها من الإعدادات بعد الدخول
            </div>
            {% endif %}
        </div>
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
    con.close()

    expiring_rows = "".join(f"""
        <tr>
            <td>{esc(s['name'])}</td>
            <td>{esc(s['username'])}</td>
            <td>{esc(s['phone'])}</td>
            <td>{esc(s['plan_name'])}</td>
            <td>{esc(s['expiry_date'])}</td>
            <td>{status_badge(s['expiry_date'])}</td>
            <td><a href="{url_for('transfer', subscriber_id=s['id'])}" class="btn btn-sm btn-warning">ترحيل</a></td>
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
        <div class="col-md-3">
            <div class="card balance-card stat-card p-4">
                <h6>رصيد المحفظة الحالي</h6>
                <h2>{money(balance)} د.ع</h2>
            </div>
        </div>
        <div class="col-md-3">
            <div class="card stat-card p-4">
                <h6>عدد المشتركين</h6>
                <h2>{subscribers_count}</h2>
            </div>
        </div>
        <div class="col-md-3">
            <div class="card stat-card p-4">
                <h6>أرباح هذا الشهر</h6>
                <h2 class="text-success">{money(month['profit'])} د.ع</h2>
                <small class="text-muted">{month['ops']} عملية — مبيعات {money(month['sales'])} د.ع</small>
            </div>
        </div>
        <div class="col-md-3">
            <div class="card stat-card p-4">
                <h6>إجمالي الأرباح</h6>
                <h2>{money(stats['total_profit'])} د.ع</h2>
                <small class="text-muted">تفعيل: {stats['activations'] or 0} — ترحيل: {stats['transfers'] or 0}</small>
            </div>
        </div>
    </div>

    <div class="card p-4 mb-4">
        <h5>⏰ اشتراكات منتهية أو تنتهي خلال {EXPIRY_WARNING_DAYS} أيام</h5>
        <div class="table-responsive">
            <table class="table table-hover">
                <thead><tr><th>الاسم</th><th>اليوزر</th><th>الهاتف</th><th>الاشتراك</th>
                <th>تاريخ الانتهاء</th><th>الحالة</th><th></th></tr></thead>
                <tbody>{expiring_rows}</tbody>
            </table>
        </div>
    </div>

    <div class="card p-4">
        <h5>آخر العمليات</h5>
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
        <option value="{plan['id']}" {sel}>
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
    if plan_filter.isdigit():
        where.append("s.plan_id = ?")
        params.append(int(plan_filter))

    plans_list = con.execute("SELECT * FROM plans ORDER BY category, id").fetchall()
    subscriber_rows = con.execute(f"""
        SELECT s.*, p.name AS plan_name, p.retail_price, p.wholesale_price, p.profit
        FROM subscribers s
        JOIN plans p ON p.id = s.plan_id
        {"WHERE " + " AND ".join(where) if where else ""}
        ORDER BY s.id DESC
    """, params).fetchall()
    con.close()

    subscriber_html = ""
    for s in subscriber_rows:
        subscriber_html += f"""
        <tr>
            <td>{esc(s['name'])}</td>
            <td>{esc(s['username'])}</td>
            <td>{esc(s['phone'])}</td>
            <td>{esc(s['plan_name'])}</td>
            <td>{money(s['retail_price'])}</td>
            <td>{money(s['wholesale_price'])}</td>
            <td>{esc(s['activation_date'])}</td>
            <td>{status_badge(s['expiry_date'])}</td>
            <td class="text-nowrap">
                <a href="{url_for('activate', subscriber_id=s['id'])}" class="btn btn-sm btn-success">تفعيل</a>
                <a href="{url_for('transfer', subscriber_id=s['id'])}" class="btn btn-sm btn-warning">ترحيل</a>
                <a href="{url_for('subscriber_detail', subscriber_id=s['id'])}" class="btn btn-sm btn-outline-primary">تفاصيل</a>
            </td>
        </tr>"""

    plan_filter_opts = "".join(
        f'<option value="{p["id"]}" {"selected" if str(p["id"]) == plan_filter else ""}>'
        f'{esc(p["category"])} - {esc(p["name"])}</option>' for p in plans_list)

    body = f"""
    <h2 class="mb-4">المشتركون</h2>

    <div class="card p-4 mb-4">
        <h5>إضافة مشترك جديد</h5>
        <form method="post" class="row g-3 mt-1">
            <div class="col-md-4">
                <label class="form-label">اسم المشترك</label>
                <input name="name" class="form-control" required>
            </div>
            <div class="col-md-4">
                <label class="form-label">يوزر المشترك</label>
                <input name="username" class="form-control" required>
            </div>
            <div class="col-md-4">
                <label class="form-label">رقم الهاتف</label>
                <input name="phone" class="form-control">
            </div>
            <div class="col-md-4">
                <label class="form-label">نوع الاشتراك</label>
                <select name="plan_id" class="form-select" required>
                    <option value="">اختر الاشتراك</option>
                    {plan_options(plans_list)}
                </select>
            </div>
            <div class="col-md-4">
                <label class="form-label">تاريخ التفعيل</label>
                <input type="date" name="activation_date" value="{today}" class="form-control" required>
                <small class="text-muted">ينتهي تلقائياً بعد {SUBSCRIPTION_DAYS} يوم</small>
            </div>
            <div class="col-md-4">
                <label class="form-label">ملاحظات</label>
                <input name="notes" class="form-control">
            </div>
            <div class="col-12">
                <button class="btn btn-primary">حفظ المشترك</button>
                <small class="text-muted ms-2">الإضافة لا تخصم من المحفظة — استخدم زر "تفعيل" للخصم.</small>
            </div>
        </form>
    </div>

    <div class="card p-4">
        <div class="d-flex flex-wrap justify-content-between align-items-center mb-3 gap-2">
            <h5 class="mb-0">قائمة المشتركين ({len(subscriber_rows)})</h5>
            <form method="get" class="d-flex flex-wrap gap-2">
                <input name="q" value="{esc(q)}" class="form-control" placeholder="🔍 بحث بالاسم/اليوزر/الهاتف" style="max-width:240px">
                <select name="status" class="form-select" style="max-width:150px">
                    <option value="">كل الحالات</option>
                    <option value="active" {"selected" if status == "active" else ""}>فعّال</option>
                    <option value="expired" {"selected" if status == "expired" else ""}>منتهي</option>
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
                        <th>سعر البيع</th><th>سعر الجملة</th><th>آخر تفعيل</th><th>الحالة</th><th>الإجراء</th>
                    </tr>
                </thead>
                <tbody>{subscriber_html or '<tr><td colspan="9" class="text-muted text-center">لا يوجد مشتركون</td></tr>'}</tbody>
            </table>
        </div>
    </div>
    """
    return page(body)


def get_subscriber(con, subscriber_id):
    return con.execute("""
        SELECT s.*, p.name AS plan_name, p.category,
               p.retail_price, p.wholesale_price, p.profit
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
    con.close()

    total_paid = sum(h["retail_price"] or 0 for h in history)
    total_profit = sum(h["profit"] or 0 for h in history)

    history_rows = "".join(f"""
        <tr>
            <td>{esc(h['operation_type'])}</td>
            <td>{esc(h['plan_name'])}</td>
            <td>{money(h['retail_price'])}</td>
            <td>{money(h['amount'])}</td>
            <td>{money(h['profit'])}</td>
            <td>{esc(h['notes'])}</td>
            <td>{esc(h['operation_date'])}</td>
        </tr>""" for h in history) or '<tr><td colspan="7" class="text-muted text-center">لا توجد عمليات لهذا المشترك</td></tr>'

    if history:
        delete_html = '<p class="text-muted small mb-0">لا يمكن حذف مشترك لديه عمليات مسجلة حتى لا تختل الحسابات والتقارير.</p>'
    else:
        delete_html = f"""
        <form method="post" action="{url_for('delete_subscriber', subscriber_id=subscriber_id)}"
              data-confirm="متأكد تريد تحذف هذا المشترك؟">
            <button class="btn btn-outline-danger">🗑️ حذف المشترك</button>
        </form>"""

    body = f"""
    <div class="d-flex justify-content-between align-items-center mb-4">
        <h2 class="mb-0">المشترك: {esc(s['name'])}</h2>
        <div>
            <a href="{url_for('activate', subscriber_id=subscriber_id)}" class="btn btn-success">تفعيل</a>
            <a href="{url_for('transfer', subscriber_id=subscriber_id)}" class="btn btn-warning">ترحيل</a>
            <a href="{url_for('subscribers')}" class="btn btn-secondary">رجوع</a>
        </div>
    </div>

    <div class="row g-3 mb-4">
        <div class="col-md-4"><div class="card p-3"><h6>الحالة</h6><div>{status_badge(s['expiry_date'])}</div>
            <small class="text-muted">ينتهي: {esc(s['expiry_date'])}</small></div></div>
        <div class="col-md-4"><div class="card p-3"><h6>مجموع المبيعات له</h6><h4>{money(total_paid)} د.ع</h4></div></div>
        <div class="col-md-4"><div class="card p-3"><h6>مجموع الأرباح منه</h6><h4 class="text-success">{money(total_profit)} د.ع</h4></div></div>
    </div>

    <div class="card p-4 mb-4">
        <h5>تعديل البيانات</h5>
        <form method="post" class="row g-3 mt-1">
            <div class="col-md-4">
                <label class="form-label">اسم المشترك</label>
                <input name="name" value="{esc(s['name'])}" class="form-control" required>
            </div>
            <div class="col-md-4">
                <label class="form-label">اليوزر</label>
                <input name="username" value="{esc(s['username'])}" class="form-control" required>
            </div>
            <div class="col-md-4">
                <label class="form-label">رقم الهاتف</label>
                <input name="phone" value="{esc(s['phone'])}" class="form-control">
            </div>
            <div class="col-md-4">
                <label class="form-label">نوع الاشتراك</label>
                <select name="plan_id" class="form-select" required>{plan_options(plans_list, s['plan_id'])}</select>
            </div>
            <div class="col-md-2">
                <label class="form-label">تاريخ التفعيل</label>
                <input type="date" name="activation_date" value="{esc(s['activation_date'])}" class="form-control" required>
            </div>
            <div class="col-md-2">
                <label class="form-label">تاريخ الانتهاء</label>
                <input type="date" name="expiry_date" value="{esc(s['expiry_date'])}" class="form-control">
            </div>
            <div class="col-md-4">
                <label class="form-label">ملاحظات</label>
                <input name="notes" value="{esc(s['notes'])}" class="form-control">
            </div>
            <div class="col-12">
                <button class="btn btn-primary">حفظ التعديلات</button>
            </div>
        </form>
    </div>

    <div class="card p-4 mb-4">
        <h5>سجل العمليات</h5>
        <div class="table-responsive">
            <table class="table table-striped">
                <thead><tr><th>العملية</th><th>الاشتراك</th><th>سعر البيع</th><th>سعر الجملة</th>
                <th>الربح</th><th>الملاحظات</th><th>التاريخ</th></tr></thead>
                <tbody>{history_rows}</tbody>
            </table>
        </div>
    </div>

    <div class="card p-4 border border-danger-subtle">{delete_html}</div>
    """
    return page(body)


@app.route("/subscribers/<int:subscriber_id>/delete", methods=["POST"])
@login_required
def delete_subscriber(subscriber_id):
    con = db()
    has_ops = con.execute("SELECT 1 FROM activations WHERE subscriber_id = ?", (subscriber_id,)).fetchone()
    if has_ops:
        con.close()
        flash("لا يمكن حذف مشترك لديه عمليات مسجلة", "danger")
        return redirect(url_for("subscriber_detail", subscriber_id=subscriber_id))
    con.execute("DELETE FROM subscribers WHERE id = ?", (subscriber_id,))
    con.commit()
    con.close()
    flash("تم حذف المشترك", "success")
    return redirect(url_for("subscribers"))


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

    if balance < amount:
        con.rollback()
        con.close()
        flash(f"الرصيد غير كافٍ. المطلوب {money(amount)} د.ع، والرصيد الحالي {money(balance)} د.ع", "danger")
        return redirect(url_for("subscribers"))

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

    con.execute("""
        INSERT INTO activations
        (subscriber_id, plan_id, operation_type, amount, retail_price, profit, notes, operation_date)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (subscriber_id, plan_id, operation_type, amount,
          plan["retail_price"], plan["profit"], notes, now))

    con.execute("""
        UPDATE subscribers
        SET activation_date = ?, expiry_date = ?, plan_id = ?,
            notes = CASE WHEN ? <> '' THEN ? ELSE notes END
        WHERE id = ?
    """, (today_str(), new_expiry, plan_id, notes, notes, subscriber_id))

    con.commit()
    con.close()

    flash(f"تمت عملية {operation_type} للمشترك {subscriber['username']} وخصم {money(amount)} د.ع "
          f"من المحفظة — ينتهي الاشتراك بتاريخ {new_expiry}", "success")
    return redirect(url_for("subscribers"))


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

    body = f"""
    <div class="card p-4">
        <h3>{operation_type} الاشتراك</h3>
        <hr>
        <p><b>اسم المشترك:</b> {esc(subscriber['name'])}</p>
        <p><b>اليوزر:</b> {esc(subscriber['username'])}</p>
        <p><b>الحالة الحالية:</b> {status_badge(subscriber['expiry_date'])} (ينتهي {esc(subscriber['expiry_date'])})</p>
        <p><b>الرصيد الحالي:</b> <span class="{'text-danger' if low else 'text-success'}">{money(balance)} د.ع</span></p>

        <form method="post">
            <label class="form-label">الباقة</label>
            <select name="plan_id" class="form-select mb-2">{plan_options(plans_list, subscriber['plan_id'])}</select>
            <small class="text-muted d-block mb-3">المبلغ المخصوم من المحفظة = سعر الجملة للباقة المختارة.
            {"عند الترحيل تُضاف المدة على تاريخ الانتهاء الحالي إذا كان الاشتراك بعده فعّال." if operation_type == "ترحيل" else ""}</small>

            <label class="form-label">ملاحظات {operation_type}</label>
            <textarea name="notes" class="form-control mb-3" placeholder="اكتب الملاحظات هنا"></textarea>

            <button class="btn btn-success">تأكيد {operation_type}</button>
            <a href="{url_for('subscribers')}" class="btn btn-secondary">رجوع</a>
        </form>
    </div>
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
    <div class="d-flex justify-content-between align-items-center mb-4">
        <h2 class="mb-0">التقارير</h2>
        <button onclick="window.print()" class="btn btn-outline-secondary no-print">🖨️ طباعة</button>
    </div>

    <div class="card p-4 mb-4 no-print">
        <form method="get" class="row g-3 align-items-end">
            <div class="col-md-3"><label class="form-label">من تاريخ</label>
                <input type="date" name="from" value="{start}" class="form-control"></div>
            <div class="col-md-3"><label class="form-label">إلى تاريخ</label>
                <input type="date" name="to" value="{end}" class="form-control"></div>
            <div class="col-md-2"><button class="btn btn-primary w-100">عرض</button></div>
            <div class="col-md-4 text-end">
                <a href="{url_for('export_activations')}{qs}" class="btn btn-success">📥 Excel العمليات</a>
                <a href="{url_for('export_charges')}{qs}" class="btn btn-primary">📥 Excel المحفظة</a>
            </div>
        </form>
    </div>

    <p class="text-muted">الفترة: من <b>{start}</b> إلى <b>{end}</b></p>

    <div class="row g-3 mb-4">
        <div class="col-md-3"><div class="card p-3"><h6>عدد العمليات</h6><h3>{len(activations)}</h3></div></div>
        <div class="col-md-3"><div class="card p-3"><h6>المبيعات</h6><h3>{money(total_sales)}</h3></div></div>
        <div class="col-md-3"><div class="card p-3"><h6>الخصم من المحفظة</h6><h3 class="text-danger">{money(total_cost)}</h3></div></div>
        <div class="col-md-3"><div class="card p-3"><h6>صافي الربح</h6><h3 class="text-success">{money(total_profit)}</h3></div></div>
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
                <thead><tr><th>العملية</th><th>المشترك</th><th>اليوزر</th><th>الاشتراك</th><th>سعر البيع</th>
                <th>سعر الجملة</th><th>الربح</th><th>الملاحظات</th><th>التاريخ</th></tr></thead>
                <tbody>{activation_rows or '<tr><td colspan="9" class="text-muted text-center">لا توجد عمليات بهذه الفترة</td></tr>'}</tbody>
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
               "سعر البيع", "سعر الجملة (المخصوم)", "الربح", "الملاحظات", "التاريخ"]
    totals = ["المجموع", "", "", "", "", "",
              sum(r["retail_price"] or 0 for r in rows),
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
                <h5>🔑 تغيير كلمة المرور</h5>
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
                <h5>👤 المستخدمون</h5>
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
                <h5>💾 النسخ الاحتياطي</h5>
                <p class="text-muted mb-3">نزّل نسخة كاملة من قاعدة البيانات واحتفظ بيها بمكان آمن (يُنصح بعمل نسخة كل أسبوع).</p>
                <div><a href="{url_for('backup')}" class="btn btn-dark">تحميل نسخة احتياطية</a></div>
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
