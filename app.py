from flask import (
    Flask, request, redirect, url_for, session,
    render_template_string, send_file, flash
)
import sqlite3
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from functools import wraps
from markupsafe import escape
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = "change-this-secret-key"

DB_NAME = "internet_accounting.db"
PROFIT = 5000

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


def esc(value):
    """تنظيف النصوص المدخلة من المستخدم قبل عرضها في HTML."""
    return escape(value if value is not None else "")


def db():
    connection = sqlite3.connect(DB_NAME)
    connection.row_factory = sqlite3.Row
    return connection


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

    user = cur.execute(
        "SELECT id FROM users WHERE username = ?",
        ("admin",)
    ).fetchone()

    if not user:
        cur.execute(
            "INSERT INTO users(username, password) VALUES (?, ?)",
            ("admin", generate_password_hash("admin123"))
        )

    count = cur.execute("SELECT COUNT(*) AS c FROM plans").fetchone()["c"]

    if count == 0:
        for category, name, retail_price in PLANS:
            cur.execute("""
                INSERT INTO plans
                (category, name, retail_price, profit, wholesale_price)
                VALUES (?, ?, ?, ?, ?)
            """, (
                category,
                name,
                retail_price,
                PROFIT,
                retail_price - PROFIT
            ))

    con.commit()
    con.close()


def login_required(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return func(*args, **kwargs)
    return wrapper


def wallet_balance():
    con = db()
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
    con.close()
    return result["balance"]


BASE_HTML = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>برنامج محاسبة الإنترنت</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.rtl.min.css" rel="stylesheet">

    <style>
        body {
            background: #f1f5f9;
            font-family: Tahoma, Arial, sans-serif;
        }

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
            display: block;
            color: white;
            text-decoration: none;
            padding: 12px;
            border-radius: 10px;
            margin-bottom: 7px;
        }

        .sidebar a:hover {
            background: rgba(255,255,255,.18);
        }

        .content {
            padding: 25px;
        }

        .card {
            border: none;
            border-radius: 16px;
            box-shadow: 0 5px 18px rgba(15,23,42,.08);
        }

        .balance-card {
            background: linear-gradient(135deg, #0f766e, #14b8a6);
            color: white;
        }

        .stat-card {
            min-height: 125px;
        }

        .table {
            vertical-align: middle;
        }

        .btn {
            border-radius: 9px;
        }

        .form-control, .form-select {
            border-radius: 9px;
            padding: 10px;
        }

        .brand {
            font-weight: bold;
            font-size: 20px;
        }
    </style>
</head>

<body>
<div class="container-fluid">
    <div class="row">

        {% if session.get('user_id') %}
        <aside class="col-md-2 sidebar">
            <h4 class="brand">📡 محاسبة الإنترنت</h4>

            <a href="{{ url_for('dashboard') }}">🏠 الرئيسية</a>
            <a href="{{ url_for('subscribers') }}">👥 المشتركون</a>
            <a href="{{ url_for('wallet') }}">💰 المحفظة</a>
            <a href="{{ url_for('reports') }}">📊 التقارير</a>
            <a href="{{ url_for('logout') }}">🚪 تسجيل الخروج</a>
        </aside>
        {% endif %}

        <main class="{% if session.get('user_id') %}col-md-10{% else %}col-md-12{% endif %} content">

            {% with messages = get_flashed_messages() %}
                {% if messages %}
                    {% for message in messages %}
                        <div class="alert alert-info">{{ message }}</div>
                    {% endfor %}
                {% endif %}
            {% endwith %}

            {{ body|safe }}
        </main>
    </div>
</div>
</body>
</html>
"""

LOGIN_PAGE = """
<div class="row justify-content-center mt-5">
    <div class="col-md-5">
        <div class="card p-4">
            <h3 class="text-center mb-4">تسجيل الدخول</h3>

            <form method="post">
                <label class="form-label">اسم المستخدم</label>
                <input name="username" class="form-control mb-3" required>

                <label class="form-label">كلمة المرور</label>
                <input name="password" type="password" class="form-control mb-3" required>

                <button class="btn btn-primary w-100">دخول</button>
            </form>

            <div class="text-muted mt-3 text-center">
                الدخول الافتراضي: admin / admin123
            </div>
        </div>
    </div>
</div>
"""


@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        con = db()
        user = con.execute(
            "SELECT * FROM users WHERE username = ?",
            (username,)
        ).fetchone()
        con.close()

        if user and check_password_hash(user["password"], password):
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            return redirect(url_for("dashboard"))

        flash("اسم المستخدم أو كلمة المرور غير صحيحة")

    return render_template_string(
        BASE_HTML,
        body=render_template_string(LOGIN_PAGE)
    )


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/dashboard")
@login_required
def dashboard():
    con = db()

    subscribers_count = con.execute(
        "SELECT COUNT(*) AS c FROM subscribers"
    ).fetchone()["c"]

    activations_count = con.execute(
        "SELECT COUNT(*) AS c FROM activations WHERE operation_type = 'تفعيل'"
    ).fetchone()["c"]

    transfers_count = con.execute(
        "SELECT COUNT(*) AS c FROM activations WHERE operation_type = 'ترحيل'"
    ).fetchone()["c"]

    total_profit = con.execute("""
        SELECT COALESCE(SUM(p.profit), 0) AS total
        FROM activations a
        JOIN plans p ON p.id = a.plan_id
    """).fetchone()["total"]

    con.close()

    body = f"""
    <h2 class="mb-4">لوحة التحكم</h2>

    <div class="row g-3 mb-4">
        <div class="col-md-3">
            <div class="card balance-card stat-card p-4">
                <h6>رصيد المحفظة الحالي</h6>
                <h2>{wallet_balance():,} د.ع</h2>
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
                <h6>عمليات التفعيل</h6>
                <h2>{activations_count}</h2>
            </div>
        </div>

        <div class="col-md-3">
            <div class="card stat-card p-4">
                <h6>إجمالي الأرباح</h6>
                <h2>{total_profit:,} د.ع</h2>
            </div>
        </div>
    </div>

    <div class="card p-4">
        <h5>ملخص النظام</h5>
        <p>عدد عمليات الترحيل: <b>{transfers_count}</b></p>
        <p>سعر الربح الثابت لكل اشتراك: <b>{PROFIT:,} د.ع</b></p>
    </div>
    """

    return render_template_string(BASE_HTML, body=body)


@app.route("/wallet", methods=["GET", "POST"])
@login_required
def wallet():
    if request.method == "POST":
        try:
            amount = int(request.form["amount"])
        except (KeyError, ValueError):
            flash("أدخل مبلغاً صحيحاً")
            return redirect(url_for("wallet"))

        notes = request.form.get("notes", "")
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if amount <= 0:
            flash("يجب أن يكون مبلغ الشحن أكبر من صفر")
            return redirect(url_for("wallet"))

        con = db()
        con.execute("""
            INSERT INTO wallet_transactions
            (transaction_type, amount, notes, created_at)
            VALUES (?, ?, ?, ?)
        """, ("شحن", amount, notes, now))
        con.commit()
        con.close()

        flash("تم شحن المحفظة بنجاح")
        return redirect(url_for("wallet"))

    con = db()
    transactions = con.execute("""
        SELECT * FROM wallet_transactions
        ORDER BY id DESC
    """).fetchall()
    con.close()

    rows = ""

    for item in transactions:
        color = "text-success" if item["transaction_type"] == "شحن" else "text-danger"
        sign = "+" if item["transaction_type"] == "شحن" else "-"

        rows += f"""
        <tr>
            <td>{esc(item['transaction_type'])}</td>
            <td class="{color}">{sign}{item['amount']:,}</td>
            <td>{esc(item['notes'])}</td>
            <td>{item['created_at']}</td>
        </tr>
        """

    body = f"""
    <h2 class="mb-4">المحفظة</h2>

    <div class="card p-4 mb-4">
        <h4>الرصيد الحالي: <span class="text-success">{wallet_balance():,} د.ع</span></h4>

        <form method="post" class="row g-3 mt-2">
            <div class="col-md-4">
                <label class="form-label">مبلغ الشحن</label>
                <input type="number" name="amount" class="form-control" required>
            </div>

            <div class="col-md-6">
                <label class="form-label">ملاحظات الشحن</label>
                <input name="notes" class="form-control" placeholder="مثال: شحن من الإدارة">
            </div>

            <div class="col-md-2 d-flex align-items-end">
                <button class="btn btn-success w-100">شحن المحفظة</button>
            </div>
        </form>
    </div>

    <div class="card p-4">
        <h5>حركات المحفظة</h5>

        <div class="table-responsive">
            <table class="table table-striped">
                <thead>
                    <tr>
                        <th>النوع</th>
                        <th>المبلغ</th>
                        <th>الملاحظات</th>
                        <th>التاريخ</th>
                    </tr>
                </thead>
                <tbody>{rows}</tbody>
            </table>
        </div>
    </div>
    """

    return render_template_string(BASE_HTML, body=body)


@app.route("/subscribers", methods=["GET", "POST"])
@login_required
def subscribers():
    con = db()

    if request.method == "POST":
        name = request.form["name"]
        username = request.form["username"]
        phone = request.form.get("phone", "")
        plan_id = int(request.form["plan_id"])
        activation_date = request.form["activation_date"]
        notes = request.form.get("notes", "")
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        con.execute("""
            INSERT INTO subscribers
            (name, username, phone, plan_id, activation_date, notes, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            name,
            username,
            phone,
            plan_id,
            activation_date,
            notes,
            now
        ))

        con.commit()
        con.close()

        flash("تمت إضافة المشترك بنجاح")
        return redirect(url_for("subscribers"))

    plans = con.execute("SELECT * FROM plans ORDER BY category, id").fetchall()

    subscriber_rows = con.execute("""
        SELECT s.*, p.name AS plan_name, p.retail_price,
               p.wholesale_price, p.profit
        FROM subscribers s
        JOIN plans p ON p.id = s.plan_id
        ORDER BY s.id DESC
    """).fetchall()

    con.close()

    plan_options = ""

    for plan in plans:
        plan_options += f"""
        <option value="{plan['id']}">
            {esc(plan['category'])} - {esc(plan['name'])}
            | البيع {plan['retail_price']:,}
            | الجملة {plan['wholesale_price']:,}
        </option>
        """

    subscriber_html = ""

    for s in subscriber_rows:
        subscriber_html += f"""
        <tr>
            <td>{esc(s['name'])}</td>
            <td>{esc(s['username'])}</td>
            <td>{esc(s['phone'])}</td>
            <td>{esc(s['plan_name'])}</td>
            <td>{s['retail_price']:,}</td>
            <td>{s['wholesale_price']:,}</td>
            <td>{s['activation_date']}</td>
            <td>
                <a href="{url_for('activate', subscriber_id=s['id'])}"
                   class="btn btn-sm btn-success">
                   تفعيل
                </a>

                <a href="{url_for('transfer', subscriber_id=s['id'])}"
                   class="btn btn-sm btn-warning">
                   ترحيل
                </a>
            </td>
        </tr>
        """

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
                    {plan_options}
                </select>
            </div>

            <div class="col-md-4">
                <label class="form-label">تاريخ التفعيل</label>
                <input type="date" name="activation_date"
                       value="{datetime.now().strftime('%Y-%m-%d')}"
                       class="form-control" required>
            </div>

            <div class="col-md-4">
                <label class="form-label">ملاحظات</label>
                <input name="notes" class="form-control">
            </div>

            <div class="col-12">
                <button class="btn btn-primary">حفظ المشترك</button>
            </div>
        </form>
    </div>

    <div class="card p-4">
        <h5>قائمة المشتركين</h5>

        <div class="table-responsive">
            <table class="table table-hover">
                <thead>
                    <tr>
                        <th>الاسم</th>
                        <th>اليوزر</th>
                        <th>الهاتف</th>
                        <th>الاشتراك</th>
                        <th>سعر البيع</th>
                        <th>سعر الجملة</th>
                        <th>التاريخ</th>
                        <th>الإجراء</th>
                    </tr>
                </thead>
                <tbody>{subscriber_html}</tbody>
            </table>
        </div>
    </div>
    """

    return render_template_string(BASE_HTML, body=body)


def perform_operation(subscriber_id, operation_type):
    con = db()

    subscriber = con.execute("""
        SELECT s.*, p.name AS plan_name, p.wholesale_price
        FROM subscribers s
        JOIN plans p ON p.id = s.plan_id
        WHERE s.id = ?
    """, (subscriber_id,)).fetchone()

    if not subscriber:
        con.close()
        flash("المشترك غير موجود")
        return redirect(url_for("subscribers"))

    amount = subscriber["wholesale_price"]
    balance = wallet_balance()

    if balance < amount:
        con.close()
        flash(
            f"الرصيد غير كافٍ. المطلوب {amount:,} د.ع، "
            f"والرصيد الحالي {balance:,} د.ع"
        )
        return redirect(url_for("subscribers"))

    notes = request.form.get("notes", "")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    con.execute("""
        INSERT INTO wallet_transactions
        (transaction_type, amount, notes, created_at)
        VALUES (?, ?, ?, ?)
    """, (
        "خصم",
        amount,
        f"{operation_type} - {subscriber['username']} - {notes}",
        now
    ))

    con.execute("""
        INSERT INTO activations
        (subscriber_id, plan_id, operation_type, amount, notes, operation_date)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        subscriber_id,
        subscriber["plan_id"],
        operation_type,
        amount,
        notes,
        now
    ))

    con.execute("""
        UPDATE subscribers
        SET activation_date = ?, notes = ?
        WHERE id = ?
    """, (
        now[:10],
        notes,
        subscriber_id
    ))

    con.commit()
    con.close()

    flash(
        f"تمت عملية {operation_type} للمشترك "
        f"{subscriber['username']} وخصم {amount:,} د.ع من المحفظة"
    )

    return redirect(url_for("subscribers"))


def operation_page(subscriber_id, operation_type):
    con = db()
    subscriber = con.execute("""
        SELECT s.*, p.name AS plan_name,
               p.retail_price, p.wholesale_price
        FROM subscribers s
        JOIN plans p ON p.id = s.plan_id
        WHERE s.id = ?
    """, (subscriber_id,)).fetchone()
    con.close()

    if not subscriber:
        flash("المشترك غير موجود")
        return redirect(url_for("subscribers"))

    body = f"""
    <div class="card p-4">
        <h3>{operation_type} الاشتراك</h3>

        <hr>

        <p><b>اسم المشترك:</b> {esc(subscriber['name'])}</p>
        <p><b>اليوزر:</b> {esc(subscriber['username'])}</p>
        <p><b>نوع الاشتراك:</b> {esc(subscriber['plan_name'])}</p>
        <p><b>سعر البيع:</b> {subscriber['retail_price']:,} د.ع</p>
        <p><b>المبلغ المخصوم من المحفظة:</b>
            <span class="text-danger">
                {subscriber['wholesale_price']:,} د.ع
            </span>
        </p>
        <p><b>الرصيد الحالي:</b> {wallet_balance():,} د.ع</p>

        <form method="post">
            <label class="form-label">ملاحظات {operation_type}</label>
            <textarea name="notes" class="form-control mb-3"
                      placeholder="اكتب الملاحظات هنا"></textarea>

            <button class="btn btn-success">تأكيد {operation_type}</button>
            <a href="{url_for('subscribers')}" class="btn btn-secondary">
                رجوع
            </a>
        </form>
    </div>
    """

    return render_template_string(BASE_HTML, body=body)


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


@app.route("/reports")
@login_required
def reports():
    con = db()

    activations = con.execute("""
        SELECT a.*, s.name AS subscriber_name,
               s.username AS subscriber_username,
               p.name AS plan_name,
               p.retail_price,
               p.profit
        FROM activations a
        JOIN subscribers s ON s.id = a.subscriber_id
        JOIN plans p ON p.id = a.plan_id
        ORDER BY a.id DESC
    """).fetchall()

    charges = con.execute("""
        SELECT * FROM wallet_transactions
        WHERE transaction_type = 'شحن'
        ORDER BY id DESC
    """).fetchall()

    con.close()

    activation_rows = ""

    for a in activations:
        activation_rows += f"""
        <tr>
            <td>{esc(a['operation_type'])}</td>
            <td>{esc(a['subscriber_name'])}</td>
            <td>{esc(a['subscriber_username'])}</td>
            <td>{esc(a['plan_name'])}</td>
            <td>{a['retail_price']:,}</td>
            <td>{a['amount']:,}</td>
            <td>{a['profit']:,}</td>
            <td>{esc(a['notes'])}</td>
            <td>{a['operation_date']}</td>
        </tr>
        """

    charge_rows = ""

    for c in charges:
        charge_rows += f"""
        <tr>
            <td>{c['amount']:,}</td>
            <td>{esc(c['notes'])}</td>
            <td>{c['created_at']}</td>
        </tr>
        """

    body = f"""
    <h2 class="mb-4">التقارير</h2>

    <div class="row g-4">
        <div class="col-md-6">
            <div class="card p-4">
                <h5>تحميل تقرير اليوزرات</h5>
                <p>يتضمن التفعيل والترحيل وسعر الجملة والأرباح.</p>
                <a href="{url_for('export_activations')}"
                   class="btn btn-success">
                   تحميل Excel
                </a>
            </div>
        </div>

        <div class="col-md-6">
            <div class="card p-4">
                <h5>تحميل تقرير شحن المحفظة</h5>
                <p>يتضمن جميع مبالغ الشحن والملاحظات والتواريخ.</p>
                <a href="{url_for('export_charges')}"
                   class="btn btn-primary">
                   تحميل Excel
                </a>
            </div>
        </div>
    </div>

    <div class="card p-4 mt-4">
        <h5>تقرير التفعيل والترحيل</h5>

        <div class="table-responsive">
            <table class="table table-striped">
                <thead>
                    <tr>
                        <th>العملية</th>
                        <th>المشترك</th>
                        <th>اليوزر</th>
                        <th>الاشتراك</th>
                        <th>سعر البيع</th>
                        <th>سعر الجملة</th>
                        <th>الربح</th>
                        <th>الملاحظات</th>
                        <th>التاريخ</th>
                    </tr>
                </thead>
                <tbody>{activation_rows}</tbody>
            </table>
        </div>
    </div>

    <div class="card p-4 mt-4">
        <h5>تقرير شحن المحفظة</h5>

        <div class="table-responsive">
            <table class="table table-striped">
                <thead>
                    <tr>
                        <th>المبلغ</th>
                        <th>الملاحظات</th>
                        <th>التاريخ</th>
                    </tr>
                </thead>
                <tbody>{charge_rows}</tbody>
            </table>
        </div>
    </div>
    """

    return render_template_string(BASE_HTML, body=body)


@app.route("/export/activations")
@login_required
def export_activations():
    con = db()

    rows = con.execute("""
        SELECT a.operation_type,
               s.name AS subscriber_name,
               s.username AS subscriber_username,
               s.phone,
               p.category,
               p.name AS plan_name,
               p.retail_price,
               p.wholesale_price,
               p.profit,
               a.amount,
               a.notes,
               a.operation_date
        FROM activations a
        JOIN subscribers s ON s.id = a.subscriber_id
        JOIN plans p ON p.id = a.plan_id
        ORDER BY a.id DESC
    """).fetchall()

    con.close()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "تقرير اليوزرات"

    headers = [
        "نوع العملية",
        "اسم المشترك",
        "اليوزر",
        "رقم الهاتف",
        "الفئة",
        "نوع الاشتراك",
        "سعر البيع",
        "سعر الجملة",
        "الربح",
        "المبلغ المخصوم",
        "الملاحظات",
        "التاريخ"
    ]

    sheet.append(headers)

    for row in rows:
        sheet.append(list(row))

    file = BytesIO()
    workbook.save(file)
    file.seek(0)

    return send_file(
        file,
        as_attachment=True,
        download_name="users_report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/export/charges")
@login_required
def export_charges():
    con = db()

    rows = con.execute("""
        SELECT amount, notes, created_at
        FROM wallet_transactions
        WHERE transaction_type = 'شحن'
        ORDER BY id DESC
    """).fetchall()

    con.close()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "شحن المحفظة"

    sheet.append([
        "مبلغ الشحن",
        "الملاحظات",
        "التاريخ"
    ])

    for row in rows:
        sheet.append(list(row))

    file = BytesIO()
    workbook.save(file)
    file.seek(0)

    return send_file(
        file,
        as_attachment=True,
        download_name="wallet_charges.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# تهيئة قاعدة البيانات عند التشغيل (سواء بـ python app.py أو flask run)
init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
