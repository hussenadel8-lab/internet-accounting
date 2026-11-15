"""
إعادة تعيين كلمة السر (إذا نسيتها).

يشتغل من الشاشة السوداء (Bash) بـ PythonAnywhere:
    cd ~/internet-accounting
    python3 reset_password.py                 ← يعرض أسماء المستخدمين
    python3 reset_password.py admin 12345678  ← يغيّر كلمة سر admin إلى 12345678
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from werkzeug.security import generate_password_hash  # noqa: E402
from app import db  # noqa: E402


def main():
    con = db()
    users = con.execute("SELECT username, role FROM users ORDER BY id").fetchall()
    if len(sys.argv) < 3:
        print("المستخدمين:")
        for u in users:
            kind = "حساب الدفع من قبلي" if u["role"] == "self" else "البرنامج الرئيسي"
            print(f"  - {u['username']}  ({kind})")
        print("\nلتغيير كلمة السر اكتب:  python3 reset_password.py اسم_المستخدم كلمة_السر_الجديدة")
        return
    username, password = sys.argv[1], sys.argv[2]
    if len(password) < 6:
        print("❌ كلمة السر لازم تكون 6 أحرف أو أكثر")
        return
    cur = con.execute("UPDATE users SET password = ? WHERE username = ?",
                      (generate_password_hash(password), username))
    con.commit()
    if cur.rowcount:
        print(f"✅ تم تغيير كلمة سر ({username}). ادخل بيها هسه.")
    else:
        print(f"❌ ماكو مستخدم باسم ({username}). شغّل الأمر بدون إضافات حتى تشوف الأسماء.")


if __name__ == "__main__":
    main()
