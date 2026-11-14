"""
النسخ الاحتياطي اليومي.

يسوي نسخة من قاعدة البيانات بمجلد backups (يحتفظ بآخر 30 يوم)
ويبعثها للتليكرام إذا مربوط من صفحة الإعدادات.

على PythonAnywhere: صفحة Tasks ← أضف مهمة يومية بالأمر:
    python3 /home/USERNAME/internet-accounting/backup_job.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import run_backup  # noqa: E402

if __name__ == "__main__":
    print(run_backup())
