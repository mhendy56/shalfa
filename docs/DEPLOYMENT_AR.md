# Shalfa Operations Enterprise V18

## البنية
- `main.py`: نقطة تشغيل ASGI فقط.
- `app/factory.py`: إنشاء التطبيق وإضافة Routers.
- `app/legacy.py`: الوظائف التشغيلية الحالية محفوظة دون فقد أي شاشة.
- `app/config.py`: إعدادات البيئة.
- `app/database.py`: طبقة توافق SQLite/PostgreSQL.
- `app/routers/`: المسارات الجديدة المنفصلة.
- `scripts/migrate_sqlite_to_postgres.py`: نقل البيانات القديمة.

## النشر على Render
1. ارفع الملفات إلى GitHub.
2. New → Blueprint.
3. اختر المستودع؛ سيقرأ `render.yaml` وينشئ PostgreSQL والخدمة.
4. اختبر `/health/database`.

## نقل البيانات
بعد إنشاء جداول النسخة الجديدة، شغّل:
`python scripts/migrate_sqlite_to_postgres.py data/shalfa.db`
مع تعريف `DATABASE_URL`.

## ملاحظة
طبقة PostgreSQL مجهزة، لكن يظل الاختبار النهائي لترحيل كامل بيانات الإنتاج مطلوبًا على نسخة احتياطية قبل تحويل الخدمة الرئيسية.
