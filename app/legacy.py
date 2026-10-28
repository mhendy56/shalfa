
from __future__ import annotations
import os, json, hashlib, hmac, secrets, shutil, re, asyncio
from datetime import datetime, date, timedelta
from pathlib import Path
from app.config import settings
from app.database import connect as database_connect, IntegrityError
from app.unified_master import ensure_unified_schema, reconcile_unified_data, relationship_summary, record_relationship_event
from app.services.master_data_service import MasterDataService
from typing import Optional

from fastapi import FastAPI, Request, Form, UploadFile, File, HTTPException, Depends
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from itsdangerous import URLSafeSerializer, BadSignature
from openpyxl import load_workbook, Workbook
from io import BytesIO

BASE = Path(__file__).resolve().parent.parent
EVIDENCE_DIR = BASE / "data" / "access_evidence"
EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = settings.database_path
DB_PATH.parent.mkdir(parents=True, exist_ok=True)
SECRET_KEY = settings.secret_key

# Ensure required folders/files exist even if GitHub upload omitted directories.
STATIC_DIR = BASE / "static"
TEMPLATES_DIR = BASE / "templates"
STATIC_DIR.mkdir(parents=True, exist_ok=True)
TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)

_FALLBACK_STYLE = '\n*{box-sizing:border-box}body{margin:0;font-family:Tahoma,Arial;background:#eef3f8;color:#10233d}header{background:#082b60;color:#fff;padding:14px 20px;display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;z-index:5}header b{display:block;font-size:20px}header small{opacity:.85}header a{color:#fff;text-decoration:none;background:#d93025;padding:9px 13px;border-radius:8px}main{padding:12px;max-width:1500px;margin:auto}.filters,.import{background:#fff;border-radius:12px;padding:10px;display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px}.filters>*{flex:1;min-width:140px;padding:9px;border:1px solid #ccd6e1;border-radius:7px}.filters button,.export,.import button{background:#082b60;color:#fff;border:0;text-decoration:none;text-align:center;font-weight:bold}.export{padding:10px;border-radius:7px}.kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:8px;margin-bottom:10px}.kpis div{color:#fff;border-radius:10px;text-align:center;padding:10px}.kpis b{display:block;font-size:24px}.green{background:#188038}.red{background:#d93025}.orange{background:#f9ab00}.blue{background:#1a73e8}.gray{background:#6b7280}.purple{background:#6339a6}.table-wrap{background:#fff;border-radius:12px;overflow:auto;max-height:70vh}table{width:100%;border-collapse:collapse;font-size:11px}th{position:sticky;top:0;background:#082b60;color:#fff;padding:8px}td{border:1px solid #dce4ed;padding:7px;text-align:center}.actions{display:flex;gap:3px}.actions button{border:0;border-radius:5px;padding:6px;font-size:9px;cursor:pointer}.badge{padding:4px 8px;border-radius:999px;color:#fff;font-weight:bold}.s-حاضر{background:#188038}.s-غائب{background:#d93025}.s-إجازة{background:#f9ab00}.s-بديل{background:#1a73e8}.s-لم-يسجل{background:#6b7280}.login-body{min-height:100vh;display:grid;place-items:center;background:linear-gradient(135deg,#082b60,#0f4c8c)}.login-card{width:min(420px,92vw);background:#fff;padding:24px;border-radius:16px}.login-card h1,.login-card h2{text-align:center;color:#082b60}.login-card label{display:block;margin-top:10px}.login-card input{width:100%;padding:11px;border:1px solid #ccd6e1;border-radius:8px}.login-card button{width:100%;margin-top:15px;padding:12px;border:0;border-radius:8px;background:#082b60;color:#fff;font-weight:bold}.hint{margin-top:12px;background:#eef3f8;padding:10px;border-radius:8px;font-size:12px}.error{background:#fde8e7;color:#a61b13;padding:9px;border-radius:7px}@media(max-width:800px){.kpis{grid-template-columns:repeat(3,1fr)}header b{font-size:15px}.actions{flex-wrap:wrap}}\n'
_FALLBACK_LOGIN = '<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>دخول شلفا</title><link rel="stylesheet" href="/static/style.css"></head><body class="login-body"><form class="login-card" method="post"><h1>شركة شلفا لإدارة المرافق</h1><h2>نظام حضور البوابين المركزي</h2>{% if error %}<div class="error">{{error}}</div>{% endif %}<label>اسم المستخدم</label><input name="username" required><label>كلمة المرور</label><input name="password" type="password" required><button>دخول</button><div class="hint">الإدارة: admin / 1234<br>المنسق: الرقم الوظيفي / 1234</div></form></body></html>'
_FALLBACK_DASHBOARD = '<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>غرفة العمليات</title><link rel="stylesheet" href="/static/style.css"></head><body>\n<header><div><b>غرفة عمليات حضور البوابين</b><small>{{user.full_name}} — {{user.role}}</small></div><a href="/logout">خروج</a></header>\n<main>\n<form class="filters" method="get"><input type="date" name="work_date" value="{{work_date}}"><select name="city"><option value="">كل المناطق</option>{% for x in cities %}<option value="{{x}}" {% if city==x %}selected{% endif %}>{{x}}</option>{% endfor %}</select>{% if user.role!=\'coordinator\' %}<select name="coordinator"><option value="">كل المنسقين</option>{% for x in coords %}<option value="{{x}}" {% if coordinator==x %}selected{% endif %}>{{x}}</option>{% endfor %}</select>{% endif %}<button>تحديث</button><a class="export" href="/export.xlsx?work_date={{work_date}}&city={{city}}&coordinator={{coordinator}}">تصدير Excel</a></form>\n<section class="kpis"><div class="green">حاضر<b>{{counts[\'حاضر\']}}</b></div><div class="red">غائب<b>{{counts[\'غائب\']}}</b></div><div class="orange">إجازة<b>{{counts[\'إجازة\']}}</b></div><div class="blue">بديل<b>{{counts[\'بديل\']}}</b></div><div class="gray">لم يسجل<b>{{counts[\'لم يسجل\']}}</b></div><div class="purple">التغطية<b>{{\'%.1f\'|format(coverage)}}%</b></div></section>\n<form class="import" action="/attendance/import" method="post" enctype="multipart/form-data"><input type="hidden" name="work_date" value="{{work_date}}"><input type="file" name="file" accept=".xlsx" required><button>استيراد حضور Excel</button></form>\n<div class="table-wrap"><table><thead><tr><th>المدرسة</th><th>المدينة</th><th>المنسق</th><th>البواب</th><th>الرقم الوظيفي</th><th>الحالة</th><th>تسجيل</th></tr></thead><tbody>{% for r in rows %}<tr><td>{{r.school_name}}</td><td>{{r.city}}</td><td>{{r.coordinator_name}}</td><td>{{r.guard_name}}</td><td>{{r.guard_emp_id}}</td><td><span class="badge s-{{r.status or \'لم-يسجل\'}}">{{r.status or \'لم يسجل\'}}</span></td><td><div class="actions" data-mid="{{r.ministry_id}}">{% for st in [\'حاضر\',\'غائب\',\'إجازة\',\'بديل\'] %}<button type="button" data-status="{{st}}">{{st}}</button>{% endfor %}</div></td></tr>{% endfor %}</tbody></table></div>\n</main><script>document.querySelectorAll(\'.actions button\').forEach(b=>b.onclick=async()=>{const box=b.closest(\'.actions\');const fd=new FormData();fd.append(\'work_date\',\'{{work_date}}\');fd.append(\'ministry_id\',box.dataset.mid);fd.append(\'status\',b.dataset.status);fd.append(\'notes\',\'\');const r=await fetch(\'/attendance/save\',{method:\'POST\',body:fd});if(r.ok)location.reload();});</script></body></html>'

if not (STATIC_DIR / "style.css").exists():
    (STATIC_DIR / "style.css").write_text(_FALLBACK_STYLE, encoding="utf-8")
if not (TEMPLATES_DIR / "login.html").exists():
    (TEMPLATES_DIR / "login.html").write_text(_FALLBACK_LOGIN, encoding="utf-8")
if not (TEMPLATES_DIR / "dashboard.html").exists():
    (TEMPLATES_DIR / "dashboard.html").write_text(_FALLBACK_DASHBOARD, encoding="utf-8")

serializer = URLSafeSerializer(SECRET_KEY, salt="shalfa-session")


# V18.5 duplicate-school permanent cleanup.
_DUPLICATE_SCHOOLS_CLEANUP_FILE = BASE / "app" / "duplicate_schools_cleanup.json"

def _duplicate_school_ids():
    try:
        payload=json.loads(_DUPLICATE_SCHOOLS_CLEANUP_FILE.read_text(encoding="utf-8"))
        return {
          str(item.get("ministry_id") or "").strip()
          for item in payload.get("schools",[])
          if str(item.get("ministry_id") or "").strip()
        }
    except Exception:
        return set()

DUPLICATE_SCHOOL_IDS=_duplicate_school_ids()

def _safe_delete_by_ministry(con,table,column,ids):
    if not ids:
        return 0
    deleted=0
    try:
        for offset in range(0,len(ids),100):
            chunk=ids[offset:offset+100]
            placeholders=",".join("?" for _ in chunk)
            cur=con.execute(
              f"DELETE FROM {table} WHERE {column} IN ({placeholders})",
              chunk
            )
            deleted+=getattr(cur,"rowcount",0) or 0
    except Exception as exc:
        print(f"V18.5 cleanup skipped {table}: {exc}")
    return deleted



# V18.6 coordinator cleanup master
_COORDINATOR_CLEANUP_FILE = BASE / "app" / "coordinator_cleanup_master.json"

def _load_coordinator_cleanup_master():
    try:
        return json.loads(_COORDINATOR_CLEANUP_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"coordinators":[],"school_links":[]}



# V19.4 — authoritative guard and coordinator replacement from the approved Excel file.
_AUTHORITATIVE_PERSONNEL_FILE = BASE / "app" / "authoritative_personnel_master.json"

def _load_authoritative_personnel():
    try:
        return json.loads(_AUTHORITATIVE_PERSONNEL_FILE.read_text(encoding="utf-8"))
    except Exception as exc:
        print("Authoritative personnel load warning:",exc)
        return {"schools":[],"guards":[],"coordinators":[]}

def _replace_guards_and_coordinators_from_authoritative_file():
    data=_load_authoritative_personnel()
    schools=data.get("schools",[])
    guards=data.get("guards",[])
    coordinators=data.get("coordinators",[])
    if not schools:
        return {"schools":0,"guards":0,"coordinators":0}

    con=db()
    now=datetime.utcnow().isoformat()
    try:
        # Clear current assignments first so every screen reads the same approved data.
        con.execute("""UPDATE schools SET
          guard_name='',guard_emp_id='',guard_id='',guard_mobile='',
          coordinator_name='',coordinator_emp_id='',coordinator_mobile=''""")

        # Delete all old current master records. Historical attendance, violations,
        # documents, incidents and assignment logs are intentionally retained.
        con.execute("DELETE FROM guard_master")
        con.execute("DELETE FROM coordinator_master")

        # Also reset the V19 unified current-state tables; they will be rebuilt below.
        for table in (
          "school_guard_links","school_coordinator_links",
          "master_guards","master_coordinators"
        ):
            try:
                con.execute(f"DELETE FROM {table}")
            except Exception:
                pass

        guard_cols={r["name"] for r in con.execute("PRAGMA table_info(guard_master)").fetchall()}
        for guard in guards:
            emp_id=str(guard.get("emp_id") or "").strip()
            if not emp_id:
                continue
            con.execute("""INSERT INTO guard_master(
              emp_id,emp_name,id_no,mobile,status,linked_ministry_id,linked_at,linked_by
            ) VALUES(?,?,?,?,?,?,?,NULL)""",(
              emp_id,str(guard.get("emp_name") or "").strip(),
              str(guard.get("id_no") or "").strip(),
              str(guard.get("mobile") or "").strip(),
              str(guard.get("status") or "على رأس العمل").strip(),
              str(guard.get("linked_ministry_id") or "").strip() or None,now
            ))
            if "service_status" in guard_cols:
                con.execute("UPDATE guard_master SET service_status='على رأس العمل' WHERE emp_id=?",(emp_id,))
            if "is_active" in guard_cols:
                con.execute("UPDATE guard_master SET is_active=1 WHERE emp_id=?",(emp_id,))

        coord_cols={r["name"] for r in con.execute("PRAGMA table_info(coordinator_master)").fetchall()}
        coordinator_ids={}
        for coord in coordinators:
            emp_id=str(coord.get("emp_id") or "").strip()
            full_name=re.sub(r"\s+"," ",str(coord.get("full_name") or "").strip())
            if not full_name and not emp_id:
                continue
            cur=con.execute("""INSERT INTO coordinator_master(
              emp_id,full_name,mobile,city,is_active,notes,created_by,created_at,updated_at
            ) VALUES(?,?,?,?,1,?,NULL,?,?)""",(
              emp_id or None,full_name,str(coord.get("mobile") or "").strip(),
              str(coord.get("city") or "").strip(),
              "مستورد من الملف المعتمد",now,now
            ))
            coord_id=cur.lastrowid
            if not coord_id:
                row=con.execute(
                  "SELECT id FROM coordinator_master WHERE emp_id=? OR full_name=? ORDER BY id DESC LIMIT 1",
                  (emp_id,full_name)
                ).fetchone()
                coord_id=row["id"] if row else None
            coordinator_ids[emp_id or f"NAME:{full_name}"]=coord_id
            if "sector" in coord_cols:
                con.execute("UPDATE coordinator_master SET sector=? WHERE id=?",
                  (str(coord.get("sector") or "").strip(),coord_id))

        # Reapply all approved school relationships by ministry ID.
        for school in schools:
            ministry_id=str(school.get("ministry_id") or "").strip()
            if not ministry_id:
                continue
            con.execute("""UPDATE schools SET
              guard_name=?,guard_emp_id=?,guard_id=?,guard_mobile=?,
              coordinator_name=?,coordinator_emp_id=?,coordinator_mobile=?
              WHERE ministry_id=?""",(
              str(school.get("guard_name") or "").strip(),
              str(school.get("guard_emp_id") or "").strip(),
              str(school.get("guard_id") or "").strip(),
              str(school.get("guard_mobile") or "").strip(),
              re.sub(r"\s+"," ",str(school.get("coordinator_name") or "").strip()),
              str(school.get("coordinator_emp_id") or "").strip(),
              str(school.get("coordinator_mobile") or "").strip(),
              ministry_id
            ))

        # Keep current personnel_master rows aligned where this table exists.
        try:
            con.execute("DELETE FROM personnel_master WHERE role_title IN ('بواب','منسق / مشرف','منسق','مشرف')")
            for guard in guards:
                con.execute("""INSERT INTO personnel_master(
                  emp_id,full_name,id_no,mobile,role_title,city,sector,
                  linked_ministry_id,is_active
                ) VALUES(?,?,?,?,?,?,?,?,1)
                ON CONFLICT(emp_id) DO UPDATE SET
                  full_name=excluded.full_name,id_no=excluded.id_no,mobile=excluded.mobile,
                  role_title=excluded.role_title,linked_ministry_id=excluded.linked_ministry_id,
                  is_active=1""",(
                  str(guard.get("emp_id") or "").strip(),
                  str(guard.get("emp_name") or "").strip(),
                  str(guard.get("id_no") or "").strip(),
                  str(guard.get("mobile") or "").strip(),
                  "بواب","","",
                  str(guard.get("linked_ministry_id") or "").strip()
                ))
            for coord in coordinators:
                emp_id=str(coord.get("emp_id") or "").strip()
                if not emp_id:
                    continue
                con.execute("""INSERT INTO personnel_master(
                  emp_id,full_name,id_no,mobile,role_title,city,sector,
                  linked_ministry_id,is_active
                ) VALUES(?,?,?,?,?,?,?,?,1)
                ON CONFLICT(emp_id) DO UPDATE SET
                  full_name=excluded.full_name,mobile=excluded.mobile,
                  role_title=excluded.role_title,city=excluded.city,sector=excluded.sector,
                  is_active=1""",(
                  emp_id,re.sub(r"\s+"," ",str(coord.get("full_name") or "").strip()),
                  "",str(coord.get("mobile") or "").strip(),"منسق / مشرف",
                  str(coord.get("city") or "").strip(),str(coord.get("sector") or "").strip(),""
                ))
        except Exception as exc:
            print("Personnel master alignment warning:",exc)

        con.execute("""INSERT INTO app_settings(setting_key,setting_value,updated_at)
          VALUES('authoritative_personnel_source',?,?)
          ON CONFLICT(setting_key) DO UPDATE SET
          setting_value=excluded.setting_value,updated_at=excluded.updated_at""",
          (str(data.get("source_file") or ""),now))
        con.commit()
    except Exception:
        con.rollback();con.close()
        raise
    con.close()

    # Rebuild the unified V19 current-state layer after replacement.
    reconcile_unified_data("authoritative_personnel_replace",emit_event=True)
    return {
      "schools":len(schools),"guards":len(guards),"coordinators":len(coordinators)
    }


# V21 TBC Facilities Management — maintenance and cleaning.
def _init_tbc_fm_schema(con):
    con.executescript("""
    CREATE TABLE IF NOT EXISTS fm_work_orders(
      id INTEGER PRIMARY KEY AUTOINCREMENT,work_order_no TEXT UNIQUE NOT NULL,
      ministry_id TEXT,school_name TEXT,city TEXT,category TEXT NOT NULL,
      subcategory TEXT,title TEXT NOT NULL,description TEXT,priority TEXT DEFAULT 'متوسط',
      source TEXT DEFAULT 'داخلي',status TEXT DEFAULT 'جديد',assigned_to TEXT,
      technician_mobile TEXT,sla_hours INTEGER DEFAULT 24,cost REAL DEFAULT 0,
      materials_cost REAL DEFAULT 0,before_photo TEXT,after_photo TEXT,
      opened_by INTEGER,opened_by_name TEXT,opened_at TEXT NOT NULL,
      accepted_at TEXT,started_at TEXT,resolved_at TEXT,closed_at TEXT,
      resolution_notes TEXT,customer_rating INTEGER,customer_notes TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_fm_wo_status ON fm_work_orders(status,priority,opened_at);
    CREATE INDEX IF NOT EXISTS idx_fm_wo_school ON fm_work_orders(ministry_id,category);

    CREATE TABLE IF NOT EXISTS fm_assets(
      id INTEGER PRIMARY KEY AUTOINCREMENT,asset_code TEXT UNIQUE NOT NULL,
      ministry_id TEXT,school_name TEXT,city TEXT,asset_type TEXT NOT NULL,
      asset_name TEXT,manufacturer TEXT,model TEXT,serial_number TEXT,
      install_date TEXT,warranty_expiry TEXT,condition_status TEXT DEFAULT 'جيدة',
      criticality TEXT DEFAULT 'متوسط',location_description TEXT,qr_code TEXT,
      last_maintenance_date TEXT,next_maintenance_date TEXT,is_active INTEGER DEFAULT 1,
      notes TEXT,created_at TEXT NOT NULL,updated_at TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_fm_assets_school ON fm_assets(ministry_id,asset_type);

    CREATE TABLE IF NOT EXISTS fm_pm_plans(
      id INTEGER PRIMARY KEY AUTOINCREMENT,plan_name TEXT NOT NULL,asset_id INTEGER,
      ministry_id TEXT,asset_type TEXT,frequency_days INTEGER NOT NULL,
      checklist TEXT,assigned_team TEXT,priority TEXT DEFAULT 'متوسط',
      next_due_date TEXT,last_generated_at TEXT,is_active INTEGER DEFAULT 1,
      created_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_fm_pm_due ON fm_pm_plans(next_due_date,is_active);

    CREATE TABLE IF NOT EXISTS fm_cleaning_inspections(
      id INTEGER PRIMARY KEY AUTOINCREMENT,inspection_no TEXT UNIQUE NOT NULL,
      inspection_date TEXT NOT NULL,ministry_id TEXT,school_name TEXT,city TEXT,
      inspector_name TEXT,cleaning_team TEXT,attendance_count INTEGER DEFAULT 0,
      classrooms_score INTEGER DEFAULT 0,toilets_score INTEGER DEFAULT 0,
      corridors_score INTEGER DEFAULT 0,yard_score INTEGER DEFAULT 0,
      offices_score INTEGER DEFAULT 0,overall_score REAL DEFAULT 0,
      status TEXT DEFAULT 'مقبول',before_photo TEXT,after_photo TEXT,
      findings TEXT,corrective_action TEXT,follow_up_date TEXT,
      recorded_by INTEGER,recorded_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_cleaning_date ON fm_cleaning_inspections(inspection_date,city);

    CREATE TABLE IF NOT EXISTS fm_material_usage(
      id INTEGER PRIMARY KEY AUTOINCREMENT,work_order_id INTEGER,material_name TEXT,
      quantity REAL DEFAULT 1,unit TEXT,unit_cost REAL DEFAULT 0,total_cost REAL DEFAULT 0,
      issued_by TEXT,used_at TEXT NOT NULL,notes TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_fm_material_wo ON fm_material_usage(work_order_id);

    CREATE TABLE IF NOT EXISTS fm_employees(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      emp_id TEXT UNIQUE NOT NULL,full_name TEXT NOT NULL,id_no TEXT,mobile TEXT,email TEXT,
      department TEXT DEFAULT 'الصيانة والنظافة',job_title TEXT,trade TEXT,city TEXT,
      team_name TEXT,supervisor_name TEXT,employment_status TEXT DEFAULT 'نشط',
      hire_date TEXT,contract_expiry TEXT,notes TEXT,created_at TEXT NOT NULL,updated_at TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_fm_emp_department ON fm_employees(department,job_title,city);

    CREATE TABLE IF NOT EXISTS fm_inventory_batches(
      id INTEGER PRIMARY KEY AUTOINCREMENT,batch_no TEXT UNIQUE NOT NULL,
      batch_name TEXT,source_files TEXT,total_rows INTEGER DEFAULT 0,
      added_count INTEGER DEFAULT 0,updated_count INTEGER DEFAULT 0,error_count INTEGER DEFAULT 0,
      imported_by INTEGER,imported_by_name TEXT,imported_at TEXT NOT NULL,notes TEXT
    );

    CREATE TABLE IF NOT EXISTS fm_price_catalog(
      id INTEGER PRIMARY KEY AUTOINCREMENT,price_code TEXT UNIQUE NOT NULL,
      asset_type TEXT NOT NULL,asset_name TEXT,manufacturer TEXT,model TEXT,
      condition_status TEXT DEFAULT 'الكل',unit TEXT DEFAULT 'عدد',
      base_price REAL DEFAULT 0,condition_factor REAL DEFAULT 1,
      calculated_unit_price REAL DEFAULT 0,currency TEXT DEFAULT 'SAR',
      effective_from TEXT,effective_to TEXT,is_active INTEGER DEFAULT 1,
      notes TEXT,created_at TEXT NOT NULL,updated_at TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_fm_price_match ON fm_price_catalog(asset_type,asset_name,model,condition_status,is_active);

    CREATE TABLE IF NOT EXISTS fm_asset_price_history(
      id INTEGER PRIMARY KEY AUTOINCREMENT,asset_id INTEGER,price_catalog_id INTEGER,
      old_price REAL,new_price REAL,condition_status TEXT,quantity REAL,
      total_price REAL,pricing_method TEXT,priced_by INTEGER,priced_at TEXT NOT NULL,notes TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_fm_asset_price_history ON fm_asset_price_history(asset_id,priced_at);
    """)


app = FastAPI(title="Shalfa Central Attendance")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
app.mount("/evidence", StaticFiles(directory=EVIDENCE_DIR), name="evidence")
templates = Jinja2Templates(directory=TEMPLATES_DIR)

# V19.0 unified synchronization middleware.
_UNIFIED_MUTATION_PREFIXES=(
  "/schools","/school-import","/unlinked-guards","/coordinators",
  "/attendance","/documents","/system-settings"
)

@app.middleware("http")
async def unified_master_sync_middleware(request: Request,call_next):
    response=await call_next(request)
    if request.method in ("POST","PUT","PATCH","DELETE") and request.url.path.startswith(_UNIFIED_MUTATION_PREFIXES):
        try:
            reconcile_unified_data(f"http:{request.url.path}",emit_event=True)
        except Exception as exc:
            print("Unified reconciliation warning:",exc)
    return response


def db():
    return database_connect()

def hash_password(password: str, salt: Optional[str] = None) -> str:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 120_000).hex()
    return f"{salt}${digest}"

def verify_password(password: str, stored: str) -> bool:
    try:
        salt, digest = stored.split("$", 1)
        chk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 120_000).hex()
        return hmac.compare_digest(chk, digest)
    except Exception:
        return False

def init_db():
    con = db()
    con.executescript("""
    CREATE TABLE IF NOT EXISTS users(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      full_name TEXT NOT NULL,
      role TEXT NOT NULL,
      coordinator_name TEXT,
      is_active INTEGER NOT NULL DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS schools(
      ministry_id TEXT PRIMARY KEY,
      school_name TEXT,
      city TEXT,
      school_type TEXT,
      gender TEXT,
      size TEXT,
      guard_name TEXT,
      guard_emp_id TEXT,
      guard_id TEXT,
      guard_mobile TEXT,
      coordinator_name TEXT,
      coordinator_emp_id TEXT,
      coordinator_mobile TEXT
    );
    CREATE TABLE IF NOT EXISTS attendance(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      work_date TEXT NOT NULL,
      ministry_id TEXT NOT NULL,
      status TEXT NOT NULL,
      replacement_name TEXT,
      replacement_emp_id TEXT,
      notes TEXT,
      source TEXT,
      recorded_by INTEGER,
      recorded_at TEXT NOT NULL,
      UNIQUE(work_date, ministry_id),
      FOREIGN KEY(ministry_id) REFERENCES schools(ministry_id),
      FOREIGN KEY(recorded_by) REFERENCES users(id)
    );
    CREATE INDEX IF NOT EXISTS idx_attendance_date ON attendance(work_date);
    CREATE INDEX IF NOT EXISTS idx_school_city ON schools(city);
    CREATE INDEX IF NOT EXISTS idx_school_coord ON schools(coordinator_name);
    CREATE TABLE IF NOT EXISTS violations(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      violation_date TEXT NOT NULL,
      guard_emp_id TEXT,
      guard_name TEXT,
      ministry_id TEXT,
      school_name TEXT,
      city TEXT,
      coordinator_name TEXT,
      violation_type TEXT NOT NULL,
      severity TEXT,
      action_taken TEXT,
      notes TEXT,
      recorded_by INTEGER,
      recorded_at TEXT NOT NULL,
      FOREIGN KEY(recorded_by) REFERENCES users(id)
    );
    CREATE INDEX IF NOT EXISTS idx_violation_date ON violations(violation_date);
    CREATE INDEX IF NOT EXISTS idx_violation_guard ON violations(guard_emp_id);
    CREATE TABLE IF NOT EXISTS guard_tours(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      tour_date TEXT NOT NULL,
      ministry_id TEXT NOT NULL,
      school_name TEXT,
      city TEXT,
      coordinator_name TEXT,
      checkpoint_code TEXT,
      visit_time TEXT,
      latitude REAL,
      longitude REAL,
      status TEXT NOT NULL DEFAULT 'تمت',
      notes TEXT,
      recorded_by INTEGER,
      recorded_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_tour_date ON guard_tours(tour_date);

    CREATE TABLE IF NOT EXISTS incidents(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      incident_date TEXT NOT NULL,
      incident_time TEXT,
      ministry_id TEXT,
      school_name TEXT,
      city TEXT,
      incident_type TEXT NOT NULL,
      severity TEXT,
      description TEXT,
      action_taken TEXT,
      status TEXT NOT NULL DEFAULT 'مفتوح',
      assigned_to TEXT,
      closed_at TEXT,
      recorded_by INTEGER,
      recorded_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_incident_date ON incidents(incident_date);

    CREATE TABLE IF NOT EXISTS assets(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_type TEXT NOT NULL,
      asset_code TEXT,
      asset_name TEXT,
      serial_number TEXT,
      ministry_id TEXT,
      school_name TEXT,
      assigned_guard_emp_id TEXT,
      assigned_guard_name TEXT,
      quantity INTEGER NOT NULL DEFAULT 1,
      condition_status TEXT NOT NULL DEFAULT 'جيدة',
      received_date TEXT,
      return_date TEXT,
      notes TEXT,
      recorded_by INTEGER,
      recorded_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS shifts(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      work_date TEXT NOT NULL,
      ministry_id TEXT NOT NULL,
      school_name TEXT,
      city TEXT,
      shift_name TEXT NOT NULL,
      required_count INTEGER NOT NULL DEFAULT 1,
      assigned_count INTEGER NOT NULL DEFAULT 0,
      shortage_count INTEGER NOT NULL DEFAULT 0,
      supervisor_name TEXT,
      notes TEXT,
      recorded_by INTEGER,
      recorded_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_shift_date ON shifts(work_date);

    CREATE TABLE IF NOT EXISTS guard_evaluations(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      evaluation_date TEXT NOT NULL,
      guard_emp_id TEXT,
      guard_name TEXT,
      ministry_id TEXT,
      school_name TEXT,
      city TEXT,
      coordinator_name TEXT,
      discipline INTEGER NOT NULL,
      attendance_score INTEGER NOT NULL,
      appearance INTEGER NOT NULL,
      performance INTEGER NOT NULL,
      customer_service INTEGER NOT NULL,
      total_score REAL NOT NULL,
      notes TEXT,
      recorded_by INTEGER,
      recorded_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS employee_documents(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      guard_emp_id TEXT NOT NULL,
      guard_name TEXT,
      document_type TEXT NOT NULL,
      document_number TEXT,
      issue_date TEXT,
      expiry_date TEXT,
      status TEXT,
      notes TEXT,
      recorded_by INTEGER,
      recorded_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_doc_expiry ON employee_documents(expiry_date);
    CREATE TABLE IF NOT EXISTS school_import_batches(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      file_name TEXT,
      imported_by INTEGER,
      imported_at TEXT NOT NULL,
      total_rows INTEGER NOT NULL DEFAULT 0,
      added_count INTEGER NOT NULL DEFAULT 0,
      updated_count INTEGER NOT NULL DEFAULT 0,
      unchanged_count INTEGER NOT NULL DEFAULT 0,
      missing_count INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS school_change_log(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      batch_id INTEGER,
      ministry_id TEXT,
      school_name TEXT,
      change_type TEXT NOT NULL,
      field_name TEXT,
      old_value TEXT,
      new_value TEXT,
      changed_by INTEGER,
      changed_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_school_change_batch ON school_change_log(batch_id);
    CREATE TABLE IF NOT EXISTS user_permissions(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      user_id INTEGER NOT NULL,
      permission_key TEXT NOT NULL,
      is_allowed INTEGER NOT NULL DEFAULT 1,
      UNIQUE(user_id,permission_key),
      FOREIGN KEY(user_id) REFERENCES users(id)
    );
    CREATE INDEX IF NOT EXISTS idx_user_permissions_user ON user_permissions(user_id);
    CREATE TABLE IF NOT EXISTS school_access_permits(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      permit_date TEXT NOT NULL,
      ministry_id TEXT NOT NULL,
      school_name TEXT,
      city TEXT,
      coordinator_name TEXT,
      visit_type TEXT NOT NULL,
      company_name TEXT,
      visitor_names TEXT NOT NULL,
      visitor_ids TEXT,
      visitor_mobiles TEXT,
      vehicle_plate TEXT,
      work_order_no TEXT,
      visit_reason TEXT,
      planned_entry_time TEXT,
      planned_exit_time TEXT,
      approval_status TEXT NOT NULL DEFAULT 'بانتظار الموافقة',
      approved_by_name TEXT,
      approved_by_role TEXT,
      approved_at TEXT,
      entry_time TEXT,
      entry_latitude REAL,
      entry_longitude REAL,
      school_opened_by TEXT,
      exit_time TEXT,
      exit_latitude REAL,
      exit_longitude REAL,
      school_closed_by TEXT,
      actual_status TEXT NOT NULL DEFAULT 'لم يدخل',
      notes TEXT,
      created_by INTEGER,
      created_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_access_permit_date ON school_access_permits(permit_date);
    CREATE INDEX IF NOT EXISTS idx_access_permit_school ON school_access_permits(ministry_id);
    CREATE INDEX IF NOT EXISTS idx_access_permit_status ON school_access_permits(approval_status,actual_status);
    
    CREATE TABLE IF NOT EXISTS coordinator_master(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      emp_id TEXT UNIQUE,
      full_name TEXT NOT NULL,
      mobile TEXT,
      city TEXT,
      is_active INTEGER NOT NULL DEFAULT 1,
      notes TEXT,
      created_by INTEGER,
      created_at TEXT NOT NULL,
      updated_at TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_coordinator_master_active
      ON coordinator_master(is_active,city);

    CREATE TABLE IF NOT EXISTS coordinator_assignment_log(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      ministry_id TEXT NOT NULL,
      school_name TEXT,
      city TEXT,
      old_coordinator_id INTEGER,
      old_coordinator_name TEXT,
      new_coordinator_id INTEGER,
      new_coordinator_name TEXT,
      action_type TEXT NOT NULL,
      changed_by INTEGER,
      changed_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_coord_assignment_school
      ON coordinator_assignment_log(ministry_id,changed_at);
    CREATE TABLE IF NOT EXISTS guard_master(
      emp_id TEXT PRIMARY KEY, emp_name TEXT, id_no TEXT, mobile TEXT, status TEXT,
      linked_ministry_id TEXT, linked_at TEXT, linked_by INTEGER
    );
    CREATE INDEX IF NOT EXISTS idx_guard_master_linked ON guard_master(linked_ministry_id);
    CREATE TABLE IF NOT EXISTS school_operational_exclusions(
      ministry_id TEXT PRIMARY KEY, exclusion_reason TEXT,
      is_excluded INTEGER NOT NULL DEFAULT 1, excluded_at TEXT, excluded_by INTEGER
    );
    CREATE TABLE IF NOT EXISTS app_settings(
      setting_key TEXT PRIMARY KEY,
      setting_value TEXT,
      updated_by INTEGER,
      updated_at TEXT
    );






    """)
    con.commit()

    seed = json.loads((BASE / "seed_data.json").read_text(encoding="utf-8"))
    count = con.execute("SELECT COUNT(*) c FROM schools").fetchone()["c"]
    if count == 0:
        for s in seed["schools"]:
            con.execute("""
            INSERT OR REPLACE INTO schools VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                s.get("ministry_id"), s.get("school_name"), s.get("city"),
                s.get("school_type"), s.get("gender"), s.get("size"),
                s.get("guard_name"), s.get("guard_emp_id"), s.get("guard_id"),
                s.get("guard_mobile"), s.get("coordinator_name"),
                s.get("coordinator_emp_id"), s.get("coordinator_mobile")
            ))
        con.commit()

    if con.execute("SELECT COUNT(*) c FROM users").fetchone()["c"] == 0:
        demo_users = [
            ("admin", "1234", "غرفة العمليات", "admin", None),
            ("manager", "1234", "مدير خدمات أعمال البوابين", "manager", None),
            ("supervisor", "1234", "المشرف العام", "supervisor", None),
            ("client", "1234", "ممثل العميل", "client", None),
        ]
        for u,p,n,r,c in demo_users:
            con.execute("INSERT INTO users(username,password_hash,full_name,role,coordinator_name) VALUES(?,?,?,?,?)",
                        (u, hash_password(p), n, r, c))
        used=set()
        for c in seed["coordinators"]:
            name=(c.get("name") or "").strip()
            if not name or name=="غير محدد" or name in used: continue
            used.add(name)
            username=(c.get("emp_id") or f"coord{len(used)}").strip()
            if not username: username=f"coord{len(used)}"
            try:
                con.execute("INSERT INTO users(username,password_hash,full_name,role,coordinator_name) VALUES(?,?,?,?,?)",
                            (username, hash_password("1234"), name, "coordinator", name))
            except IntegrityError:
                pass
        con.commit()
    con.close()

@app.on_event("startup")
def startup():
    init_db()
    ensure_unified_schema()
    con=db()
    _init_tbc_fm_schema(con)
    # V22 inventory, photos and automatic pricing columns.
    fm_asset_cols={r["name"] for r in con.execute("PRAGMA table_info(fm_assets)").fetchall()}
    for col,col_type in {
      "quantity":"REAL DEFAULT 1","unit":"TEXT DEFAULT 'عدد'",
      "inventory_batch_id":"INTEGER","inventory_date":"TEXT",
      "inventoried_by":"TEXT","photo_path":"TEXT","photo_caption":"TEXT",
      "estimated_unit_price":"REAL DEFAULT 0","estimated_total_price":"REAL DEFAULT 0",
      "price_catalog_id":"INTEGER","pricing_status":"TEXT DEFAULT 'غير مسعر'",
      "pricing_notes":"TEXT","source_file":"TEXT","source_row":"INTEGER"
    }.items():
        if col not in fm_asset_cols:
            con.execute(f"ALTER TABLE fm_assets ADD COLUMN {col} {col_type}")
    # V18.6: normalize coordinators, merge duplicates, deactivate obsolete rows,
    # and re-link schools to the authoritative 15 coordinators from Excel.
    payload=_load_coordinator_cleanup_master()
    master_rows=payload.get("coordinators",[])
    school_links=payload.get("school_links",[])
    if master_rows:
        now=datetime.utcnow().isoformat()
        keep_ids=[]
        key_to_id={}
        for item in master_rows:
            emp_id=str(item.get("emp_id") or "").strip()
            full_name=re.sub(r"\s+"," ",str(item.get("full_name") or "").strip())
            mobile=str(item.get("mobile") or "").strip()
            city=str(item.get("city") or "").strip()
            current=None
            if emp_id:
                current=con.execute(
                  "SELECT * FROM coordinator_master WHERE emp_id=? ORDER BY id LIMIT 1",
                  (emp_id,)
                ).fetchone()
            if not current:
                current=con.execute(
                  "SELECT * FROM coordinator_master WHERE TRIM(full_name)=? ORDER BY id LIMIT 1",
                  (full_name,)
                ).fetchone()
            if current:
                coord_id=current["id"]
                con.execute("""UPDATE coordinator_master SET
                  emp_id=?,full_name=?,mobile=?,city=?,is_active=1,updated_at=?
                  WHERE id=?""",(emp_id or None,full_name,mobile,city,now,coord_id))
            else:
                cur=con.execute("""INSERT INTO coordinator_master(
                  emp_id,full_name,mobile,city,is_active,created_at,updated_at
                ) VALUES(?,?,?,?,1,?,?)""",(emp_id or None,full_name,mobile,city,now,now))
                coord_id=cur.lastrowid
                if not coord_id:
                    coord_id=con.execute(
                      "SELECT id FROM coordinator_master WHERE full_name=? ORDER BY id DESC LIMIT 1",
                      (full_name,)
                    ).fetchone()["id"]
            keep_ids.append(coord_id)
            key_to_id[emp_id or f"NAME:{full_name}"]=coord_id

        # Deactivate all obsolete coordinators instead of deleting history.
        if keep_ids:
            placeholders=",".join("?" for _ in keep_ids)
            con.execute(
              f"UPDATE coordinator_master SET is_active=0,updated_at=? WHERE id NOT IN ({placeholders})",
              [now]+keep_ids
            )

        # Re-link schools using ministry ID to the canonical coordinator row.
        for link in school_links:
            ministry_id=str(link.get("ministry_id") or "").strip()
            key=str(link.get("key") or "").strip()
            if not ministry_id or key not in key_to_id:
                continue
            coord=con.execute(
              "SELECT * FROM coordinator_master WHERE id=?",
              (key_to_id[key],)
            ).fetchone()
            if coord:
                con.execute("""UPDATE schools SET
                  coordinator_name=?,coordinator_emp_id=?,coordinator_mobile=?
                  WHERE ministry_id=?""",(
                    coord["full_name"],coord["emp_id"] or "",coord["mobile"] or "",
                    ministry_id
                ))
        con.commit()
    # V18.5: permanently remove the confirmed duplicate schools by ministry ID.
    if DUPLICATE_SCHOOL_IDS:
        duplicate_ids=sorted(DUPLICATE_SCHOOL_IDS)
        _safe_delete_by_ministry(con,"schools","ministry_id",duplicate_ids)
        _safe_delete_by_ministry(con,"school_operational_exclusions","ministry_id",duplicate_ids)
        _safe_delete_by_ministry(con,"coordinator_assignment_log","ministry_id",duplicate_ids)
        _safe_delete_by_ministry(con,"guard_assignment_log","ministry_id",duplicate_ids)
        _safe_delete_by_ministry(con,"personnel_master","linked_ministry_id",duplicate_ids)
        # Historical attendance, incidents, violations and permits are retained.
        con.commit()
    # V16 master data, school lifecycle and guard documents migrations
    school_cols={r["name"] for r in con.execute("PRAGMA table_info(schools)").fetchall()}
    for col,typ in {
      "status":"TEXT DEFAULT 'تعمل'",
      "merged_into_ministry_id":"TEXT",
      "status_notes":"TEXT",
      "status_updated_at":"TEXT",
      "status_updated_by":"INTEGER",
      "sector":"TEXT",
      "administrative_name":"TEXT",
      "administrative_emp_id":"TEXT",
      "administrative_mobile":"TEXT"
    }.items():
        if col not in school_cols:
            con.execute(f"ALTER TABLE schools ADD COLUMN {col} {typ}")
    # Ensure appointment dates exist even before the next personnel import.
    for col in ("guard_hire_date","coordinator_hire_date"):
        if col not in school_cols:
            con.execute(f"ALTER TABLE schools ADD COLUMN {col} TEXT")
    con.execute("UPDATE schools SET status='تعمل' WHERE status IS NULL OR TRIM(status)=''")
    con.executescript("""
    CREATE TABLE IF NOT EXISTS personnel_master(
      id INTEGER PRIMARY KEY AUTOINCREMENT, emp_id TEXT UNIQUE, full_name TEXT, id_no TEXT,
      mobile TEXT, role_title TEXT, city TEXT, sector TEXT, linked_ministry_id TEXT,
      is_active INTEGER NOT NULL DEFAULT 1, imported_at TEXT, notes TEXT
    );
    CREATE TABLE IF NOT EXISTS master_import_log(
      id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT, imported_at TEXT, imported_by INTEGER,
      schools_count INTEGER DEFAULT 0, guards_count INTEGER DEFAULT 0, coordinators_count INTEGER DEFAULT 0,
      administrators_count INTEGER DEFAULT 0, archived_schools_count INTEGER DEFAULT 0, mode TEXT, notes TEXT
    );
    CREATE TABLE IF NOT EXISTS guard_document_requirements(
      id INTEGER PRIMARY KEY AUTOINCREMENT, guard_emp_id TEXT NOT NULL, document_type TEXT NOT NULL,
      requirement_status TEXT NOT NULL DEFAULT 'مطلوب', reason TEXT, updated_by INTEGER, updated_at TEXT,
      UNIQUE(guard_emp_id,document_type)
    );
    CREATE TABLE IF NOT EXISTS command_center_alert_actions(
      id INTEGER PRIMARY KEY AUTOINCREMENT, alert_key TEXT UNIQUE NOT NULL,
      status TEXT NOT NULL DEFAULT 'مفتوح', assigned_to TEXT, notes TEXT,
      updated_by INTEGER, updated_at TEXT
    );
    """)
    doc_cols={r["name"] for r in con.execute("PRAGMA table_info(employee_documents)").fetchall()}
    for col,typ in {
      "file_path":"TEXT", "original_file_name":"TEXT", "uploaded_at":"TEXT",
      "requirement_status":"TEXT DEFAULT 'مطلوب'"
    }.items():
        if col not in doc_cols:
            con.execute(f"ALTER TABLE employee_documents ADD COLUMN {col} {typ}")
    con.commit()
    cols=[r["name"] for r in con.execute("PRAGMA table_info(schools)").fetchall()]
    if "operational_type" not in cols:
        con.execute("ALTER TABLE schools ADD COLUMN operational_type TEXT")
    # V18.1: normalize existing scope values. Blank or unknown records are never assumed to be main.
    con.execute("""
      UPDATE schools
      SET operational_type=
        CASE
          WHEN COALESCE(operational_type,'') LIKE '%ملحق%' OR COALESCE(school_type,'') LIKE '%ملحق%' THEN 'ملحقة'
          WHEN COALESCE(operational_type,'') LIKE '%فرع%' OR COALESCE(school_type,'') LIKE '%فرع%' THEN 'فرعية'
          WHEN TRIM(COALESCE(operational_type,'')) IN ('رئيسية','رئيسي','أساسية','اساسية','أساسي','اساسي','main','primary') THEN 'رئيسية'
          WHEN TRIM(COALESCE(school_type,'')) IN ('رئيسية','رئيسي','أساسية','اساسية','أساسي','اساسي','main','primary') THEN 'رئيسية'
          WHEN TRIM(COALESCE(operational_type,'')) IN ('رئيسية','فرعية','ملحقة') THEN TRIM(operational_type)
          ELSE 'غير محدد'
        END
    """)
    con.commit()
    # V18.2: coordinator / supervisor import fields.
    coord_cols={r["name"] for r in con.execute("PRAGMA table_info(coordinator_master)").fetchall()}
    for col,typ in {
      "email":"TEXT",
      "sector":"TEXT",
      "imported_at":"TEXT",
      "import_batch":"TEXT"
    }.items():
        if col not in coord_cols:
            con.execute(f"ALTER TABLE coordinator_master ADD COLUMN {col} {typ}")
    con.commit()

    # Initialize default permissions for all users who have none.
    users=con.execute("SELECT id,role FROM users").fetchall()
    for u in users:
        exists=con.execute("SELECT COUNT(*) c FROM user_permissions WHERE user_id=?",(u["id"],)).fetchone()["c"]
        if exists==0:
            for key in ROLE_DEFAULTS.get(u["role"],[]):
                con.execute("INSERT OR IGNORE INTO user_permissions(user_id,permission_key,is_allowed) VALUES(?,?,1)",(u["id"],key))
    con.commit()
    permit_cols=[r["name"] for r in con.execute("PRAGMA table_info(school_access_permits)").fetchall()]
    permit_additions={
      "entry_photo_path":"TEXT","exit_photo_path":"TEXT",
      "entry_server_time":"TEXT","exit_server_time":"TEXT",
      "entry_verified_by":"TEXT","exit_verified_by":"TEXT",
      "entry_evidence_no":"TEXT","exit_evidence_no":"TEXT",
      "entry_evidence_hash":"TEXT","exit_evidence_hash":"TEXT"
    }
    for col,col_type in permit_additions.items():
        if col not in permit_cols:
            con.execute(f"ALTER TABLE school_access_permits ADD COLUMN {col} {col_type}")
    permit_cols=[r["name"] for r in con.execute("PRAGMA table_info(school_access_permits)").fetchall()]
    for col,col_type in {
      "approval_token":"TEXT","exit_token":"TEXT","whatsapp_approval_at":"TEXT","whatsapp_exit_at":"TEXT"
    }.items():
        if col not in permit_cols:
            con.execute(f"ALTER TABLE school_access_permits ADD COLUMN {col} {col_type}")
    default_settings={
      "company_name":"شركة شلفا لإدارة المرافق",
      "system_name":"نظام إدارة العمليات الأمنية وخدمات البوابين",
      "project_name":"مشروع مدارس القصيم",
      "system_version":"V22 TBC FM Inventory & Pricing",
      "gate_services_manager":"عاصم عبدالله سعيد الحصيني الزهراني",
      "gate_services_manager_mobile":"0507675525",
      "general_supervisor":"عبدالله الميموني",
      "general_supervisor_mobile":"0535177867",
      "security_sector_manager":"فضل بن مسعود البيشي",
      "security_sector_manager_mobile":"0567888875",
      "prepared_by":"محمد هندي",
      "prepared_by_mobile":"0507596454"
    }
    for k,v in default_settings.items():
        con.execute("INSERT OR IGNORE INTO app_settings(setting_key,setting_value,updated_at) VALUES(?,?,?)",(k,v,datetime.utcnow().isoformat()))
    # Load corrected guard master and updated school/guard data from bundled Excel.
    corrected_file=STATIC_DIR / "بيانات_المدارس_والبوابين_المصححة.xlsx"
    if corrected_file.exists():
        try:
            wb=load_workbook(corrected_file,data_only=True,read_only=True)
            if "Employee (64)" in wb.sheetnames:
                ews=wb["Employee (64)"]
                headers={str(c.value or "").strip():i+1 for i,c in enumerate(ews[1])}
                for row in ews.iter_rows(min_row=2,values_only=True):
                    def ev(name):
                        idx=headers.get(name); return row[idx-1] if idx else ""
                    if str(ev("Cat_name") or "").strip()=="بواب":
                        emp_id=str(ev("Emp_ID") or "").strip().replace(".0","")
                        if emp_id:
                            con.execute("""INSERT INTO guard_master(emp_id,emp_name,id_no,mobile,status)
                              VALUES(?,?,?,?,?) ON CONFLICT(emp_id) DO UPDATE SET
                              emp_name=excluded.emp_name,id_no=excluded.id_no,
                              mobile=excluded.mobile,status=excluded.status""",
                              (emp_id,str(ev("Emp_name") or "").strip(),
                               str(ev("ID_No") or "").strip().replace(".0",""),
                               str(ev("Mobile") or "").strip().replace(".0",""),
                               str(ev("Status_name") or "").strip()))
            if "جميع المدارس" in wb.sheetnames:
                sws=wb["جميع المدارس"]
                sh={str(c.value or "").strip():i+1 for i,c in enumerate(sws[1])}
                for row in sws.iter_rows(min_row=2,values_only=True):
                    def sv(name):
                        idx=sh.get(name); return row[idx-1] if idx else ""
                    mid=str(sv("الرقم الوزاري") or "").strip().replace(".0","")
                    if not mid: continue
                    geid=str(sv("الرقم الوظيفي للبواب") or "").strip().replace(".0","")
                    con.execute("""UPDATE schools SET guard_name=?,guard_emp_id=?,guard_id=?,guard_mobile=?
                      WHERE ministry_id=?""",
                      (str(sv("اسم البواب") or "").strip(),geid,
                       str(sv("رقم الهوية للبواب") or "").strip().replace(".0",""),
                       str(sv("جوال البواب") or "").strip().replace(".0",""),mid))
                    if geid:
                        con.execute("UPDATE guard_master SET linked_ministry_id=? WHERE emp_id=?",(mid,geid))
        except Exception as exc:
            print("Corrected Excel load warning:",exc)

    # Exclude 235 branch/annex schools from operational counters without deleting records.
    if con.execute("SELECT COUNT(*) c FROM school_operational_exclusions WHERE is_excluded=1").fetchone()["c"]==0:
        candidates=con.execute("""SELECT ministry_id FROM schools
          WHERE COALESCE(operational_type,'') IN ('فرعية','ملحقة')
             OR school_type LIKE '%فرع%' OR school_type LIKE '%ملحق%'
          ORDER BY ministry_id LIMIT 235""").fetchall()
        for s in candidates:
            con.execute("""INSERT OR IGNORE INTO school_operational_exclusions
              (ministry_id,exclusion_reason,is_excluded,excluded_at)
              VALUES(?,?,1,?)""",
              (s["ministry_id"],"مستبعدة من المؤشرات التشغيلية",datetime.utcnow().isoformat()))

    # Sync existing school coordinators into one master table.
    existing_coords=con.execute("""
      SELECT coordinator_emp_id,coordinator_name,coordinator_mobile,city
      FROM schools
      WHERE coordinator_name IS NOT NULL AND TRIM(coordinator_name)<>''
      GROUP BY coordinator_emp_id,coordinator_name,coordinator_mobile,city
    """).fetchall()
    for c in existing_coords:
        emp_id=(c["coordinator_emp_id"] or "").strip()
        name=(c["coordinator_name"] or "").strip()
        mobile=(c["coordinator_mobile"] or "").strip()
        city=(c["city"] or "").strip()
        current=None
        if emp_id:
            current=con.execute("SELECT id FROM coordinator_master WHERE emp_id=?",(emp_id,)).fetchone()
        if not current:
            current=con.execute("SELECT id FROM coordinator_master WHERE full_name=?",(name,)).fetchone()
        if current:
            con.execute("""UPDATE coordinator_master
              SET emp_id=CASE WHEN COALESCE(emp_id,'')='' THEN ? ELSE emp_id END,
                  mobile=CASE WHEN COALESCE(mobile,'')='' THEN ? ELSE mobile END,
                  city=CASE WHEN COALESCE(city,'')='' THEN ? ELSE city END,
                  updated_at=?
              WHERE id=?""",(emp_id or None,mobile,city,datetime.utcnow().isoformat(),current["id"]))
        else:
            try:
                con.execute("""INSERT INTO coordinator_master(
                  emp_id,full_name,mobile,city,is_active,created_at
                ) VALUES(?,?,?,?,1,?)""",(emp_id or None,name,mobile,city,datetime.utcnow().isoformat()))
            except IntegrityError:
                pass

    # Approved short application identity.
    con.execute("""INSERT INTO app_settings(setting_key,setting_value,updated_at)
      VALUES('system_name','Shalfa Ops',?)
      ON CONFLICT(setting_key) DO UPDATE SET setting_value='Shalfa Ops',
      updated_at=excluded.updated_at""",(datetime.utcnow().isoformat(),))
    con.execute("""INSERT INTO app_settings(setting_key,setting_value,updated_at)
      VALUES('system_version','V13 Enterprise',?)
      ON CONFLICT(setting_key) DO UPDATE SET setting_value='V13 Enterprise',
      updated_at=excluded.updated_at""",(datetime.utcnow().isoformat(),))

    # V14 external permit requester fields.
    user_cols=[r["name"] for r in con.execute("PRAGMA table_info(users)").fetchall()]
    for col,col_type in {
      "external_department":"TEXT",
      "external_company":"TEXT",
      "external_mobile":"TEXT"
    }.items():
        if col not in user_cols:
            con.execute(f"ALTER TABLE users ADD COLUMN {col} {col_type}")

    permit_cols=[r["name"] for r in con.execute("PRAGMA table_info(school_access_permits)").fetchall()]
    for col,col_type in {
      "requester_user_id":"INTEGER",
      "requester_name":"TEXT",
      "requester_department":"TEXT",
      "requester_company":"TEXT",
      "requester_mobile":"TEXT",
      "work_summary":"TEXT"
    }.items():
        if col not in permit_cols:
            con.execute(f"ALTER TABLE school_access_permits ADD COLUMN {col} {col_type}")

    # V14.2 guard lifecycle and assignment management.
    guard_cols=[r["name"] for r in con.execute("PRAGMA table_info(guard_master)").fetchall()]
    guard_additions={
      "service_status":"TEXT DEFAULT 'على رأس العمل'",
      "termination_date":"TEXT",
      "termination_reason":"TEXT",
      "updated_at":"TEXT",
      "updated_by":"INTEGER"
    }
    for col,col_type in guard_additions.items():
        if col not in guard_cols:
            con.execute(f"ALTER TABLE guard_master ADD COLUMN {col} {col_type}")
    con.execute("""UPDATE guard_master SET service_status='على رأس العمل'
      WHERE service_status IS NULL OR TRIM(service_status)=''""")
    con.execute("""CREATE TABLE IF NOT EXISTS guard_assignment_log(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      emp_id TEXT NOT NULL,
      emp_name TEXT,
      old_ministry_id TEXT,
      old_school_name TEXT,
      new_ministry_id TEXT,
      new_school_name TEXT,
      action_type TEXT NOT NULL,
      notes TEXT,
      changed_by INTEGER,
      changed_at TEXT NOT NULL
    )""")
    con.execute("""CREATE INDEX IF NOT EXISTS idx_guard_assignment_log_emp
      ON guard_assignment_log(emp_id,changed_at)""")

    # V14.4 load school coordinates extracted from the approved legacy maps.
    school_cols=[r["name"] for r in con.execute("PRAGMA table_info(schools)").fetchall()]
    if "latitude" not in school_cols:
        con.execute("ALTER TABLE schools ADD COLUMN latitude REAL")
    if "longitude" not in school_cols:
        con.execute("ALTER TABLE schools ADD COLUMN longitude REAL")
    coordinates_file=STATIC_DIR / "school_coordinates.json"
    if coordinates_file.exists():
        try:
            coords_data=json.loads(coordinates_file.read_text(encoding="utf-8"))
            for ministry_id,coords in coords_data.items():
                con.execute("""UPDATE schools SET latitude=?,longitude=?
                  WHERE ministry_id=?""",(
                  float(coords.get("latitude")),float(coords.get("longitude")),str(ministry_id)
                ))
        except Exception as exc:
            print("COORDINATE LOAD ERROR:",repr(exc))

    # V14.5 extended school card fields.
    school_cols=[r["name"] for r in con.execute("PRAGMA table_info(schools)").fetchall()]
    school_card_extended_fields={
      "district":"TEXT","stage":"TEXT","classrooms_count":"INTEGER",
      "capacity":"INTEGER","education_admin":"TEXT","education_office":"TEXT",
      "establishment_date":"TEXT","school_status":"TEXT DEFAULT 'نشطة'",
      "school_notes":"TEXT","has_parking":"TEXT","has_playground":"TEXT",
      "has_sports_hall":"TEXT","main_image_url":"TEXT","gallery_images":"TEXT"
    }
    for col,col_type in school_card_extended_fields.items():
        if col not in school_cols:
            con.execute(f"ALTER TABLE schools ADD COLUMN {col} {col_type}")

    # V14.7 client dashboard scope.
    user_cols=[r["name"] for r in con.execute("PRAGMA table_info(users)").fetchall()]
    for col,col_type in {
      "client_city":"TEXT",
      "client_coordinator":"TEXT",
      "client_name":"TEXT"
    }.items():
        if col not in user_cols:
            con.execute(f"ALTER TABLE users ADD COLUMN {col} {col_type}")

    # V15: apply the reviewed school/guard corrections on every new deployment.
    v15_import_file=STATIC_DIR / "V15_تحديث_المدارس_والبوابين.xlsx"
    if v15_import_file.exists():
        try:
            v15_wb=load_workbook(v15_import_file,data_only=True,read_only=True)
            if "استيراد المدارس" in v15_wb.sheetnames:
                v15_ws=v15_wb["استيراد المدارس"]
                v15_headers=[str(c.value or "").strip() for c in v15_ws[1]]
                for values in v15_ws.iter_rows(min_row=2,values_only=True):
                    row=dict(zip(v15_headers,values))
                    def vv(key):
                        value=row.get(key)
                        if value is None:return ""
                        text=str(value).strip()
                        return text[:-2] if text.endswith(".0") and text[:-2].isdigit() else text
                    ministry_id=vv("الرقم الوزاري")
                    if not ministry_id:
                        continue
                    guard_emp_id=vv("الرقم الوظيفي للبواب")
                    guard_name=vv("اسم البواب")
                    guard_id=vv("رقم الهوية للبواب")
                    guard_mobile=vv("جوال البواب")
                    school_name=vv("اسم المدرسة")
                    exists=con.execute("SELECT 1 FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
                    if not exists:
                        continue
                    con.execute("""UPDATE schools SET
                      school_name=CASE WHEN ?<>'' THEN ? ELSE school_name END,
                      guard_emp_id=?,guard_name=?,guard_id=?,guard_mobile=?
                      WHERE ministry_id=?""",(
                      school_name,school_name,guard_emp_id,guard_name,guard_id,guard_mobile,ministry_id
                    ))
                    if guard_emp_id:
                        con.execute("""INSERT INTO guard_master(
                          emp_id,emp_name,id_no,mobile,status,service_status,
                          linked_ministry_id,linked_at
                        ) VALUES(?,?,?,?,?,'على رأس العمل',?,?)
                        ON CONFLICT(emp_id) DO UPDATE SET
                          emp_name=excluded.emp_name,id_no=excluded.id_no,
                          mobile=excluded.mobile,status='نشط',
                          service_status='على رأس العمل',
                          linked_ministry_id=excluded.linked_ministry_id,
                          linked_at=excluded.linked_at""",(
                          guard_emp_id,guard_name,guard_id,guard_mobile,"نشط",
                          ministry_id,datetime.utcnow().isoformat()
                        ))
            v15_wb.close()
        except Exception as exc:
            print("V15 DATA SYNC ERROR:",repr(exc))

    # V15.1: seed the approved attendance snapshot for 24-07-2026.
    attendance_seed=STATIC_DIR / "V15_1_حضور_2026-07-24.xlsx"
    if attendance_seed.exists():
        try:
            awb=load_workbook(attendance_seed,data_only=True,read_only=True)
            if "تفاصيل الحضور" in awb.sheetnames:
                aws=awb["تفاصيل الحضور"]
                ah=[str(c.value or "").strip() for c in aws[1]]
                for vals in aws.iter_rows(min_row=2,values_only=True):
                    rec=dict(zip(ah,vals))
                    def av(key):
                        v=rec.get(key)
                        if v is None:return ""
                        text=str(v).strip()
                        return text[:-2] if text.endswith(".0") and text[:-2].isdigit() else text
                    mid=av("الرقم الوزاري")
                    if not mid or not con.execute("SELECT 1 FROM schools WHERE ministry_id=?",(mid,)).fetchone():
                        continue
                    wd=av("التاريخ") or "2026-07-24"
                    status=av("الحالة") or "لم يسجل"
                    recorded_at=av("وقت التسجيل") or datetime.utcnow().isoformat()
                    con.execute("""INSERT INTO attendance(
                      work_date,ministry_id,status,replacement_name,replacement_emp_id,
                      notes,source,recorded_by,recorded_at
                    ) VALUES(?,?,?,?,?,?,?,?,?)
                    ON CONFLICT(work_date,ministry_id) DO UPDATE SET
                      status=excluded.status,replacement_name=excluded.replacement_name,
                      notes=excluded.notes,source=excluded.source,recorded_at=excluded.recorded_at
                    """,(wd,mid,status,av("البديل"),"",av("الملاحظات"),
                         "excel_v15_1",None,recorded_at))
            awb.close()
        except Exception as exc:
            print("V15.1 ATTENDANCE SEED ERROR:",repr(exc))
    # V20.0_AUDIT_SCHEMA
    con.executescript("""
    CREATE TABLE IF NOT EXISTS audit_trail(
      id INTEGER PRIMARY KEY AUTOINCREMENT,user_id INTEGER,username TEXT,user_role TEXT,
      action TEXT,module TEXT,entity_type TEXT,entity_id TEXT,old_values TEXT,new_values TEXT,
      reason TEXT,request_path TEXT,http_method TEXT,status_code INTEGER,ip_address TEXT,
      user_agent TEXT,success INTEGER DEFAULT 1,created_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_audit_created ON audit_trail(created_at,user_id,module);
    """)
    # V20.1_TASKS_SCHEMA
    con.executescript("""
    CREATE TABLE IF NOT EXISTS tasks(
      id INTEGER PRIMARY KEY AUTOINCREMENT,task_no TEXT UNIQUE,title TEXT NOT NULL,description TEXT,
      task_type TEXT,priority TEXT,ministry_id TEXT,school_name TEXT,city TEXT,assigned_to_user_id INTEGER,
      assigned_to_name TEXT,assigned_to_role TEXT,status TEXT DEFAULT 'جديدة',progress INTEGER DEFAULT 0,
      due_at TEXT,sla_hours INTEGER DEFAULT 24,created_by INTEGER,created_by_name TEXT,created_at TEXT,
      accepted_at TEXT,started_at TEXT,submitted_at TEXT,completed_at TEXT,approval_notes TEXT,
      execution_notes TEXT,latitude REAL,longitude REAL
    );
    CREATE TABLE IF NOT EXISTS task_events(id INTEGER PRIMARY KEY AUTOINCREMENT,task_id INTEGER,event_type TEXT,status_from TEXT,status_to TEXT,notes TEXT,user_id INTEGER,username TEXT,created_at TEXT);
    CREATE INDEX IF NOT EXISTS idx_tasks_status ON tasks(status,assigned_to_user_id,due_at);
    """)
    con.commit()
    con.close()
    reconcile_unified_data("startup",emit_event=False)
    _replace_guards_and_coordinators_from_authoritative_file()


def current_user(request: Request):
    token = request.cookies.get("shalfa_session")
    if not token:
        return None
    try:
        payload = serializer.loads(token)
    except BadSignature:
        return None
    con=db()
    user=con.execute("SELECT * FROM users WHERE id=? AND is_active=1",(payload.get("uid"),)).fetchone()
    con.close()
    return user

def require_user(request: Request):
    user=current_user(request)
    if not user:
        raise HTTPException(status_code=401)
    return user


PERMISSION_CATALOG = {
  "home": "الرئيسية",
  "security_dashboard": "غرفة العمليات الأمنية",
  "map": "الخريطة التفاعلية",
  "attendance": "الحضور اليومي",
  "period_report": "تقارير الحضور",
  "supervisor_dashboard": "مؤشرات المشرفين",
  "guard_tours": "الجولات الأمنية",
  "incidents": "البلاغات والحوادث",
  "violations": "المخالفات والإنذارات",
  "assets": "العهد والمعدات",
  "shifts": "الورديات والعجز",
  "evaluations": "تقييم الحراس",
  "documents": "انتهاء المستندات",
  "schools": "المدارس والبوابون",
  "coordinators": "المنسقون",
  "school_import": "تحديث بيانات المدارس",
  "data_quality": "مركز جودة البيانات",
  "coordinates": "إدارة الإحداثيات",
  "reports": "التقارير",
  "school_reports": "تقارير المدارس الشاملة",
  "relationship_center": "مركز الربط والمزامنة",
  "ask_shalfa_ai": "اسأل شلفا",
  "coordinator_mobile": "تطبيق المنسق",
  "tasks": "نظام المهام",
  "audit_trail": "سجل التدقيق",
  "fm_dashboard": "مركز عمليات الصيانة والنظافة",
  "maintenance_work_orders": "بلاغات وأوامر الصيانة",
  "fm_assets": "أصول ومعدات المدارس",
  "preventive_maintenance": "الصيانة الوقائية",
  "cleaning_quality": "النظافة والجودة",
  "technician_mobile": "تطبيق الفني الميداني",
  "fm_pricing": "تسعير الأصول",
  "fm_inventory": "حصر الأصول",
  "fm_employees": "موظفو الصيانة والنظافة",
  "client_dashboard": "لوحة العميل",
  "users": "إدارة المستخدمين",
  "access_permits": "طلبات تصاريح التشغيل والصيانة",
  "system_settings": "إعدادات البرنامج"
}

ROLE_DEFAULTS = {
  "admin": list(PERMISSION_CATALOG.keys()),
  "manager": [k for k in PERMISSION_CATALOG if k not in ("users",)],
  "supervisor": [
    "home","security_dashboard","map","attendance","period_report",
    "supervisor_dashboard","guard_tours","incidents","violations",
    "assets","shifts","evaluations","documents","schools","coordinators",
    "school_import","reports","client_dashboard","access_permits"
  ,"school_reports","relationship_center","audit_trail","tasks","coordinator_mobile","ask_shalfa_ai","fm_dashboard","maintenance_work_orders","fm_assets","preventive_maintenance","cleaning_quality","technician_mobile","fm_employees","fm_inventory","fm_pricing"],
  "coordinator": [
    "home","map","attendance","period_report","guard_tours","incidents",
    "violations","assets","shifts","evaluations","schools","reports",
    "access_permits"
  ,"school_reports","relationship_center","tasks","coordinator_mobile","ask_shalfa_ai","fm_dashboard","maintenance_work_orders","cleaning_quality","fm_employees","fm_inventory","fm_pricing"],
  "client": ["home","map","period_report","client_dashboard","reports","access_permits","school_reports","fm_dashboard"],
  "external_requester": ["access_permits"]
}

def user_permissions(user):
    if not user:
        return set()
    con=db()
    rows=con.execute("SELECT permission_key,is_allowed FROM user_permissions WHERE user_id=?",(user["id"],)).fetchall()
    con.close()
    explicit={r["permission_key"]:bool(r["is_allowed"]) for r in rows}
    defaults=set(ROLE_DEFAULTS.get(user["role"],[]))
    for key,allowed in explicit.items():
        if allowed: defaults.add(key)
        else: defaults.discard(key)
    if user["role"]=="admin":
        defaults.add("users")
    return defaults

def has_permission(user,key):
    return key in user_permissions(user)

def require_permission(request,key):
    user=current_user(request)
    if not user:
        raise HTTPException(status_code=401)
    if not has_permission(user,key):
        raise HTTPException(status_code=403,detail="ليس لديك صلاحية للوصول إلى هذه الصفحة")
    return user

# Register permission helpers in Jinja only after they are defined.
templates.env.globals["has_permission"] = has_permission
templates.env.globals["PERMISSION_CATALOG"] = PERMISSION_CATALOG


def scope_clause(user):
    clauses=["COALESCE(s.status,'تعمل')='تعمل'"];args=[]
    if user["role"]=="coordinator":
        clauses.append("s.coordinator_name=?");args.append(user["coordinator_name"])
    elif user["role"]=="client":
        if "client_city" in user.keys() and user["client_city"]:
            clauses.append("s.city=?");args.append(user["client_city"])
        if "client_coordinator" in user.keys() and user["client_coordinator"]:
            clauses.append("s.coordinator_name=?");args.append(user["client_coordinator"])
    return (" AND "+" AND ".join(clauses)+" " if clauses else ""),args

# V20.0_AUDIT_MIDDLEWARE
def _audit_write(request,user,status_code):
    try:
        con=db();path=request.url.path;module=(path.strip('/').split('/')[0] or 'home')
        con.execute("""INSERT INTO audit_trail(user_id,username,user_role,action,module,entity_type,entity_id,
          old_values,new_values,reason,request_path,http_method,status_code,ip_address,user_agent,success,created_at)
          VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(
          user['id'] if user else None,user['username'] if user else 'anonymous',user['role'] if user else '',
          'تعديل بيانات',module,'','', '{}','{}','',path,request.method,status_code,
          request.client.host if request.client else '',request.headers.get('user-agent',''),1 if status_code<400 else 0,
          datetime.utcnow().isoformat()))
        con.commit();con.close()
    except Exception as exc: print('Audit warning:',exc)

@app.middleware('http')
async def v20_audit_middleware(request: Request,call_next):
    response=await call_next(request)
    if request.method in ('POST','PUT','PATCH','DELETE') and request.url.path!='/login':
        _audit_write(request,current_user(request),response.status_code)
    return response

@app.get("/", response_class=HTMLResponse)
def root(request: Request):
    user=current_user(request)
    if not user:
        destination="/login"
    elif user["role"]=="client":
        destination="/client-dashboard"
    elif user["role"]=="external_requester":
        destination="/access-permits"
    elif user["role"] in ("admin","manager","supervisor"):
        destination="/security-dashboard"
    else:
        destination="/home"
    return RedirectResponse(destination,status_code=303)


@app.get("/sw.js", include_in_schema=False)
def service_worker():
    response=FileResponse(
      STATIC_DIR / "sw.js",
      media_type="application/javascript",
      headers={
        "Service-Worker-Allowed":"/",
        "Cache-Control":"no-cache, no-store, must-revalidate"
      }
    )
    return response

@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request, error: str=""):
    return templates.TemplateResponse("login.html", {"request":request,"error":error})

@app.post("/login")
def login(request: Request, username: str=Form(...), password: str=Form(...)):
    con=db()
    user=con.execute("SELECT * FROM users WHERE username=? AND is_active=1",(username.strip(),)).fetchone()
    con.close()
    if not user or not verify_password(password,user["password_hash"]):
        return RedirectResponse("/login?error=بيانات الدخول غير صحيحة",status_code=303)
    token=serializer.dumps({"uid":user["id"]})
    destination=("/access-permits" if user["role"]=="external_requester" else
      ("/client-dashboard" if user["role"]=="client" else
       ("/security-dashboard" if user["role"] in ("admin","manager","supervisor") else "/home")))
    resp=RedirectResponse(destination,status_code=303)
    resp.set_cookie("shalfa_session",token,httponly=True,samesite="lax",max_age=60*60*24*30)
    return resp

@app.get("/logout")
def logout():
    resp=RedirectResponse("/login",status_code=303)
    resp.delete_cookie("shalfa_session")
    return resp

@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(
    request: Request,
    work_date: str="",
    city: str="",
    coordinator: str="",
    school_scope: str="all",
    status_filter: str=""
):
    user=require_permission(request,"attendance")
    work_date=work_date or date.today().isoformat()
    con=db()
    clause,args=scope_clause(user)
    filters=""
    if city:
        filters+=" AND s.city=?";args.append(city)
    if coordinator and user["role"]!="coordinator":
        filters+=" AND s.coordinator_name=?";args.append(coordinator)
    if school_scope=="main":
        filters+=" AND TRIM(COALESCE(s.operational_type,''))='رئيسية'"
    elif school_scope=="main_branch":
        filters+=" AND TRIM(COALESCE(s.operational_type,'')) IN ('رئيسية','فرعية')"
    elif school_scope=="branch":
        filters+=" AND TRIM(COALESCE(s.operational_type,''))='فرعية'"
    elif school_scope=="annex":
        filters+=" AND TRIM(COALESCE(s.operational_type,''))='ملحقة'"

    status_sql=""
    status_args=[]
    if status_filter:
        if status_filter=="لم يسجل":
            status_sql=" AND a.status IS NULL"
        else:
            status_sql=" AND a.status=?";status_args.append(status_filter)

    rows=con.execute(f"""
      SELECT s.*,COALESCE(a.status,'لم يسجل') status,a.notes,
             a.replacement_name,a.replacement_emp_id,a.recorded_at,a.source
      FROM schools s LEFT JOIN attendance a
        ON a.ministry_id=s.ministry_id AND a.work_date=?
      WHERE 1=1 {clause} {filters} {status_sql}
      ORDER BY
        CASE COALESCE(a.status,'لم يسجل')
          WHEN 'لم يسجل' THEN 1 WHEN 'غائب' THEN 2 WHEN 'إجازة' THEN 3
          WHEN 'بديل' THEN 4 ELSE 5 END,
        s.city,s.coordinator_name,s.school_name
    """,[work_date]+args+status_args).fetchall()

    summary=con.execute(f"""
      SELECT s.city,COALESCE(NULLIF(TRIM(s.coordinator_name),''),'غير متوفر') coordinator_name,
        COUNT(*) total,
        SUM(CASE WHEN a.status='حاضر' THEN 1 ELSE 0 END) present,
        SUM(CASE WHEN a.status='غائب' THEN 1 ELSE 0 END) absent,
        SUM(CASE WHEN a.status='إجازة' THEN 1 ELSE 0 END) leave_count,
        SUM(CASE WHEN a.status='بديل' THEN 1 ELSE 0 END) replacement,
        SUM(CASE WHEN a.status IS NULL THEN 1 ELSE 0 END) unrecorded,
        MAX(a.recorded_at) last_recorded_at
      FROM schools s LEFT JOIN attendance a
        ON a.ministry_id=s.ministry_id AND a.work_date=?
      WHERE 1=1 {clause} {filters}
      GROUP BY s.city,COALESCE(NULLIF(TRIM(s.coordinator_name),''),'غير متوفر')
      ORDER BY s.city,total DESC
    """,[work_date]+args).fetchall()

    counts={"حاضر":0,"غائب":0,"إجازة":0,"بديل":0,"لم يسجل":0}
    # Counts must use unfiltered population.
    for r in summary:
        counts["حاضر"]+=r["present"] or 0
        counts["غائب"]+=r["absent"] or 0
        counts["إجازة"]+=r["leave_count"] or 0
        counts["بديل"]+=r["replacement"] or 0
        counts["لم يسجل"]+=r["unrecorded"] or 0
    total=sum(counts.values())
    coverage=((counts["حاضر"]+counts["بديل"])/total*100) if total else 0
    recording_pct=((total-counts["لم يسجل"])/total*100) if total else 0
    alerts_count=counts["لم يسجل"]+counts["غائب"]
    cities=[r["city"] for r in con.execute(
      "SELECT DISTINCT city FROM schools WHERE COALESCE(city,'')<>'' ORDER BY city"
    )]
    coords=[r["coordinator_name"] for r in con.execute(
      "SELECT DISTINCT coordinator_name FROM schools WHERE COALESCE(coordinator_name,'')<>'' ORDER BY coordinator_name"
    )]
    con.close()
    summary_data=[]
    for r in summary:
        item=dict(r)
        item["attendance_pct"]=round(((r["present"] or 0)+(r["replacement"] or 0))*100/(r["total"] or 1),1)
        item["recording_pct"]=round(((r["total"] or 0)-(r["unrecorded"] or 0))*100/(r["total"] or 1),1)
        summary_data.append(item)
    return templates.TemplateResponse("dashboard.html",{
      "request":request,"user":user,"rows":rows,"summary":summary_data,
      "counts":counts,"total":total,"coverage":coverage,
      "recording_pct":recording_pct,"alerts_count":alerts_count,
      "work_date":work_date,"cities":cities,"coords":coords,
      "city":city,"coordinator":coordinator,"school_scope":school_scope,
      "status_filter":status_filter
    })

@app.post("/attendance/save")
def save_attendance(
    request: Request,
    work_date: str=Form(...), ministry_id: str=Form(...), status: str=Form(...),
    notes: str=Form(""), replacement_name: str=Form(""), replacement_emp_id: str=Form("")
):
    user=current_user(request)
    if not user: raise HTTPException(status_code=401)
    con=db()
    school=con.execute("SELECT * FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
    if not school: raise HTTPException(404)
    if user["role"]=="coordinator" and school["coordinator_name"]!=user["coordinator_name"]:
        raise HTTPException(403)
    con.execute("""
    INSERT INTO attendance(work_date,ministry_id,status,replacement_name,replacement_emp_id,notes,source,recorded_by,recorded_at)
    VALUES(?,?,?,?,?,?,?,?,?)
    ON CONFLICT(work_date,ministry_id) DO UPDATE SET
      status=excluded.status,replacement_name=excluded.replacement_name,
      replacement_emp_id=excluded.replacement_emp_id,notes=excluded.notes,
      source=excluded.source,recorded_by=excluded.recorded_by,recorded_at=excluded.recorded_at
    """,(work_date,ministry_id,status,replacement_name,replacement_emp_id,notes,"web",user["id"],datetime.utcnow().isoformat()))
    con.commit(); con.close()
    return JSONResponse({"ok":True})

@app.post("/attendance/import")
async def import_attendance(request: Request, file: UploadFile=File(...), work_date: str=Form(...)):
    user=current_user(request)
    if not user: raise HTTPException(401)
    content=await file.read()
    wb=load_workbook(BytesIO(content),data_only=True)
    con=db(); matched=0
    for ws in wb.worksheets:
        headers=[str(c.value or "").strip() for c in ws[1]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            rec=dict(zip(headers,row))
            mid=str(rec.get("الرقم الوزاري") or "").strip().replace(".0","")
            emp=str(rec.get("الرقم الوظيفي للبواب") or "").strip().replace(".0","")
            school=None
            if mid:
                school=con.execute("SELECT * FROM schools WHERE ministry_id=?",(mid,)).fetchone()
            if not school and emp:
                school=con.execute("SELECT * FROM schools WHERE guard_emp_id=?",(emp,)).fetchone()
            if not school: continue
            if user["role"]=="coordinator" and school["coordinator_name"]!=user["coordinator_name"]:
                continue
            status=str(rec.get("الحالة") or "لم يسجل").strip()
            d=str(rec.get("التاريخ") or work_date).strip()
            con.execute("""
            INSERT INTO attendance(work_date,ministry_id,status,replacement_name,replacement_emp_id,notes,source,recorded_by,recorded_at)
            VALUES(?,?,?,?,?,?,?,?,?)
            ON CONFLICT(work_date,ministry_id) DO UPDATE SET
              status=excluded.status,replacement_name=excluded.replacement_name,
              replacement_emp_id=excluded.replacement_emp_id,notes=excluded.notes,
              source=excluded.source,recorded_by=excluded.recorded_by,recorded_at=excluded.recorded_at
            """,(d,school["ministry_id"],status,str(rec.get("اسم البديل") or rec.get("البديل") or ""),str(rec.get("الرقم الوظيفي للبديل") or ""),
                 str(rec.get("ملاحظات") or ""),"excel",user["id"],
                 str(rec.get("وقت التسجيل") or datetime.utcnow().isoformat())))
            matched+=1
    con.commit();con.close()
    return RedirectResponse(f"/dashboard?work_date={work_date}&imported={matched}",status_code=303)

@app.get("/export.xlsx")
def export_excel(request: Request, work_date: str="", city: str="", coordinator: str="", school_scope: str="all"):
    user=current_user(request)
    if not user: raise HTTPException(401)
    work_date=work_date or date.today().isoformat()
    con=db()
    clause,args=scope_clause(user)
    filters=""
    if city: filters+=" AND s.city=?";args.append(city)
    if coordinator and user["role"]!="coordinator": filters+=" AND s.coordinator_name=?";args.append(coordinator)
    if school_scope=="main": filters+=" AND TRIM(COALESCE(s.operational_type,''))='رئيسية'"
    elif school_scope=="main_branch": filters+=" AND TRIM(COALESCE(s.operational_type,'')) IN ('رئيسية','فرعية')"
    elif school_scope=="branch": filters+=" AND TRIM(COALESCE(s.operational_type,''))='فرعية'"
    elif school_scope=="annex": filters+=" AND TRIM(COALESCE(s.operational_type,''))='ملحقة'"
    rows=con.execute(f"""
      SELECT s.*,COALESCE(a.status,'لم يسجل') status,a.notes,a.replacement_name,a.replacement_emp_id,a.recorded_at
      FROM schools s LEFT JOIN attendance a ON a.ministry_id=s.ministry_id AND a.work_date=?
      WHERE 1=1 {clause} {filters} ORDER BY s.city,s.school_name
    """,[work_date]+args).fetchall()
    con.close()
    wb=Workbook();ws=wb.active;ws.title="الحضور"
    headers=["التاريخ","الرقم الوزاري","اسم المدرسة","المدينة","اسم المنسق","اسم البواب","الرقم الوظيفي","جوال البواب","الحالة","اسم البديل","ملاحظات","وقت التسجيل"]
    ws.append(headers)
    for r in rows:
        ws.append([work_date,r["ministry_id"],r["school_name"],r["city"],r["coordinator_name"],r["guard_name"],r["guard_emp_id"],r["guard_mobile"],r["status"],r["replacement_name"],r["notes"],r["recorded_at"]])
    for cell in ws[1]:
        cell.font=cell.font.copy(bold=True,color="FFFFFF");cell.fill=cell.fill.copy(fill_type="solid",fgColor="082B60")
    for col in ["A","B","G","H"]:
        for cell in ws[col]: cell.number_format="@"
    bio=BytesIO();wb.save(bio);bio.seek(0)
    filename=f"attendance_{work_date}.xlsx"
    return StreamingResponse(bio,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      headers={"Content-Disposition":f'attachment; filename="{filename}"'})

@app.get("/api/attendance")
def api_attendance(request: Request, work_date: str="", city: str=""):
    user=current_user(request)
    if not user: raise HTTPException(401)
    work_date=work_date or date.today().isoformat()
    con=db();clause,args=scope_clause(user)
    if city: clause+=" AND s.city=?";args.append(city)
    rows=con.execute(f"""
      SELECT s.ministry_id,s.school_name,s.city,s.guard_name,s.guard_emp_id,
             s.coordinator_name,COALESCE(a.status,'لم يسجل') status,a.notes,a.recorded_at
      FROM schools s LEFT JOIN attendance a ON a.ministry_id=s.ministry_id AND a.work_date=?
      WHERE 1=1 {clause}
    """,[work_date]+args).fetchall()
    con.close()
    return [dict(r) for r in rows]


@app.get("/home", response_class=HTMLResponse)
def home_page(request: Request):
    user=require_permission(request,"home")
    con=db()
    today=date.today().isoformat()
    totals=con.execute("""
      SELECT
        SUM(CASE WHEN e.ministry_id IS NULL THEN 1 ELSE 0 END) total,
        SUM(CASE WHEN e.ministry_id IS NULL AND s.school_type LIKE '%رئيس%' THEN 1 ELSE 0 END) main_count,
        SUM(CASE WHEN e.ministry_id IS NULL AND s.school_type LIKE '%فرع%' THEN 1 ELSE 0 END) branch_count,
        COUNT(DISTINCT CASE WHEN e.ministry_id IS NULL THEN s.city END) cities,
        COUNT(DISTINCT CASE WHEN e.ministry_id IS NULL THEN s.coordinator_name END) coordinators,
        COUNT(DISTINCT CASE WHEN e.ministry_id IS NULL AND s.guard_emp_id<>'' THEN s.guard_emp_id END) guards
      FROM schools s
      LEFT JOIN school_operational_exclusions e ON e.ministry_id=s.ministry_id AND e.is_excluded=1
    """).fetchone()
    daily=con.execute("""
      SELECT
        SUM(CASE WHEN status='حاضر' THEN 1 ELSE 0 END) present_count,
        SUM(CASE WHEN status='غائب' THEN 1 ELSE 0 END) absent_count,
        SUM(CASE WHEN status='إجازة' THEN 1 ELSE 0 END) leave_count,
        SUM(CASE WHEN status='بديل' THEN 1 ELSE 0 END) replacement_count,
        COUNT(*) recorded_count
      FROM attendance WHERE work_date=?
    """,(today,)).fetchone()
    operations=con.execute("""
      SELECT
        (SELECT COUNT(*) FROM guard_tours WHERE tour_date=?) tours_today,
        (SELECT COUNT(*) FROM violations WHERE violation_date=?) violations_today,
        (SELECT COUNT(*) FROM incidents WHERE status<>'مغلق') open_incidents,
        (SELECT COUNT(*) FROM school_access_permits WHERE actual_status='داخل المدرسة') teams_inside,
        (SELECT COUNT(*) FROM school_access_permits WHERE approval_status='بانتظار الموافقة') pending_permits
    """,(today,today)).fetchone()
    active_coordinators=con.execute(
      "SELECT COUNT(*) c FROM coordinator_master WHERE is_active=1"
    ).fetchone()["c"]
    total_schools=totals["total"] or 0
    present=daily["present_count"] or 0
    replacement=daily["replacement_count"] or 0
    coverage=((present+replacement)/total_schools*100) if total_schools else 0
    unrecorded=max(total_schools-(daily["recorded_count"] or 0),0)
    guard_stats=con.execute("""
      SELECT COUNT(*) total_guards,
       SUM(CASE WHEN linked_ministry_id IS NOT NULL AND linked_ministry_id<>'' THEN 1 ELSE 0 END) linked_guards,
       SUM(CASE WHEN linked_ministry_id IS NULL OR linked_ministry_id='' THEN 1 ELSE 0 END) unlinked_guards
      FROM guard_master
    """).fetchone()
    con.close()
    return templates.TemplateResponse("home.html",{
      "request":request,"user":user,"totals":totals,"daily":daily,
      "operations":operations,"coverage":coverage,"unrecorded":unrecorded,
      "today":today,"settings":get_app_settings(),"guard_stats":guard_stats,"active_coordinators":active_coordinators
    })


@app.get("/api/map-schools")
def map_schools_api(request: Request):
    user=require_permission(request,"map")
    con=db()
    try:
        where=[
          "latitude IS NOT NULL",
          "longitude IS NOT NULL"
        ]
        args=[]
        if user["role"]=="coordinator":
            where.append("coordinator_name=?")
            args.append(user["coordinator_name"] or "")
        rows=con.execute("""
          SELECT ministry_id,school_name,city,school_type,operational_type,
            gender,size,guard_name,guard_emp_id,guard_mobile,
            coordinator_name,coordinator_emp_id,coordinator_mobile,
            guard_hire_date,coordinator_hire_date,
            latitude,longitude,district,stage,classrooms_count,capacity,
            education_admin,education_office,building_ownership,unit_classification,establishment_date,school_status,
            school_notes,has_parking,has_playground,has_sports_hall,
            main_image_url,gallery_images,
            (SELECT a.status FROM attendance a WHERE a.ministry_id=schools.ministry_id
             ORDER BY a.work_date DESC,a.recorded_at DESC LIMIT 1) latest_attendance_status,
            (SELECT a.work_date FROM attendance a WHERE a.ministry_id=schools.ministry_id
             ORDER BY a.work_date DESC,a.recorded_at DESC LIMIT 1) latest_attendance_date,
            (SELECT a.recorded_at FROM attendance a WHERE a.ministry_id=schools.ministry_id
             ORDER BY a.work_date DESC,a.recorded_at DESC LIMIT 1) latest_attendance_time
          FROM schools
          WHERE """+" AND ".join(where)+"""
          ORDER BY city,school_name
        """,args).fetchall()
        return JSONResponse([{
          "ministry_id":r["ministry_id"] or "",
          "school_name":r["school_name"] or "",
          "city":r["city"] or "",
          "school_type":r["school_type"] or "",
          "operational_type":r["operational_type"] or "",
          "gender":r["gender"] or "",
          "size":r["size"] or "",
          "guard_name":r["guard_name"] or "",
          "guard_emp_id":r["guard_emp_id"] or "",
          "guard_mobile":r["guard_mobile"] or "",
          "guard_hire_date":r["guard_hire_date"] or "",
          "coordinator_name":r["coordinator_name"] or "",
          "coordinator_emp_id":r["coordinator_emp_id"] or "",
          "coordinator_mobile":r["coordinator_mobile"] or "",
          "coordinator_hire_date":r["coordinator_hire_date"] or "",
          "latitude":float(r["latitude"]),
          "longitude":float(r["longitude"]),
          "district":r["district"] or "",
          "stage":r["stage"] or "",
          "classrooms_count":r["classrooms_count"] or "",
          "capacity":r["capacity"] or "",
          "education_admin":r["education_admin"] or "",
          "education_office":r["education_office"] or "",
          "building_ownership":r["building_ownership"] or "",
          "unit_classification":r["unit_classification"] or "",
          "establishment_date":r["establishment_date"] or "",
          "school_status":r["school_status"] or "نشطة",
          "school_notes":r["school_notes"] or "",
          "has_parking":r["has_parking"] or "",
          "has_playground":r["has_playground"] or "",
          "has_sports_hall":r["has_sports_hall"] or "",
          "main_image_url":r["main_image_url"] or "",
          "gallery_images":r["gallery_images"] or "",
          "attendance_status":r["latest_attendance_status"] or "لم يسجل",
          "attendance_date":r["latest_attendance_date"] or "",
          "attendance_time":r["latest_attendance_time"] or ""
        } for r in rows])
    except Exception as exc:
        print("MAP API ERROR:",repr(exc))
        return JSONResponse(
          {"error":"تعذر تحميل بيانات الخريطة","details":str(exc)},
          status_code=500
        )
    finally:
        con.close()


@app.get("/schools/{ministry_id}/card.pdf", response_class=HTMLResponse)
def school_card_print(request: Request,ministry_id: str):
    user=require_permission(request,"map")
    con=db()
    school=con.execute("SELECT * FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
    if not school:
        con.close();raise HTTPException(404,detail="المدرسة غير موجودة")
    guard_log=con.execute("""SELECT l.*,u.full_name changed_by_name
      FROM guard_assignment_log l
      LEFT JOIN users u ON u.id=l.changed_by
      WHERE l.new_ministry_id=? OR l.old_ministry_id=?
      ORDER BY l.id DESC LIMIT 1""",(ministry_id,ministry_id)).fetchone()
    con.close()
    response=templates.TemplateResponse("school_card_print.html",{
      "request":request,"user":user,"school":school,"guard_log":guard_log
    })
    response.headers["Cache-Control"]="no-store"
    return response

@app.get("/map", response_class=HTMLResponse)
def map_page(request: Request):
    user=require_permission(request,"map")
    response=templates.TemplateResponse("map.html",{"request":request,"user":user})
    response.headers["Cache-Control"]="no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"]="no-cache"
    response.headers["Expires"]="0"
    return response

@app.get("/attendance", response_class=HTMLResponse)
def attendance_page(request: Request, work_date: str="", city: str="", coordinator: str="", school_scope: str="auto"):
    return dashboard(request,work_date,city,coordinator,school_scope)

@app.get("/schools", response_class=HTMLResponse)
def schools_page(request: Request, q: str="", city: str="", coordinator: str="", school_type: str="all", status: str="all"):
    user=require_permission(request,"schools")
    con=db();where=["1=1"];args=[]
    if user["role"]=="coordinator":
        where.append("coordinator_name=?");args.append(user["coordinator_name"])
    if q:
        where.append("(school_name LIKE ? OR ministry_id LIKE ? OR guard_name LIKE ? OR guard_emp_id LIKE ?)")
        like=f"%{q}%";args += [like,like,like,like]
    if city: where.append("city=?");args.append(city)
    if status and status!="all": where.append("COALESCE(status,'تعمل')=?");args.append(status)
    if coordinator and user["role"]!="coordinator":
        where.append("coordinator_name=?");args.append(coordinator)
    if school_type=="main":
        where.append("TRIM(COALESCE(operational_type,''))='رئيسية'")
    elif school_type=="main_branch":
        where.append("TRIM(COALESCE(operational_type,'')) IN ('رئيسية','فرعية')")
    elif school_type=="branch":
        where.append("TRIM(COALESCE(operational_type,''))='فرعية'")
    elif school_type=="annex":
        where.append("TRIM(COALESCE(operational_type,''))='ملحقة'")
    rows=con.execute("SELECT * FROM schools WHERE "+" AND ".join(where)+" ORDER BY city,school_name",args).fetchall()
    cities=[r["city"] for r in con.execute("SELECT DISTINCT city FROM schools WHERE city<>'' ORDER BY city")]
    coords=[r["coordinator_name"] for r in con.execute("SELECT DISTINCT coordinator_name FROM schools WHERE coordinator_name<>'' ORDER BY coordinator_name")]
    con.close()
    return templates.TemplateResponse("schools.html",{"request":request,"user":user,"rows":rows,"q":q,"city":city,"coordinator":coordinator,"school_type":school_type,"status":status,"cities":cities,"coords":coords})

@app.post("/schools/status")
def school_status_update(request: Request, ministry_id: str=Form(...), status: str=Form(...), merged_into_ministry_id: str=Form(""), status_notes: str=Form("")):
    user=current_user(request)
    if not user or user["role"] not in ("admin","manager","supervisor"): raise HTTPException(403)
    if status not in ("تعمل","مغلقة","مدمجة"): raise HTTPException(400)
    if status=="مدمجة" and (not merged_into_ministry_id or merged_into_ministry_id==ministry_id): raise HTTPException(400,detail="اختر مدرسة دمج صحيحة")
    con=db();con.execute("UPDATE schools SET status=?,merged_into_ministry_id=?,status_notes=?,status_updated_at=?,status_updated_by=? WHERE ministry_id=?",
      (status,merged_into_ministry_id if status=="مدمجة" else "",status_notes,datetime.utcnow().isoformat(),user["id"],ministry_id))
    con.commit();con.close();return RedirectResponse("/schools",303)

@app.get("/coordinators", response_class=HTMLResponse)
def coordinators_page(
    request: Request,
    city: str="",
    current_coordinator: str="",
    school_type: str="all",
    q: str=""
):
    user=require_permission(request,"coordinators")
    con=db()
    coordinators=con.execute("""
      SELECT c.*,
        COUNT(s.ministry_id) total,
        SUM(CASE WHEN s.school_type LIKE '%رئيس%' THEN 1 ELSE 0 END) main_count,
        SUM(CASE WHEN s.school_type LIKE '%فرع%' THEN 1 ELSE 0 END) branch_count
      FROM coordinator_master c
      LEFT JOIN schools s ON s.coordinator_name=c.full_name
      GROUP BY c.id
      ORDER BY c.is_active DESC,total DESC,c.full_name
    """).fetchall()

    where=["1=1"];args=[]
    if city:
        where.append("s.city=?");args.append(city)
    if current_coordinator=="unassigned":
        where.append("(s.coordinator_name IS NULL OR TRIM(s.coordinator_name)='')")
    elif current_coordinator:
        where.append("s.coordinator_name=?");args.append(current_coordinator)
    if school_type=="main":
        where.append("s.school_type LIKE '%رئيس%'")
    elif school_type=="branch":
        where.append("s.school_type LIKE '%فرع%'")
    elif school_type=="annex":
        where.append("COALESCE(s.operational_type,'')='ملحقة'")
    if q:
        where.append("(s.school_name LIKE ? OR s.ministry_id LIKE ?)")
        like=f"%{q}%";args.extend([like,like])

    schools=con.execute("""
      SELECT s.ministry_id,s.school_name,s.city,s.school_type,
             s.coordinator_name,s.coordinator_emp_id,s.coordinator_mobile,
             s.guard_name,s.guard_emp_id
      FROM schools s
      WHERE """+" AND ".join(where)+"""
      ORDER BY s.city,s.school_name
      LIMIT 2500
    """,args).fetchall()
    cities=[r["city"] for r in con.execute(
      "SELECT DISTINCT city FROM schools WHERE COALESCE(city,'')<>'' ORDER BY city"
    )]
    unassigned=con.execute("""
      SELECT COUNT(*) c FROM schools
      WHERE coordinator_name IS NULL OR TRIM(coordinator_name)=''
    """).fetchone()["c"]
    con.close()
    return templates.TemplateResponse("coordinators.html",{
      "request":request,"user":user,"rows":coordinators,"schools":schools,
      "cities":cities,"city":city,"current_coordinator":current_coordinator,
      "school_type":school_type,"q":q,"unassigned":unassigned
    })


def _coordinator_status(value):
    text=(value or "").strip().lower()
    if text in ("0","موقوف","غير نشط","متوقف","inactive","disabled","لا"):
        return 0
    return 1


def _normalize_coordinator_header(value):
    text=str(value or "").replace("\xa0"," ").strip()
    text=re.sub(r"\s+"," ",text)
    text=text.replace("／","/").replace("\\","/")
    return text

def _coordinator_text(value):
    if value is None:
        return ""
    text=str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text

def _coordinator_header_map(headers):
    exact={}
    compact={}
    for original in headers:
        normalized=_normalize_coordinator_header(original)
        if not normalized:
            continue
        exact[normalized]=original
        compact[re.sub(r"[\s/_-]+","",normalized)]=original
    return exact,compact

def _coordinator_find_header(headers,*aliases):
    exact,compact=_coordinator_header_map(headers)
    for alias in aliases:
        normalized=_normalize_coordinator_header(alias)
        if normalized in exact:
            return exact[normalized]
        key=re.sub(r"[\s/_-]+","",normalized)
        if key in compact:
            return compact[key]
    return None

def _coordinator_import_rows(workbook):
    rows=[]
    for ws in workbook.worksheets:
        headers=[_normalize_coordinator_header(c.value) for c in ws[1]]
        name_col=_coordinator_find_header(
          headers,
          "المنسق / المشرف","اسم المنسق / المشرف","اسم المنسق/المشرف",
          "اسم المنسق","اسم المشرف","المنسق","المشرف","الاسم","full_name","name"
        )
        if not name_col:
            continue
        emp_col=_coordinator_find_header(
          headers,
          "الرقم الوظيفي للمنسق","الرقم الوظيفي للمنسق / المشرف",
          "الرقم الوظيفي","رقم المنسق","رقم المشرف","emp_id"
        )
        mobile_col=_coordinator_find_header(
          headers,
          "جوال المنسق","جوال المنسق / المشرف","جوال المشرف",
          "رقم جوال المنسق","الجوال","رقم الجوال","mobile"
        )
        email_col=_coordinator_find_header(headers,"البريد الإلكتروني","البريد","الايميل","email")
        city_col=_coordinator_find_header(headers,"المدينة","المحافظة","city")
        sector_col=_coordinator_find_header(headers,"القطاع","النطاق","sector")
        status_col=_coordinator_find_header(headers,"الحالة","حالة المنسق","حالة المنسق / المشرف","status")
        notes_col=_coordinator_find_header(headers,"ملاحظات","الملاحظات","notes")
        ministry_col=_coordinator_find_header(headers,"الرقم الوزاري","رقم وزاري","رقم المدرسة","ministry_id")

        for values in ws.iter_rows(min_row=2,values_only=True):
            row=dict(zip(headers,values))
            name=_coordinator_text(row.get(name_col))
            if not name:
                continue
            rows.append({
              "full_name":name,
              "emp_id":_coordinator_text(row.get(emp_col)) if emp_col else "",
              "mobile":_coordinator_text(row.get(mobile_col)) if mobile_col else "",
              "email":_coordinator_text(row.get(email_col)) if email_col else "",
              "city":_coordinator_text(row.get(city_col)) if city_col else "",
              "sector":_coordinator_text(row.get(sector_col)) if sector_col else "",
              "status":(_coordinator_text(row.get(status_col)) if status_col else "") or "نشط",
              "notes":_coordinator_text(row.get(notes_col)) if notes_col else "",
              "ministry_id":(
                _coordinator_text(row.get(ministry_col))
                if ministry_col and _coordinator_text(row.get(ministry_col)) not in DUPLICATE_SCHOOL_IDS
                else ""
              )
            })
    return rows


@app.get("/coordinators/import-template.xlsx")
def coordinators_import_template(request: Request):
    require_permission(request,"coordinators")
    wb=Workbook();ws=wb.active;ws.title="المنسقون والمدارس"
    ws.append(["الرقم الوظيفي","اسم المنسق / المشرف","الجوال","البريد الإلكتروني",
               "المدينة","القطاع","الحالة","الرقم الوزاري","ملاحظات"])
    ws.append(["1001","مثال: أحمد محمد","05xxxxxxxx","name@example.com",
               "بريدة","القطاع الأول","نشط","51001",""])
    ws.append(["1001","مثال: أحمد محمد","05xxxxxxxx","name@example.com",
               "بريدة","القطاع الأول","نشط","51002","يمكن تكرار المنسق لكل مدرسة"])
    for cell in ws[1]:
        cell.font=cell.font.copy(bold=True,color="FFFFFF")
        cell.fill=cell.fill.copy(fill_type="solid",fgColor="082B60")
    widths={"A":18,"B":34,"C":18,"D":28,"E":18,"F":20,"G":14,"H":18,"I":38}
    for col,width in widths.items(): ws.column_dimensions[col].width=width
    ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
    bio=BytesIO();wb.save(bio);bio.seek(0)
    return StreamingResponse(bio,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      headers={"Content-Disposition":'attachment; filename="coordinators_import_template.xlsx"'})


@app.get("/coordinators/master.xlsx")
def coordinators_master_export(request: Request):
    require_permission(request,"coordinators");con=db()
    rows=con.execute("""SELECT c.*,s.ministry_id,s.school_name FROM coordinator_master c
      LEFT JOIN schools s ON s.coordinator_name=c.full_name
      ORDER BY c.is_active DESC,c.full_name,s.city,s.school_name""").fetchall();con.close()
    wb=Workbook();ws=wb.active;ws.title="المنسقون والمدارس"
    ws.append(["الرقم الوظيفي","اسم المنسق / المشرف","الجوال","البريد الإلكتروني",
               "المدينة","القطاع","الحالة","الرقم الوزاري","اسم المدرسة","ملاحظات"])
    for r in rows:
        ws.append([r["emp_id"],r["full_name"],r["mobile"],r["email"],r["city"],r["sector"],
                   "نشط" if r["is_active"] else "موقوف",r["ministry_id"],r["school_name"],r["notes"]])
    for cell in ws[1]:
        cell.font=cell.font.copy(bold=True,color="FFFFFF")
        cell.fill=cell.fill.copy(fill_type="solid",fgColor="082B60")
    ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width=min(max(len(str(c.value or "")) for c in col)+3,42)
    bio=BytesIO();wb.save(bio);bio.seek(0)
    return StreamingResponse(bio,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      headers={"Content-Disposition":'attachment; filename="coordinators_master.xlsx"'})


@app.post("/coordinators/import")
async def coordinators_import(
    request: Request,
    file: UploadFile=File(...),
    update_school_assignments: str=Form("yes"),
    deactivate_missing: str=Form("")
):
    user=require_permission(request,"coordinators")
    if user["role"] not in ("admin","manager","supervisor"):
        raise HTTPException(403)
    content=await file.read()
    try:
        workbook=load_workbook(BytesIO(content),data_only=True,read_only=True)
    except Exception as exc:
        raise HTTPException(400,detail=f"تعذر قراءة ملف Excel: {exc}")
    imported_rows=_coordinator_import_rows(workbook)
    if not imported_rows:
        read_columns=[]
        for ws in workbook.worksheets:
            read_columns.extend([
              _normalize_coordinator_header(c.value)
              for c in ws[1] if c.value not in (None,"")
            ])
        workbook.close()
        raise HTTPException(400,detail={
          "message":"لم يتم العثور على عمود اسم المنسق / المشرف داخل الملف",
          "columns_read":read_columns,
          "accepted_columns":[
            "المنسق / المشرف","اسم المنسق / المشرف","اسم المنسق",
            "اسم المشرف","المنسق","المشرف"
          ]
        })
    workbook.close()

    con=db();now=datetime.utcnow().isoformat();batch=datetime.utcnow().strftime("COORD-%Y%m%d%H%M%S")
    added=updated=assigned=unchanged=duplicates=missing_schools=deactivated=0
    imported_keys=set();seen_rows=set();processed_coords={}
    try:
        for row in imported_rows:
            coord_key=row["emp_id"] or f"NAME:{row['full_name']}"
            row_key=(coord_key,row["ministry_id"])
            if row_key in seen_rows:
                duplicates+=1;continue
            seen_rows.add(row_key)
            imported_keys.add(coord_key)
            if coord_key in processed_coords:
                coord_id=processed_coords[coord_key]
            else:
                current=None
                if row["emp_id"]:
                    current=con.execute("SELECT * FROM coordinator_master WHERE emp_id=?",(row["emp_id"],)).fetchone()
                if not current:
                    current=con.execute("SELECT * FROM coordinator_master WHERE full_name=? AND is_active=1",(row["full_name"],)).fetchone()
                is_active=_coordinator_status(row["status"])
                if current:
                    old_name=current["full_name"]
                    changed=any(str(current[k] or "").strip()!=str(row[k] or "").strip()
                      for k in ("full_name","emp_id","mobile","email","city","sector","notes")) or int(current["is_active"] or 0)!=is_active
                    con.execute("""UPDATE coordinator_master SET full_name=?,emp_id=?,mobile=?,email=?,city=?,sector=?,
                      is_active=?,notes=?,imported_at=?,import_batch=?,updated_at=? WHERE id=?""",(
                      row["full_name"],row["emp_id"] or None,row["mobile"],row["email"],row["city"],row["sector"],
                      is_active,row["notes"],now,batch,now,current["id"]))
                    if old_name!=row["full_name"]:
                        con.execute("""UPDATE schools SET coordinator_name=?,coordinator_emp_id=?,coordinator_mobile=?
                          WHERE coordinator_name=?""",(row["full_name"],row["emp_id"],row["mobile"],old_name))
                    updated+=1 if changed else 0;unchanged+=0 if changed else 1
                    coord_id=current["id"]
                else:
                    cur=con.execute("""INSERT INTO coordinator_master(emp_id,full_name,mobile,email,city,sector,
                      is_active,notes,created_by,created_at,imported_at,import_batch) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",(
                      row["emp_id"] or None,row["full_name"],row["mobile"],row["email"],row["city"],row["sector"],
                      is_active,row["notes"],user["id"],now,now,batch))
                    coord_id=cur.lastrowid
                    if not coord_id:
                        coord_id=con.execute("SELECT id FROM coordinator_master WHERE full_name=?",(row["full_name"],)).fetchone()["id"]
                    added+=1
                processed_coords[coord_key]=coord_id

            if update_school_assignments=="yes" and row["ministry_id"]:
                school=con.execute("SELECT * FROM schools WHERE ministry_id=?",(row["ministry_id"],)).fetchone()
                if not school:
                    missing_schools+=1;continue
                if (school["coordinator_name"] or "")!=row["full_name"] or (school["coordinator_emp_id"] or "")!=row["emp_id"]:
                    old=con.execute("SELECT id FROM coordinator_master WHERE full_name=?",(school["coordinator_name"],)).fetchone() if school["coordinator_name"] else None
                    con.execute("""UPDATE schools SET coordinator_name=?,coordinator_emp_id=?,coordinator_mobile=?
                      WHERE ministry_id=?""",(row["full_name"],row["emp_id"],row["mobile"],row["ministry_id"]))
                    con.execute("""INSERT INTO coordinator_assignment_log(ministry_id,school_name,city,
                      old_coordinator_id,old_coordinator_name,new_coordinator_id,new_coordinator_name,
                      action_type,changed_by,changed_at) VALUES(?,?,?,?,?,?,?,?,?,?)""",(
                      row["ministry_id"],school["school_name"],school["city"],old["id"] if old else None,
                      school["coordinator_name"],coord_id,row["full_name"],"استيراد Excel",user["id"],now))
                    assigned+=1

        if deactivate_missing=="yes":
            active_rows=con.execute("SELECT id,emp_id,full_name FROM coordinator_master WHERE is_active=1").fetchall()
            for c in active_rows:
                key=c["emp_id"] or f"NAME:{c['full_name']}"
                if key not in imported_keys:
                    con.execute("UPDATE coordinator_master SET is_active=0,updated_at=? WHERE id=?",(now,c["id"]))
                    deactivated+=1
        con.commit()
    except IntegrityError:
        con.rollback();con.close()
        raise HTTPException(400,detail="يوجد رقم وظيفي مكرر أو مستخدم لمنسق آخر داخل الملف")
    except Exception as exc:
        con.rollback();con.close()
        raise HTTPException(500,detail=f"فشل استيراد المنسقين وتم التراجع عن التغييرات: {exc}")
    con.close()
    return RedirectResponse(
      f"/coordinators?import_added={added}&import_updated={updated}&import_unchanged={unchanged}"
      f"&import_assigned={assigned}&import_deactivated={deactivated}&import_duplicates={duplicates}"
      f"&import_missing_schools={missing_schools}",status_code=303)


@app.post("/coordinators/create")
def coordinator_create(
    request: Request,
    full_name: str=Form(...),
    emp_id: str=Form(""),
    mobile: str=Form(""),
    email: str=Form(""),
    city: str=Form(""),
    sector: str=Form(""),
    notes: str=Form("")
):
    user=require_permission(request,"coordinators")
    con=db()
    try:
        con.execute("""INSERT INTO coordinator_master(
          emp_id,full_name,mobile,email,city,sector,is_active,notes,created_by,created_at
        ) VALUES(?,?,?,?,?,?,1,?,?,?)""",(
          emp_id.strip() or None,full_name.strip(),mobile.strip(),email.strip(),city.strip(),sector.strip(),
          notes.strip(),user["id"],datetime.utcnow().isoformat()
        ))
        con.commit()
    except IntegrityError:
        con.close()
        raise HTTPException(400,detail="الرقم الوظيفي مستخدم لمنسق آخر")
    con.close()
    return RedirectResponse("/coordinators?created=1",status_code=303)

@app.post("/coordinators/{coordinator_id}/update")
def coordinator_update(
    request: Request,
    coordinator_id: int,
    full_name: str=Form(...),
    emp_id: str=Form(""),
    mobile: str=Form(""),
    email: str=Form(""),
    city: str=Form(""),
    sector: str=Form(""),
    notes: str=Form("")
):
    user=require_permission(request,"coordinators")
    con=db()
    old=con.execute("SELECT * FROM coordinator_master WHERE id=?",(coordinator_id,)).fetchone()
    if not old:
        con.close();raise HTTPException(404)
    try:
        con.execute("""UPDATE coordinator_master SET
          full_name=?,emp_id=?,mobile=?,email=?,city=?,sector=?,notes=?,updated_at=?
          WHERE id=?""",(
          full_name.strip(),emp_id.strip() or None,mobile.strip(),email.strip(),city.strip(),sector.strip(),
          notes.strip(),datetime.utcnow().isoformat(),coordinator_id
        ))
        # Keep all assigned schools synchronized with edited coordinator details.
        con.execute("""UPDATE schools SET
          coordinator_name=?,coordinator_emp_id=?,coordinator_mobile=?
          WHERE coordinator_name=?""",(
          full_name.strip(),emp_id.strip(),mobile.strip(),old["full_name"]
        ))
        con.commit()
    except IntegrityError:
        con.rollback();con.close()
        raise HTTPException(400,detail="الرقم الوظيفي مستخدم لمنسق آخر")
    con.close()
    return RedirectResponse("/coordinators?updated=1",status_code=303)

@app.post("/coordinators/{coordinator_id}/toggle")
def coordinator_toggle(request: Request,coordinator_id: int):
    user=require_permission(request,"coordinators")
    con=db()
    row=con.execute("SELECT is_active FROM coordinator_master WHERE id=?",(coordinator_id,)).fetchone()
    if not row:
        con.close();raise HTTPException(404)
    con.execute("UPDATE coordinator_master SET is_active=?,updated_at=? WHERE id=?",
                (0 if row["is_active"] else 1,datetime.utcnow().isoformat(),coordinator_id))
    con.commit();con.close()
    return RedirectResponse("/coordinators",status_code=303)

@app.post("/coordinators/{coordinator_id}/delete")
def coordinator_delete(request: Request,coordinator_id: int):
    user=require_permission(request,"coordinators")
    con=db()
    coord=con.execute("SELECT * FROM coordinator_master WHERE id=?",(coordinator_id,)).fetchone()
    if not coord:
        con.close();raise HTTPException(404)
    assigned=con.execute("SELECT COUNT(*) c FROM schools WHERE coordinator_name=?",
                         (coord["full_name"],)).fetchone()["c"]
    if assigned:
        con.close()
        raise HTTPException(400,detail="انقل مدارس المنسق أولاً قبل حذفه")
    con.execute("DELETE FROM coordinator_master WHERE id=?",(coordinator_id,))
    con.commit();con.close()
    return RedirectResponse("/coordinators?deleted=1",status_code=303)

@app.post("/coordinators/assign")
async def coordinator_assign_schools(request: Request):
    user=require_permission(request,"coordinators")
    form=await request.form()
    coordinator_id=str(form.get("coordinator_id") or "").strip()
    ministry_ids=[str(v).strip() for v in form.getlist("ministry_ids") if str(v).strip()]
    if not coordinator_id or not ministry_ids:
        raise HTTPException(400,detail="اختر المنسق ومدرسة واحدة على الأقل")
    con=db()
    coord=con.execute("SELECT * FROM coordinator_master WHERE id=? AND is_active=1",
                      (coordinator_id,)).fetchone()
    if not coord:
        con.close();raise HTTPException(404,detail="المنسق غير موجود أو غير نشط")
    for mid in ministry_ids:
        school=con.execute("SELECT * FROM schools WHERE ministry_id=?",(mid,)).fetchone()
        if not school:
            continue
        old_id=None
        if school["coordinator_name"]:
            old=con.execute("SELECT id FROM coordinator_master WHERE full_name=?",
                            (school["coordinator_name"],)).fetchone()
            old_id=old["id"] if old else None
        con.execute("""UPDATE schools SET coordinator_name=?,
          coordinator_emp_id=?,coordinator_mobile=? WHERE ministry_id=?""",(
          coord["full_name"],coord["emp_id"] or "",coord["mobile"] or "",mid
        ))
        con.execute("""INSERT INTO coordinator_assignment_log(
          ministry_id,school_name,city,old_coordinator_id,old_coordinator_name,
          new_coordinator_id,new_coordinator_name,action_type,changed_by,changed_at
        ) VALUES(?,?,?,?,?,?,?,?,?,?)""",(
          mid,school["school_name"],school["city"],old_id,
          school["coordinator_name"],coord["id"],coord["full_name"],
          "نقل/إسناد مدارس",user["id"],datetime.utcnow().isoformat()
        ))
    con.commit();con.close()
    return RedirectResponse(f"/coordinators?assigned={len(ministry_ids)}",status_code=303)

@app.post("/coordinators/unassign")
async def coordinator_unassign_schools(request: Request):
    user=require_permission(request,"coordinators")
    form=await request.form()
    ministry_ids=[str(v).strip() for v in form.getlist("ministry_ids") if str(v).strip()]
    con=db()
    for mid in ministry_ids:
        school=con.execute("SELECT * FROM schools WHERE ministry_id=?",(mid,)).fetchone()
        if not school: continue
        old=con.execute("SELECT id FROM coordinator_master WHERE full_name=?",
                        (school["coordinator_name"],)).fetchone() if school["coordinator_name"] else None
        con.execute("""UPDATE schools SET coordinator_name='',coordinator_emp_id='',
          coordinator_mobile='' WHERE ministry_id=?""",(mid,))
        con.execute("""INSERT INTO coordinator_assignment_log(
          ministry_id,school_name,city,old_coordinator_id,old_coordinator_name,
          action_type,changed_by,changed_at
        ) VALUES(?,?,?,?,?,?,?,?)""",(
          mid,school["school_name"],school["city"],old["id"] if old else None,
          school["coordinator_name"],"فك إسناد",user["id"],datetime.utcnow().isoformat()
        ))
    con.commit();con.close()
    return RedirectResponse(f"/coordinators?unassigned={len(ministry_ids)}",status_code=303)

@app.post("/coordinators/redistribute")
async def coordinator_redistribute(request: Request):
    user=require_permission(request,"coordinators")
    form=await request.form()
    city=str(form.get("redistribute_city") or "").strip()
    school_scope=str(form.get("redistribute_scope") or "all").strip()
    selected_ids=[str(v).strip() for v in form.getlist("redistribute_coordinators") if str(v).strip()]
    con=db()
    if selected_ids:
        placeholders=",".join("?" for _ in selected_ids)
        coords=con.execute(f"""SELECT * FROM coordinator_master
          WHERE is_active=1 AND id IN ({placeholders}) ORDER BY full_name""",selected_ids).fetchall()
    else:
        coords=con.execute("SELECT * FROM coordinator_master WHERE is_active=1 ORDER BY full_name").fetchall()
    if not coords:
        con.close();raise HTTPException(400,detail="لا يوجد منسقون نشطون للتوزيع")
    where=["1=1"];args=[]
    if city:
        where.append("city=?");args.append(city)
    if school_scope=="main":
        where.append("school_type LIKE '%رئيس%'")
    elif school_scope=="branch":
        where.append("school_type LIKE '%فرع%'")
    schools=con.execute("SELECT * FROM schools WHERE "+" AND ".join(where)+
                        " ORDER BY city,school_type,school_name",args).fetchall()
    loads={c["id"]:0 for c in coords}
    changed=0
    for school in schools:
        coord=min(coords,key=lambda c:(loads[c["id"]],c["full_name"]))
        old=con.execute("SELECT id FROM coordinator_master WHERE full_name=?",
                        (school["coordinator_name"],)).fetchone() if school["coordinator_name"] else None
        con.execute("""UPDATE schools SET coordinator_name=?,coordinator_emp_id=?,
          coordinator_mobile=? WHERE ministry_id=?""",(
          coord["full_name"],coord["emp_id"] or "",coord["mobile"] or "",school["ministry_id"]
        ))
        con.execute("""INSERT INTO coordinator_assignment_log(
          ministry_id,school_name,city,old_coordinator_id,old_coordinator_name,
          new_coordinator_id,new_coordinator_name,action_type,changed_by,changed_at
        ) VALUES(?,?,?,?,?,?,?,?,?,?)""",(
          school["ministry_id"],school["school_name"],school["city"],
          old["id"] if old else None,school["coordinator_name"],coord["id"],
          coord["full_name"],"إعادة توزيع تلقائي",user["id"],datetime.utcnow().isoformat()
        ))
        loads[coord["id"]]+=1;changed+=1
    con.commit();con.close()
    return RedirectResponse(f"/coordinators?redistributed={changed}",status_code=303)

@app.get("/coordinators/distribution.xlsx")
def coordinators_distribution_export(request: Request):
    user=require_permission(request,"coordinators")
    con=db()
    rows=con.execute("""SELECT ministry_id,school_name,city,school_type,
      coordinator_name,coordinator_emp_id,coordinator_mobile,
      guard_name,guard_emp_id FROM schools ORDER BY city,coordinator_name,school_name""").fetchall()
    con.close()
    wb=Workbook();ws=wb.active;ws.title="توزيع المدارس"
    ws.append(["الرقم الوزاري","اسم المدرسة","المدينة","نوع المدرسة",
               "المنسق / المشرف","رقمه الوظيفي","جواله","البواب","رقم البواب"])
    for r in rows:
        ws.append([r["ministry_id"],r["school_name"],r["city"],r["school_type"],
                   r["coordinator_name"],r["coordinator_emp_id"],r["coordinator_mobile"],
                   r["guard_name"],r["guard_emp_id"]])
    for cell in ws[1]:
        cell.font=cell.font.copy(bold=True,color="FFFFFF")
        cell.fill=cell.fill.copy(fill_type="solid",fgColor="082B60")
    ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
    bio=BytesIO();wb.save(bio);bio.seek(0)
    return StreamingResponse(
      bio,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      headers={"Content-Disposition":'attachment; filename="shalfa_ops_coordinator_distribution.xlsx"'}
    )


# V18.7 Advanced school reports
SCHOOL_REPORT_LABELS={
  "ministry_id":"الرقم الوزاري","school_name":"اسم المدرسة","city":"المدينة",
  "district":"الحي","stage":"المرحلة","gender":"النوع","school_type":"نوع المدرسة",
  "operational_type":"التصنيف التشغيلي","size":"الحجم","status":"حالة المدرسة",
  "latitude":"Latitude","longitude":"Longitude",
  "guard_emp_id":"الرقم الوظيفي للبواب","guard_name":"اسم البواب",
  "guard_id":"هوية البواب","guard_mobile":"جوال البواب",
  "coordinator_emp_id":"الرقم الوظيفي للمنسق","coordinator_name":"المنسق / المشرف",
  "coordinator_mobile":"جوال المنسق"
}

def _school_report_filters(user,report_type,city,coordinator,school_status,operational_type,gender,q):
    # Preserve user data scope without forcing active schools, so closed and merged reports work.
    clauses=[];args=[]
    if user["role"]=="coordinator":
        clauses.append("s.coordinator_name=?");args.append(user["coordinator_name"])
    elif user["role"]=="client":
        if "client_city" in user.keys() and user["client_city"]:
            clauses.append("s.city=?");args.append(user["client_city"])
        if "client_coordinator" in user.keys() and user["client_coordinator"]:
            clauses.append("s.coordinator_name=?");args.append(user["client_coordinator"])
    clause=(" AND "+" AND ".join(clauses)) if clauses else ""
    filters=[]
    if city:
        filters.append("s.city=?");args.append(city)
    if coordinator and user["role"]!="coordinator":
        filters.append("s.coordinator_name=?");args.append(coordinator)
    if school_status:
        filters.append("COALESCE(s.status,'تعمل')=?");args.append(school_status)
    if operational_type:
        filters.append("COALESCE(s.operational_type,'غير محدد')=?");args.append(operational_type)
    if gender:
        filters.append("s.gender=?");args.append(gender)
    if q:
        filters.append("""(
          s.ministry_id LIKE ? OR s.school_name LIKE ? OR
          COALESCE(s.guard_name,'') LIKE ? OR COALESCE(s.guard_emp_id,'') LIKE ? OR
          COALESCE(s.coordinator_name,'') LIKE ? OR COALESCE(s.coordinator_emp_id,'') LIKE ?
        )""")
        term=f"%{q}%";args.extend([term]*6)
    if report_type=="without_guard":
        filters.append("COALESCE(TRIM(s.guard_emp_id),'')=''")
    elif report_type=="without_coordinator":
        filters.append("COALESCE(TRIM(s.coordinator_name),'')=''")
    elif report_type=="missing_coordinates":
        filters.append("(s.latitude IS NULL OR s.longitude IS NULL)")
    elif report_type=="main_schools":
        filters.append("TRIM(COALESCE(s.operational_type,''))='رئيسية'")
    elif report_type=="branch_schools":
        filters.append("TRIM(COALESCE(s.operational_type,''))='فرعية'")
    elif report_type=="annex_schools":
        filters.append("TRIM(COALESCE(s.operational_type,''))='ملحقة'")
    elif report_type=="closed_schools":
        filters.append("COALESCE(s.status,'تعمل')='مغلقة'")
    elif report_type=="merged_schools":
        filters.append("COALESCE(s.status,'تعمل')='مدمجة'")
    where=" ".join([clause]+[f" AND {x}" for x in filters])
    return where,args

def _school_report_rows(con,user,report_type="all",city="",coordinator="",school_status="",
                        operational_type="",gender="",q=""):
    where,args=_school_report_filters(
      user,report_type,city,coordinator,school_status,operational_type,gender,q
    )
    columns=[r["name"] for r in con.execute("PRAGMA table_info(schools)").fetchall()]
    preferred=[
      "ministry_id","school_name","city","district","stage","gender","school_type",
      "operational_type","size","status","latitude","longitude",
      "guard_emp_id","guard_name","guard_id","guard_mobile",
      "coordinator_emp_id","coordinator_name","coordinator_mobile"
    ]
    selected=[c for c in preferred if c in columns]
    selected.extend([c for c in columns if c not in selected])
    sql="SELECT "+",".join(f"s.{c}" for c in selected)+f" FROM schools s WHERE 1=1 {where} ORDER BY s.city,s.coordinator_name,s.school_name"
    rows=con.execute(sql,args).fetchall()
    return selected,rows



@app.get("/school-reports",response_class=HTMLResponse)
def school_reports_page(
    request: Request,report_type: str="all",city: str="",coordinator: str="",
    school_status: str="",operational_type: str="",gender: str="",q: str=""
):
    user=require_permission(request,"school_reports")
    con=db()
    columns,rows=_school_report_rows(
      con,user,report_type,city,coordinator,school_status,operational_type,gender,q
    )
    cities=[r["city"] for r in con.execute("SELECT DISTINCT city FROM schools WHERE COALESCE(city,'')<>'' ORDER BY city")]
    coords=[r["coordinator_name"] for r in con.execute("SELECT DISTINCT coordinator_name FROM schools WHERE COALESCE(coordinator_name,'')<>'' ORDER BY coordinator_name")]
    statuses=[r["status"] for r in con.execute("SELECT DISTINCT COALESCE(status,'تعمل') status FROM schools ORDER BY status")]
    operational_types=[r["operational_type"] for r in con.execute("SELECT DISTINCT COALESCE(operational_type,'غير محدد') operational_type FROM schools ORDER BY operational_type")]
    genders=[r["gender"] for r in con.execute("SELECT DISTINCT gender FROM schools WHERE COALESCE(gender,'')<>'' ORDER BY gender")]
    con.close()
    return templates.TemplateResponse("school_reports.html",{
      "request":request,"user":user,"columns":columns,"labels":SCHOOL_REPORT_LABELS,
      "rows":rows[:500],"total_rows":len(rows),"report_type":report_type,"city":city,
      "coordinator":coordinator,"school_status":school_status,
      "operational_type":operational_type,"gender":gender,"q":q,
      "cities":cities,"coords":coords,"statuses":statuses,
      "operational_types":operational_types,"genders":genders
    })

@app.get("/school-reports.xlsx")
def school_reports_excel(
    request: Request,report_type: str="all",city: str="",coordinator: str="",
    school_status: str="",operational_type: str="",gender: str="",q: str=""
):
    user=require_permission(request,"school_reports")
    con=db()
    columns,rows=_school_report_rows(
      con,user,report_type,city,coordinator,school_status,operational_type,gender,q
    )
    con.close()
    wb=Workbook();ws=wb.active;ws.title="المدارس والبوابين والمنسقين"
    headers=[SCHOOL_REPORT_LABELS.get(c,c) for c in columns]
    ws.append(headers)
    for r in rows:
        ws.append([r[c] for c in columns])
    for cell in ws[1]:
        cell.font=cell.font.copy(bold=True,color="FFFFFF")
        cell.fill=cell.fill.copy(fill_type="solid",fgColor="082B60")
    ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
    widths={}
    for idx,col in enumerate(columns,1):
        label=SCHOOL_REPORT_LABELS.get(col,col)
        width=18
        if col in ("school_name","guard_name","coordinator_name"):width=35
        elif col in ("ministry_id","guard_emp_id","guard_id","guard_mobile","coordinator_emp_id","coordinator_mobile"):width=20
        widths[idx]=width
    from openpyxl.utils import get_column_letter
    for idx,width in widths.items():ws.column_dimensions[get_column_letter(idx)].width=width
    for idx,col in enumerate(columns,1):
        if col in ("ministry_id","guard_emp_id","guard_id","guard_mobile","coordinator_emp_id","coordinator_mobile"):
            for cell in ws.iter_cols(min_col=idx,max_col=idx,min_row=2):
                for x in cell:x.number_format="@"

    # Add summary sheets for multiple management reports.
    summary=wb.create_sheet("ملخص المدن")
    summary.append(["المدينة","إجمالي المدارس","مرتبطة ببواب","بدون بواب","مرتبطة بمنسق","بدون منسق"])
    groups={}
    for r in rows:
        city_name=r["city"] if "city" in columns else ""
        g=groups.setdefault(city_name,{"total":0,"guard":0,"coord":0})
        g["total"]+=1
        if "guard_emp_id" in columns and str(r["guard_emp_id"] or "").strip():g["guard"]+=1
        if "coordinator_name" in columns and str(r["coordinator_name"] or "").strip():g["coord"]+=1
    for city_name,g in sorted(groups.items()):
        summary.append([city_name,g["total"],g["guard"],g["total"]-g["guard"],g["coord"],g["total"]-g["coord"]])

    dist=wb.create_sheet("توزيع المنسقين")
    dist.append(["المنسق / المشرف","الرقم الوظيفي","الجوال","عدد المدارس","المدن"])
    coord_groups={}
    for r in rows:
        name=str(r["coordinator_name"] or "").strip() if "coordinator_name" in columns else ""
        if not name:continue
        g=coord_groups.setdefault(name,{"emp":"","mobile":"","schools":0,"cities":set()})
        g["schools"]+=1
        if "coordinator_emp_id" in columns:g["emp"]=r["coordinator_emp_id"] or g["emp"]
        if "coordinator_mobile" in columns:g["mobile"]=r["coordinator_mobile"] or g["mobile"]
        if "city" in columns and r["city"]:g["cities"].add(str(r["city"]))
    for name,g in sorted(coord_groups.items()):
        dist.append([name,g["emp"],g["mobile"],g["schools"],"، ".join(sorted(g["cities"]))])

    for sheet in (summary,dist):
        for cell in sheet[1]:
            cell.font=cell.font.copy(bold=True,color="FFFFFF")
            cell.fill=cell.fill.copy(fill_type="solid",fgColor="082B60")
        sheet.freeze_panes="A2";sheet.auto_filter.ref=sheet.dimensions
        for col in sheet.columns:
            sheet.column_dimensions[get_column_letter(col[0].column)].width=24

    bio=BytesIO();wb.save(bio);bio.seek(0)
    filename=f"school_report_{report_type}.xlsx"
    return StreamingResponse(
      bio,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      headers={"Content-Disposition":f'attachment; filename="{filename}"'}
    )

@app.get("/school-reports/print",response_class=HTMLResponse)
def school_reports_print(
    request: Request,report_type: str="all",city: str="",coordinator: str="",
    school_status: str="",operational_type: str="",gender: str="",q: str=""
):
    user=require_permission(request,"school_reports")
    con=db()
    columns,rows=_school_report_rows(
      con,user,report_type,city,coordinator,school_status,operational_type,gender,q
    )
    con.close()
    return templates.TemplateResponse("school_reports_print.html",{
      "request":request,"user":user,"columns":columns,"labels":SCHOOL_REPORT_LABELS,
      "rows":rows,"report_type":report_type
    })



# V19.1 unified service endpoints.
@app.post("/api/unified/guard/assign")
def unified_guard_assign(request: Request,ministry_id: str=Form(...),emp_id: str=Form(...),reason: str=Form("")):
    user=require_permission(request,"schools")
    try:
        MasterDataService.assign_guard(ministry_id,emp_id,user["id"],reason)
    except ValueError as exc:
        raise HTTPException(400,detail=str(exc))
    return JSONResponse({"ok":True,"message":"تم ربط البواب وتحديث جميع الشاشات"})

@app.post("/api/unified/guard/unassign")
def unified_guard_unassign(request: Request,ministry_id: str=Form(...),reason: str=Form("")):
    user=require_permission(request,"schools")
    try:
        MasterDataService.unassign_guard(ministry_id,user["id"],reason)
    except ValueError as exc:
        raise HTTPException(400,detail=str(exc))
    return JSONResponse({"ok":True,"message":"تم فك الربط وتحديث جميع الشاشات"})

@app.post("/api/unified/coordinator/assign")
def unified_coordinator_assign(request: Request,ministry_id: str=Form(...),coordinator_id: int=Form(...),reason: str=Form("")):
    user=require_permission(request,"coordinators")
    try:
        MasterDataService.assign_coordinator(ministry_id,coordinator_id,user["id"],reason)
    except ValueError as exc:
        raise HTTPException(400,detail=str(exc))
    return JSONResponse({"ok":True,"message":"تم ربط المنسق / المشرف وتحديث جميع الشاشات"})

@app.post("/api/unified/school/status")
def unified_school_status(request: Request,ministry_id: str=Form(...),status: str=Form(...),reason: str=Form("")):
    user=require_permission(request,"schools")
    try:
        MasterDataService.set_school_status(ministry_id,status,user["id"],reason)
    except ValueError as exc:
        raise HTTPException(400,detail=str(exc))
    return JSONResponse({"ok":True,"message":"تم تحديث حالة المدرسة في جميع الشاشات"})

@app.post("/api/unified/reconcile")
def unified_reconcile(request: Request):
    user=require_permission(request,"schools")
    return JSONResponse({"ok":True,"stats":reconcile_unified_data("manual",emit_event=True)})



# V19.2 relationship center and near-real-time events.
@app.get("/relationship-center",response_class=HTMLResponse)
def relationship_center_page(request: Request):
    user=require_permission(request,"relationship_center")
    reconcile_unified_data("relationship_center_open",emit_event=False)
    summary=relationship_summary()
    con=db()
    no_guard=con.execute("""SELECT ministry_id,school_name,city,coordinator_name
      FROM schools WHERE COALESCE(status,'تعمل')='تعمل'
      AND COALESCE(TRIM(guard_emp_id),'')='' ORDER BY city,school_name LIMIT 200""").fetchall()
    no_coord=con.execute("""SELECT ministry_id,school_name,city,guard_name
      FROM schools WHERE COALESCE(status,'تعمل')='تعمل'
      AND COALESCE(TRIM(coordinator_name),'')='' ORDER BY city,school_name LIMIT 200""").fetchall()
    no_coords=con.execute("""SELECT ministry_id,school_name,city FROM schools
      WHERE latitude IS NULL OR longitude IS NULL ORDER BY city,school_name LIMIT 200""").fetchall()
    unlinked_guards=con.execute("""SELECT emp_id,emp_name,mobile,service_status
      FROM guard_master WHERE COALESCE(TRIM(linked_ministry_id),'')=''
      AND COALESCE(service_status,'على رأس العمل')<>'منتهي الخدمة'
      ORDER BY emp_name LIMIT 200""").fetchall()
    duplicate_guards=con.execute("""SELECT guard_emp_id,MAX(guard_name) guard_name,COUNT(*) school_count
      FROM schools WHERE COALESCE(TRIM(guard_emp_id),'')<>''
      GROUP BY guard_emp_id HAVING COUNT(*)>1 ORDER BY school_count DESC""").fetchall()
    events=con.execute("""SELECT * FROM relationship_events ORDER BY id DESC LIMIT 50""").fetchall()
    con.close()
    return templates.TemplateResponse("relationship_center.html",{
      "request":request,"user":user,"summary":summary,"no_guard":no_guard,
      "no_coord":no_coord,"no_coords":no_coords,"unlinked_guards":unlinked_guards,
      "duplicate_guards":duplicate_guards,"events":events
    })

@app.get("/api/relationship-summary")
def relationship_summary_api(request: Request):
    require_permission(request,"relationship_center")
    return JSONResponse(relationship_summary())

@app.get("/api/relationship-events")
def relationship_events_api(request: Request,since: int=0):
    require_permission(request,"relationship_center")
    con=db()
    rows=con.execute("""SELECT * FROM relationship_events WHERE id>? ORDER BY id LIMIT 100""",(since,)).fetchall()
    con.close()
    return JSONResponse({"events":[dict(r) for r in rows],"summary":relationship_summary()})

@app.get("/relationship-events/stream")
async def relationship_events_stream(request: Request,since: int=0):
    require_permission(request,"relationship_center")
    async def generate():
        last_id=since
        while True:
            if await request.is_disconnected():
                break
            con=db()
            rows=con.execute("""SELECT * FROM relationship_events WHERE id>? ORDER BY id LIMIT 100""",(last_id,)).fetchall()
            con.close()
            if rows:
                last_id=rows[-1]["id"]
                payload={"events":[dict(r) for r in rows],"summary":relationship_summary(),"last_id":last_id}
                yield f"data: {json.dumps(payload,ensure_ascii=False)}\n\n"
            else:
                yield f"event: ping\ndata: {json.dumps({'last_id':last_id})}\n\n"
            await asyncio.sleep(5)
    return StreamingResponse(generate(),media_type="text/event-stream",
      headers={"Cache-Control":"no-cache","Connection":"keep-alive"})

@app.post("/relationship-center/repair")
def relationship_center_repair(request: Request,action: str=Form("reconcile")):
    user=require_permission(request,"relationship_center")
    if action=="duplicate_guards":
        con=db()
        groups=con.execute("""SELECT guard_emp_id FROM schools
          WHERE COALESCE(TRIM(guard_emp_id),'')<>'' GROUP BY guard_emp_id HAVING COUNT(*)>1""").fetchall()
        fixed=0
        for group in groups:
            emp_id=group["guard_emp_id"]
            schools=con.execute("""SELECT ministry_id,school_name FROM schools
              WHERE guard_emp_id=? ORDER BY ministry_id""",(emp_id,)).fetchall()
            keep=schools[0] if schools else None
            for school in schools[1:]:
                con.execute("""UPDATE schools SET guard_name='',guard_emp_id='',guard_id='',guard_mobile=''
                  WHERE ministry_id=?""",(school["ministry_id"],));fixed+=1
            if keep:
                try:con.execute("UPDATE guard_master SET linked_ministry_id=? WHERE emp_id=?",(keep["ministry_id"],emp_id))
                except Exception:pass
        con.commit();con.close()
        stats=reconcile_unified_data("repair_duplicate_guards",emit_event=True)
        record_relationship_event("system","guards","repair_duplicates",{"fixed":fixed,"user_id":user["id"]})
        return RedirectResponse(f"/relationship-center?fixed={fixed}",status_code=303)
    stats=reconcile_unified_data("relationship_repair",emit_event=True)
    return RedirectResponse("/relationship-center?reconciled=1",status_code=303)


# V20.0_AUDIT_ROUTES
@app.get('/audit-trail',response_class=HTMLResponse)
def audit_trail_page(request: Request,username: str='',module: str='',date_from: str='',date_to: str='',success: str=''):
    user=require_permission(request,'audit_trail');con=db();where=['1=1'];args=[]
    if username:where.append('username LIKE ?');args.append(f'%{username}%')
    if module:where.append('module=?');args.append(module)
    if date_from:where.append('date(created_at)>=date(?)');args.append(date_from)
    if date_to:where.append('date(created_at)<=date(?)');args.append(date_to)
    if success in ('0','1'):where.append('success=?');args.append(int(success))
    rows=con.execute('SELECT * FROM audit_trail WHERE '+' AND '.join(where)+' ORDER BY id DESC LIMIT 3000',args).fetchall()
    modules=[r['module'] for r in con.execute("SELECT DISTINCT module FROM audit_trail WHERE COALESCE(module,'')<>'' ORDER BY module")]
    con.close();return templates.TemplateResponse('audit_trail.html',{'request':request,'user':user,'rows':rows,'modules':modules,'username':username,'module':module,'date_from':date_from,'date_to':date_to,'success':success})

@app.get('/audit-trail.xlsx')
def audit_trail_excel(request: Request):
    require_permission(request,'audit_trail');con=db();rows=con.execute('SELECT * FROM audit_trail ORDER BY id DESC').fetchall();con.close()
    wb=Workbook();ws=wb.active;ws.title='سجل التدقيق';ws.append(['الوقت','المستخدم','الدور','العملية','الوحدة','المسار','الحالة','IP','الجهاز','نجاح'])
    for r in rows:ws.append([r['created_at'],r['username'],r['user_role'],r['action'],r['module'],r['request_path'],r['status_code'],r['ip_address'],r['user_agent'],r['success']])
    bio=BytesIO();wb.save(bio);bio.seek(0);return StreamingResponse(bio,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={'Content-Disposition':'attachment; filename="audit_trail.xlsx"'})

# V20.1_TASK_ROUTES
def _task_scope(user):return ('assigned_to_user_id=?',[user['id']]) if user['role']=='coordinator' else ('1=1',[])
@app.get('/tasks',response_class=HTMLResponse)
def tasks_page(request: Request,status: str='',priority: str='',q: str=''):
    user=require_permission(request,'tasks');con=db();scope,args=_task_scope(user);where=[scope]
    if status:where.append('status=?');args.append(status)
    if priority:where.append('priority=?');args.append(priority)
    if q:where.append('(task_no LIKE ? OR title LIKE ? OR school_name LIKE ?)');args.extend([f'%{q}%']*3)
    rows=con.execute('SELECT * FROM tasks WHERE '+' AND '.join(where)+' ORDER BY id DESC',args).fetchall()
    users=con.execute("SELECT id,username,full_name,role FROM users WHERE is_active=1 ORDER BY full_name,username").fetchall();schools=con.execute("SELECT ministry_id,school_name,city FROM schools WHERE COALESCE(status,'تعمل')='تعمل' ORDER BY city,school_name").fetchall();con.close()
    return templates.TemplateResponse('tasks.html',{'request':request,'user':user,'rows':rows,'users':users,'schools':schools,'status':status,'priority':priority,'q':q})
@app.post('/tasks/create')
def task_create(request: Request,title: str=Form(...),description: str=Form(''),task_type: str=Form('عام'),priority: str=Form('متوسط'),ministry_id: str=Form(''),assigned_to_user_id: int=Form(...),due_at: str=Form(''),sla_hours: int=Form(24)):
    user=require_permission(request,'tasks');con=db();assignee=con.execute('SELECT * FROM users WHERE id=?',(assigned_to_user_id,)).fetchone();school=con.execute('SELECT * FROM schools WHERE ministry_id=?',(ministry_id,)).fetchone() if ministry_id else None
    if not assignee:con.close();raise HTTPException(400,detail='المستخدم المكلف غير موجود')
    no='TSK-'+datetime.utcnow().strftime('%Y%m%d%H%M%S%f');now=datetime.utcnow().isoformat()
    con.execute("""INSERT INTO tasks(task_no,title,description,task_type,priority,ministry_id,school_name,city,assigned_to_user_id,assigned_to_name,assigned_to_role,status,due_at,sla_hours,created_by,created_by_name,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(no,title,description,task_type,priority,ministry_id,school['school_name'] if school else '',school['city'] if school else '',assigned_to_user_id,assignee['full_name'] or assignee['username'],assignee['role'],'جديدة',due_at,sla_hours,user['id'],user['username'],now));tid=con.execute('SELECT id FROM tasks WHERE task_no=?',(no,)).fetchone()['id'];con.execute("INSERT INTO task_events(task_id,event_type,status_to,notes,user_id,username,created_at) VALUES(?,?,?,?,?,?,?)",(tid,'إنشاء','جديدة','',user['id'],user['username'],now));con.commit();con.close();return RedirectResponse('/tasks',status_code=303)
@app.post('/tasks/{task_id}/update')
def task_update(request: Request,task_id: int,status: str=Form(...),progress: int=Form(0),notes: str=Form(''),latitude: str=Form(''),longitude: str=Form('')):
    user=require_permission(request,'tasks');con=db();task=con.execute('SELECT * FROM tasks WHERE id=?',(task_id,)).fetchone()
    if not task:con.close();raise HTTPException(404)
    if user['role']=='coordinator' and task['assigned_to_user_id']!=user['id']:con.close();raise HTTPException(403)
    now=datetime.utcnow().isoformat();timestamps={'مستلمة':'accepted_at','قيد التنفيذ':'started_at','بانتظار اعتماد':'submitted_at','مكتملة':'completed_at'};field=timestamps.get(status);extra=f',{field}=?' if field else '';params=[status,max(0,min(100,progress)),notes,float(latitude) if latitude else None,float(longitude) if longitude else None]+([now] if field else [])+[task_id]
    con.execute(f'UPDATE tasks SET status=?,progress=?,execution_notes=?,latitude=?,longitude=? {extra} WHERE id=?',params);con.execute("INSERT INTO task_events(task_id,event_type,status_from,status_to,notes,user_id,username,created_at) VALUES(?,?,?,?,?,?,?,?)",(task_id,'تحديث',task['status'],status,notes,user['id'],user['username'],now));con.commit();con.close();return RedirectResponse('/tasks',status_code=303)

# V20.2_MOBILE_ROUTES
@app.get('/coordinator-mobile',response_class=HTMLResponse)
def coordinator_mobile_page(request: Request):
    user=require_permission(request,'coordinator_mobile');con=db()
    if user['role']=='coordinator':schools=con.execute("SELECT * FROM schools WHERE coordinator_name=? AND COALESCE(status,'تعمل')='تعمل' ORDER BY city,school_name",(user['coordinator_name'],)).fetchall();tasks=con.execute("SELECT * FROM tasks WHERE assigned_to_user_id=? AND status NOT IN ('مكتملة','مرفوضة') ORDER BY id DESC",(user['id'],)).fetchall()
    else:schools=con.execute("SELECT * FROM schools WHERE COALESCE(status,'تعمل')='تعمل' ORDER BY city,school_name LIMIT 500").fetchall();tasks=con.execute("SELECT * FROM tasks WHERE status NOT IN ('مكتملة','مرفوضة') ORDER BY id DESC LIMIT 100").fetchall()
    con.close();return templates.TemplateResponse('coordinator_mobile.html',{'request':request,'user':user,'schools':schools,'tasks':tasks})
@app.get('/coordinator-mobile/school/{ministry_id}',response_class=HTMLResponse)
def coordinator_mobile_school(request: Request,ministry_id: str):
    user=require_permission(request,'coordinator_mobile');con=db();school=con.execute('SELECT * FROM schools WHERE ministry_id=?',(ministry_id,)).fetchone();con.close()
    if not school:raise HTTPException(404)
    if user['role']=='coordinator' and school['coordinator_name']!=user['coordinator_name']:raise HTTPException(403)
    return templates.TemplateResponse('coordinator_mobile_school.html',{'request':request,'user':user,'school':school})

# V20.3_ASK_SHALFA_ROUTES
def _ask_shalfa(question,user):
    q=(question or '').strip();con=db();scope,args=scope_clause(user)
    if 'بدون بواب' in q:
        rows=con.execute(f"SELECT school_name,city FROM schools s WHERE 1=1 {scope} AND COALESCE(TRIM(guard_emp_id),'')='' LIMIT 50",args).fetchall();ans=f"عدد المدارس بدون بواب: {len(rows)}\n"+'\n'.join(f"- {r['school_name']} ({r['city']})" for r in rows)
    elif 'بدون منسق' in q or 'بدون مشرف' in q:
        rows=con.execute(f"SELECT school_name,city FROM schools s WHERE 1=1 {scope} AND COALESCE(TRIM(coordinator_name),'')='' LIMIT 50",args).fetchall();ans=f"عدد المدارس بدون منسق / مشرف: {len(rows)}\n"+'\n'.join(f"- {r['school_name']} ({r['city']})" for r in rows)
    elif 'مهام' in q:
        rows=con.execute('SELECT status,COUNT(*) c FROM tasks WHERE assigned_to_user_id=? GROUP BY status',(user['id'],)).fetchall() if user['role']=='coordinator' else con.execute('SELECT status,COUNT(*) c FROM tasks GROUP BY status').fetchall();ans='ملخص المهام:\n'+'\n'.join(f"- {r['status']}: {r['c']}" for r in rows)
    elif 'حضور' in q:
        d=date.today().isoformat();rows=con.execute('SELECT status,COUNT(*) c FROM attendance WHERE work_date=? GROUP BY status',(d,)).fetchall();ans=f'حضور اليوم {d}:\n'+'\n'.join(f"- {r['status']}: {r['c']}" for r in rows)
    elif 'مخالفة' in q:
        rows=con.execute('SELECT violation_type,COUNT(*) c FROM violations GROUP BY violation_type ORDER BY c DESC LIMIT 10').fetchall();ans='أكثر أنواع المخالفات:\n'+'\n'.join(f"- {r['violation_type']}: {r['c']}" for r in rows)
    else:
        x=relationship_summary();ans=f"يمكنك السؤال عن الحضور والمدارس بدون بواب أو منسق والمهام والمخالفات.\n- المدارس: {x['schools']}\n- البوابون: {x['guards']}\n- المنسقون: {x['coordinators']}\n- بدون بواب: {x['schools_without_guard']}\n- بدون منسق: {x['schools_without_coordinator']}"
    con.close();return ans
@app.get('/ask-shalfa',response_class=HTMLResponse)
def ask_shalfa_page(request: Request,q: str=''):
    user=require_permission(request,'ask_shalfa_ai');return templates.TemplateResponse('ask_shalfa.html',{'request':request,'user':user,'q':q,'answer':_ask_shalfa(q,user) if q else ''})
@app.post('/api/ask-shalfa')
async def ask_shalfa_api(request: Request):
    user=require_permission(request,'ask_shalfa_ai');data=await request.json();return JSONResponse({'answer':_ask_shalfa(str(data.get('question','')),user)})



# V22 TBC FM — personnel import, multi-file inventory and automatic pricing.
FM_ASSET_UPLOAD_DIR=STATIC_DIR / "uploads" / "fm_assets"
FM_ASSET_UPLOAD_DIR.mkdir(parents=True,exist_ok=True)

def _fm_header(value):
    value=str(value or "").replace("\xa0"," ").strip()
    return re.sub(r"\s+"," ",value).replace("／","/").replace("\\","/")

def _fm_cell(value):
    if value is None:return ""
    value=str(value).strip()
    if value.endswith(".0") and value[:-2].isdigit():return value[:-2]
    return value

def _fm_find_col(headers,*aliases):
    exact={_fm_header(x):x for x in headers if _fm_header(x)}
    compact={re.sub(r"[\s/_-]+","",_fm_header(x)):x for x in headers if _fm_header(x)}
    for alias in aliases:
        a=_fm_header(alias)
        if a in exact:return exact[a]
        c=re.sub(r"[\s/_-]+","",a)
        if c in compact:return compact[c]
    return None

def _fm_float(value,default=0.0):
    try:
        if value in (None,""):return default
        return float(str(value).replace(",","").strip())
    except Exception:return default

def _fm_safe_filename(filename):
    base=Path(filename or "upload.bin").name
    safe=re.sub(r"[^A-Za-z0-9._-]+","_",base)
    return f"{datetime.utcnow().strftime('%Y%m%d%H%M%S%f')}_{safe}"

def _fm_save_photo(upload):
    if not upload or not getattr(upload,"filename",""):return ""
    content=upload.file.read()
    if not content:return ""
    safe=_fm_safe_filename(upload.filename)
    target=FM_ASSET_UPLOAD_DIR / safe
    target.write_bytes(content)
    return f"/static/uploads/fm_assets/{safe}"

def _fm_condition_default_factor(status):
    return {
      "جديدة":1.00,"ممتازة":0.95,"جيدة":0.85,"متوسطة":0.65,
      "ضعيفة":0.40,"تالفة":0.10,"خارج الخدمة":0.05
    }.get(str(status or "").strip(),1.00)

def _fm_price_asset(con,asset,priced_by=None,notes=""):
    asset_type=str(asset["asset_type"] or "").strip()
    asset_name=str(asset["asset_name"] or "").strip()
    manufacturer=str(asset["manufacturer"] or "").strip()
    model=str(asset["model"] or "").strip()
    condition=str(asset["condition_status"] or "جيدة").strip()
    quantity=float(asset["quantity"] or 1) if "quantity" in asset.keys() else 1.0

    # Exact rows score higher. Generic 'الكل' condition is a fallback.
    rows=con.execute("""SELECT * FROM fm_price_catalog WHERE is_active=1
      AND asset_type=? AND condition_status IN (?, 'الكل','ALL','')
      ORDER BY
        CASE WHEN condition_status=? THEN 0 ELSE 1 END,
        CASE WHEN COALESCE(model,'')=? AND ?<>'' THEN 0 ELSE 1 END,
        CASE WHEN COALESCE(manufacturer,'')=? AND ?<>'' THEN 0 ELSE 1 END,
        CASE WHEN COALESCE(asset_name,'')=? AND ?<>'' THEN 0 ELSE 1 END,
        id DESC""",(
          asset_type,condition,condition,model,model,manufacturer,manufacturer,asset_name,asset_name
        )).fetchall()
    catalog=rows[0] if rows else None
    if not catalog:
        old=float(asset["estimated_unit_price"] or 0) if "estimated_unit_price" in asset.keys() else 0
        con.execute("""UPDATE fm_assets SET estimated_unit_price=0,estimated_total_price=0,
          price_catalog_id=NULL,pricing_status='لا يوجد سعر مطابق',pricing_notes=?,updated_at=?
          WHERE id=?""",(notes,datetime.utcnow().isoformat(),asset["id"]))
        return {"matched":False,"unit_price":0,"total_price":0}

    factor=float(catalog["condition_factor"] or 0)
    if factor<=0:factor=_fm_condition_default_factor(condition)
    base=float(catalog["base_price"] or 0)
    calculated=float(catalog["calculated_unit_price"] or 0)
    unit_price=calculated if calculated>0 else base*factor
    total=unit_price*quantity
    old=float(asset["estimated_unit_price"] or 0) if "estimated_unit_price" in asset.keys() else 0
    con.execute("""UPDATE fm_assets SET estimated_unit_price=?,estimated_total_price=?,
      price_catalog_id=?,pricing_status='تم التسعير',pricing_notes=?,updated_at=? WHERE id=?""",
      (unit_price,total,catalog["id"],notes,datetime.utcnow().isoformat(),asset["id"]))
    con.execute("""INSERT INTO fm_asset_price_history(asset_id,price_catalog_id,old_price,new_price,
      condition_status,quantity,total_price,pricing_method,priced_by,priced_at,notes)
      VALUES(?,?,?,?,?,?,?,?,?,?,?)""",(
      asset["id"],catalog["id"],old,unit_price,condition,quantity,total,
      "مطابقة قائمة الأسعار",priced_by,datetime.utcnow().isoformat(),notes
    ))
    return {"matched":True,"unit_price":unit_price,"total_price":total}

def _fm_reprice_all(con,priced_by=None):
    rows=con.execute("SELECT * FROM fm_assets WHERE is_active=1 ORDER BY id").fetchall()
    matched=unmatched=0;total=0
    for asset in rows:
        result=_fm_price_asset(con,asset,priced_by,"إعادة تسعير جماعية")
        if result["matched"]:matched+=1;total+=result["total_price"]
        else:unmatched+=1
    return {"matched":matched,"unmatched":unmatched,"total":total}


# V21 TBC FM routes
def _fm_scope(user,alias='s'):
    conditions=[];args=[]
    if user['role']=='coordinator':
        conditions.append(f"{alias}.coordinator_name=?");args.append(user['coordinator_name'])
    elif user['role']=='client':
        if 'client_city' in user.keys() and user['client_city']:
            conditions.append(f"{alias}.city=?");args.append(user['client_city'])
    return (' AND '+' AND '.join(conditions) if conditions else ''),args


# V22 TBC FM employees, inventory and pricing routes.
@app.get('/fm-employees',response_class=HTMLResponse)
def fm_employees_page(request: Request,department: str='',city: str='',status: str='',q: str=''):
    user=require_permission(request,'fm_employees');con=db();where=['1=1'];args=[]
    if department:where.append('department=?');args.append(department)
    if city:where.append('city=?');args.append(city)
    if status:where.append('employment_status=?');args.append(status)
    if q:where.append('(emp_id LIKE ? OR full_name LIKE ? OR id_no LIKE ? OR mobile LIKE ?)');args.extend([f'%{q}%']*4)
    rows=con.execute("SELECT * FROM fm_employees WHERE "+' AND '.join(where)+" ORDER BY department,city,full_name",args).fetchall()
    departments=[r['department'] for r in con.execute("SELECT DISTINCT department FROM fm_employees WHERE COALESCE(department,'')<>'' ORDER BY department")]
    cities=[r['city'] for r in con.execute("SELECT DISTINCT city FROM fm_employees WHERE COALESCE(city,'')<>'' ORDER BY city")]
    con.close();return templates.TemplateResponse('fm_employees.html',{'request':request,'user':user,'rows':rows,'departments':departments,'cities':cities,'department':department,'city':city,'status':status,'q':q})

@app.post('/fm-employees/create')
def fm_employee_create(request: Request,emp_id: str=Form(...),full_name: str=Form(...),id_no: str=Form(''),
  mobile: str=Form(''),email: str=Form(''),department: str=Form('الصيانة'),job_title: str=Form(''),
  trade: str=Form(''),city: str=Form(''),team_name: str=Form(''),supervisor_name: str=Form(''),
  employment_status: str=Form('نشط'),hire_date: str=Form(''),contract_expiry: str=Form(''),notes: str=Form('')):
    require_permission(request,'fm_employees');con=db();now=datetime.utcnow().isoformat()
    con.execute("""INSERT INTO fm_employees(emp_id,full_name,id_no,mobile,email,department,job_title,trade,
      city,team_name,supervisor_name,employment_status,hire_date,contract_expiry,notes,created_at,updated_at)
      VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
      ON CONFLICT(emp_id) DO UPDATE SET full_name=excluded.full_name,id_no=excluded.id_no,
      mobile=excluded.mobile,email=excluded.email,department=excluded.department,
      job_title=excluded.job_title,trade=excluded.trade,city=excluded.city,
      team_name=excluded.team_name,supervisor_name=excluded.supervisor_name,
      employment_status=excluded.employment_status,hire_date=excluded.hire_date,
      contract_expiry=excluded.contract_expiry,notes=excluded.notes,updated_at=excluded.updated_at""",
      (emp_id.strip(),full_name.strip(),id_no.strip(),mobile.strip(),email.strip(),department.strip(),
       job_title.strip(),trade.strip(),city.strip(),team_name.strip(),supervisor_name.strip(),
       employment_status.strip(),hire_date,contract_expiry,notes.strip(),now,now))
    con.commit();con.close();return RedirectResponse('/fm-employees',status_code=303)

@app.get('/fm-employees/template.xlsx')
def fm_employees_template(request: Request):
    require_permission(request,'fm_employees');wb=Workbook();ws=wb.active;ws.title='موظفو الصيانة والنظافة'
    ws.append(['الرقم الوظيفي','اسم الموظف','رقم الهوية','الجوال','البريد الإلكتروني','القسم',
      'المسمى الوظيفي','التخصص/الحرفة','المدينة','الفريق','المشرف','الحالة','تاريخ التعيين','انتهاء العقد','ملاحظات'])
    ws.append(['FM1001','مثال: أحمد محمد','10xxxxxxxx','05xxxxxxxx','name@example.com','الصيانة',
      'فني صيانة','كهرباء','بريدة','فريق 1','مشرف الصيانة','نشط','','',''])
    for c in ws[1]:c.font=c.font.copy(bold=True,color='FFFFFF');c.fill=c.fill.copy(fill_type='solid',fgColor='082B60')
    ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
    bio=BytesIO();wb.save(bio);bio.seek(0)
    return StreamingResponse(bio,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      headers={'Content-Disposition':'attachment; filename="fm_employees_template.xlsx"'})

@app.post('/fm-employees/import')
async def fm_employees_import(request: Request,file: UploadFile=File(...),replace_old: str=Form('')):
    user=require_permission(request,'fm_employees');content=await file.read()
    try:wb=load_workbook(BytesIO(content),data_only=True,read_only=True)
    except Exception as exc:raise HTTPException(400,detail=f'تعذر قراءة الملف: {exc}')
    con=db();added=updated=errors=0;imported=set();now=datetime.utcnow().isoformat()
    try:
        if replace_old=='yes':con.execute("DELETE FROM fm_employees")
        for ws in wb.worksheets:
            headers=[_fm_header(c.value) for c in ws[1]]
            emp_col=_fm_find_col(headers,'الرقم الوظيفي','رقم الموظف','emp_id')
            name_col=_fm_find_col(headers,'اسم الموظف','الاسم','full_name','name')
            if not emp_col or not name_col:continue
            cols={
              'id_no':_fm_find_col(headers,'رقم الهوية','الهوية','id_no'),
              'mobile':_fm_find_col(headers,'الجوال','رقم الجوال','mobile'),
              'email':_fm_find_col(headers,'البريد الإلكتروني','البريد','email'),
              'department':_fm_find_col(headers,'القسم','department'),
              'job_title':_fm_find_col(headers,'المسمى الوظيفي','الوظيفة','job_title'),
              'trade':_fm_find_col(headers,'التخصص/الحرفة','التخصص','الحرفة','trade'),
              'city':_fm_find_col(headers,'المدينة','city'),
              'team_name':_fm_find_col(headers,'الفريق','اسم الفريق','team'),
              'supervisor_name':_fm_find_col(headers,'المشرف','اسم المشرف','supervisor'),
              'employment_status':_fm_find_col(headers,'الحالة','حالة الموظف','status'),
              'hire_date':_fm_find_col(headers,'تاريخ التعيين','hire_date'),
              'contract_expiry':_fm_find_col(headers,'انتهاء العقد','تاريخ انتهاء العقد','contract_expiry'),
              'notes':_fm_find_col(headers,'ملاحظات','notes')
            }
            for values in ws.iter_rows(min_row=2,values_only=True):
                row=dict(zip(headers,values));emp_id=_fm_cell(row.get(emp_col));name=_fm_cell(row.get(name_col))
                if not emp_id or not name:errors+=1;continue
                exists=con.execute("SELECT id FROM fm_employees WHERE emp_id=?",(emp_id,)).fetchone()
                data={k:_fm_cell(row.get(v)) if v else '' for k,v in cols.items()}
                con.execute("""INSERT INTO fm_employees(emp_id,full_name,id_no,mobile,email,department,job_title,
                  trade,city,team_name,supervisor_name,employment_status,hire_date,contract_expiry,notes,created_at,updated_at)
                  VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                  ON CONFLICT(emp_id) DO UPDATE SET full_name=excluded.full_name,id_no=excluded.id_no,
                  mobile=excluded.mobile,email=excluded.email,department=excluded.department,
                  job_title=excluded.job_title,trade=excluded.trade,city=excluded.city,
                  team_name=excluded.team_name,supervisor_name=excluded.supervisor_name,
                  employment_status=excluded.employment_status,hire_date=excluded.hire_date,
                  contract_expiry=excluded.contract_expiry,notes=excluded.notes,updated_at=excluded.updated_at""",
                  (emp_id,name,data['id_no'],data['mobile'],data['email'],data['department'] or 'الصيانة والنظافة',
                   data['job_title'],data['trade'],data['city'],data['team_name'],data['supervisor_name'],
                   data['employment_status'] or 'نشط',data['hire_date'],data['contract_expiry'],data['notes'],now,now))
                imported.add(emp_id);updated+=1 if exists else 0;added+=0 if exists else 1
        con.commit()
    except Exception as exc:
        con.rollback();con.close();wb.close();raise HTTPException(500,detail=f'فشل الاستيراد: {exc}')
    con.close();wb.close()
    return RedirectResponse(f'/fm-employees?added={added}&updated={updated}&errors={errors}',status_code=303)

@app.get('/fm-inventory',response_class=HTMLResponse)
def fm_inventory_page(request: Request,city: str='',asset_type: str='',condition_status: str='',pricing_status: str='',q: str=''):
    user=require_permission(request,'fm_inventory');con=db();where=['a.is_active=1'];args=[]
    if city:where.append('a.city=?');args.append(city)
    if asset_type:where.append('a.asset_type=?');args.append(asset_type)
    if condition_status:where.append('a.condition_status=?');args.append(condition_status)
    if pricing_status:where.append('a.pricing_status=?');args.append(pricing_status)
    if q:where.append('(a.asset_code LIKE ? OR a.asset_name LIKE ? OR a.school_name LIKE ? OR a.serial_number LIKE ?)');args.extend([f'%{q}%']*4)
    rows=con.execute("SELECT a.* FROM fm_assets a WHERE "+' AND '.join(where)+" ORDER BY a.city,a.school_name,a.asset_type,a.id DESC",args).fetchall()
    schools=con.execute("SELECT ministry_id,school_name,city FROM schools WHERE COALESCE(status,'تعمل')='تعمل' ORDER BY city,school_name").fetchall()
    batches=con.execute("SELECT * FROM fm_inventory_batches ORDER BY id DESC LIMIT 20").fetchall()
    cities=[r['city'] for r in con.execute("SELECT DISTINCT city FROM fm_assets WHERE COALESCE(city,'')<>'' ORDER BY city")]
    types=[r['asset_type'] for r in con.execute("SELECT DISTINCT asset_type FROM fm_assets WHERE COALESCE(asset_type,'')<>'' ORDER BY asset_type")]
    totals=con.execute("""SELECT COUNT(*) records,COALESCE(SUM(quantity),0) quantity,
      COALESCE(SUM(estimated_total_price),0) value,
      SUM(CASE WHEN pricing_status='تم التسعير' THEN 1 ELSE 0 END) priced
      FROM fm_assets WHERE is_active=1""").fetchone()
    con.close();return templates.TemplateResponse('fm_inventory.html',{'request':request,'user':user,'rows':rows,'schools':schools,'batches':batches,'cities':cities,'types':types,'totals':totals,'city':city,'asset_type':asset_type,'condition_status':condition_status,'pricing_status':pricing_status,'q':q})

@app.post('/fm-inventory/create')
async def fm_inventory_create(request: Request,ministry_id: str=Form(...),asset_type: str=Form(...),
  asset_name: str=Form(''),manufacturer: str=Form(''),model: str=Form(''),serial_number: str=Form(''),
  quantity: float=Form(1),unit: str=Form('عدد'),condition_status: str=Form('جيدة'),
  criticality: str=Form('متوسط'),location_description: str=Form(''),inventory_date: str=Form(''),
  inventoried_by: str=Form(''),notes: str=Form(''),photo: UploadFile|None=File(None)):
    user=require_permission(request,'fm_inventory');con=db();school=con.execute("SELECT * FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
    if not school:con.close();raise HTTPException(404,detail='المدرسة غير موجودة')
    photo_path=''
    if photo and photo.filename:
        safe=_fm_safe_filename(photo.filename);target=FM_ASSET_UPLOAD_DIR/safe;target.write_bytes(await photo.read());photo_path=f'/static/uploads/fm_assets/{safe}'
    code='AST-'+datetime.utcnow().strftime('%Y%m%d%H%M%S%f');now=datetime.utcnow().isoformat()
    cur=con.execute("""INSERT INTO fm_assets(asset_code,ministry_id,school_name,city,asset_type,asset_name,
      manufacturer,model,serial_number,condition_status,criticality,location_description,qr_code,
      quantity,unit,inventory_date,inventoried_by,photo_path,notes,created_at,updated_at)
      VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(code,ministry_id,school['school_name'],school['city'],
      asset_type,asset_name,manufacturer,model,serial_number,condition_status,criticality,location_description,
      code,quantity,unit,inventory_date or date.today().isoformat(),inventoried_by or user['username'],photo_path,notes,now,now))
    asset_id=cur.lastrowid
    if not asset_id:asset_id=con.execute("SELECT id FROM fm_assets WHERE asset_code=?",(code,)).fetchone()['id']
    asset=con.execute("SELECT * FROM fm_assets WHERE id=?",(asset_id,)).fetchone();_fm_price_asset(con,asset,user['id'],'تسعير تلقائي عند الحصر')
    con.commit();con.close();return RedirectResponse('/fm-inventory',status_code=303)

@app.get('/fm-inventory/template.xlsx')
def fm_inventory_template(request: Request):
    require_permission(request,'fm_inventory');wb=Workbook();ws=wb.active;ws.title='حصر الأصول'
    ws.append(['الرقم الوزاري','نوع الأصل','اسم الأصل','الشركة المصنعة','الموديل','الرقم التسلسلي',
      'الكمية','الوحدة','حالة الأصل','درجة الأهمية','مكان الأصل','تاريخ الحصر','القائم بالحصر','ملاحظات'])
    ws.append(['47138','تكييف','مكيف سبليت','LG','S18','','1','عدد','جيدة','مرتفع','الفصل 1','','',''])
    for c in ws[1]:c.font=c.font.copy(bold=True,color='FFFFFF');c.fill=c.fill.copy(fill_type='solid',fgColor='082B60')
    ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
    bio=BytesIO();wb.save(bio);bio.seek(0)
    return StreamingResponse(bio,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      headers={'Content-Disposition':'attachment; filename="fm_asset_inventory_template.xlsx"'})

@app.post('/fm-inventory/import')
async def fm_inventory_import(request: Request,files: list[UploadFile]=File(...),replace_existing: str=Form('')):
    user=require_permission(request,'fm_inventory');con=db();now=datetime.utcnow().isoformat()
    batch_no='INV-'+datetime.utcnow().strftime('%Y%m%d%H%M%S%f');source_names=[];added=updated=errors=total_rows=0
    try:
        if replace_existing=='yes':con.execute("UPDATE fm_assets SET is_active=0,updated_at=?",(now,))
        batch_cur=con.execute("""INSERT INTO fm_inventory_batches(batch_no,batch_name,source_files,imported_by,
          imported_by_name,imported_at) VALUES(?,?,?,?,?,?)""",(batch_no,'حصر متعدد الملفات','',user['id'],user['username'],now))
        batch_id=batch_cur.lastrowid
        if not batch_id:batch_id=con.execute("SELECT id FROM fm_inventory_batches WHERE batch_no=?",(batch_no,)).fetchone()['id']
        for upload in files:
            source_names.append(upload.filename or 'unknown.xlsx');content=await upload.read()
            try:wb=load_workbook(BytesIO(content),data_only=True,read_only=True)
            except Exception:errors+=1;continue
            for ws in wb.worksheets:
                headers=[_fm_header(c.value) for c in ws[1]]
                mid_col=_fm_find_col(headers,'الرقم الوزاري','رقم المدرسة','ministry_id')
                type_col=_fm_find_col(headers,'نوع الأصل','فئة الأصل','asset_type')
                if not mid_col or not type_col:continue
                names={
                  'asset_name':_fm_find_col(headers,'اسم الأصل','الأصل','asset_name'),
                  'manufacturer':_fm_find_col(headers,'الشركة المصنعة','الماركة','manufacturer'),
                  'model':_fm_find_col(headers,'الموديل','model'),
                  'serial_number':_fm_find_col(headers,'الرقم التسلسلي','السيريال','serial_number'),
                  'quantity':_fm_find_col(headers,'الكمية','quantity'),
                  'unit':_fm_find_col(headers,'الوحدة','unit'),
                  'condition_status':_fm_find_col(headers,'حالة الأصل','الحالة','condition_status'),
                  'criticality':_fm_find_col(headers,'درجة الأهمية','الخطورة','criticality'),
                  'location_description':_fm_find_col(headers,'مكان الأصل','الموقع داخل المدرسة','location'),
                  'inventory_date':_fm_find_col(headers,'تاريخ الحصر','inventory_date'),
                  'inventoried_by':_fm_find_col(headers,'القائم بالحصر','الحاصر','inventoried_by'),
                  'notes':_fm_find_col(headers,'ملاحظات','notes')
                }
                row_no=1
                for values in ws.iter_rows(min_row=2,values_only=True):
                    row_no+=1;row=dict(zip(headers,values));mid=_fm_cell(row.get(mid_col));atype=_fm_cell(row.get(type_col))
                    if not mid or not atype:errors+=1;continue
                    school=con.execute("SELECT * FROM schools WHERE ministry_id=?",(mid,)).fetchone()
                    if not school:errors+=1;continue
                    data={k:_fm_cell(row.get(v)) if v else '' for k,v in names.items()}
                    serial=data['serial_number'];asset_name=data['asset_name']
                    existing=None
                    if serial:existing=con.execute("SELECT * FROM fm_assets WHERE ministry_id=? AND serial_number=? AND is_active=1",(mid,serial)).fetchone()
                    if not existing:
                        existing=con.execute("""SELECT * FROM fm_assets WHERE ministry_id=? AND asset_type=?
                          AND COALESCE(asset_name,'')=? AND COALESCE(location_description,'')=? AND is_active=1
                          ORDER BY id LIMIT 1""",(mid,atype,asset_name,data['location_description'])).fetchone()
                    quantity=_fm_float(data['quantity'],1);condition=data['condition_status'] or 'جيدة'
                    if existing:
                        con.execute("""UPDATE fm_assets SET asset_name=?,manufacturer=?,model=?,serial_number=?,
                          quantity=?,unit=?,condition_status=?,criticality=?,location_description=?,
                          inventory_batch_id=?,inventory_date=?,inventoried_by=?,notes=?,source_file=?,source_row=?,
                          is_active=1,updated_at=? WHERE id=?""",(asset_name,data['manufacturer'],data['model'],serial,
                          quantity,data['unit'] or 'عدد',condition,data['criticality'] or 'متوسط',
                          data['location_description'],batch_id,data['inventory_date'] or date.today().isoformat(),
                          data['inventoried_by'] or user['username'],data['notes'],upload.filename,row_no,now,existing['id']))
                        asset_id=existing['id'];updated+=1
                    else:
                        code='AST-'+datetime.utcnow().strftime('%Y%m%d%H%M%S%f')
                        cur=con.execute("""INSERT INTO fm_assets(asset_code,ministry_id,school_name,city,asset_type,
                          asset_name,manufacturer,model,serial_number,condition_status,criticality,location_description,
                          qr_code,quantity,unit,inventory_batch_id,inventory_date,inventoried_by,notes,source_file,
                          source_row,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                          (code,mid,school['school_name'],school['city'],atype,asset_name,data['manufacturer'],
                           data['model'],serial,condition,data['criticality'] or 'متوسط',
                           data['location_description'],code,quantity,data['unit'] or 'عدد',batch_id,
                           data['inventory_date'] or date.today().isoformat(),data['inventoried_by'] or user['username'],
                           data['notes'],upload.filename,row_no,now,now))
                        asset_id=cur.lastrowid
                        if not asset_id:asset_id=con.execute("SELECT id FROM fm_assets WHERE asset_code=?",(code,)).fetchone()['id']
                        added+=1
                    asset=con.execute("SELECT * FROM fm_assets WHERE id=?",(asset_id,)).fetchone()
                    _fm_price_asset(con,asset,user['id'],f'تسعير تلقائي من الملف {upload.filename}')
                    total_rows+=1
            wb.close()
        con.execute("""UPDATE fm_inventory_batches SET source_files=?,total_rows=?,added_count=?,
          updated_count=?,error_count=? WHERE id=?""",(', '.join(source_names),total_rows,added,updated,errors,batch_id))
        con.commit()
    except Exception as exc:
        con.rollback();con.close();raise HTTPException(500,detail=f'فشل استيراد الحصر: {exc}')
    con.close();return RedirectResponse(f'/fm-inventory?added={added}&updated={updated}&errors={errors}',status_code=303)

@app.post('/fm-inventory/{asset_id}/photo')
async def fm_inventory_photo(request: Request,asset_id: int,photo: UploadFile=File(...),caption: str=Form('')):
    require_permission(request,'fm_inventory');con=db();asset=con.execute("SELECT * FROM fm_assets WHERE id=?",(asset_id,)).fetchone()
    if not asset:con.close();raise HTTPException(404)
    safe=_fm_safe_filename(photo.filename);target=FM_ASSET_UPLOAD_DIR/safe;target.write_bytes(await photo.read())
    path=f'/static/uploads/fm_assets/{safe}'
    con.execute("UPDATE fm_assets SET photo_path=?,photo_caption=?,updated_at=? WHERE id=?",(path,caption,datetime.utcnow().isoformat(),asset_id))
    con.commit();con.close();return RedirectResponse('/fm-inventory',status_code=303)

@app.get('/fm-pricing',response_class=HTMLResponse)
def fm_pricing_page(request: Request,asset_type: str='',condition_status: str='',q: str=''):
    user=require_permission(request,'fm_pricing');con=db();where=['is_active=1'];args=[]
    if asset_type:where.append('asset_type=?');args.append(asset_type)
    if condition_status:where.append('condition_status=?');args.append(condition_status)
    if q:where.append('(price_code LIKE ? OR asset_name LIKE ? OR manufacturer LIKE ? OR model LIKE ?)');args.extend([f'%{q}%']*4)
    rows=con.execute("SELECT * FROM fm_price_catalog WHERE "+' AND '.join(where)+" ORDER BY asset_type,asset_name,condition_status",args).fetchall()
    types=[r['asset_type'] for r in con.execute("SELECT DISTINCT asset_type FROM fm_price_catalog WHERE is_active=1 ORDER BY asset_type")]
    summary=con.execute("""SELECT COUNT(*) price_rows,COALESCE(SUM(estimated_total_price),0) total_value,
      SUM(CASE WHEN pricing_status='تم التسعير' THEN 1 ELSE 0 END) priced_assets,
      SUM(CASE WHEN pricing_status<>'تم التسعير' OR pricing_status IS NULL THEN 1 ELSE 0 END) unpriced_assets
      FROM fm_assets WHERE is_active=1""").fetchone()
    con.close();return templates.TemplateResponse('fm_pricing.html',{'request':request,'user':user,'rows':rows,'types':types,'summary':summary,'asset_type':asset_type,'condition_status':condition_status,'q':q})

@app.post('/fm-pricing/create')
def fm_pricing_create(request: Request,asset_type: str=Form(...),asset_name: str=Form(''),
  manufacturer: str=Form(''),model: str=Form(''),condition_status: str=Form('الكل'),
  unit: str=Form('عدد'),base_price: float=Form(...),condition_factor: float=Form(1),
  effective_from: str=Form(''),effective_to: str=Form(''),notes: str=Form('')):
    user=require_permission(request,'fm_pricing');con=db();code='PRC-'+datetime.utcnow().strftime('%Y%m%d%H%M%S%f')
    calculated=base_price*condition_factor;now=datetime.utcnow().isoformat()
    con.execute("""INSERT INTO fm_price_catalog(price_code,asset_type,asset_name,manufacturer,model,
      condition_status,unit,base_price,condition_factor,calculated_unit_price,effective_from,effective_to,
      notes,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(code,asset_type,asset_name,
      manufacturer,model,condition_status,unit,base_price,condition_factor,calculated,effective_from,
      effective_to,notes,now,now))
    _fm_reprice_all(con,user['id']);con.commit();con.close();return RedirectResponse('/fm-pricing',status_code=303)

@app.get('/fm-pricing/template.xlsx')
def fm_pricing_template(request: Request):
    require_permission(request,'fm_pricing');wb=Workbook();ws=wb.active;ws.title='قائمة أسعار الأصول'
    ws.append(['نوع الأصل','اسم الأصل','الشركة المصنعة','الموديل','حالة الأصل','الوحدة',
      'السعر الأساسي','معامل الحالة','السعر المحسوب','سريان من','سريان إلى','ملاحظات'])
    ws.append(['تكييف','مكيف سبليت','LG','S18','الكل','عدد',2000,1,2000,'','','السعر المرجعي'])
    ws.append(['تكييف','مكيف سبليت','LG','S18','جيدة','عدد',2000,0.85,1700,'','','سعر الحالة الجيدة'])
    for c in ws[1]:c.font=c.font.copy(bold=True,color='FFFFFF');c.fill=c.fill.copy(fill_type='solid',fgColor='082B60')
    ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
    bio=BytesIO();wb.save(bio);bio.seek(0)
    return StreamingResponse(bio,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      headers={'Content-Disposition':'attachment; filename="fm_asset_pricing_template.xlsx"'})

@app.post('/fm-pricing/import')
async def fm_pricing_import(request: Request,file: UploadFile=File(...),replace_old: str=Form('')):
    user=require_permission(request,'fm_pricing');content=await file.read()
    try:wb=load_workbook(BytesIO(content),data_only=True,read_only=True)
    except Exception as exc:raise HTTPException(400,detail=f'تعذر قراءة ملف الأسعار: {exc}')
    con=db();added=updated=errors=0;now=datetime.utcnow().isoformat()
    try:
        if replace_old=='yes':con.execute("UPDATE fm_price_catalog SET is_active=0,updated_at=?",(now,))
        for ws in wb.worksheets:
            headers=[_fm_header(c.value) for c in ws[1]]
            type_col=_fm_find_col(headers,'نوع الأصل','فئة الأصل','asset_type')
            base_col=_fm_find_col(headers,'السعر الأساسي','السعر','base_price','unit_price')
            if not type_col or not base_col:continue
            cols={
              'asset_name':_fm_find_col(headers,'اسم الأصل','asset_name'),
              'manufacturer':_fm_find_col(headers,'الشركة المصنعة','الماركة','manufacturer'),
              'model':_fm_find_col(headers,'الموديل','model'),
              'condition_status':_fm_find_col(headers,'حالة الأصل','الحالة','condition_status'),
              'unit':_fm_find_col(headers,'الوحدة','unit'),
              'condition_factor':_fm_find_col(headers,'معامل الحالة','نسبة الحالة','condition_factor'),
              'calculated_unit_price':_fm_find_col(headers,'السعر المحسوب','سعر الحالة','calculated_unit_price'),
              'effective_from':_fm_find_col(headers,'سريان من','effective_from'),
              'effective_to':_fm_find_col(headers,'سريان إلى','effective_to'),
              'notes':_fm_find_col(headers,'ملاحظات','notes')
            }
            for values in ws.iter_rows(min_row=2,values_only=True):
                row=dict(zip(headers,values));atype=_fm_cell(row.get(type_col));base=_fm_float(row.get(base_col),0)
                if not atype or base<=0:errors+=1;continue
                data={k:_fm_cell(row.get(v)) if v else '' for k,v in cols.items()}
                factor=_fm_float(data['condition_factor'],0)
                condition=data['condition_status'] or 'الكل'
                if factor<=0:factor=_fm_condition_default_factor(condition)
                calculated=_fm_float(data['calculated_unit_price'],0) or base*factor
                existing=con.execute("""SELECT * FROM fm_price_catalog WHERE is_active=1 AND asset_type=?
                  AND COALESCE(asset_name,'')=? AND COALESCE(manufacturer,'')=? AND COALESCE(model,'')=?
                  AND COALESCE(condition_status,'الكل')=? ORDER BY id DESC LIMIT 1""",(atype,data['asset_name'],
                  data['manufacturer'],data['model'],condition)).fetchone()
                if existing:
                    con.execute("""UPDATE fm_price_catalog SET unit=?,base_price=?,condition_factor=?,
                      calculated_unit_price=?,effective_from=?,effective_to=?,notes=?,updated_at=?
                      WHERE id=?""",(data['unit'] or 'عدد',base,factor,calculated,data['effective_from'],
                      data['effective_to'],data['notes'],now,existing['id']));updated+=1
                else:
                    code='PRC-'+datetime.utcnow().strftime('%Y%m%d%H%M%S%f')
                    con.execute("""INSERT INTO fm_price_catalog(price_code,asset_type,asset_name,manufacturer,
                      model,condition_status,unit,base_price,condition_factor,calculated_unit_price,
                      effective_from,effective_to,notes,created_at,updated_at)
                      VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(code,atype,data['asset_name'],data['manufacturer'],
                      data['model'],condition,data['unit'] or 'عدد',base,factor,calculated,
                      data['effective_from'],data['effective_to'],data['notes'],now,now));added+=1
        result=_fm_reprice_all(con,user['id']);con.commit()
    except Exception as exc:
        con.rollback();con.close();wb.close();raise HTTPException(500,detail=f'فشل استيراد الأسعار: {exc}')
    con.close();wb.close()
    return RedirectResponse(f"/fm-pricing?added={added}&updated={updated}&errors={errors}&priced={result['matched']}",status_code=303)

@app.post('/fm-pricing/recalculate')
def fm_pricing_recalculate(request: Request):
    user=require_permission(request,'fm_pricing');con=db();result=_fm_reprice_all(con,user['id']);con.commit();con.close()
    return RedirectResponse(f"/fm-pricing?priced={result['matched']}&unpriced={result['unmatched']}",status_code=303)

@app.get('/fm-inventory-report.xlsx')
def fm_inventory_report(request: Request):
    require_permission(request,'fm_inventory');con=db()
    assets=con.execute("SELECT * FROM fm_assets WHERE is_active=1 ORDER BY city,school_name,asset_type").fetchall()
    employees=con.execute("SELECT * FROM fm_employees ORDER BY department,city,full_name").fetchall()
    prices=con.execute("SELECT * FROM fm_price_catalog WHERE is_active=1 ORDER BY asset_type,asset_name,condition_status").fetchall()
    con.close();wb=Workbook();ws=wb.active;ws.title='حصر وتسعير الأصول'
    ws.append(['كود الأصل','الرقم الوزاري','المدرسة','المدينة','نوع الأصل','اسم الأصل','الشركة','الموديل',
      'الرقم التسلسلي','الكمية','الوحدة','الحالة','مكان الأصل','القائم بالحصر','تاريخ الحصر',
      'سعر الوحدة','القيمة الإجمالية','حالة التسعير','رابط الصورة'])
    for a in assets:ws.append([a['asset_code'],a['ministry_id'],a['school_name'],a['city'],a['asset_type'],
      a['asset_name'],a['manufacturer'],a['model'],a['serial_number'],a['quantity'],a['unit'],
      a['condition_status'],a['location_description'],a['inventoried_by'],a['inventory_date'],
      a['estimated_unit_price'],a['estimated_total_price'],a['pricing_status'],a['photo_path']])
    es=wb.create_sheet('موظفو الصيانة والنظافة');es.append(['الرقم الوظيفي','الاسم','الهوية','الجوال','القسم',
      'المسمى','التخصص','المدينة','الفريق','المشرف','الحالة','انتهاء العقد'])
    for e in employees:es.append([e['emp_id'],e['full_name'],e['id_no'],e['mobile'],e['department'],e['job_title'],
      e['trade'],e['city'],e['team_name'],e['supervisor_name'],e['employment_status'],e['contract_expiry']])
    ps=wb.create_sheet('قائمة الأسعار');ps.append(['الكود','نوع الأصل','اسم الأصل','الشركة','الموديل','الحالة',
      'الوحدة','السعر الأساسي','معامل الحالة','سعر الوحدة'])
    for p in prices:ps.append([p['price_code'],p['asset_type'],p['asset_name'],p['manufacturer'],p['model'],
      p['condition_status'],p['unit'],p['base_price'],p['condition_factor'],p['calculated_unit_price']])
    for sh in wb.worksheets:
        for c in sh[1]:c.font=c.font.copy(bold=True,color='FFFFFF');c.fill=c.fill.copy(fill_type='solid',fgColor='082B60')
        sh.freeze_panes='A2';sh.auto_filter.ref=sh.dimensions
    bio=BytesIO();wb.save(bio);bio.seek(0)
    return StreamingResponse(bio,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      headers={'Content-Disposition':'attachment; filename="tbc_fm_inventory_pricing_report.xlsx"'})


@app.get('/fm-dashboard',response_class=HTMLResponse)
def fm_dashboard(request: Request):
    user=require_permission(request,'fm_dashboard');con=db();today=date.today().isoformat()
    scope,args=_fm_scope(user,'s')
    schools=con.execute(f"SELECT ministry_id FROM schools s WHERE COALESCE(s.status,'تعمل')='تعمل' {scope}",args).fetchall()
    school_ids=[r['ministry_id'] for r in schools]
    def count(sql,params=()):
        r=con.execute(sql,params).fetchone();return int(r[0] if r else 0)
    if school_ids:
        ph=','.join('?' for _ in school_ids)
        open_wo=count(f"SELECT COUNT(*) FROM fm_work_orders WHERE ministry_id IN ({ph}) AND status NOT IN ('مغلق','ملغي')",school_ids)
        critical=count(f"SELECT COUNT(*) FROM fm_work_orders WHERE ministry_id IN ({ph}) AND priority='حرج' AND status NOT IN ('مغلق','ملغي')",school_ids)
        overdue=count(f"SELECT COUNT(*) FROM fm_work_orders WHERE ministry_id IN ({ph}) AND status NOT IN ('مغلق','ملغي') AND datetime(opened_at,'+'||sla_hours||' hours')<datetime('now')",school_ids)
        assets=count(f"SELECT COUNT(*) FROM fm_assets WHERE ministry_id IN ({ph}) AND is_active=1",school_ids)
        pm_due=count(f"SELECT COUNT(*) FROM fm_pm_plans WHERE ministry_id IN ({ph}) AND is_active=1 AND date(next_due_date)<=date(?)",school_ids+[today])
        inspections=count(f"SELECT COUNT(*) FROM fm_cleaning_inspections WHERE ministry_id IN ({ph}) AND inspection_date=?",school_ids+[today])
        low_clean=count(f"SELECT COUNT(*) FROM fm_cleaning_inspections WHERE ministry_id IN ({ph}) AND inspection_date=? AND overall_score<70",school_ids+[today])
        work_orders=con.execute(f"SELECT * FROM fm_work_orders WHERE ministry_id IN ({ph}) ORDER BY id DESC LIMIT 20",school_ids).fetchall()
        cleaning=con.execute(f"SELECT * FROM fm_cleaning_inspections WHERE ministry_id IN ({ph}) ORDER BY id DESC LIMIT 15",school_ids).fetchall()
    else:
        open_wo=critical=overdue=assets=pm_due=inspections=low_clean=0;work_orders=[];cleaning=[]
    con.close()
    return templates.TemplateResponse('fm_dashboard.html',{'request':request,'user':user,'open_wo':open_wo,'critical':critical,'overdue':overdue,'assets':assets,'pm_due':pm_due,'inspections':inspections,'low_clean':low_clean,'work_orders':work_orders,'cleaning':cleaning})

@app.get('/maintenance-work-orders',response_class=HTMLResponse)
def maintenance_work_orders(request: Request,status: str='',category: str='',priority: str='',city: str=''):
    user=require_permission(request,'maintenance_work_orders');con=db();scope,args=_fm_scope(user,'s')
    where=["1=1"];params=[]
    if status:where.append('w.status=?');params.append(status)
    if category:where.append('w.category=?');params.append(category)
    if priority:where.append('w.priority=?');params.append(priority)
    if city:where.append('w.city=?');params.append(city)
    rows=con.execute(f"""SELECT w.* FROM fm_work_orders w LEFT JOIN schools s ON s.ministry_id=w.ministry_id
      WHERE {' AND '.join(where)} {scope} ORDER BY CASE w.priority WHEN 'حرج' THEN 1 WHEN 'مرتفع' THEN 2 WHEN 'متوسط' THEN 3 ELSE 4 END,w.id DESC""",params+args).fetchall()
    schools=con.execute(f"SELECT s.ministry_id,s.school_name,s.city FROM schools s WHERE COALESCE(s.status,'تعمل')='تعمل' {scope} ORDER BY s.city,s.school_name",args).fetchall()
    cities=[r['city'] for r in con.execute("SELECT DISTINCT city FROM schools WHERE COALESCE(city,'')<>'' ORDER BY city")]
    con.close();return templates.TemplateResponse('maintenance_work_orders.html',{'request':request,'user':user,'rows':rows,'schools':schools,'cities':cities,'status':status,'category':category,'priority':priority,'city':city})

@app.post('/maintenance-work-orders/create')
def maintenance_work_order_create(request: Request,ministry_id: str=Form(...),category: str=Form(...),subcategory: str=Form(''),title: str=Form(...),description: str=Form(''),priority: str=Form('متوسط'),assigned_to: str=Form(''),technician_mobile: str=Form(''),sla_hours: int=Form(24)):
    user=require_permission(request,'maintenance_work_orders');con=db();school=con.execute('SELECT * FROM schools WHERE ministry_id=?',(ministry_id,)).fetchone()
    if not school:con.close();raise HTTPException(404,detail='المدرسة غير موجودة')
    no='WO-'+datetime.utcnow().strftime('%Y%m%d%H%M%S%f')
    con.execute("""INSERT INTO fm_work_orders(work_order_no,ministry_id,school_name,city,category,subcategory,title,description,priority,status,assigned_to,technician_mobile,sla_hours,opened_by,opened_by_name,opened_at)
      VALUES(?,?,?,?,?,?,?,?,?,'جديد',?,?,?,?,?,?)""",(no,ministry_id,school['school_name'],school['city'],category,subcategory,title,description,priority,assigned_to,technician_mobile,sla_hours,user['id'],user['username'],datetime.utcnow().isoformat()))
    con.commit();con.close();return RedirectResponse('/maintenance-work-orders',status_code=303)

@app.post('/maintenance-work-orders/{wo_id}/update')
def maintenance_work_order_update(request: Request,wo_id: int,status: str=Form(...),assigned_to: str=Form(''),resolution_notes: str=Form(''),cost: float=Form(0),materials_cost: float=Form(0)):
    user=require_permission(request,'maintenance_work_orders');con=db();wo=con.execute('SELECT * FROM fm_work_orders WHERE id=?',(wo_id,)).fetchone()
    if not wo:con.close();raise HTTPException(404)
    now=datetime.utcnow().isoformat();timestamps={'مقبول':'accepted_at','قيد التنفيذ':'started_at','تم الإصلاح':'resolved_at','مغلق':'closed_at'};extra='';params=[status,assigned_to,resolution_notes,cost,materials_cost]
    if status in timestamps:extra=f",{timestamps[status]}=?";params.append(now)
    params.append(wo_id);con.execute(f"UPDATE fm_work_orders SET status=?,assigned_to=?,resolution_notes=?,cost=?,materials_cost=? {extra} WHERE id=?",params)
    con.commit();con.close();return RedirectResponse('/maintenance-work-orders',status_code=303)

@app.get('/fm-assets',response_class=HTMLResponse)
def fm_assets_page(request: Request,asset_type: str='',city: str=''):
    user=require_permission(request,'fm_assets');con=db();scope,args=_fm_scope(user,'s');where=['a.is_active=1'];params=[]
    if asset_type:where.append('a.asset_type=?');params.append(asset_type)
    if city:where.append('a.city=?');params.append(city)
    rows=con.execute(f"SELECT a.* FROM fm_assets a LEFT JOIN schools s ON s.ministry_id=a.ministry_id WHERE {' AND '.join(where)} {scope} ORDER BY a.city,a.school_name,a.asset_type",params+args).fetchall()
    schools=con.execute(f"SELECT s.ministry_id,s.school_name,s.city FROM schools s WHERE COALESCE(s.status,'تعمل')='تعمل' {scope} ORDER BY s.city,s.school_name",args).fetchall();con.close()
    return templates.TemplateResponse('fm_assets.html',{'request':request,'user':user,'rows':rows,'schools':schools,'asset_type':asset_type,'city':city})

@app.post('/fm-assets/create')
def fm_asset_create(request: Request,ministry_id: str=Form(...),asset_type: str=Form(...),asset_name: str=Form(''),manufacturer: str=Form(''),model: str=Form(''),serial_number: str=Form(''),condition_status: str=Form('جيدة'),criticality: str=Form('متوسط'),location_description: str=Form(''),warranty_expiry: str=Form(''),next_maintenance_date: str=Form('')):
    require_permission(request,'fm_assets');con=db();school=con.execute('SELECT * FROM schools WHERE ministry_id=?',(ministry_id,)).fetchone()
    if not school:con.close();raise HTTPException(404)
    code='AST-'+datetime.utcnow().strftime('%Y%m%d%H%M%S%f')
    con.execute("""INSERT INTO fm_assets(asset_code,ministry_id,school_name,city,asset_type,asset_name,manufacturer,model,serial_number,warranty_expiry,condition_status,criticality,location_description,qr_code,next_maintenance_date,created_at)
      VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(code,ministry_id,school['school_name'],school['city'],asset_type,asset_name,manufacturer,model,serial_number,warranty_expiry,condition_status,criticality,location_description,code,next_maintenance_date,datetime.utcnow().isoformat()))
    con.commit();con.close();return RedirectResponse('/fm-assets',status_code=303)

@app.get('/preventive-maintenance',response_class=HTMLResponse)
def preventive_maintenance_page(request: Request):
    user=require_permission(request,'preventive_maintenance');con=db();plans=con.execute('SELECT p.*,a.asset_code,a.asset_name,s.school_name FROM fm_pm_plans p LEFT JOIN fm_assets a ON a.id=p.asset_id LEFT JOIN schools s ON s.ministry_id=p.ministry_id WHERE p.is_active=1 ORDER BY p.next_due_date').fetchall();assets=con.execute('SELECT id,asset_code,asset_name,asset_type,ministry_id,school_name FROM fm_assets WHERE is_active=1 ORDER BY school_name,asset_type').fetchall();con.close();return templates.TemplateResponse('preventive_maintenance.html',{'request':request,'user':user,'plans':plans,'assets':assets})

@app.post('/preventive-maintenance/create')
def preventive_maintenance_create(request: Request,plan_name: str=Form(...),asset_id: int=Form(...),frequency_days: int=Form(...),checklist: str=Form(''),assigned_team: str=Form(''),priority: str=Form('متوسط'),next_due_date: str=Form(...)):
    require_permission(request,'preventive_maintenance');con=db();asset=con.execute('SELECT * FROM fm_assets WHERE id=?',(asset_id,)).fetchone()
    if not asset:con.close();raise HTTPException(404)
    con.execute("INSERT INTO fm_pm_plans(plan_name,asset_id,ministry_id,asset_type,frequency_days,checklist,assigned_team,priority,next_due_date,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)",(plan_name,asset_id,asset['ministry_id'],asset['asset_type'],frequency_days,checklist,assigned_team,priority,next_due_date,datetime.utcnow().isoformat()))
    con.commit();con.close();return RedirectResponse('/preventive-maintenance',status_code=303)

@app.post('/preventive-maintenance/generate')
def preventive_maintenance_generate(request: Request):
    user=require_permission(request,'preventive_maintenance');con=db();today=date.today().isoformat();plans=con.execute("SELECT p.*,a.school_name,a.city FROM fm_pm_plans p LEFT JOIN fm_assets a ON a.id=p.asset_id WHERE p.is_active=1 AND date(p.next_due_date)<=date(?)",(today,)).fetchall();generated=0
    for p in plans:
        no='WO-PM-'+datetime.utcnow().strftime('%Y%m%d%H%M%S%f')
        con.execute("""INSERT INTO fm_work_orders(work_order_no,ministry_id,school_name,city,category,subcategory,title,description,priority,status,assigned_to,sla_hours,opened_by,opened_by_name,opened_at)
          VALUES(?,?,?,?, 'صيانة وقائية',?,?,?,?, 'جديد',?,72,?,?,?)""",(no,p['ministry_id'],p['school_name'],p['city'],p['asset_type'],p['plan_name'],p['checklist'],p['priority'],p['assigned_team'],user['id'],user['username'],datetime.utcnow().isoformat()))
        next_due=(date.fromisoformat(p['next_due_date'])+timedelta(days=int(p['frequency_days']))).isoformat();con.execute('UPDATE fm_pm_plans SET next_due_date=?,last_generated_at=? WHERE id=?',(next_due,datetime.utcnow().isoformat(),p['id']));generated+=1
    con.commit();con.close();return RedirectResponse(f'/preventive-maintenance?generated={generated}',status_code=303)

@app.get('/cleaning-quality',response_class=HTMLResponse)
def cleaning_quality_page(request: Request,city: str='',date_from: str='',date_to: str=''):
    user=require_permission(request,'cleaning_quality');con=db();scope,args=_fm_scope(user,'s');where=['1=1'];params=[]
    if city:where.append('c.city=?');params.append(city)
    if date_from:where.append('c.inspection_date>=?');params.append(date_from)
    if date_to:where.append('c.inspection_date<=?');params.append(date_to)
    rows=con.execute(f"SELECT c.* FROM fm_cleaning_inspections c LEFT JOIN schools s ON s.ministry_id=c.ministry_id WHERE {' AND '.join(where)} {scope} ORDER BY c.inspection_date DESC,c.id DESC",params+args).fetchall();schools=con.execute(f"SELECT s.ministry_id,s.school_name,s.city FROM schools s WHERE COALESCE(s.status,'تعمل')='تعمل' {scope} ORDER BY s.city,s.school_name",args).fetchall();con.close();return templates.TemplateResponse('cleaning_quality.html',{'request':request,'user':user,'rows':rows,'schools':schools,'city':city,'date_from':date_from,'date_to':date_to})

@app.post('/cleaning-quality/create')
def cleaning_quality_create(request: Request,inspection_date: str=Form(...),ministry_id: str=Form(...),inspector_name: str=Form(...),cleaning_team: str=Form(''),attendance_count: int=Form(0),classrooms_score: int=Form(...),toilets_score: int=Form(...),corridors_score: int=Form(...),yard_score: int=Form(...),offices_score: int=Form(...),findings: str=Form(''),corrective_action: str=Form(''),follow_up_date: str=Form('')):
    user=require_permission(request,'cleaning_quality');con=db();school=con.execute('SELECT * FROM schools WHERE ministry_id=?',(ministry_id,)).fetchone()
    if not school:con.close();raise HTTPException(404)
    scores=[classrooms_score,toilets_score,corridors_score,yard_score,offices_score];overall=sum(scores)/len(scores);status='ممتاز' if overall>=90 else 'جيد' if overall>=80 else 'مقبول' if overall>=70 else 'يحتاج معالجة';no='CLN-'+datetime.utcnow().strftime('%Y%m%d%H%M%S%f')
    con.execute("""INSERT INTO fm_cleaning_inspections(inspection_no,inspection_date,ministry_id,school_name,city,inspector_name,cleaning_team,attendance_count,classrooms_score,toilets_score,corridors_score,yard_score,offices_score,overall_score,status,findings,corrective_action,follow_up_date,recorded_by,recorded_at)
      VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(no,inspection_date,ministry_id,school['school_name'],school['city'],inspector_name,cleaning_team,attendance_count,*scores,overall,status,findings,corrective_action,follow_up_date,user['id'],datetime.utcnow().isoformat()))
    if overall<70:
        wo='WO-CLN-'+datetime.utcnow().strftime('%Y%m%d%H%M%S%f');con.execute("""INSERT INTO fm_work_orders(work_order_no,ministry_id,school_name,city,category,title,description,priority,status,sla_hours,opened_by,opened_by_name,opened_at) VALUES(?,?,?,?, 'نظافة','معالجة تدني مستوى النظافة',?,'مرتفع','جديد',12,?,?,?)""",(wo,ministry_id,school['school_name'],school['city'],findings,user['id'],user['username'],datetime.utcnow().isoformat()))
    con.commit();con.close();return RedirectResponse('/cleaning-quality',status_code=303)

@app.get('/technician-mobile',response_class=HTMLResponse)
def technician_mobile_page(request: Request):
    user=require_permission(request,'technician_mobile');con=db();rows=con.execute("SELECT * FROM fm_work_orders WHERE status NOT IN ('مغلق','ملغي') ORDER BY CASE priority WHEN 'حرج' THEN 1 WHEN 'مرتفع' THEN 2 ELSE 3 END,id DESC LIMIT 150").fetchall();con.close();return templates.TemplateResponse('technician_mobile.html',{'request':request,'user':user,'rows':rows})

@app.get('/fm-report.xlsx')
def fm_report_excel(request: Request):
    require_permission(request,'fm_dashboard');con=db();wo=con.execute('SELECT * FROM fm_work_orders ORDER BY id DESC').fetchall();clean=con.execute('SELECT * FROM fm_cleaning_inspections ORDER BY id DESC').fetchall();assets=con.execute('SELECT * FROM fm_assets WHERE is_active=1 ORDER BY city,school_name').fetchall();con.close();wb=Workbook();ws=wb.active;ws.title='أوامر الصيانة';h=['الرقم','المدرسة','المدينة','الفئة','العنوان','الأولوية','الحالة','المكلف','وقت الفتح','التكلفة'];ws.append(h)
    for r in wo:ws.append([r['work_order_no'],r['school_name'],r['city'],r['category'],r['title'],r['priority'],r['status'],r['assigned_to'],r['opened_at'],r['cost']+r['materials_cost']])
    cs=wb.create_sheet('تقييم النظافة');cs.append(['الرقم','التاريخ','المدرسة','المدينة','المفتش','العمال','الدرجة','الحالة','الملاحظات'])
    for r in clean:cs.append([r['inspection_no'],r['inspection_date'],r['school_name'],r['city'],r['inspector_name'],r['attendance_count'],r['overall_score'],r['status'],r['findings']])
    ass=wb.create_sheet('الأصول');ass.append(['الكود','المدرسة','المدينة','النوع','الاسم','الحالة','الخطورة','الصيانة القادمة'])
    for r in assets:ass.append([r['asset_code'],r['school_name'],r['city'],r['asset_type'],r['asset_name'],r['condition_status'],r['criticality'],r['next_maintenance_date']])
    for sh in wb.worksheets:
        for c in sh[1]:c.font=c.font.copy(bold=True,color='FFFFFF');c.fill=c.fill.copy(fill_type='solid',fgColor='082B60')
        sh.freeze_panes='A2';sh.auto_filter.ref=sh.dimensions
    bio=BytesIO();wb.save(bio);bio.seek(0);return StreamingResponse(bio,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={'Content-Disposition':'attachment; filename="tbc_facilities_report.xlsx"'})


@app.get("/reports", response_class=HTMLResponse)
def reports_page(request: Request):
    user=require_permission(request,"reports")
    return templates.TemplateResponse("reports.html",{"request":request,"user":user})



@app.get("/period-report", response_class=HTMLResponse)
def period_report_page(
    request: Request,
    date_from: str="",
    date_to: str="",
    city: str="",
    coordinator: str="",
    school_scope: str="all"
):
    user=require_permission(request,"period_report")
    today=date.today()
    if not date_from:
        date_from=today.replace(day=1).isoformat()
    if not date_to:
        date_to=today.isoformat()
    con=db()
    clause,args=scope_clause(user)
    filters=""
    if city:
        filters+=" AND s.city=?"; args.append(city)
    if coordinator and user["role"]!="coordinator":
        filters+=" AND s.coordinator_name=?"; args.append(coordinator)

    rows=con.execute(f"""
      SELECT s.city,s.coordinator_name,s.ministry_id,s.school_name,s.guard_name,s.guard_emp_id,
             a.work_date,COALESCE(a.status,'لم يسجل') status,a.notes,a.replacement_name
      FROM schools s
      LEFT JOIN attendance a
        ON a.ministry_id=s.ministry_id AND a.work_date BETWEEN ? AND ?
      WHERE 1=1 {clause} {filters}
      ORDER BY s.city,s.coordinator_name,s.school_name,a.work_date
    """,[date_from,date_to]+args).fetchall()

    summary=con.execute(f"""
      WITH RECURSIVE dates(d) AS (
        SELECT date(?)
        UNION ALL SELECT date(d,'+1 day') FROM dates WHERE d < date(?)
      ),
      base AS (
        SELECT d.d work_date,s.city,s.coordinator_name,s.ministry_id
        FROM dates d CROSS JOIN schools s
        WHERE 1=1 {clause} {filters}
      )
      SELECT city,coordinator_name,
             COUNT(*) total_slots,
             SUM(CASE WHEN COALESCE(a.status,'لم يسجل')='حاضر' THEN 1 ELSE 0 END) present_count,
             SUM(CASE WHEN COALESCE(a.status,'لم يسجل')='غائب' THEN 1 ELSE 0 END) absent_count,
             SUM(CASE WHEN COALESCE(a.status,'لم يسجل')='إجازة' THEN 1 ELSE 0 END) leave_count,
             SUM(CASE WHEN COALESCE(a.status,'لم يسجل')='بديل' THEN 1 ELSE 0 END) replacement_count,
             SUM(CASE WHEN a.status IS NULL OR a.status='لم يسجل' THEN 1 ELSE 0 END) missing_count
      FROM base b
      LEFT JOIN attendance a ON a.ministry_id=b.ministry_id AND a.work_date=b.work_date
      GROUP BY city,coordinator_name
      ORDER BY city,coordinator_name
    """,[date_from,date_to]+args).fetchall()

    cities=[r["city"] for r in con.execute("SELECT DISTINCT city FROM schools WHERE city<>'' ORDER BY city")]
    coords=[r["coordinator_name"] for r in con.execute("SELECT DISTINCT coordinator_name FROM schools WHERE coordinator_name<>'' ORDER BY coordinator_name")]
    con.close()
    return templates.TemplateResponse("period_report.html",{
      "request":request,"user":user,"rows":rows,"summary":summary,
      "date_from":date_from,"date_to":date_to,"city":city,"coordinator":coordinator,"school_scope":school_scope,
      "cities":cities,"coords":coords
    })

@app.get("/period-report.xlsx")
def period_report_excel(
    request: Request,
    date_from: str="",
    date_to: str="",
    city: str="",
    coordinator: str="",
    school_scope: str="all"
):
    user=current_user(request)
    if not user: raise HTTPException(401)
    today=date.today()
    date_from=date_from or today.replace(day=1).isoformat()
    date_to=date_to or today.isoformat()
    con=db()
    clause,args=scope_clause(user)
    filters=""
    if city: filters+=" AND s.city=?";args.append(city)
    if coordinator and user["role"]!="coordinator": filters+=" AND s.coordinator_name=?";args.append(coordinator)
    if school_scope=="main": filters+=" AND TRIM(COALESCE(s.operational_type,''))='رئيسية'"
    elif school_scope=="main_branch": filters+=" AND TRIM(COALESCE(s.operational_type,'')) IN ('رئيسية','فرعية')"
    elif school_scope=="branch": filters+=" AND TRIM(COALESCE(s.operational_type,''))='فرعية'"
    elif school_scope=="annex": filters+=" AND TRIM(COALESCE(s.operational_type,''))='ملحقة'"

    rows=con.execute(f"""
      WITH RECURSIVE dates(d) AS (
        SELECT date(?)
        UNION ALL SELECT date(d,'+1 day') FROM dates WHERE d < date(?)
      )
      SELECT d.d work_date,s.city,s.coordinator_name,s.ministry_id,s.school_name,
             s.guard_name,s.guard_emp_id,s.guard_mobile,
             COALESCE(a.status,'لم يسجل') status,a.replacement_name,a.notes,a.recorded_at
      FROM dates d CROSS JOIN schools s
      LEFT JOIN attendance a ON a.ministry_id=s.ministry_id AND a.work_date=d.d
      WHERE 1=1 {clause} {filters}
      ORDER BY d.d,s.city,s.coordinator_name,s.school_name
    """,[date_from,date_to]+args).fetchall()
    con.close()

    wb=Workbook()
    ws=wb.active
    ws.title="تفاصيل الحضور"
    headers=["التاريخ","المدينة","المنسق","الرقم الوزاري","المدرسة","البواب","الرقم الوظيفي","جوال البواب","الحالة","البديل","الملاحظات","وقت التسجيل"]
    ws.append(headers)
    for r in rows:
        ws.append([r["work_date"],r["city"],r["coordinator_name"],r["ministry_id"],r["school_name"],r["guard_name"],r["guard_emp_id"],r["guard_mobile"],r["status"],r["replacement_name"],r["notes"],r["recorded_at"]])

    summary=wb.create_sheet("ملخص المشرفين")
    summary.append(["المدينة","المنسق","إجمالي السجلات","حاضر","غائب","إجازة","بديل","لم يسجل","نسبة الحضور","نسبة التغطية"])
    groups={}
    for r in rows:
        key=(r["city"],r["coordinator_name"])
        g=groups.setdefault(key,{"total":0,"حاضر":0,"غائب":0,"إجازة":0,"بديل":0,"لم يسجل":0})
        g["total"]+=1
        g[r["status"] if r["status"] in g else "لم يسجل"]+=1
    for (city_name,coord),g in groups.items():
        attendance_pct=(g["حاضر"]/g["total"]*100) if g["total"] else 0
        coverage_pct=((g["حاضر"]+g["بديل"])/g["total"]*100) if g["total"] else 0
        summary.append([city_name,coord,g["total"],g["حاضر"],g["غائب"],g["إجازة"],g["بديل"],g["لم يسجل"],attendance_pct,coverage_pct])

    for sheet in [ws,summary]:
        for cell in sheet[1]:
            cell.font=cell.font.copy(bold=True,color="FFFFFF")
            cell.fill=cell.fill.copy(fill_type="solid",fgColor="082B60")
        sheet.freeze_panes="A2"
    for col in ["A","D","G","H"]:
        for cell in ws[col]: cell.number_format="@"
    for row in summary.iter_rows(min_row=2,min_col=9,max_col=10):
        for cell in row: cell.number_format='0.0%'

    bio=BytesIO();wb.save(bio);bio.seek(0)
    return StreamingResponse(
      bio,
      media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      headers={"Content-Disposition":f'attachment; filename="attendance_{date_from}_to_{date_to}.xlsx"'}
    )

@app.get("/violations", response_class=HTMLResponse)
def violations_page(request: Request, guard_emp_id: str="", city: str="", date_from: str="", date_to: str=""):
    user=require_permission(request,"violations")
    con=db()
    where=["1=1"];args=[]
    if user["role"]=="coordinator":
        where.append("coordinator_name=?");args.append(user["coordinator_name"])
    if guard_emp_id:
        where.append("(guard_emp_id LIKE ? OR guard_name LIKE ?)");like=f"%{guard_emp_id}%";args += [like,like]
    if city: where.append("city=?");args.append(city)
    if date_from: where.append("violation_date>=?");args.append(date_from)
    if date_to: where.append("violation_date<=?");args.append(date_to)
    rows=con.execute("SELECT * FROM violations WHERE "+" AND ".join(where)+" ORDER BY violation_date DESC,id DESC",args).fetchall()
    schools=con.execute("SELECT ministry_id,school_name,city,guard_name,guard_emp_id,coordinator_name FROM schools ORDER BY city,school_name").fetchall()
    cities=[r["city"] for r in con.execute("SELECT DISTINCT city FROM schools WHERE city<>'' ORDER BY city")]
    con.close()
    return templates.TemplateResponse("violations.html",{
      "request":request,"user":user,"rows":rows,"schools":schools,"cities":cities,
      "guard_emp_id":guard_emp_id,"city":city,"date_from":date_from,"date_to":date_to
    })

@app.post("/violations/save")
def save_violation(
    request: Request,
    violation_date: str=Form(...),
    ministry_id: str=Form(...),
    violation_type: str=Form(...),
    severity: str=Form("متوسطة"),
    action_taken: str=Form(""),
    notes: str=Form("")
):
    user=current_user(request)
    if not user: raise HTTPException(401)
    con=db()
    school=con.execute("SELECT * FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
    if not school: raise HTTPException(404)
    if user["role"]=="coordinator" and school["coordinator_name"]!=user["coordinator_name"]:
        raise HTTPException(403)
    con.execute("""
      INSERT INTO violations(
        violation_date,guard_emp_id,guard_name,ministry_id,school_name,city,
        coordinator_name,violation_type,severity,action_taken,notes,recorded_by,recorded_at
      ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
    """,(
      violation_date,school["guard_emp_id"],school["guard_name"],school["ministry_id"],
      school["school_name"],school["city"],school["coordinator_name"],violation_type,
      severity,action_taken,notes,user["id"],datetime.utcnow().isoformat()
    ))
    con.commit();con.close()
    return RedirectResponse("/violations",status_code=303)

@app.get("/supervisor-dashboard", response_class=HTMLResponse)
def supervisor_dashboard(request: Request, date_from: str="", date_to: str=""):
    user=require_permission(request,"supervisor_dashboard")
    today=date.today()
    date_from=date_from or today.replace(day=1).isoformat()
    date_to=date_to or today.isoformat()
    con=db()
    clause,args=scope_clause(user)
    rows=con.execute(f"""
      WITH RECURSIVE dates(d) AS (
        SELECT date(?)
        UNION ALL SELECT date(d,'+1 day') FROM dates WHERE d < date(?)
      ),
      base AS (
        SELECT d.d work_date,s.city,s.coordinator_name,s.ministry_id
        FROM dates d CROSS JOIN schools s
        WHERE 1=1 {clause}
      )
      SELECT city,coordinator_name,COUNT(*) total_slots,
        SUM(CASE WHEN COALESCE(a.status,'لم يسجل')='حاضر' THEN 1 ELSE 0 END) present_count,
        SUM(CASE WHEN COALESCE(a.status,'لم يسجل')='غائب' THEN 1 ELSE 0 END) absent_count,
        SUM(CASE WHEN COALESCE(a.status,'لم يسجل')='إجازة' THEN 1 ELSE 0 END) leave_count,
        SUM(CASE WHEN COALESCE(a.status,'لم يسجل')='بديل' THEN 1 ELSE 0 END) replacement_count,
        SUM(CASE WHEN a.status IS NULL OR a.status='لم يسجل' THEN 1 ELSE 0 END) missing_count
      FROM base b LEFT JOIN attendance a
        ON a.ministry_id=b.ministry_id AND a.work_date=b.work_date
      GROUP BY city,coordinator_name
      ORDER BY city,coordinator_name
    """,[date_from,date_to]+args).fetchall()
    con.close()
    return templates.TemplateResponse("supervisor_dashboard.html",{
      "request":request,"user":user,"rows":rows,"date_from":date_from,"date_to":date_to
    })


def _cc_percent(part,total):
    return round((float(part or 0)/float(total or 1))*100,1) if total else 0.0


def _command_center_payload(user, work_date: str):
    con=db()
    scope,args=scope_clause(user)
    yesterday=(date.fromisoformat(work_date)-timedelta(days=1)).isoformat()

    schools=con.execute(f"""
      SELECT s.ministry_id,s.school_name,s.city,s.district,s.stage,s.gender,
             s.operational_type,s.school_type,s.status,s.guard_name,s.guard_emp_id,
             s.guard_mobile,s.coordinator_name,s.coordinator_mobile,s.latitude,s.longitude,
             COALESCE(a.status,'لم يسجل') attendance_status,a.recorded_at attendance_time,
             (SELECT COUNT(*) FROM incidents i WHERE i.ministry_id=s.ministry_id AND i.status<>'مغلق') open_incidents,
             (SELECT COUNT(*) FROM school_access_permits p WHERE p.ministry_id=s.ministry_id AND p.actual_status='داخل المدرسة') active_permits,
             (SELECT COUNT(*) FROM violations v WHERE v.ministry_id=s.ministry_id AND v.violation_date=?) violations_today,
             (SELECT MAX(gt.recorded_at) FROM guard_tours gt WHERE gt.ministry_id=s.ministry_id) last_tour_at,
             (SELECT MAX(gt.tour_date) FROM guard_tours gt WHERE gt.ministry_id=s.ministry_id) last_tour_date
      FROM schools s
      LEFT JOIN attendance a ON a.ministry_id=s.ministry_id AND a.work_date=?
      WHERE 1=1 {scope}
      ORDER BY s.city,s.school_name
    """,[work_date,work_date]+args).fetchall()
    school_rows=[dict(r) for r in schools]
    total=len(school_rows)
    attendance_counts={"حاضر":0,"غائب":0,"إجازة":0,"بديل":0,"لم يسجل":0}
    for r in school_rows:
        st=r.get("attendance_status") or "لم يسجل"
        attendance_counts[st if st in attendance_counts else "لم يسجل"]+=1

    # Yesterday comparison for the same scope.
    yrow=con.execute(f"""
      SELECT COUNT(*) total,
       SUM(CASE WHEN a.status='حاضر' THEN 1 ELSE 0 END) present,
       SUM(CASE WHEN a.status='بديل' THEN 1 ELSE 0 END) replacement
      FROM schools s LEFT JOIN attendance a ON a.ministry_id=s.ministry_id AND a.work_date=?
      WHERE 1=1 {scope}
    """,[yesterday]+args).fetchone()
    yesterday_coverage=_cc_percent((yrow["present"] or 0)+(yrow["replacement"] or 0),yrow["total"] or 0)

    active_guards=con.execute("""SELECT COUNT(*) c FROM guard_master
      WHERE COALESCE(service_status,'على رأس العمل')<>'منتهي الخدمة'""").fetchone()["c"]
    unlinked_guards=con.execute("""SELECT COUNT(*) c FROM guard_master
      WHERE COALESCE(service_status,'على رأس العمل')<>'منتهي الخدمة'
        AND COALESCE(TRIM(linked_ministry_id),'')=''""").fetchone()["c"]
    active_coords=con.execute("SELECT COUNT(*) c FROM coordinator_master WHERE is_active=1").fetchone()["c"]
    active_users=con.execute("SELECT COUNT(*) c FROM users WHERE is_active=1").fetchone()["c"]

    op=con.execute("""SELECT
      (SELECT COUNT(*) FROM incidents WHERE status<>'مغلق') open_incidents,
      (SELECT COUNT(*) FROM incidents WHERE status<>'مغلق' AND COALESCE(severity,'') IN ('حرج','عالية','مرتفع')) critical_incidents,
      (SELECT COUNT(*) FROM guard_tours WHERE tour_date=?) tours_today,
      (SELECT COUNT(*) FROM violations WHERE violation_date=?) violations_today,
      (SELECT COUNT(*) FROM school_access_permits WHERE actual_status='داخل المدرسة') teams_inside,
      (SELECT COUNT(*) FROM school_access_permits WHERE approval_status='بانتظار الموافقة') pending_permits,
      (SELECT COUNT(*) FROM school_access_permits WHERE actual_status='داخل المدرسة' AND permit_date<?) overdue_exits,
      (SELECT COALESCE(SUM(shortage_count),0) FROM shifts WHERE work_date=?) shortages,
      (SELECT COUNT(*) FROM employee_documents WHERE expiry_date<>'' AND expiry_date IS NOT NULL AND expiry_date>=? AND expiry_date<=?) expiring_docs
    """,(work_date,work_date,work_date,work_date,work_date,(date.fromisoformat(work_date)+timedelta(days=30)).isoformat())).fetchone()
    operations=dict(op)

    # Required documents missing (guard-level requirements).
    try:
        missing_docs=con.execute("""SELECT COUNT(*) c FROM guard_document_requirements r
          LEFT JOIN employee_documents d ON d.guard_emp_id=r.guard_emp_id AND d.document_type=r.document_type
          WHERE r.requirement_status='مطلوب' AND d.id IS NULL""").fetchone()["c"]
    except Exception:
        missing_docs=0

    no_guard=sum(1 for r in school_rows if not (r.get("guard_emp_id") or "").strip())
    no_coord=sum(1 for r in school_rows if not (r.get("coordinator_name") or "").strip() or (r.get("coordinator_name") or "").strip()=="غير متوفر")
    no_coords=sum(1 for r in school_rows if r.get("latitude") is None or r.get("longitude") is None)

    # Coordinator performance.
    coord_base={}
    for s in school_rows:
        key=(s.get("city") or "غير محدد",s.get("coordinator_name") or "غير متوفر")
        c=coord_base.setdefault(key,{"city":key[0],"name":key[1],"schools":0,"present":0,"absent":0,"unrecorded":0,"replacement":0,"phone":s.get("coordinator_mobile") or ""})
        c["schools"]+=1
        st=s.get("attendance_status") or "لم يسجل"
        if st=="حاضر": c["present"]+=1
        elif st=="بديل": c["replacement"]+=1
        elif st=="غائب": c["absent"]+=1
        elif st=="لم يسجل": c["unrecorded"]+=1
    tour_map={(r["city"] or "غير محدد",r["coordinator_name"] or "غير متوفر"):r["c"] for r in con.execute("""SELECT city,coordinator_name,COUNT(*) c FROM guard_tours WHERE tour_date=? GROUP BY city,coordinator_name""",(work_date,)).fetchall()}
    incident_map={(r["city"] or "غير محدد",r["coordinator_name"] or "غير متوفر"):r["c"] for r in con.execute("""SELECT s.city,s.coordinator_name,COUNT(i.id) c FROM schools s LEFT JOIN incidents i ON i.ministry_id=s.ministry_id AND i.status<>'مغلق' GROUP BY s.city,s.coordinator_name""").fetchall()}
    violation_map={(r["city"] or "غير محدد",r["coordinator_name"] or "غير متوفر"):r["c"] for r in con.execute("""SELECT city,coordinator_name,COUNT(*) c FROM violations WHERE violation_date=? GROUP BY city,coordinator_name""",(work_date,)).fetchall()}
    coordinators=[]
    for key,c in coord_base.items():
        c["coverage"]=_cc_percent(c["present"]+c["replacement"],c["schools"])
        c["recording"]=_cc_percent(c["schools"]-c["unrecorded"],c["schools"])
        c["tours"]=tour_map.get(key,0);c["incidents"]=incident_map.get(key,0);c["violations"]=violation_map.get(key,0)
        c["score"]=max(0,round(c["coverage"]*.55+min(c["tours"]/(c["schools"] or 1),1)*25+max(0,20-c["incidents"]*3-c["violations"]*2),1))
        coordinators.append(c)
    coordinators.sort(key=lambda x:(x["score"],-x["schools"]))

    coverage=_cc_percent(attendance_counts["حاضر"]+attendance_counts["بديل"],total)
    recording=_cc_percent(total-attendance_counts["لم يسجل"],total)
    tour_target=max(active_coords,1)
    tour_rate=min(100,_cc_percent(operations["tours_today"],tour_target))
    data_quality=max(0,100-_cc_percent(no_guard+no_coord+no_coords,total*3 if total else 1))
    docs_health=max(0,100-min(100,missing_docs/max(active_guards,1)*100))
    incident_penalty=min(100,operations["open_incidents"]*2+operations["critical_incidents"]*8)
    health=max(0,min(100,round(coverage*.35+recording*.1+tour_rate*.15+data_quality*.15+docs_health*.15+(100-incident_penalty)*.1,1)))

    # Alerts generated from current operational state.
    alerts=[]
    def add_alert(key,severity,kind,title,detail,url):
        alerts.append({"key":key,"severity":severity,"kind":kind,"title":title,"detail":detail,"url":url})
    if attendance_counts["غائب"]: add_alert(f"absent:{work_date}","حرج","حضور",f"{attendance_counts['غائب']} مدارس بها غياب","تحتاج متابعة فورية مع المنسقين / المشرفين","/dashboard?work_date="+work_date+"&status_filter=غائب")
    if attendance_counts["لم يسجل"]: add_alert(f"unrecorded:{work_date}","مرتفع","حضور",f"{attendance_counts['لم يسجل']} مدارس لم تسجل","قائمة المدارس غير المسجلة حتى الآن","/dashboard?work_date="+work_date+"&status_filter=لم يسجل")
    if operations["critical_incidents"]: add_alert("critical-incidents","حرج","بلاغات",f"{operations['critical_incidents']} بلاغات حرجة","بلاغات مفتوحة ذات أولوية عالية","/incidents")
    if operations["overdue_exits"]: add_alert("overdue-exits","مرتفع","تصاريح",f"{operations['overdue_exits']} تصاريح دون مغادرة","فرق ما زالت مسجلة داخل المدارس","/access-permits")
    if no_guard: add_alert("no-guard","مرتفع","جودة البيانات",f"{no_guard} مدارس بدون بواب","يلزم الربط أو إضافة بواب بديل","/unlinked-guards")
    if no_coord: add_alert("no-coordinator","متوسط","جودة البيانات",f"{no_coord} مدارس بدون منسق / مشرف","يلزم إعادة توزيع المدارس","/coordinators")
    if no_coords: add_alert("no-coordinates","متوسط","جودة البيانات",f"{no_coords} مدارس بدون إحداثيات","لن تظهر هذه المدارس على الخريطة","/coordinates")
    if missing_docs: add_alert("missing-documents","متوسط","المستندات",f"{missing_docs} مستندات مطلوبة ناقصة","متابعة ملفات البوابين غير المكتملة","/documents")
    if operations["pending_permits"]: add_alert("pending-permits","متوسط","تصاريح",f"{operations['pending_permits']} تصاريح بانتظار الاعتماد","طلبات تشغيل وصيانة جديدة","/access-permits")

    handled={r["alert_key"]:dict(r) for r in con.execute("SELECT * FROM command_center_alert_actions WHERE status='تمت المعالجة'").fetchall()}
    for a in alerts:
        a["handled"]=a["key"] in handled
    alerts.sort(key=lambda a:({"حرج":0,"مرتفع":1,"متوسط":2,"منخفض":3}.get(a["severity"],4),a["handled"]))

    # Live event feed from real records.
    events=[]
    for r in con.execute("""SELECT a.recorded_at ts,'حضور' event_type,s.school_name title,a.status detail,s.city FROM attendance a LEFT JOIN schools s ON s.ministry_id=a.ministry_id WHERE a.work_date=? ORDER BY a.recorded_at DESC LIMIT 12""",(work_date,)).fetchall(): events.append(dict(r))
    for r in con.execute("""SELECT recorded_at ts,'بلاغ' event_type,school_name title,incident_type||' — '||status detail,city FROM incidents ORDER BY recorded_at DESC LIMIT 10""").fetchall(): events.append(dict(r))
    for r in con.execute("""SELECT recorded_at ts,'جولة' event_type,school_name title,status detail,city FROM guard_tours ORDER BY recorded_at DESC LIMIT 10""").fetchall(): events.append(dict(r))
    for r in con.execute("""SELECT created_at ts,'تصريح' event_type,school_name title,visit_type||' — '||approval_status detail,city FROM school_access_permits ORDER BY created_at DESC LIMIT 10""").fetchall(): events.append(dict(r))
    for r in con.execute("""SELECT recorded_at ts,'مخالفة' event_type,school_name title,violation_type detail,city FROM violations ORDER BY recorded_at DESC LIMIT 10""").fetchall(): events.append(dict(r))
    events.sort(key=lambda e:e.get("ts") or "",reverse=True);events=events[:30]

    # Map risk state.
    map_points=[]
    for r in school_rows:
        if r.get("latitude") is None or r.get("longitude") is None: continue
        st=r.get("attendance_status") or "لم يسجل"
        risk=0
        if st=="غائب": risk+=55
        elif st=="لم يسجل": risk+=35
        elif st=="إجازة": risk+=12
        risk+=min(25,(r.get("open_incidents") or 0)*12)+min(12,(r.get("violations_today") or 0)*5)+min(8,(r.get("active_permits") or 0)*2)
        if not (r.get("guard_emp_id") or "").strip():risk+=25
        status="حرج" if risk>=60 else "تنبيه" if risk>=30 else "مستقر"
        point=dict(r);point["risk_score"]=min(100,risk);point["project_status"]=status
        map_points.append(point)

    con.close()
    return {
      "work_date":work_date,"generated_at":datetime.now().isoformat(),"health_score":health,
      "coverage":coverage,"yesterday_coverage":yesterday_coverage,"recording":recording,
      "attendance":attendance_counts,"totals":{"schools":total,"guards":active_guards,"coordinators":active_coords,"users":active_users},
      "operations":operations,"quality":{"no_guard":no_guard,"no_coord":no_coord,"no_coords":no_coords,"unlinked_guards":unlinked_guards,"missing_docs":missing_docs},
      "rates":{"tour_rate":tour_rate,"data_quality":round(data_quality,1),"docs_health":round(docs_health,1)},
      "alerts":alerts,"events":events,"coordinators":coordinators,"map_points":map_points
    }


@app.get("/security-dashboard", response_class=HTMLResponse)
def security_dashboard(request: Request, work_date: str=""):
    user=require_permission(request,"security_dashboard")
    work_date=work_date or date.today().isoformat()
    payload=_command_center_payload(user,work_date)
    return templates.TemplateResponse("security_dashboard.html",{
      "request":request,"user":user,"work_date":work_date,"payload":payload,
      "settings":get_app_settings()
    })


@app.get("/api/command-center")
def command_center_api(request: Request,work_date: str=""):
    user=require_permission(request,"security_dashboard")
    work_date=work_date or date.today().isoformat()
    return JSONResponse(_command_center_payload(user,work_date))


@app.post("/command-center/alerts/action")
def command_center_alert_action(request: Request,alert_key: str=Form(...),status: str=Form("تمت المعالجة"),assigned_to: str=Form(""),notes: str=Form("")):
    user=require_permission(request,"security_dashboard");con=db();now=datetime.now().isoformat()
    con.execute("""INSERT INTO command_center_alert_actions(alert_key,status,assigned_to,notes,updated_by,updated_at)
      VALUES(?,?,?,?,?,?) ON CONFLICT(alert_key) DO UPDATE SET status=excluded.status,assigned_to=excluded.assigned_to,
      notes=excluded.notes,updated_by=excluded.updated_by,updated_at=excluded.updated_at""",
      (alert_key,status,assigned_to,notes,user["id"],now))
    con.commit();con.close();return JSONResponse({"ok":True})

@app.get("/guard-tours", response_class=HTMLResponse)
def guard_tours_page(request: Request, date_from: str="", date_to: str="", city: str=""):
    user=require_permission(request,"guard_tours")
    con=db();where=["1=1"];args=[]
    if user["role"]=="coordinator":
        where.append("coordinator_name=?");args.append(user["coordinator_name"])
    if date_from: where.append("tour_date>=?");args.append(date_from)
    if date_to: where.append("tour_date<=?");args.append(date_to)
    if city: where.append("city=?");args.append(city)
    rows=con.execute("SELECT * FROM guard_tours WHERE "+" AND ".join(where)+" ORDER BY tour_date DESC,visit_time DESC",args).fetchall()
    schools=con.execute("SELECT ministry_id,school_name,city,coordinator_name FROM schools ORDER BY city,school_name").fetchall()
    cities=[r["city"] for r in con.execute("SELECT DISTINCT city FROM schools WHERE city<>'' ORDER BY city")]
    con.close()
    return templates.TemplateResponse("guard_tours.html",{"request":request,"user":user,"rows":rows,"schools":schools,"cities":cities,"date_from":date_from,"date_to":date_to,"city":city})

@app.post("/guard-tours/save")
def guard_tour_save(
 request: Request, tour_date: str=Form(...), ministry_id: str=Form(...),
 checkpoint_code: str=Form(""), visit_time: str=Form(""), latitude: str=Form(""),
 longitude: str=Form(""), status: str=Form("تمت"), notes: str=Form("")
):
    user=current_user(request)
    if not user: raise HTTPException(401)
    con=db();s=con.execute("SELECT * FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
    if not s: raise HTTPException(404)
    if user["role"]=="coordinator" and s["coordinator_name"]!=user["coordinator_name"]: raise HTTPException(403)
    con.execute("""INSERT INTO guard_tours(tour_date,ministry_id,school_name,city,coordinator_name,checkpoint_code,visit_time,latitude,longitude,status,notes,recorded_by,recorded_at)
    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",(tour_date,ministry_id,s["school_name"],s["city"],s["coordinator_name"],checkpoint_code,visit_time,float(latitude or 0),float(longitude or 0),status,notes,user["id"],datetime.utcnow().isoformat()))
    con.commit();con.close()
    return RedirectResponse("/guard-tours",status_code=303)

@app.get("/incidents", response_class=HTMLResponse)
def incidents_page(request: Request, status: str="", city: str=""):
    user=require_permission(request,"incidents")
    con=db();where=["1=1"];args=[]
    if status: where.append("status=?");args.append(status)
    if city: where.append("city=?");args.append(city)
    rows=con.execute("SELECT * FROM incidents WHERE "+" AND ".join(where)+" ORDER BY incident_date DESC,incident_time DESC",args).fetchall()
    schools=con.execute("SELECT ministry_id,school_name,city FROM schools ORDER BY city,school_name").fetchall()
    cities=[r["city"] for r in con.execute("SELECT DISTINCT city FROM schools WHERE city<>'' ORDER BY city")]
    con.close()
    return templates.TemplateResponse("incidents.html",{"request":request,"user":user,"rows":rows,"schools":schools,"cities":cities,"status":status,"city":city})

@app.post("/incidents/save")
def incident_save(
 request: Request, incident_date: str=Form(...), incident_time: str=Form(""),
 ministry_id: str=Form(""), incident_type: str=Form(...), severity: str=Form("متوسطة"),
 description: str=Form(""), action_taken: str=Form(""), status: str=Form("مفتوح"),
 assigned_to: str=Form("")
):
    user=current_user(request)
    if not user: raise HTTPException(401)
    con=db();s=con.execute("SELECT * FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
    school_name=s["school_name"] if s else ""; city=s["city"] if s else ""
    con.execute("""INSERT INTO incidents(incident_date,incident_time,ministry_id,school_name,city,incident_type,severity,description,action_taken,status,assigned_to,recorded_by,recorded_at)
    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",(incident_date,incident_time,ministry_id,school_name,city,incident_type,severity,description,action_taken,status,assigned_to,user["id"],datetime.utcnow().isoformat()))
    con.commit();con.close()
    return RedirectResponse("/incidents",status_code=303)

@app.post("/incidents/{incident_id}/close")
def incident_close(request: Request, incident_id: int):
    user=current_user(request)
    if not user: raise HTTPException(401)
    con=db();con.execute("UPDATE incidents SET status='مغلق',closed_at=? WHERE id=?",(datetime.utcnow().isoformat(),incident_id));con.commit();con.close()
    return RedirectResponse("/incidents",status_code=303)

@app.get("/assets", response_class=HTMLResponse)
def assets_page(request: Request):
    user=require_permission(request,"assets")
    con=db();rows=con.execute("SELECT * FROM assets ORDER BY id DESC").fetchall()
    schools=con.execute("SELECT ministry_id,school_name,city,guard_name,guard_emp_id FROM schools ORDER BY city,school_name").fetchall()
    con.close()
    return templates.TemplateResponse("assets.html",{"request":request,"user":user,"rows":rows,"schools":schools})

@app.post("/assets/save")
def asset_save(
 request: Request, asset_type: str=Form(...), asset_code: str=Form(""), asset_name: str=Form(""),
 serial_number: str=Form(""), ministry_id: str=Form(""), assigned_guard_emp_id: str=Form(""),
 quantity: int=Form(1), condition_status: str=Form("جيدة"), received_date: str=Form(""), notes: str=Form("")
):
    user=current_user(request)
    if not user: raise HTTPException(401)
    con=db();s=con.execute("SELECT * FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
    con.execute("""INSERT INTO assets(asset_type,asset_code,asset_name,serial_number,ministry_id,school_name,assigned_guard_emp_id,assigned_guard_name,quantity,condition_status,received_date,notes,recorded_by,recorded_at)
    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(asset_type,asset_code,asset_name,serial_number,ministry_id,s["school_name"] if s else "",assigned_guard_emp_id,s["guard_name"] if s else "",quantity,condition_status,received_date,notes,user["id"],datetime.utcnow().isoformat()))
    con.commit();con.close()
    return RedirectResponse("/assets",status_code=303)

@app.get("/shifts", response_class=HTMLResponse)
def shifts_page(request: Request, work_date: str=""):
    user=require_permission(request,"shifts")
    work_date=work_date or date.today().isoformat()
    con=db();rows=con.execute("SELECT * FROM shifts WHERE work_date=? ORDER BY city,school_name,shift_name",(work_date,)).fetchall()
    schools=con.execute("SELECT ministry_id,school_name,city,coordinator_name FROM schools ORDER BY city,school_name").fetchall()
    total_shortage=sum(r["shortage_count"] for r in rows)
    con.close()
    return templates.TemplateResponse("shifts.html",{"request":request,"user":user,"rows":rows,"schools":schools,"work_date":work_date,"total_shortage":total_shortage})

@app.post("/shifts/save")
def shift_save(
 request: Request, work_date: str=Form(...), ministry_id: str=Form(...),
 shift_name: str=Form(...), required_count: int=Form(1), assigned_count: int=Form(0), notes: str=Form("")
):
    user=current_user(request)
    if not user: raise HTTPException(401)
    con=db();s=con.execute("SELECT * FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
    shortage=max(required_count-assigned_count,0)
    con.execute("""INSERT INTO shifts(work_date,ministry_id,school_name,city,shift_name,required_count,assigned_count,shortage_count,supervisor_name,notes,recorded_by,recorded_at)
    VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",(work_date,ministry_id,s["school_name"],s["city"],shift_name,required_count,assigned_count,shortage,s["coordinator_name"],notes,user["id"],datetime.utcnow().isoformat()))
    con.commit();con.close()
    return RedirectResponse(f"/shifts?work_date={work_date}",status_code=303)

@app.get("/evaluations", response_class=HTMLResponse)
def evaluations_page(request: Request):
    user=require_permission(request,"evaluations")
    con=db();rows=con.execute("SELECT * FROM guard_evaluations ORDER BY evaluation_date DESC,id DESC").fetchall()
    schools=con.execute("SELECT ministry_id,school_name,city,guard_name,guard_emp_id,coordinator_name FROM schools ORDER BY city,school_name").fetchall()
    con.close()
    return templates.TemplateResponse("evaluations.html",{"request":request,"user":user,"rows":rows,"schools":schools})

@app.post("/evaluations/save")
def evaluation_save(
 request: Request, evaluation_date: str=Form(...), ministry_id: str=Form(...),
 discipline: int=Form(...), attendance_score: int=Form(...), appearance: int=Form(...),
 performance: int=Form(...), customer_service: int=Form(...), notes: str=Form("")
):
    user=current_user(request)
    if not user: raise HTTPException(401)
    con=db();s=con.execute("SELECT * FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
    total=(discipline+attendance_score+appearance+performance+customer_service)/5
    con.execute("""INSERT INTO guard_evaluations(evaluation_date,guard_emp_id,guard_name,ministry_id,school_name,city,coordinator_name,discipline,attendance_score,appearance,performance,customer_service,total_score,notes,recorded_by,recorded_at)
    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(evaluation_date,s["guard_emp_id"],s["guard_name"],ministry_id,s["school_name"],s["city"],s["coordinator_name"],discipline,attendance_score,appearance,performance,customer_service,total,notes,user["id"],datetime.utcnow().isoformat()))
    con.commit();con.close()
    return RedirectResponse("/evaluations",status_code=303)

@app.get("/documents", response_class=HTMLResponse)
def documents_page(request: Request, q: str="", document_filter: str="all"):
    user=require_permission(request,"documents")
    con=db()
    guards=con.execute("""SELECT DISTINCT guard_emp_id,guard_name,guard_id,guard_mobile
      FROM schools WHERE COALESCE(TRIM(guard_emp_id),'')<>'' ORDER BY guard_name""").fetchall()
    doc_types=["عقد العمل","الهوية / الإقامة","عقد قوى","خلو السوابق","فحص السموم","شهادة الآيبان","العنوان الوطني","الإقرار"]
    result=[]
    for g in guards:
        if q and q not in (g["guard_name"] or "") and q not in (g["guard_emp_id"] or ""): continue
        docs={r["document_type"]:dict(r) for r in con.execute("SELECT * FROM employee_documents WHERE guard_emp_id=? ORDER BY id DESC",(g["guard_emp_id"],)).fetchall()}
        reqs={r["document_type"]:dict(r) for r in con.execute("SELECT * FROM guard_document_requirements WHERE guard_emp_id=?",(g["guard_emp_id"],)).fetchall()}
        items=[];required=uploaded=missing=0
        for dt in doc_types:
            req=reqs.get(dt,{}).get("requirement_status","مطلوب")
            d=docs.get(dt)
            state="غير مطلوب" if req=="غير مطلوب" else ("مرفوع" if d and d.get("file_path") else "ناقص")
            if req!="غير مطلوب":
                required+=1
                if state=="مرفوع": uploaded+=1
                else: missing+=1
            items.append({"type":dt,"requirement":req,"state":state,"doc":d})
        completion=round(uploaded*100/required) if required else 100
        if document_filter=="complete" and missing: continue
        if document_filter=="missing" and not missing: continue
        result.append({"guard":dict(g),"items":items,"required":required,"uploaded":uploaded,"missing":missing,"completion":completion})
    con.close()
    return templates.TemplateResponse("documents.html",{"request":request,"user":user,"guards":result,"q":q,"document_filter":document_filter,"doc_types":doc_types})

@app.post("/documents/requirement")
def document_requirement(request: Request, guard_emp_id: str=Form(...), document_type: str=Form(...), requirement_status: str=Form(...), reason: str=Form("")):
    user=current_user(request)
    if not user or user["role"] not in ("admin","manager","hr","supervisor"): raise HTTPException(403)
    if requirement_status not in ("مطلوب","غير مطلوب"): raise HTTPException(400)
    con=db();con.execute("""INSERT INTO guard_document_requirements(guard_emp_id,document_type,requirement_status,reason,updated_by,updated_at)
      VALUES(?,?,?,?,?,?) ON CONFLICT(guard_emp_id,document_type) DO UPDATE SET requirement_status=excluded.requirement_status,reason=excluded.reason,updated_by=excluded.updated_by,updated_at=excluded.updated_at""",
      (guard_emp_id,document_type,requirement_status,reason,user["id"],datetime.utcnow().isoformat()))
    con.commit();con.close();return RedirectResponse("/documents",303)

@app.post("/documents/upload")
async def document_upload(request: Request, guard_emp_id: str=Form(...), guard_name: str=Form(""), document_type: str=Form(...), document_number: str=Form(""), issue_date: str=Form(""), expiry_date: str=Form(""), notes: str=Form(""), file: UploadFile=File(...)):
    user=current_user(request)
    if not user: raise HTTPException(401)
    allowed={".pdf",".png",".jpg",".jpeg"};suffix=Path(file.filename or "").suffix.lower()
    if suffix not in allowed: raise HTTPException(400,detail="يسمح بملفات PDF والصور فقط")
    folder=BASE/"data"/"guard_documents"/re.sub(r'[^0-9A-Za-z_-]','_',guard_emp_id);folder.mkdir(parents=True,exist_ok=True)
    safe=f"{re.sub(r'[^0-9A-Za-z_-]','_',document_type)}_{datetime.utcnow().strftime('%Y%m%d%H%M%S')}{suffix}"
    target=folder/safe;target.write_bytes(await file.read())
    rel=str(target.relative_to(BASE)).replace('\\','/')
    con=db();con.execute("DELETE FROM employee_documents WHERE guard_emp_id=? AND document_type=?",(guard_emp_id,document_type))
    con.execute("""INSERT INTO employee_documents(guard_emp_id,guard_name,document_type,document_number,issue_date,expiry_date,status,notes,recorded_by,recorded_at,file_path,original_file_name,uploaded_at,requirement_status)
      VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(guard_emp_id,guard_name,document_type,document_number,issue_date,expiry_date,"مرفوع",notes,user["id"],datetime.utcnow().isoformat(),rel,file.filename,datetime.utcnow().isoformat(),"مطلوب"))
    con.commit();con.close();return RedirectResponse("/documents",303)

@app.get("/documents/file/{doc_id}")
def document_file(request: Request, doc_id: int):
    user=current_user(request)
    if not user: raise HTTPException(401)
    con=db();row=con.execute("SELECT * FROM employee_documents WHERE id=?",(doc_id,)).fetchone();con.close()
    if not row or not row["file_path"]: raise HTTPException(404)
    path=BASE/row["file_path"]
    if not path.exists(): raise HTTPException(404)
    return FileResponse(path,filename=row["original_file_name"] or path.name)

@app.post("/documents/delete/{doc_id}")
def document_delete(request: Request, doc_id: int):
    user=current_user(request)
    if not user or user["role"] not in ("admin","manager","hr","supervisor"): raise HTTPException(403)
    con=db();row=con.execute("SELECT file_path FROM employee_documents WHERE id=?",(doc_id,)).fetchone()
    if row and row["file_path"]:
        path=BASE/row["file_path"]
        if path.exists(): path.unlink()
    con.execute("DELETE FROM employee_documents WHERE id=?",(doc_id,));con.commit();con.close()
    return RedirectResponse("/documents",303)

@app.get("/client-dashboard", response_class=HTMLResponse)
def client_dashboard(
    request: Request,
    date_from: str="",
    date_to: str="",
    city: str=""
):
    user=require_permission(request,"client_dashboard")
    today=date.today()
    date_to=date_to or today.isoformat()
    date_from=date_from or (today.replace(day=1)).isoformat()
    con=db()

    scope,args=scope_clause(user)
    extra="";extra_args=[]
    if city and not (user["role"]=="client" and "client_city" in user.keys() and user["client_city"]):
        extra+=" AND s.city=?";extra_args.append(city)

    base_where=f" WHERE 1=1 {scope} {extra} "
    school_args=args+extra_args

    school_stats=con.execute(f"""SELECT
      COUNT(*) total_schools,
      SUM(CASE WHEN COALESCE(TRIM(s.guard_emp_id),'')<>'' THEN 1 ELSE 0 END) covered_schools,
      SUM(CASE WHEN COALESCE(TRIM(s.guard_emp_id),'')='' THEN 1 ELSE 0 END) uncovered_schools,
      COUNT(DISTINCT s.city) cities,
      COUNT(DISTINCT NULLIF(TRIM(s.coordinator_name),'')) coordinators
      FROM schools s {base_where}""",school_args).fetchone()

    attendance=con.execute(f"""SELECT
      COUNT(DISTINCT s.ministry_id) expected,
      SUM(CASE WHEN a.status='حاضر' THEN 1 ELSE 0 END) present,
      SUM(CASE WHEN a.status='غائب' THEN 1 ELSE 0 END) absent,
      SUM(CASE WHEN a.status='بديل' THEN 1 ELSE 0 END) replacement,
      SUM(CASE WHEN a.status='إجازة' THEN 1 ELSE 0 END) leave_count
      FROM schools s
      LEFT JOIN attendance a ON a.ministry_id=s.ministry_id AND a.work_date=?
      {base_where}""",[date_to]+school_args).fetchone()

    violations_count=con.execute(f"""SELECT COUNT(*) c FROM violations v
      JOIN schools s ON s.ministry_id=v.ministry_id
      {base_where} AND v.violation_date BETWEEN ? AND ?""",school_args+[date_from,date_to]).fetchone()["c"]

    tours=con.execute(f"""SELECT
      COUNT(*) total,
      SUM(CASE WHEN gt.status='تمت' THEN 1 ELSE 0 END) completed,
      SUM(CASE WHEN gt.status<>'تمت' THEN 1 ELSE 0 END) pending
      FROM guard_tours gt JOIN schools s ON s.ministry_id=gt.ministry_id
      {base_where} AND gt.tour_date BETWEEN ? AND ?""",school_args+[date_from,date_to]).fetchone()

    open_incidents=con.execute(f"""SELECT COUNT(*) c FROM incidents i
      JOIN schools s ON s.ministry_id=i.ministry_id
      {base_where} AND COALESCE(i.status,'مفتوح')<>'مغلق'""",school_args).fetchone()["c"]

    city_distribution=con.execute(f"""SELECT s.city,COUNT(*) total,
      SUM(CASE WHEN COALESCE(TRIM(s.guard_emp_id),'')<>'' THEN 1 ELSE 0 END) covered
      FROM schools s {base_where}
      GROUP BY s.city ORDER BY total DESC""",school_args).fetchall()

    alerts=con.execute(f"""SELECT * FROM (
      SELECT v.violation_date event_date,v.school_name title,
        'مخالفة: '||COALESCE(v.violation_type,'غير محدد') message,'violation' event_type,
        v.city
      FROM violations v JOIN schools s ON s.ministry_id=v.ministry_id
      {base_where} AND v.violation_date BETWEEN ? AND ?
      UNION ALL
      SELECT i.incident_date event_date,i.school_name title,
        'بلاغ: '||COALESCE(i.incident_type,'غير محدد') message,'incident' event_type,
        i.city
      FROM incidents i JOIN schools s ON s.ministry_id=i.ministry_id
      {base_where} AND COALESCE(i.status,'مفتوح')<>'مغلق'
      UNION ALL
      SELECT gt.tour_date event_date,gt.school_name title,
        'جولة: '||COALESCE(gt.status,'—') message,'tour' event_type,
        gt.city
      FROM guard_tours gt JOIN schools s ON s.ministry_id=gt.ministry_id
      {base_where} AND gt.tour_date BETWEEN ? AND ?
      ) ORDER BY event_date DESC LIMIT 8""",
      school_args+[date_from,date_to]+school_args+school_args+[date_from,date_to]).fetchall()

    latest_visits=con.execute(f"""SELECT gt.*,s.school_name,s.city,
      COALESCE(gt.visit_time,'—') visit_time_display
      FROM guard_tours gt JOIN schools s ON s.ministry_id=gt.ministry_id
      {base_where} AND gt.tour_date BETWEEN ? AND ?
      ORDER BY gt.tour_date DESC,gt.visit_time DESC LIMIT 8""",school_args+[date_from,date_to]).fetchall()

    attendance_trend=con.execute(f"""SELECT a.work_date,
      SUM(CASE WHEN a.status='حاضر' THEN 1 ELSE 0 END) present,
      SUM(CASE WHEN a.status='غائب' THEN 1 ELSE 0 END) absent
      FROM attendance a JOIN schools s ON s.ministry_id=a.ministry_id
      {base_where} AND a.work_date BETWEEN ? AND ?
      GROUP BY a.work_date ORDER BY a.work_date""",school_args+[date_from,date_to]).fetchall()

    violation_types=con.execute(f"""SELECT COALESCE(v.violation_type,'أخرى') label,COUNT(*) total
      FROM violations v JOIN schools s ON s.ministry_id=v.ministry_id
      {base_where} AND v.violation_date BETWEEN ? AND ?
      GROUP BY COALESCE(v.violation_type,'أخرى') ORDER BY total DESC LIMIT 6""",school_args+[date_from,date_to]).fetchall()

    map_schools=con.execute(f"""SELECT ministry_id,school_name,city,latitude,longitude,
      guard_name,coordinator_name FROM schools s
      {base_where} AND latitude IS NOT NULL AND longitude IS NOT NULL
      ORDER BY city,school_name""",school_args).fetchall()

    all_cities=[r["city"] for r in con.execute("SELECT DISTINCT city FROM schools WHERE COALESCE(city,'')<>'' ORDER BY city")]
    con.close()

    coverage_pct=round((school_stats["covered_schools"] or 0)*100/(school_stats["total_schools"] or 1),1)
    attendance_pct=round((attendance["present"] or 0)*100/(attendance["expected"] or 1),1)
    tours_pct=round((tours["completed"] or 0)*100/(tours["total"] or 1),1)

    return templates.TemplateResponse("client_dashboard.html",{
      "request":request,"user":user,"date_from":date_from,"date_to":date_to,
      "selected_city":city,"cities":all_cities,"school_stats":school_stats,
      "attendance":attendance,"violations_count":violations_count,
      "tours":tours,"open_incidents":open_incidents,
      "city_distribution":city_distribution,"alerts":alerts,
      "latest_visits":latest_visits,"attendance_trend":[dict(r) for r in attendance_trend],
      "violation_types":violation_types,
      "map_schools":[dict(r) for r in map_schools],
      "coverage_pct":coverage_pct,"attendance_pct":attendance_pct,"tours_pct":tours_pct,
      "client_name":(user["client_name"] if "client_name" in user.keys() and user["client_name"] else user["full_name"])
    })



def get_app_settings():
    con=db()
    rows=con.execute("SELECT setting_key,setting_value FROM app_settings").fetchall()
    con.close()
    return {r["setting_key"]:r["setting_value"] for r in rows}

def normalize_whatsapp_number(number):
    digits="".join(ch for ch in str(number or "") if ch.isdigit())
    if digits.startswith("00"):
        digits=digits[2:]
    if digits.startswith("0") and len(digits)==10:
        digits="966"+digits[1:]
    elif not digits.startswith("966") and len(digits)==9:
        digits="966"+digits
    return digits

def ensure_permit_tokens(con, permit_id):
    row=con.execute("SELECT approval_token,exit_token FROM school_access_permits WHERE id=?",(permit_id,)).fetchone()
    if not row:
        return None,None
    approval=row["approval_token"] or uuid.uuid4().hex
    exit_token=row["exit_token"] or uuid.uuid4().hex
    con.execute("UPDATE school_access_permits SET approval_token=?,exit_token=? WHERE id=?",(approval,exit_token,permit_id))
    con.commit()
    return approval,exit_token

def save_access_evidence(upload, permit_id, event_type):
    if not upload or not getattr(upload,"filename",""):
        return "", ""
    ext=Path(upload.filename).suffix.lower()
    if ext not in (".jpg",".jpeg",".png",".webp"):
        raise HTTPException(400,detail="صيغة صورة الإثبات غير مدعومة")
    evidence_no=f"{event_type.upper()}-{permit_id}-{uuid.uuid4().hex[:10].upper()}"
    filename=f"{evidence_no}{ext}"
    path=EVIDENCE_DIR / filename
    content=upload.file.read()
    if len(content)>8*1024*1024:
        raise HTTPException(400,detail="حجم صورة الإثبات أكبر من 8 ميجابايت")
    path.write_bytes(content)
    digest=hashlib.sha256(content).hexdigest()
    return filename,digest

def _norm_excel_value(value):
    if value is None:
        return ""
    text=str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text=text[:-2]
    return text

def _pick(row, *aliases):
    normalized={str(k or "").strip():v for k,v in row.items()}
    for alias in aliases:
        if alias in normalized:
            return _norm_excel_value(normalized[alias])
    return ""

def _normalize_operational_type(value, fallback=""):
    """Normalize school scope without treating blank values as main schools."""
    text=(value or "").strip().replace("ـ", "")
    fallback_text=(fallback or "").strip().replace("ـ", "")
    combined=f"{text} {fallback_text}".strip()
    if "ملحق" in combined:
        return "ملحقة"
    if "فرع" in combined:
        return "فرعية"
    main_aliases={"رئيسية","رئيسي","أساسية","اساسية","أساسي","اساسي","main","primary"}
    if text.lower() in main_aliases or fallback_text.lower() in main_aliases:
        return "رئيسية"
    # A non-empty exact standard value is retained; an unresolved blank stays unclassified.
    if text in ("رئيسية","فرعية","ملحقة"):
        return text
    return "غير محدد"


# Employee payroll exports are HTML tables encoded as UTF-16 despite .xls extension.
from html.parser import HTMLParser
class _EmployeeTableParser(HTMLParser):
    def __init__(self):
        super().__init__();self.rows=[];self.row=None;self.cell=None
    def handle_starttag(self,tag,attrs):
        if tag.lower()=='tr':self.row=[]
        elif tag.lower() in ('td','th') and self.row is not None:self.cell=[]
    def handle_data(self,data):
        if self.cell is not None:self.cell.append(data)
    def handle_endtag(self,tag):
        if tag.lower() in ('td','th') and self.cell is not None and self.row is not None:
            self.row.append(''.join(self.cell).strip());self.cell=None
        elif tag.lower()=='tr' and self.row is not None:
            self.rows.append(self.row);self.row=None

@app.post('/school-import/hire-dates')
async def import_employee_hire_dates(request: Request, files: list[UploadFile]=File(...)):
    user=require_permission(request,'school_import')
    if user['role'] not in ('admin','manager','supervisor'):raise HTTPException(403)
    if not files or len(files)>5:raise HTTPException(400,detail='اختر من ملف إلى خمسة ملفات')
    by_emp={};by_id={};ambiguous_emp=set();ambiguous_id=set()
    for upload in files:
        if not (upload.filename or '').lower().endswith('.xls'):
            raise HTTPException(400,detail='هذه الشاشة تستقبل ملفات Employee بصيغة xls فقط')
        raw=await upload.read(12*1024*1024+1)
        if len(raw)>12*1024*1024:raise HTTPException(400,detail='حجم الملف أكبر من الحد المسموح')
        try:
            decoded=raw.decode('utf-16') if raw.startswith((bytes([255,254]),bytes([254,255]))) else raw.decode('utf-8-sig')
            parser=_EmployeeTableParser();parser.feed(decoded)
            if not parser.rows:raise ValueError('لا يوجد جدول')
            headers=parser.rows[0]
            columns=[headers.index(k) for k in ('Emp_ID','ID_No','DOAppointment')]
        except Exception as exc:
            raise HTTPException(400,detail=f'صيغة ملف الموظفين غير صحيحة: {exc}')
        for row in parser.rows[1:]:
            if len(row)<=max(columns):continue
            emp,identity,hire=(row[i].strip() for i in columns)
            if not hire:continue
            for key,value,mapping,conflicts in ((emp,hire,by_emp,ambiguous_emp),(identity,hire,by_id,ambiguous_id)):
                if not key:continue
                if key in mapping and mapping[key]!=value:conflicts.add(key)
                else:mapping[key]=value
    con=db();updated_guard=updated_supervisor=0
    try:
        school_columns={c['name'] for c in con.execute('PRAGMA table_info(schools)')}
        for col in ('guard_hire_date','coordinator_hire_date'):
            if col not in school_columns:con.execute(f'ALTER TABLE schools ADD COLUMN {col} TEXT')
        for r in con.execute('SELECT ministry_id,guard_emp_id,guard_id,coordinator_emp_id FROM schools').fetchall():
            for field,emp_key,id_key in (('guard_hire_date','guard_emp_id','guard_id'),('coordinator_hire_date','coordinator_emp_id',None)):
                emp=str(r[emp_key] or '').strip();identity=str(r[id_key] or '').strip() if id_key else ''
                date_value=by_emp.get(emp,'') if emp and emp not in ambiguous_emp else ''
                if not date_value and identity and identity not in ambiguous_id:date_value=by_id.get(identity,'')
                if not date_value:continue
                con.execute(f'UPDATE schools SET {field}=? WHERE ministry_id=?',(date_value,r['ministry_id']))
                if field=='guard_hire_date':updated_guard+=1
                else:updated_supervisor+=1
        con.commit()
    except Exception:
        con.rollback();raise
    finally:con.close()
    return RedirectResponse(f'/school-import?hire_guards={updated_guard}&hire_supervisors={updated_supervisor}',status_code=303)


@app.post("/school-import/school-details")
async def import_school_details(request: Request, file: UploadFile=File(...)):
    user=require_permission(request,"school_import")
    if user["role"] not in ("admin","manager","supervisor"):raise HTTPException(403)
    try:
        await file.seek(0)
        wb=load_workbook(file.file,read_only=True,data_only=True)
    except Exception as exc:
        raise HTTPException(400,detail=f"تعذر فتح ملف المدارس: {exc}")
    updates={}
    try:
        if "Data" not in wb.sheetnames:raise HTTPException(400,detail="ملف Master Data يجب أن يحتوي على صفحة Data")
        ws=wb["Data"]
        it=ws.iter_rows(values_only=True)
        headers=[str(v or "").strip() for v in next(it,())]
        def ix(name):
            return headers.index(name) if name in headers else None
        fields={"size":"حجم المبني","gender":"النوع","stage":"المرحلة","classrooms_count":"عدد الفصول",
                "district":"الحي","education_office":"مكتب التعليم","building_ownership":"ملكية المبنى",
                "unit_classification":"تصنيف الوحدة"}
        idcols=[ix(x) for x in ("الرقم الوزاري 1","الرقم الوزاري 2","الرقم الوزاري 3")]
        for values in it:
            record={}
            for field,header in fields.items():
                j=ix(header)
                if j is None or j>=len(values):continue
                v=values[j]
                if v is None or str(v).strip() in ("","#NAME?"):continue
                value=str(v).strip()
                if field=="size":
                    value={"صغير":"صغيرة","متوسط":"متوسطة","كبير":"كبيرة"}.get(value,"")
                    if not value:continue
                record[field]=value
            if not record:continue
            for j in idcols:
                if j is None or j>=len(values):continue
                mid=_norm_excel_value(values[j])
                if not mid:continue
                if mid in updates:
                    # Conflicting records must be reviewed rather than silently overwritten.
                    for k,v in record.items():
                        if k in updates[mid] and updates[mid][k]!=v:
                            updates[mid].pop(k,None)
                        elif k not in updates[mid]:updates[mid][k]=v
                else:updates[mid]=record.copy()
    finally:wb.close()
    if not updates:raise HTTPException(400,detail="لم يتم العثور على بيانات صالحة")
    con=db();matched=0;changed=0
    try:
        cols={r["name"] for r in con.execute("PRAGMA table_info(schools)")}
        for col in ("building_ownership","unit_classification"):
            if col not in cols:con.execute(f"ALTER TABLE schools ADD COLUMN {col} TEXT")
        for mid,record in updates.items():
            if not record:continue
            if not con.execute("SELECT 1 FROM schools WHERE ministry_id=?",(mid,)).fetchone():continue
            matched+=1
            columns=list(record)
            con.execute("UPDATE schools SET "+",".join(f"{k}=?" for k in columns)+" WHERE ministry_id=?",
                        [record[k] for k in columns]+[mid])
            changed+=1
        if not matched:raise HTTPException(400,detail="لم تتم مطابقة أي مدرسة؛ لم يتم تعديل البيانات")
        con.commit()
    except Exception:
        con.rollback();raise
    finally:con.close()
    return RedirectResponse(f"/school-import?details_updated={changed}&details_matched={matched}",status_code=303)


# Monthly payroll: manual daily overtime and deductions, preserving historical months.
def _payroll_schema(con):
    con.executescript("""
    CREATE TABLE IF NOT EXISTS payroll_people (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      role TEXT NOT NULL,
      employee_id TEXT NOT NULL,
      employee_name TEXT NOT NULL,
      base_salary REAL NOT NULL,
      identity_no TEXT NOT NULL DEFAULT '',
      ministry_id TEXT NOT NULL DEFAULT '',
      school_name TEXT NOT NULL DEFAULT '',
      hire_date TEXT NOT NULL DEFAULT '',
      UNIQUE(role,employee_id)
    );
    CREATE TABLE IF NOT EXISTS payroll_template (
      id INTEGER PRIMARY KEY CHECK(id=1),
      workbook BLOB NOT NULL
    );
    CREATE TABLE IF NOT EXISTS payroll_school_order (
      sequence_no INTEGER PRIMARY KEY,
      ministry_id TEXT NOT NULL DEFAULT '',
      school_name TEXT NOT NULL DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS payroll_assignments (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      person_id INTEGER NOT NULL,
      ministry_id TEXT NOT NULL,
      school_name TEXT NOT NULL,
      start_date TEXT NOT NULL,
      end_date TEXT,
      event_type TEXT NOT NULL DEFAULT 'initial',
      created_at TEXT NOT NULL,
      UNIQUE(person_id,ministry_id,start_date)
    );
    CREATE INDEX IF NOT EXISTS idx_payroll_assignments_person ON payroll_assignments(person_id,start_date);
    CREATE TABLE IF NOT EXISTS payroll_terminations (
      person_id INTEGER PRIMARY KEY,
      last_work_date TEXT NOT NULL,
      reason TEXT NOT NULL DEFAULT '',
      recorded_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS payroll_daily (
      person_id INTEGER NOT NULL,
      work_date TEXT NOT NULL,
      attendance_status TEXT NOT NULL DEFAULT '',
      overtime_amount REAL NOT NULL DEFAULT 0,
      deduction_amount REAL NOT NULL DEFAULT 0,
      notes TEXT NOT NULL DEFAULT '',
      PRIMARY KEY(person_id,work_date)
    );
    """)
    cols={x["name"] for x in con.execute("PRAGMA table_info(payroll_people)")}
    if "identity_no" not in cols:con.execute("ALTER TABLE payroll_people ADD COLUMN identity_no TEXT NOT NULL DEFAULT ''")
    for column in ("ministry_id","school_name","hire_date"):
        if column not in cols:con.execute(f"ALTER TABLE payroll_people ADD COLUMN {column} TEXT NOT NULL DEFAULT ''")

def _payroll_sync(con):
    # School master is only a CURRENT snapshot; never fabricate past transfer dates.
    rows=con.execute("""SELECT ministry_id,school_name,guard_emp_id,guard_name,guard_id,
      guard_hire_date,coordinator_emp_id,coordinator_name,coordinator_hire_date FROM schools""").fetchall()
    for r in rows:
        for role,emp,name,salary,identity,hire in (
            ("guard",r["guard_emp_id"],r["guard_name"],3000,r["guard_id"],r["guard_hire_date"]),
            ("supervisor",r["coordinator_emp_id"],r["coordinator_name"],4500,"",r["coordinator_hire_date"])):
            if not name or not str(name).strip():continue
            emp=str(emp or "").strip() or ("NAME:"+str(name).strip())
            con.execute("""INSERT OR IGNORE INTO payroll_people(role,employee_id,employee_name,base_salary,
              identity_no,ministry_id,school_name,hire_date) VALUES(?,?,?,?,?,?,?,?)""",
              (role,emp,str(name).strip(),salary,str(identity or ""),str(r["ministry_id"] or ""),
               str(r["school_name"] or ""),str(hire or "")))
            con.execute("""UPDATE payroll_people SET
              identity_no=CASE WHEN identity_no='' THEN ? ELSE identity_no END,
              hire_date=CASE WHEN hire_date='' THEN ? ELSE hire_date END,
              ministry_id=CASE WHEN ministry_id='' THEN ? ELSE ministry_id END,
              school_name=CASE WHEN school_name='' THEN ? ELSE school_name END
              WHERE role=? AND employee_id=?""",
              (str(identity or ""),str(hire or ""),str(r["ministry_id"] or ""),
               str(r["school_name"] or ""),role,emp))
            p=con.execute("SELECT id,hire_date FROM payroll_people WHERE role=? AND employee_id=?",(role,emp)).fetchone()
            # Seed only once. A known hire date is a factual lower bound; otherwise
            # use today's date and require manual backdating for historical cycles.
            if not con.execute("SELECT 1 FROM payroll_assignments WHERE person_id=?",(p["id"],)).fetchone() and not con.execute("SELECT 1 FROM payroll_terminations WHERE person_id=?",(p["id"],)).fetchone():
                hire_day=_payroll_hire_date(p["hire_date"])
                effective=hire_day if hire_day and hire_day<=date.today() else date.today()
                con.execute("""INSERT OR IGNORE INTO payroll_assignments
                  (person_id,ministry_id,school_name,start_date,event_type,created_at)
                  VALUES(?,?,?,?,'initial',?)""",
                  (p["id"],str(r["ministry_id"] or ""),str(r["school_name"] or ""),
                   effective.isoformat(),datetime.now().isoformat(timespec="seconds")))

def _payroll_find_person(con,search):
    key=_payroll_normalize(search)
    if not key:raise HTTPException(400,"ادخل رقم الموظف أو هويته أو اسمه")
    people=con.execute("SELECT * FROM payroll_people").fetchall()
    exact=[p for p in people if key in (_payroll_normalize(p["employee_id"]),_payroll_normalize(p["identity_no"]),_payroll_normalize(p["employee_name"]))]
    matches=exact if exact else [p for p in people if key in _payroll_normalize(p["employee_name"])]
    if not matches:raise HTTPException(404,"الموظف غير موجود")
    if len(matches)!=1:raise HTTPException(409,"أكثر من موظف مطابق؛ استخدم الرقم الوظيفي أو الهوية")
    return matches[0]

def _payroll_apply_movement(con,person,action,effective_date,ministry_id="",reason=""):
    # One open assignment at most; date ranges use inclusive endpoints.
    if action not in ("transfer","termination"):raise HTTPException(400,"حركة غير صالحة")
    try:effective=date.fromisoformat(effective_date)
    except ValueError:raise HTTPException(400,"تاريخ غير صالح")
    open_rows=con.execute("""SELECT * FROM payroll_assignments WHERE person_id=? AND end_date IS NULL
      ORDER BY start_date DESC""",(person["id"],)).fetchall()
    if len(open_rows)!=1:raise HTTPException(409,"يجب مراجعة فترات تسكين الموظف قبل تسجيل الحركة")
    old=open_rows[0]
    if effective<date.fromisoformat(old["start_date"]) or (action=="transfer" and effective==date.fromisoformat(old["start_date"])):raise HTTPException(409,"تاريخ الحركة لا يتوافق مع بداية التسكين الحالي")
    if con.execute("SELECT 1 FROM payroll_terminations WHERE person_id=?",(person["id"],)).fetchone():
        raise HTTPException(409,"الموظف منتهي الخدمة بالفعل")
    if action=="transfer":
        school=con.execute("SELECT ministry_id,school_name FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
        if not school:raise HTTPException(404,"الرقم الوزاري للمدرسة الجديدة غير موجود")
        if old["ministry_id"]==str(ministry_id):raise HTTPException(409,"الموظف مسكن بالفعل في المدرسة المحددة")
        con.execute("UPDATE payroll_assignments SET end_date=? WHERE id=?",((effective-timedelta(days=1)).isoformat(),old["id"]))
        con.execute("""INSERT INTO payroll_assignments(person_id,ministry_id,school_name,start_date,event_type,created_at)
          VALUES(?,?,?,?,'transfer',?)""",(person["id"],str(ministry_id),str(school["school_name"] or ""),
          effective.isoformat(),datetime.now().isoformat(timespec="seconds")))
    else:
        # Termination date is LAST paid work day, not first day off payroll.
        con.execute("UPDATE payroll_assignments SET end_date=? WHERE id=?",(effective.isoformat(),old["id"]))
        con.execute("""INSERT INTO payroll_terminations(person_id,last_work_date,reason,recorded_at)
          VALUES(?,?,?,?)""",(person["id"],effective.isoformat(),reason[:500],datetime.now().isoformat(timespec="seconds")))

@app.post("/payroll/movement")
def payroll_movement(request:Request,search:str=Form(...),action:str=Form(...),
                     effective_date:str=Form(...),ministry_id:str=Form(""),reason:str=Form("")):
    user=require_permission(request,"school_import")
    if user["role"] not in ("admin","manager"):raise HTTPException(403)
    con=db()
    try:
        _payroll_schema(con)
        _payroll_sync(con)
        person=_payroll_find_person(con,search)
        _payroll_apply_movement(con,person,action,effective_date,ministry_id.strip(),reason)
        con.commit()
    finally:con.close()
    day=date.fromisoformat(effective_date)
    label=(day if day.day<=23 else day+timedelta(days=12)).strftime("%Y-%m")
    return RedirectResponse("/payroll?month="+label,status_code=303)

def _payroll_period(month):
    # Payroll labeled YYYY-MM ends on the 23rd of that month.
    year,number=map(int,month.split("-"))
    end=date(year,number,23)
    previous=(date(year,number,1)-timedelta(days=1))
    return date(previous.year,previous.month,24),end

def _payroll_hire_date(value):
    value=str(value or "").strip()
    for fmt in ("%Y-%m-%d","%Y/%m/%d","%d/%m/%Y","%d-%m-%Y"):
        try:return datetime.strptime(value[:10],fmt).date()
        except ValueError:pass
    return None

@app.post("/payroll/export-order")
async def payroll_import_export_order(request:Request,file:UploadFile=File(...)):
    user=require_permission(request,"school_import")
    if user["role"] not in ("admin","manager"):raise HTTPException(403)
    raw=await file.read()
    if len(raw)>15*1024*1024:raise HTTPException(413,"الملف كبير")
    try:
        wb=load_workbook(BytesIO(raw),read_only=True,data_only=True)
        ws=wb["مسير البوابين"] if "مسير البوابين" in wb.sheetnames else wb.active
        def normalize_cell(value):
            return _payroll_normalize(value).replace(" ","").replace("_","")
        # October reference layout: B=sequence, C=school, D=ministry IDs.
        # C/D headers are blank in the supplied workbook; rely on its fixed layout.
        records=[]
        seen=set()
        for line,row in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
            school=str(row[2] or "").strip() if len(row)>2 else ""
            ministry=str(row[3] or "").strip() if len(row)>3 else ""
            if not school:continue
            if ministry.endswith(".0"):ministry=ministry[:-2]
            key=("id",ministry) if ministry else ("name",_payroll_normalize(school))
            if key in seen:continue
            seen.add(key)
            records.append((len(records)+1,ministry,school))
        if not records:raise ValueError("لم يتم العثور على مدارس في العمود C")
        wb.close()
    except Exception as exc:
        raise HTTPException(400,"تعذر قراءة تسلسل المسير: "+str(exc)[:120])
    con=db()
    try:
        _payroll_schema(con)
        con.execute("INSERT OR REPLACE INTO payroll_template(id,workbook) VALUES(1,?)",(raw,))
        con.execute("DELETE FROM payroll_school_order")
        for seq,ministry,school in records:
            con.execute("""INSERT INTO payroll_school_order(sequence_no,ministry_id,school_name)
              VALUES(?,?,?)""",(seq,ministry,school))
        con.commit()
    finally:con.close()
    return RedirectResponse("/payroll?order_imported="+str(len(records)),status_code=303)

@app.get("/payroll/export")
def payroll_export_excel(request:Request,month:str=""):
    require_permission(request,"school_import")
    month=month or (date.today() if date.today().day<=23 else date.today()+timedelta(days=12)).strftime("%Y-%m")
    if not re.fullmatch(r"20\d{2}-(0[1-9]|1[0-2])",month):
        raise HTTPException(400,"شهر غير صالح")
    period_start,period_end=_payroll_period(month)
    con=db()
    try:
        _payroll_schema(con)
        _payroll_sync(con)
        con.commit()
        people=con.execute("""SELECT p.*,
            COALESCE(SUM(d.overtime_amount),0) overtime,
            COALESCE(SUM(d.deduction_amount),0) deductions,
            SUM(CASE WHEN d.attendance_status='غائب' THEN 1 ELSE 0 END) absent_days
            FROM payroll_people p LEFT JOIN payroll_daily d
            ON p.id=d.person_id AND d.work_date BETWEEN ? AND ?
            GROUP BY p.id ORDER BY p.school_name,p.employee_name""",
            (period_start.isoformat(),period_end.isoformat())).fetchall()
        assignments=con.execute("""SELECT * FROM payroll_assignments
            WHERE start_date<=? AND (end_date IS NULL OR end_date>=?)
            ORDER BY school_name,person_id,start_date""",
            (period_end.isoformat(),period_start.isoformat())).fetchall()
        termination={r["person_id"]:r["last_work_date"] for r in con.execute(
            "SELECT person_id,last_work_date FROM payroll_terminations").fetchall()}
        grouped={}
        for a in assignments:grouped.setdefault(a["person_id"],[]).append(a)
        template=con.execute("SELECT workbook FROM payroll_template WHERE id=1").fetchone()
        wb=load_workbook(BytesIO(template["workbook"])) if template else Workbook()
        ws=wb["مسير البوابين"] if template and "مسير البوابين" in wb.sheetnames else wb.active
        if template:
            return _payroll_export_from_template(wb,ws,people,grouped,termination,period_start,period_end,month)

        ws.title="مسير "+month
        ws.sheet_view.rightToLeft=True
        ws.append(["المدرسة","الرقم الوزاري","اسم الموظف","الوظيفة","الرقم الوظيفي","الهوية",
                   "تاريخ التعيين","بداية التسكين","نهاية التسكين","أيام المدرسة",
                   "إجمالي أيام الموظف","الغياب","الراتب المستحق","الإضافي",
                   "خصم الغياب","خصومات أخرى","إضافة يدوية","خصم يدوي","الصافي بعد التعديل","ملاحظات"])
        rows=[]
        supervisor_rows=[]
        for p in people:
            segments=[]
            for a in grouped.get(p["id"],[]):
                first=max(period_start,date.fromisoformat(a["start_date"]))
                last=min(period_end,date.fromisoformat(a["end_date"]) if a["end_date"] else period_end)
                hire=_payroll_hire_date(p["hire_date"])
                if hire:first=max(first,hire)
                term=_payroll_hire_date(termination.get(p["id"]))
                if term:last=min(last,term)
                if first<=last:segments.append((a,first,last,(last-first).days+1))
            if not segments:continue
            days=min(30,sum(item[3] for item in segments))
            base=round(float(p["base_salary"])*days/30,2)
            absence=(int(p["absent_days"] or 0))*(200 if p["role"]=="guard" else 300)
            if p["role"]=="supervisor":
                supervisor_rows.append((p["employee_name"],p["employee_id"],p["identity_no"],p["hire_date"],days,int(p["absent_days"] or 0),base,float(p["overtime"]),absence,float(p["deductions"])))
                continue
            for index,(a,first,last,segment_days) in enumerate(segments):
                rows.append((a["school_name"] or "",a["ministry_id"] or "",p["employee_name"],
                    "بواب" if p["role"]=="guard" else "مشرف",p["employee_id"],p["identity_no"],
                    p["hire_date"],first.isoformat(),last.isoformat(),segment_days,
                    days if index==0 else None,int(p["absent_days"] or 0) if index==0 else None,
                    base if index==0 else None,float(p["overtime"]) if index==0 else None,
                    absence if index==0 else None,float(p["deductions"]) if index==0 else None,
                    None,None,None,"" if index==0 else "صف توزيع فقط؛ الراتب في الصف الأول للموظف"))
        ordering=con.execute("SELECT ministry_id,school_name,sequence_no FROM payroll_school_order").fetchall()
        order_by_ministry={str(r["ministry_id"]).strip():r["sequence_no"] for r in ordering if r["ministry_id"]}
        order_by_school={_payroll_normalize(r["school_name"]):r["sequence_no"] for r in ordering if r["school_name"]}
        def school_sequence(row):
            ministry=str(row[1]).strip()
            school=_payroll_normalize(row[0])
            return order_by_ministry.get(ministry,order_by_school.get(school,10**9))
        rows.sort(key=lambda row:(school_sequence(row),str(row[0]),str(row[2]),str(row[7])))
        existing={(str(x[1]).strip(),_payroll_normalize(x[0])) for x in rows}
        for o in ordering:
            key=(str(o["ministry_id"] or "").strip(),_payroll_normalize(o["school_name"]))
            if key not in existing:
                rows.append((o["school_name"],o["ministry_id"],"","","","","","","",None,None,None,None,None,None,None,None,None,None,"لا يوجد بواب مسجل"))
        rows.sort(key=lambda x:(school_sequence(x),str(x[0]),str(x[2])))
        for item in rows:
            ws.append(list(item))
            n=ws.max_row
            if item[12] is not None:
                ws.cell(n,19,f"=M{n}+N{n}-O{n}-P{n}+Q{n}-R{n}")
        from openpyxl.styles import Font,PatternFill,Alignment
        from openpyxl.utils import get_column_letter
        sup=wb.create_sheet("مسير المشرفين")
        sup.sheet_view.rightToLeft=True
        sup.append(["اسم المشرف","الرقم الوظيفي","الهوية","تاريخ التعيين","أيام الاستحقاق","الغياب","الراتب المستحق","الإضافي","خصم الغياب","خصومات أخرى","إضافة يدوية","خصم يدوي","الصافي"])
        for record in supervisor_rows:
            sup.append(list(record)+[None,None,None])
            n=sup.max_row
            sup.cell(n,13,f"=G{n}+H{n}-I{n}-J{n}+K{n}-L{n}")
        sup.freeze_panes="C2"
        sup.auto_filter.ref=sup.dimensions
        for col in range(1,14):
            sup.column_dimensions[get_column_letter(col)].width=34 if col==1 else 20
        for cell in sup[1]:
            cell.fill=PatternFill("solid",fgColor="17365D")
            cell.font=Font(color="FFFFFF",bold=True)
        for row in sup.iter_rows(min_row=2):
            for col in (11,12):row[col-1].fill=PatternFill("solid",fgColor="FFF2CC")
        from openpyxl.utils import get_column_letter
        for c in ws[1]:
            c.fill=PatternFill("solid",fgColor="17365D")
            c.font=Font(color="FFFFFF",bold=True)
            c.alignment=Alignment(horizontal="center")
        widths=[30,17,31,13,19,19,17,17,17,16,18,13,18,16,17,17,17,17,21,43]
        for i,width in enumerate(widths,1):
            ws.column_dimensions[get_column_letter(i)].width=width
        for row in ws.iter_rows(min_row=2):
            for col in (17,18):
                row[col-1].fill=PatternFill("solid",fgColor="FFF2CC")
            for col in (13,14,15,16,17,18,19):
                row[col-1].number_format='#,##0.00'
        ws.freeze_panes="D2"
        ws.auto_filter.ref=ws.dimensions
        summary=wb.create_sheet("تعليمات")
        summary.sheet_view.rightToLeft=True
        for line in [
            ["مسير الرواتب",month],
            ["بداية الفترة",period_start.isoformat()],
            ["نهاية الفترة",period_end.isoformat()],
            ["طريقة الحساب","الراتب الأساسي ÷ 30 × أيام الخدمة بحد أقصى 30 يوم"],
            ["خصم غياب البواب","200 ريال / يوم"],
            ["خصم غياب المشرف","300 ريال / يوم"],
            ["الأعمدة الصفراء","Q إضافة يدوية، R خصم يدوي"],
            ["هام","ترتيب مدارس أكتوبر للبوابين فقط؛ المشرفون في شيت منفصل دون مدارس"],
            ["تنبيه","التعديل على ملف Excel لا يحدث بيانات البرنامج تلقائيًا"],
            ["إجمالي الصافي بعد التعديل","='مسير "+month+"'!S2" if len(rows)==1 else "=SUM('مسير "+month+"'!S2:S"+str(max(2,ws.max_row))+")"]
        ]:summary.append(line)
        summary.column_dimensions["A"].width=33
        summary.column_dimensions["B"].width=90
        output=BytesIO()
        wb.save(output)
        output.seek(0)
    finally:con.close()
    return StreamingResponse(output,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      headers={"Content-Disposition":'attachment; filename="payroll-'+month+'.xlsx"'})

@app.get("/payroll",response_class=HTMLResponse)
def payroll_page(request: Request,month: str=""):
    require_permission(request,"school_import")
    from html import escape
    month=month or (date.today() if date.today().day<=23 else date.today()+timedelta(days=12)).strftime("%Y-%m")
    if not re.fullmatch(r"20\d{2}-(0[1-9]|1[0-2])",month):
        raise HTTPException(400,"شهر غير صالح")
    period_start,period_end=_payroll_period(month)
    con=db()
    try:
        _payroll_schema(con)
        _payroll_sync(con)
        con.commit()
        people=con.execute("""SELECT p.*,
          COALESCE(SUM(d.overtime_amount),0) overtime,
          COALESCE(SUM(d.deduction_amount),0) deductions,
          SUM(CASE WHEN d.attendance_status='غائب' THEN 1 ELSE 0 END) absent_days
          FROM payroll_people p LEFT JOIN payroll_daily d
          ON p.id=d.person_id AND d.work_date BETWEEN ? AND ?
          GROUP BY p.id ORDER BY p.school_name,p.role,p.employee_name""",
          (period_start.isoformat(),period_end.isoformat())).fetchall()
        assignments=con.execute("""SELECT * FROM payroll_assignments
          WHERE start_date<=? AND (end_date IS NULL OR end_date>=?)
          ORDER BY school_name,person_id,start_date""",
          (period_end.isoformat(),period_start.isoformat())).fetchall()
        by_person={}
        for row in assignments:by_person.setdefault(row["person_id"],[]).append(row)
        terminations={r["person_id"]:r["last_work_date"] for r in
          con.execute("SELECT person_id,last_work_date FROM payroll_terminations").fetchall()}
        def segments(person):
            hire=_payroll_hire_date(person["hire_date"])
            termination=_payroll_hire_date(terminations.get(person["id"]))
            out=[]
            for assignment in by_person.get(person["id"],[]):
                first=max(period_start,date.fromisoformat(assignment["start_date"]))
                last=min(period_end,date.fromisoformat(assignment["end_date"]) if assignment["end_date"] else period_end)
                if hire:first=max(first,hire)
                if termination:last=min(last,termination)
                if first<=last:out.append((assignment,(last-first).days+1))
            return out
        def payroll_amount(person):
            days=min(30,sum(days for _,days in segments(person)))
            base=round(float(person["base_salary"])*days/30,2)
            absent=int(person["absent_days"] or 0)
            absence_cost=absent*(200 if person["role"]=="guard" else 300)
            return days,base,absence_cost,base+float(person["overtime"])-float(person["deductions"])-absence_cost
        active_people=[person for person in people if segments(person)]
        total=sum(payroll_amount(person)[3] for person in active_people)
        absence_total=sum(int(person["absent_days"] or 0) for person in active_people)
        opts="".join('<option value="'+str(person["id"])+'">'+escape(person["employee_name"])+' ('+escape(person["employee_id"])+')</option>' for person in active_people)
        # School rows are distribution rows. The financial total is counted ONCE per person.
        # Monetary columns are displayed only on the first school row to avoid duplication.
        school_rows=[]
        for person in active_people:
            for idx,(assignment,days) in enumerate(segments(person)):
                school_rows.append((assignment["school_name"] or "غير محددة",person,days,idx))
        school_rows.sort(key=lambda item:(item[0],item[1]["employee_name"]))
        trs="".join("<tr><td>"+escape(school)+"</td><td>"+escape(person["employee_name"])+"</td><td>"+("بواب" if person["role"]=="guard" else "مشرف")+"</td><td>"+escape(person["employee_id"])+"</td><td>"+escape(person["hire_date"] or "غير مسجل")+"</td><td>"+str(days)+"</td><td>"+(str(person["absent_days"] or 0) if idx==0 else "—")+"</td><td>"+(f'{payroll_amount(person)[1]:,.2f}' if idx==0 else "—")+"</td><td>"+(f'{person["overtime"]:,.2f}' if idx==0 else "—")+"</td><td>"+(f'{payroll_amount(person)[2]:,.2f}' if idx==0 else "—")+"</td><td>"+(f'{person["deductions"]:,.2f}' if idx==0 else "—")+"</td><td>"+(f'{payroll_amount(person)[3]:,.2f}' if idx==0 else "—")+"</td></tr>" for school,person,days,idx in school_rows)
        return HTMLResponse("""<!doctype html><html lang="ar" dir="rtl"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>مسير الرواتب</title>
        <style>body{{font-family:Tahoma,Arial;background:#f3f6fb;color:#152b4a;margin:20px}}section{{background:white;padding:18px;border-radius:12px;margin:12px 0}}input,select,button{{padding:10px;margin:5px;border:1px solid #bbc7d8;border-radius:6px}}button{{background:#143b72;color:white;cursor:pointer}}table{{border-collapse:collapse;width:100%}}td,th{{padding:9px;border:1px solid #d8e0e9;text-align:right}}th{{background:#143b72;color:white}}.scroll{{overflow:auto}}</style>
        <h1>مسير الرواتب الشهري</h1><p><a href="/">الرئيسية</a> | البواب 3000 ريال، المشرف 4500 ريال | خصم الغياب: 200 للبواب و300 للمشرف</p>
        <section><form method="get"><label>شهر المسير <input type="month" name="month" value="{month}"></label><button>عرض المسير</button></form>
        <h3>الفترة: {period_start} إلى {period_end}</h3><p><a href="/payroll/export?month={month}" style="display:inline-block;padding:11px;background:#167347;color:white;border-radius:7px;text-decoration:none">تصدير مسير Excel للتعديل اليدوي</a></p><h2>إجمالي المستحقات: {total:,.2f} ريال</h2>
        <p>الحساب يعتمد على فترات التسكين التاريخية المسجلة، ولا يستنتج تواريخ النقل القديمة تلقائيًا. جميع الموظفين حاضرون افتراضيًا؛ الغياب المسجل {absence_total} يوم. الموظف الجديد يبدأ احتسابه من تاريخ التعيين. أساس تجزئة راتب المنضم خلال الدورة: 30 يومًا.</p></section>
        <section><h2>تسلسل المدارس في المسير الأصلي</h2><p>ارفع مسير أكتوبر مرة واحدة لتثبيت ترتيب المدارس لجميع الشهور القادمة. لا يتغير الترتيب عند اختيار شهر مختلف، ولا يُستبدل إلا برفع مرجع جديد عمدًا.</p>
        <form method="post" action="/payroll/export-order" enctype="multipart/form-data"><input type="file" name="file" accept=".xlsx" required><button>اعتماد تسلسل المسير</button></form></section>
        <section><h2>تسجيل نقل أو إنهاء خدمة</h2><p>صلاحية المدير فقط. تاريخ النقل هو أول يوم في المدرسة الجديدة، وتاريخ الفصل هو آخر يوم عمل مستحق.</p>
        <form method="post" action="/payroll/movement"><label>الموظف <input name="search" required placeholder="الرقم الوظيفي / الهوية / الاسم"></label>
        <label>الحركة <select name="action"><option value="transfer">نقل</option><option value="termination">إنهاء خدمة</option></select></label>
        <label>تاريخ الحركة <input type="date" name="effective_date" required></label>
        <label>الرقم الوزاري للمدرسة الجديدة (للنقل) <input name="ministry_id"></label>
        <label>سبب / ملاحظات <input name="reason" maxlength="500"></label><button>تسجيل الحركة</button></form></section>
        <section><h2>تسجيل غياب بالرقم الوظيفي أو الاسم أو الهوية</h2><form method="post" action="/payroll/absence"><label>الموظف <input name="search" required placeholder="الرقم الوظيفي / الهوية / الاسم"></label><label>اليوم <input type="date" name="work_date" value="{today}" min="{period_start}" max="{period_end}" required></label><label>ملاحظة <input name="notes" maxlength="500"></label><button>تسجيل غياب</button></form></section>
        <section><h2>الإضافي والخصومات وتعديل حالة يوم</h2><form method="post" action="/payroll/daily">
        <label>الموظف <select name="person_id" required>{opts}</select></label>
        <label>اليوم <input type="date" name="work_date" value="{today}" min="{period_start}" max="{period_end}" required></label>
        <label>الحالة <select name="attendance_status"><option value="حاضر">حاضر</option><option>غائب</option><option>إجازة</option><option>راحة</option></select></label>
        <label>الإضافي بالريال <input name="overtime_amount" type="number" min="0" step="0.01" value="0"></label>
        <label>خصومات أخرى <input name="deduction_amount" type="number" min="0" step="0.01" value="0"></label>
        <label>ملاحظات <input name="notes" maxlength="500"></label><button>حفظ</button></form></section>
        <section class="scroll"><h2>كشف المسير حسب المدرسة</h2><p>المنقول يظهر تحت كل مدرسة بأيامها؛ المبالغ تظهر مرة واحدة فقط لكل موظف، ولا تُكرر عند جمع المسير.</p><table><thead><tr><th>المدرسة</th><th>الموظف</th><th>الوظيفة</th><th>الرقم الوظيفي</th><th>تاريخ التعيين</th><th>أيام الفترة المحتسبة</th><th>غياب</th><th>الأساسي المستحق</th><th>الإضافي</th><th>خصم الغياب</th><th>خصومات أخرى</th><th>الإجمالي</th></tr></thead><tbody>{trs}</tbody></table></section>
        </html>""".format(month=month,period_start=period_start,period_end=period_end,total=total,
          absence_total=absence_total,opts=opts,trs=trs,
          today=min(max(date.today(),period_start),period_end)))
    finally:con.close()

def _payroll_normalize(value):
    value=str(value or "").strip().casefold()
    for a,b in (("أ","ا"),("إ","ا"),("آ","ا"),("ى","ي"),("ة","ه"),("ـ","")):
        value=value.replace(a,b)
    return " ".join(value.split())

@app.post("/payroll/absence")
def payroll_record_absence(request:Request,search:str=Form(...),work_date:str=Form(...),notes:str=Form("")):
    require_permission(request,"school_import")
    try:day=date.fromisoformat(work_date)
    except ValueError:raise HTTPException(400,"تاريخ غير صالح")
    if day.year<2020 or day.year>2100:raise HTTPException(400,"تاريخ غير صالح")
    key=_payroll_normalize(search)
    if not key or len(key)>150:raise HTTPException(400,"ادخل اسم الموظف أو رقمه أو هويته")
    con=db()
    try:
        _payroll_schema(con)
        _payroll_sync(con)
        con.commit()
        people=con.execute("SELECT id,employee_id,employee_name,identity_no FROM payroll_people").fetchall()
        exact=[p for p in people if key in (_payroll_normalize(p["employee_id"]),_payroll_normalize(p["identity_no"]),_payroll_normalize(p["employee_name"]))]
        matches=exact if exact else [p for p in people if key in _payroll_normalize(p["employee_name"])]
        if not matches:raise HTTPException(404,"لم يتم العثور على موظف مطابق؛ راجع الاسم أو الرقم أو الهوية")
        if len(matches)!=1:raise HTTPException(409,"وجدت أكثر من موظف مطابق؛ استخدم رقم الموظف أو الهوية بشكل محدد")
        person=matches[0]
        if not con.execute("""SELECT 1 FROM payroll_assignments WHERE person_id=? AND start_date<=?
          AND (end_date IS NULL OR end_date>=?)""",(person["id"],work_date,work_date)).fetchone():
            raise HTTPException(409,"تاريخ الغياب خارج فترة خدمة الموظف المسجلة")
        previous=con.execute("SELECT overtime_amount,deduction_amount FROM payroll_daily WHERE person_id=? AND work_date=?",(person["id"],work_date)).fetchone()
        con.execute("""INSERT INTO payroll_daily(person_id,work_date,attendance_status,overtime_amount,deduction_amount,notes)
          VALUES(?,?,'غائب',?,?,?) ON CONFLICT(person_id,work_date) DO UPDATE SET
          attendance_status='غائب',notes=excluded.notes""",
          (person["id"],work_date,float(previous["overtime_amount"]) if previous else 0,float(previous["deduction_amount"]) if previous else 0,notes[:500]))
        con.commit()
    finally:con.close()
    return RedirectResponse("/payroll?month="+(date.fromisoformat(work_date) if int(work_date[8:10])<=23 else date.fromisoformat(work_date)+timedelta(days=12)).strftime("%Y-%m"),status_code=303)

@app.post("/payroll/daily")
def payroll_daily_save(request:Request,person_id:int=Form(...),work_date:str=Form(...),
                       attendance_status:str=Form(""),overtime_amount:float=Form(0),
                       deduction_amount:float=Form(0),notes:str=Form("")):
    require_permission(request,"school_import")
    try:day=date.fromisoformat(work_date)
    except ValueError:raise HTTPException(400,"تاريخ غير صالح")
    if day.year<2020 or day.year>2100 or attendance_status not in ("","حاضر","غائب","إجازة","راحة"):
        raise HTTPException(400,"بيانات غير صالحة")
    import math
    if not all(math.isfinite(x) and 0<=x<=1000000 for x in (overtime_amount,deduction_amount)):
        raise HTTPException(400,"المبلغ غير صالح")
    con=db()
    try:
        _payroll_schema(con)
        if not con.execute("SELECT 1 FROM payroll_people WHERE id=?",(person_id,)).fetchone():
            raise HTTPException(404,"الموظف غير موجود")
        if not con.execute("""SELECT 1 FROM payroll_assignments WHERE person_id=? AND start_date<=?
          AND (end_date IS NULL OR end_date>=?)""",(person_id,work_date,work_date)).fetchone():
            raise HTTPException(409,"التاريخ خارج فترة خدمة الموظف المسجلة")
        con.execute("""INSERT INTO payroll_daily(person_id,work_date,attendance_status,overtime_amount,deduction_amount,notes)
          VALUES(?,?,?,?,?,?) ON CONFLICT(person_id,work_date) DO UPDATE SET
          attendance_status=excluded.attendance_status,overtime_amount=excluded.overtime_amount,
          deduction_amount=excluded.deduction_amount,notes=excluded.notes""",
          (person_id,work_date,attendance_status,overtime_amount,deduction_amount,notes[:500]))
        con.commit()
    finally:con.close()
    return RedirectResponse("/payroll?month="+work_date[:7],status_code=303)

@app.get("/school-import", response_class=HTMLResponse)
def school_import_page(request: Request):
    user=require_permission(request,"school_import")
    if user["role"] not in ("admin","manager","supervisor"):
        raise HTTPException(403)
    con=db()
    batches=con.execute("""
      SELECT b.*,u.full_name imported_by_name
      FROM school_import_batches b LEFT JOIN users u ON u.id=b.imported_by
      ORDER BY b.id DESC LIMIT 20
    """).fetchall()
    changes=con.execute("""
      SELECT c.*,u.full_name changed_by_name
      FROM school_change_log c LEFT JOIN users u ON u.id=c.changed_by
      ORDER BY c.id DESC LIMIT 200
    """).fetchall()
    con.close()
    return templates.TemplateResponse("school_import.html",{
      "request":request,"user":user,"batches":batches,"changes":changes
    })

@app.post("/school-import")
async def school_import_execute(request: Request, file: UploadFile=File(...), replace_master: str=Form(""), confirmation: str=Form("")):
    user=current_user(request)
    if not user or user["role"] not in ("admin","manager","supervisor"): raise HTTPException(403)
    if replace_master!="yes" or confirmation.strip()!="استبدال البيانات": raise HTTPException(400,detail="يلزم تفعيل الاستبدال وكتابة: استبدال البيانات")
    # Stream workbook from the spooled upload file to reduce memory usage.
    try:
        await file.seek(0)
        wb=load_workbook(file.file,data_only=True,read_only=True)
    except Exception as exc:
        raise HTTPException(400,detail=f"تعذر قراءة ملف Excel: {exc}")
    rows=[]
    try:
      for ws in wb.worksheets:
        headers=[str(c or "").strip() for c in next(ws.iter_rows(min_row=1,max_row=1,values_only=True),())]
        if not any(h in headers for h in ("الرقم الوزاري","رقم وزاري","ministry_id")): continue
        for values in ws.iter_rows(min_row=2,values_only=True):
            row=dict(zip(headers,values));mid=_pick(row,"الرقم الوزاري","رقم وزاري","ministry_id")
            if mid and mid not in DUPLICATE_SCHOOL_IDS:
                rows.append(row)
    finally:
      wb.close()
    if not rows: raise HTTPException(400,detail="لم يتم العثور على بيانات مدارس")
    # October Qassim placement sheet: replace ONLY current personnel/assignments.
    # Never delete schools or historical attendance, documents or assignment logs.
    if any("البواب" == str(k).strip() for k in rows[0]) and any("اسم المشرف" == str(k).strip() for k in rows[0]):
        incoming={}
        for row in rows:
            mid=_pick(row,"الرقم الوزاري","رقم وزاري","ministry_id")
            if mid in incoming:
                raise HTTPException(400,detail=f"رقم وزاري مكرر في الملف: {mid}")
            incoming[mid]=row
        con=db()
        try:
            known={r["ministry_id"] for r in con.execute("SELECT ministry_id FROM schools")}
            matched=known.intersection(incoming)
            if not matched:
                raise HTTPException(400,detail="لا توجد أرقام وزارية مطابقة للمدارس الحالية؛ لم يتم تغيير أي بيانات")
            if len(matched)*2 < len(incoming):
                raise HTTPException(400,detail=f"نسبة المطابقة منخفضة ({len(matched)} من {len(incoming)})؛ تم إلغاء الاستبدال لحماية البيانات")
            now=datetime.utcnow().isoformat()
            for table, cols in {"schools":("guard_hire_date","coordinator_hire_date"),"guard_master":("hire_date",),"coordinator_master":("hire_date",),"personnel_master":("hire_date",)}.items():
                existing={c["name"] for c in con.execute(f"PRAGMA table_info({table})")}
                for col in cols:
                    if col not in existing: con.execute(f"ALTER TABLE {table} ADD COLUMN {col} TEXT")
            # Clear old CURRENT placements, including schools absent from the new file.
            con.execute("UPDATE schools SET guard_name='',guard_emp_id='',guard_id='',guard_mobile='',coordinator_name='',coordinator_emp_id='',coordinator_mobile='',guard_hire_date='',coordinator_hire_date=''")
            con.execute("DELETE FROM personnel_master WHERE role_title IN ('بواب','منسق / مشرف')")
            con.execute("DELETE FROM guard_master")
            con.execute("DELETE FROM coordinator_master")
            guards={};supervisors={}
            for mid in matched:
                row=incoming[mid]
                guard=_pick(row,"البواب");gmobile=_pick(row,"رقم جوال البواب");gid=_pick(row,"رقم هوية البواب")
                supervisor=_pick(row,"اسم المشرف");smobile=_pick(row,"رقم جوال المشرف");sid=_pick(row,"رقم هوية المشرف")
                ghire=_pick(row,"تاريخ تعيين البواب") if guard and _pick(row,"مطابقة البواب")=="مطابق" else ""
                shire=_pick(row,"تاريخ تعيين المشرف") if supervisor and _pick(row,"مطابقة المشرف")=="مطابق" else ""
                gemp=("G-"+gid) if gid else ("G-"+guard) if guard else ""
                semp=("S-"+sid) if sid else ("S-"+supervisor) if supervisor else ""
                con.execute("UPDATE schools SET guard_name=?,guard_emp_id=?,guard_id=?,guard_mobile=?,coordinator_name=?,coordinator_emp_id=?,coordinator_mobile=?,guard_hire_date=?,coordinator_hire_date=? WHERE ministry_id=?",
                    (guard,gemp,gid,gmobile,supervisor,semp,smobile,ghire,shire,mid))
                if guard and gemp:
                    guards[gemp]=(guard,gid,gmobile,mid,ghire)
                if supervisor and semp:
                    supervisors[semp]=(supervisor,sid,smobile,mid,shire)
            for emp,(name,id_no,mobile,mid,hire_date) in guards.items():
                con.execute("INSERT INTO guard_master(emp_id,emp_name,id_no,mobile,status,linked_ministry_id,hire_date) VALUES(?,?,?,?,?,?,?)",
                    (emp,name,id_no,mobile,"نشط",mid,hire_date))
                con.execute("INSERT INTO personnel_master(emp_id,full_name,id_no,mobile,role_title,linked_ministry_id,is_active,imported_at,hire_date) VALUES(?,?,?,?,?,?,1,?,?)",
                    (emp,name,id_no,mobile,"بواب",mid,now,hire_date))
            for emp,(name,id_no,mobile,mid,hire_date) in supervisors.items():
                con.execute("INSERT INTO coordinator_master(emp_id,full_name,mobile,is_active,created_at,updated_at,hire_date) VALUES(?,?,?,1,?,?,?)",
                    (emp,name,mobile,now,now,hire_date))
                con.execute("INSERT INTO personnel_master(emp_id,full_name,id_no,mobile,role_title,linked_ministry_id,is_active,imported_at,hire_date) VALUES(?,?,?,?,?,?,1,?,?)",
                    (emp,name,id_no,mobile,"منسق / مشرف",mid,now,hire_date))
            con.execute("INSERT INTO master_import_log(file_name,imported_at,imported_by,schools_count,guards_count,coordinators_count,administrators_count,archived_schools_count,mode,notes) VALUES(?,?,?,?,?,?,?,?,?,?)",
                (file.filename,now,user["id"],len(matched),len(guards),len(supervisors),0,0,"استبدال تسكين الموظفين",f"Matched {len(matched)}/{len(incoming)} schools"))
            con.commit()
        except Exception:
            con.rollback()
            raise
        finally:
            con.close()
        return RedirectResponse("/school-import?success=1",303)
    con=db();now=datetime.utcnow().isoformat()
    # Preserve operational history; replace master assignments and personnel only.
    old_ids={r["ministry_id"] for r in con.execute("SELECT ministry_id FROM schools")};new_ids=set()
    con.execute("DELETE FROM personnel_master");con.execute("DELETE FROM coordinator_master");con.execute("DELETE FROM guard_master")
    counts={"guards":0,"coords":0,"admins":0}
    for row in rows:
        mid=_pick(row,"الرقم الوزاري","رقم وزاري","ministry_id");new_ids.add(mid)
        vals={
          "school_name":_pick(row,"اسم المدرسة","المدرسة","school_name"),"city":_pick(row,"المدينة","المنطقة","city"),
          "school_type":_pick(row,"رئيسية / فرعية","نوع المدرسة","school_type"),"gender":_pick(row,"النوع","بنين / بنات","gender"),
          "size":_pick(row,"حجم المدرسة","size"),"guard_name":_pick(row,"اسم البواب","اسم الحارس","guard_name"),
          "guard_emp_id":_pick(row,"الرقم الوظيفي للبواب","رقم موظف البواب","guard_emp_id"),"guard_id":_pick(row,"رقم الهوية للبواب","هوية البواب","guard_id"),
          "guard_mobile":_pick(row,"جوال البواب","رقم جوال البواب","guard_mobile"),"coordinator_name":_pick(row,"اسم المنسق / المشرف","اسم المنسق","اسم المشرف","coordinator_name"),
          "coordinator_emp_id":_pick(row,"الرقم الوظيفي للمنسق / المشرف","الرقم الوظيفي للمنسق","رقم موظف المنسق","coordinator_emp_id"),
          "coordinator_mobile":_pick(row,"جوال المنسق / المشرف","جوال المنسق","رقم جوال المنسق","coordinator_mobile"),
          "operational_type":_pick(row,"نوع المدرسة التشغيلي","رئيسية / فرعية / ملحقة","operational_type"),
          "status":_pick(row,"حالة المدرسة","status") or "تعمل","merged":_pick(row,"المدرسة المدمج معها","الرقم الوزاري للمدرسة المدمج معها","merged_into_ministry_id"),
          "sector":_pick(row,"القطاع","sector"),"admin_name":_pick(row,"اسم الإداري","الإداري","administrative_name"),
          "admin_emp":_pick(row,"الرقم الوظيفي للإداري","administrative_emp_id"),"admin_mobile":_pick(row,"جوال الإداري","administrative_mobile")}
        vals["operational_type"]=_normalize_operational_type(vals["operational_type"],vals["school_type"])
        if vals["status"] not in ("تعمل","مغلقة","مدمجة"): vals["status"]="تعمل"
        con.execute("""INSERT INTO schools(ministry_id,school_name,city,school_type,gender,size,guard_name,guard_emp_id,guard_id,guard_mobile,coordinator_name,coordinator_emp_id,coordinator_mobile,operational_type,status,merged_into_ministry_id,sector,administrative_name,administrative_emp_id,administrative_mobile)
          VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(ministry_id) DO UPDATE SET school_name=excluded.school_name,city=excluded.city,school_type=excluded.school_type,gender=excluded.gender,size=excluded.size,guard_name=excluded.guard_name,guard_emp_id=excluded.guard_emp_id,guard_id=excluded.guard_id,guard_mobile=excluded.guard_mobile,coordinator_name=excluded.coordinator_name,coordinator_emp_id=excluded.coordinator_emp_id,coordinator_mobile=excluded.coordinator_mobile,operational_type=excluded.operational_type,status=excluded.status,merged_into_ministry_id=excluded.merged_into_ministry_id,sector=excluded.sector,administrative_name=excluded.administrative_name,administrative_emp_id=excluded.administrative_emp_id,administrative_mobile=excluded.administrative_mobile""",
          (mid,vals["school_name"],vals["city"],vals["school_type"],vals["gender"],vals["size"],vals["guard_name"],vals["guard_emp_id"],vals["guard_id"],vals["guard_mobile"],vals["coordinator_name"],vals["coordinator_emp_id"],vals["coordinator_mobile"],vals["operational_type"],vals["status"],vals["merged"],vals["sector"],vals["admin_name"],vals["admin_emp"],vals["admin_mobile"]))
        for role,n,e,i,m in (("بواب",vals["guard_name"],vals["guard_emp_id"],vals["guard_id"],vals["guard_mobile"]),("منسق / مشرف",vals["coordinator_name"],vals["coordinator_emp_id"],"",vals["coordinator_mobile"]),("إداري",vals["admin_name"],vals["admin_emp"],"",vals["admin_mobile"])):
            if not n: continue
            emp=e or f"{role}-{n}"
            con.execute("INSERT OR REPLACE INTO personnel_master(emp_id,full_name,id_no,mobile,role_title,city,sector,linked_ministry_id,is_active,imported_at) VALUES(?,?,?,?,?,?,?,?,1,?)",(emp,n,i,m,role,vals["city"],vals["sector"],mid,now))
            if role=="بواب":
                con.execute("INSERT OR REPLACE INTO guard_master(emp_id,emp_name,id_no,mobile,status,linked_ministry_id) VALUES(?,?,?,?,?,?)",(emp,n,i,m,"نشط",mid));counts["guards"]+=1
            elif role=="منسق / مشرف":
                existing=con.execute("SELECT id FROM coordinator_master WHERE emp_id=? OR full_name=?",(e,n)).fetchone()
                if not existing: con.execute("INSERT INTO coordinator_master(emp_id,full_name,mobile,city,is_active,created_at,updated_at) VALUES(?,?,?,?,1,?,?)",(e,n,m,vals["city"],now,now));counts["coords"]+=1
            else: counts["admins"]+=1
    archived=old_ids-new_ids
    for mid in archived: con.execute("UPDATE schools SET status='مغلقة',status_notes='غير موجودة في آخر ملف Master Data',status_updated_at=? WHERE ministry_id=?",(now,mid))
    con.execute("INSERT INTO master_import_log(file_name,imported_at,imported_by,schools_count,guards_count,coordinators_count,administrators_count,archived_schools_count,mode) VALUES(?,?,?,?,?,?,?,?,?)",(file.filename,now,user["id"],len(new_ids),counts["guards"],counts["coords"],counts["admins"],len(archived),"استبدال كامل"))
    con.commit();con.close();return RedirectResponse("/school-import?success=1",303)

@app.get("/school-import/template.xlsx")
def school_import_template(request: Request):
    user=current_user(request)
    if not user: raise HTTPException(401)
    wb=Workbook();ws=wb.active;ws.title="بيانات المدارس"
    headers=[
      "الرقم الوزاري","اسم المدرسة","المدينة","رئيسية / فرعية","النوع","حجم المدرسة",
      "اسم البواب","الرقم الوظيفي للبواب","رقم الهوية للبواب","جوال البواب",
      "اسم المنسق / المشرف","الرقم الوظيفي للمنسق / المشرف","جوال المنسق / المشرف","القطاع",
      "اسم الإداري","الرقم الوظيفي للإداري","جوال الإداري","نوع المدرسة التشغيلي",
      "حالة المدرسة","الرقم الوزاري للمدرسة المدمج معها"
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font=cell.font.copy(bold=True,color="FFFFFF")
        cell.fill=cell.fill.copy(fill_type="solid",fgColor="082B60")
    bio=BytesIO();wb.save(bio);bio.seek(0)
    return StreamingResponse(
      bio,
      media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      headers={"Content-Disposition":'attachment; filename="school_import_template.xlsx"'}
    )


@app.get("/users", response_class=HTMLResponse)
def users_page(request: Request):
    user=require_permission(request,"users")
    con=db()
    users=con.execute("""
      SELECT id,username,full_name,role,coordinator_name,is_active,
             external_department,external_company,external_mobile,
             client_city,client_coordinator,client_name
      FROM users ORDER BY id
    """).fetchall()
    coordinators=[r["coordinator_name"] for r in con.execute("""
      SELECT DISTINCT coordinator_name FROM schools
      WHERE coordinator_name<>'' ORDER BY coordinator_name
    """)]
    cities=[r["city"] for r in con.execute("""
      SELECT DISTINCT city FROM schools WHERE COALESCE(city,'')<>'' ORDER BY city
    """)]
    permissions={}
    for u in users:
        rows=con.execute("SELECT permission_key,is_allowed FROM user_permissions WHERE user_id=?",(u["id"],)).fetchall()
        permissions[u["id"]]={r["permission_key"]:bool(r["is_allowed"]) for r in rows}
    con.close()
    return templates.TemplateResponse("users.html",{
      "request":request,"user":user,"users":users,"coordinators":coordinators,
      "permission_catalog":PERMISSION_CATALOG,"permissions":permissions,
      "cities":cities
    })

@app.post("/users/create")
def user_create(
    request: Request,
    username: str=Form(...),
    password: str=Form(...),
    full_name: str=Form(...),
    role: str=Form(...),
    coordinator_name: str=Form(""),
    external_department: str=Form(""),
    external_company: str=Form(""),
    external_mobile: str=Form(""),
    client_city: str=Form(""),
    client_coordinator: str=Form(""),
    client_name: str=Form("")
):
    admin=require_permission(request,"users")
    con=db()
    try:
        cur=con.execute("""
          INSERT INTO users(
            username,password_hash,full_name,role,coordinator_name,is_active,
            external_department,external_company,external_mobile,
            client_city,client_coordinator,client_name
          ) VALUES(?,?,?,?,?,1,?,?,?,?,?,?)
        """,(
          username.strip(),hash_password(password),full_name.strip(),role,
          coordinator_name.strip() or None,external_department.strip() or None,
          external_company.strip() or None,external_mobile.strip() or None,
          client_city.strip() or None,client_coordinator.strip() or None,
          client_name.strip() or None
        ))
        uid=cur.lastrowid
        for key in ROLE_DEFAULTS.get(role,[]):
            con.execute("INSERT OR IGNORE INTO user_permissions(user_id,permission_key,is_allowed) VALUES(?,?,1)",(uid,key))
        con.commit()
    except IntegrityError:
        con.close()
        return RedirectResponse("/users?error=username_exists",status_code=303)
    con.close()
    return RedirectResponse("/users?created=1",status_code=303)

@app.post("/users/{user_id}/update")
def user_update(
    request: Request,
    user_id: int,
    full_name: str=Form(...),
    role: str=Form(...),
    coordinator_name: str=Form(""),
    external_department: str=Form(""),
    external_company: str=Form(""),
    external_mobile: str=Form(""),
    client_city: str=Form(""),
    client_coordinator: str=Form(""),
    client_name: str=Form(""),
    is_active: int=Form(1)
):
    admin=require_permission(request,"users")
    if user_id==admin["id"] and not is_active:
        raise HTTPException(400,detail="لا يمكن تعطيل حسابك الحالي")
    con=db()
    con.execute("""
      UPDATE users SET full_name=?,role=?,coordinator_name=?,
        external_department=?,external_company=?,external_mobile=?,
        client_city=?,client_coordinator=?,client_name=?,is_active=?
        WHERE id=?
    """,(
      full_name.strip(),role,coordinator_name.strip() or None,
      external_department.strip() or None,external_company.strip() or None,
      external_mobile.strip() or None,client_city.strip() or None,
      client_coordinator.strip() or None,client_name.strip() or None,is_active,user_id
    ))
    con.commit();con.close()
    return RedirectResponse("/users",status_code=303)

@app.post("/users/{user_id}/password")
def user_password(
    request: Request,
    user_id: int,
    new_password: str=Form(...)
):
    require_permission(request,"users")
    con=db()
    con.execute("UPDATE users SET password_hash=? WHERE id=?",(hash_password(new_password),user_id))
    con.commit();con.close()
    return RedirectResponse("/users?password=1",status_code=303)

@app.post("/users/{user_id}/permissions")
async def user_permissions_save(request: Request,user_id: int):
    require_permission(request,"users")
    form=await request.form()
    con=db()
    con.execute("DELETE FROM user_permissions WHERE user_id=?",(user_id,))
    for key in PERMISSION_CATALOG:
        allowed=1 if form.get(f"perm_{key}")=="on" else 0
        con.execute("""
          INSERT INTO user_permissions(user_id,permission_key,is_allowed) VALUES(?,?,?)
        """,(user_id,key,allowed))
    con.commit();con.close()
    return RedirectResponse("/users?permissions=1",status_code=303)

@app.post("/users/{user_id}/toggle")
def user_toggle(request: Request,user_id: int):
    admin=require_permission(request,"users")
    if user_id==admin["id"]:
        raise HTTPException(400,detail="لا يمكن تعطيل حسابك الحالي")
    con=db()
    row=con.execute("SELECT is_active FROM users WHERE id=?",(user_id,)).fetchone()
    if not row:
        con.close();raise HTTPException(404)
    con.execute("UPDATE users SET is_active=? WHERE id=?",(0 if row["is_active"] else 1,user_id))
    con.commit();con.close()
    return RedirectResponse("/users",status_code=303)

@app.post("/users/{user_id}/delete")
def user_delete(request: Request,user_id: int):
    admin=require_permission(request,"users")
    if user_id==admin["id"]:
        raise HTTPException(400,detail="لا يمكن حذف حسابك الحالي")
    con=db()
    con.execute("DELETE FROM user_permissions WHERE user_id=?",(user_id,))
    con.execute("DELETE FROM users WHERE id=?",(user_id,))
    con.commit();con.close()
    return RedirectResponse("/users?deleted=1",status_code=303)


@app.get("/access-permits", response_class=HTMLResponse)
def access_permits_page(
    request: Request,
    date_from: str="",
    date_to: str="",
    city: str="",
    visit_type: str="",
    approval_status: str="",
    actual_status: str=""
):
    user=require_permission(request,"access_permits")
    con=db()
    where=["1=1"];args=[]
    if user["role"]=="coordinator":
        where.append("coordinator_name=?");args.append(user["coordinator_name"])
    elif user["role"]=="external_requester":
        where.append("requester_user_id=?");args.append(user["id"])
    if date_from: where.append("permit_date>=?");args.append(date_from)
    if date_to: where.append("permit_date<=?");args.append(date_to)
    if city: where.append("city=?");args.append(city)
    if visit_type: where.append("visit_type=?");args.append(visit_type)
    if approval_status: where.append("approval_status=?");args.append(approval_status)
    if actual_status: where.append("actual_status=?");args.append(actual_status)

    rows=con.execute(
      "SELECT * FROM school_access_permits WHERE "+" AND ".join(where)+
      " ORDER BY permit_date DESC,planned_entry_time DESC,id DESC",args
    ).fetchall()
    schools=con.execute("""
      SELECT ministry_id,school_name,city,coordinator_name,guard_name,guard_emp_id
      FROM schools ORDER BY city,school_name
    """).fetchall()
    cities=[r["city"] for r in con.execute("SELECT DISTINCT city FROM schools WHERE city<>'' ORDER BY city")]
    stats=con.execute("""
      SELECT COUNT(*) total,
       SUM(CASE WHEN approval_status='بانتظار الموافقة' THEN 1 ELSE 0 END) pending,
       SUM(CASE WHEN approval_status='موافق' THEN 1 ELSE 0 END) approved,
       SUM(CASE WHEN actual_status='داخل المدرسة' THEN 1 ELSE 0 END) inside_count,
       SUM(CASE WHEN actual_status='تم الخروج' THEN 1 ELSE 0 END) completed
      FROM school_access_permits WHERE """+" AND ".join(where),args).fetchone()
    con.close()
    return templates.TemplateResponse("access_permits.html",{
      "request":request,"user":user,"rows":rows,"schools":schools,"cities":cities,
      "stats":stats,"date_from":date_from,"date_to":date_to,"city":city,
      "visit_type":visit_type,"approval_status":approval_status,
      "actual_status":actual_status,"is_external":user["role"]=="external_requester"
    })

@app.post("/access-permits/create")
def access_permit_create(
    request: Request,
    permit_date: str=Form(...),
    ministry_id: str=Form(...),
    visit_type: str=Form(...),
    company_name: str=Form(""),
    visitor_names: str=Form(...),
    visitor_ids: str=Form(""),
    visitor_mobiles: str=Form(""),
    vehicle_plate: str=Form(""),
    work_order_no: str=Form(""),
    visit_reason: str=Form(""),
    planned_entry_time: str=Form(""),
    planned_exit_time: str=Form(""),
    notes: str=Form("")
):
    user=require_permission(request,"access_permits")
    con=db()
    s=con.execute("SELECT * FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
    if not s: raise HTTPException(404)
    if user["role"]=="coordinator" and s["coordinator_name"]!=user["coordinator_name"]:
        raise HTTPException(403)
    con.execute("""
      INSERT INTO school_access_permits(
        permit_date,ministry_id,school_name,city,coordinator_name,visit_type,
        company_name,visitor_names,visitor_ids,visitor_mobiles,vehicle_plate,
        work_order_no,visit_reason,planned_entry_time,planned_exit_time,notes,
        created_by,created_at,requester_user_id,requester_name,
        requester_department,requester_company,requester_mobile
      ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """,(
      permit_date,ministry_id,s["school_name"],s["city"],s["coordinator_name"],
      visit_type,company_name,visitor_names,visitor_ids,visitor_mobiles,
      vehicle_plate,work_order_no,visit_reason,planned_entry_time,
      planned_exit_time,notes,user["id"],datetime.utcnow().isoformat(),
      user["id"],user["full_name"],
      user["external_department"] if "external_department" in user.keys() else "",
      user["external_company"] if "external_company" in user.keys() else company_name,
      user["external_mobile"] if "external_mobile" in user.keys() else visitor_mobiles
    ))
    permit_id=con.execute("SELECT last_insert_rowid() id").fetchone()["id"]
    ensure_permit_tokens(con,permit_id)
    con.commit();con.close()
    return RedirectResponse("/access-permits?created=1",status_code=303)

@app.post("/access-permits/{permit_id}/approve")
def access_permit_approve(
    request: Request,
    permit_id: int,
    decision: str=Form(...),
    notes: str=Form("")
):
    user=require_permission(request,"access_permits")
    if decision not in ("موافق","مرفوض"):
        raise HTTPException(400)
    con=db()
    permit=con.execute("SELECT * FROM school_access_permits WHERE id=?",(permit_id,)).fetchone()
    if not permit: raise HTTPException(404)
    if user["role"]=="coordinator" and permit["coordinator_name"]!=user["coordinator_name"]:
        raise HTTPException(403)
    con.execute("""
      UPDATE school_access_permits SET approval_status=?,approved_by_name=?,
      approved_by_role=?,approved_at=?,notes=CASE WHEN ?<>'' THEN ? ELSE notes END
      WHERE id=?
    """,(decision,user["full_name"],user["role"],datetime.utcnow().isoformat(),notes,notes,permit_id))
    con.commit();con.close()
    return RedirectResponse("/access-permits",status_code=303)

@app.post("/access-permits/{permit_id}/entry")
def access_permit_entry(
    request: Request,
    permit_id: int,
    entry_latitude: str=Form(""),
    entry_longitude: str=Form(""),
    school_opened_by: str=Form(...),
    visitor_confirmation: str=Form(...),
    entry_photo: UploadFile=File(...),
    notes: str=Form("")
):
    user=require_permission(request,"access_permits")
    con=db()
    permit=con.execute("SELECT * FROM school_access_permits WHERE id=?",(permit_id,)).fetchone()
    if not permit:
        con.close();raise HTTPException(404)
    if permit["approval_status"]!="موافق":
        con.close();raise HTTPException(400,detail="لا يمكن تسجيل الدخول قبل الموافقة")
    if permit["actual_status"]!="لم يدخل":
        con.close();raise HTTPException(400,detail="تم تسجيل الدخول مسبقًا")
    photo_path,evidence_hash=save_access_evidence(entry_photo,permit_id,"entry")
    now=datetime.utcnow().isoformat()
    local_time=datetime.now().strftime("%H:%M:%S")
    evidence_no=Path(photo_path).stem
    con.execute("""
      UPDATE school_access_permits SET entry_time=?,entry_server_time=?,
      entry_latitude=?,entry_longitude=?,school_opened_by=?,
      entry_verified_by=?,entry_photo_path=?,entry_evidence_no=?,
      entry_evidence_hash=?,actual_status='داخل المدرسة',
      notes=CASE WHEN ?<>'' THEN COALESCE(notes,'') || CHAR(10) || ? ELSE notes END
      WHERE id=?
    """,(local_time,now,float(entry_latitude or 0),float(entry_longitude or 0),
         school_opened_by,f'{user["full_name"]} | إقرار الزائر: {visitor_confirmation}',
         photo_path,evidence_no,evidence_hash,notes,notes,permit_id))
    con.commit();con.close()
    return RedirectResponse("/access-permits?entry_verified=1",status_code=303)

@app.post("/access-permits/{permit_id}/exit")
def access_permit_exit(
    request: Request,
    permit_id: int,
    exit_latitude: str=Form(""),
    exit_longitude: str=Form(""),
    school_closed_by: str=Form(...),
    visitor_confirmation: str=Form(...),
    exit_photo: UploadFile=File(...),
    notes: str=Form("")
):
    user=require_permission(request,"access_permits")
    con=db()
    permit=con.execute("SELECT * FROM school_access_permits WHERE id=?",(permit_id,)).fetchone()
    if not permit:
        con.close();raise HTTPException(404)
    if permit["actual_status"]!="داخل المدرسة":
        con.close();raise HTTPException(400,detail="لا يمكن تسجيل الخروج قبل تسجيل الدخول")
    photo_path,evidence_hash=save_access_evidence(exit_photo,permit_id,"exit")
    now=datetime.utcnow().isoformat()
    local_time=datetime.now().strftime("%H:%M:%S")
    evidence_no=Path(photo_path).stem
    con.execute("""
      UPDATE school_access_permits SET exit_time=?,exit_server_time=?,
      exit_latitude=?,exit_longitude=?,school_closed_by=?,
      exit_verified_by=?,exit_photo_path=?,exit_evidence_no=?,
      exit_evidence_hash=?,actual_status='تم الخروج',
      notes=CASE WHEN ?<>'' THEN COALESCE(notes,'') || CHAR(10) || ? ELSE notes END
      WHERE id=?
    """,(local_time,now,float(exit_latitude or 0),float(exit_longitude or 0),
         school_closed_by,f'{user["full_name"]} | إقرار الزائر: {visitor_confirmation}',
         photo_path,evidence_no,evidence_hash,notes,notes,permit_id))
    con.commit();con.close()
    return RedirectResponse("/access-permits?exit_verified=1",status_code=303)

@app.get("/access-permits.xlsx")
def access_permits_export(
    request: Request,
    date_from: str="",
    date_to: str="",
    city: str="",
    visit_type: str=""
):
    user=require_permission(request,"access_permits")
    con=db();where=["1=1"];args=[]
    if user["role"]=="coordinator":
        where.append("coordinator_name=?");args.append(user["coordinator_name"])
    if date_from: where.append("permit_date>=?");args.append(date_from)
    if date_to: where.append("permit_date<=?");args.append(date_to)
    if city: where.append("city=?");args.append(city)
    if visit_type: where.append("visit_type=?");args.append(visit_type)
    rows=con.execute(
      "SELECT * FROM school_access_permits WHERE "+" AND ".join(where)+
      " ORDER BY permit_date,city,school_name,entry_time",args
    ).fetchall()
    con.close()

    wb=Workbook();ws=wb.active;ws.title="دخول المدارس"
    headers=[
      "التاريخ","المدينة","المدرسة","الرقم الوزاري","المشرف","نوع الزيارة",
      "الشركة","أسماء الزوار","الهويات","الجوالات","المركبة","أمر العمل",
      "سبب الزيارة","حالة الموافقة","اعتمد بواسطة","وقت فتح المدرسة",
      "فتح المدرسة بواسطة","GPS الدخول","وقت الخروج","إغلاق المدرسة بواسطة",
      "GPS الخروج","الحالة الفعلية","رقم إثبات الدخول","صورة الدخول",
      "اعتماد الدخول","توقيت السيرفر للدخول","رقم إثبات الخروج","صورة الخروج",
      "اعتماد الخروج","توقيت السيرفر للخروج","الملاحظات"
    ]
    ws.append(headers)
    for r in rows:
        ws.append([
          r["permit_date"],r["city"],r["school_name"],r["ministry_id"],
          r["coordinator_name"],r["visit_type"],r["company_name"],
          r["visitor_names"],r["visitor_ids"],r["visitor_mobiles"],
          r["vehicle_plate"],r["work_order_no"],r["visit_reason"],
          r["approval_status"],r["approved_by_name"],r["entry_time"],
          r["school_opened_by"],
          f'{r["entry_latitude"] or ""},{r["entry_longitude"] or ""}',
          r["exit_time"],r["school_closed_by"],
          f'{r["exit_latitude"] or ""},{r["exit_longitude"] or ""}',
          r["actual_status"],r["entry_evidence_no"],
          f'/evidence/{r["entry_photo_path"]}' if r["entry_photo_path"] else "",
          r["entry_verified_by"],r["entry_server_time"],r["exit_evidence_no"],
          f'/evidence/{r["exit_photo_path"]}' if r["exit_photo_path"] else "",
          r["exit_verified_by"],r["exit_server_time"],r["notes"]
        ])
    for c in ws[1]:
        c.font=c.font.copy(bold=True,color="FFFFFF")
        c.fill=c.fill.copy(fill_type="solid",fgColor="082B60")
    ws.freeze_panes="A2"
    bio=BytesIO();wb.save(bio);bio.seek(0)
    return StreamingResponse(
      bio,
      media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      headers={"Content-Disposition":'attachment; filename="school_access_report.xlsx"'}
    )



@app.get("/data-quality", response_class=HTMLResponse)
def data_quality_page(request: Request):
    user=require_permission(request,"data_quality")
    con=db()
    stats=con.execute("""SELECT
      COUNT(*) total_schools,
      SUM(CASE WHEN COALESCE(TRIM(guard_emp_id),'')='' OR COALESCE(TRIM(guard_name),'')='' THEN 1 ELSE 0 END) no_guard,
      SUM(CASE WHEN COALESCE(TRIM(coordinator_name),'')='' OR TRIM(coordinator_name)='غير متوفر' THEN 1 ELSE 0 END) no_coordinator,
      SUM(CASE WHEN latitude IS NULL OR longitude IS NULL THEN 1 ELSE 0 END) no_coordinates
      FROM schools""").fetchone()
    unlinked=con.execute("""SELECT COUNT(*) c FROM guard_master
      WHERE COALESCE(service_status,'على رأس العمل')<>'منتهي الخدمة'
      AND COALESCE(TRIM(linked_ministry_id),'')=''""").fetchone()["c"]
    duplicate_guards=con.execute("""SELECT COUNT(*) c FROM (
      SELECT guard_emp_id FROM schools WHERE COALESCE(TRIM(guard_emp_id),'')<>''
      GROUP BY guard_emp_id HAVING COUNT(*)>1
    )""").fetchone()["c"]
    no_coord_rows=con.execute("""SELECT ministry_id,school_name,city,guard_name,guard_emp_id
      FROM schools WHERE COALESCE(TRIM(coordinator_name),'')='' OR TRIM(coordinator_name)='غير متوفر'
      ORDER BY city,school_name LIMIT 200""").fetchall()
    no_coord_count=con.execute("""SELECT COUNT(*) c FROM schools
      WHERE COALESCE(TRIM(coordinator_name),'')='' OR TRIM(coordinator_name)='غير متوفر'""").fetchone()["c"]
    no_coords_rows=con.execute("""SELECT ministry_id,school_name,city,coordinator_name,guard_name
      FROM schools WHERE latitude IS NULL OR longitude IS NULL
      ORDER BY city,school_name LIMIT 200""").fetchall()
    unlinked_rows=con.execute("""SELECT emp_id,emp_name,id_no,mobile,service_status
      FROM guard_master WHERE COALESCE(service_status,'على رأس العمل')<>'منتهي الخدمة'
      AND COALESCE(TRIM(linked_ministry_id),'')=''
      ORDER BY emp_name LIMIT 200""").fetchall()
    con.close()
    return templates.TemplateResponse("data_quality.html",{
      "request":request,"user":user,"stats":stats,"unlinked":unlinked,
      "duplicate_guards":duplicate_guards,"no_coord_rows":no_coord_rows,
      "no_coord_count":no_coord_count,"no_coords_rows":no_coords_rows,
      "unlinked_rows":unlinked_rows
    })

@app.get("/coordinates", response_class=HTMLResponse)
def coordinates_page(request: Request,q: str="",city: str="",status: str="missing"):
    user=require_permission(request,"coordinates")
    con=db();where=["1=1"];args=[]
    if status=="missing":
        where.append("(latitude IS NULL OR longitude IS NULL)")
    elif status=="available":
        where.append("(latitude IS NOT NULL AND longitude IS NOT NULL)")
    if q:
        where.append("(school_name LIKE ? OR ministry_id LIKE ?)")
        args.extend([f"%{q}%",f"%{q}%"])
    if city:
        where.append("city=?");args.append(city)
    rows=con.execute("SELECT * FROM schools WHERE "+" AND ".join(where)+" ORDER BY city,school_name LIMIT 1000",args).fetchall()
    cities=[r["city"] for r in con.execute("SELECT DISTINCT city FROM schools WHERE COALESCE(city,'')<>'' ORDER BY city")]
    counts=con.execute("""SELECT COUNT(*) total,
      SUM(CASE WHEN latitude IS NULL OR longitude IS NULL THEN 1 ELSE 0 END) missing,
      SUM(CASE WHEN latitude IS NOT NULL AND longitude IS NOT NULL THEN 1 ELSE 0 END) available
      FROM schools""").fetchone()
    con.close()
    return templates.TemplateResponse("coordinates.html",{
      "request":request,"user":user,"rows":rows,"cities":cities,
      "q":q,"city":city,"status":status,"counts":counts
    })

@app.post("/coordinates/{ministry_id}/update")
def coordinate_update(
    request: Request,ministry_id: str,
    latitude: str=Form(...),longitude: str=Form(...)
):
    user=require_permission(request,"coordinates")
    try:
        lat=float(str(latitude).replace(",",".").strip())
        lng=float(str(longitude).replace(",",".").strip())
    except ValueError:
        raise HTTPException(400,detail="الإحداثيات غير صحيحة")
    if not (-90<=lat<=90 and -180<=lng<=180):
        raise HTTPException(400,detail="الإحداثيات خارج النطاق")
    con=db()
    con.execute("UPDATE schools SET latitude=?,longitude=? WHERE ministry_id=?",(lat,lng,ministry_id))
    con.commit();con.close()
    return RedirectResponse("/coordinates?updated=1",status_code=303)

@app.post("/coordinates/import")
async def coordinates_import(request: Request,file: UploadFile=File(...)):
    user=require_permission(request,"coordinates")
    content=await file.read()
    wb=load_workbook(BytesIO(content),data_only=True,read_only=True)
    updated=skipped=0
    con=db()
    for ws in wb.worksheets:
        headers=[str(c.value or "").strip() for c in ws[1]]
        for values in ws.iter_rows(min_row=2,values_only=True):
            row=dict(zip(headers,values))
            ministry_id=_pick(row,"الرقم الوزاري","رقم وزاري","ministry_id")
            lat=_pick(row,"Latitude","latitude","خط العرض")
            lng=_pick(row,"Longitude","longitude","خط الطول")
            if not ministry_id or not lat or not lng:
                skipped+=1;continue
            try:
                lat_v=float(str(lat).replace(",","."))
                lng_v=float(str(lng).replace(",","."))
                if not (-90<=lat_v<=90 and -180<=lng_v<=180):
                    raise ValueError
            except Exception:
                skipped+=1;continue
            cur=con.execute("UPDATE schools SET latitude=?,longitude=? WHERE ministry_id=?",
                            (lat_v,lng_v,ministry_id))
            if cur.rowcount:updated+=1
            else:skipped+=1
    con.commit();con.close();wb.close()
    return RedirectResponse(f"/coordinates?imported={updated}&skipped={skipped}",status_code=303)

@app.get("/downloads/v15-review")
def download_v15_review(request: Request):
    require_permission(request,"data_quality")
    return FileResponse(STATIC_DIR/"V15_تقرير_مراجعة_البيانات.xlsx",
      filename="V15_تقرير_مراجعة_البيانات.xlsx")

@app.get("/downloads/v15-import")
def download_v15_import(request: Request):
    require_permission(request,"school_import")
    return FileResponse(STATIC_DIR/"V15_تحديث_المدارس_والبوابين.xlsx",
      filename="V15_تحديث_المدارس_والبوابين.xlsx")

@app.get("/system-settings", response_class=HTMLResponse)
def system_settings_page(request: Request):
    user=require_permission(request,"system_settings")
    return templates.TemplateResponse("system_settings.html",{
      "request":request,"user":user,"settings":get_app_settings()
    })

@app.post("/system-settings")
async def system_settings_save(request: Request):
    user=require_permission(request,"system_settings")
    form=await request.form()
    allowed=[
      "company_name","system_name","project_name","system_version",
      "gate_services_manager","gate_services_manager_mobile",
      "general_supervisor","general_supervisor_mobile",
      "security_sector_manager","security_sector_manager_mobile",
      "prepared_by","prepared_by_mobile"
    ]
    con=db()
    for key in allowed:
        if key in form:
            con.execute("""
              INSERT INTO app_settings(setting_key,setting_value,updated_by,updated_at)
              VALUES(?,?,?,?)
              ON CONFLICT(setting_key) DO UPDATE SET
                setting_value=excluded.setting_value,
                updated_by=excluded.updated_by,
                updated_at=excluded.updated_at
            """,(key,str(form.get(key,"")).strip(),user["id"],datetime.utcnow().isoformat()))
    con.commit();con.close()
    return RedirectResponse("/system-settings?saved=1",status_code=303)

@app.get("/access-permits/{permit_id}/whatsapp-approval")
def access_permit_whatsapp_approval(request: Request,permit_id: int):
    user=require_permission(request,"access_permits")
    con=db()
    permit=con.execute("""
      SELECT p.*,s.coordinator_mobile FROM school_access_permits p
      LEFT JOIN schools s ON s.ministry_id=p.ministry_id
      WHERE p.id=?
    """,(permit_id,)).fetchone()
    if not permit:
        con.close();raise HTTPException(404)
    approval_token,_=ensure_permit_tokens(con,permit_id)
    con.execute("UPDATE school_access_permits SET whatsapp_approval_at=? WHERE id=?",(datetime.utcnow().isoformat(),permit_id))
    con.commit();con.close()
    number=normalize_whatsapp_number(permit["coordinator_mobile"])
    if not number:
        raise HTTPException(400,detail="رقم واتساب المشرف غير موجود")
    approval_url=str(request.base_url).rstrip("/") + f"/permit/approve/{approval_token}"
    message=(
      f"طلب تصريح دخول {permit['visit_type']}\n"
      f"المدرسة: {permit['school_name']}\n"
      f"المدينة: {permit['city']}\n"
      f"الشركة: {permit['company_name'] or 'غير محدد'}\n"
      f"الزوار: {permit['visitor_names']}\n"
      f"التاريخ: {permit['permit_date']} - الوقت: {permit['planned_entry_time'] or 'غير محدد'}\n"
      f"سبب الزيارة: {permit['visit_reason'] or 'غير محدد'}\n"
      f"للموافقة أو الرفض افتح الرابط:\n{approval_url}"
    )
    return RedirectResponse(f"https://wa.me/{number}?text={quote(message)}",status_code=302)

@app.get("/permit/approve/{token}", response_class=HTMLResponse)
def public_permit_approval_page(request: Request,token: str):
    con=db()
    permit=con.execute("SELECT * FROM school_access_permits WHERE approval_token=?",(token,)).fetchone()
    con.close()
    if not permit: raise HTTPException(404)
    return templates.TemplateResponse("public_permit_approval.html",{"request":request,"permit":permit,"token":token})

@app.post("/permit/approve/{token}")
def public_permit_approval_submit(
    request: Request,token: str,decision: str=Form(...),
    approver_name: str=Form(...),approver_mobile: str=Form(""),notes: str=Form("")
):
    if decision not in ("موافق","مرفوض"): raise HTTPException(400)
    con=db()
    permit=con.execute("SELECT * FROM school_access_permits WHERE approval_token=?",(token,)).fetchone()
    if not permit: con.close();raise HTTPException(404)
    if permit["approval_status"]!="بانتظار الموافقة":
        con.close()
        return templates.TemplateResponse("public_result.html",{"request":request,"message":"تم اتخاذ القرار على هذا التصريح مسبقًا."})
    con.execute("""
      UPDATE school_access_permits SET approval_status=?,approved_by_name=?,
      approved_by_role='موافقة عبر واتساب',approved_at=?,
      notes=CASE WHEN ?<>'' THEN COALESCE(notes,'')||CHAR(10)||? ELSE notes END
      WHERE id=?
    """,(decision,f"{approver_name} | {approver_mobile}",datetime.utcnow().isoformat(),notes,notes,permit["id"]))
    con.commit();con.close()
    return templates.TemplateResponse("public_result.html",{"request":request,"message":f"تم تسجيل القرار: {decision}"})

@app.get("/access-permits/{permit_id}/whatsapp-exit")
def access_permit_whatsapp_exit(request: Request,permit_id: int):
    user=require_permission(request,"access_permits")
    con=db()
    permit=con.execute("SELECT * FROM school_access_permits WHERE id=?",(permit_id,)).fetchone()
    if not permit: con.close();raise HTTPException(404)
    _,exit_token=ensure_permit_tokens(con,permit_id)
    con.execute("UPDATE school_access_permits SET whatsapp_exit_at=? WHERE id=?",(datetime.utcnow().isoformat(),permit_id))
    con.commit();con.close()
    number=normalize_whatsapp_number((permit["visitor_mobiles"] or "").split(",")[0].split("،")[0])
    if not number: raise HTTPException(400,detail="رقم واتساب فريق الصيانة أو النظافة غير موجود")
    exit_url=str(request.base_url).rstrip("/") + f"/permit/exit/{exit_token}"
    message=(
      f"إثبات انتهاء العمل والخروج من المدرسة\n"
      f"المدرسة: {permit['school_name']}\n"
      f"أمر العمل: {permit['work_order_no'] or 'غير محدد'}\n"
      f"بعد انتهاء العمل سجّل صورة الخروج والموقع من الرابط:\n{exit_url}"
    )
    return RedirectResponse(f"https://wa.me/{number}?text={quote(message)}",status_code=302)

@app.get("/permit/exit/{token}", response_class=HTMLResponse)
def public_permit_exit_page(request: Request,token: str):
    con=db()
    permit=con.execute("SELECT * FROM school_access_permits WHERE exit_token=?",(token,)).fetchone()
    con.close()
    if not permit: raise HTTPException(404)
    return templates.TemplateResponse("public_permit_exit.html",{"request":request,"permit":permit,"token":token})

@app.post("/permit/exit/{token}")
def public_permit_exit_submit(
    request: Request,token: str,
    exit_latitude: str=Form(""),exit_longitude: str=Form(""),
    representative_name: str=Form(...),school_closed_by: str=Form(...),
    exit_photo: UploadFile=File(...),notes: str=Form("")
):
    con=db()
    permit=con.execute("SELECT * FROM school_access_permits WHERE exit_token=?",(token,)).fetchone()
    if not permit: con.close();raise HTTPException(404)
    if permit["actual_status"]!="داخل المدرسة":
        con.close()
        return templates.TemplateResponse("public_result.html",{"request":request,"message":"لا يمكن تسجيل الخروج قبل وجود دخول مثبت، أو تم تسجيل الخروج مسبقًا."})
    photo_path,evidence_hash=save_access_evidence(exit_photo,permit["id"],"exit")
    now=datetime.utcnow().isoformat()
    local_time=datetime.now().strftime("%H:%M:%S")
    evidence_no=Path(photo_path).stem
    con.execute("""
      UPDATE school_access_permits SET exit_time=?,exit_server_time=?,
      exit_latitude=?,exit_longitude=?,school_closed_by=?,
      exit_verified_by=?,exit_photo_path=?,exit_evidence_no=?,
      exit_evidence_hash=?,actual_status='تم الخروج',
      notes=CASE WHEN ?<>'' THEN COALESCE(notes,'')||CHAR(10)||? ELSE notes END
      WHERE id=?
    """,(local_time,now,float(exit_latitude or 0),float(exit_longitude or 0),
         school_closed_by,f"{representative_name} | إثبات خروج عبر الرابط",
         photo_path,evidence_no,evidence_hash,notes,notes,permit["id"]))
    con.commit();con.close()
    return templates.TemplateResponse("public_result.html",{"request":request,"message":f"تم إثبات الخروج بنجاح. رقم الإثبات: {evidence_no}"})


@app.get("/unlinked-guards", response_class=HTMLResponse)
def unlinked_guards_page(request: Request, q: str="", view: str="active"):
    user=require_permission(request,"schools")
    con=db()
    where=["1=1"];args=[]
    if view=="terminated":
        where.append("COALESCE(service_status,'على رأس العمل')='منتهي الخدمة'")
    elif view=="all":
        pass
    else:
        where.append("COALESCE(service_status,'على رأس العمل')<>'منتهي الخدمة'")
        where.append("(linked_ministry_id IS NULL OR TRIM(linked_ministry_id)='')")
    if q:
        where.append("(emp_name LIKE ? OR emp_id LIKE ? OR id_no LIKE ? OR mobile LIKE ?)")
        like=f"%{q}%";args.extend([like,like,like,like])
    guards=con.execute(
      "SELECT * FROM guard_master WHERE "+" AND ".join(where)+" ORDER BY emp_name",args
    ).fetchall()
    empty_schools=con.execute("""
      SELECT s.ministry_id,s.school_name,s.city,s.coordinator_name,
             s.guard_name,s.guard_emp_id
      FROM schools s
      LEFT JOIN school_operational_exclusions e
        ON e.ministry_id=s.ministry_id AND e.is_excluded=1
      WHERE e.ministry_id IS NULL
        AND (COALESCE(TRIM(s.guard_name),'')='' OR COALESCE(TRIM(s.guard_emp_id),'')='')
      ORDER BY s.city,s.school_name
    """).fetchall()
    all_schools=con.execute("""
      SELECT ministry_id,school_name,city,guard_name,guard_emp_id,coordinator_name
      FROM schools ORDER BY city,school_name
    """).fetchall()
    active_unlinked=con.execute("""SELECT COUNT(*) c FROM guard_master
      WHERE COALESCE(service_status,'على رأس العمل')<>'منتهي الخدمة'
      AND (linked_ministry_id IS NULL OR TRIM(linked_ministry_id)='')""").fetchone()["c"]
    terminated_count=con.execute("""SELECT COUNT(*) c FROM guard_master
      WHERE COALESCE(service_status,'على رأس العمل')='منتهي الخدمة'""").fetchone()["c"]
    con.close()
    return templates.TemplateResponse("unlinked_guards.html",{
      "request":request,"user":user,"guards":guards,"schools":empty_schools,
      "all_schools":all_schools,"q":q,"view":view,
      "school_count_without_guard":len(empty_schools),
      "active_unlinked":active_unlinked,"terminated_count":terminated_count
    })

@app.post("/unlinked-guards/{emp_id}/link")
def link_guard_to_school(request: Request,emp_id: str,ministry_id: str=Form(...)):
    user=require_permission(request,"schools")
    con=db()
    guard=con.execute("SELECT * FROM guard_master WHERE emp_id=?",(emp_id,)).fetchone()
    school=con.execute("SELECT * FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
    if not guard or not school:
        con.close();raise HTTPException(404)
    if (guard["service_status"] or "على رأس العمل")=="منتهي الخدمة":
        con.close();raise HTTPException(400,detail="لا يمكن ربط بواب منتهي الخدمة")
    if guard["linked_ministry_id"]:
        con.close();raise HTTPException(400,detail="هذا البواب مرتبط بمدرسة بالفعل")
    if school["guard_emp_id"]:
        con.close();raise HTTPException(400,detail="المدرسة مرتبطة ببواب بالفعل")
    con.execute("""UPDATE schools SET guard_name=?,guard_emp_id=?,guard_id=?,guard_mobile=?
      WHERE ministry_id=?""",(guard["emp_name"],guard["emp_id"],guard["id_no"],guard["mobile"],ministry_id))
    con.execute("""UPDATE guard_master SET linked_ministry_id=?,linked_at=?,linked_by=?,
      updated_at=?,updated_by=? WHERE emp_id=?""",(
      ministry_id,datetime.utcnow().isoformat(),user["id"],datetime.utcnow().isoformat(),user["id"],emp_id))
    con.execute("""INSERT INTO guard_assignment_log(
      emp_id,emp_name,new_ministry_id,new_school_name,action_type,changed_by,changed_at
    ) VALUES(?,?,?,?,?,?,?)""",(emp_id,guard["emp_name"],ministry_id,school["school_name"],
      "ربط بمدرسة",user["id"],datetime.utcnow().isoformat()))
    con.commit();con.close()
    return RedirectResponse("/unlinked-guards?linked=1",status_code=303)

@app.post("/guards/create")
def guard_create(
    request: Request,
    emp_id: str=Form(...),
    emp_name: str=Form(...),
    id_no: str=Form(""),
    mobile: str=Form(""),
    ministry_id: str=Form("")
):
    user=require_permission(request,"schools")
    con=db()
    try:
        con.execute("""INSERT INTO guard_master(
          emp_id,emp_name,id_no,mobile,status,service_status,updated_at,updated_by
        ) VALUES(?,?,?,?,?,'على رأس العمل',?,?)""",(
          emp_id.strip(),emp_name.strip(),id_no.strip(),mobile.strip(),
          "نشط",datetime.utcnow().isoformat(),user["id"]))
    except IntegrityError:
        con.close();raise HTTPException(400,detail="الرقم الوظيفي موجود مسبقًا")
    con.commit();con.close()
    if ministry_id.strip():
        return link_guard_to_school(request,emp_id.strip(),ministry_id.strip())
    return RedirectResponse("/unlinked-guards?created=1",status_code=303)

@app.post("/guards/{emp_id}/terminate")
def guard_terminate(
    request: Request,
    emp_id: str,
    termination_date: str=Form(""),
    termination_reason: str=Form("")
):
    user=require_permission(request,"schools")
    con=db()
    guard=con.execute("SELECT * FROM guard_master WHERE emp_id=?",(emp_id,)).fetchone()
    if not guard:
        con.close();raise HTTPException(404)
    old_school=None
    if guard["linked_ministry_id"]:
        old_school=con.execute("SELECT * FROM schools WHERE ministry_id=?",
                               (guard["linked_ministry_id"],)).fetchone()
        con.execute("""UPDATE schools SET guard_name='',guard_emp_id='',guard_id='',guard_mobile=''
          WHERE ministry_id=?""",(guard["linked_ministry_id"],))
    con.execute("""UPDATE guard_master SET service_status='منتهي الخدمة',
      termination_date=?,termination_reason=?,linked_ministry_id=NULL,
      updated_at=?,updated_by=? WHERE emp_id=?""",(
      termination_date or date.today().isoformat(),termination_reason.strip(),
      datetime.utcnow().isoformat(),user["id"],emp_id))
    con.execute("""INSERT INTO guard_assignment_log(
      emp_id,emp_name,old_ministry_id,old_school_name,action_type,notes,changed_by,changed_at
    ) VALUES(?,?,?,?,?,?,?,?)""",(
      emp_id,guard["emp_name"],guard["linked_ministry_id"],
      old_school["school_name"] if old_school else None,"إنهاء خدمة",
      termination_reason.strip(),user["id"],datetime.utcnow().isoformat()))
    con.commit();con.close()
    return RedirectResponse("/unlinked-guards?terminated=1",status_code=303)

@app.post("/guards/{emp_id}/restore")
def guard_restore(request: Request,emp_id: str):
    user=require_permission(request,"schools")
    con=db()
    con.execute("""UPDATE guard_master SET service_status='على رأس العمل',
      termination_date=NULL,termination_reason=NULL,updated_at=?,updated_by=?
      WHERE emp_id=?""",(datetime.utcnow().isoformat(),user["id"],emp_id))
    con.commit();con.close()
    return RedirectResponse("/unlinked-guards?restored=1",status_code=303)

@app.post("/guards/{emp_id}/unlink")
def guard_unlink(request: Request,emp_id: str,notes: str=Form("")):
    user=require_permission(request,"schools")
    con=db()
    guard=con.execute("SELECT * FROM guard_master WHERE emp_id=?",(emp_id,)).fetchone()
    if not guard:
        con.close();raise HTTPException(404)
    old_school=None
    if guard["linked_ministry_id"]:
        old_school=con.execute("SELECT * FROM schools WHERE ministry_id=?",
                               (guard["linked_ministry_id"],)).fetchone()
        con.execute("""UPDATE schools SET guard_name='',guard_emp_id='',guard_id='',guard_mobile=''
          WHERE ministry_id=?""",(guard["linked_ministry_id"],))
    con.execute("""UPDATE guard_master SET linked_ministry_id=NULL,linked_at=NULL,
      updated_at=?,updated_by=? WHERE emp_id=?""",(datetime.utcnow().isoformat(),user["id"],emp_id))
    con.execute("""INSERT INTO guard_assignment_log(
      emp_id,emp_name,old_ministry_id,old_school_name,action_type,notes,changed_by,changed_at
    ) VALUES(?,?,?,?,?,?,?,?)""",(
      emp_id,guard["emp_name"],guard["linked_ministry_id"],
      old_school["school_name"] if old_school else None,"فك ارتباط",
      notes.strip(),user["id"],datetime.utcnow().isoformat()))
    con.commit();con.close()
    return RedirectResponse("/unlinked-guards?unlinked=1",status_code=303)

@app.post("/guards/{emp_id}/transfer")
def guard_transfer(
    request: Request,
    emp_id: str,
    new_ministry_id: str=Form(...),
    notes: str=Form("")
):
    user=require_permission(request,"schools")
    con=db()
    guard=con.execute("SELECT * FROM guard_master WHERE emp_id=?",(emp_id,)).fetchone()
    new_school=con.execute("SELECT * FROM schools WHERE ministry_id=?",(new_ministry_id,)).fetchone()
    if not guard or not new_school:
        con.close();raise HTTPException(404)
    if (guard["service_status"] or "على رأس العمل")=="منتهي الخدمة":
        con.close();raise HTTPException(400,detail="البواب منتهي الخدمة")
    if new_school["guard_emp_id"] and new_school["guard_emp_id"]!=emp_id:
        con.close();raise HTTPException(400,detail="المدرسة الجديدة مرتبطة ببواب آخر")
    old_school=None
    if guard["linked_ministry_id"]:
        old_school=con.execute("SELECT * FROM schools WHERE ministry_id=?",
                               (guard["linked_ministry_id"],)).fetchone()
        con.execute("""UPDATE schools SET guard_name='',guard_emp_id='',guard_id='',guard_mobile=''
          WHERE ministry_id=?""",(guard["linked_ministry_id"],))
    con.execute("""UPDATE schools SET guard_name=?,guard_emp_id=?,guard_id=?,guard_mobile=?
      WHERE ministry_id=?""",(guard["emp_name"],guard["emp_id"],guard["id_no"],guard["mobile"],new_ministry_id))
    con.execute("""UPDATE guard_master SET linked_ministry_id=?,linked_at=?,linked_by=?,
      updated_at=?,updated_by=? WHERE emp_id=?""",(
      new_ministry_id,datetime.utcnow().isoformat(),user["id"],
      datetime.utcnow().isoformat(),user["id"],emp_id))
    con.execute("""INSERT INTO guard_assignment_log(
      emp_id,emp_name,old_ministry_id,old_school_name,new_ministry_id,new_school_name,
      action_type,notes,changed_by,changed_at
    ) VALUES(?,?,?,?,?,?,?,?,?,?)""",(
      emp_id,guard["emp_name"],guard["linked_ministry_id"],
      old_school["school_name"] if old_school else None,new_ministry_id,new_school["school_name"],
      "نقل بين المدارس",notes.strip(),user["id"],datetime.utcnow().isoformat()))
    con.commit();con.close()
    return RedirectResponse("/unlinked-guards?transferred=1",status_code=303)

@app.post("/access-permits/{permit_id}/external-entry")
def external_permit_entry(
    request: Request,
    permit_id: int,
    entry_latitude: str=Form(""),
    entry_longitude: str=Form(""),
    entry_photo: UploadFile=File(...),
    notes: str=Form("")
):
    user=require_permission(request,"access_permits")
    if user["role"]!="external_requester":
        raise HTTPException(403)
    con=db()
    permit=con.execute(
      "SELECT * FROM school_access_permits WHERE id=? AND requester_user_id=?",
      (permit_id,user["id"])
    ).fetchone()
    if not permit:
        con.close();raise HTTPException(404)
    if permit["approval_status"]!="موافق":
        con.close();raise HTTPException(400,detail="التصريح لم تتم الموافقة عليه")
    if permit["actual_status"]!="لم يدخل":
        con.close();raise HTTPException(400,detail="تم تسجيل الدخول مسبقًا")
    photo_path,evidence_hash=save_access_evidence(entry_photo,permit_id,"entry")
    now=datetime.utcnow().isoformat()
    evidence_no=Path(photo_path).stem
    con.execute("""UPDATE school_access_permits SET
      entry_time=?,entry_server_time=?,entry_latitude=?,entry_longitude=?,
      entry_verified_by=?,entry_photo_path=?,entry_evidence_no=?,
      entry_evidence_hash=?,actual_status='داخل المدرسة',
      notes=CASE WHEN ?<>'' THEN COALESCE(notes,'')||CHAR(10)||? ELSE notes END
      WHERE id=?""",(
      datetime.now().strftime("%H:%M:%S"),now,float(entry_latitude or 0),
      float(entry_longitude or 0),user["full_name"],photo_path,evidence_no,
      evidence_hash,notes,notes,permit_id
    ))
    con.commit();con.close()
    return RedirectResponse("/access-permits?entry=1",status_code=303)

@app.post("/access-permits/{permit_id}/external-exit")
def external_permit_exit(
    request: Request,
    permit_id: int,
    exit_latitude: str=Form(""),
    exit_longitude: str=Form(""),
    exit_photo: UploadFile=File(...),
    work_summary: str=Form(...),
    notes: str=Form("")
):
    user=require_permission(request,"access_permits")
    if user["role"]!="external_requester":
        raise HTTPException(403)
    con=db()
    permit=con.execute(
      "SELECT * FROM school_access_permits WHERE id=? AND requester_user_id=?",
      (permit_id,user["id"])
    ).fetchone()
    if not permit:
        con.close();raise HTTPException(404)
    if permit["actual_status"]!="داخل المدرسة":
        con.close();raise HTTPException(400,detail="يجب إثبات الدخول أولاً")
    photo_path,evidence_hash=save_access_evidence(exit_photo,permit_id,"exit")
    now=datetime.utcnow().isoformat()
    evidence_no=Path(photo_path).stem
    con.execute("""UPDATE school_access_permits SET
      exit_time=?,exit_server_time=?,exit_latitude=?,exit_longitude=?,
      exit_verified_by=?,exit_photo_path=?,exit_evidence_no=?,
      exit_evidence_hash=?,work_summary=?,actual_status='تم الخروج',
      notes=CASE WHEN ?<>'' THEN COALESCE(notes,'')||CHAR(10)||? ELSE notes END
      WHERE id=?""",(
      datetime.now().strftime("%H:%M:%S"),now,float(exit_latitude or 0),
      float(exit_longitude or 0),user["full_name"],photo_path,evidence_no,
      evidence_hash,work_summary,notes,notes,permit_id
    ))
    con.commit();con.close()
    return RedirectResponse("/access-permits?exit=1",status_code=303)
