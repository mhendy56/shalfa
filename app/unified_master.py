
from __future__ import annotations
import json
from datetime import datetime
from typing import Any
from .database import connect

def _now() -> str:
    return datetime.utcnow().isoformat()

def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()

def ensure_unified_schema() -> None:
    con=connect()
    con.executescript("""
    CREATE TABLE IF NOT EXISTS master_schools(
      ministry_id TEXT PRIMARY KEY,
      school_name TEXT,city TEXT,district TEXT,stage TEXT,gender TEXT,
      school_type TEXT,operational_type TEXT,size TEXT,status TEXT,
      latitude REAL,longitude REAL,updated_at TEXT
    );
    CREATE TABLE IF NOT EXISTS master_guards(
      emp_id TEXT PRIMARY KEY,full_name TEXT,id_no TEXT,mobile TEXT,
      service_status TEXT,is_active INTEGER DEFAULT 1,updated_at TEXT
    );
    CREATE TABLE IF NOT EXISTS master_coordinators(
      coordinator_key TEXT PRIMARY KEY,emp_id TEXT,full_name TEXT,mobile TEXT,
      city TEXT,sector TEXT,is_active INTEGER DEFAULT 1,updated_at TEXT
    );
    CREATE TABLE IF NOT EXISTS school_guard_links(
      ministry_id TEXT PRIMARY KEY,guard_emp_id TEXT,linked_at TEXT,source TEXT
    );
    CREATE TABLE IF NOT EXISTS school_coordinator_links(
      ministry_id TEXT PRIMARY KEY,coordinator_key TEXT,linked_at TEXT,source TEXT
    );
    CREATE TABLE IF NOT EXISTS relationship_events(
      id INTEGER PRIMARY KEY AUTOINCREMENT,entity_type TEXT,entity_id TEXT,
      action TEXT,payload_json TEXT,created_at TEXT
    );
    CREATE TABLE IF NOT EXISTS unified_sync_state(
      sync_key TEXT PRIMARY KEY,sync_value TEXT,updated_at TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_master_school_city ON master_schools(city);
    CREATE INDEX IF NOT EXISTS idx_school_guard_emp ON school_guard_links(guard_emp_id);
    CREATE INDEX IF NOT EXISTS idx_school_coord_key ON school_coordinator_links(coordinator_key);
    CREATE INDEX IF NOT EXISTS idx_relationship_events_created ON relationship_events(created_at);
    """)
    con.commit();con.close()

def record_relationship_event(entity_type: str, entity_id: str, action: str, payload: dict|None=None) -> None:
    con=connect()
    con.execute("""INSERT INTO relationship_events(entity_type,entity_id,action,payload_json,created_at)
      VALUES(?,?,?,?,?)""",(entity_type,entity_id,action,json.dumps(payload or {},ensure_ascii=False),_now()))
    con.commit();con.close()

def reconcile_unified_data(source: str="system", emit_event: bool=False) -> dict:
    ensure_unified_schema()
    con=connect();now=_now()
    stats={"schools":0,"guards":0,"coordinators":0,"guard_links":0,"coordinator_links":0}
    school_cols={r["name"] for r in con.execute("PRAGMA table_info(schools)").fetchall()}
    rows=con.execute("SELECT * FROM schools").fetchall()
    for row in rows:
        ministry_id=_text(row["ministry_id"])
        if not ministry_id: continue
        def get(name,default=""):
            return row[name] if name in school_cols else default
        con.execute("""INSERT INTO master_schools(
          ministry_id,school_name,city,district,stage,gender,school_type,
          operational_type,size,status,latitude,longitude,updated_at
        ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(ministry_id) DO UPDATE SET
          school_name=excluded.school_name,city=excluded.city,district=excluded.district,
          stage=excluded.stage,gender=excluded.gender,school_type=excluded.school_type,
          operational_type=excluded.operational_type,size=excluded.size,status=excluded.status,
          latitude=excluded.latitude,longitude=excluded.longitude,updated_at=excluded.updated_at""",(
          ministry_id,_text(get("school_name")),_text(get("city")),_text(get("district")),
          _text(get("stage")),_text(get("gender")),_text(get("school_type")),
          _text(get("operational_type")),_text(get("size")),_text(get("status") or "تعمل"),
          get("latitude",None),get("longitude",None),now
        ))
        stats["schools"]+=1
        guard_emp=_text(get("guard_emp_id"))
        if guard_emp:
            con.execute("""INSERT INTO master_guards(emp_id,full_name,id_no,mobile,service_status,is_active,updated_at)
              VALUES(?,?,?,?,?,1,?)
              ON CONFLICT(emp_id) DO UPDATE SET full_name=excluded.full_name,id_no=excluded.id_no,
              mobile=excluded.mobile,is_active=1,updated_at=excluded.updated_at""",(
              guard_emp,_text(get("guard_name")),_text(get("guard_id")),_text(get("guard_mobile")),"نشط",now
            ))
            con.execute("""INSERT INTO school_guard_links(ministry_id,guard_emp_id,linked_at,source)
              VALUES(?,?,?,?) ON CONFLICT(ministry_id) DO UPDATE SET
              guard_emp_id=excluded.guard_emp_id,linked_at=excluded.linked_at,source=excluded.source""",
              (ministry_id,guard_emp,now,source))
            stats["guards"]+=1;stats["guard_links"]+=1
        else:
            con.execute("DELETE FROM school_guard_links WHERE ministry_id=?",(ministry_id,))
        coord_name=_text(get("coordinator_name"))
        coord_emp=_text(get("coordinator_emp_id"))
        if coord_name or coord_emp:
            coord_key=coord_emp or f"NAME:{coord_name}"
            con.execute("""INSERT INTO master_coordinators(
              coordinator_key,emp_id,full_name,mobile,city,sector,is_active,updated_at
            ) VALUES(?,?,?,?,?,?,1,?)
            ON CONFLICT(coordinator_key) DO UPDATE SET emp_id=excluded.emp_id,
              full_name=excluded.full_name,mobile=excluded.mobile,city=excluded.city,
              is_active=1,updated_at=excluded.updated_at""",(
              coord_key,coord_emp,coord_name,_text(get("coordinator_mobile")),
              _text(get("city")),"",now
            ))
            con.execute("""INSERT INTO school_coordinator_links(ministry_id,coordinator_key,linked_at,source)
              VALUES(?,?,?,?) ON CONFLICT(ministry_id) DO UPDATE SET
              coordinator_key=excluded.coordinator_key,linked_at=excluded.linked_at,source=excluded.source""",
              (ministry_id,coord_key,now,source))
            stats["coordinators"]+=1;stats["coordinator_links"]+=1
        else:
            con.execute("DELETE FROM school_coordinator_links WHERE ministry_id=?",(ministry_id,))
    # Remove links to schools that no longer exist.
    con.execute("DELETE FROM school_guard_links WHERE ministry_id NOT IN (SELECT ministry_id FROM schools)")
    con.execute("DELETE FROM school_coordinator_links WHERE ministry_id NOT IN (SELECT ministry_id FROM schools)")
    con.execute("""INSERT INTO unified_sync_state(sync_key,sync_value,updated_at)
      VALUES('last_reconcile',?,?) ON CONFLICT(sync_key) DO UPDATE SET
      sync_value=excluded.sync_value,updated_at=excluded.updated_at""",(json.dumps(stats,ensure_ascii=False),now))
    if emit_event:
        con.execute("""INSERT INTO relationship_events(entity_type,entity_id,action,payload_json,created_at)
          VALUES('system','all','reconcile',?,?)""",(json.dumps(stats,ensure_ascii=False),now))
    con.commit();con.close()
    return stats

def relationship_summary() -> dict:
    ensure_unified_schema()
    con=connect()
    def scalar(sql,params=()):
        row=con.execute(sql,params).fetchone()
        return int(row[0] if row else 0)
    result={
      "schools":scalar("SELECT COUNT(*) FROM master_schools"),
      "guards":scalar("SELECT COUNT(*) FROM master_guards WHERE is_active=1"),
      "coordinators":scalar("SELECT COUNT(*) FROM master_coordinators WHERE is_active=1"),
      "schools_without_guard":scalar("""SELECT COUNT(*) FROM master_schools s
        LEFT JOIN school_guard_links l ON l.ministry_id=s.ministry_id
        WHERE COALESCE(l.guard_emp_id,'')='' AND COALESCE(s.status,'تعمل')='تعمل'"""),
      "schools_without_coordinator":scalar("""SELECT COUNT(*) FROM master_schools s
        LEFT JOIN school_coordinator_links l ON l.ministry_id=s.ministry_id
        WHERE COALESCE(l.coordinator_key,'')='' AND COALESCE(s.status,'تعمل')='تعمل'"""),
      "unlinked_guards":scalar("""SELECT COUNT(*) FROM master_guards g
        LEFT JOIN school_guard_links l ON l.guard_emp_id=g.emp_id
        WHERE l.ministry_id IS NULL AND g.is_active=1"""),
      "coordinators_without_schools":scalar("""SELECT COUNT(*) FROM master_coordinators c
        LEFT JOIN school_coordinator_links l ON l.coordinator_key=c.coordinator_key
        WHERE l.ministry_id IS NULL AND c.is_active=1"""),
      "duplicate_guard_links":scalar("""SELECT COUNT(*) FROM(
        SELECT guard_emp_id FROM school_guard_links WHERE COALESCE(guard_emp_id,'')<>''
        GROUP BY guard_emp_id HAVING COUNT(*)>1
      ) x"""),
      "last_event_id":scalar("SELECT COALESCE(MAX(id),0) FROM relationship_events")
    }
    con.close();return result
