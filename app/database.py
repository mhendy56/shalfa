from __future__ import annotations
import re
import sqlite3
from collections.abc import Iterable, Iterator, Mapping
from typing import Any
from .config import settings

IntegrityError = sqlite3.IntegrityError

class CompatRow(Mapping):
    def __init__(self, columns: list[str], values: Iterable[Any]):
        self._columns = columns
        self._values = tuple(values)
        self._data = dict(zip(columns, self._values))
    def __getitem__(self, key):
        if isinstance(key, int): return self._values[key]
        return self._data[key]
    def __iter__(self) -> Iterator[str]: return iter(self._columns)
    def __len__(self) -> int: return len(self._columns)
    def keys(self): return self._data.keys()

class CursorCompat:
    def __init__(self, cursor, postgres: bool):
        self._cursor=cursor; self._postgres=postgres
        self.lastrowid=getattr(cursor,"lastrowid",None)
        self.rowcount=getattr(cursor,"rowcount",-1)
    def _row(self,row):
        if row is None:return None
        if isinstance(row,(sqlite3.Row,CompatRow)):return row
        cols=[d.name if hasattr(d,"name") else d[0] for d in (self._cursor.description or [])]
        return CompatRow(cols,row)
    def fetchone(self): return self._row(self._cursor.fetchone())
    def fetchall(self): return [self._row(r) for r in self._cursor.fetchall()]
    def __iter__(self):
        for r in self._cursor: yield self._row(r)

class ConnectionCompat:
    def __init__(self, raw, postgres: bool): self.raw=raw; self.postgres=postgres
    def execute(self,sql:str,params:Iterable[Any]=()):
        if self.postgres:
            sql,params=_postgres_sql(sql,tuple(params)); cur=self.raw.cursor(); cur.execute(sql,params)
        else: cur=self.raw.execute(sql,tuple(params))
        return CursorCompat(cur,self.postgres)
    def executemany(self,sql:str,seq):
        if self.postgres:
            sql,_=_postgres_sql(sql,()); cur=self.raw.cursor(); cur.executemany(sql,seq)
        else: cur=self.raw.executemany(sql,seq)
        return CursorCompat(cur,self.postgres)
    def executescript(self,script:str):
        if not self.postgres: self.raw.executescript(script); return
        for statement in _split_sql(script):
            if statement.strip(): self.execute(_postgres_ddl(statement))
    def commit(self): self.raw.commit()
    def rollback(self): self.raw.rollback()
    def close(self): self.raw.close()
    def __enter__(self): return self
    def __exit__(self,exc_type,exc,tb):
        if exc:self.rollback()
        else:self.commit()
        self.close()


def connect() -> ConnectionCompat:
    if settings.database_engine=="postgresql":
        import psycopg
        global IntegrityError
        IntegrityError=psycopg.IntegrityError
        url=settings.database_url.replace("postgres://","postgresql://",1)
        return ConnectionCompat(psycopg.connect(url),True)
    settings.database_path.parent.mkdir(parents=True,exist_ok=True)
    raw=sqlite3.connect(settings.database_path); raw.row_factory=sqlite3.Row
    raw.execute("PRAGMA foreign_keys=ON")
    return ConnectionCompat(raw,False)


def database_engine()->str:return settings.database_engine

def _split_sql(script:str)->list[str]:
    return [x for x in re.split(r";\s*(?:\n|$)",script) if x.strip()]

def _postgres_ddl(sql:str)->str:
    sql=re.sub(r"INTEGER\s+PRIMARY\s+KEY\s+AUTOINCREMENT","BIGSERIAL PRIMARY KEY",sql,flags=re.I)
    sql=re.sub(r"\bAUTOINCREMENT\b","",sql,flags=re.I)
    return sql

def _qmark_to_percent(sql:str)->str:
    out=[]; quote=None
    for ch in sql:
        if ch in ("'",'\"'):
            if quote==ch:quote=None
            elif quote is None:quote=ch
        if ch=='?' and quote is None:out.append('%s')
        else:out.append(ch)
    return ''.join(out)

def _postgres_sql(sql:str,params:tuple[Any,...]):
    stripped=sql.strip()
    m=re.match(r"PRAGMA\s+table_info\(([^)]+)\)",stripped,re.I)
    if m:
        table=m.group(1).strip(" '\"")
        return ("SELECT column_name AS name,data_type AS type FROM information_schema.columns WHERE table_schema='public' AND table_name=%s ORDER BY ordinal_position",(table,))
    sql=_postgres_ddl(sql)
    sql=re.sub(r"INSERT\s+OR\s+IGNORE\s+INTO","INSERT INTO",sql,flags=re.I)
    if re.search(r"INSERT\s+INTO",sql,re.I) and "OR IGNORE" not in stripped.upper() and "ON CONFLICT" not in sql.upper() and "INSERT OR IGNORE" in stripped.upper():
        sql=sql.rstrip()+" ON CONFLICT DO NOTHING"
    # SQLite date recursive CTE compatibility.
    sql=re.sub(r"date\((%s)\)",r"CAST(\1 AS date)",sql,flags=re.I)
    sql=sql.replace("date(d,'+1 day')","(d + INTERVAL '1 day')::date")
    return _qmark_to_percent(sql),params
