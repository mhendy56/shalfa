from __future__ import annotations
import os,sqlite3,sys
import psycopg
from psycopg import sql

def main(source:str):
    url=os.environ.get('DATABASE_URL','')
    if not url:raise SystemExit('DATABASE_URL is required')
    src=sqlite3.connect(source);src.row_factory=sqlite3.Row
    dst=psycopg.connect(url.replace('postgres://','postgresql://',1))
    tables=[r['name'] for r in src.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")]
    with dst.cursor() as cur:
        for table in tables:
            cols=[r['name'] for r in src.execute(f'PRAGMA table_info("{table}")')]
            if not cols:continue
            rows=src.execute(f'SELECT * FROM "{table}"').fetchall()
            if not rows:continue
            query=sql.SQL('INSERT INTO {} ({}) VALUES ({}) ON CONFLICT DO NOTHING').format(
                sql.Identifier(table),sql.SQL(',').join(map(sql.Identifier,cols)),sql.SQL(',').join(sql.Placeholder()*len(cols)))
            cur.executemany(query,[tuple(r[c] for c in cols) for r in rows])
            print(table,len(rows))
    dst.commit();dst.close();src.close()
if __name__=='__main__':main(sys.argv[1] if len(sys.argv)>1 else 'data/shalfa.db')
