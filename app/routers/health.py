from fastapi import APIRouter
from app.database import connect,database_engine
router=APIRouter(tags=["health"])
@router.get("/health")
def health(): return {"ok":True,"service":"shalfa-operations"}
@router.get("/health/database")
def database_health():
    with connect() as con:
        row=con.execute("SELECT 1 AS ok").fetchone()
    return {"ok":bool(row["ok"]),"engine":database_engine()}
