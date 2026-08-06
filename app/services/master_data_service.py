
from __future__ import annotations
from datetime import datetime
from ..database import connect
from ..unified_master import reconcile_unified_data,record_relationship_event

class MasterDataService:
    @staticmethod
    def assign_guard(ministry_id:str,emp_id:str,user_id:int|None=None,reason:str="") -> None:
        con=connect()
        school=con.execute("SELECT * FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
        guard=con.execute("SELECT * FROM guard_master WHERE emp_id=?",(emp_id,)).fetchone()
        if not school: con.close();raise ValueError("المدرسة غير موجودة")
        if not guard: con.close();raise ValueError("البواب غير موجود")
        # One active school per guard.
        con.execute("""UPDATE schools SET guard_name='',guard_emp_id='',guard_id='',guard_mobile=''
          WHERE guard_emp_id=? AND ministry_id<>?""",(emp_id,ministry_id))
        con.execute("""UPDATE schools SET guard_name=?,guard_emp_id=?,guard_id=?,guard_mobile=?
          WHERE ministry_id=?""",(guard["emp_name"],emp_id,guard["id_no"] or "",guard["mobile"] or "",ministry_id))
        try: con.execute("UPDATE guard_master SET linked_ministry_id=? WHERE emp_id=?",(ministry_id,emp_id))
        except Exception: pass
        con.commit();con.close()
        reconcile_unified_data("guard_assignment")
        record_relationship_event("guard",emp_id,"assign",{"ministry_id":ministry_id,"reason":reason,"user_id":user_id})

    @staticmethod
    def unassign_guard(ministry_id:str,user_id:int|None=None,reason:str="") -> None:
        con=connect();school=con.execute("SELECT * FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
        if not school: con.close();raise ValueError("المدرسة غير موجودة")
        old=school["guard_emp_id"] or ""
        con.execute("UPDATE schools SET guard_name='',guard_emp_id='',guard_id='',guard_mobile='' WHERE ministry_id=?",(ministry_id,))
        try:
            if old: con.execute("UPDATE guard_master SET linked_ministry_id=NULL WHERE emp_id=?",(old,))
        except Exception: pass
        con.commit();con.close()
        reconcile_unified_data("guard_unassignment")
        record_relationship_event("guard",old,"unassign",{"ministry_id":ministry_id,"reason":reason,"user_id":user_id})

    @staticmethod
    def assign_coordinator(ministry_id:str,coordinator_id:int,user_id:int|None=None,reason:str="") -> None:
        con=connect()
        school=con.execute("SELECT * FROM schools WHERE ministry_id=?",(ministry_id,)).fetchone()
        coord=con.execute("SELECT * FROM coordinator_master WHERE id=? AND is_active=1",(coordinator_id,)).fetchone()
        if not school: con.close();raise ValueError("المدرسة غير موجودة")
        if not coord: con.close();raise ValueError("المنسق / المشرف غير موجود أو موقوف")
        con.execute("""UPDATE schools SET coordinator_name=?,coordinator_emp_id=?,coordinator_mobile=?
          WHERE ministry_id=?""",(coord["full_name"],coord["emp_id"] or "",coord["mobile"] or "",ministry_id))
        con.commit();con.close()
        reconcile_unified_data("coordinator_assignment")
        record_relationship_event("coordinator",coord["emp_id"] or str(coordinator_id),"assign",
          {"ministry_id":ministry_id,"reason":reason,"user_id":user_id})

    @staticmethod
    def set_school_status(ministry_id:str,status:str,user_id:int|None=None,reason:str="") -> None:
        if status not in ("تعمل","مغلقة","مدمجة"): raise ValueError("حالة المدرسة غير صحيحة")
        con=connect();con.execute("UPDATE schools SET status=? WHERE ministry_id=?",(status,ministry_id))
        con.commit();con.close()
        reconcile_unified_data("school_status")
        record_relationship_event("school",ministry_id,"status",{"status":status,"reason":reason,"user_id":user_id})
