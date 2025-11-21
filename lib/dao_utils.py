# dao_utils.py
# - DAO 엔진 감지 + 테이블 목록 + 필드 타입
import datetime
from decimal import Decimal
from win32com.client import Dispatch


def get_dao_engine():
    """
    DAO 엔진 자동 감지 (WinXP~Win10 호환)
    """
    candidates = [
        "DAO.DBEngine.120",  # Jet 4.0 (Win7/Win10)
        "DAO.DBEngine.36",   # Jet 3.6 (WinXP)
        "DAO.DBEngine.35",   # Jet 3.5 (Win9x/XP)
        "DAO.DBEngine"       # fallback
    ]
    for v in candidates:
        try:
            return Dispatch(v)
        except:
            continue
    raise RuntimeError("No usable DAO engine found.")


def detect_mdb_version(mdb_path, password=""):
    """
    Jet Engine Engine Type	의미: 
    1	Access 95 (Jet 3.0)
    2	Access 97 (Jet 3.5)
    3	Access 2000 (Jet 4.0)
    4	Access 2002-2003 (Jet 4.0, updated)
    5	DAO 4.x / ACE 초기 포맷

    Version	의미
    "3.0"	Access 95
    "3.5"	Access 97
    "4.0"	Access 2000~2003
    "12.0"	Access 2007+ (ACCDB 전환 가능)

    MEDICAL.mdb : version = 3.0, eng_type = None
    """
    engine = get_dao_engine()
    connect = f";PWD={password}"
    db = engine.OpenDatabase(mdb_path, False, False, connect)

    version = db.Version
    try:
        eng_type = db.Properties("Jet OLEDB:Engine Type").Value
    except:
        eng_type = None

    db.Close()
    return version, eng_type


def list_tables(db):
    """Access 시스템 테이블 제외한 사용자 테이블 목록"""
    tables = []
    for t in db.TableDefs:
        name = t.Name
        if not name.startswith("MSys"):
            tables.append(name)
    return tables


def get_field_types(tabledef):
    """테이블의 필드 타입 dict 반환 {필드명: 타입코드}"""
    out = {}
    for f in tabledef.Fields:
        out[f.Name] = f.Type
    return out


def get_primary_index(tabledef):
    """
    테이블 정의에서 Primary Index를 찾아 인덱스 이름과 필드 목록을 반환.
    없으면 (None, []) 리턴.
    """
    for idx in tabledef.Indexes:
        try:
            if idx.Primary:
                fields = [f.Name for f in idx.Fields]
                return idx.Name, fields
        except Exception:
            continue
    return None, []


# ───────────────────────────────────────────
# 값 변환 (MDB → JSON)
# ───────────────────────────────────────────
def normalize_value(v):
    """DAO 값 → JSON-friendly 값"""
    if isinstance(v, Decimal):
        return float(v)
    if isinstance(v, datetime.datetime):
        return v.isoformat()
    if isinstance(v, datetime.date):
        return v.isoformat()
    return v


# ───────────────────────────────────────────
# 값 역변환 (JSON → MDB)
# ───────────────────────────────────────────
def restore_value(v, field_type):
    """JSON 값 → Access DAO Field 값으로 역변환"""
    # None 그대로
    if v is None:
        return None

    # DAO Field Type 참고:
    # https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/fieldtype-enumeration-dao
    #
    # 주요 타입:
    # 1 YESNO
    # 3 INTEGER
    # 4 LONG
    # 5 CURRENCY
    # 6 SINGLE
    # 7 DOUBLE
    # 8 DATETIME
    # 10 TEXT
    # 12 MEMO

    # 숫자
    if field_type in (3, 4, 5, 6, 7):
        return Decimal(str(v))

    # 날짜/시간
    if field_type == 8:
        try:
            return datetime.datetime.fromisoformat(v)
        except Exception:
            return v

    # TEXT/MEMO/기타
    return v
