# dao_utils.py
# - DAO 엔진 감지 + 테이블 목록 + 필드 타입
import datetime
from decimal import Decimal
from win32com.client import Dispatch
# from lib.singleton import SingletonMeta

from lib.log import logger

class DbEngine(): 
    def __init__(self):
        self.mdb_path = None
        self.password = None

        # DAO 엔진 자동 감지 (WinXP~Win10 호환)
        self.engine = None
        self.engine_name = None
        candidates = [
            "DAO.DBEngine.35",   # Jet 3.5 (Win9x/XP)
            "DAO.DBEngine.36",   # Jet 3.6 (WinXP)
            "DAO.DBEngine.120",  # Jet 4.0 (Win7/Win10)
            "DAO.DBEngine"       # fallback
        ]
        for v in candidates:
            try:
                self.engine_name = v
                self.engine = Dispatch(v)
            except:
                continue
        if self.engine == None: 
            raise RuntimeError("No usable DAO engine found.")
        logger.info("DAO Engine: %s", self.engine_name)


    def open_db(self):
        connect = ";PWD={}".format(self.password)
        return self.engine.OpenDatabase(self.mdb_path, False, False, connect)

    def close_db(self, db): 
        db.Close() 


    def detect_version(self, db):
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
        self.version = db.Version
        logger.info("engine: name={}, version={}".format(\
                    self.engine_name, self.version))
        return self.version

    # 값 변환 (MDB → JSON)
    def normalize_value(self, v):
        """DAO 값 → JSON-friendly 값"""
        if isinstance(v, Decimal):
            return float(v)
        if isinstance(v, datetime.datetime):
            return v.isoformat()
        if isinstance(v, datetime.date):
            return v.isoformat()
        return v

    # 값 역변환 (JSON → MDB)
    def restore_value(self, v, field_type):
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


    # Read a whole table
    def read_table(self, db, table_name):
        rs = db.OpenRecordset(table_name)

        rows = []
        fields = [f.Name for f in rs.Fields]

        if not rs.EOF:
            rs.MoveFirst()
            while not rs.EOF:
                row = {f: self.normalize_value(rs.Fields(f).Value) for f in fields}
                rows.append(row)
                rs.MoveNext()

        rs.Close()
        return rows

    # Export a whole list or the result of a SQL as a json file
    def export_json(self, db, json_path, table_name = "", sql = ""): 
        if table_name: 
            rows = self.read_table(db, table_name)
            save_json(json_path)
            return rows

        elif sql: 
            rows = self.sql.query(db, sql)
            save_json(json_path)

    # Insert the given rows of the same schema into a table
    def insert_table(self, db, table_name, rows):
        tdef = db.TableDefs(table_name)
        field_types = self.get_field_types(tdef)

        rs = db.OpenRecordset(table_name)

        for row in rows:
            rs.AddNew()
            for field_name, value in row.items():

                if field_name not in field_types:
                    continue  # JSON에만 있고 MDB에는 없는 필드 방지

                ftype = field_types[field_name]
                v2 = self.restore_value(value, ftype)

                rs.Fields(field_name).Value = v2
            rs.Update()

        rs.Close()


    # ----------------------
    # For initial inspection
    # ----------------------
    def mdb_ddl_to_json(self, db, json_path = ""):
        """
        MDB 전체 테이블 구조(DDL에 대응하는 메타데이터)를 하나의 JSON으로 저장.
        """
        tables = self.list_tables()
        logger.info("Tables:", tables)

        ddl = {}

        for tbl in tables:
            tdef = db.TableDefs(tbl)
            ddl[tbl] = extract_table_ddl(tdef)

        if json_path: 
            save_json(ddl, json_path)
            logger.info("DDL of tables saved:", json_path)
        return ddl

    def list_tables(self, db):
        """Access 시스템 테이블 제외한 사용자 테이블 목록"""
        tables = []
        for t in db.TableDefs:
            name = t.Name
            if not name.startswith("MSys"):
                tables.append(name)
        return tables

    def get_field_types(self, tabledef):
        """테이블의 필드 타입 dict 반환 {필드명: 타입코드}"""
        out = {}
        for f in tabledef.Fields:
            out[f.Name] = f.Type
        return out

    def get_primary_index(self, tabledef):
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

    def get_table_ddl(self, tabledef):
        """DAO TableDef에서 필드, 인덱스 정보를 추출해 dict로 반환"""
        fields = []
        for f in tabledef.Fields:
            fields.append({
                "name": f.Name,
                "type": f.Type,
                "size": getattr(f, "Size", None),
                "attributes": getattr(f, "Attributes", None),
                # 필요하면 Required 등 추가 가능
            })

        indexes = []
        for idx in tabledef.Indexes:
            try:
                idx_fields = [f.Name for f in idx.Fields]
                indexes.append({
                    "name": idx.Name,
                    "primary": bool(getattr(idx, "Primary", False)),
                    "unique": bool(getattr(idx, "Unique", False)),
                    "required": bool(getattr(idx, "Required", False)),
                    "fields": idx_fields,
                })
            except Exception:
                continue

        return {
            "fields": fields,
            "indexes": indexes,
        }
    
    