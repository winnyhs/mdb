# -*- coding: utf-8 -*-
# - DAO 엔진 감지 + 테이블 목록 + 필드 타입
import os, shutil, time, datetime
import pythoncom
from decimal import Decimal
from win32com.client import Dispatch
from functools import lru_cache

from lib.singleton import SingletonMeta
from lib.log import logger

class DbCtrl: 
    def __init__(self, mdb_path, password = ""):
        self.engine, self.engine_name = self.detect_engine()
        # 강제로 workgroup 초기화 (보안파일 detach)
        self.engine.SystemDB = ""

        ws = self.engine.CreateWorkspace("", "admin", "")
        self.mdb_path = mdb_path
        self.password = password 

        db = self.open_db()
        self.version = DbCtrl.detect_version(db)
        db.Close()

    @staticmethod
    def detect_engine(): #self): # DAO 엔진 자동 감지 (WinXP~Win10 호환)
        engine, engine_name = None, None
        candidates = [
            "DAO.DBEngine.35",   # Jet 3.5 (Win9x/XP)
            "DAO.DBEngine.36",   # Jet 3.6 (WinXP)
            "DAO.DBEngine.120",  # Jet 4.0 (Win7/Win10)
            "DAO.DBEngine"       # fallback
        ]
        for v in candidates:
            try:
                engine_name = v
                engine = Dispatch(v)
            except:
                continue
        if engine == None: 
            raise RuntimeError("No usable DAO engine found.")

        logger.info("DAO Engine: %s", engine_name)
        return engine, engine_name
    
    @staticmethod
    def detect_version(db):
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
        logger.info("version={}".format(db.version))
        return db.version

    def open_db(self):
        logger.info("DB Opened: %s", self.mdb_path)
        # print(os.access(self.mdb_path, os.W_OK)) True. No prob
        connect = "" if self.password is None else ";PWD={}".format(self.password)
        # Exclusive: False, Read-Only: False
        return self.engine.OpenDatabase(self.mdb_path, False, False, connect)

    def close_db(self, db):
        logger.info("DB Closeded: %s", self.mdb_path)
        db.Close() 

    def compact_db(self, db_path, password=None):
        """
        :param db_path: 원본 MDB 경로
        :param compact_path: compact 결과 저장할 파일 (None이면 자동 생성)
        :param password: DB 비밀번호 (없으면 None)
        :return: True/False
        """
        if db_path is None: 
            dp_path = self.mdb_path
            password = self.password

        if not os.path.exists(db_path):
            logger.error("ERROR: Source DB does not exist: %s", db_path)
            return False

        # Compact 후 생성될 파일
        base, ext = os.path.splitext(db_path)
        compact_path = base + "_compact" + ext

        # 이전에 있으면 삭제
        if os.path.exists(compact_path):
            try:
                os.remove(compact_path)
            except:
                logger.error("ERROR: Could not remove old compact file: %s", compact_path)
                return False

        # COM 초기화 (XP 필수)
        pythoncom.CoInitialize()
        self.engine, self.engine_name = self.detect_engine()
        src_conn = ";PWD=" + password if password else ""
        
        logger.debug("------ %s", password)
        # 압축 실행
        logger.info("Compacting %s...", db_path)
        try:
            self.engine.CompactDatabase(
                db_path,
                compact_path,
                None,  # default locale
                0,     # options
                src_conn
            )
        except Exception as e:
            logger.exception("ERROR: %s: Compact error:", e)
            return False

        # Compact 성공, 원본 교체
        try:
            backup_path = db_path + ".bak_" + time.strftime("%Y%m%d_%H%M%S")
            shutil.move(db_path, backup_path)
            shutil.move(compact_path, db_path)
        except Exception as e:
            logger.exception("ERROR: %s: Compacted but failed in replacing:", e)
            return False
        else: 
            logger.info("Backup created: %s", backup_path)
            logger.info("Compact DB replaced original file.")

        return True

    # 값 변환 (MDB → JSON)
    @staticmethod
    @lru_cache(maxsize=50000) 
    def normalize_value(v): # self, v):
        """DAO 값 → JSON-friendly 값"""
        if isinstance(v, Decimal):
            return float(v)
        if isinstance(v, datetime.datetime):
            return v.isoformat()
        if isinstance(v, datetime.date):
            return v.isoformat()
        return v

    # 값 역변환 (JSON → MDB)
    @staticmethod
    @lru_cache(maxsize=50000) 
    def restore_value(v, field_type): # self, v, field_type):
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




    # -------- Used for initial analysis
    # -------- Not used any more

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



def diagnose_dao(path, password = "I1D2E3A4"):
    print("=== DAO Diagnostic ===")

    dao = win32com.client.Dispatch("DAO.DBEngine.36")
    dao.SystemDB = ""

    connect = ";PWD={}".format(password) if password else ""

    try:
        db = dao.OpenDatabase(path, False, False, connect)
        print("[OK] Database opened")
    except Exception as e:
        print("[ERR] OpenDatabase:", e)
        return

    # 2) Connection Mode -------------------------------------
    try:
        print("db.Updatable:", db.Updatable)
    except Exception as e:
        print("Updatable read ERROR:", e)

    try:
        print("db.ReadOnly:", db.ReadOnly)
    except Exception as e:
        print("ReadOnly read ERROR:", e)

    # 3) Try BeginTrans --------------------------------------
    try:
        db.BeginTrans()
        print("BeginTrans: OK")
        db.Rollback()
    except Exception as e:
        print("BeginTrans ERROR:", e)

    # 4) Workspace-based test --------------------------------
    print("\n-- Workspace test --")
    ws = dao.CreateWorkspace("", "admin", "")
    try:
        db2 = ws.OpenDatabase(path, False, False, connect)
        print("WS.Updatable:", db2.Updatable)
        print("WS.ReadOnly:", db2.ReadOnly)
        db2.BeginTrans()
        print("WS BeginTrans: OK")
        db2.Rollback()
    except Exception as e:
        print("WS BeginTrans ERROR:", e)

    print("=== END ===")

