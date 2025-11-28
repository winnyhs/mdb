import datetime
from decimal import Decimal

from lib.singleton import SingletonMeta
from lib.log import logger
from utils.db_ctrl import DbCtrl

class __SQL(metaclass=SingletonMeta): 

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


    # Generic Execute SQL (Non-SELECT)
    def execute(self, db, sql):
        """
        UPDATE, INSERT, DELETE 등 결과를 반환하지 않는 SQL 실행.
        """
        logger.debug("SQL: %s", sql)
        try:
            db.Execute(sql)
        except Exception as e:
            raise RuntimeError(f"SQL Error: {sql} → {e}")

    # SELECT 실행 → 리스트(dict) 반환
    def query(self, db, sql):
        """
        SELECT 쿼리를 실행하고 결과를 list[dict] 형태로 반환.
        """
        # logger.debug("SQL: %s", sql)
        try:
            rs = db.OpenRecordset(sql)
        except Exception as e:
            raise RuntimeError("ERROR: %s: SQL = %s", e, sql)

        rows = []
        if rs.EOF:
            rs.Close()
            return rows

        fields = [f.Name for f in rs.Fields]

        rs.MoveFirst()
        while not rs.EOF:
            # row = {f: rs.Fields(f).Value for f in fields}
            row = {f: DbCtrl.normalize_value(rs.Fields(f).Value) for f in fields}
            rows.append(row)
            rs.MoveNext()

        rs.Close()
        return rows

    # 단일 값 SELECT (예: COUNT)
    def query_scalar(self, db, sql):
        """
        SELECT COUNT(*) AS CNT FROM ... 같이 단일 값만 반환하는 쿼리.
        """
        result = query_sql(db, sql)
        if not result:
            return None

        # 첫 row의 첫 value
        return next(iter(result[0].values()))

    # 각 테이블의 Row Count 얻기
    def table_row_count(self, db, table_name):
        sql = f"SELECT COUNT(*) AS CNT FROM [{table_name}]"
        return query_scalar(db, sql)

    def all_table_row_counts(self, db):
        """
        DB의 모든 테이블에 대해 Row Count 반환
        { "TableA": 123, "TableB": 4421, ... }
        """
        tables = [t.Name for t in db.TableDefs if not t.Name.startswith("MSys")]

        counts = {}
        for tbl in tables:
            cnt = table_row_count(db, tbl)
            counts[tbl] = cnt
        return counts

Sql = __SQL()