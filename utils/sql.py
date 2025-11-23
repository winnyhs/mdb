from lib.singleton import SingletonMeta
from lib.log import logger

class __SQL(metaclass=SingletonMeta): 

    # Generic Execute SQL (Non-SELECT)
    def execute(db, sql):
        """
        UPDATE, INSERT, DELETE 등 결과를 반환하지 않는 SQL 실행.
        """
        logger.debug("SQL: %s", sql)
        try:
            db.Execute(sql)
        except Exception as e:
            raise RuntimeError(f"SQL Error: {sql} → {e}")

    # SELECT 실행 → 리스트(dict) 반환
    def query(db, sql):
        """
        SELECT 쿼리를 실행하고 결과를 list[dict] 형태로 반환.
        """
        logger.debug("SQL: %s", sql)
        try:
            rs = db.OpenRecordset(sql)
        except Exception as e:
            raise RuntimeError(f"SQL Error: {sql} → {e}")

        rows = []
        if rs.EOF:
            rs.Close()
            return rows

        fields = [f.Name for f in rs.Fields]

        rs.MoveFirst()
        while not rs.EOF:
            row = {f: rs.Fields(f).Value for f in fields}
            rows.append(row)
            rs.MoveNext()

        rs.Close()
        return rows



    # 단일 값 SELECT (예: COUNT)
    def query_scalar(db, sql):
        """
        SELECT COUNT(*) AS CNT FROM ... 같이 단일 값만 반환하는 쿼리.
        """
        result = query_sql(db, sql)
        if not result:
            return None

        # 첫 row의 첫 value
        return next(iter(result[0].values()))

    # 각 테이블의 Row Count 얻기
    def table_row_count(db, table_name):
        sql = f"SELECT COUNT(*) AS CNT FROM [{table_name}]"
        return query_scalar(db, sql)

    def all_table_row_counts(db):
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