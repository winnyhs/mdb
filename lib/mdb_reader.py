# mdb_reader.py
# - MDB → JSON 변환기
import os
import json
from lib.dao_utils import (
    get_dao_engine,
    list_tables,
    normalize_value
)


def read_table_to_list(db, table_name):
    rs = db.OpenRecordset(table_name)
    result = []

    fields = [f.Name for f in rs.Fields]

    if rs.EOF:
        rs.Close()
        return result

    rs.MoveFirst()
    while not rs.EOF:
        row = {}
        for f in fields:
            v = rs.Fields(f).Value
            row[f] = normalize_value(v)
        result.append(row)
        rs.MoveNext()

    rs.Close()
    return result


def mdb_to_json_files(mdb_path, password, out_dir):
    """테이블별 JSON 파일로 저장"""
    engine = get_dao_engine()
    connect = f";PWD={password}"
    db = engine.OpenDatabase(mdb_path, False, False, connect)

    tables = list_tables(db)
    print("Tables:", tables)

    os.makedirs(out_dir, exist_ok=True)

    for tbl in tables:
        print(f"Reading table: {tbl}")

        data = read_table_to_list(db, tbl)

        fname = os.path.join(out_dir, f"{tbl}.json")
        with open(fname, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        print("Saved:", fname)

    db.Close()
    print("Done.")
