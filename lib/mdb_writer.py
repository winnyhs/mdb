# mdb_writer.py
# JSON → MDB 복원기
import json
from decimal import Decimal
from lib.dao_utils import (
    get_dao_engine,
    get_field_types,
    restore_value
)


def insert_table_rows(db, table_name, rows):
    tdef = db.TableDefs(table_name)
    field_types = get_field_types(tdef)

    rs = db.OpenRecordset(table_name)

    for row in rows:
        rs.AddNew()
        for field_name, value in row.items():

            if field_name not in field_types:
                continue  # JSON에만 있고 MDB에는 없는 필드 방지

            ftype = field_types[field_name]
            v2 = restore_value(value, ftype)

            rs.Fields(field_name).Value = v2

        rs.Update()

    rs.Close()


def json_to_mdb(json_file, mdb_path, password):
    engine = get_dao_engine()
    connect = f";PWD={password}"

    db = engine.OpenDatabase(mdb_path, False, False, connect)

    # 파일 이름 = 테이블 명
    import os
    base = os.path.basename(json_file)
    tbl = base.replace(".json", "")

    with open(json_file, "r", encoding="utf-8") as f:
        rows = json.load(f)

    print(f"Inserting {len(rows)} rows into {tbl}…")
    insert_table_rows(db, tbl, rows)

    db.Close()
    print("Done writing:", tbl)
