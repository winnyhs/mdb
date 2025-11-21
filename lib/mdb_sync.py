# mdb_sync.py
import json
from lib.dao_utils import (
    get_dao_engine,
    get_field_types,
    get_primary_index,
    restore_value,
)


def sync_table_from_json(mdb_path, password, table_name, json_file, update_fields):
    """
    JSON 파일을 읽어서 해당 테이블의 특정 필드만 업데이트.
    - PK(Primary Index) 기준으로 매칭
    - update_fields 목록에 있는 필드만 수정
    - JSON에 있는 row 중 PK가 없는 경우 → 무시
    """
    engine = get_dao_engine()
    connect = f";PWD={password}"
    db = engine.OpenDatabase(mdb_path, False, False, connect)

    tdef = db.TableDefs(table_name)
    field_types = get_field_types(tdef)
    pk_index_name, pk_fields = get_primary_index(tdef)

    if not pk_fields:
        raise RuntimeError(f"Table '{table_name}' has no primary index. Can't sync safely.")

    print(f"Table: {table_name}")
    print(f"Primary index: {pk_index_name}, fields: {pk_fields}")
    print(f"Update fields: {update_fields}")

    # JSON 읽기
    with open(json_file, "r", encoding="utf-8") as f:
        rows = json.load(f)

    # DAO: Table type Recordset을 열고, Index + Seek 사용
    # dbOpenTable = 1
    rs = db.OpenRecordset(table_name, 1)
    rs.Index = pk_index_name

    updated_count = 0
    skipped_no_pk = 0
    skipped_no_match = 0

    for row in rows:
        # PK 값 준비
        try:
            key_values = [row[pk] for pk in pk_fields]
        except KeyError:
            skipped_no_pk += 1
            continue

        # Seek("=" , pk1, pk2, ...)
        # 한 개 PK: rs.Seek("=", key)
        # 여러 개 PK: rs.Seek("=", key1, key2, ...)
        args = ["="] + key_values
        rs.Seek(*args)

        if rs.NoMatch:
            skipped_no_match += 1
            continue

        # 매칭된 레코드의 update_fields만 수정
        for fname in update_fields:
            if fname not in row:
                continue
            if fname not in field_types:
                continue

            ftype = field_types[fname]
            v = restore_value(row[fname], ftype)
            rs.Fields(fname).Value = v

        rs.Update()
        updated_count += 1

    rs.Close()
    db.Close()

    print(f"Updated rows: {updated_count}")
    print(f"Skipped (no pk in JSON): {skipped_no_pk}")
    print(f"Skipped (no match in DB): {skipped_no_match}")
