# mdb_ddl_reader.py
import json
import os
from lib.dao_utils import get_dao_engine, list_tables


def extract_table_ddl(tabledef):
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


def mdb_ddl_to_json(mdb_path, password, out_json):
    """
    MDB 전체 테이블 구조(DDL에 대응하는 메타데이터)를 하나의 JSON으로 저장.
    """
    engine = get_dao_engine()
    connect = f";PWD={password}"
    db = engine.OpenDatabase(mdb_path, False, False, connect)

    tables = list_tables(db)
    print("Tables:", tables)

    ddl = {}

    for tbl in tables:
        tdef = db.TableDefs(tbl)
        ddl[tbl] = extract_table_ddl(tdef)

    db.Close()

    os.makedirs(os.path.dirname(out_json) or ".", exist_ok=True)
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(ddl, f, indent=2, ensure_ascii=False)

    print("DDL JSON saved:", out_json)
