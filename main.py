import os, sys, json

from lib.dao_utils import detect_mdb_version
from lib.mdb_reader import mdb_to_json_files
from lib.mdb_writer import json_to_mdb
from lib.mdb_ddl import mdb_ddl_to_json
from lib.mdb_sync import sync_table_from_json
from lib.mdb_insert import process_prescription
from lib.sql_utils import get_dao_engine, all_table_row_counts


def inspect_table(mdb_path, password, table="M_HISTORY"):
    engine = get_dao_engine()
    connect = f";PWD={password}"
    db = engine.OpenDatabase(mdb_path, False, False, connect)

    tbl = db.TableDefs(table)
    print("=== Fields ===")
    for f in tbl.Fields:
        print(f"Name={f.Name}, Type={f.Type}, Size={f.Size}, Attr={f.Attributes}")

    print("\n=== Indexes ===")
    for idx in tbl.Indexes:
        print(f"Index={idx.Name}, Primary={idx.Primary}, Unique={idx.Unique}")
        for f in idx.Fields:
            print("   Field:", f.Name)

    db.Close()



if __name__ == "__main__":

    mdb_path = r"E:\App\Data\VSCode\mdb\db\MEDICAL.mdb"
    password = "I1D2E3A4"
    in_dir = r"E:\App\Data\VSCode\mdb\output"
    out_dir = r"E:\App\Data\VSCode\mdb\temp"

    print(f"{mdb_path} version: {detect_mdb_version(mdb_path, password)}")
    inspect_table(mdb_path, password, "M_HISTORY")

    cmd = sys.argv[1]
    if cmd == 'read': 
        mdb_to_json_files(mdb_path, password, out_dir)
    
    elif cmd == 'write': 
        for fname in os.listdir(out_dir):
            if fname.lower().endswith(".json"):
                json_file = os.path.join(out_dir, fname)
                json_to_mdb(json_file, mdb_path, password)

    elif cmd == 'read_ddl': 
        out_json = os.path.join(out_dir, r"ddl.json")
        mdb_ddl_to_json(mdb_path, password, out_json)

    elif cmd == "insert": 
        json_file = os.path.join(in_dir, r"must-have.json")
        add_cnt, add_list, drop_cnt, drop_list = \
            process_prescription(mdb_path, password, json_file, "test-same-mm", True)
    
    # elif cmd == 'update': # ??? Won't be used. Never mind
    #     # 예: "M_DATA" 테이블의 "description", "prescription" 필드만 업데이트
    #     table_name = "M_DATA"
    #     json_file = os.path.join(out_dir, "M_DATA.json")
    #     update_fields = ["description", "prescription"]  # 수정하려는 필드 목록
    #     sync_table_from_json(mdb_path, password, table_name, json_file, update_fields)

    elif cmd == "count": 
        engine = get_dao_engine()
        connect = f";PWD={password}"
        db = engine.OpenDatabase(mdb_path, False, False, connect)

        counts = all_table_row_counts(db)
        db.Close()

        json_file = os.path.join(out_dir, "count.json")
        with open(json_file, "w", encoding="utf-8") as f:
            json.dump(counts, f, indent=2, ensure_ascii=False)
        for tbl, c in counts.items():
            print(f"{tbl}: {c}")

    else:
        print(f"Invalid command: {cmd}")