# mdb_insert_history.py

import json
import datetime, time
from decimal import Decimal

from lib.dao_utils import get_dao_engine
from lib.str_norm import norm



# 테이블 a 데이터를 메모리로 로딩 (SELECT 없이 DAO.OpenRecordset)
def load_table(db, table_name):
    rs = db.OpenRecordset(table_name)
    rows = []
    fields = [f.Name for f in rs.Fields]

    if not rs.EOF:
        rs.MoveFirst()
        while not rs.EOF:
            row = {f: rs.Fields(f).Value for f in fields}
            rows.append(row)
            rs.MoveNext()

    rs.Close()
    return rows

# Build index for exact match: KEY → 여러 개 row(list) 가능
def build_hash_key(table_rows, str_len):
    index = {}
    for row in table_rows:
        key = (
            norm(row.get("TYPE"), str_len),
            norm(row.get("ITEM"), str_len),
            norm(row.get("MEMO"), str_len)
        )
        index.setdefault(key, []).append(row)

    for item in index: 
        if item[0] == "오관" and item[1] == "코" and item[2].startswith("정맥동염"):
            print(f"    {item}")
    return index

# Load JSON file
def load_json(json_path):
    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)

# Exact match json item with table hash: (O(1) lookup
def exact_match(item, hash_key, prefix_len):
    key = (
        norm(item.get("cat"), prefix_len),
        norm(item.get("subcat"), prefix_len),
        norm(item.get("description"), prefix_len)
    )
    return hash_key.get(key)   # 없으면 []


# Insert new row into M_HISTORY
def insert_into_m_history(db, table_row, timestamp, prog_name) :
    """
    M_HISTORY schema:
    HTIME (datetime)
    DESP (text)
    CODE (text)
    NAME (text)
    DATA1 (currency/numeric)
    DATA2 (currency/numeric)
    GRP (text)
    VIDEO (text)
    """
    rs = db.OpenRecordset("M_HISTORY")

    rs.AddNew()
    rs.Fields("HTIME").Value = timestamp
    rs.Fields("DESP").Value = prog_name   # 처방 이름
    rs.Fields("CODE").Value = table_row.get("CODE")
    rs.Fields("NAME").Value = table_row.get("MEMO")

    # numeric → Decimal
    rs.Fields("DATA1").Value = Decimal(str(table_row.get("DATA1") or 0))
    rs.Fields("DATA2").Value = Decimal(str(table_row.get("DATA2") or 0))

    rs.Fields("GRP").Value = table_row.get("GRP")
    rs.Fields("VIDEO").Value = table_row.get("VIDEO")

    rs.Update()
    rs.Close()

# 전체 프로세스 실행
def process_prescription(mdb_path, password, json_fname, prog_name, verbose = True):
    prefix_len = 24
    table_name = "M_DATA"
    timestamp = datetime.datetime.now().replace(microsecond=0, second=0)

    # 1) DAO open
    if verbose: 
        print("Opening DAO connection...", flush=True)
        # time.sleep(0.01)
    engine = get_dao_engine()
    connect = f";PWD={password}"
    db = engine.OpenDatabase(mdb_path, False, False, connect)

    # 2) table a 읽기
    if verbose:
        print(f"Loading Table {table_name} from MEDICAL.mdb...", flush=True)
        # time.sleep(0.01)
    table_rows = load_table(db, table_name)

    if verbose: 
        print("Building hash key...", flush=True)
        # time.sleep(0.01)
    hash_key = build_hash_key(table_rows, prefix_len)  # 24 characters only

    if verbose:
        print("Loading json file...", flush=True)
        # time.sleep(0.01)
    json_items = load_json(json_fname)

    if verbose:
        print("Inserting json data into database...", flush=True)
    total_added_rows = 0   # 실제 insert된 row 개수
    skipped_json_items = 0 # json element 기준으로 실패한 개수
    added_list = []
    skipped_list = []
    for item in json_items:
        json_norm_desc = norm(item.get("description") or "", prefix_len)

        # exact match: 여러 row 가능
        matched_rows = exact_match(item, hash_key, prefix_len)
        if not matched_rows:
            skipped_list.append([item, json_norm_desc])
            skipped_json_items += 1
            print(f"ERROR: --- drop {item}")
            continue

        # 매칭된 row 전부 M_HISTORY에 insert
        print(f"--- Insert {len(matched_rows)} row(s) for {item}")

        for table_row in matched_rows:
            insert_into_m_history(db, table_row, timestamp, prog_name)
            total_added_rows += 1

            added_list.append({
                "code": table_row.get("코드"),
                "name": table_row.get("설명"),
                "cat": table_row.get("대분류"), 
                "subcat": table_row.get("소분류1"), 
                "grp": table_row.get("처방"),
                "desc_json": item.get("description"),
                "desc_a": table_row.get("설명"),
            })

    db.Close()
    
    print("\n──────────────────────────────")
    print("Completed.")
    print("    Inserted Rows      :", total_added_rows)
    print("    Dropped JSON json_items :", skipped_json_items)
    print("──────────────────────────────")

    return total_added_rows, added_list, skipped_json_items, skipped_list





# 실행
if __name__ == "__main__":
    mdb_path = r"C:\medical\MEDICAL.mdb"
    password = "xxxx"   # 실제 비밀번호 입력
    json_path = r":\json\prescription.json"

    process_prescription(mdb_path, password, json_path, "ttt")

