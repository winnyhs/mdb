# mdb_insert_history.py

import json
import datetime, time
from decimal import Decimal

from lib.dao_utils import get_dao_engine
from lib.str_norm import norm
from lib.log import logger



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
def load_json(json_file):
    with open(json_file, "r", encoding="utf-8") as f:
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
def insert_into_m_history(db, table_row, htime, prog_name) :
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
    rs.Fields("HTIME").Value = htime
    rs.Fields("DESP").Value = prog_name   # 처방전 이름
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
def process_prescription(mdb_file, password, json_file, prog_name):
    prefix_len = 24
    table_name = "M_DATA"
    htime = datetime.datetime.now().replace(microsecond=0, second=0)

    # 1) DAO open
    logger.info("Opening DAO connection...") #, flush=True)
    engine = get_dao_engine()
    connect = f";PWD={password}"
    db = engine.OpenDatabase(mdb_file, False, False, connect)

    # 2) table M_DATA 읽기
    logger.info(f"Loading Table {table_name} from MEDICAL.mdb...") #, flush=True)
    table_rows = load_table(db, table_name)

    logger.info("Building hash key...") # , flush=True)
    hash_key = build_hash_key(table_rows, prefix_len)  # 24 characters only

    logger.info("Loading json file...") #, flush=True)
    json_items = load_json(json_file)

    logger.info("Inserting json data into database...") #, flush=True)
    added_row_cnt = 0   # 실제 insert된 row 개수
    skipped_json_cnt = 0 # json element 기준으로 실패한 개수
    added_list = []
    skipped_list = []
    for item in json_items:
        json_norm_desc = norm(item.get("description") or "", prefix_len)

        # exact match: 여러 row 가능
        matched_rows = exact_match(item, hash_key, prefix_len)
        if not matched_rows:
            skipped_list.append([item, json_norm_desc])
            skipped_json_cnt += 1
            logger.error(f"ERROR: --- drop {item}: norm_desc={json_norm_desc}")
            continue

        # 매칭된 row 전부 M_HISTORY에 insert
        logger.info(f"--- Insert {len(matched_rows)} row(s) for {item}")

        for table_row in matched_rows:
            insert_into_m_history(db, table_row, htime, prog_name)
            added_row_cnt += 1

            added_list.append({
                "code": table_row.get("CODE"),
                "cat": table_row.get("TYPE"), 
                "subcat": table_row.get("ITEM"),
                "frequency": table_row.get("DATA1"),  
                "grp": table_row.get("GRP"),
                "name": table_row.get("MEMO"),

                "json_desc": item.get("description"),
            })

    db.Close()
    
    logger.info("")
    logger.info("──────────────────────────────")
    logger.info("Completed.")
    logger.info(f"    Inserted table rows : {added_row_cnt}")
    logger.info(f"    Dropped json items  : {skipped_json_cnt}")
    logger.info("──────────────────────────────")

    return added_row_cnt, added_list, skipped_json_cnt, skipped_list

