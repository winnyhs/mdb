# mdb_sync.py
import json, os, sys, shutil, datetime
import unicodedata, re

from lib.singleton import SingletonMeta
from lib.log import logger
from utils.json import save_json, load_json
from utils.sql import Sql
from utils.db_engine import DbEngine
from utils.config import Config

class MedicalDB(DbEngine): 
    def __init__(self, mdb_path, password = ""):
        super().__init__()
        self.mdb_path = mdb_path
        self.password = password 
        self.sql = Sql
    

class Program(metaclass = SingletonMeta): 
    def __init__(self, config): 
        self.data_table = "M_DATA"
        self.program_table = "M_HISTORY"
        self.cfg = config

        # This instance is to update MEDICAL.mdb in the system
        self.sys_db_engine = MedicalDB(config.sys_drv["mdb_path"], config.password)
        self.ext_db_engine = MedicalDB(config.ext_drv["mdb_path"], config.password)
        for engine in [self.sys_db_engine, self.ext_db_engine]: 
            db = engine.open_db()
            engine.detect_version(db)
            db.Close()

        self.prefix_len = 24  # cat2's string length to compare
        self.htime = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        # self.htime = 

    # Remove all white spaces and keep the prefix of <len> length
    def str_normalize(self, s, len):
        # Normalize Korean/English text for matching while preserving () , [] {}.
        if not s:
            return ""

        # 1) Unicode Normalize
        s = unicodedata.normalize("NFKC", s)

        # 2) Lowercase (영문만 영향)
        s = s.lower()

        # 3) Strip leading/trailing spaces
        s = s.strip()

        # 4) Remove ALL white spaces (space/tabs/newlines)
        s = re.sub(r"[ \t\r\n]+", "", s)

        # 5) Remove punctuation EXCEPT (), [], {}, and ,
        #    we remove: . , ; : ! ? ' " - _ / \ |
        s = re.sub(r"[.;:!?'\"/_\\\-|]+", "", s)

        return s[:len]

        # TODO: 그냥 matching 하는 것과 비교해 보기. 맨 앞 20글자만 matching 했을 때와도 비교하기
        # 개수를 봐서 차이가 있는지 없는지 확인하기 
        # TODO: table a, b, c, M_LIST, M_DATA 상호 비교해서 차이 정리하기 
    
    # Build the hash for exact match: key = (TYPE, ITEM, MEMO), value = row
    def build_hash(self, table_rows, str_len):
        index = {}
        for row in table_rows:
            key = (
                self.str_normalize(row.get("TYPE"), str_len),
                self.str_normalize(row.get("ITEM"), str_len),
                self.str_normalize(row.get("MEMO"), str_len)
            )
            index.setdefault(key, []).append(row)

        # for testing
        for item in index: 
            if item[0] == "오관" and item[1] == "코" and item[2].startswith("정맥동염"):
                print(f"    {item}")
        return index

    # Exact match json item with table hash: (O(1) lookup
    def exact_match(self, item, hash_key):
        key = (
            self.sys_db_engine.str_normalize(item.get("cat"),         self.prefix_len),
            self.sys_db_engine.str_normalize(item.get("subcat"),      self.prefix_len),
            self.sys_db_engine.str_normalize(item.get("description"), self.prefix_len)
        )
        return hash_key.get(key)   # 없으면 []

    # Insert new row into M_HISTORY
    def insert_1row(self, db, table_row, htime, prog_name) :
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
        rs.Fields("HTIME").Value = htime      # 처방전 key
        rs.Fields("DESP").Value = prog_name   # 처방전 이름; 처방전 key
        rs.Fields("CODE").Value = table_row.get("CODE")
        rs.Fields("NAME").Value = table_row.get("MEMO")

        # numeric → Decimal
        rs.Fields("DATA1").Value = Decimal(str(table_row.get("DATA1") or 0))
        rs.Fields("DATA2").Value = Decimal(str(table_row.get("DATA2") or 0))

        rs.Fields("GRP").Value = table_row.get("GRP")
        rs.Fields("VIDEO").Value = table_row.get("VIDEO")

        rs.Update()
        rs.Close()

    # Build a program for the given prescription and then, insert it into sys db
    def build_program_from_anlaysis_result(self, db, analysis_json_file, prog_name):
        # 1) Read table M_DATA
        logger.info(f"Loading Table {self.data_table} from MEDICAL.mdb...") #, flush=True)
        table_rows = self.sys_db_engine.read_table(db, self.data_table)

        # 1.2) Build a hash table for table M_DATA
        logger.info("Building hash table for fast exact matching...") # , flush=True)
        hash_key = self.build_hash(table_rows, self.prefix_len)  # 24 characters only

        # 2) Read analysis result data, that is new prescription to add into M_HISTORY
        logger.info("Loading analysis result json file...") #, flush=True)
        json_path = os.path.join(self.cfg.ext_drv["json_dir"], analysis_json_file)
        json_items = load_json(json_path)

        # 3) Convert analysis json data into a program json data
        logger.info("Converting analysis result data into a program json...") #, flush=True)
        added_row_cnt = 0   # 실제 insert된 row 개수
        skipped_json_cnt = 0 # json element 기준으로 실패한 개수
        added_list = []
        skipped_list = []
        for item in json_items:
            # 3.1) Normalize a json element: cat2 string 
            norm_desc = self.str_normalize(item.get("description") or "", self.prefix_len)

            # 3.2) Do exact match of a json element and M_DATA table 
            #    and get rows in M_DATA table
            matched_rows = self.exact_match(item, hash_key)
            if not matched_rows:
                skipped_list.append([item, norm_desc])
                skipped_json_cnt += 1
                logger.error(f"ERROR: --- drop {item}: norm_desc={norm_desc}")
                continue

            # 6) Add the matched rows into M_HISTORY table
            logger.info(f"--- Insert {len(matched_rows)} row(s) for {item}")
            for table_row in matched_rows:
                added_row_cnt += 1
                added_list.append({
                    "HTIME": self.htime, 
                    "DESP": prog_name, 
                    "CODE": table_row.get("CODE"),
                    "NAME": table_row.get("MEMO"),  
                    "DATA1": table_row.get("DATA1"),
                    "DATA2": table_row.get("DATA2"),
                    "GRP": table_row.get("GRP"),
                    "VIDEO": table_row.get("VIDEO")
                })
        
        logger.info("")
        logger.info("───────────────────────────────────")
        logger.info("Completed.")
        logger.info("    Analysis result items: ", len(a_items))
        logger.info("        * Dropped items  : ", skipped_json_cnt)
        logger.info("    Program json items   : ", added_row_cnt)
        logger.info("───────────────────────────────────")

        # 4) Export the program json as json file
        save_json(added_list, self.cfg.ext_drv["json_path"])
        return added_row_cnt, added_list, skipped_json_cnt, skipped_list

    # Insert a program json file into sys db
    def insert_program(self, db): 
        # 1. Load the program json
        prog_json = load_json(self.ext_drv["json_path"])

        # 2. Insert the program 
        # 2.1. Start transaction (performance and data integrity)
        db.BeginTrans()
        try:
            for row in prog_json:
                # 2.2. Convert a row to two lists: cols, vals
                cols = []
                vals = []
                for k, v in row.items():
                    cols.append(k)
                    v = normalize_value(v)
                    vals.append(v)

                # 2.3. Build SQL: INSERT INTO M_HISTORY (col1, col2) VALUES ('v1', 'v2')
                col_list = ", ".join(cols)

                # Convert value (문자열은 따옴표, 숫자는 그대로)
                val_list = []
                for v in vals:
                    if v is None:
                        val_list.append("NULL")
                    elif isinstance(v, (int, float)):
                        val_list.append(str(v))
                    else: # string: add escape
                        s = str(v).replace("'", "''")
                        val_list.append("'" + s + "'")
                val_expr = ", ".join(val_list)

                sql = "INSERT INTO %s (%s) VALUES (%s)" % (
                    self.program_table, col_list, val_expr
                )

                # 2.4. Execute the SQL
                db.Execute(sql)

            # 3. Commit the transaction
            db.CommitTrans()

        except Exception as e:
            # 3.1. Rollback the transaction
            logger.exception("Error: %s: Rollback", e)
            db.Rollback()
        # finally:
        #     db.Close()

        logger.info("Inserted rows: ", len(prog_json))

    def export_program(self, db): 
        data = self.sys_db_engine.read_table(db, self.program_table)
        save_json(data, self.cfg.ext_drv["json_path"])
    
    def import_progam(self, db, only_diff = True):
        if onlY_diff == True: 
            print("NOT IMPLEMENTED YET")
        else: 
            insert_progam(db)


if __name__ == "__main__": 

    c = Config.configure()
    p = Program(c)

    # 1. Load must-have.json
    # 2. Build the program from the analysis result json
    db = p.sys_db_engine.open_db()
    p.build_program_from_anlaysis_result(db, "must-have.json", "99test")
    print(os.listdir(c.ext_drv["json_dir"]))

    # 3. Insert the program into the system mdb
    p.insert_program(db)

    # 4. Export the program table (M_HISTORY) to the external driver
    p.export_program(db)
    db.Close()
    print(os.listdir(c.ext_drv["json_dir"]))

    # --- DO delete some program from the system mdb --- 
    input("Enter when ready: ")

    # 5. Import the programs that are not in the system mdb
    db = p.sys_db_engine.open_db()
    p.import_program(db)
    db.Close()


