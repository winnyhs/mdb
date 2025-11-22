import os, sys, subprocess, time, json

from lib.log import logger
from lib.dao_utils import detect_mdb_version
from lib.mdb_reader import mdb_to_json_files
from lib.mdb_writer import json_to_mdb
from lib.mdb_ddl import mdb_ddl_to_json
from lib.mdb_sync import sync_table_from_json
from lib.mdb_insert import process_prescription
from lib.sql_utils import get_dao_engine, all_table_row_counts


def inspect_table(mdb_file, password, table="M_HISTORY"):
    engine = get_dao_engine()
    connect = f";PWD={password}"
    db = engine.OpenDatabase(mdb_file, False, False, connect)

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

def main_inspect(argv): 
    mdb_file = r"E:\App\Data\VSCode\mdb\db\MEDICAL.mdb"
    password = "I1D2E3A4"
    in_dir = r"E:\App\Data\VSCode\mdb\output"
    out_dir = r"E:\App\Data\VSCode\mdb\temp"

    print(f"{mdb_file} version: {detect_mdb_version(mdb_file, password)}")
    inspect_table(mdb_file, password, "M_HISTORY")

    cmd = sys.argv[1]
    if cmd == 'read': 
        mdb_to_json_files(mdb_file, password, out_dir)
    
    elif cmd == 'write': 
        for fname in os.listdir(out_dir):
            if fname.lower().endswith(".json"):
                json_file = os.path.join(out_dir, fname)
                json_to_mdb(json_file, mdb_file, password)

    elif cmd == 'read_ddl': 
        out_json = os.path.join(out_dir, r"ddl.json")
        mdb_ddl_to_json(mdb_file, password, out_json)

    elif cmd == "insert": 
        json_file = os.path.join(in_dir, r"must-have.json")
        add_cnt, add_list, drop_cnt, drop_list = \
            process_prescription(mdb_file, password, json_file, "test-same-mm", True)
        print(f"{add_cnt} rows are added. {drip_cnt} jsons are dropped")
        print(f"added rows are: {add_list}")
        print(f"dropped rows are: {dropped_list}")

    elif cmd == "count": 
        engine = get_dao_engine()
        connect = f";PWD={password}"
        db = engine.OpenDatabase(mdb_file, False, False, connect)

        counts = all_table_row_counts(db)
        db.Close()

        json_file = os.path.join(out_dir, "count.json")
        with open(json_file, "w", encoding="utf-8") as f:
            json.dump(counts, f, indent=2, ensure_ascii=False)
        for tbl, c in counts.items():
            print(f"{tbl}: {c}")

    else:
        print(f"Invalid command: {cmd}")

def is_running(client_exe):
    try:
        out = subprocess.check_output("tasklist", shell=True)
        out = out.decode("cp949", "ignore").lower()
        return exe_name.lower() in out
    except:
        return False

def kill_processes_startswith(prefix):
    prefix = prefix.lower()

    try:
        # 현재 실행 중인 프로세스 목록 가져오기
        out = subprocess.check_output("tasklist", shell=True)
        lines = out.decode("cp949", "ignore").splitlines()
    except:
        return

    # medical.exe처럼 시작하는 모든 프로세스 종료
    for line in lines:
        parts = line.split()
        if not parts:
            continue

        exe = parts[0].lower()  # 프로세스 이름

        if exe.startswith(prefix):
            logger.info(f"{exe} is running and so is being terminated now")
            
            # 강제 종료
            for i in range(2): # try 2 times
                os.system(f"taskkill /F /IM {exe} >nul 2>nul")
                time.sleep(0.5)
                for j in range(10):   # exit in 2 seconds
                    if not is_running(exe):
                        return True
                    time.sleep(0.2)

            logger.error(f"ERROR: -----------------------------------------------------------------")
            logger.error(f"ERROR: Failed in killing medical.exe: TERMINATE medical.exe by yourself!")
            logger.error(f"ERROR: -----------------------------------------------------------------")
            return False

def mdb_insert(json_file, prog_name):
    # 1. Check the mdb file path on the client pc
    client_exe = "medical.exe"
    client_mdb = None
    for p in [r"C:\Program Files\medical", r"C:\Program Files (x86)\medical"]:
        if os.path.isdir(p):
            client_mdb = os.path.join(p, r".\MEDICAL.mdb")
            break
    if client_mdb is None:
        return {"result": "ERROR: No medical database", 
                "add_cnt": 0, "add_list": [], "drop_cnt": 0, "drop_list": []}
    
    # 2. Kill the running, if any
    kill_processes_startswith(client_exe)
    
    # 3. Start inserting new program
    add_cnt, add_list, drop_cnt, drop_list = \
                process_prescription(client_mdb, "I1D2E3A4", json_file, prog_name)
    if drop_cnt > 0: 
        result = f"ERORR: {drop_cnt} json items are dropped"
    else: 
        result = f"{add_cnt} table rows are inserted"

    return {"result": result, 
            "mdb_file" : client_mdb, 
            "prog_name" : prog_name, 
            "add_cnt": add_cnt, "add_list": add_list, 
            "drop_cnt": drop_cnt, "drop_list": drop_list}


if __name__ == "__main__":

    # Parameters
    mdb_file = os.path.abspath(r".\db\MEDICAL.mdb")
    password = "I1D2E3A4"
    in_dir = os.path.abspath(r".\output")
    out_dir = os.path.abspath(r".\temp")
    prog_name = "test"

    logger.info(f"{mdb_file} version: {detect_mdb_version(mdb_file, password)}")
    # inspect_table(mdb_file, password, "M_HISTORY")

    # 
    result = mdb_insert(os.path.join(in_dir, r"must-have.json"), prog_name)
    logger.info(f"--------- mdb file: {result['mdb_file']}")
    logger.info(f"--------- added program name: {result['prog_name']}")
    logger.info(f"--------- {result['add_cnt']} added rows are:\n{result['add_list']}")
    logger.info(f"--------- {result['drop_cnt']} dropped rows are:\n{result['drop_list']}")



    # main_inspect(sys.argv)