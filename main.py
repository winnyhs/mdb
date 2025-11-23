import os, sys, subprocess, shutil, time, json

from lib.log import logger
from utils.medical import Program
from utils.config import Config

config = Config()
program = Program(config)


def copy_mdb_to_external(): 
    # copy PC's MEDICAL.mdb file into external drive's pre-defined directory
    src = config["int_drive"]["mdb_path"]
    dst = config["ext_drive"]["mdb_path"]
    try:
        shutil.copy2(src, dst)
        logger.info("%s is copied as %s", src, dst)
    except Exception as e:
        logger.exception("%s: ERROR: Copy failed: %s -> %s", e, src, dst)


def mdb_import_json(): 
    # Import json-exported row data into MEDICAL.mdb
    # - If the program name of the imported is in use in MEDICAL.mdb, 
    #   it is suffixed with "_1"

    # 1. Check that the given json_file is there 
    json_path = os.path.join(config["ext_drive"]["json_dir"], config["json_file"])
    return {"result": "ERROR: No such json file: {}".format(json_path), 
            "add_cnt" : add_cnt, "add_list" : add_list, 
            "drop_cnt" : drop_cnt, "drop_list": drop_list}
    rows = load_json(json_path)

    # 2. Check that the given prog_name can be used in the new mdb
    prog_name_list = []
    for r in rows: 
        if r["DESC"] not in prog_name_list: 
            prog_name_list.append()
    
    for p in prog_name_list: 

    return {"result": result, 
            "add_cnt" : add_cnt, "add_list" : add_list, 
            "drop_cnt" : drop_cnt, "drop_list": drop_list}

    # 3. Import the json elements into the mdb
    return mdb_insert_json(json_file, prog_name)

# Insert  
def insert_analysis_json(self, json_file, prog_name):
    
    # 1. Kill the running, if any
    # kill_processes_startswith(config["exe_file"])
    
    # 2. Start inserting new program
    json_path = os.path.join(config["ext_drive"]["json_dir"], json_file)
    add_cnt, add_list, drop_cnt, drop_list = \
        process_prescription(config["int_drive"]["mdb_path"], 
                            config["password"], 
                            json_Path, 
                            prog_name)
    if drop_cnt > 0: 
        result = f"ERORR: {drop_cnt} json items are dropped"
    else: 
        result = f"{add_cnt} table rows are inserted"

    return {"result" : result, 
            "mdb_file" : config["int_drive"]["mdb_path"], 
            "json_file" : json_path, 
            "prog_name" : prog_name, 
            "add_cnt" : add_cnt, "add_list": add_list, 
            "drop_cnt" : drop_cnt, "drop_list": drop_list}

