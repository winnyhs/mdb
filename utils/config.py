import os, shutil
from lib.singleton import SingletonMeta
from lib.log import logger
from utils.sys import choose_external_drive_name

class __Config(metaclass = SingletonMeta): 
    def __init__(self): 
        self.exe_file = "medical.exe"
        self.mdb_file = "MEDICAL.mdb"
        self.password = "I1D2E3A4"
        

        self.data_table_file = "data_table.json" # export of M_DATA table
        self.program_table_file = "program_table.json" # the export of M_HISTORY table 

        self.program_file = "program.json"       # the final level result of auto_analyzer

        self.must_have_file = "must-have.json"
        self.good_to_have_file = "good_to_have.json"
        self.virus_file = "virun.json"

        self.sys_drv = {
            "top_dir_candidate" : [ r"C:\Program Files\medical", 
                                    r"C:\Program Files (x86)\medical", 
                                    r"C:\medical" ], 
            "top_dir" : None, 
            "mdb_path" : None, 
            "exe_path" : None,
        }

        '''In case of running on E: Drive
        E:\medical\
            +-- auto_analyzer
            |   +-- config
            |   +-- db
            |   +-- temp
            |       +-- json
            |       +-- mdb
            +-- client
            +-- install_package
        '''
        ''' 
        C:\Program Files (32)\medical
            +-- medical.exe
            +-- MEDICAL.mdb
        C:\Python_38
        C:\Python_venv
        '''
        '''In case of running on C: Drive
        C:\Program Files (32)\auto_analyzer
            +-- mdb_copier.exe  # Run on the clients PC. 
            |                   # Copy(delete and then copy) its MEDICAL.mdb 
            |                   # under E:\medical\mdb
            +-- auto_analyzer.exe
            +-- config
            |   +-- config.py
            +-- db
            |   +-- data_table.json   # Exort of M_DATA table
            +-- temp
                +-- json
                |   +-- program.json       # program, final level analysis result from auto_analyzer
                |   +-- program_table.json # export of M_HISTORY table
                |   +-- must-have.json     # level-1 analysis result from auto_analyzer
                |   +-- good-to-have.json, virus.json, ...
                |
                +-- mdb
                    +-- MEDICAL.mdb

        E:\medical\
            +-- install_package
            |   +-- auto_analyzer_install.exe # To install python 3.8.10, python 3.7, pywin32, ...
            |   +-- auto_start_install.exe    # To install the automatic start of medical.exe
            |   +-- MDBPlus.exe               # To read a mdb file and to run SQL on it
            |
            +-- client
                +-- 김나라
                |   +-- 김나라_profile.json
                |   +-- 김나라_memo.txt
                |   +-- 2025-11-11T10-00
                |   |   +-- html
                |   |   |   +-- image
                |   |   +-- json
                |   |       +-- must-have.json, good-to-have.json, virus.json, ...
                |   +-- 2025-12-01T10-00
                |
                +-- 배나라 
        '''
        self.run_drv = {
            "name": os.path.splitdrive(os.getcwd())[0] + "\\", # ex) "E:\\" for "E:\"
            "top_dir": os.getcwd(),  # os.path.dirname(os.getcwd())
        }
        # Export of M_DATA table
        self.run_drv["data_table_path"] = os.path.join(self.run_drv["top_dir"], r"db\data_table.json") 
        self.run_drv["ddl_path"]        = os.path.join(self.run_drv["top_dir"], r"db\ddl.json") 
        # Storage for json files
        self.run_drv["json_dir"]     = os.path.join(self.run_drv["top_dir"], r"temp\json")
        self.run_drv["program_path"] = os.path.join(self.run_drv["json_dir"], r"program.json")

        self.ext_drv = {  # need to be reconfigured
            # Python doc says: 
            #   - Raw string literal means backslashes do not introduce escape sequences.
            #     They are included in the string exactly as typed, except that
            #     a raw string cannot end with a single backslash.
            "name": choose_external_drive_name() + ":\\", # ex) "E:\\" for "E:\"
        }
        self.ext_drv["top_dir"]    = self.ext_drv["name"] + r"medical" 
        self.ext_drv["client_dir"] = os.path.join(self.ext_drv["top_dir"], r"client") 
        self.ext_drv["mdb_dir"]    = os.path.join(self.ext_drv["top_dir"], r"temp\mdb")
        self.ext_drv["mdb_path"]   = os.path.join(self.ext_drv["top_dir"], r"temp\mdb\MEDICAL.mdb")
    
    def configure(self, sys_drv_top_dir = None, ext_drv_name = None): 
        # 1. Configure the internal directory, if not provided under self.sysdrv["top_dir"]
        if sys_drv_top_dir == None: # self.sys_drv["top_dir"] is None: 
            for p in self.sys_drv["top_dir_candidate"]:
                if os.path.isdir(p):
                    if os.path.isfile(os.path.join(p, self.mdb_file)): 
                        self.sys_drv["top_dir"] = p
                        self.sys_drv["mdb_path"] = os.path.join(p, self.mdb_file)
                        self.sys_drv["exe_path"] = os.path.join(p, self.exe_file)
                        break
            if self.sys_drv["top_dir"] is None: 
                logger.error("ERROR: No medical program is installed")
                return None
        else: 
            if not os.path.isdir(sys_drv_top_dir) or \
                not os.path.isfile(os.path.join(sys_drv_top_dir, self.mdb_file)):
                logger.error("ERROR: No medical program is installed under %s", sys_drv_top_dir)
                return None 
        
        # 2. Reconfigure the external drive directory
        if ext_drv_name is not None: 
            if self.ext_drv["name"][0].upper() != ext_drv_name[0].upper(): 
                for i in self.ext_drv: 
                    self.ext_drv[i][0] = ext_drv_name[0]

        # self.clean_temp()
        logger.info("------------------------------------------------")
        logger.info("config: ")
        logger.info("    sys_drv: %s", self.sys_drv["top_dir"])
        logger.info("    run_drv: %s", self.run_drv["top_dir"])
        logger.info("    ext_drv: %s", self.ext_drv["top_dir"])
        logger.info("            sys mdb_path: %s", self.sys_drv["mdb_path"])
        logger.info("            ext mdb_path: %s", self.ext_drv["mdb_path"])
        logger.info("------------------------------------------------")
        return self

    def clean_temp(self):
        top_dir  = self.run_drv["top_dir"]
        json_dir = self.run_drv["json_dir"]
        mdb_dir  = self.ext_drv["mdb_dir"]

        # Delete all under json and mdb folder 
        for d in (json_dir, mdb_dir): 
            try: 
                shutil.rmtree(d)
                logger.info("%s is deleted", d)
            except Exception as e:
                logger.exception("ERROR: %s: Failed in deleting folders", e)
                return False

        for d in (json_dir, mdb_dir):
            try:
                os.makedirs(d)
                logger.info("%s is created", d)
            except Exception as e:
                logger.exception("ERROR: %s: Failed in creating folders", e)
                return False
        return True
    

Config = __Config()
