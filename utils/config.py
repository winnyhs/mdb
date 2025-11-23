import os
from lib.singleton import SingletonMeta
from lib.log import logger
from utils.sys import choose_external_drive_name

class __Config(metaclass = SingletonMeta): 
    def __init__(self): 
        self.exe_file = "medical.exe"
        self.mdb_file = "MEDICAL.mdb"
        self.password = "I1D2E3A4"
        
        self.program_json_file = "program.json" # export of M_HISTORY or build from must-have.py
        self.must_have_json_file = "must-have.json"
        self.good_to_have_json_file = "good_to_have.json"
        self.virus_json_file = "virun.json"

        self.sys_drv = {
            "top_dir_candidate" : [ r"C:\Program Files\medical", 
                                    r"C:\Program Files (x86)\medical", 
                                    r"C:\medical" ], 
            "top_dir" : None, 
            "mdb_path" : None, 
            "exe_path" : None,
        }
        
        '''External drive directory structure (like USB)
            E:\medical\
                +--- bin
                +--- json
                +--- mdb
        '''
        self.ext_drv = {  # need to be reconfigured
            "name": None, # ex) "E", 
            "top_dir" : r"medical",  
            "bin_dir" : r"bin", 
            "json_dir" : r"json", 
            "mdb_dir" : r"mdb", 
            "json_path": None, 
            "mdb_path": None
        }
    
    def configure(self): 
        # 1. Find the internal directory
        for p in self.sys_drv["top_dir_candidate"]:
            if os.path.isdir(p):
                if os.path.isfile(os.path.join(p, self.mdb_file)): 
                    self.sys_drv["top_dir"] = p
                    self.sys_drv["mdb_path"] = os.path.join(p, self.mdb_file)
                    self.sys_drv["exe_path"] = os.path.join(p, self.exe_file)
                    break
        if self.sys_drv["top_dir"] is None: 
            return False 
        
        # 2. Find the external directory
        # Raw string literal means backslashes do not introduce escape sequences.
        # They are included in the string exactly as typed, except that
        # a raw string cannot end with a single backslash.
        self.ext_drv["name"] = choose_external_drive_name() + ":\\"
        self.ext_drv["top_dir"]  = self.ext_drv["name"] + self.ext_drv["top_dir"]
        self.ext_drv["bin_dir"]  = os.path.join(self.ext_drv["top_dir"], self.ext_drv["bin_dir"])
        self.ext_drv["json_dir"] = os.path.join(self.ext_drv["top_dir"], self.ext_drv["json_dir"])
        self.ext_drv["mdb_dir"]  = os.path.join(self.ext_drv["top_dir"], self.ext_drv["mdb_dir"])
        self.ext_drv["json_path"] = os.path.join(self.ext_drv["json_dir"], self.program_json_file)
        self.ext_drv["mdb_path"] = os.path.join(self.ext_drv["mdb_dir"], self.mdb_file)

        logger.info("--------------------\n config: ")
        logger.info("sys_drv: %s", self.sys_drv)
        logger.info("ext_drv: %s", self.ext_drv)
        logger.info("--------------------")
        return self

    def clean_ext_drive():
        top_dir  = self.ext_drv["top_dir"]
        bin_dir  = self.ext_drv["bin_dir"]
        json_dir = self.ext_drv["json_dir"]
        mdb_dir  = self.ext_drv["mdb_dir"]

        if not os.path.isdir(top_dir):
            os.makedirs(top_dir)
            os.makedirs(bin_dir)
            os.makedirs(json_dir)
            os.makedirs(mdb_dir)
            logger.error("ERROR: %s doesn't have exe files", bin_dir)
            return False

        # Delete all under json and mdb folder 
        for d in (json_dir, mdb_dir): 
            try: 
                shutil.rmtree(d)
                logger.info("%s is deleted", d)
            except Exception as e:
                logger.exception("%s: Failed in deleting folders", e)
                return False

        for d in (json_dir, mdb_dir):
            try:
                os.makedirs(d)
                logger.info("%s is created", d)
            except Exception as e:
                logger.exception("%s: Failed in creating folders", e)
                return False
        return True
    

Config = __Config()
