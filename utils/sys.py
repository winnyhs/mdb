# -*- coding: utf-8 -*-
import ctypes, os

from lib.log import logger

# ------------------------
# Process
# ------------------------

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


# ------------------------
# path type
# ------------------------
def path_type(path):
    is_abs = os.path.isabs(path)    # 절대경로인지?
    dirname = os.path.dirname(path) # 디렉토리 부분 추출
    base = os.path.basename(path)   # basename (파일명만인지 여부)

    if is_abs:
        return "absolute"
    else:
        # absolute가 아닌 경우: relative or file name only
        if dirname == "":
            return "file_name_only"
        else:
            return "relative"


# ------------------------
# External Drives
# ------------------------

# Windows DriveType Constants
DRIVE_UNKNOWN = 0
DRIVE_NO_ROOT_DIR = 1
DRIVE_REMOVABLE = 2   # USB / SD Card
DRIVE_FIXED = 3       # 내장 HDD / 외장 HDD
DRIVE_REMOTE = 4
DRIVE_CDROM = 5
DRIVE_RAMDISK = 6

def get_drive_label(drive_letter):
    kernel32 = ctypes.windll.kernel32

    volume_name_buf = ctypes.create_unicode_buffer(256)
    fs_name_buf = ctypes.create_unicode_buffer(256)

    serial_number = ctypes.c_uint(0)
    max_component_len = ctypes.c_uint(0)
    file_system_flags = ctypes.c_uint(0)

    root_path = drive_letter + ":\\"

    res = kernel32.GetVolumeInformationW(
        ctypes.c_wchar_p(root_path),
        volume_name_buf,
        ctypes.sizeof(volume_name_buf),
        ctypes.byref(serial_number),
        ctypes.byref(max_component_len),
        ctypes.byref(file_system_flags),
        fs_name_buf,
        ctypes.sizeof(fs_name_buf)
    )

    if res == 0:
        return ""
    return volume_name_buf.value

def list_removable_drive_labels():
    """USB / SD Card 만 출력"""
    kernel32 = ctypes.windll.kernel32
    labels = {}

    bitmask = kernel32.GetLogicalDrives()

    for i in range(26):
        if bitmask & (1 << i):
            d = chr(ord('A') + i)
            drive_type = kernel32.GetDriveTypeW(u"%s:\\" % d)

            if drive_type == DRIVE_REMOVABLE:
                # USB / SD card 만 필터링
                labels[d] = get_drive_label(d)

    return labels

def choose_external_drive_name(): # The earliest detected one
    labels = list_removable_drive_labels()
    selected = "Z"
    for drive, label in labels.items():
        if drive.lower() < selected.lower(): 
            selected = drive

    print("selected drive: %s" % (drive)) # E
    return selected
    
# -----------------------
# 사용 예
# -----------------------
if __name__ == "__main__":

    # -----
    labels = list_removable_drive_labels()
    for drive, label in labels.items():
        print("%s: %s" % (drive, label)) # E: SONY

    # -----
    paths = [
        r"C:\temp\data.txt",
        r"E:data.txt",
        r"temp\data.txt",
        r"data.txt",
        r".\data.txt",
        r"..\config\set.ini"
    ]
    for p in paths:
        print(p, "→", classify_path(p))
