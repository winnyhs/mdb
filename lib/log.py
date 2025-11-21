# log.py
import logging
import os

# 로그 포맷 설정
LOG_FORMAT = "%(asctime)s [%(filename)s:%(lineno)d] %(message)s"
# LOG_FORMAT = "%(asctime)s [%(levelname)s] [%(filename)s:%(lineno)d] %(message)s"
# LOG_FORMAT = "[%(filename)s:%(lineno)d] %(message)s"

# 로그 파일 경로 (필요시)
# LOG_FILE = os.path.join(os.path.dirname(__file__), "app.log")

# 기본 설정 (콘솔 + 파일)
logging.basicConfig(
    level=logging.DEBUG,                # DEBUG / INFO / WARNING / ERROR / CRITICAL
    format=LOG_FORMAT,
    handlers=[
        logging.StreamHandler(),       # 콘솔 출력
        # logging.FileHandler(LOG_FILE, encoding="utf-8")  # 파일 저장
    ]
)

# 모듈 단위 logger 제공 (각 파일에서 import 해 사용)
logger = logging.getLogger(__name__)
