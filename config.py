import os
from datetime import datetime

# 기본 경로 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
LOG_DIR = os.path.join(BASE_DIR, "logs")
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")

# URL 설정
BASE_URL = "https://www.msit.go.kr"
LIST_URL = f"{BASE_URL}/bbs/list.do?sCode=user&mPid=208&mId=307"

# 크롤링 설정
TARGET_YEAR = 2024
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"
}

# 재시도 설정
MAX_RETRIES = 5
BACKOFF_FACTOR = 1
TIMEOUT = 60

# 파일 저장 설정
TODAY_STR = datetime.now().strftime("%Y%m%d")
EXCEL_FILENAME = f"press_releases_{TODAY_STR}.xlsx"
EXCEL_PATH = os.path.join(DATA_DIR, EXCEL_FILENAME)
