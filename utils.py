import re
from datetime import datetime

def clean_text(text):
    """텍스트 공백 정리"""
    if not text:
        return ""
    return re.sub(r'\s+', ' ', text).strip()

def parse_date(date_str):
    """
    날짜 문자열 파싱
    지원 형식: YYYY.MM.DD, YYYY-MM-DD
    """
    if not date_str:
        return None
        
    date_str = date_str.strip()
    
    formats = ["%Y.%m.%d", "%Y-%m-%d"]
    
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
            
    return None

def normalize_date(date_str: str) -> str:
    """
    다양한 날짜 포맷을 YYYY-MM-DD로 표준화합니다.
    
    지원 형식:
    - Feb 6, 2026 (영어)
    - 2026. 2. 6 (한국어 공백 포함)
    - 2026.02.06 (한국어 공백 없음)
    - 2026-02-06 (이미 표준)
    
    Returns:
        str: YYYY-MM-DD 형식 문자열. 파싱 실패 시 원본 반환.
    """
    if not date_str:
        return ""
    
    date_str = date_str.strip().rstrip('.')
    
    # 이미 표준 형식인 경우 바로 반환
    if re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        return date_str
    
    # 영어 형식: Feb 6, 2026
    if ',' in date_str:
        months = {
            'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
            'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
        }
        parts = date_str.replace(',', '').split()
        if len(parts) == 3:
            month = months.get(parts[0], '01')
            day = parts[1].zfill(2)
            year = parts[2]
            return f"{year}-{month}-{day}"
    
    # 한국어/점 형식: 2026. 2. 6 또는 2026.02.06
    if '.' in date_str:
        nums = re.findall(r'\d+', date_str)
        if len(nums) >= 3:
            year = nums[0]
            month = nums[1].zfill(2)
            day = nums[2].zfill(2)
            return f"{year}-{month}-{day}"
    
    # 파싱 실패 시 원본 반환
    return date_str

def summarize_text(text, num_sentences=3):
    """규칙 기반 요약: 본문의 첫 N개 문장 추출 (길이 체크 포함)"""
    if not text:
        return ""
    
    # 문장 분리 (단순하게 . ? ! 등으로 분리)
    sentences = re.split(r'(?<=[.?!])\s+', text)
    
    # 너무 짧은 문장은 제외 (예: 제목의 일부 등)
    # 20자 이상인 문장만 유효하다고 판단
    valid_sentences = [s for s in sentences if len(s.strip()) > 20]
    
    # 유효 문장이 없으면 그냥 앞부분 반환
    if not valid_sentences:
        return text[:200] + "..." if len(text) > 200 else text
        
    summary = " ".join(valid_sentences[:num_sentences])
    return summary
