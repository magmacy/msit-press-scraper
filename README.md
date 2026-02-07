# 과학기술정보통신부 보도자료 스크래퍼

과학기술정보통신부 웹사이트의 보도자료를 수집하여 엑셀 파일로 저장하고 첨부파일을 다운로드하는 도구입니다.

## 주요 기능

- **자동화된 수집**: JavaScript로 렌더링되는 목록 및 상세 페이지 자동 파싱
- **이어받기**: 중단된 시점부터 수집 재개 (중복 데이터 건너뜀)
- **안정성**: 네트워크 불안정 시 자동 재시도 및 로깅 기능
- **첨부파일**: 게시글별 첨부파일 자동 다운로드
- **요약**: 본문 내용의 핵심 3문장 요약 제공

## 설치 방법

1. 필요한 패키지 설치

```bash
pip install -r requirements.txt
```

## 사용 방법

### 기본 실행 (2024년 이후 데이터)

```bash
python scraper.py
```

### 테스트 실행 (5건만 수집)

```bash
python scraper.py --test
```

### 옵션 사용

```bash
# 특정 페이지부터 시작
python scraper.py --page 10

# 특정 연도 이후 데이터 수집
python scraper.py --year 2023
```

## 결과물

- **엑셀 파일**: `data/press_releases_YYYYMMDD.xlsx`
- **첨부파일**: `downloads/YYYY-MM-DD_제목/`
- **로그 파일**: `logs/scraper_YYYYMMDD.log`

## 프로젝트 구조

```
.
├── scraper.py          # 메인 실행 파일
├── config.py           # 설정 (URL, 경로, 헤더 등)
├── utils.py            # 유틸리티 함수 (텍스트 정제, 날짜 파싱 등)
├── data/               # 수집된 엑셀 파일 저장소
├── downloads/          # 첨부파일 다운로드 경로
└── logs/               # 실행 로그
```
