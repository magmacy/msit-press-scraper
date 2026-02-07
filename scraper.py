import os
import sys
import time
import logging
import argparse
import requests
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin, unquote
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from tqdm import tqdm
import re
from datetime import datetime

import config
import utils
import migrate_folders

# 로깅 설정
def setup_logging():
    os.makedirs(config.LOG_DIR, exist_ok=True)
    
    log_file = os.path.join(config.LOG_DIR, f"scraper_{config.TODAY_STR}.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return logging.getLogger(__name__)

logger = setup_logging()

class PressReleaseScraper:
    def __init__(self, year=config.TARGET_YEAR, output_file=config.EXCEL_PATH):
        self.target_year = year
        self.output_file = output_file
        self.session = self._setup_session()
        self.collected_data = []
        self.seen_ids = set()
        
        # 이어받기: 기존 파일이 있으면 ID 로드
        if os.path.exists(self.output_file):
            try:
                df = pd.read_excel(self.output_file)
                if '번호' in df.columns:
                    self.seen_ids = set(df['번호'].astype(str).tolist())
                    logger.info(f"기존 데이터 {len(self.seen_ids)}건 로드 완료. 중복 수집을 건너뜁니다.")
            except Exception as e:
                logger.warning(f"기존 파일 로드 실패: {e}")

    def _setup_session(self):
        """안정적인 네트워크 요청을 위한 세션 설정"""
        session = requests.Session()
        session.headers.update(config.HEADERS)
        
        retry_strategy = Retry(
            total=config.MAX_RETRIES,
            backoff_factor=config.BACKOFF_FACTOR,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["HEAD", "GET", "OPTIONS"]
        )
        
        adapter = HTTPAdapter(max_retries=retry_strategy)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        return session

    def download_attachment(self, url, folder_name):
        """첨부파일 다운로드"""
        try:
            response = self.session.get(url, stream=True, timeout=config.TIMEOUT)
            response.raise_for_status()
            
            filename = ""
            # Content-Disposition 헤더 확인
            if "Content-Disposition" in response.headers:
                cd = response.headers["Content-Disposition"]
                # RFC 5987: filename*=UTF-8''EncodedString
                matches = re.findall(r"filename\*=UTF-8''(.+)", cd)
                if matches:
                    filename = unquote(matches[0])
                else:
                    # filename="Name"
                    matches = re.findall(r'filename="([^"]+)"', cd)
                    if matches:
                        filename = unquote(matches[0])
            
            # 헤더에서 실패했거나 없는 경우 URL에서 추출
            if not filename:
                # 리다이렉트된 최종 URL 기준
                if response.history:
                    url = response.url
                filename = unquote(os.path.basename(url))
            
            # 파일명 정제 (특수문자 제거)
            filename = re.sub(r'[\\/*?:"<>|]', "", filename)
            
            # 파일명 길이 제한 (Windows MAX_PATH 고려, 100자로 제한)
            name, ext = os.path.splitext(filename)
            if len(name) > 80:
                name = name[:80]
            filename = f"{name}{ext}"
            
            save_dir = os.path.join(config.DOWNLOAD_DIR, folder_name)
            os.makedirs(save_dir, exist_ok=True)
            
            file_path = os.path.join(save_dir, filename)
            
            # 이미 있으면 스킵
            if os.path.exists(file_path):
                return filename, file_path

            with open(file_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
                    
            return filename, file_path
            
        except Exception as e:
            logger.error(f"파일 다운로드 실패 ({url}): {e}")
            return None, None

    def _extract_dates_from_script(self, soup):
        """
        페이지 내 스크립트에서 동적으로 할당되는 날짜 정보를 추출합니다.
        형식: $('#td_'+'REG_DT'+'_0').html('Feb 6, 2026');
        반환값: {index: date_str} 딕셔너리
        """
        date_map = {}
        scripts = soup.find_all('script')
        target_script = None

        for script in scripts:
            if script.string and "$('#td_'" in script.string and "REG_DT" in script.string:
                target_script = script.string
                break
        
        if target_script:
            # 정규식으로 인덱스와 날짜 추출
            # 예: $('#td_'+'REG_DT'+'_0').html('Feb 6, 2026');
            pattern = re.compile(r"\$\('#td_'\s*\+\s*'REG_DT'\s*\+\s*'_(\d+)'\)\.html\('([^']+)'\)")
            matches = pattern.findall(target_script)
            
            for idx, date_str in matches:
                # utils.normalize_date로 날짜 표준화
                date_map[int(idx)] = utils.normalize_date(date_str) 

        return date_map

    def get_list_page(self, page):
        """목록 페이지 파싱"""
        url = f"{config.LIST_URL}&pageIndex={page}"
        try:
            response = self.session.get(url, timeout=config.TIMEOUT)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 스크립트에서 날짜 추출
            script_dates = self._extract_dates_from_script(soup)
            
            # 링크 찾기 및 인덱스 매핑
            # 스크립트의 인덱스(_0, _1...)는 .board_list 내의 순서와 일치함
            links = soup.select('.board_list .toggle > a[onclick^="fn_detail"]')

            items = []
            seen_page_ids = set()

            for idx, link in enumerate(links):
                onclick = link['onclick']
                match = re.search(r"fn_detail\((\d+)\)", onclick)
                if not match:
                    continue
                ntt_id = match.group(1)

                if ntt_id in seen_page_ids:
                    continue
                seen_page_ids.add(ntt_id)

                # 날짜 가져오기: 스크립트 매핑 우선
                date_str = script_dates.get(idx, "")
                
                # HTML 백업 (혹시 모를 상황 대비)
                if not date_str:
                    li = link.find_parent('div', class_='toggle')
                    if li:
                        date_div = li.find('div', class_='date')
                        if date_div:
                            date_str = date_div.get_text(strip=True).replace('등록일', '').strip()

                if not date_str:
                    date_str = datetime.now().strftime("%Y-%m-%d")

                items.append((ntt_id, date_str))

            return items
            
        except Exception as e:
            logger.error(f"목록 페이지 {page} 로드 실패: {e}")
            return []

    def get_detail_page(self, ntt_id, date_str):
        """상세 페이지 파싱"""
        url = f"{config.BASE_URL}/bbs/view.do?sCode=user&mPid=208&mId=307&bbsSeqNo=94&nttSeqNo={ntt_id}"
        
        try:
            response = self.session.get(url, timeout=config.TIMEOUT)
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 제목
            title_elem = soup.select_one('.view_head h2')
            title = utils.clean_text(title_elem.get_text()) if title_elem else f"제목없음_{ntt_id}"
            
            # 부서
            dept = ""
            for dt in soup.select('.tit_con dt'):
                if "부서" in dt.get_text():
                    dd = dt.find_next_sibling('dd')
                    if dd:
                        dept = utils.clean_text(dd.get_text())
                    break
            
            # 본문
            content_div = soup.select_one('.board_notcon') or soup.select_one('.board_pc')
            content = utils.clean_text(content_div.get_text()) if content_div else ""
            
            # 요약
            summary = utils.summarize_text(content)
            
            # 첨부파일 처리
            attachments = []
            file_paths = []
            
            # JS 다운로드 패턴: fn_download('atch_no', 'file_ord', 'ext')
            download_scripts = re.findall(r"fn_download\('(\d+)',\s*'(\d+)',\s*'([^']+)'\)", response.text)
            downloaded_set = set()
            
            title_clean = re.sub(r'[\\\\/*?:\"<>|]', '', title)
            folder_name = f"{date_str}_{title_clean[:30].strip()}"
            
            for atch_no, file_ord, _ in download_scripts:
                down_url = f"{config.BASE_URL}/ssm/file/fileDown.do?atchFileNo={atch_no}&fileOrd={file_ord}&fileBtn=A"
                if down_url in downloaded_set:
                    continue
                    
                fname, fpath = self.download_attachment(down_url, folder_name)
                if fname:
                    attachments.append(fname)
                    file_paths.append(fpath)
                    downloaded_set.add(down_url)
            
            return {
                '번호': ntt_id,
                '제목': title,
                '등록일': date_str,
                '부서': dept,
                '상세URL': url,
                '본문': content,
                '핵심요약': summary,
                '첨부파일목록': ", ".join(attachments),
                '첨부파일경로': ", ".join(file_paths)
            }
            
        except Exception as e:
            logger.error(f"상세 페이지 {ntt_id} 파싱 실패: {e}")
            return None

    def save_data(self):
        """데이터 저장"""
        if not self.collected_data:
            return

        new_df = pd.DataFrame(self.collected_data)
        
        if os.path.exists(self.output_file):
            try:
                old_df = pd.read_excel(self.output_file)
                # 번호 기준 중복 제거 후 병합
                combined_df = pd.concat([old_df, new_df]).drop_duplicates(subset=['번호'], keep='last')
                try:
                    combined_df.to_excel(self.output_file, index=False)
                except PermissionError:
                    # 파일이 열려있어서 저장이 안되는 경우
                    new_filename = self.output_file.replace(".xlsx", f"_backup_{datetime.now().strftime('%H%M%S')}.xlsx")
                    logger.warning(f"파일이 열려있어 저장할 수 없습니다. 백업 파일로 저장합니다: {new_filename}")
                    combined_df.to_excel(new_filename, index=False)
            except Exception as e:
                logger.error(f"데이터 병합 저장 중 오류: {e}")
                # 병합 실패 시 현재 데이터라도 따로 저장
                new_filename = self.output_file.replace(".xlsx", f"_partial_{datetime.now().strftime('%H%M%S')}.xlsx")
                new_df.to_excel(new_filename, index=False)
        else:
            try:
                new_df.to_excel(self.output_file, index=False)
            except PermissionError:
                new_filename = self.output_file.replace(".xlsx", f"_new_{datetime.now().strftime('%H%M%S')}.xlsx")
                logger.warning(f"파일이 열려있어 저장할 수 없습니다. 새 파일로 저장합니다: {new_filename}")
                new_df.to_excel(new_filename, index=False)
            
        logger.info(f"데이터 저장 완료: {self.output_file}")
        # 메모리 정리
        self.collected_data = []

    def run(self, start_page=1, test_mode=False):
        logger.info(f"수집 시작 (대상 연도: {self.target_year}년 이상)")
        if test_mode:
            logger.info(">> 테스트 모드: 수집 건수가 5건에 도달하면 종료합니다.")

        os.makedirs(config.DATA_DIR, exist_ok=True)
        os.makedirs(config.DOWNLOAD_DIR, exist_ok=True)
        
        page = start_page
        stop_flag = False
        total_collected = 0
        
        # tqdm 설정
        pbar = tqdm(desc="페이지 수집", unit="page")
        
        while not stop_flag:
            pbar.set_description(f"Page {page}")
            items = self.get_list_page(page)
            
            if not items:
                logger.info("더 이상 게시글이 없거나 파싱에 실패했습니다.")
                break
                
            new_page_items = 0
            
            for idx, (ntt_id, date_str) in enumerate(items):
                # 연도 체크
                dt = utils.parse_date(date_str)
                if dt and dt.year < self.target_year:
                    logger.info(f"기준 연도({self.target_year}) 이전 데이터 도달 ({date_str}). 종료합니다.")
                    stop_flag = True
                    break
                
                # 중복 체크
                if str(ntt_id) in self.seen_ids and not test_mode:
                    continue
                    
                # 상세 수집
                # 진행 상황 로그 (터미널 출력용)
                tqdm.write(f"  - [{idx+1}/{len(items)}] 상세 수집 중: {ntt_id} ({date_str})")
                
                data = self.get_detail_page(ntt_id, date_str)
                if data:
                    self.collected_data.append(data)
                    self.seen_ids.add(str(ntt_id))
                    new_page_items += 1
                    total_collected += 1
                    
                    tqdm.write(f"    Target: {data['제목'][:30]}...")

                    if test_mode and total_collected >= 5:
                        logger.info("테스트 목표 달성 (5건). 종료합니다.")
                        stop_flag = True
                        break
                    
                time.sleep(1) # 부하 조절
            
            # 페이지 단위 저장
            if self.collected_data:
                self.save_data()
                
            if new_page_items == 0 and not stop_flag and not test_mode:
                logger.info(f"페이지 {page}의 모든 데이터가 이미 수집되었습니다. (중복)")
                # 연속으로 중복이면 종료하는 로직도 고려 가능하지만 일단 진행
                
            page += 1
            pbar.update(1)
            
            if test_mode and stop_flag:
                break
            
        pbar.close()
        logger.info("수집 종료")

def main():
    parser = argparse.ArgumentParser(description="과학기술정보통신부 보도자료 스크래퍼")
    parser.add_argument("--page", type=int, default=1, help="시작 페이지 번호")
    parser.add_argument("--year", type=int, default=config.TARGET_YEAR, help="수집 기준 연도 (이후 데이터 수집)")
    parser.add_argument("--test", action="store_true", help="테스트 모드 (1페이지만 수집하고 종료)")
    
    args = parser.parse_args()
    
    # 설정 오버라이드
    if args.year:
        config.TARGET_YEAR = args.year
        
    scraper = PressReleaseScraper(year=config.TARGET_YEAR)
    
    if args.test:
        # 테스트 모드 시 파일명 변경 (덮어쓰기 방지)
        scraper.output_file = config.EXCEL_PATH.replace(".xlsx", "_test.xlsx")
        
    scraper.run(start_page=args.page, test_mode=args.test)
    
    # 수집 완료 후 폴더명 변경 (마이그레이션) 자동 실행
    if not args.test: # 테스트 모드가 아닐 때만 실행하거나, 필요에 따라 조정
        logger.info("폴더명 마이그레이션(날짜 수정) 시작...")
        try:
            migrate_folders.migrate_folders()
        except Exception as e:
            logger.error(f"마이그레이션 실행 중 실패: {e}")

if __name__ == "__main__":
    main()
