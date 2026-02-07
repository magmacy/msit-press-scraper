import os
import shutil
import pandas as pd
from datetime import datetime
import logging
import config
import utils
import re

logger = logging.getLogger(__name__)

def migrate_folders():
    download_dir = config.DOWNLOAD_DIR
    if not os.path.exists(download_dir):
        logger.info(f"다운로드 폴더가 없습니다: {download_dir}")
        return

    # 모든 엑셀 파일 로드 (날짜 매핑용)
    data_dir = config.DATA_DIR
    title_to_date = {}
    
    if os.path.exists(data_dir):
        excel_files = [f for f in os.listdir(data_dir) if f.endswith('.xlsx') and not f.endswith('_test.xlsx')]
        for excel_file in excel_files:
            excel_path = os.path.join(data_dir, excel_file)
            try:
                df = pd.read_excel(excel_path)
                # 제목과 등록일 매핑
                for _, row in df.iterrows():
                    title = str(row['제목']).strip()
                    date_val = row['등록일']
                    
                    # utils.normalize_date 사용
                    if isinstance(date_val, datetime):
                        date_str = date_val.strftime("%Y-%m-%d")
                    else:
                        date_str = utils.normalize_date(str(date_val))
                    
                    if date_str:
                        clean_title = re.sub(r'[\\/*?:"<>|]', "", title).strip()
                        title_to_date[clean_title] = date_str

                logger.info(f"엑셀 로드: {excel_file} ({len(df)}건)")
            except Exception as e:
                logger.error(f"엑셀 파일 로드 실패 ({excel_file}): {e}")
        logger.info(f"총 {len(title_to_date)}건 매핑 완료")
    else:
        logger.warning("데이터 폴더가 없습니다.")

    print(f"폴더 마이그레이션 시작: {download_dir}")
    count = 0
    
    for folder_name in os.listdir(download_dir):
        folder_path = os.path.join(download_dir, folder_name)
        if not os.path.isdir(folder_path):
            continue
            
        # 기존 패턴: DATE_TITLE
        # 날짜 부분 추출 시도
        try:
            parts = folder_name.split('_', 1)
            if len(parts) != 2:
                continue
                
            current_date_part = parts[0]
            title_part = parts[1].strip()
            
            # 정확한 날짜 찾기
            new_date_part = current_date_part
            
            # 1. 엑셀 매핑 확인 (제목 30자 제한 고려)
            # 폴더명에 쓰인 제목은 30자로 잘려있을 수 있음.
            # title_to_date 키들 중 title_part로 시작하는 것을 찾거나
            # 반대로 title_to_date 키가 title_part를 포함하는지 확인
            
            matched_date = None
            for full_title, date in title_to_date.items():
                # 폴더명 생성 규칙과 동일하게 변환 후 비교
                clean_full_title = re.sub(r'[\\/*?:"<>|]', "", full_title).strip()
                
                # 다양한 매칭 방법 시도
                # 1. 정확히 30자 잘린 경우
                trunc_title = clean_full_title[:30].strip()
                
                # 2. 폴더 제목이 전체 제목으로 시작하는지
                # 3. 전체 제목이 폴더 제목으로 시작하는지
                if (trunc_title == title_part or 
                    clean_full_title.startswith(title_part) or
                    title_part.startswith(clean_full_title[:len(title_part)])):
                    matched_date = date
                    break
            
            if matched_date:
                new_date_part = matched_date
            else:
               # 매핑 실패 시 포맷만이라도 통일
               new_date_part = current_date_part.replace('.', '-')

            new_folder_name = f"{new_date_part}_{title_part}"
            
            if new_folder_name != folder_name:
                new_folder_path = os.path.join(download_dir, new_folder_name)
                
                if os.path.exists(new_folder_path):
                    print(f"[MERGING] {folder_name} -> {new_folder_name}")
                    # 내용물 이동
                    for item in os.listdir(folder_path):
                        src = os.path.join(folder_path, item)
                        dst = os.path.join(new_folder_path, item)
                        try:
                            if os.path.exists(dst):
                                os.remove(dst) # 덮어쓰기를 위해 기존 파일 삭제
                            shutil.move(src, dst)
                        except Exception as e:
                            print(f"  - 파일 이동 실패 ({item}): {e}")
                    
                    # 빈 폴더 삭제
                    try:
                        os.rmdir(folder_path)
                        print(f"[MERGE COMPLETE] {folder_name}")
                        count += 1
                    except OSError:
                        print(f"  - 폴더 삭제 실패 (비어있지 않음): {folder_name}")
                else:
                    try:
                        os.rename(folder_path, new_folder_path)
                        print(f"[RENAME] {folder_name} -> {new_folder_name}")
                        count += 1
                    except Exception as e:
                        print(f"[ERROR] Rename failed: {e}")
        except Exception as e:
            print(f"[ERROR] {folder_name} 처리 중 오류: {e}")
            
    print(f"마이그레이션 완료. {count}개 폴더 변경됨.")

if __name__ == "__main__":
    migrate_folders()
