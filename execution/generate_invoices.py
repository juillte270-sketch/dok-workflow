import os
import glob
import re
import datetime
import zipfile
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import range_boundaries

# ==================== 설정 ====================
# 날짜 설정 (기본값: 오늘)
# TARGET_DATE = datetime.date.today()
TARGET_DATE = datetime.date(2026, 1, 30) # Fixed for this task
DATE_STR = TARGET_DATE.strftime("%Y년 %m월 %d일")
DATE_SHORT = TARGET_DATE.strftime("%m%d")
DATE_FILENAME = TARGET_DATE.strftime("%Y%m%d")

# 경로 설정
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_FILE = os.path.join(PROJECT_ROOT, "data", "inputs", "processed_orders.txt")
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "data", "outputs", f"invoices_{DATE_FILENAME}")
OUTPUT_ZIP = os.path.join(PROJECT_ROOT, "data", "outputs", f"거래명세서_{DATE_FILENAME}.zip")

# 드라이브 템플릿 루트
DRIVE_ROOT = r"G:\내 드라이브\1. 도크_주문 명세서\1. 거래명세서"

# 원산지 판정 리스트
IMPORT_ITEMS = ['세척당근', '오렌지', '표고버섯2L', '표고버섯', '맛김치', '포기김치', '파김치', '레몬', '백김치']

def get_origin(item):
    for imp in IMPORT_ITEMS:
        if imp in item:
            return '수입산'
    return '국내산'

# ==================== 데이터 파싱 ====================
def parse_processed_orders():
    orders = {}
    if not os.path.exists(INPUT_FILE):
        print(f"Error: {INPUT_FILE} not found.")
        return orders

    current_store = None
    
    # Regex to handle items like "청경채 4박스" or "청경채 4.5박스"
    # Same regex as update_order_sheet.py
    item_pattern = re.compile(r'^(.+?)\s+(\d+(?:\.\d+)?)([a-zA-Z가-힣]+).*$')

    with open(INPUT_FILE, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        if line.startswith('-'):
            store_name = line.lstrip('-').strip()
            # Ignore suppliers and internal categories
            if store_name in ["영운농산", "다모아버섯", "건영농산", "오복상회", "명진농산", "풍경농산", 
                          "나물향기", "태현상회", "가야웰빙", "이모유통", "초원농산", "소리농산", 
                          "명화농산", "우신농산", "정복상회", "정운", "재고, 창고소분", "시장구매, 시장소분"]:
                current_store = None
                continue
            current_store = store_name
            if current_store not in orders:
                orders[current_store] = []
            continue

        if current_store:
            if line.startswith('('):
                continue
                
            match = item_pattern.match(line)
            if match:
                item_name = match.group(1).strip()
                qty = float(match.group(2))
                unit = match.group(3).strip()
                orders[current_store].append((item_name, unit, qty))
                
    return orders

# ==================== 템플릿 찾기 ====================
def find_template(store_name):
    """
    거래처 이름에 맞는 템플릿 파일 경로를 찾습니다.
    1. 거래처 폴더 매핑 확인
    2. 연/월 폴더 확인 (예: 2026년 1월)
    3. 최신 .xlsx 파일 반환
    """
    
    # 1. 매핑 정의
    folder_map = {
        '샤브야키 분당점': '1. 샤브야키_분당점',
        '샤브야키 평택점': '2. 샤브야키 평택점',
        '샤브야키 주안점': '3. 샤브야키 주안점',
        '샤브야키 동탄점': '4. 샤브야키 동탄점',
        '고른햇살': '6. 고른햇살',
        '브럭시': '11. 브럭시',
        '선데이버거클럽': '12. 선데이버거클럽',
        '육회꽃필무렵': '13. 육회꽃필무렵 송파점',
        '파라이': '14. 파라이 카이카이',
        '소유프루트': '8. 성동 과일 클래스 소유',
        '봄날': '9. 가락 과일 클래스 봄날',
        '오레노이키루미치 하남점': r'7. 오레노이키루미치\오레노이키루미치 하남점',
        '오레노이키루미치 압구정점': r'7. 오레노이키루미치\오레노이키루미치 압구정점',
        '부엉이산장 강남점': r'5. 부엉이 산장\1. 부엉이산장 강남점',
        '부엉이산장 구월점': r'5. 부엉이 산장\2. 부엉이산장 구월점',
        '부엉이산장 마곡점': r'5. 부엉이 산장\3. 부엉이산장 마곡점',
        # 필요한 경우 추가 매핑
    }
    
    # 일부 이름 매칭 (예: '파라이(송파)' -> '파라이')
    target_folder = None
    
    # Exact match first
    if store_name in folder_map:
        target_folder = folder_map[store_name]
    else:
        # Partial match
        for key, val in folder_map.items():
            if key in store_name:
                target_folder = val
                break
    
    if not target_folder:
        print(f"  [Warning] No folder mapping found for {store_name}")
        return None

    full_path = os.path.join(DRIVE_ROOT, target_folder)
    
    # 2. 연도/월 폴더 찾기 (e.g., "2026년 1월")
    # 2. 연도/월 폴더 찾기
    # 다양한 포맷 시도
    # - 2026년 1월 (분당점)
    # - 26년1월 (주안점)
    # - 2026년 01월
    # - 26년 01월
    
    y = TARGET_DATE.year
    yy = y % 100
    m = TARGET_DATE.month
    
    candidate_folders = [
        f"{y}년 {m}월",
        f"{y}년 {m:02d}월",
        f"{y}년{m}월",
        f"{y}년{m:02d}월",
        f"{yy}년{m}월",
        f"{yy}년{m:02d}월",
        f"{yy}년 {m}월",
        f"{yy}년 {m:02d}월"
    ]
    
    ym_path = None
    for folder_name in candidate_folders:
        check_path = os.path.join(full_path, folder_name)
        if os.path.exists(check_path):
            ym_path = check_path
            break
            
    if not ym_path:
        print(f"  [Info] Month folder not found in {full_path}. Tried: {candidate_folders}")
        # Root fallback check?
        # Only fallback if explicitly safe, but user said 'latest in folder'. 
        # Let's try root as a last resort if defined, otherwise fail.
        # Check files in root
        if glob.glob(os.path.join(full_path, "*.xlsx")):
             print("  [Info] Found xlsx in root, using root.")
             ym_path = full_path
        else:
             print(f"  [Error] Path not found: {os.path.join(full_path, candidate_folders[0])} (and others)")
             return None
        
    # 3. 최신 xlsx 찾기
    files = glob.glob(os.path.join(ym_path, "*.xlsx"))
    if not files:
        print(f"  [Error] No .xlsx files found in {ym_path}")
        return None
        
    # Sort by modification time (files might be named by date, but mod time is safer if names vary)
    files.sort(key=os.path.getmtime, reverse=True)
    
    # Skip temporary files (~$)
    valid_files = [f for f in files if not os.path.basename(f).startswith('~$')]
    
    if not valid_files:
        return None
        
    return valid_files[0]

# ==================== 엑셀 유틸 ====================
def get_merge_start(sheet, row, col):
    for merged_range in sheet.merged_cells.ranges:
        min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
        if min_row <= row <= max_row and min_col <= col <= max_col:
            return min_row, min_col
    return row, col

def safe_set_value(sheet, row, col, value):
    start_row, start_col = get_merge_start(sheet, row, col)
    sheet.cell(row=start_row, column=start_col).value = value

# ==================== 송장 생성 ====================
def create_invoice(store_name, items):
    template_path = find_template(store_name)
    if not template_path:
        print(f"  [Skip] Skipping {store_name}: Template not found.")
        return None

    print(f"  Processing {store_name} using {os.path.basename(template_path)}...")
    
    wb = load_workbook(template_path)
    sheet = wb.active # Assume first sheet is the target
    
    # Update Date
    # Try finding "202" or "년" in top rows to verify standard layout?
    # User sample code hardcoded 'D6'. Let's trust that for now, but be careful.
    try:
        sheet['D6'] = DATE_STR
    except:
         # Some templates might be different. Let's try to find a cell with date-like format?
         # For now, safe_set_value to D6
         safe_set_value(sheet, 6, 4, DATE_STR)

    # 4, 2 might be store name (B4)
    safe_set_value(sheet, 4, 2, store_name)
    
    # Clear existing items
    # Assume items start at row 16, end at 32 (17 rows)
    start_row = 16
    max_rows = 17 # typical max for one page
    
    for r in range(start_row, start_row + max_rows):
        safe_set_value(sheet, r, 2, None) # 품목
        safe_set_value(sheet, r, 4, None) # 단위
        safe_set_value(sheet, r, 5, None) # 수량
        safe_set_value(sheet, r, 6, None) # 원산지
        safe_set_value(sheet, r, 10, None) # 단가
        safe_set_value(sheet, r, 11, None) # 금액
        safe_set_value(sheet, r, 12, None) # 비고
        
    # Fill items
    for i, (item_name, unit, qty) in enumerate(items):
        if i >= max_rows:
            print(f"  [Warning] Too many items for {store_name} ({len(items)}). Truncating.")
            break
            
        row = start_row + i
        safe_set_value(sheet, row, 2, item_name)
        safe_set_value(sheet, row, 4, unit)
        safe_set_value(sheet, row, 5, qty)
        safe_set_value(sheet, row, 6, get_origin(item_name))
        
        # Formula for Amount (Qty * Price). Price is Col J (10), Amount is Col K (11)
        # =E16*J16
        sheet.cell(row=row, column=11).value = f"=E{row}*J{row}"
        
    # Save
    filename = f"{DATE_FILENAME}_{store_name.replace(' ', '_')}_거래명세서.xlsx"
    output_path = os.path.join(OUTPUT_DIR, filename)
    wb.save(output_path)
    return output_path

# ==================== 메인 ====================
def main():
    print(f"=== Generating Invoices for {DATE_STR} ===")
    
    orders = parse_processed_orders()
    if not orders:
        print("No orders found to process.")
        return

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    created_files = []
    
    for store_name, items in orders.items():
        # Clean store name for special cases
        # e.g. "파라이송파점" should map to "파라이"? or "육회꽃필무렵" matches folder "육회꽃필무렵 송파점"
        # The fuzzy match in find_template should handle most.
        
        # Mapping checks
        filepath = create_invoice(store_name, items)
        if filepath:
            created_files.append(filepath)
            
    if created_files:
        with zipfile.ZipFile(OUTPUT_ZIP, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in created_files:
                zf.write(f, os.path.basename(f))
        print(f"\nAll done. Created {len(created_files)} invoices.")
        print(f"Output Directory: {OUTPUT_DIR}")
        print(f"ZIP File: {OUTPUT_ZIP}")
    else:
        print("\nNo invoices created.")

if __name__ == "__main__":
    main()
