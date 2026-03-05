import os
import sys
import glob
import re
import json
import datetime
import zipfile
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import range_boundaries

sys.stdout.reconfigure(encoding='utf-8')

# ==================== 설정 ====================
# 날짜 설정 (기본값: 오늘)
import argparse
parser = argparse.ArgumentParser()
parser.add_argument("--date", help="Target date (YYYY-MM-DD)", default=None)
parser.add_argument("--local-only", help="Skip uploading to Drive", action="store_true")
args, _ = parser.parse_known_args() # Ignore unknown args if called from another script with more args

if args.date:
    print(f"Using provided target date: {args.date}")
    TARGET_DATE = datetime.datetime.strptime(args.date, "%Y-%m-%d").date()
else:
    TARGET_DATE = datetime.date.today()

DATE_STR = TARGET_DATE.strftime("%Y년 %m월 %d일")
DATE_SHORT = TARGET_DATE.strftime("%m%d")
DATE_FILENAME = TARGET_DATE.strftime("%Y%m%d")

# 경로 설정
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# Use date-specific file to preserve historical data
INPUT_FILE = os.path.join(PROJECT_ROOT, "data", "inputs", f"processed_orders_{DATE_FILENAME}.txt")
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "data", "outputs", f"invoices_{DATE_FILENAME}")
OUTPUT_ZIP = os.path.join(PROJECT_ROOT, "data", "outputs", f"거래명세서_{DATE_FILENAME}.zip")

# 드라이브 템플릿 루트
DRIVE_ROOT = r"G:\내 드라이브\1. 도크_주문 명세서\1. 거래명세서"

# 원산지 판정 리스트
IMPORT_ITEMS = ['세척당근', '오렌지', '표고버섯2L', '표고버섯', '맛김치', '포기김치', '파김치', '레몬', '백김치', '매운고춧가루', '굵은고춧가루', '씻은지']

def get_origin(item):
    for imp in IMPORT_ITEMS:
        if imp in item:
            return '수입산'
    return '국내산'

# ==================== 데이터 파싱 ====================
def _load_supplier_names():
    """Load canonical supplier names from mappings.json."""
    mappings_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                                 "skills", "order_processing", "resources", "mappings.json")
    try:
        with open(mappings_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        names = set(data.get('supplier_names', []))
    except Exception:
        names = set()
    names.update(["재고, 창고소분", "시장구매, 시장소분"])
    return names

def parse_processed_orders():
    orders = {}

    # Validation: Check if input file exists
    if not os.path.exists(INPUT_FILE):
        print(f"❌ Error: {INPUT_FILE} not found.")
        print(f"Please run first: python execution/process_orders.py --input inputs/orders_{DATE_FILENAME}.txt --date {TARGET_DATE.strftime('%Y-%m-%d')}")
        return orders

    print(f"✓ Reading from: {INPUT_FILE}")

    current_store = None

    # Regex to handle items like "청경채 4박스" or "청경채 4.5박스" or "스테비아토마토1박스" (공백 없는 경우 포함)
    item_pattern = re.compile(r'^(.+?)\s+(\d+(?:\.\d+)?)([a-zA-Z가-힣]+).*$')
    # Fallback: 공백 없이 품명+수량+단위 붙어있는 경우 (예: 스테비아토마토1박스)
    item_pattern_nospace = re.compile(r'^(.+?)(\d+(?:\.\d+)?)([a-zA-Z가-힣]+).*$')

    supplier_names = _load_supplier_names()

    with open(INPUT_FILE, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if line.startswith('-'):
            store_name = line.lstrip('-').strip()
            # Ignore suppliers and internal categories (loaded from mappings.json)
            if store_name in supplier_names:
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
            if not match:
                match = item_pattern_nospace.match(line)
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
    
    # 1. 매핑 정의 (더 구체적인 매핑을 먼저 배치)
    folder_map = {
        # 파라이 - 송파점과 역삼점 구분 (구체적인 것 먼저)
        '파라이 송파점': r'14. 파라이 카이카이\파라이 송파점',
        '파라이 역삼점': r'14. 파라이 카이카이\파라이 역삼점',
        '파라이': r'14. 파라이 카이카이\파라이 역삼점',  # 기본값은 역삼점
        
        # 샤브야키
        '샤브야키 분당점': '1. 샤브야키 분당점',
        '샤브야키 평택점': '2. 샤브야키 평택점',
        '샤브야키 주안점': '3. 샤브야키 주안점',
        '샤브야키 동탄점': '4. 샤브야키 동탄점',
        '샤브야키 도봉점': '4. 샤브야키 도봉점',
        '샤브야키 광교점': '4. 샤브야키 광교점',
        
        # 부엉이산장
        '부엉이산장 강남점': r'5. 부엉이 산장\1. 부엉이산장 강남점',
        '부엉이산장 구월점': r'5. 부엉이 산장\2. 부엉이산장 구월점',
        '부엉이산장 마곡점': r'5. 부엉이 산장\3. 부엉이산장 마곡점',
        
        # 오레노이키루미치
        '오레노이키루미치 하남점': r'7. 오레노이키루미치\오레노이키루미치 하남점',
        '오레노이키루미치 압구정점': r'7. 오레노이키루미치\오레노이키루미치 압구정점',
        
        # 기타
        '고른햇살': '6. 고른햇살',
        '브럭시': '11. 브럭시',
        '선데이버거클럽': '12. 선데이버거클럽',
        '육회꽃필무렵': '13. 육회꽃필무렵 송파점',
        '소유 프루트': '8. 성동 과일 클래스 소유',
        '소유프루트': '8. 성동 과일 클래스 소유',
        '봄날': '9. 가락 과일 클래스 봄날',
        '이너프유': '0. 이너프유',
    }
    
    # Normalize: remove spaces/underscores for fuzzy comparison
    def _norm(s): return s.replace(" ", "").replace("_", "")

    # 매칭 로직: Exact → Partial → Normalized fallback
    target_folder = None

    # 1. Exact match first (최우선)
    if store_name in folder_map:
        target_folder = folder_map[store_name]
        print(f"  [Exact Match] {store_name} → {target_folder}")
    else:
        # 2. Longest partial match (가장 긴 매칭을 우선)
        matched_keys = [(key, val) for key, val in folder_map.items() if key in store_name]
        if matched_keys:
            matched_keys.sort(key=lambda x: len(x[0]), reverse=True)
            target_folder = matched_keys[0][1]
            print(f"  [Partial Match] {store_name} → {matched_keys[0][0]} → {target_folder}")
        else:
            # 3. Normalized fallback (spaces/underscores removed)
            norm_store = _norm(store_name)
            norm_matched = [(key, val) for key, val in folder_map.items()
                           if _norm(key) in norm_store or norm_store in _norm(key)]
            if norm_matched:
                norm_matched.sort(key=lambda x: len(_norm(x[0])), reverse=True)
                target_folder = norm_matched[0][1]
                print(f"  [Normalized Match] {store_name} → {norm_matched[0][0]} → {target_folder}")

    if not target_folder:
        print(f"  [Warning] No folder mapping found for {store_name}")
        return None

    full_path = os.path.join(DRIVE_ROOT, target_folder)
    
    # Prioritize "AI참고용" template in the store folder (recursive)
    # Search extensively
    print(f"  [Debug] Scanning {full_path} for templates...")
    if os.path.exists(full_path):
        for root, dirs, files in os.walk(full_path):
            print(f"  [Debug] Checking {root}, Files: {files}")
            for f in files:
                # Normalizing nfc for reliable comparison
                import unicodedata
                norm_f = unicodedata.normalize('NFC', f)
                if ("AI참고용" in norm_f or "AI" in norm_f) and f.endswith(".xlsx") and not f.startswith("~$"):
                    print(f"  [Info] Found AI Reference Template: {f}")
                    return os.path.join(root, f)
    else:
        print(f"  [Error] Store folder not found: {full_path}")
        return None

    # 2. 연도/월 폴더 찾기 (e.g., "2026년 1월")
    # 다양한 포맷 시도
    y = TARGET_DATE.year
    yy = y % 100
    m = TARGET_DATE.month
    
    # Previous Month Calculation
    if m == 1:
        pm_y = y - 1
        pm_m = 12
    else:
        pm_y = y
        pm_m = m - 1
    
    pm_yy = pm_y % 100

    candidate_folders = [
        f"{y}년 {m}월",
        f"{y}년 {m:02d}월",
        f"{y}년{m}월",
        f"{y}년{m:02d}월",
        f"{yy}년{m}월",
        f"{yy}년{m:02d}월",
        f"{yy}년 {m}월",
        f"{yy}년 {m:02d}월",
        # Previous Month Candidates
        f"{pm_y}년 {pm_m}월",
        f"{pm_y}년 {pm_m:02d}월",
        f"{pm_yy}년 {pm_m}월",
        f"{pm_yy}년 {pm_m:02d}월"
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
    
    # Update Phone Number (K10)
    # Target: 010-3275-0188
    # Using safe_set_value in case of merge, though K10 usually isn't merged in standard templates.
    safe_set_value(sheet, 10, 11, "010-3275-0188")
    
    # Template capacity detection
    # Assume items start at row 16
    start_row = 16
    
    # Scan for Footer ("합계") to determine max capacity
    footer_row_initial = -1
    scan_limit = 100
    for r in range(start_row, start_row + scan_limit):
        # Check cols 2 to 10
        found = False
        for c_idx in range(2, 11):
            val = sheet.cell(row=r, column=c_idx).value
            if val and isinstance(val, str) and ("합계" in val or "계" == val.strip()):
                footer_row_initial = r
                found = True
                break
        if found:
            break
            
    if footer_row_initial != -1:
        max_rows = footer_row_initial - start_row
        print(f"  [Info] Detected template capacity: {max_rows} rows (Footer at {footer_row_initial})")
    else:
        max_rows = 17 # Fallback
        print(f"  [Warning] Could not detect footer. Using default capacity: {max_rows}")
    
    # Clear existing items within range
    for r in range(start_row, start_row + max_rows):
        safe_set_value(sheet, r, 2, None) # 품목
        safe_set_value(sheet, r, 4, None) # 단위
        safe_set_value(sheet, r, 5, None) # 수량
        safe_set_value(sheet, r, 6, None) # 원산지
        safe_set_value(sheet, r, 10, None) # 단가
        safe_set_value(sheet, r, 11, None) # 금액
        safe_set_value(sheet, r, 12, None) # 비고
        
    # Style definitions
    thin_side = Side(style='thin')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    font_body = Font(name='Malgun Gothic', size=10) 

    # Fill Items
    last_item_row = start_row
    for i, (item_name, unit, qty) in enumerate(items):
        if i >= max_rows:
            print(f"  [Warning] Too many items for {store_name} ({len(items)}). Truncating at 17.")
            break
            
        row = start_row + i
        last_item_row = row
        
        # Values
        safe_set_value(sheet, row, 2, item_name)
        safe_set_value(sheet, row, 4, unit)
        safe_set_value(sheet, row, 5, qty)
        safe_set_value(sheet, row, 6, get_origin(item_name))
        
        # Value: Supply Price (Col 11) -> 0
        sheet.cell(row=row, column=11).value = 0
        
        # No explicit styling - keep template style

    # 3.5 Find Footer
    footer_row = -1
    scan_limit = 100
    current_scan = last_item_row + 1
    
    while current_scan < scan_limit:
        found_footer = False
        # Check columns 2 to 10 for "합계"
        for c_idx in range(2, 11):
            val = sheet.cell(row=current_scan, column=c_idx).value
            if val and isinstance(val, str) and ("합계" in val or "계" == val.strip()):
                footer_row = current_scan
                found_footer = True
                break
        if found_footer:
            break
        current_scan += 1

    # 4. Zero out Total Amounts (Top & Bottom)
    # Target keywords
    zero_targets = ["합계", "당월금액", "당일금액", "총사용액", "총", "입금액", "합계금액", "금액"]
    
    # 4-1. Top Section Scan (rows 1-15)
    for r in range(1, 16):
        for c in range(1, 15):
            val = sheet.cell(row=r, column=c).value
            if val and isinstance(val, str):
                norm_val = val.replace(" ", "").replace(":", "").replace("-", "")
                is_target = False
                for t in zero_targets:
                    if t in norm_val:
                        is_target = True
                        break
                
                if is_target:
                    # Found a label. Try to zero out nearby value cells.
                    # Heuristic: Check right neighbor cells (c+1 to c+5) and cells below (r+1)
                    # Specific to Top Amount: often right next to it or below it.
                    
                    found_val = False
                    # Check current row, next columns
                    for search_c in range(c + 1, c + 6):
                        if search_c > 15: break
                        cell_v = sheet.cell(row=r, column=search_c).value
                        if isinstance(cell_v, (int, float)) or (isinstance(cell_v, str) and cell_v.startswith("=")):
                            try:
                                safe_set_value(sheet, r, search_c, 0)
                                found_val = True
                            except: pass
                            
                    # Check next row, same column or next columns (common in some headers)
                    if not found_val:
                        for search_c in range(c, c + 3):
                             if search_c > 15: break
                             cell_v = sheet.cell(row=r+1, column=search_c).value
                             if isinstance(cell_v, (int, float)) or (isinstance(cell_v, str) and cell_v.startswith("=")):
                                 try:
                                     safe_set_value(sheet, r+1, search_c, 0)
                                 except: pass

    if footer_row != -1:
        # 4-2. Hide Empty Rows
        rows_to_hide_start = last_item_row + 1
        rows_to_hide_end = footer_row - 1
        
        if rows_to_hide_end >= rows_to_hide_start:
            print(f"  [Info] Hiding empty rows {rows_to_hide_start} to {rows_to_hide_end}")
            for r in range(rows_to_hide_start, rows_to_hide_end + 1):
                sheet.row_dimensions[r].hidden = True
        
        # 4-3. Zero out Footer Amounts (Bottom)
        # Scan footer area (footer_row to footer_row + 10)
        for r in range(footer_row, footer_row + 15):
             for c in range(2, 14): # Scan cols 2 to 14
                val = sheet.cell(row=r, column=c).value
                if val and isinstance(val, str):
                    norm_val = val.replace(" ", "").replace(":", "").replace("-", "")
                    
                    is_target = False
                    for t in zero_targets:
                        if t in norm_val:
                            is_target = True
                            break
                    
                    if is_target:
                        # "합계" row -> Col 11 is Total Amount (Standard)
                        if "합계" in norm_val and r == footer_row:
                             safe_set_value(sheet, r, 11, 0)
                        
                        # General search for values to zero
                        for search_c in range(c + 1, 14):
                             cell_v = sheet.cell(row=r, column=search_c).value
                             if (cell_v and isinstance(cell_v, str) and cell_v.startswith("=")) or isinstance(cell_v, (int, float)):
                                 try:
                                     safe_set_value(sheet, r, search_c, 0)
                                 except: pass
                        
                        # Fallback for "합계" row if Col 11 wasn't zeroed by loop
                        # (Because sometimes Col 11 is empty but has a formula we can't see, or is just empty)
                        # We force 0 to Col 11 anyway if it's the footer row
                        if r == footer_row:
                            safe_set_value(sheet, r, 11, 0)

    else:
        print(f"  [Warning] Could not find footer ('합계' etc). Skipping row hiding/zeroing.")

    # Save Local
    filename = f"{DATE_FILENAME}_{store_name.replace(' ', '_')}_거래명세서.xlsx"
    output_path = os.path.join(OUTPUT_DIR, filename)
    wb.save(output_path)
    
    # Save to Drive (Month Folder)
    # Drive Root for this client
    template_dir_name = os.path.basename(os.path.dirname(template_path))
    # Simple check if the template is inside a month folder (contains '월' and '년')
    if "월" in template_dir_name and "년" in template_dir_name:
        drive_client_root = os.path.dirname(os.path.dirname(template_path))
    else:
        drive_client_root = os.path.dirname(template_path)
    
    # Check for existing month folder to avoid duplicates
    # Patterns: 2026년 1월, 26년 1월, 2026년1월, 26년1월, etc.
    y = TARGET_DATE.year
    yy = y % 100
    m = TARGET_DATE.month
    
    candidate_month_folders = [
        f"{y}년 {m}월",
        f"{yy}년 {m}월",      # 26년 1월
        f"{y}년{m}월",
        f"{yy}년{m}월",
        f"{y}년 {m:02d}월",
        f"{yy}년 {m:02d}월"
    ]
    
    target_month_folder = None
    
    # 1. Search for existing folder
    # Helper to clean folder names for comparison
    def clean_name(n): return n.replace(" ", "").replace("0", "")

    target_clean = f"{y}년{m}월".replace("0", "")
    target_clean_short = f"{yy}년{m}월".replace("0", "")
    
    # List actual directories
    try:
        if os.path.exists(drive_client_root):
            actual_subdirs = [d for d in os.listdir(drive_client_root) if os.path.isdir(os.path.join(drive_client_root, d))]
            for d in actual_subdirs:
                dc = clean_name(d)
                if target_clean in dc or target_clean_short in dc:
                     # Check if it looks like a date folder
                     if "년" in d and "월" in d:
                         target_month_folder = d
                         print(f"  [Info] Found existing month folder (fuzzy): {target_month_folder}")
                         break
    except: pass

    if not target_month_folder:
        for folder_name in candidate_month_folders:
            check_path = os.path.join(drive_client_root, folder_name)
            if os.path.exists(check_path):
                target_month_folder = folder_name
                print(f"  [Info] Found existing month folder: {target_month_folder}")
                break
            
    # 2. If not found, default to standard "YYYY년 M월" (or what the user prefers?)
    # User said "2026년2월" exists. So if not found, maybe I create "2026년 2월" (with space) or without?
    # I'll stick to space as default.
    if not target_month_folder:
        target_month_folder = f"{y}년 {m}월"
        print(f"  [Info] Creating new month folder: {target_month_folder}")

    drive_target_dir = os.path.join(drive_client_root, target_month_folder)
    
    try:
        if not args.local_only:
            os.makedirs(drive_target_dir, exist_ok=True)
            drive_path = os.path.join(drive_target_dir, filename)
            shutil.copy2(output_path, drive_path)
            print(f"  [Drive] Saved to: {drive_path}")
        else:
             print(f"  [Info] Skipping Drive upload (local-only): {drive_target_dir}")
    except Exception as e:
        print(f"  [Drive Error] Could not save to {drive_target_dir}: {e}")
    
    # Save to Drive (Same folder as template or Client Root)
    # User requested to ONLY save in the month folder, so we skip saving to the root/template folder.
    # drive_folder = os.path.dirname(template_path)
    # drive_path = os.path.join(drive_folder, filename)
    # ... code removed ...

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

        from _notify import notify
        notify("stage2_invoice", date_str=args.date)
    else:
        print("\nNo invoices created.")

if __name__ == "__main__":
    main()
