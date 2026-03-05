import argparse
import os
import openpyxl
from copy import copy
from datetime import datetime
from openpyxl.utils import range_boundaries
import sys

# Set encoding
sys.stdout.reconfigure(encoding='utf-8')
from openpyxl.styles import Alignment, Border, Side

# ==================== HELPER FUNCTIONS ====================

def parse_order_list(order_text):
    """
    주문 목록 텍스트를 파싱하여 발주된 핵심 품목 키워드를 추출합니다.
    """
    ordered_keywords = set()
    
    # 무시할 단어들
    ignore_words = ['박스', '봉', '망', 'kg', 'kg', 'g', '개', '팩', '단', '관', '수', '국산', '수입']
    
    lines = order_text.strip().split('\n')
    
    for line in lines:
        trimmed = line.strip()
        if not trimmed: continue
        
        # 주석이나 섹션 헤더 처리
        if trimmed.startswith('-') or trimmed.startswith('(') or trimmed.endswith(')'):
             # 섹션 헤더일 수 있지만, 품목이 포함될 수도 있음 (예: (선발주))
             if trimmed.startswith('-'): continue
             if trimmed.startswith('(') and '금일' in trimmed: continue
        
        # 내용이 있는 라인에서 품목 추출
        words = trimmed.split()
        if not words: continue
        
        first_word = words[0]
        
        # 첫 단어가 수량이거나 단위면 건너뜀 (숫자로 시작하면 무시)
        if first_word[0].isdigit(): continue
        
        # 키워드 정제: 괄호 제거
        keyword = first_word.split('(')[0]
        
        # 정제: 숫자+단위 혹은 호수 등 제거 (예: 양배추45 -> 양배추, 완숙토마토1호 -> 완숙토마토)
        # 단순하게 끝에 붙은 숫자/호/L 등을 제거하는 매핑 추가
        import re
        # remove pattern: number + optional "호", "L", "K", "kg", "g" at the end
        # But be careful not to kill "2L" if product name is short?
        # Safe strategy: Custom Cleanups
        
        if "양배추" in keyword: keyword = "양배추"
        if "깐양파" in keyword: keyword = "깐양파" # 깐양파2호 -> 깐양파
        if "표고" in keyword: keyword = "표고버섯" # 표고버섯2L -> 표고버섯
        if "토마토" in keyword: keyword = "토마토" # 완숙토마토1호 -> 토마토 (Master might vary)
        if "완숙토마토" in keyword: keyword = "완숙토마토"
        
        if not keyword: continue
        
        # 예외 처리 및 매핑
        if keyword == "일자콩나물": keyword = "콩나물"
        if keyword in ["피양파", "깐양파", "양파"]: keyword = "양파"
        if keyword in ["흙대파", "깐대파", "대파"]: keyword = "대파"
        if keyword in ["깐쪽파", "쪽파"]: keyword = "쪽파"
        if keyword == "취청오이": keyword = "오이"
        if keyword == "세척무": keyword = "무"
        # 마늘: 간마늘/깐마늘 구분 유지 (뭉치지 않음)
        if keyword == "세척당근": keyword = "당근"
        if keyword == "청경채": keyword = "청경채" # Explicit
        if keyword == "잎로메인": keyword = "로메인"  # 잎로메인 → 로메인(일반) 매칭용

        ordered_keywords.add(keyword)

    # 추가 매핑 (Reverse Mapping for Matching: If user orders 'Onion', match 'Peeled Onion' too)
    base_keywords = list(ordered_keywords)
    for k in base_keywords:
        if k == "양파":
            ordered_keywords.add("깐양파"); ordered_keywords.add("피양파")
        if k == "대파":
            ordered_keywords.add("흙대파"); ordered_keywords.add("깐대파")
        # 마늘: 깐마늘 주문 시 간마늘 자동 추가하지 않음 (별개 품목)
        if k == "오이":
            ordered_keywords.add("취청오이")
        if k == "토마토":
             ordered_keywords.add("완숙토마토")
        if k == "완숙토마토":
             ordered_keywords.add("토마토")
        if k == "표고버섯":
             ordered_keywords.add("생표고"); ordered_keywords.add("표고")

    return ordered_keywords

def get_grade_priority(item_name):
    """품목명의 등급 우선순위 반환 (낮을수록 우선순위 높음)"""
    if '특' in item_name: return 1
    if '상' in item_name: return 2
    if '보통' in item_name: return 3
    return 4 # 등급 명시 없음

def copy_cell(src_cell, dest_cell):
    """셀 값과 스타일 복사"""
    dest_cell.value = src_cell.value
    if src_cell.has_style:
        dest_cell.font = copy(src_cell.font)
        dest_cell.border = copy(src_cell.border)
        dest_cell.fill = copy(src_cell.fill)
        dest_cell.number_format = copy(src_cell.number_format)
        dest_cell.protection = copy(src_cell.protection)
        dest_cell.alignment = copy(src_cell.alignment)

def delete_old_date_rows(ws, target_date):
    """Delete all rows in the sheet that match the target date in column 2."""
    from datetime import datetime
    
    if isinstance(target_date, str):
        dt_obj = datetime.strptime(target_date, "%Y-%m-%d")
    else:
        dt_obj = target_date
    
    rows_to_delete = []
    
    for r in range(2, ws.max_row + 1):
        date_val = ws.cell(row=r, column=2).value
        
        if date_val is None:
            continue
            
        # Handle both datetime objects and strings
        if isinstance(date_val, datetime):
            if date_val.year == dt_obj.year and date_val.month == dt_obj.month and date_val.day == dt_obj.day:
                rows_to_delete.append(r)
        elif isinstance(date_val, str):
            # Try to parse common formats
            try:
                if "월" in date_val and "일" in date_val:
                    # Format: "02월 04일"
                    parsed = datetime.strptime(f"{dt_obj.year}년 {date_val}", "%Y년 %m월 %d일")
                    if parsed.month == dt_obj.month and parsed.day == dt_obj.day:
                        rows_to_delete.append(r)
            except:
                pass
    
    # Delete in reverse order
    for r in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(r, 1)
    
    return len(rows_to_delete)

def create_price_sheet(input_file, master_file, target_date, filter_only=False):
    print(f"Reading orders from {input_file}...")
    with open(input_file, 'r', encoding='utf-8') as f:
        order_text = f.read()
    
    ordered_keywords = parse_order_list(order_text)
    print(f"Parsed ordered keywords: {ordered_keywords}")
    
    print(f"Opening master file: {master_file}")
    wb = openpyxl.load_workbook(master_file)
    ws_form = wb["단가양식"]
    ws_price = wb["단가"]
    
    # 1. 공간 확보 (Insert Rows) 및 양식 복사
    form_max_row = ws_form.max_row
    for r in range(form_max_row, 1, -1):
        if ws_form.cell(row=r, column=3).value:
            form_max_row = r
            break
    row_count = form_max_row - 1

    if not filter_only:
        # Delete existing rows for this date first
        print(f"Deleting existing rows for {target_date}...")
        deleted_count = delete_old_date_rows(ws_price, target_date)
        print(f"✓ Deleted {deleted_count} existing rows for this date.")
        
        print(f"Inserting {row_count} rows at the top...")
        ws_price.insert_rows(2, amount=row_count)
        
        # 2. 양식 복사
        print("Copying template data...")
        dt_obj = datetime.strptime(target_date, "%Y-%m-%d")
        
        for i in range(row_count):
            src_row = 2 + i
            dest_row = 2 + i
            # Row Height Copy
            ws_price.row_dimensions[dest_row].height = ws_form.row_dimensions[src_row].height
            
            for col in range(1, 15): 
                src_cell = ws_form.cell(row=src_row, column=col)
                dest_cell = ws_price.cell(row=dest_row, column=col)
                copy_cell(src_cell, dest_cell)
                
                # 날짜 입력
                val = str(src_cell.value) if src_cell.value else ""
                is_header = "동원" in val or "재고" in val or "날짜" in val
                
                if col == 2 and not is_header:
                        item_name = ws_form.cell(row=src_row, column=3).value
                        if item_name:
                            dest_cell.value = dt_obj
                            dest_cell.number_format = 'mm"월" dd"일"'

        # 병합 셀 적용
        print("Applying merged cells...")
        for rng in ws_form.merged_cells.ranges:
            min_col, min_row, max_col, max_row = range_boundaries(str(rng))
            if min_row >= 2 and max_row <= form_max_row:
                ws_price.merge_cells(start_row=min_row, start_column=min_col,
                                    end_row=max_row, end_column=max_col)
    else:
        print("ℹ️ Skipping Insertion and Copying (--filter-only mode active)")

    # 2.5 Force Unmerge Supplier Column (Col 10) & Fill Values
    # User Request: "공급업체 절대 병합하지마" & "스타일 유지"
    print("Forcing Unmerge on Supplier Column (Col 10)...")
    
    # We must iterate a copy of the ranges because we are modifying the list
    all_merged = list(ws_price.merged_cells.ranges)
    
    for rng in all_merged:
        min_col, min_row, max_col, max_row = range_boundaries(str(rng))
        
        # Check if this merge involves Column 10
        if min_col <= 10 <= max_col and min_row >= 2:
            # Get the value AND style from the top-left cell BEFORE unmerging
            top_cell = ws_price.cell(row=min_row, column=min_col)
            top_val = top_cell.value
            
            # Unmerge
            ws_price.unmerge_cells(str(rng))
            
            # Fill the entire range with value AND copy style
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    target_cell = ws_price.cell(row=r, column=c)
                    target_cell.value = top_val
                    
                    if top_cell.has_style:
                        target_cell.font = copy(top_cell.font)
                        target_cell.border = copy(top_cell.border)
                        target_cell.fill = copy(top_cell.fill)
                        target_cell.number_format = copy(top_cell.number_format)
                        target_cell.protection = copy(top_cell.protection)
                        target_cell.alignment = copy(top_cell.alignment)

    # 3. 필터링 (V3: Supplier Matching & Deduplication)
    print("Filtering (V3 Logic)...")
    
    # Helper to parse supplier->items from text
    def parse_supplier_orders_local(text):
        scanner = {}
        current_supp = None
        lines = text.strip().split('\n')
        for line in lines:
            line = line.strip()
            if not line: continue
            if line.startswith('-'):
                # New supplier
                current_supp = line.lstrip('-').strip()
                if current_supp not in scanner: scanner[current_supp] = [] # Use List to count occurrences
                continue
            
            if not current_supp: continue
            if line.startswith('(') and '금일' in line: continue
            if line.startswith('(선발주)'): continue 
            
            parts = line.split()
            if not parts: continue
            if parts[0][0].isdigit(): continue
            
            keyword = parts[0].split('(')[0]
            
            # Additional Regex-like Cleanup (Same as global parser)
            if "양배추" in keyword: keyword = "양배추"
            if "깐양파" in keyword: keyword = "깐양파"
            if "표고" in keyword: keyword = "표고버섯"
            if "토마토" in keyword: keyword = "완숙토마토" # Default normalize
            if "완숙토마토" in keyword: keyword = "완숙토마토"
            
            # Mappings (Same as before)
            if keyword == "일자콩나물": keyword = "콩나물"
            if keyword in ["피양파", "깐양파", "양파"]: keyword = "양파"
            if keyword in ["흙대파", "깐대파", "대파"]: keyword = "대파"
            if keyword in ["깐쪽파", "쪽파"]: keyword = "쪽파"
            if keyword == "취청오이": keyword = "오이"
            if keyword == "세척무": keyword = "무"
            # 마늘: 간마늘/깐마늘 구분 유지 (뭉치지 않음)
            if keyword == "세척당근": keyword = "당근"
            if keyword == "알베기배추": keyword = "알배기"
            if keyword == "알배기": keyword = "알배기"
            if keyword == "청경채": keyword = "청경채"
            if keyword == "잎로메인": keyword = "로메인"  # 잎로메인 → 로메인(일반) 매칭용
            if keyword == "적채": keyword = "빨간양배추"  # 적채 = 빨간양배추 국산

            if keyword:
                scanner[current_supp].append(keyword) # Append allows duplicates (count based)
                # 표고버섯 발주 시 생표고 수입도 단가표에 포함
                if keyword == "표고버섯":
                    scanner[current_supp].append("생표고")
        return scanner
        return scanner

    supplier_map = parse_supplier_orders_local(order_text)
    
    # Define Garlic Exceptions
    garlic_exceptions = [
        ("상", "1K"), ("보통", "1K"), ("상", "20K"), ("상", "1kg"), ("보통", "1kg"), ("상", "20kg")
    ]
    
    def is_garlic_keep(grade, unit):
        g = str(grade).strip()
        u = str(unit).strip().upper() 
        for ex_g, ex_u in garlic_exceptions:
            if ex_g in g and ex_u.upper() in u:
                return True
        return False

    scan_end_row = 2 + row_count - 1
    rows_to_delete = set()
    candidates = {}
    current_supplier_val = ""

    for r in range(2, scan_end_row + 1):
        item_val = ws_price.cell(row=r, column=3).value
        grade_val = ws_price.cell(row=r, column=4).value
        unit_val = ws_price.cell(row=r, column=5).value
        supp_val = ws_price.cell(row=r, column=10).value
        
        # Update current supplier if value exists
        if supp_val:
            current_supplier_val = str(supp_val).strip()
        
        cat_val = ws_price.cell(row=r, column=1).value
        cat_str = str(cat_val).strip() if cat_val else ""

        supp_str = current_supplier_val
        item_str = str(item_val).strip() if item_val else ""
        grade_str = str(grade_val).strip() if grade_val else ""
        unit_str = str(unit_val).strip() if unit_val else ""
        
        # 1. Protection Logic
        is_protected = False
        check_text = f"{cat_str} {item_str} {supp_str}"
        
        # User defined protection
        if "소분" in check_text: is_protected = True
        
        # Header Protection (New) - Preserves '공급업체' row above Dongwon
        if "공급업체" in check_text: is_protected = True
        
        # Legacy protection (just in case)
        if "식봄" in check_text: is_protected = True

        if is_protected:
            continue
        
        
        # 2. Supplier Match (with Aliasing)
        matched_supp_key = None
        
        # Aliases Map (Template Name <-> Order Name)
        # Order Name (Key in map) -> Template Name (Value in Excel)
        # We need to map Excel Name -> Order Name to find key in map
        # Or check if keys in map match Excel Name OR its aliases
        
        SUPPLIER_ALIASES = {
            "경향농산": ["다모아버섯", "경향농산"],
            "다모아버섯": ["경향농산", "다모아버섯"],
            # Add more common aliases
            "풍경농산": ["풍경", "풍경농산"],
            "명진농산": ["명진", "명진농산"],
            "영운농산": ["영운", "영운농산"],
            "나물향기": ["나물향기", "나물"],
            "태현상회": ["태현", "태현상회"],
            "초원농산": ["경북상회", "초원농산"],
            "경북상회": ["초원농산", "경북상회"],
        }
        
        # Iterate through keys in map (Order Suppliers)
        for s_key in supplier_map.keys():
            if not s_key: continue
            
            # 1. Direct Match
            if s_key in supp_str:
                matched_supp_key = s_key
                break
                
            # 2. Alias Match
            # If supp_str is "경향농산", check if s_key ("다모아버섯") is an alias
            if supp_str in SUPPLIER_ALIASES:
                possible_aliases = SUPPLIER_ALIASES[supp_str]
                if s_key in possible_aliases:
                    matched_supp_key = s_key
                    break

        
        if not matched_supp_key:
            # Fallback: 공급처 매칭 실패해도 ordered_keywords에 있으면 유지
            # (템플릿 공급처 ≠ 발주 공급처인 경우: 치커리, 청양고추 등)
            item_in_orders = False
            for kw in ordered_keywords:
                if kw in item_str or item_str in kw:
                    item_in_orders = True
                    break
            if not item_in_orders:
                rows_to_delete.add(r)
            continue

        # 3. Item Match (Bidirectional)
        # Check if any keyword in supplier_map[matched_supp_key] is in item_str
        # OR if item_str contains any keyword
        matched_item_key = None

        # Item aliases: 주문 키워드 → 템플릿 매칭용 키워드 (이름 순서 차이 등)
        ITEM_MATCH_ALIASES = {
            "완숙토마토": "토마토",  # Template: "토마토 완숙"
        }

        # Match block list (양방향 모두 적용)
        MATCH_BLOCK = {
            "근대": ["적근대"],   # "근대" 템플릿이 "적근대" 키워드를 잡지 않도록
            "오이": ["백다다기오이"],  # "취청오이→오이"가 "백다다기오이"를 잡지 않도록
        }

        # Pass 1: Forward match (keyword in Excel item name) - more precise
        for k in supplier_map[matched_supp_key]:
            k_alias = ITEM_MATCH_ALIASES.get(k, k)
            if k in item_str or (k_alias != k and k_alias in item_str):
                # 차단 목록 확인
                blocked = MATCH_BLOCK.get(k, [])
                if item_str in blocked:
                    continue
                matched_item_key = k
                break
        # Pass 2: Reverse match only if forward failed (e.g. "깐대파" in order, "대파" in template)
        # 단, 접두사가 다른 별개 품목은 제외 (적근대≠근대, 적채≠배추 등)
        if not matched_item_key:
            for k in supplier_map[matched_supp_key]:
                if item_str in k:
                    # 역방향 매칭 차단 목록 확인
                    blocked = MATCH_BLOCK.get(item_str, [])
                    if k in blocked:
                        continue
                    matched_item_key = k
                    break
        
        if not matched_item_key:
            rows_to_delete.add(r)
            continue
            
        # 4. Garlic Special Logic
        if "마늘" in matched_item_key or "마늘" in item_str:
            if is_garlic_keep(grade_str, unit_str):
                continue
        
        # 5. Add to Candidates for Deduplication
        pair_key = (matched_supp_key, matched_item_key)
        if pair_key not in candidates: candidates[pair_key] = []
        
        candidates[pair_key].append({
            'row': r,
            'grade': grade_str,
            'score': get_grade_priority(grade_str)
        })

    # Deduplication Execution
    for key, group in candidates.items():
        # key is (SupplierName, ItemName)
        supp_key, item_key = key
        
        # Determine how many times this item was requested by this supplier
        requested_count = 1
        if supp_key in supplier_map:
             # Count occurrences in the list
             requested_count = supplier_map[supp_key].count(item_key)
             if requested_count == 0: requested_count = 1 # Fallback
        
        # If we have more candidates than requested count, we delete the excess (worse grades)
        if len(group) > requested_count:
            group.sort(key=lambda x: x['score']) # Score 1 is best
            
            # Keep top 'requested_count' items
            # Delete indices from [requested_count : end]
            for other in group[requested_count:]:
                rows_to_delete.add(other['row'])

    # Track what we kept for Verification
    kept_items = []

    for r in range(2, ws_price.max_row + 1):
        if r in rows_to_delete: continue
        
        # We need to capture what's actually on the sheet to verify against scanner
        # But wait, we haven't deleted yet.
        pass

    # Execute Delete
    for r in sorted(list(rows_to_delete), reverse=True):
        ws_price.delete_rows(r, 1)

    print(f"✓ Deleted {len(rows_to_delete)} rows (V3 Logic).")
    
    # --- VERIFICATION STEP ---
    print("\n🔍 Verifying Missing Items...")
    
    # 1. Scan the FINAL sheet state to see what exists
    final_inventory = {} # { Supplier : [Items...] }
    
    max_r = ws_price.max_row
    current_supp = ""
    
    for r in range(2, max_r + 1):
        val_supp = ws_price.cell(row=r, column=10).value
        # If merged, we might need to handle logic, but we forced unmerge earlier!
        # Good thing we did that.
        
        if val_supp:
            current_supp = str(val_supp).strip()
            
        val_item = ws_price.cell(row=r, column=3).value
        val_item_str = str(val_item).strip() if val_item else ""
        
        if not current_supp or not val_item_str: continue
        if "공급업체" in current_supp: continue # Header
        
        if current_supp not in final_inventory: final_inventory[current_supp] = []
        final_inventory[current_supp].append(val_item_str)

    # 2. Compare scanner (Requests) vs final_inventory (Result)
    missing_report = []
    
    for req_supp, req_items in supplier_map.items():
        # Resolve Aliases for Verification
        # We need to guess what 'req_supp' maps to in 'final_inventory'
        # Try Direct, Try Aliases
        
        target_inventory_items = []
        
        # Try finding the supplier in inventory
        found_supp_key = None
        if req_supp in final_inventory:
            found_supp_key = req_supp
        else:
            # Check aliases
            if req_supp in SUPPLIER_ALIASES:
                for alias in SUPPLIER_ALIASES[req_supp]:
                    if alias in final_inventory:
                        found_supp_key = alias
                        break
            # Check partial match (e.g. 명진농산 -> 명진)
            if not found_supp_key:
                for inv_supp in final_inventory.keys():
                    if inv_supp in req_supp or req_supp in inv_supp:
                        found_supp_key = inv_supp
                        break
        
        if found_supp_key:
            target_inventory_items = final_inventory[found_supp_key]
        else:
            # Supplier totally missing
            missing_report.append(f"🔴 Supplier Missing: {req_supp} (Items: {req_items})")
            continue
            
        # Check Items
        # req_items is a list e.g. ['대파', '대파', '쪽파']
        # target_inventory_items is list of strings in Excel
        
        import copy
        inv_copy = copy.deepcopy(target_inventory_items)
        
        for req_item in req_items:
            # Find fuzzy match in inventory
            match_found = False
            match_idx = -1
            
            for i, inv_item in enumerate(inv_copy):
                # Logic must match the filtering logic!
                # mapping reverses? "대파" matches "깐대파", "흙대파"
                
                # Simple check: inclusion
                if req_item in inv_item: 
                    match_found = True
                    match_idx = i
                    break
                    
            if match_found:
                # Consume this item so we don't double count
                inv_copy.pop(match_idx)
            else:
                missing_report.append(f"❌ Missing Item: [{req_supp}] {req_item}")

    if missing_report:
        print("\n" + "="*40)
        print("⚠️  MISSING ITEMS REPORT")
        print("="*40)
        for line in missing_report:
            print(line)
        print("="*40 + "\n")
    else:
        print("\n✅ All requested items verified present!\n")

    print(f"Saving to {master_file}...")
    wb.save(master_file)
    print("✓ Done! Inserted, Copied, and Filtered (V3).")

    # 4. Final Cleanup & Layout Normalization (Deletion Only)
    print("Performing Final Cleanup (Deletion Only)...")
    
    cleanup_rows = set()
    max_r = ws_price.max_row
    dt_obj = datetime.strptime(target_date, "%Y-%m-%d")

    for r in range(2, max_r + 1):
        item_val = ws_price.cell(row=r, column=3).value
        item_str = str(item_val).strip() if item_val else ""
        
        # Check Section Headers (Preserve)
        row_values = [str(ws_price.cell(row=r, column=c).value) for c in range(1, 15)]
        row_text = "".join(row_values)
        
        is_section_header = False
        if "재고" in row_text and "소분" in row_text: is_section_header = True
        if "식봄" in row_text: is_section_header = True
        
        # Protect Dongwon Header here too
        if "공급업체" in row_text: is_section_header = True

        if is_section_header:
            continue
            
        # Delete conditions
        if item_str == "품목": # Repeated Header
            cleanup_rows.add(r)
            continue
            
        if not item_str: # Empty Item
            cleanup_rows.add(r)
            continue
            
    # Execute Deletion
    for r in sorted(list(cleanup_rows), reverse=True):
        ws_price.delete_rows(r, 1)
        
    print(f"✓ Cleanup: Deleted {len(cleanup_rows)} empty/garbage rows.")

    # 3. Final Polish (Dates only) - 당일 삽입 행만 대상, 다른 날짜 행은 절대 건드리지 않음
    print("Polishing Layout (Dates only - today's rows only)...")
    current_max_row = ws_price.max_row

    for r in range(2, current_max_row + 1):
        # 다른 날짜 행은 건드리지 않음 (핵심 수정: 기존 데이터 보호)
        date_val = ws_price.cell(row=r, column=2).value
        if isinstance(date_val, datetime) and date_val.date() != dt_obj.date():
            continue

        item_val = ws_price.cell(row=r, column=3).value
        item_str = str(item_val) if item_val else ""
        if "단가" in item_str or "매입가" in item_str: continue

        row_vals = [str(ws_price.cell(row=r, column=c).value) for c in range(1, 15)]
        row_txt = "".join(row_vals)
        if "재고" in row_txt and "소분" in row_txt: continue
        if "공급업체" in row_txt: continue

        # Fix date clearing on header lines matching logic
        if not item_val: continue

        # Set Date (Column 2) if row is valid item (당일 또는 날짜 없는 행만)
        if item_val:
            date_cell = ws_price.cell(row=r, column=2)
            date_cell.value = dt_obj
            date_cell.number_format = 'mm"월" dd"일"'
            date_cell.alignment = Alignment(horizontal='center', vertical='center')

    print(f"Saving to {master_file}...")
    wb.save(master_file)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="Path to processed orders file")
    parser.add_argument("--master", required=True, help="Path to master Excel file")
    parser.add_argument("--date", required=True, help="Target date (YYYY-MM-DD)")
    parser.add_argument("--filter-only", action="store_true", help="Skip insertion and only filter existing rows")
    
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        print(f"❌ Error: Input file not found: {args.input}")
        exit(1)
        
    if not os.path.exists(args.master):
        print(f"❌ Error: Master file not found: {args.master}")
        exit(1)
    
    import sys
    try:
        create_price_sheet(args.input, args.master, args.date, filter_only=args.filter_only)
    except SystemExit:
        raise
    except Exception:
        import traceback; traceback.print_exc()
        sys.exit(1)

    from _notify import notify
    notify("stage3_price", date_str=args.date)
