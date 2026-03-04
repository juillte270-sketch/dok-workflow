import openpyxl
import os
import argparse
from datetime import datetime

def parse_order_list(order_text):
    """
    주문 목록 텍스트를 파싱하여 발주된 핵심 품목 키워드를 추출합니다.
    (create_price_sheet.py와 동일한 로직, 디버깅 용)
    """
    ordered_keywords = set()
    
    # 무시할 단어들
    ignore_words = ['박스', '봉', '망', 'kg', 'kg', 'g', '개', '팩', '단', '관', '수', '국산', '수입']
    
    lines = order_text.strip().split('\n')
    
    for line in lines:
        trimmed = line.strip()
        if not trimmed: continue
        if trimmed.startswith('-'): continue
        if trimmed.startswith('(') and '금일' in trimmed: continue
        
        words = trimmed.split()
        if not words: continue
        
        first_word = words[0]
        if first_word.isdigit(): continue
        
        keyword = first_word.split('(')[0]
        
        # Mapping rules
        if keyword == "일자콩나물": keyword = "콩나물"
        if keyword == "피양파" or keyword == "깐양파": keyword = "양파"
        if keyword == "흙대파" or keyword == "깐대파": keyword = "대파"
        if keyword == "깐쪽파": keyword = "쪽파"
        if keyword == "취청오이": keyword = "오이"
        if keyword == "세척무": keyword = "무"
        if keyword == "간마늘" or keyword == "깐마늘": keyword = "마늘"
        if keyword == "세척당근": keyword = "당근"
        
        ordered_keywords.add(keyword)
        
    # Additional mappings
    base_keywords = list(ordered_keywords)
    for k in base_keywords:
        if k == "양파": 
            ordered_keywords.add("깐양파"); ordered_keywords.add("피양파")
        if k == "대파":
            ordered_keywords.add("흙대파"); ordered_keywords.add("깐대파")
        if k == "마늘":
            ordered_keywords.add("간마늘"); ordered_keywords.add("깐마늘")
        if k == "오이":
            ordered_keywords.add("취청오이")
            
    return ordered_keywords

def clean_price_sheet(master_file, order_file, target_date):
    print(f"Cleaning price sheet in {master_file} for {target_date}...")
    
    with open(order_file, 'r', encoding='utf-8') as f:
        order_text = f.read()
    
    ordered_keywords = parse_order_list(order_text)
    print(f"Target Keywords: {ordered_keywords}")
    
    wb = openpyxl.load_workbook(master_file)
    ws = wb["단가"]
    
    # Parse target date
    target_dt = datetime.strptime(target_date, "%Y-%m-%d")
    
    rows_to_delete = set()
    
    # Identify rows to verify
    # Scan from top (row 2) until we hit a date or empty that is NOT our target date.
    # Since we insert at top, our target date rows should be continuous from row 2.
    
    print("Scanning rows...")
    current_row = 2
    match_count = 0
    
    while True:
        # Safety break
        if current_row > 1000: break
        
        date_cell = ws.cell(row=current_row, column=2)
        item_cell = ws.cell(row=current_row, column=3)
        
        val = date_cell.value
        item_name = str(item_cell.value).strip() if item_cell.value else ""
        
        if not val and not item_name:
            # End of data potentially
            # But assume we stop when date changes or is significantly old
            # However, if date is None, it might be a header within the block?
            # Re-check logic: inserted rows have date.
            pass
            
        is_target_date = False
        if isinstance(val, datetime):
            if val.date() == target_dt.date():
                is_target_date = True
        elif isinstance(val, str):
            if target_date in val:
                is_target_date = True
        
        # If we encounter a row with a date DIFFERENT from target date, we stop scanning?
        # Assuming the new block is at the top.
        if val and not is_target_date and (isinstance(val, datetime) or (isinstance(val, str) and "-" in val)):
            print(f"Encountered different date at row {current_row}: {val}. Stopping scan.")
            break
            
        # Only process if it LOOKS like part of our block (date matches OR empty date but valid item)
        # Note: In 'create_price_sheet', we put date for every item row.
        # But section headers might not have date.
        
        if is_target_date:
            match_count += 1
            
            # Filtering Logic
            keep = False
            
            # 1. Always keep special sections
            if "재고" in item_name or "소분" in item_name: keep = True
            elif "동원" in item_name: keep = True
            elif "날짜" in item_name: keep = True # Header row
            
            # 2. Check keywords
            else:
                for k in ordered_keywords:
                    if k in item_name:
                        keep = True
                        break
            
            if not keep:
                print(f"  [DELETE] Row {current_row}: {item_name} (Not in order list)")
                rows_to_delete.add(current_row)
            else:
                 # Debug: print what we kept
                 # print(f"  [KEEP]   Row {current_row}: {item_name}")
                 pass

        current_row += 1
        
    if not rows_to_delete:
        print("No rows found to delete.")
        return

    # Delete rows (reverse order)
    sorted_rows = sorted(list(rows_to_delete), reverse=True)
    for r in sorted_rows:
        ws.delete_rows(r, 1)
        
    print(f"✓ Deleted {len(sorted_rows)} rows.")
    wb.save(master_file)
    print("File saved.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--master", required=True)
    parser.add_argument("--orders", required=True)
    parser.add_argument("--date", required=True)
    args = parser.parse_args()
    
    clean_price_sheet(args.master, args.orders, args.date)
