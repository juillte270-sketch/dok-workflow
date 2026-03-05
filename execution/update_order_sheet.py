import pandas as pd
import os
import sys
import json
from datetime import datetime
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

# Paths
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MAPPINGS_PATH = os.path.join(project_root, "skills", "order_processing", "resources", "mappings.json")

def _load_supplier_names():
    """Load canonical supplier names from mappings.json."""
    try:
        with open(MAPPINGS_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
        names = set(data.get('supplier_names', []))
    except Exception:
        names = set()
    # Always include these special categories
    names.update(["재고, 창고소분", "시장구매, 시장소분"])
    return names
DEFAULT_MASTER_FILE = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx"
sheet_name = "발주"

def parse_processed_orders(target_date_str):
    orders = []
    # Use date-specific file path
    processed_orders_path = os.path.join(project_root, "data", "inputs", f"processed_orders_{target_date_str}.txt")
    
    if not os.path.exists(processed_orders_path):
        print(f"Error: {processed_orders_path} not found.")
        print(f"Please run: python execution/process_orders.py --date {target_date_str[:4]}-{target_date_str[4:6]}-{target_date_str[6:8]}")
        return orders
    
    print(f"✓ Reading from: {processed_orders_path}")

    current_store = None
    in_supplier_section = False
    with open(processed_orders_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    import re
    # Pattern for items: "Item Name QtyUnit" (e.g., "청경채 4박스")
    item_pattern = re.compile(r'^(.+?)\s+(\d+(?:\.\d+)?)([a-zA-Z가-힣]+).*$')
    # Fallback: 공백 없이 붙어있는 경우 (예: 스테비아토마토1박스)
    item_pattern_nospace = re.compile(r'^(.+?)(\d+(?:\.\d+)?)([a-zA-Z가-힣]+).*$')

    # Supplier names loaded from canonical source (mappings.json)
    supplier_names = _load_supplier_names()

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if line.startswith('-'):
            store_name = line.lstrip('-').strip()
            # Once we hit the first supplier (영운농산), everything after is supplier data
            if store_name in supplier_names:
                in_supplier_section = True
                current_store = None
                continue
            # If already in supplier section, skip all remaining headers
            if in_supplier_section:
                current_store = None
                continue
            current_store = store_name
            continue

        if current_store:
            # Skip memos in parentheses
            if line.startswith('('):
                continue
                
            match = item_pattern.match(line)
            if not match:
                match = item_pattern_nospace.match(line)
            if match:
                item_name = match.group(1).strip()
                qty = float(match.group(2))
                unit = match.group(3).strip()

                orders.append({
                    "매장명": current_store,
                    "품목": item_name,
                    "입고수량": qty,
                    "단위": unit
                })
    return orders

def update_order_sheet():
    # Argument Parsing
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", help="Target date (YYYY-MM-DD)", default=None)
    parser.add_argument("--start-row", help="Force start row (integer)", default=None)
    parser.add_argument("--master-file", help="Path to master excel file", default=None)
    # Use parse_known_args in case it's called with other args
    args, _ = parser.parse_known_args()

    if args.master_file:
        master_file_path = args.master_file
        print(f"Using provided master file: {master_file_path}")
    else:
        master_file_path = DEFAULT_MASTER_FILE

    if not args.date:
        print("Error: --date argument is required (format: YYYY-MM-DD)")
        return
    
    target_date_str = args.date.replace("-", "")
    
    # Validation: Check if processed orders file exists
    processed_orders_path = os.path.join(project_root, "data", "inputs", f"processed_orders_{target_date_str}.txt")
    if not os.path.exists(processed_orders_path):
        print(f"❌ Error: Processed orders file not found: {processed_orders_path}")
        print(f"Please run first: python execution/process_orders.py --input inputs/orders_{target_date_str}.txt --date {args.date}")
        return
    
    orders = parse_processed_orders(target_date_str)
    if not orders:
        print("No orders to enter.")
        return
    
    # Validation: Verify we have store orders (not just supplier data)
    store_count = len(set(order['매장명'] for order in orders))
    print(f"✓ Validation passed: Found {len(orders)} items from {store_count} stores")

    if not os.path.exists(master_file_path):
        print(f"Error: Master file not found at {master_file_path}")
        return

    from openpyxl.styles import Font, Border, Side, Alignment
    from datetime import datetime

    # Style definitions
    font_malgun = Font(name='Malgun Gothic', sz=11)
    font_calibri = Font(name='Calibri', sz=11)
    align_center = Alignment(horizontal='center', vertical='center')
    border_double_bottom = Border(bottom=Side(style='double'))
    border_none = Border()
    
    # Formats
    date_format = 'mm"월" dd"일"'
    price_format = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@'
    rate_format = '0.0'

    print(f"Updating 85 rows in {master_file_path} starting at row 19409...")

    # Set fixed date per user request: "1월 32일" -> Mapping to 2026-01-31 for validity
    # Set date per user request or default to now
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", help="Target date (YYYY-MM-DD)", default=None)
    parser.add_argument("--start-row", help="Force start row (integer)", default=None)
    parser.add_argument("--master-file", help="Path to master excel file", default=None)
    # Use parse_known_args in case it's called with other args
    args, _ = parser.parse_known_args()
    
    if args.master_file:
        master_file_path = args.master_file
        print(f"Using provided master file: {master_file_path}")
    else:
        master_file_path = DEFAULT_MASTER_FILE

    if args.date:
         target_date_obj = datetime.strptime(args.date, "%Y-%m-%d")
         print(f"Using provided target date: {target_date_obj.date()}")
    else:
         target_date_obj = datetime.now()

    year = target_date_obj.year
    month = target_date_obj.month

    try:
        wb = openpyxl.load_workbook(master_file_path)
        ws = wb[sheet_name]
        
        # Find range of existing data for this date
        existing_rows = []
        max_r_scan = ws.max_row
        
        # Determine where relevant data ends (optimization)
        # Scan backward from max row
        final_data_row = 0
        for r in range(max_r_scan, 1, -1):
            if ws.cell(row=r, column=1).value is not None:
                final_data_row = r
                break
        
        # Scan for existing target date
        print(f"Scanning {final_data_row} rows for existing data of {target_date_obj.date()}...")
        for r in range(2, final_data_row + 1):
             val = ws.cell(row=r, column=4).value
             match = False
             if isinstance(val, datetime):
                 if val.date() == target_date_obj.date():
                     match = True
             elif isinstance(val, str):
                 # Handle string formats if necessary, e.g. "2026-02-01"
                 try:
                     if val.startswith(target_date_str):
                         match = True
                 except: pass
             
             if match:
                 existing_rows.append(r)
        
        if existing_rows:
            # Check for continuity
            is_continuous = True
            if len(existing_rows) > 1:
                for i in range(len(existing_rows) - 1):
                    if existing_rows[i+1] != existing_rows[i] + 1:
                        is_continuous = False
                        break
            
            if not is_continuous:
                print("Warning: Existing data for this date is likely scattered. Deleting all found rows.")
                # Delete from bottom up to avoid index shift issues
                for r in reversed(existing_rows):
                    ws.delete_rows(r, 1)
                start_row = existing_rows[0] # Insert at the first occurrence
            else:
                print(f"Found {len(existing_rows)} existing rows ({existing_rows[0]}~{existing_rows[-1]}). Replacing...")
                ws.delete_rows(existing_rows[0], len(existing_rows))
                start_row = existing_rows[0]
        else:
            print("No existing data for this date. Appending...")
            start_row = final_data_row + 1

        print(f"Inserting {len(orders)} new rows at row {start_row}...")
        ws.insert_rows(start_row, amount=len(orders))
        
        current_row = start_row

        # Group orders by store to determine ranges
        from collections import defaultdict
        grouped_orders = defaultdict(list)
        all_store_names = set()
        
        for order in orders:
            s_name = order['매장명']
            all_store_names.add(s_name)
            grouped_orders[s_name].append(order)
        
        # Define Sorting Logic
        # 1. Shabuyaki
        # 2. Bueong-i Sanjang
        # 3. Rest (Alphabetical, excluding Dongwon)
        # 4. Dongwon 1st
        
        shabu_stores = sorted([s for s in all_store_names if "샤브야키" in s])
        owl_stores = sorted([s for s in all_store_names if "부엉이산장" in s])
        dongwon_stores = sorted([s for s in all_store_names if "동원1차" in s])
        other_stores = sorted([s for s in all_store_names if "샤브야키" not in s and "부엉이산장" not in s and "동원1차" not in s])
        
        store_order = shabu_stores + owl_stores + other_stores + dongwon_stores
        
        print("Insertion Order:")
        print(f"  1. Shabuyaki: {shabu_stores}")
        print(f"  2. Owl: {owl_stores}")
        print(f"  3. Other: {other_stores}")
        print(f"  4. Dongwon: {dongwon_stores}")
        
        for s_name in store_order:
            store_items = grouped_orders[s_name]
            store_start_row = current_row
            store_end_row = current_row + len(store_items) - 1
            
            # Aggregate Formula: (Sum(Sales)-Sum(Cost))/Sum(Sales)
            # 이익률 = (판매-매입) / 판매 (마진율)
            # Range for this store: row {store_start_row} to {store_end_row}
            # Cost Range: J{store_start_row}:J{store_end_row}
            # Sales Range: L{store_start_row}:L{store_end_row}
            
            profit_formula = (
                f"=IF(SUM(L{store_start_row}:L{store_end_row})>0,"
                f"(SUM(L{store_start_row}:L{store_end_row})-SUM(J{store_start_row}:J{store_end_row}))/SUM(L{store_start_row}:L{store_end_row})*100,0)"
            )
            
            for idx, order in enumerate(store_items):
                # Calculate profit formula only for the first row of the store
                profit_val = profit_formula if idx == 0 else ""

                # Fill data
                cells_data = {
                    1: (order['매장명'], font_malgun, 'General'),
                    2: (year, font_calibri, 'General'),
                    3: (month, font_calibri, 'General'),
                    4: (target_date_obj, font_calibri, date_format),
                    5: (order['품목'], font_malgun, 'General'),
                    7: (order['단위'], font_malgun, 'General'),
                    8: (order['입고수량'], font_malgun, 'General'),
                    9: (None, font_calibri, price_format), # 매입단가
                    10: (f"=H{current_row}*I{current_row}", font_calibri, price_format), # 총 금액
                    11: (None, font_calibri, price_format), # 판매단가
                    12: (f"=K{current_row}*H{current_row}", font_calibri, price_format), # 총 판매금액
                    13: (f"=IF(L{current_row}>0,(L{current_row}-J{current_row})/L{current_row}*100,0)", font_calibri, rate_format), # 마진율 (Per item)
                    14: (profit_val, font_calibri, rate_format) # 이익률 (Aggregate per store, 1st row only)
                }
                
                for col, (val, font, fmt) in cells_data.items():
                    cell = ws.cell(row=current_row, column=col, value=val)
                    cell.font = font
                    cell.alignment = align_center
                    cell.number_format = fmt
                    cell.border = border_none
                
                # Empty columns
                for col in [6, 15]:
                    cell = ws.cell(row=current_row, column=col, value=None)
                    cell.alignment = align_center
                    cell.border = border_none

                current_row += 1
            
            # Apply double border to the last row of the store
            for col in range(1, 15):
                ws.cell(row=store_end_row, column=col).border = border_double_bottom

        wb.save(master_file_path)
        print(f"✓ Successfully inserted {len(orders)} rows starting from row {start_row} with Aggregate Profit Rate.")

        from _notify import notify
        notify("stage2_order", date_str=args.date)

    except Exception as e:
        import traceback
        print(f"Error updating excel: {e}")
        traceback.print_exc()
        import sys; sys.exit(1)

if __name__ == "__main__":
    import sys
    try:
        update_order_sheet()
    except SystemExit:
        raise
    except Exception:
        import traceback; traceback.print_exc()
        sys.exit(1)
