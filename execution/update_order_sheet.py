import pandas as pd
import os
from datetime import datetime
import openpyxl

# Paths
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
processed_orders_path = os.path.join(project_root, "data", "inputs", "processed_orders.txt")
master_file_path = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터-AI접근용.xlsx"
sheet_name = "발주"

def parse_processed_orders():
    orders = []
    if not os.path.exists(processed_orders_path):
        print(f"Error: {processed_orders_path} not found.")
        return orders

    current_store = None
    with open(processed_orders_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    import re
    # Pattern for items: "Item Name QtyUnit" (e.g., "청경채 4박스")
    item_pattern = re.compile(r'^(.+?)\s+(\d+(?:\.\d+)?)([a-zA-Z가-힣]+).*$')

    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        if line.startswith('-'):
            store_name = line.lstrip('-').strip()
            # Stop if we hit supplier sections or stock/market
            if store_name in ["영운농산", "다모아버섯", "건영농산", "오복상회", "명진농산", "풍경농산", 
                          "나물향기", "태현상회", "가야웰빙", "이모유통", "초원농산", "소리농산", 
                          "명화농산", "우신농산", "정복상회", "정운", "재고, 창고소분", "시장구매, 시장소분"]:
                current_store = None
                continue
            current_store = store_name
            continue

        if current_store:
            # Skip memos in parentheses
            if line.startswith('('):
                continue
                
            match = item_pattern.match(line)
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
    orders = parse_processed_orders()
    if not orders:
        print("No orders to enter.")
        return

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
    # target_date_obj = datetime(2026, 1, 31)
    target_date_obj = datetime.now()
    year = target_date_obj.year
    month = target_date_obj.month

    try:
        wb = openpyxl.load_workbook(master_file_path)
        ws = wb[sheet_name]
        
        start_row = 19409
        
        # Group orders by store to determine ranges
        from collections import defaultdict
        grouped_orders = defaultdict(list)
        # Orders are already sorted by store roughly, but let's be safe and preserve order
        store_order = []
        for order in orders:
            s_name = order['매장명']
            if s_name not in store_order:
                store_order.append(s_name)
            grouped_orders[s_name].append(order)
            
        current_row = start_row
        
        for s_name in store_order:
            store_items = grouped_orders[s_name]
            store_start_row = current_row
            store_end_row = current_row + len(store_items) - 1
            
            # Aggregate Formula: (Sum(Sales)-Sum(Cost))/Sum(Cost)
            # Range for this store: row {store_start_row} to {store_end_row}
            # Cost Range: J{store_start_row}:J{store_end_row}
            # Sales Range: L{store_start_row}:L{store_end_row}
            
            profit_formula = (
                f"=IF(SUM(J{store_start_row}:J{store_end_row})>0,"
                f"(SUM(L{store_start_row}:L{store_end_row})-SUM(J{store_start_row}:J{store_end_row}))/SUM(J{store_start_row}:J{store_end_row})*100,0)"
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

    except Exception as e:
        import traceback
        print(f"Error updating excel: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    update_order_sheet()
