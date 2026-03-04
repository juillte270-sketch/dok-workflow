import openpyxl
from datetime import datetime

def clean_today(master_path, target_date_str):
    print(f"Cleaning rows for {target_date_str} in {master_path}...")
    wb = openpyxl.load_workbook(master_path)
    ws = wb["단가"]
    
    # Parse date to match Excel format if necessary, but usually string match or obj match
    # Excel date might be datetime.
    
    rows_to_delete = []
    
    # Scan top 500 rows
    for r in range(2, 501):
        cell_val = ws.cell(row=r, column=2).value
        # Check if matches target date
        is_match = False
        if isinstance(cell_val, datetime):
            if cell_val.strftime("%Y-%m-%d") == target_date_str:
                is_match = True
        elif isinstance(cell_val, str):
            # Checking for '02월 03일' format or '2026-02-03'
            if "02월 03일" in cell_val: # This is the format seen in screenshots
                is_match = True
        
        if is_match:
            rows_to_delete.append(r)
            
    if not rows_to_delete:
        print("No rows found for today to clean.")
    else:
        print(f"Deleting {len(rows_to_delete)} rows...")
        for r in sorted(rows_to_delete, reverse=True):
            ws.delete_rows(r, 1)
        wb.save(master_path)
        print("Cleanup done.")

if __name__ == "__main__":
    clean_today(r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx", "2026-02-03")
