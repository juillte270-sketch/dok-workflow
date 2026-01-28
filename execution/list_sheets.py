import pandas as pd
import os

file_path = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx"

def list_all_sheets():
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return

    try:
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names
        
        with open(".tmp/sheets_summary.txt", "w", encoding="utf-8") as f:
            f.write(f"Sheet names: {sheet_names}\n\n")
            for name in sheet_names:
                df = xl.parse(name, nrows=2)
                f.write(f"--- Sheet: {name} ---\n")
                f.write(f"Columns: {df.columns.tolist()}\n\n")
                
        print("✓ Sheets summary written to .tmp/sheets_summary.txt")

    except Exception as e:
        print(f"Error reading excel: {e}")

if __name__ == "__main__":
    list_all_sheets()
