import pandas as pd
import os

file_path = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx"
sheet_name = "발주"

def analyze_sheet():
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return

    print(f"Analyzing file: {file_path}")
    try:
        # Load only the first few rows to save memory/time
        df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=10)
        
        with open(".tmp/master_analysis.txt", "w", encoding="utf-8") as f:
            f.write(f"Columns: {df.columns.tolist()}\n\n")
            f.write("First 5 rows:\n")
            f.write(df.head().to_string())
            
        print("✓ Analysis written to .tmp/master_analysis.txt")

    except Exception as e:
        print(f"Error reading excel: {e}")

if __name__ == "__main__":
    analyze_sheet()
