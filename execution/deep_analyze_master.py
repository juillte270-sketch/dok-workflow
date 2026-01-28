import pandas as pd
import os

file_path = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx"
sheet_name = "발주"

def deep_analyze():
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return

    try:
        # Read the first 20 rows without headers to see where they are
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=20)
        
        with open(".tmp/deep_analysis.txt", "w", encoding="utf-8") as f:
            f.write("Raw rows 1-20 (No Header Mode):\n")
            f.write(df.to_string())
            
        print("✓ Deep analysis written to .tmp/deep_analysis.txt")

    except Exception as e:
        print(f"Error reading excel: {e}")

if __name__ == "__main__":
    deep_analyze()
