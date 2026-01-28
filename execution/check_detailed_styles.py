import openpyxl
import os

master_file_path = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터-AI접근용.xlsx"
sheet_name = "발주"

def check_formulas_and_formats(row=19000):
    if not os.path.exists(master_file_path):
        print(f"Error: Master file not found at {master_file_path}")
        return

    try:
        wb = openpyxl.load_workbook(master_file_path, data_only=False)
        ws = wb[sheet_name]
        
        print(f"Analyzing row {row}...")
        
        with open(".tmp/detailed_format_check.txt", "w", encoding="utf-8") as f:
            for col in range(1, 16):
                cell = ws.cell(row=row, column=col)
                f.write(f"Column {cell.column_letter}:\n")
                f.write(f"  Value/Formula: {cell.value}\n")
                f.write(f"  Number Format: {cell.number_format}\n")
                f.write(f"  Font: {cell.font.name}\n")
                f.write("-" * 20 + "\n")
                
        print("✓ Detailed check written to .tmp/detailed_format_check.txt")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_formulas_and_formats(19000)
