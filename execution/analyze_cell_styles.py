import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment
import os

master_file_path = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터-AI접근용.xlsx"
sheet_name = "발주"

def analyze_styles(target_row=19408):
    if not os.path.exists(master_file_path):
        print(f"Error: Master file not found at {master_file_path}")
        return

    try:
        wb = openpyxl.load_workbook(master_file_path, data_only=False)
        ws = wb[sheet_name]
        
        print(f"Analyzing Styles for row {target_row}...")
        
        with open(".tmp/style_analysis.txt", "w", encoding="utf-8") as f:
            for col in range(1, 16): # Analyze columns A to O
                cell = ws.cell(row=target_row, column=col)
                f.write(f"Column {cell.column_letter}:\n")
                f.write(f"  Value: {cell.value}\n")
                f.write(f"  Font: Name={cell.font.name}, Size={cell.font.sz}, Bold={cell.font.b}, Color={cell.font.color.rgb if cell.font.color else 'None'}\n")
                f.write(f"  Border: Top={cell.border.top.style}, Bottom={cell.border.bottom.style}, Left={cell.border.left.style}, Right={cell.border.right.style}\n")
                f.write(f"  Alignment: Horizontal={cell.alignment.horizontal}, Vertical={cell.alignment.vertical}\n")
                f.write("-" * 20 + "\n")
                
        print("✓ Style analysis written to .tmp/style_analysis.txt")

    except Exception as e:
        print(f"Error analyzing styles: {e}")

if __name__ == "__main__":
    analyze_styles(19408)
