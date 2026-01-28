import openpyxl
import os

master_file_path = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터-AI접근용.xlsx"
sheet_name = "발주"

def verify_insertion(start_row=19409, num_rows=5):
    if not os.path.exists(master_file_path):
        print(f"Error: Master file not found at {master_file_path}")
        return

    try:
        wb = openpyxl.load_workbook(master_file_path, data_only=False)
        ws = wb[sheet_name]
        
        print(f"Verifying Rows {start_row} to {start_row + num_rows - 1}...")
        
        with open(".tmp/insertion_verification.txt", "w", encoding="utf-8") as f:
            for row in range(start_row, start_row + num_rows):
                f.write(f"Row {row}:\n")
                for col in range(1, 15):
                    cell = ws.cell(row=row, column=col)
                    b_style = cell.border.bottom.style if cell.border and cell.border.bottom else "None"
                    f.write(f"  Col {col}: Val={cell.value}, Font={cell.font.name}, BorderBottom={b_style}\n")
                f.write("-" * 20 + "\n")
                
        print("✓ Verification written to .tmp/insertion_verification.txt")

    except Exception as e:
        print(f"Error verifying: {e}")

if __name__ == "__main__":
    verify_insertion(19409, 10) # Verify first 10 inserted rows
