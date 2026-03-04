import os
import glob
import openpyxl
from openpyxl import load_workbook

# Configuration
DRIVE_ROOT = r"G:\내 드라이브\1. 도크_주문 명세서"
TARGET_FILENAME_KEYWORD = "AI참고용"
NEW_PHONE_NUMBER = "010-3275-0188"
TARGET_CELL = "K10" # Row 10, Column 11

def batch_update_templates():
    print(f"Scanning {DRIVE_ROOT} for templates...")
    
    updated_count = 0
    error_count = 0
    
    for root, dirs, files in os.walk(DRIVE_ROOT):
        for f in files:
            # Check if file matches criteria
            if TARGET_FILENAME_KEYWORD in f and f.endswith(".xlsx") and not f.startswith("~$"):
                full_path = os.path.join(root, f)
                print(f"Found template: {full_path}")
                
                try:
                    wb = load_workbook(full_path)
                    
                    # Assume first sheet is the target
                    sheet = wb.active
                    
                    # Update Value
                    # Row 10, Column 11
                    current_val = sheet.cell(row=10, column=11).value
                    
                    if str(current_val) == NEW_PHONE_NUMBER:
                        print(f"  -> Already up to date.")
                    else:
                        print(f"  -> Updating phone number: {current_val} -> {NEW_PHONE_NUMBER}")
                        sheet.cell(row=10, column=11).value = NEW_PHONE_NUMBER
                        wb.save(full_path)
                        print(f"  -> Saved.")
                        updated_count += 1
                        
                except Exception as e:
                    print(f"  -> [Error] Failed to update: {e}")
                    error_count += 1

    print("-" * 50)
    print(f"Summary: Updated {updated_count} files. Errors: {error_count}")

if __name__ == "__main__":
    batch_update_templates()
