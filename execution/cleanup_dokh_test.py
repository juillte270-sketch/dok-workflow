import os
import glob

# Configuration
DRIVE_ROOT = r"G:\내 드라이브\1. 도크_주문 명세서\1. 거래명세서"
TARGET_DATE_STR = "20260131"

def cleanup_test_files():
    print(f"Scanning {DRIVE_ROOT} for files containing '{TARGET_DATE_STR}'...")
    
    deleted_count = 0
    
    for root, dirs, files in os.walk(DRIVE_ROOT):
        for f in files:
            # Match the file pattern generated: "20260131_StoreName_거래명세서.xlsx"
            if TARGET_DATE_STR in f and "거래명세서" in f and f.endswith(".xlsx"):
                full_path = os.path.join(root, f)
                try:
                    os.remove(full_path)
                    print(f"  [Deleted] {full_path}")
                    deleted_count += 1
                except Exception as e:
                    print(f"  [Error] Could not delete {full_path}: {e}")
                    
    print(f"\nCleanup complete. Total files deleted: {deleted_count}")

if __name__ == "__main__":
    cleanup_test_files()
