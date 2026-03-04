import argparse
import openpyxl
import os
import time

# Import Scraper
try:
    from fetch_foodspring_prices import FoodspringScraper
except ImportError:
    # Handle running from root vs execution dir
    import sys
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    from fetch_foodspring_prices import FoodspringScraper

def fetch_market_price(scraper, item_name):
    """
    Search for item on Foodspring using the provided scraper instance.
    """
    price, supplier = scraper.search_item(item_name)
    return price, supplier

def update_prices(master_file, target_date, row_start=None, row_end=None):
    print(f"Opening master file: {master_file}")
    try:
        wb = openpyxl.load_workbook(master_file)
    except Exception as e:
        print(f"Error opening file: {e}")
        return

    ws = wb["단가"]

    updates_count = 0

    start_row = row_start if row_start else 2
    max_row = row_end if row_end else ws.max_row
    
    print(f"Scanning {max_row} rows for date {target_date}...")
    
    # Initialize Scraper
    scraper = FoodspringScraper()
    if not scraper.login():
        print("Failed to login to Foodspring. Aborting updates.")
        scraper.close()
        return

    try:
        # First pass: Collect all target rows
        target_rows = []
        for row_idx in range(start_row, max_row + 1):
            date_cell = ws.cell(row=row_idx, column=2)
            cell_date = date_cell.value
            is_match = False
            if str(cell_date).startswith(target_date):
                 is_match = True
            elif hasattr(cell_date, 'strftime'):
                if cell_date.strftime("%Y-%m-%d") == target_date:
                    is_match = True
            
            if is_match and ws.cell(row=row_idx, column=3).value:
                target_rows.append(row_idx)

        total_files = len(target_rows)
        print(f"Found {total_files} items to update for date {target_date}.")
        
        current_idx = 0
        for row_idx in target_rows:
            current_idx += 1
            item_name = str(ws.cell(row=row_idx, column=3).value).strip()
            
            # Search Mappings - Use specific search terms
            SEARCH_MAPPING = {
                # 깻잎
                "깻잎(4KG,큰잎)": "찹큰",
                "깻잎(4kg,큰잎)": "찹큰",
                "깻잎 4kg": "찹큰",
                "깻잎(4K,큰잎)": "찹큰",
                "깻잎(1KG)": "깻잎",
                
                # 쑥갓 - 4KG
                "쑥갓": "쑥갓 4KG",
                
                # 알배기
                "알배기배추": "알배기",
                "알배기": "알배기",
                
                # 표고버섯
                "생표고 수입": "표고버섯 2L",
                "생표고버섯 수입": "표고버섯 2L",
                
                # 로메인
                "로메인(일반)": "청로메인",
                "통로메인": "통로메인",
                
                # 양상추
                "양상추(일반)": "양상추",
                "양상추(국산/수입)": "양상추",
                
                # 마늘
                "깐마늘 대서(상OR대, 1K)": "깐마늘 대 1KG",
                "깐마늘 대서(보통OR중,1K)": "깐마늘 중 1KG",
                "깐마늘 대서(보통OR중, 20K)": "깐마늘 중 20KG",

                # 쪽파/부추
                "깐쪽파": "쪽파",
                "부추(일반)": "부추",

                # 추가 매핑 (2026-02-23 피드백)
                "케일": "케일",
                "깐대파": "깐대파 국내산",
                "미나리": "미나리 1단",
                "배추": "배추52",
                "레몬": "레몬 1박스",
                "간마늘": "다진마늘 1kg",
            }
            
            search_keyword = SEARCH_MAPPING.get(item_name, item_name)
            print(f"[{current_idx}/{total_files}] Processing: {item_name} (Search: {search_keyword})")
                
            # Fetch Price
            price, supplier = fetch_market_price(scraper, search_keyword)
            
            if price:
                # Update Price (Column 8 / H)
                price_cell = ws.cell(row=row_idx, column=8) # H
                price_cell.value = price
                
                # Update Supplier (Column 13 / M)
                supplier_cell = ws.cell(row=row_idx, column=13) # M
                supplier_cell.value = supplier
                
                print(f"  -> [Updated] {price}원 / {supplier}")
                updates_count += 1
            else:
                print(f"  -> [Skip] No price found")
                    
    finally:
        scraper.close()
                
    if updates_count > 0:
        print(f"Saving {updates_count} updates to {master_file}...")
        wb.save(master_file)
        print("Done.")
    else:
        print("No updates made (Scraper not active or no matches).")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--master", required=True)
    parser.add_argument("--date", required=True)
    parser.add_argument("--start-row", type=int, default=None, help="Start row (inclusive)")
    parser.add_argument("--end-row", type=int, default=None, help="End row (inclusive)")

    args = parser.parse_args()

    update_prices(args.master, args.date, args.start_row, args.end_row)
