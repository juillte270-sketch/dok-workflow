# -*- coding: utf-8 -*-
"""
Batch Transaction Ledger Generator - Based on working single-store script
Generates transaction ledgers for multiple stores for February 1-6, 2026
"""

import pandas as pd
import openpyxl
import os
import shutil
from datetime import datetime
import zipfile
import sys

# Store mapping based on diagnosis (Final Verification)
STORE_MAPPING = {
    # 샤브야키: Uses SPACES
    "0. 샤브야키 광교점": "샤브야키 광교점",  
    "1. 샤브야키 분당점": "샤브야키 분당점",
    "2. 샤브야키 평택점": "샤브야키 평택점",
    "3. 샤브야키 주안점": "샤브야키 주안점",
    "4. 샤브야키 동탄점": "샤브야키 동탄점",
    
    # 부엉이: Uses SPACES. "지오다노" likely refers to Gangnam branch.
    "5. 부엉이 산장": ["부엉이산장 지오다노", "부엉이산장 구월점", "부엉이산장 마곡점"],
    
    # 고른햇살: Exact match
    "6. 고른햇살": "고른햇살",
    
    # 오레노: Found versions with and without underscores. Using fallbacks if list logic supports logic update.
    # Note: My fallback logic in process_store handles variations of EACH name in the list.
    # So "오레노이키루미치_압구정점" will also try "오레노이키루미치 압구정점".
    "7. 오레노이키루미치": ["오레노이키루미치_압구정점", "오레노이키루미치_하남점"],
    
    # 소유: Try multiple variations
    "8. 성동 과일 클래스 소유": ["소유프루트", "소유 프루트", "소유"],
    
    # 봄날: Exact match
    "9. 가락 과일 클래스 봄날": "봄날",
    
    # 브럭시: "브럭시 송파점"
    "11. 브럭시": ["브럭시 송파점", "브럭시"],
    
    # 선데이버거클럽: Exact match ("선데이버거클럽 신사동점" seen too, but "선데이버거클럽" worked)
    "12. 선데이버거클럽": "선데이버거클럽",
    
    # 육회: "육회꽃필무렵 송파점"
    "13. 육회꽃필무렵 송파점": ["육회꽃필무렵 송파점", "육회꽃필무렵"],
    
    # 파라이: Space pattern confirmed
    "14. 파라이 카이카이": ["파라이 송파점", "파라이 역삼점"],
    
    # 이너프유: Exact match
    "0. 이너프유": "이너프유",
}

# Configuration
MASTER_DATA_PATH = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx"
TRANSACTION_BASE_PATH = r"G:\내 드라이브\1. 도크_주문 명세서\1. 거래명세서"
OUTPUT_DIR = r"c:\Users\DoKH_D\OneDrive\Desktop\안티그래비티 프로젝트 (3)\data\outputs"
START_DATE = "2026-02-01"
END_DATE = "2026-02-06"

# Ensure output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

def find_template_file(store_folder_path):
    """Find the template file for a given store folder"""
    try:
        for file in os.listdir(store_folder_path):
            if "(거래원장 템플릿)" in file and file.endswith(".xlsx"):
                return os.path.join(store_folder_path, file)
    except Exception as e:
        print(f"    Error listing directory: {e}")
    return None


def process_store(store_folder, store_names):
    """Process a single store (or multiple locations for one store)"""
    print(f"\n{'='*70}")
    print(f"Processing: {store_folder}")
    print(f"{'='*70}")
    
    store_folder_path = os.path.join(TRANSACTION_BASE_PATH, store_folder)
    
    # Check if folder exists
    if not os.path.exists(store_folder_path):
        print(f"  ❌ Folder not found: {store_folder_path}")
        return None
    
    # Find template
    template_file = find_template_file(store_folder_path)
    if not template_file:
        print(f"  ❌ Template not found")
        return None
    
    print(f"  ✓ Template: {os.path.basename(template_file)}")
    
    # Load master data
    try:
        df_master = pd.read_excel(MASTER_DATA_PATH, sheet_name='발주')
    except Exception as e:
        print(f"  ❌ Error loading master data: {e}")
        return None
    
    # Handle multiple store names (list of names or list of lists for fallbacks)
    # Strategy: 
    # If store_names is a list of strings, it could mean:
    # 1. Multiple branches (e.g. Owl, Oreno) -> Process all and sum them up?
    # 2. Or Fallback names? (e.g. Try this OR that)
    #
    # Let's standardize: 
    # STORE_MAPPING values will be a LIST of branch targets.
    # Each branch target can be a STRING (exact name) or a TUPLE (fallback options).
    
    # But wait, existing logic summed up all entries in the list.
    # Identifying if it's a fallback or aggregation is tricky with just a list.
    # Let's assume the user provided list in STORE_MAPPING contains ALL branches that belong to this folder.
    # For each branch in the list, we might need fallback names (e.g. "Branch A" or "Branch_A").
    
    # REVISED STRATEGY: 
    # The input `store_names` should be a list of "Branch Identifiers".
    # For each "Branch Identifier", we try to find it in the master data.
    # If the identifier matches exactly, good.
    # If not, we try variations (replace space with underscore, etc).
    # We sum up data for ALL found branches.
    
    if isinstance(store_names, str):
        store_names = [store_names]
        
    df_filtered_all = pd.DataFrame()
    found_any = False
    
    for name_query in store_names:
        # Try exact match first
        matched_names = []
        if name_query in df_master['매장명'].values:
            matched_names.append(name_query)
        else:
            # Try variations
            variations = [
                name_query,
                name_query.replace(" ", "_"),
                name_query.replace("_", " "),
                name_query.replace(" ", ""),
            ]
            for var in variations:
                if var in df_master['매장명'].values:
                    matched_names.append(var)
                    break # Found a match for this query
        
        if matched_names:
            print(f"  ✓ Found master data for: {matched_names}")
            for match in matched_names:
                df_store = df_master[df_master['매장명'] == match].copy()
                df_store['기준일'] = pd.to_datetime(df_store['기준일'])
                mask = (df_store['기준일'] >= pd.to_datetime(START_DATE)) & \
                       (df_store['기준일'] <= pd.to_datetime(END_DATE))
                df_filtered = df_store.loc[mask]
                df_filtered_all = pd.concat([df_filtered_all, df_filtered])
                found_any = True
        else:
             print(f"  ⚠️  No master data found for: '{name_query}' (or variations)")

    if not found_any or df_filtered_all.empty:
        print(f"  ⚠️  No data found for any branches in period")
        return None

    daily_sales = df_filtered_all.groupby('기준일')['총 판매금액'].sum().reset_index()
    daily_sales = daily_sales.sort_values('기준일')
    
    print(f"  ✓ Total Data rows: {len(daily_sales)}")
    
    # Create output folder
    start_dt = pd.to_datetime(START_DATE)
    folder_name = f"{start_dt.year}년 {start_dt.month}월"
    output_folder = os.path.join(store_folder_path, folder_name)
    os.makedirs(output_folder, exist_ok=True)
    
    # Create output file
    store_display_name = store_folder.split('. ', 1)[1] if '. ' in store_folder else store_folder
    file_name = f"{start_dt.year}{start_dt.month:02d}_{store_display_name}_월말정산_거래원장.xlsx"
    output_path = os.path.join(output_folder, file_name)
    
    # Copy template
    shutil.copy(template_file, output_path)
    
    # Update ledger (using same logic as single-store script)
    wb = openpyxl.load_workbook(output_path)
    ws = wb.active
    
    # Clear existing data
    for r in range(17, 44):
        ws.cell(row=r, column=2).value = None
        ws.cell(row=r, column=13).value = None
    
    # Write daily data
    start_row = 17
    total_sum = 0
    for idx, row in daily_sales.iterrows():
        current_row = start_row + idx
        ws.cell(row=current_row, column=2).value = row['기준일']
        ws.cell(row=current_row, column=2).number_format = 'yyyy-mm-dd'
        ws.cell(row=current_row, column=13).value = row['총 판매금액']
        ws.cell(row=current_row, column=13).number_format = '#,##0'
        total_sum += row['총 판매금액']
    
    # Update title (B5)
    month_str = f"{start_dt.month:02d}월"
    title_text = f"{store_display_name}\\n{month_str} 월말 정산 내역"
    ws.cell(row=5, column=2).value = title_text
    
    # Update date range (D7)
    end_dt = pd.to_datetime(END_DATE)
    date_range_text = f"{start_dt.year}년 {start_dt.month:02d}월 {start_dt.day:02d}일 ~ {end_dt.month:02d}월 {end_dt.day:02d}일"
    ws.cell(row=7, column=4).value = date_range_text
    
    # Update bottom summary fields
    ws.cell(row=47, column=6).value = total_sum
    ws.cell(row=47, column=6).number_format = '#,##0'
    ws.cell(row=49, column=6).value = total_sum
    ws.cell(row=49, column=6).number_format = '#,##0'
    
    wb.save(output_path)
    print(f"  ✅ Saved: {file_name}")
    print(f"  Total: {total_sum:,}원")
    
    return output_path


def main():
    """Main execution"""
    print("="*70)
    print("Batch Transaction Ledger Generator")
    print(f"Period: {START_DATE} ~ {END_DATE}")
    print("="*70)
    
    generated_files = []
    success_count = 0
    
    for store_folder, store_names in STORE_MAPPING.items():
        try:
            output_path = process_store(store_folder, store_names)
            if output_path:
                generated_files.append(output_path)
                success_count += 1
        except Exception as e:
            print(f"  ❌ Error: {e}")
            import traceback
            traceback.print_exc()
    
    # Create zip file
    if generated_files:
        zip_path = os.path.join(OUTPUT_DIR, "거래원장_2026년2월_1-6일.zip")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in generated_files:
                arcname = os.path.basename(file_path)
                zipf.write(file_path, arcname)
        
        print(f"\n{'='*70}")
        print(f"✅ Zip file created: {zip_path}")
        print(f"   Total files: {len(generated_files)}")
        print(f"{'='*70}")
    
    print(f"\n✅ Batch processing complete!")
    print(f"   Successfully processed: {success_count}/{len(STORE_MAPPING)} stores")


if __name__ == "__main__":
    main()
