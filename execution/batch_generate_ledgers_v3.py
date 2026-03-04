# -*- coding: utf-8 -*-
"""
Batch Transaction Ledger Generator v3
- Handles recursive template search for multi-branch stores
- Generates individual ledgers for each sub-branch
- Period: February 1-6, 2026
"""

import pandas as pd
import openpyxl
import os
import shutil
from datetime import datetime
import zipfile
import sys

# Set stdout to utf-8
if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
# Store mapping: Top folder name -> Master data branch keywords
# If the value is a string, it's a single branch.
# If it's a list, it contains multiple branch identifiers to look for.
# Store mapping: Top folder name -> Master data branch keywords
STORE_MAPPING = {
    "4. 샤브야키 광교점": "샤브야키 광교점",
    "4. 샤브야키 도봉점": "샤브야키 도봉점",
    "1. 샤브야키 분당점": "샤브야키 분당점",
    "2. 샤브야키 평택점": "샤브야키 평택점",
    "3. 샤브야키 주안점": "샤브야키 주안점",
    "4. 샤브야키 동탄점": "샤브야키 동탄점",
    "5. 부엉이 산장": ["부엉이산장 강남점", "부엉이산장 지오다노", "부엉이산장 구월점", "부엉이산장 마곡점"],
    "6. 고른햇살": "고른햇살",
    "7. 오레노이키루미치": ["오레노이키루미치 압구정점", "오레노이키루미치 하남점"],
    "8. 성동 과일 클래스 소유": ["소유프루트", "소유 프루트", "소유"],
    "9. 가락 과일 클래스 봄날": "봄날",
    "11. 브럭시": ["브럭시 송파점", "브럭시"],
    "12. 선데이버거클럽": "선데이버거클럽",
    "13. 육회꽃필무렵 송파점": ["육회꽃필무렵 송파점", "육회꽃필무렵"],
    "14. 파라이 카이카이": ["파라이 송파점", "파라이 역삼점"],
    "0. 이너프유": "이너프유",
    "10. 다비다웨딩홀뷔페 연제점": "다비다",
}

# Configuration
# Default Configuration
MASTER_DATA_PATH = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx"
TRANSACTION_BASE_PATH = r"G:\내 드라이브\1. 도크_주문 명세서\1. 거래명세서"
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "outputs")
START_DATE = "2026-02-01"
END_DATE = "2026-02-06"

# Ensure output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

def find_templates_recursively(root_path):
    """Find all potential template files in a directory tree"""
    templates = []
    try:
        if not os.path.exists(root_path):
            return []
        for root, dirs, files in os.walk(root_path):
            for file in files:
                if "(거래원장 템플릿)" in file and file.endswith(".xlsx") and "~$" not in file:
                    templates.append(os.path.join(root, file))
    except Exception as e:
        print(f"    Error walking directory {root_path}: {e}")
    return templates

def match_branch_name(template_path, branch_queries, master_names):
    """
    Match a template path to one of the branch queries and find the actual master data name.
    """
    def normalize_name(name):
        return name.replace(" ", "").replace("_", "").replace(".", "").lower()
        
    path_norm = normalize_name(template_path)
    
    if isinstance(branch_queries, str):
        branch_queries = [branch_queries]
        
    for query in branch_queries:
        query_norm = normalize_name(query)
        # Check if the query is in the path or filename
        if query_norm in path_norm:
            # Found a match in queries. Now find the exact name in master_names.
            # We use normalized comparison to master names as well.
            for m_name in master_names:
                if normalize_name(m_name) == query_norm:
                    return m_name
    
    # Fallback: if query is not in path but there is only one query, just try to match it
    if len(branch_queries) == 1:
        query = branch_queries[0]
        query_norm = normalize_name(query)
        for m_name in master_names:
            if normalize_name(m_name) == query_norm:
                return m_name
                
    return None

def process_single_ledger(df_master, template_file, branch_name):
    """Process a single ledger for a specific branch and template"""
    print(f"\n    Processing Branch: {branch_name}")
    print(f"    Template: {os.path.basename(template_file)}")
    
    # Filter by Store: always use normalized matching to handle name variants
    # (e.g., "오레노이키루미치_하남점" vs "오레노이키루미치 하남점")
    norm_branch = branch_name.replace(" ", "").replace("_", "")
    mask = df_master['매장명'].apply(
        lambda x: norm_branch in str(x).replace(" ", "").replace("_", "") or
                  str(x).replace(" ", "").replace("_", "") in norm_branch if pd.notna(x) else False
    )
    df_store = df_master[mask].copy()
    if df_store.empty:
        print(f"    ⚠️ No data found for branch '{branch_name}' in master data")
        return None
    matched_names = df_store['매장명'].unique().tolist()
    if len(matched_names) > 1 or matched_names[0] != branch_name:
        print(f"    [Name Match] '{branch_name}' → {matched_names}")
        
    df_store['기준일'] = pd.to_datetime(df_store['기준일'])
    mask = (df_store['기준일'] >= pd.to_datetime(START_DATE)) & (df_store['기준일'] <= pd.to_datetime(END_DATE))
    df_filtered = df_store.loc[mask]
    
    if df_filtered.empty:
        print(f"    ⚠️ No data found for '{branch_name}' in period {START_DATE}~{END_DATE}")
        return None
    
    daily_sales = df_filtered.groupby('기준일')['총 판매금액'].sum().reset_index()
    daily_sales = daily_sales.sort_values('기준일')
    
    # Create output folder (relative to template)
    start_dt = pd.to_datetime(START_DATE)
    end_dt = pd.to_datetime(END_DATE)
    template_dir = os.path.dirname(template_file)
    folder_name = f"{start_dt.year}년 {start_dt.month}월"
    output_folder = os.path.join(template_dir, folder_name)
    os.makedirs(output_folder, exist_ok=True)
    
    # File name
    # Extract store display name from folder or template name
    folder_basename = os.path.basename(template_dir)
    if ". " in folder_basename:
        display_name = folder_basename.split(". ", 1)[1]
    else:
        display_name = folder_basename
        
    # If display_name is too generic (like a single digit), use branch name
    if len(display_name) <= 2:
        display_name = branch_name
        
    file_name = f"{start_dt.year}{start_dt.month:02d}_{display_name}_월말정산_거래원장.xlsx"
    output_path = os.path.join(output_folder, file_name)
    
    print(f"    ✅ Generating: {file_name}")
    
    # Copy template
    shutil.copy(template_file, output_path)
    
    # Update Excel
    wb = openpyxl.load_workbook(output_path)
    ws = wb.active
    
    # Clear data (rows 17-43)
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
    title_text = f"{display_name}\n{month_str} 월말 정산 내역"
    ws.cell(row=5, column=2).value = title_text
    
    # Update date range (D7)
    date_range_text = f"{start_dt.year}년 {start_dt.month:02d}월 {start_dt.day:02d}일 ~ {end_dt.month:02d}월 {end_dt.day:02d}일"
    ws.cell(row=7, column=4).value = date_range_text
    
    # Update bottom summary (F47, F49)
    ws.cell(row=47, column=6).value = total_sum
    ws.cell(row=47, column=6).number_format = '#,##0'
    ws.cell(row=49, column=6).value = total_sum
    ws.cell(row=49, column=6).number_format = '#,##0'
    
    wb.save(output_path)
    print(f"    💰 Total: {total_sum:,}원")
    
    return output_path

def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--start", help="Start Date (YYYY-MM-DD)", default="2026-02-01")
    parser.add_argument("--end", help="End Date (YYYY-MM-DD)", default="2026-02-09")
    parser.add_argument("--master", help="Master Data Path", default=None)
    parser.add_argument("--out", help="Output Filename (in data/outputs)", default=None)
    args = parser.parse_args()

    global START_DATE, END_DATE, MASTER_DATA_PATH
    START_DATE = args.start
    END_DATE = args.end
    if args.master:
        MASTER_DATA_PATH = args.master

    print("="*80)
    print("Batch Transaction Ledger Generator v3")
    print(f"Period: {START_DATE} ~ {END_DATE}")
    print(f"Master: {MASTER_DATA_PATH}")
    print("="*80)
    
    # 1. Load Master Data early
    print("Loading Master Data...")
    try:
        df_master = pd.read_excel(MASTER_DATA_PATH, sheet_name='발주')
        master_names = df_master['매장명'].dropna().unique().tolist()
        print(f"Loaded master data with {len(master_names)} unique stores.")
    except Exception as e:
        print(f"❌ Error loading master data: {e}")
        return

    generated_files = []
    
    for top_folder, branch_queries in STORE_MAPPING.items():
        print(f"\nProcessing Folder: {top_folder}")
        root_path = os.path.join(TRANSACTION_BASE_PATH, top_folder)
        
        # 1. Recursive search for templates
        templates = find_templates_recursively(root_path)
        if not templates:
            print(f"  ❌ No templates found in {top_folder}")
            continue
            
        print(f"  ✓ Found {len(templates)} templates recursively.")
        
        # 2. Process each template
        for tpl in templates:
            branch_match = match_branch_name(tpl, branch_queries, master_names)
            if branch_match:
                try:
                    out = process_single_ledger(df_master, tpl, branch_match)
                    if out:
                        generated_files.append(out)
                except Exception as e:
                    print(f"    ❌ Error processing {branch_match}: {e}")
            else:
                print(f"    ⚠️ Could not match template to any known branch: {tpl}")

    # 3. Create zip file
    if generated_files:
        zip_filename = args.out if args.out else f"거래원장_2026년2월_{START_DATE[-2:]}-{END_DATE[-2:]}일.zip"
        zip_path = os.path.join(OUTPUT_DIR, zip_filename)
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in generated_files:
                arcname = os.path.basename(file_path)
                # To distinguish between same-named files from different folders in zip:
                # Use a slightly more descriptive arcname if needed, but let's stick to basename first.
                zipf.write(file_path, arcname)
        
        print("\n" + "="*80)
        print(f"✅ Zip file created: {zip_path}")
        print(f"   Total ledgers generated: {len(generated_files)}")
        print("="*80)
    else:
        print("\n❌ No ledgers were generated.")

if __name__ == "__main__":
    main()
