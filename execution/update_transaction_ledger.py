import pandas as pd
import openpyxl
import os
import argparse
from datetime import datetime
import shutil

def update_transaction_ledger(master_file, template_file, target_dir, store_name, start_date, end_date):
    """
    Update transaction ledger for a specific store and date range.
    """
    print(f"Processing Store: {store_name}")
    print(f"Date Range: {start_date} ~ {end_date}")
    
    # 1. Load Master Data
    print(f"Loading Master Data: {master_file}")
    try:
        df_master = pd.read_excel(master_file, sheet_name='발주')
    except Exception as e:
        print(f"Error loading master file: {e}")
        return

    # 2. Filter Master Data
    # Ensure correct column names from inspection: 
    # 0: 매장명, 3: 기준일, 13: 총 판매금액 (Assuming based on inspection, needs verification if column names are used)
    # Inspection showed headers on row 0.
    
    # Filter by Store
    df_store = df_master[df_master['매장명'] == store_name].copy()
    
    # Filter by Date
    # Convert '기준일' to datetime
    df_store['기준일'] = pd.to_datetime(df_store['기준일'])
    
    mask = (df_store['기준일'] >= pd.to_datetime(start_date)) & (df_store['기준일'] <= pd.to_datetime(end_date))
    df_filtered = df_store.loc[mask]
    
    if df_filtered.empty:
        print("No records found for the specified date range and store.")
        return

    # 3. Aggregate Daily Sales
    # Group by Date and sum '총 판매금액'
    daily_sales = df_filtered.groupby('기준일')['총 판매금액'].sum().reset_index()
    daily_sales = daily_sales.sort_values('기준일')
    
    print("\nDaily Sales Summary:")
    print(daily_sales)
    
    # 4. Prepare Target File
    # Create Target Directory
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)
        print(f"Created directory: {target_dir}")
        
    start_dt = pd.to_datetime(start_date)
    file_name = f"{start_dt.strftime('%Y%m')}_{store_name}_월말정산_거래원장.xlsx"
    target_path = os.path.join(target_dir, file_name)
    
    # Copy Template
    print(f"Creating Ledger File: {target_path}")
    shutil.copy(template_file, target_path)
    
    # 5. Write to Ledger
    try:
        wb = openpyxl.load_workbook(target_path)
        ws = wb.active
        
        # Identify Key Rows (Hardcoded based on inspection, but logic added for safety)
        header_row_idx = 16
        data_start_row = 17
        
        # Verify Header
        cell_b16 = ws.cell(row=16, column=2).value # B16
        if cell_b16 != "발주일":
             print(f"Warning: Expected '발주일' at B16, found '{cell_b16}'. Please check template.")
        
        # Find Bottom Total Row (Look for "합계금액" or "합계" around row 49)
        bottom_total_row = None
        for r in range(40, 60):
            val = ws.cell(row=r, column=2).value # Column B
            if val and ("합계" in str(val) or "합 계" in str(val)):
                bottom_total_row = r
                print(f"Found Bottom Total Row at: {r}")
                break
        
        if not bottom_total_row:
            print("Could not find bottom total row. Using default 49.")
            bottom_total_row = 49

        # Search for Top Total Row (around row 9)
        top_total_row = 9
        
        # Clear existing data
        print(f"Clearing data from row {data_start_row} to {bottom_total_row-1}")
        for r in range(data_start_row, bottom_total_row):
             cell_date = ws.cell(row=r, column=2)
             cell_price = ws.cell(row=r, column=13)
             # print(f"Clearing Row {r}: B={type(cell_date)}, M={type(cell_price)}")
             if isinstance(cell_date, openpyxl.cell.cell.MergedCell) or isinstance(cell_price, openpyxl.cell.cell.MergedCell):
                  print(f"Row {r} contains merged cells. Skipping clear for safety.")
                  continue
             cell_date.value = None
             cell_price.value = None

        # Write Data
        current_row = data_start_row
        total_sum = 0
        
        for index, row in daily_sales.iterrows():
            date_val = row['기준일']
            amount = row['총 판매금액']
            
            print(f"Writing Row {current_row}: Date={date_val}, Amount={amount}")
            # Write Date (B column) - Format MM월 DD일 usually, or YYYY-MM-DD
            # ws.cell(row=current_row, column=2).value = date_val.strftime("%Y-%m-%d")
            # User example: 2월1일 ~ 2월6일. Usually Excel date format.
            ws.cell(row=current_row, column=2).value = date_val
            ws.cell(row=current_row, column=2).number_format = 'mm"월" dd"일"'
            
            # Write Amount (M column) - 공급가액
            ws.cell(row=current_row, column=13).value = amount
            ws.cell(row=current_row, column=13).number_format = '#,##0'
            
            total_sum += amount
            current_row += 1
            
        print(f"Written {len(daily_sales)} records. Total Sum: {total_sum}")
        
        # Write Totals
        # 3. 하단 합계에는 거래한 날짜의 공급가액을 다 합쳐서 입력
        # Bottom Total Row, Column M (13)
        print(f"Writing Bottom Total at Row {bottom_total_row}, Col 13")
        ws.cell(row=bottom_total_row, column=13).value = total_sum
        ws.cell(row=bottom_total_row, column=13).number_format = '#,##0'
        
        # Update Title in B5 (change month from 01월 to correct month)
        print("Updating Title in B5...")
        start_dt = pd.to_datetime(start_date)
        end_dt = pd.to_datetime(end_date)
        month_str = f"{start_dt.month:02d}월"
        title_text = f"샤브야키_도봉점\n{month_str} 월말 정산 내역"
        ws.cell(row=5, column=2).value = title_text  # B5
        print(f"  Title updated to: {title_text.replace(chr(10), ' ')}")
        
        # Write Date Range to D7
        print("Writing Date Range to D7...")
        date_range_text = f"{start_dt.year}년 {start_dt.month:02d}월 {start_dt.day:02d}일 ~ {end_dt.month:02d}월 {end_dt.day:02d}일"
        ws.cell(row=7, column=4).value = date_range_text  # D7
        print(f"  Date range written to D7: {date_range_text}")
        
        # Write Bottom Summary Fields
        print("Writing Bottom Summary Fields...")
        # F47: 당월금액
        ws.cell(row=47, column=6).value = total_sum
        ws.cell(row=47, column=6).number_format = '#,##0'
        print(f"  당월금액 (F47): {total_sum:,}")
        
        # F49: 합계금액
        ws.cell(row=49, column=6).value = total_sum
        ws.cell(row=49, column=6).number_format = '#,##0'
        print(f"  합계금액 (F49): {total_sum:,}")
        
        # D9 and F9 already have formulas "=M44" in template, so they will auto-calculate
        # No need to write values, formulas will reference the bottom total
        print("D9 and F9 formulas (=M44) will auto-calculate from bottom total")

        wb.save(target_path)
        print(f"Successfully saved ledger to: {target_path}")

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error processing ledger file: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--date_range", help="Date range (YYYY-MM-DD~YYYY-MM-DD)", default="2026-02-01~2026-02-06")
    parser.add_argument("--store", help="Store Name", default="샤브야키 도봉점")
    args = parser.parse_args()
    
    start_date, end_date = args.date_range.split("~")
    
    master_file = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx"
    template_file = r"G:\내 드라이브\1. 도크_주문 명세서\1. 거래명세서\0. 샤브야키 도봉점\(거래원장 템플릿)202501_샤브야키 도봉점_월말정산_거래원장.xlsx"
    
    # Extract year/month from start_date for folder name
    dt = datetime.strptime(start_date, "%Y-%m-%d")
    folder_name = f"{dt.year}년 {dt.month}월"
    
    target_dir = os.path.join(r"G:\내 드라이브\1. 도크_주문 명세서\1. 거래명세서\0. 샤브야키 도봉점", folder_name)
    
    update_transaction_ledger(master_file, template_file, target_dir, args.store, start_date, end_date)
