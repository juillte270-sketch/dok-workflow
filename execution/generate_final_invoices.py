"""
최종 거래명세서 생성 스크립트 (Fuzzy Matching 추가)

1단계에서 생성된 가명세서(수량만 있는 상태)를 읽어서:
1. 마스터 데이터(도크발주관리데이터.xlsx)에서 판매가를 조회하여 입력
2. 당월 누적 금액, 당일 금액, 총 사용액 자동 계산 및 입력
3. 최종 결과물을 Excel 및 PDF로 저장
4. Google Drive로 자동 전송 (스마트 폴더 매칭)

사용법:
    python execution/generate_final_invoices.py --date "2026-02-05"
"""

import os
import glob
import openpyxl
from openpyxl.styles import Font
from datetime import datetime, timedelta
import argparse
import re
import shutil
import unicodedata

# 프로젝트 루트
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 마스터 파일 경로
MASTER_FILE = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx"
MASTER_SHEET = "발주"

def normalize_key(text):
    """매칭 키 정규화 (공백/언더스코어 제거 + NFC 정규화)"""
    if not text: return ""
    text_str = str(text)
    # NFC 정규화 (자소 분리 방지)
    text_nfc = unicodedata.normalize('NFC', text_str)
    return text_nfc.replace(" ", "").replace("_", "").strip()

def load_master_prices(target_date):
    """마스터 파일에서 판매가 정보 로드"""
    print(f"\n📊 마스터 파일에서 판매가 로드 중...")
    
    wb = openpyxl.load_workbook(MASTER_FILE, data_only=True)
    ws = wb[MASTER_SHEET]
    
    prices = {}
    print(f"   총 행 수: {ws.max_row:,}개 (데이터 추출 중...)")
    
    for row in range(2, ws.max_row + 1):
        if row % 10000 == 0:
            print(f"   처리 중: {row:,}/{ws.max_row:,} 행...")
        
        매장명 = ws.cell(row, 1).value
        날짜 = ws.cell(row, 4).value
        품목 = ws.cell(row, 5).value
        단위 = ws.cell(row, 7).value or ''
        판매가 = ws.cell(row, 11).value
        총판매금액 = ws.cell(row, 12).value

        if not all([매장명, 날짜, 품목]):
            continue

        날짜_str = ""
        if isinstance(날짜, datetime):
            날짜_str = 날짜.strftime("%Y-%m-%d")
        elif isinstance(날짜, str):
            try:
                날짜_str = datetime.strptime(날짜[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
            except:
                continue
        else:
            continue

        # 키 정규화 (공백 제거) — 단위 포함으로 동일 품목 다른 단위 구분
        매장_key = normalize_key(매장명)
        품목_key = normalize_key(품목)
        단위_key = normalize_key(단위)
        key = f"{날짜_str}|{매장_key}|{품목_key}|{단위_key}"

        def safe_float(value):
            if value is None or value == '' or value == '-': return 0
            try: return float(value)
            except: return 0

        # 동일 키 충돌 시 총판매금액 누적 (같은 품목+단위가 여러 행일 수 있음)
        if key in prices:
            prices[key]['총판매금액'] += safe_float(총판매금액)
        else:
            prices[key] = {
                '판매가': safe_float(판매가),
                '총판매금액': safe_float(총판매금액)
            }
    
    wb.close()
    
    # 디버깅: 해당 날짜 데이터 개수 확인
    target_str = target_date.strftime("%Y-%m-%d")
    count_target = sum(1 for k in prices if k.startswith(target_str))
    print(f"✅ 판매가 정보 로드 완료: 총 {len(prices):,}개 중 [{target_str}] 데이터 {count_target:,}개 포함됨.")
    
    return prices

def find_cell_by_keyword(sheet, keyword, search_range=None):
    if search_range is None:
        min_row, max_row = 1, 50
        min_col, max_col = 1, 10
    else:
        min_row, max_row, min_col, max_col = search_range
        
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            val = sheet.cell(row, col).value
            if val and keyword in str(val):
                return row, col
    return None, None

def find_value_cell_in_row(sheet, row, start_col=2):
    """해당 행에서 값을 넣을만한 셀(빈칸, -, 0, #)을 찾아서 좌표 반환"""
    for col in range(start_col + 1, 15): 
        cell = sheet.cell(row, col)
        
        is_merged = isinstance(cell, openpyxl.cell.cell.MergedCell)
        top_left = (row, col)
        if is_merged:
            for mr in sheet.merged_cells.ranges:
                if cell.coordinate in mr:
                    top_left = (mr.min_row, mr.min_col)
                    break
        else:
             top_left = (row, col)
             
        real_cell = sheet.cell(top_left[0], top_left[1])
        val = str(real_cell.value).strip() if real_cell.value else ""
        
        if not val or val in ['-', '0'] or '#' in val or '원' in val:
            return top_left[0], top_left[1]
            
    # 못 찾으면 G열(7) 기본값
    g_cell = sheet.cell(row, 7)
    if isinstance(g_cell, openpyxl.cell.cell.MergedCell):
        for mr in sheet.merged_cells.ranges:
             if g_cell.coordinate in mr:
                 return mr.min_row, mr.min_col
    return row, 7

def set_value_safe(sheet, row, col, value, number_format='#,##0', bold=False):
    """병합된 셀이라도 안전하게 값 입력"""
    try:
        cell = sheet.cell(row, col)
        for mr in sheet.merged_cells.ranges:
            if cell.coordinate in mr:
                cell = sheet.cell(mr.min_row, mr.min_col)
                break
        
        cell.value = value
        cell.number_format = number_format
        if bold:
            cell.font = Font(bold=True)
        else:
            pass 
    except Exception as e:
        print(f"   ⚠️ 값 입력 오류 ({row},{col}): {e}")

class ValidationReport:
    def __init__(self, date_str):
        self.date_str = date_str
        self.records = []
        self.timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
    def add(self, store, is_success, invoice_daily_total=0, zero_price_count=0, msg=""):
        self.records.append({
            'store': store,
            'is_success': is_success,
            'invoice_total': invoice_daily_total,
            'zero_count': zero_price_count,
            'msg': msg
        })
        
    def save(self, project_root):
        report_filename = f"validation_report_{self.date_str}.txt"
        output_dir = os.path.join(project_root, "data", "outputs", f"final_invoices_{self.date_str}")
        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        filepath = os.path.join(output_dir, report_filename)
        
        success_count = sum(1 for r in self.records if r['is_success'])
        fail_count = len(self.records) - success_count
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"==================================================\n")
            f.write(f" [최종 거래명세서 검증 리포트]\n")
            f.write(f" 생성 일시: {self.timestamp}\n")
            f.write(f" 대상 날짜: {self.date_str}\n")
            f.write(f"==================================================\n\n")
            f.write(f"📊 요약:\n")
            f.write(f"   - 총 처리: {len(self.records)}건\n")
            f.write(f"   - 성공: {success_count}건\n")
            f.write(f"   - 실패: {fail_count}건\n\n")
            f.write(f"📋 상세 내역:\n")
            f.write(f"{'매장명':<20} | {'상태':<5} | {'당일합계':>12} | {'단가0원품목':>10} | {'비고'}\n")
            f.write(f"-"*80 + "\n")
            for r in self.records:
                status = "성공" if r['is_success'] else "실패"
                f.write(f"{r['store']:<20} | {status:<5} | {r['invoice_total']:>12,} | {r['zero_count']:>10} | {r['msg']}\n")
            f.write(f"\n==================================================\n")
        print(f"\n📝 검증 리포트 생성됨: {filepath}")

def find_drive_folder(base_path, store_name):
    """드라이브 내의 해당 매장 폴더 찾기 (지점별 하위폴더 포함)"""
    # 명시적 매핑 (자동 매칭이 안 되는 경우)
    EXPLICIT_MAP = {
        '파라이역삼점': r'14. 파라이 카이카이\파라이 역삼점',
        '파라이송파점': r'14. 파라이 카이카이\파라이 송파점',
        '소유프루트': '8. 성동 과일 클래스 소유',
        '소유': '8. 성동 과일 클래스 소유',
    }
    normalized_name = store_name.replace(' ', '').replace('_', '')
    if normalized_name in EXPLICIT_MAP:
        return EXPLICIT_MAP[normalized_name]
    try:
        normalized_name = store_name.replace(' ', '').replace('_', '')
        candidates = []
        if os.path.exists(base_path):
            for name in os.listdir(base_path):
                full_path = os.path.join(base_path, name)
                if not os.path.isdir(full_path): continue
                name_no_digit = re.sub(r'^\d+\.', '', name).strip()
                clean_name = name.replace(' ', '').replace('_', '').replace('.', '')
                clean_no_digit = name_no_digit.replace(' ', '').replace('_', '')
                if normalized_name in clean_name or clean_no_digit in normalized_name:
                    candidates.append((name, len(name)))
        if not candidates:
            return store_name.replace('_', ' ')

        # 상위 폴더 찾기 (가장 짧은 이름 = 상위 폴더)
        candidates.sort(key=lambda x: x[1])
        parent_folder = candidates[0][0]
        parent_path = os.path.join(base_path, parent_folder)

        # 하위에 지점별 폴더가 있는지 확인
        sub_candidates = []
        for sub_name in os.listdir(parent_path):
            sub_full = os.path.join(parent_path, sub_name)
            if not os.path.isdir(sub_full): continue
            # 월별 폴더(예: "2026년 2월", "26년1월", "25년")는 건너뜀
            if re.match(r'^\d{2,4}년', sub_name): continue
            if sub_name in ('정산서', '거래중지'): continue

            sub_no_digit = re.sub(r'^\d+\.', '', sub_name).strip()
            clean_sub = sub_name.replace(' ', '').replace('_', '').replace('.', '')
            clean_sub_no_digit = sub_no_digit.replace(' ', '').replace('_', '')
            if normalized_name in clean_sub or clean_sub_no_digit in normalized_name:
                sub_candidates.append((os.path.join(parent_folder, sub_name), len(clean_sub)))

        if sub_candidates:
            # 가장 정확한 매칭 (긴 이름이 더 구체적)
            sub_candidates.sort(key=lambda x: x[1], reverse=True)
            return sub_candidates[0][0]

        return parent_folder
    except:
        return store_name.replace('_', ' ')

def find_price_fuzzy(prices, target_date, store_key, item_name):
    """유사 매칭(포함 관계)으로 가격 찾기"""
    # 1. 1차 필터링: 해당 날짜+매장의 모든 품목 추출
    # prices 키 형식: 날짜|매장|품목
    # prefix = f"{target_date}|{store_key}|"
    prefix = f"{target_date.strftime('%Y-%m-%d')}|{store_key}|"
    
    candidates = {}
    for k, v in prices.items():
        if k.startswith(prefix):
            # key 형식: 날짜|매장|품목|단위 → 품목은 [2]
            parts = k.split('|')
            item_n = parts[2] if len(parts) >= 3 else parts[-1]
            if item_n not in candidates:
                candidates[item_n] = v
            else:
                # 동일 품목 다른 단위 → 총판매금액 누적
                candidates[item_n] = {
                    '판매가': v['판매가'],
                    '총판매금액': candidates[item_n]['총판매금액'] + v['총판매금액']
                }
            
    if not candidates:
        return None
        
    tgt = normalize_key(item_name)
    
    # 2. 포함 관계 확인
    for master_item, val in candidates.items():
        # master_item은 이미 normalize_key 된 상태로 prices에 저장됨
        if tgt in master_item or master_item in tgt:
            print(f"      [Fuzzy Match] '{item_name}' ~= '{master_item}'")
            return val
            
    return None

def add_prices_to_invoice(invoice_path, prices, target_date, store_name, validator):
    print(f"\n💰 {store_name} - 판매가 입력 처리")
    
    try:
        wb = openpyxl.load_workbook(invoice_path)
        sheet = wb.active
        
        단가_col = None
        공급가액_col = None
        header_row = 15
        
        r, c = find_cell_by_keyword(sheet, "단가", (10, 20, 1, 20))
        if r:
            header_row = r
            단가_col = c
        
        if header_row:
            for col in range(1, 20):
                val = sheet.cell(header_row, col).value
                if val and ('공급가액' in str(val) or '금액' in str(val)) and (not 단가_col or col != 단가_col):
                    공급가액_col = col
                    break
        
        if not 단가_col: 단가_col = 10
        if not 공급가액_col: 공급가액_col = 11  # K열
        
        current_row = header_row + 1
        daily_total = 0
        items_processed = 0
        zero_price_items = 0
        
        store_key = normalize_key(store_name)
        
        합계_row = None

        while current_row < 100:
            품목_cell = sheet.cell(current_row, 2)
            if not 품목_cell.value or str(품목_cell.value).strip() == '':
                is_empty = True
                for check in range(1, 4):
                    if sheet.cell(current_row + check, 2).value:
                        is_empty = False
                        break
                if is_empty: break
                current_row += 1
                continue

            품목 = str(품목_cell.value).strip()

            # 중간 합계 (합계/계) — 위치만 기록, 나중에 실제 합계로 입력
            if '합계' in 품목 or '계' == 품목:
                합계_row = current_row
                break

            # 가격 매칭 (1. 정확 품목+단위, 2. 품목만 prefix, 3. Fuzzy)
            품목_key = normalize_key(품목)
            단위_cell = sheet.cell(current_row, 4).value
            단위_key = normalize_key(단위_cell) if 단위_cell else ''
            date_prefix = target_date.strftime('%Y-%m-%d')
            key_with_unit = f"{date_prefix}|{store_key}|{품목_key}|{단위_key}"
            key_prefix = f"{date_prefix}|{store_key}|{품목_key}|"

            price_info = None

            if key_with_unit in prices:
                price_info = prices[key_with_unit]
            else:
                # 단위 무관 prefix 매칭 (같은 품목의 첫 번째 단위)
                for k, v in prices.items():
                    if k.startswith(key_prefix):
                        price_info = v
                        break
                if not price_info:
                    # Fuzzy Matching 시도
                    price_info = find_price_fuzzy(prices, target_date, store_key, 품목)

            if price_info:
                판매가 = price_info['판매가']
                총판매금액 = price_info['총판매금액']

                set_value_safe(sheet, current_row, 단가_col, 판매가)
                set_value_safe(sheet, current_row, 공급가액_col, 총판매금액)

                daily_total += 총판매금액
                items_processed += 1
                if 판매가 == 0: zero_price_items += 1
            else:
                if items_processed < 3:
                     print(f"   ⚠️ 매칭 실패: [{key}] (품목: {품목})")

            current_row += 1

        # 루프에서 합계 행을 못 찾은 경우, 빈 행 뒤에서 탐색 (숨김 행이 많을 수 있으므로 넓게)
        if not 합계_row:
            r_sum, c_sum = find_cell_by_keyword(sheet, "합계", (current_row, min(current_row + 60, sheet.max_row), 1, 5))
            if r_sum:
                합계_row = r_sum

        # 실제 공급가액 셀 값을 읽어서 합계 재계산
        end_row = 합계_row if 합계_row else current_row
        actual_daily_total = 0
        for r in range(header_row + 1, end_row):
            cell = sheet.cell(r, 공급가액_col)
            if isinstance(cell, openpyxl.cell.cell.MergedCell):
                for mr in sheet.merged_cells.ranges:
                    if cell.coordinate in mr:
                        cell = sheet.cell(mr.min_row, mr.min_col)
                        break
            val = cell.value
            if isinstance(val, (int, float)):
                actual_daily_total += val

        # 실제 셀 합계가 더 크면 (기존 값이 있는 경우) 그 값 사용
        if actual_daily_total > daily_total:
            daily_total = actual_daily_total

        # 가격 매칭이 하나도 안 된 경우 합계에 0 입력 방지
        has_prices = items_processed > 0 or daily_total > 0

        # 합계 행 입력
        if 합계_row and has_prices:
            set_value_safe(sheet, 합계_row, 공급가액_col, daily_total)
            print(f"   ✓ 공급가액 합계 입력: {daily_total:,} (Row {합계_row})")
            if 단가_col:
                try: set_value_safe(sheet, 합계_row, 단가_col, "")
                except: pass
        elif not has_prices:
            print(f"   ⚠️ 가격 매칭 없음 — 합계 미입력 (기존 값 유지)")

        store_name_key = store_name.replace('_', ' ')
        monthly_amount = calculate_monthly_amount(prices, store_name, target_date)
        total_amount = monthly_amount + daily_total

        # 합계 입력 (하단) — 가격 매칭이 있을 때만 입력
        if has_prices:
            # (1) 상단 합계
            r_top, c_top = find_cell_by_keyword(sheet, "합계금액", (1, 12, 1, 5))
            if r_top:
                 tr, tc = find_value_cell_in_row(sheet, r_top, c_top)
                 set_value_safe(sheet, tr, tc, daily_total)
            else:
                 set_value_safe(sheet, 8, 7, daily_total)

        # (2) 하단 금액 정보
        r_month, c_month = find_cell_by_keyword(sheet, "당월", (30, 60, 1, 5))
        if r_month:
            # 1. 당월금액 (당월 누적은 항상 입력)
            tr, tc = find_value_cell_in_row(sheet, r_month, c_month)
            set_value_safe(sheet, tr, tc, monthly_amount)
            col_target = tc

            # 2. 당일금액 — 가격 매칭 있을 때만
            t_row, t_col = find_value_cell_in_row(sheet, r_month + 1, c_month)
            if has_prices:
                set_value_safe(sheet, r_month + 1, col_target, daily_total)

            # 3. 총사용액
            set_value_safe(sheet, r_month + 2, col_target, total_amount)
            
            # 5. 잔액
            payment = 0
            try:
                pay_cell = sheet.cell(r_month + 3, col_target)
                if isinstance(pay_cell, openpyxl.cell.cell.MergedCell):
                     for mr in sheet.merged_cells.ranges:
                         if pay_cell.coordinate in mr:
                             pay_cell = sheet.cell(mr.min_row, mr.min_col)
                             break
                if pay_cell.value and isinstance(pay_cell.value, (int, float)):
                    payment = pay_cell.value
            except: pass
            
            final_val = total_amount - payment
            set_value_safe(sheet, r_month + 4, col_target, final_val, bold=False)
            
            print(f"   ✓ 하단 합계 입력 완료 (@ Row {r_month+4})")
            
        else:
            set_value_safe(sheet, 41, 7, monthly_amount)
            set_value_safe(sheet, 42, 7, daily_total)
            set_value_safe(sheet, 43, 7, total_amount)
            set_value_safe(sheet, 45, 7, total_amount, bold=False)

        output_filename = f"{target_date.strftime('%Y%m%d')}_{store_name}_거래명세서.xlsx"
        output_dir = os.path.join(PROJECT_ROOT, "data", "outputs", f"final_invoices_{target_date.strftime('%Y%m%d')}")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, output_filename)
        
        wb.save(output_path)
        
        msg = ""
        if zero_price_items > 0: msg = f"단가0원 {zero_price_items}건"
        validator.add(store_name, True, daily_total, zero_price_items, msg)
        return output_path
        
    except Exception as e:
        print(f"   ❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        validator.add(store_name, False, 0, 0, str(e))
        return None

def calculate_monthly_amount(prices, store_name, target_date):
    """월 누적 금액 계산"""
    if target_date.day == 1: return 0
    month_start = target_date.replace(day=1)
    day_before = target_date - timedelta(days=1)
    total = 0
    
    store_key = normalize_key(store_name)
    
    current_date = month_start
    while current_date <= day_before:
        date_str = current_date.strftime("%Y-%m-%d")
        prefix = f"{date_str}|{store_key}|"
        
        for key, val in prices.items():
            if key.startswith(prefix):
                total += val['총판매금액']
                
        current_date += timedelta(days=1)
    return total

def convert_to_pdf(excel_path):
    try:
        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        abs_path = os.path.abspath(excel_path)
        wb = excel.Workbooks.Open(abs_path)
        pdf_path = abs_path.replace('.xlsx', '.pdf')
        ws = wb.Worksheets(1)
        # Clear stray borders beyond column N (content area is A-N)
        # Prevents right-side lines in PDF. Changes are not saved (wb.Close(False)).
        ws.Range("O:Z").ClearFormats()
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1
        wb.ExportAsFixedFormat(0, pdf_path)
        wb.Close(False)
        excel.Quit()
        return pdf_path
    except Exception as e:
        print(f"   ⚠️  PDF 변환 실패: {e}")
        return None

def main():
    parser = argparse.ArgumentParser(description="최종 거래명세서 생성")
    parser.add_argument("--date", help="YYYY-MM-DD", required=True)
    parser.add_argument("--master", help="Custom master file path", default=None)
    args = parser.parse_args()

    try: target_date = datetime.strptime(args.date, "%Y-%m-%d").date()
    except:
        print(f"FAILED: Invalid date format: {args.date}", flush=True)
        import sys; sys.exit(1)

    # Override master file if specified
    global MASTER_FILE
    if args.master:
        MASTER_FILE = args.master
        print(f"📁 마스터 파일: {MASTER_FILE}")

    print(f"📅 대상 날짜: {args.date}")
    validator = ValidationReport(target_date.strftime("%Y%m%d"))
    prices = load_master_prices(target_date) # 정규화됨
    
    src_dir = os.path.join(PROJECT_ROOT, "data", "outputs", f"invoices_{target_date.strftime('%Y%m%d')}")
    if not os.path.exists(src_dir):
        print(f"FAILED: 가명세서 폴더 없음: {src_dir}", flush=True)
        import sys; sys.exit(1)
        
    invoices = glob.glob(os.path.join(src_dir, "*.xlsx"))
    invoices = [f for f in invoices if not os.path.basename(f).startswith("~")]
    
    print(f"\n📂 가명세서 목록: {len(invoices)}개 발견")
    
    drive_root = r"G:\내 드라이브\1. 도크_주문 명세서\1. 거래명세서"
    
    for fpath in invoices:
        fname = os.path.basename(fpath)
        # 중요: 원본 가명세서는 YYYYMMDD_매장명_... 형식이 아님.
        # 기존: 20260203_선데이버거클럽_거래명세서.xlsx
        # 파싱 로직 확인 요망.
        # 원본: "20260203_매장명_거래명세서.xlsx" 인가?
        parts = fname.replace(".xlsx", "").split("_")
        # 예상: ['20260203', '선데이버거클럽', '거래명세서']
        if len(parts) >= 3:
            store_name = "_".join(parts[1:-1]) # '선데이버거클럽'
        else: continue
            
        out_excel = add_prices_to_invoice(fpath, prices, target_date, store_name, validator)
        if out_excel:
            out_pdf = convert_to_pdf(out_excel)
            if out_pdf:
                print(f"   ✅ 완료: {os.path.basename(out_pdf)}")
                try:
                    target_store_folder = find_drive_folder(drive_root, store_name)
                    month_folder = f"{target_date.year}년 {target_date.month}월"
                    drive_path = os.path.join(drive_root, target_store_folder, month_folder)
                    os.makedirs(drive_path, exist_ok=True)
                    
                    dst_pdf = os.path.join(drive_path, os.path.basename(out_pdf))
                    shutil.copy2(out_pdf, dst_pdf)
                    print(f"   📂 PDF저장: {dst_pdf}")
                    
                    dst_excel = os.path.join(drive_path, os.path.basename(out_excel))
                    shutil.copy2(out_excel, dst_excel)
                    print(f"   📂 Exc저장: {dst_excel}")
                except Exception as e:
                    print(f"   ⚠️  Drive 업로드 실패: {e}")

    validator.save(PROJECT_ROOT)

    from _notify import notify
    notify("stage8_final", date_str=args.date)

if __name__ == "__main__":
    import sys
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        import traceback; traceback.print_exc()
        sys.exit(1)
