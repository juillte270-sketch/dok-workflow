"""
재고 시트 날짜 그룹 복사 + 입고량 자동 계산 스크립트.

기능:
  A) 이전 날짜 재고현황 옆에 새 날짜 그룹(6컬럼) 생성
  B) mappings.json의 inventory_reorder_rules로 필요 입고량 계산

Usage:
  python execution/manage_inventory_date.py --prev-date 2026-02-23 --new-date 2026-02-24 --master "G:/.../도크발주관리데이터.xlsx"
  python execution/manage_inventory_date.py --prev-date 2026-02-23 --new-date 2026-02-24 --dry-run
"""
import argparse
import json
import math
import os
import time
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_MASTER = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx"
MAPPINGS_PATH = os.path.join(project_root, "skills", "order_processing", "resources", "mappings.json")

# 재고 시트 구조: 3섹션
# 채소: rows 4-31, 과일 서브헤더: row 33, 과일: 34-36, 공산품 서브헤더: row 38, 공산품: 39-43
DATA_ROWS = list(range(4, 32))       # 채소 R4-R31
FRUIT_HEADER_ROW = 33
FRUIT_ROWS = list(range(34, 37))     # 과일 R34-R36
INDUSTRIAL_HEADER_ROW = 38
INDUSTRIAL_ROWS = list(range(39, 44))  # 공산품 R39-R43
ALL_DATA_ROWS = DATA_ROWS + FRUIT_ROWS + INDUSTRIAL_ROWS
HEADER_ROW = 2
SUBHEADER_ROW = 3
SECTION_SUBHEADER_ROWS = [FRUIT_HEADER_ROW, INDUSTRIAL_HEADER_ROW]

# 6컬럼 구조: 품목(+0), 재고(+1), 1차/직납(+2), 2차출고(+3), 입고(+4), 현재고(+5)
COL_ITEM = 0
COL_STOCK = 1
COL_SHIPMENT = 2
COL_HELO = 3
COL_INCOMING = 4
COL_CURRENT = 5
GROUP_SIZE = 6


def load_mappings():
    """mappings.json 로드"""
    with open(MAPPINGS_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)


def find_date_columns(ws, target_date):
    """재고 시트에서 target_date 컬럼 그룹 시작 컬럼 찾기.
    Row 2에서 "M/DD 재고현황" 패턴 검색.
    Returns: 그룹 시작 컬럼 (품목 컬럼) 또는 None
    """
    dt = datetime.strptime(target_date, "%Y-%m-%d")
    date_prefix = f"{dt.month}/{dt.day}"

    for col_idx in range(1, ws.max_column + 1):
        cell_val = ws.cell(row=HEADER_ROW, column=col_idx).value
        if cell_val is None:
            continue
        cell_str = str(cell_val).strip()
        if "재고" not in cell_str:
            continue
        idx = cell_str.find(date_prefix)
        if idx < 0:
            continue
        # "12/23"에서 "2/23" 오매칭 방지
        if idx > 0 and cell_str[idx - 1].isdigit():
            continue
        return col_idx

    return None


def copy_cell_format(src_cell, dst_cell):
    """셀 포맷 복사 (font, fill, border, alignment, number_format)"""
    if src_cell.font:
        dst_cell.font = copy(src_cell.font)
    if src_cell.fill:
        dst_cell.fill = copy(src_cell.fill)
    if src_cell.border:
        dst_cell.border = copy(src_cell.border)
    if src_cell.alignment:
        dst_cell.alignment = copy(src_cell.alignment)
    if src_cell.number_format:
        dst_cell.number_format = src_cell.number_format


def add_new_date(ws, prev_start_col, new_date_str):
    """이전 날짜 그룹 옆에 새 날짜 그룹 6컬럼 생성.

    Args:
        ws: 재고 시트
        prev_start_col: 이전 날짜 그룹 시작 컬럼
        new_date_str: "YYYY-MM-DD" 새 날짜

    Returns: 새 그룹 시작 컬럼
    """
    new_start = prev_start_col + GROUP_SIZE
    dt = datetime.strptime(new_date_str, "%Y-%m-%d")
    header_text = f"{dt.month}/{dt.day} 재고현황"

    prev_cols = [prev_start_col + i for i in range(GROUP_SIZE)]
    new_cols = [new_start + i for i in range(GROUP_SIZE)]

    # --- Row 2: 메인 헤더 ---
    src = ws.cell(row=HEADER_ROW, column=prev_start_col)
    dst = ws.cell(row=HEADER_ROW, column=new_start)
    dst.value = header_text
    copy_cell_format(src, dst)
    # 나머지 헤더 컬럼 포맷 복사
    for offset in range(1, GROUP_SIZE):
        src2 = ws.cell(row=HEADER_ROW, column=prev_cols[offset])
        dst2 = ws.cell(row=HEADER_ROW, column=new_cols[offset])
        copy_cell_format(src2, dst2)

    # --- Row 3: 서브헤더 (품목/재고/1차/직납/2차출고/입고/현 재고) ---
    for offset in range(GROUP_SIZE):
        src = ws.cell(row=SUBHEADER_ROW, column=prev_cols[offset])
        dst = ws.cell(row=SUBHEADER_ROW, column=new_cols[offset])
        dst.value = src.value
        copy_cell_format(src, dst)

    # --- 섹션 서브헤더 (R33 과일, R38 공산품) ---
    for sh_row in SECTION_SUBHEADER_ROWS:
        for offset in range(GROUP_SIZE):
            src = ws.cell(row=sh_row, column=prev_cols[offset])
            dst = ws.cell(row=sh_row, column=new_cols[offset])
            dst.value = src.value
            copy_cell_format(src, dst)

    # --- 데이터 행 ---
    prev_current_letter = get_column_letter(prev_start_col + COL_CURRENT)  # 이전 현재고 컬럼 letter
    new_stock_letter = get_column_letter(new_start + COL_STOCK)
    new_shipment_letter = get_column_letter(new_start + COL_SHIPMENT)
    new_helo_letter = get_column_letter(new_start + COL_HELO)
    new_incoming_letter = get_column_letter(new_start + COL_INCOMING)
    new_current_letter = get_column_letter(new_start + COL_CURRENT)

    for row in ALL_DATA_ROWS:
        # col+0: 품목명 복사
        src = ws.cell(row=row, column=prev_cols[COL_ITEM])
        dst = ws.cell(row=row, column=new_cols[COL_ITEM])
        dst.value = src.value
        copy_cell_format(src, dst)

        # col+1: 재고 = 이전 현재고 수식
        dst_stock = ws.cell(row=row, column=new_cols[COL_STOCK])
        dst_stock.value = f"={prev_current_letter}{row}"
        copy_cell_format(ws.cell(row=row, column=prev_cols[COL_STOCK]), dst_stock)

        # col+2: 1차/직납 (빈칸)
        dst_ship = ws.cell(row=row, column=new_cols[COL_SHIPMENT])
        dst_ship.value = None
        copy_cell_format(ws.cell(row=row, column=prev_cols[COL_SHIPMENT]), dst_ship)

        # col+3: 2차출고 (빈칸)
        dst_helo = ws.cell(row=row, column=new_cols[COL_HELO])
        dst_helo.value = None
        copy_cell_format(ws.cell(row=row, column=prev_cols[COL_HELO]), dst_helo)

        # col+4: 입고 (빈칸, 나중에 calculate_incoming에서 채움)
        dst_inc = ws.cell(row=row, column=new_cols[COL_INCOMING])
        dst_inc.value = None
        copy_cell_format(ws.cell(row=row, column=prev_cols[COL_INCOMING]), dst_inc)

        # col+5: 현재고 수식 = 재고 + 입고 - 2차출고 - 1차/직납
        formula = (
            f"={new_stock_letter}{row}"
            f"+{new_incoming_letter}{row}"
            f"-{new_helo_letter}{row}"
            f"-{new_shipment_letter}{row}"
        )
        dst_cur = ws.cell(row=row, column=new_cols[COL_CURRENT])
        dst_cur.value = formula
        copy_cell_format(ws.cell(row=row, column=prev_cols[COL_CURRENT]), dst_cur)

    print(f"Created {header_text} at columns {get_column_letter(new_start)}-{get_column_letter(new_start + GROUP_SIZE - 1)}")
    return new_start


def get_weekday_value(rule_dict, date_str):
    """요일별 값 반환. rule_dict는 {"sun_wed": 3, "thu_fri": 4} 등의 형태."""
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    dow = dt.weekday()  # 0=Mon, 1=Tue, ..., 6=Sun

    # 키 매핑: sun_wed = 일~수(6,0,1,2,3), thu_fri = 목금(3,4)
    # sun_thu = 일~목(6,0,1,2,3), fri = 금(4)
    for key, val in rule_dict.items():
        days = key.lower().replace("-", "_")
        if days == "sun_wed" and dow in (6, 0, 1, 2):  # 일월화수
            return val
        if days == "thu_fri" and dow in (3, 4):  # 목금
            return val
        if days == "sun_thu" and dow in (6, 0, 1, 2, 3):  # 일월화수목
            return val
        if days == "fri" and dow == 4:  # 금
            return val

    # 폴백: 첫 번째 값 반환
    return list(rule_dict.values())[0]


def calculate_incoming(ws, new_start_col, new_date_str, reorder_rules, alert_rules, master_path=None, dry_run=False):
    """입고량 계산 및 기입.

    Args:
        ws: 재고 시트
        new_start_col: 새 날짜 그룹 시작 컬럼
        new_date_str: 새 날짜 "YYYY-MM-DD"
        reorder_rules: mappings.json의 inventory_reorder_rules
        alert_rules: mappings.json의 inventory_alert_rules
        master_path: 마스터 파일 경로 (data_only 로드용)

    Returns: (updates_count, alerts_list)
    """
    stock_col = new_start_col + COL_STOCK
    incoming_col = new_start_col + COL_INCOMING
    item_col = new_start_col + COL_ITEM

    # 품목명 → 행 매핑
    item_rows = {}
    for row in ALL_DATA_ROWS:
        val = ws.cell(row=row, column=item_col).value
        if val:
            item_rows[str(val).strip()] = row

    # 재고값 읽기 (수식이면 캐시된 값 사용 — openpyxl data_only 안 쓰므로 직접 계산 필요)
    # 재고(col+1)는 이전 현재고를 참조하는 수식이므로, 이전 현재고를 직접 읽는다
    prev_current_col = new_start_col + COL_CURRENT - GROUP_SIZE  # 이전 그룹의 현재고

    stock_values = {}
    for item_name, row in item_rows.items():
        cell = ws.cell(row=row, column=prev_current_col)
        val = cell.value
        if val is None:
            stock_values[item_name] = 0
        elif isinstance(val, (int, float)):
            stock_values[item_name] = val
        elif isinstance(val, str) and val.startswith("="):
            # 수식인 경우 — data_only로 재로드 필요
            stock_values[item_name] = None  # 나중에 처리
        else:
            try:
                stock_values[item_name] = float(val)
            except (ValueError, TypeError):
                stock_values[item_name] = 0

    # 수식 셀이 있으면 data_only로 한번 더 읽기
    has_formulas = any(v is None for v in stock_values.values())
    if has_formulas:
        print("  Note: Some cells have formulas. Reading cached values...")
        # openpyxl data_only=True로 캐시된 값 읽기
        try:
            wb2 = openpyxl.load_workbook(master_path or DEFAULT_MASTER, data_only=True)
            ws2 = wb2["재고"]
            for item_name, row in item_rows.items():
                if stock_values[item_name] is None:
                    cached = ws2.cell(row=row, column=prev_current_col).value
                    stock_values[item_name] = cached if isinstance(cached, (int, float)) else 0
            wb2.close()
        except Exception as e:
            print(f"  Warning: Could not read cached values: {e}")
            for k in stock_values:
                if stock_values[k] is None:
                    stock_values[k] = 0

    updates = 0
    alerts = []

    print(f"\n=== 입고량 계산 ({new_date_str}) ===")

    # 팽이 특수 처리를 위한 변수
    pengi_box_incoming = 0

    for item_name, row in item_rows.items():
        stock = stock_values.get(item_name, 0)
        if stock is None:
            stock = 0

        rule = reorder_rules.get(item_name)
        if not rule:
            continue

        rtype = rule["type"]
        incoming = 0

        if rtype == "skip":
            continue

        elif rtype == "manual":
            print(f"  {item_name}: 재고={stock} [수동 판단 필요]")
            continue

        elif rtype == "range":
            # min/max 또는 요일별 min/max
            if "min_weekday" in rule:
                min_val = get_weekday_value(rule["min_weekday"], new_date_str)
                max_val = get_weekday_value(rule["max_weekday"], new_date_str)
            else:
                min_val = rule["min"]
                max_val = rule["max"]

            if stock < min_val:
                incoming = max_val - stock
                if incoming < 0:
                    incoming = 0

        elif rtype == "threshold":
            threshold = rule["threshold"]
            order_qty = rule["order_qty"]
            if stock <= threshold:
                incoming = order_qty

        elif rtype == "tiered":
            # tiers는 threshold 오름차순 정렬 필요 — 낮은 threshold부터 체크
            tiers = sorted(rule["tiers"], key=lambda t: t["threshold"])
            for tier in tiers:
                if stock <= tier["threshold"]:
                    incoming = tier["order_qty"]
                    break  # 가장 낮은 threshold에 매칭되면 더 큰 수량 발주

        elif rtype == "target":
            target = rule["target"]
            threshold = rule.get("threshold", target)
            if stock <= threshold:
                incoming = target - stock
                if incoming < 0:
                    incoming = 0

        elif rtype == "target_unit":
            target = rule["target"]
            unit = rule["unit"]
            if stock < target:
                needed = target - stock
                incoming = math.ceil(needed / unit) * unit

        elif rtype == "special_box":
            # 팽이버섯(개): 재고<0이면 팽이(박스) 2차출고 1 + 입고 per_box
            per_box = rule["per_box"]
            if stock < 0:
                incoming = per_box
                pengi_box_incoming = 1  # 박스 2차출고 필요 표시
                print(f"  ** {item_name}: 재고={stock} < 0 → 입고 {incoming}개 (팽이(박스) 2차출고 1 필요)")

        if incoming > 0:
            cell = ws.cell(row=row, column=incoming_col)
            if dry_run:
                print(f"  {item_name}: 재고={stock} → 입고 {incoming} [DRY RUN]")
            else:
                cell.value = incoming
                print(f"  {item_name}: 재고={stock} → 입고 {incoming}")
            updates += 1
        elif rtype not in ("skip", "manual", "special_box"):
            print(f"  {item_name}: 재고={stock} → 충분 (입고 불필요)")

    # 팽이(박스) 2차출고 처리
    if pengi_box_incoming > 0 and "팽이(박스)" in item_rows:
        box_row = item_rows["팽이(박스)"]
        helo_col = new_start_col + COL_HELO
        if not dry_run:
            ws.cell(row=box_row, column=helo_col).value = pengi_box_incoming
        print(f"  팽이(박스): 2차출고 {pengi_box_incoming} 설정 {'[DRY RUN]' if dry_run else ''}")

    # 공산품 알림 체크
    print(f"\n=== 공산품 알림 체크 ===")
    for item_name, alert_rule in alert_rules.items():
        if item_name in stock_values:
            stock = stock_values[item_name]
            if stock is not None and stock <= alert_rule["threshold"]:
                msg = alert_rule["message"]
                alerts.append(f"{msg} (현재고: {stock})")
                print(f"  [!] {msg} (현재고: {stock})")
        else:
            # 부분 매칭 시도
            for sheet_item, row in item_rows.items():
                if item_name in sheet_item or sheet_item in item_name:
                    stock = stock_values.get(sheet_item, 0)
                    if stock is not None and stock <= alert_rule["threshold"]:
                        msg = alert_rule["message"]
                        alerts.append(f"{msg} (현재고: {stock})")
                        print(f"  [!] {msg} (현재고: {stock})")
                    break

    if not alerts:
        print("  공산품 알림 없음")

    return updates, alerts


def main():
    parser = argparse.ArgumentParser(description="재고 시트 날짜 그룹 복사 + 입고량 자동 계산")
    parser.add_argument("--prev-date", required=True, help="이전 날짜 YYYY-MM-DD")
    parser.add_argument("--new-date", required=True, help="새 날짜 YYYY-MM-DD")
    parser.add_argument("--master", default=DEFAULT_MASTER, help="마스터 파일 경로")
    parser.add_argument("--dry-run", action="store_true", help="미리보기 (저장 안 함)")
    parser.add_argument("--skip-copy", action="store_true", help="날짜 그룹 복사 생략 (이미 존재할 때)")
    parser.add_argument("--skip-incoming", action="store_true", help="입고량 계산 생략")
    args = parser.parse_args()

    mappings = load_mappings()
    reorder_rules = mappings.get("inventory_reorder_rules", {})
    alert_rules = mappings.get("inventory_alert_rules", {})

    print(f"Master: {args.master}")
    print(f"Previous date: {args.prev_date}")
    print(f"New date: {args.new_date}")
    dt = datetime.strptime(args.new_date, "%Y-%m-%d")
    weekday_names = ["월", "화", "수", "목", "금", "토", "일"]
    print(f"요일: {weekday_names[dt.weekday()]}요일")
    if args.dry_run:
        print("[DRY RUN MODE]")

    wb = openpyxl.load_workbook(args.master)
    ws = wb["재고"]

    # Step A: 날짜 그룹 복사
    new_start_col = None
    if not args.skip_copy:
        prev_start = find_date_columns(ws, args.prev_date)
        if prev_start is None:
            print(f"Error: {args.prev_date} 컬럼을 찾을 수 없습니다.")
            wb.close()
            return

        print(f"\n이전 날짜 그룹 시작: {get_column_letter(prev_start)} (col {prev_start})")

        # 새 날짜가 이미 있는지 체크
        existing = find_date_columns(ws, args.new_date)
        if existing:
            print(f"Warning: {args.new_date} 그룹이 이미 존재합니다 ({get_column_letter(existing)}). 복사를 건너뜁니다.")
            new_start_col = existing
        else:
            new_start_col = add_new_date(ws, prev_start, args.new_date)
    else:
        new_start_col = find_date_columns(ws, args.new_date)
        if new_start_col is None:
            print(f"Error: {args.new_date} 컬럼을 찾을 수 없습니다. --skip-copy를 제거하세요.")
            wb.close()
            return

    # Step B: 입고량 계산
    if not args.skip_incoming:
        updates, alerts = calculate_incoming(
            ws, new_start_col, args.new_date, reorder_rules, alert_rules, args.master, args.dry_run
        )
        print(f"\n=== 결과 ===")
        print(f"입고 기입: {updates}건")
        if alerts:
            print(f"공산품 알림: {len(alerts)}건")
            for a in alerts:
                print(f"  - {a}")

    # Save
    if not args.dry_run:
        saved = False
        for attempt in range(5):
            try:
                print(f"\nSaving... (attempt {attempt + 1})")
                wb.save(args.master)
                print("Done!")
                saved = True
                break
            except PermissionError:
                if attempt < 4:
                    print("  File locked, retrying in 3s...")
                    time.sleep(3)
                else:
                    import tempfile
                    tmp = os.path.join(tempfile.gettempdir(), "도크발주관리데이터_inventory_date.xlsx")
                    wb.save(tmp)
                    print(f"  Saved to temp: {tmp}")
                    print("  Close the master file and copy manually.")
                    saved = True
    else:
        print("\n[DRY RUN] No changes saved.")

    wb.close()


if __name__ == "__main__":
    import sys
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        import traceback; traceback.print_exc()
        sys.exit(1)
    from _notify import notify
    notify("stage4_inventory")
