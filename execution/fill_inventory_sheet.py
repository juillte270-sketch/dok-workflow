"""
재고 시트 1차/직납 자동 입력 스크립트.
processed_orders의 "재고, 창고소분" 섹션을 파싱하여
도크발주관리데이터.xlsx 재고 시트의 1차/직납 컬럼에 자동 입력.

Usage:
  python execution/fill_inventory_sheet.py --date 2026-02-23
  python execution/fill_inventory_sheet.py --date 2026-02-23 --dry-run
"""
import argparse
import json
import os
import re
import time

import openpyxl

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_MASTER = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx"
MAPPINGS_PATH = os.path.join(project_root, "skills", "order_processing", "resources", "mappings.json")

# processed_orders 품목명 → 재고시트 품목명 기본 매핑
DEFAULT_ITEM_MAP = {
    "감자": "감자(KG)",
    "세척당근": "당근(KG)",
    "청양고추": "고추(KG)",
    "꽈리고추": "꽈리고추(KG)",
    "팽이": "팽이버섯(팩)",
    "깐마늘(대)": "마늘1kg",
    "애호박": "애호박(개)",
    "맛느타리": "맛느타리(팩)",
    "적채": "적채(통)",
    "양상추": "양상추(12수/통)",
    "매운고춧가루": "매운 고춧가루",
    "굵은고춧가루": "굵은 고춧가루",
    "참나물": "참나물(kg)",
    "레몬": "레몬(개)",
    "새송이": "새송이(KG)",
    "알배기": "알배기(통)",
    "양배추": "양배추(통)",
    "깐양배추": "양배추(통)",
    "무": "무(통)",
    "표고버섯": "표고버섯(KG)",
    "취청오이": "취청오이(kg)",
    "청경채": "청경채(KG)",
    "깐대파": "깐대파(단)",
    "깐쪽파": "깐쪽파(단)",
    "김치": "김치(박스)",
    "쌀": "쌀",
    "골드파인애플": "골드파인애플(통)",
    "마카로니": "마카로니",
    "아카시아청": "아카시아청",
}

# 깐마늘(대) 특수 매핑: 수량별
# 4kg → 마늘1kg (count=4), 0.5kg → 마늘500g (count=1)
GARLIC_MAP = {
    "마늘1kg": 1.0,   # 1kg 이상이면 이 행, 값 = qty
    "마늘500g": 0.5,   # 0.5kg이면 이 행, 값 = qty/0.5
}


def load_item_map():
    """mappings.json에서 inventory_item_map 로드 (없으면 기본값)"""
    item_map = dict(DEFAULT_ITEM_MAP)
    try:
        with open(MAPPINGS_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if 'inventory_item_map' in data:
            item_map.update(data['inventory_item_map'])
    except Exception as e:
        print(f"⚠ mappings.json 로드 실패 (기본값 사용): {e}")
    return item_map


def parse_stock_section(processed_orders_path):
    """processed_orders에서 '-재고, 창고소분' 섹션 파싱.

    Returns: list of (item_name, qty, unit) tuples (동일 품목은 합산)
    """
    raw_items = []
    in_stock_section = False

    with open(processed_orders_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # Pattern: "품목 수량단위(매장)" e.g. "감자 15kg(강남)", "팽이 5개(마곡)"
    item_pattern = re.compile(
        r'^(.+?)\s+(\d+(?:\.\d+)?)\s*'
        r'(kg|KG|Kg|k|K|개|팩|박스|봉|통|포|단|판|망|발)'
        r'(.*)$'
    )

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if line.startswith('-'):
            section_name = line.lstrip('-').strip()
            if '재고' in section_name and '창고소분' in section_name:
                in_stock_section = True
                continue
            elif in_stock_section:
                break
            continue

        if in_stock_section:
            # Skip memo lines
            if line.startswith('('):
                continue
            match = item_pattern.match(line)
            if match:
                item_name = match.group(1).strip()
                qty = float(match.group(2))
                unit = match.group(3).lower()
                raw_items.append((item_name, qty, unit))

    # Aggregate by item name
    aggregated = {}
    for name, qty, unit in raw_items:
        if name in aggregated:
            aggregated[name] = (aggregated[name][0] + qty, unit)
        else:
            aggregated[name] = (qty, unit)

    return aggregated


def find_date_columns(ws, target_date):
    """재고 시트에서 target_date에 해당하는 컬럼 그룹 찾기.

    Row 2에서 "M/DD 재고현황" 패턴 검색.
    Returns: (item_col, shipment_col) — 품목 컬럼, 1차/직납 컬럼
    """
    from datetime import datetime
    dt = datetime.strptime(target_date, "%Y-%m-%d")
    month = dt.month
    day = dt.day

    # Exact date match: "2/23"이 "12/23" 안에서 매칭되는 것 방지
    date_prefix = f"{month}/{day}"

    for row_idx in range(1, 6):
        for col_idx in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=row_idx, column=col_idx).value
            if cell_val is None:
                continue
            cell_str = str(cell_val).strip()
            if "재고" not in cell_str:
                continue
            idx = cell_str.find(date_prefix)
            if idx < 0:
                continue
            if idx > 0 and cell_str[idx - 1].isdigit():
                continue  # e.g., "12/23" when looking for "2/23"
            # 품목 = col_idx, 재고 = +1, 1차/직납 = +2
            return col_idx, col_idx + 2

    return None, None


def find_item_rows(ws, item_col):
    """품목 컬럼에서 아이템명 → 행 번호 매핑"""
    row_map = {}
    for r in range(4, min(ws.max_row + 1, 200)):
        val = ws.cell(row=r, column=item_col).value
        if val:
            row_map[str(val).strip()] = r
    return row_map


def map_item(stock_name, qty, unit, item_map, item_rows):
    """품목명을 재고시트 행에 매핑.

    Returns: (sheet_item_name, row, final_qty) or (None, None, None)
    """
    # Special handling for 깐마늘(대) - unit determines row
    if '깐마늘' in stock_name:
        if qty < 1:
            target = "마늘500g"
            final_qty = qty / 0.5  # 0.5kg = 1 unit
        else:
            target = "마늘1kg"
            final_qty = qty  # 4kg = 4 units
        if target in item_rows:
            return target, item_rows[target], final_qty
        return None, None, None

    # 1. Direct mapping (exact, then base-keyword fuzzy)
    sheet_name = item_map.get(stock_name)
    if sheet_name:
        if sheet_name in item_rows:
            return sheet_name, item_rows[sheet_name], qty
        # Fuzzy: "팽이버섯(팩)" matches "팽이버섯(개)" by base keyword
        sn_base = sheet_name.split('(')[0].strip()
        if sn_base:
            for si, sr in item_rows.items():
                si_base = si.split('(')[0].strip()
                if sn_base == si_base:
                    return si, sr, qty

    # 2. Substring matching: check if stock_name is contained in any sheet item
    for sheet_item, row in item_rows.items():
        sheet_base = sheet_item.split('(')[0].strip()
        if stock_name == sheet_base or sheet_base in stock_name or stock_name in sheet_base:
            return sheet_item, row, qty

    return None, None, None


def fill_inventory(master_file, target_date, dry_run=False):
    """메인: processed_orders 재고 섹션 → 재고시트 1차/직납 입력"""
    date_str = target_date.replace("-", "")
    processed_path = os.path.join(
        project_root, "data", "inputs", f"processed_orders_{date_str}.txt"
    )

    if not os.path.exists(processed_path):
        print(f"Error: {processed_path} not found")
        return

    # 1. Parse stock section
    stock_items = parse_stock_section(processed_path)
    if not stock_items:
        print("No items found in '재고, 창고소분' section.")
        return

    print(f"=== Parsed stock items ({len(stock_items)}) ===")
    for name, (qty, unit) in stock_items.items():
        print(f"  {name}: {qty}{unit}")

    # 2. Load item mapping
    item_map = load_item_map()

    # 3. Open master file
    print(f"\nOpening: {master_file}")
    wb = openpyxl.load_workbook(master_file)
    ws = wb["재고"]

    # 4. Find date columns
    item_col, shipment_col = find_date_columns(ws, target_date)
    if item_col is None:
        print(f"Error: Could not find columns for {target_date} in 재고 sheet")
        wb.close()
        return

    print(f"Columns: 품목=C{item_col}, 1차/직납=C{shipment_col}")

    # 5. Build item → row mapping
    item_rows = find_item_rows(ws, item_col)
    print(f"재고 sheet items: {len(item_rows)}")

    # 6. Match and fill
    updates = 0
    no_match = []

    print(f"\n=== Filling 1차/직납 ===")
    for stock_name, (qty, unit) in stock_items.items():
        sheet_name, row, final_qty = map_item(
            stock_name, qty, unit, item_map, item_rows
        )

        if sheet_name and row:
            cell = ws.cell(row=row, column=shipment_col)
            if isinstance(cell, openpyxl.cell.cell.MergedCell):
                print(f"  {stock_name} -> {sheet_name} (R{row}): MERGED, skipped")
                continue

            if dry_run:
                print(f"  {stock_name} -> {sheet_name} (R{row}): {final_qty} [DRY RUN]")
            else:
                cell.value = final_qty
                print(f"  {stock_name} -> {sheet_name} (R{row}): {final_qty}")
            updates += 1
        else:
            no_match.append(stock_name)
            print(f"  {stock_name}: NO MATCH")

    # Results
    print(f"\n=== Results ===")
    print(f"Updated: {updates}/{len(stock_items)}")
    if no_match:
        print(f"No match: {no_match}")

    # Save
    if updates > 0 and not dry_run:
        saved = False
        for attempt in range(5):
            try:
                print(f"\nSaving... (attempt {attempt + 1})")
                wb.save(master_file)
                print("Done.")
                saved = True
                break
            except PermissionError:
                if attempt < 4:
                    print("  File locked, retrying in 3s...")
                    time.sleep(3)
                else:
                    import tempfile
                    tmp = os.path.join(
                        tempfile.gettempdir(), "도크발주관리데이터_inventory.xlsx"
                    )
                    wb.save(tmp)
                    print(f"  Saved to temp: {tmp}")
                    print("  Close the file and copy manually.")
                    saved = True
    elif dry_run:
        print("\n[DRY RUN] No changes saved.")

    wb.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="재고 시트 1차/직납 자동 입력")
    parser.add_argument("--master", default=DEFAULT_MASTER)
    parser.add_argument("--date", required=True, help="YYYY-MM-DD")
    parser.add_argument("--dry-run", action="store_true", help="Preview only")
    args = parser.parse_args()

    fill_inventory(args.master, args.date, args.dry_run)
