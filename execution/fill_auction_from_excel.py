"""
가락시장경매가.xlsx의 특정 시트에서 경매 데이터를 읽어
도크발주관리데이터.xlsx 단가 시트에 최고가/평균가를 채워넣는 스크립트.

Usage:
  python execution/fill_auction_from_excel.py \
    --auction "G:/내 드라이브/1. 도크_주문 명세서/0. 매입단가_자료/가락시장경매가.xlsx" \
    --master "G:/내 드라이브/1. 도크_주문 명세서/0. 매입단가_자료/도크발주관리데이터.xlsx" \
    --sheet 260222 \
    --date 2026-02-22
"""
import argparse
import re
import openpyxl

# 단가 시트 품목명 -> 가락시장경매가 품목명(정확히 Excel에 기재된 이름)
ITEM_MAP = {
    '청경채': '청경채',
    '쑥갓': '쑥갓',
    '알배기배추': '알배기배추',
    '근대': '근대',
    '깻잎(1KG)': '깻잎',
    '깻잎(4KG,큰잎)': '깻잎',
    '치커리': '치커리(일반)',
    '케일': '케일',
    '팽이': '팽이버섯',
    '새송이': '새송이버섯',
    '맛느타리': '맛느타리버섯',
    '꽃느타리': '꽃느타리버섯',
    '양송이': '양송이',
    '청양고추': '청양고추',
    '양파': '양파',
    '깐양파': '양파',
    '감자': '감자',
    '당근 수입': '당근 수입',
    '흙대파': '대파(일반)',
    '깐쪽파': '깐쪽파',
    '부추(일반)': '부추(일반)',
    '돌미나리': '돌미나리',
    '취청오이': '취청오이',
    '배추': '배추',
    '깐양배추45': '양배추',
    '양배추45': '양배추',
    '귤': '감귤',
    '양상추(일반)': '양상추(일반)',
    '양상추(국산/수입)': '양상추(일반)',
    '라디치오': None,  # 경매 데이터 없음
    '간마늘': None,  # '간마늘'은 경매 데이터 없음 (깐마늘만 있음)
    '깐마늘 대서': '깐마늘 대서',
    '깐마늘 대서(상OR대, 1K)': '깐마늘 대서',
    '깐마늘 대서(보통OR중,1K)': '깐마늘 대서',
    '깐마늘 대서(보통OR중, 20K)': '깐마늘 대서',
    '애호박': '애호박',
    '꽈리고추': '꽈리고추',
    '생표고 수입': '생표고 수입',
    '로메인(일반)': '로메인(일반)',
    '통로메인': '통로메인',
}


def parse_price(val):
    """콤마가 포함된 가격 문자열을 정수로 변환"""
    if val is None:
        return None
    s = str(val).replace(',', '').strip()
    if not s or s == '-':
        return None
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return None


def parse_unit_kg(unit_str):
    """단위 문자열에서 kg 수치 추출. 예: '20kg상자'→20, '1kg단'→1, '500g단'→0.5"""
    if not unit_str:
        return None
    s = str(unit_str).lower().strip()
    # kg 패턴
    m = re.search(r'([\d.]+)\s*kg', s)
    if m:
        return float(m.group(1))
    # g 패턴 (500g = 0.5kg)
    m = re.search(r'([\d.]+)\s*g', s)
    if m:
        return float(m.group(1)) / 1000
    return None


def parse_danga_unit_kg(unit_str):
    """단가 시트 단위에서 kg 수치 추출. 예: '15K'→15, '1K'→1, '0.5K'→0.5"""
    if not unit_str:
        return None
    s = str(unit_str).upper().strip()
    m = re.search(r'([\d.]+)\s*K', s)
    if m:
        return float(m.group(1))
    return None


def load_auction_data(auction_file, sheet_name):
    """가락시장경매가.xlsx에서 경매 데이터 로드 (유닛별 구분)"""
    print(f"Opening auction file: {auction_file}, sheet: {sheet_name}")
    wb = openpyxl.load_workbook(auction_file, data_only=True)
    ws = wb[sheet_name]

    # {품목명: [(등급, 거래단위str, kg수치, 최고가, 평균가), ...]}
    data = {}

    for row_idx in range(8, ws.max_row + 1):
        item = ws.cell(row=row_idx, column=5).value
        grade = ws.cell(row=row_idx, column=8).value
        unit_str = ws.cell(row=row_idx, column=10).value
        max_price_raw = ws.cell(row=row_idx, column=14).value
        avg_price_raw = ws.cell(row=row_idx, column=15).value

        if not item or not grade:
            continue

        item = str(item).strip()
        grade = str(grade).strip()
        unit_str = str(unit_str).strip() if unit_str else ''
        max_price = parse_price(max_price_raw)
        avg_price = parse_price(avg_price_raw)

        if max_price is None or avg_price is None:
            continue

        unit_kg = parse_unit_kg(unit_str)

        if item not in data:
            data[item] = []
        data[item].append((grade, unit_str, unit_kg, max_price, avg_price))

    print(f"Loaded {len(data)} unique items from auction data")
    wb.close()
    return data


def find_price(auction_data, auction_item_name, target_grade, target_kg):
    """
    경매 데이터에서 가격 조회.
    1. 품목명 매칭
    2. 단위 매칭 (target_kg 기준, 가장 가까운 것)
    3. 등급 매칭 (target_grade 우선, 없으면 상→특→보통→하 순)
    """
    if auction_item_name not in auction_data:
        return None, None, None, None

    entries = auction_data[auction_item_name]

    # 단위 필터링: target_kg가 있으면 가장 가까운 unit_kg 그룹 선택
    if target_kg and any(e[2] is not None for e in entries):
        # kg 정보가 있는 항목들에서 가장 가까운 단위 찾기
        available_kgs = set(e[2] for e in entries if e[2] is not None)
        if available_kgs:
            closest_kg = min(available_kgs, key=lambda x: abs(x - target_kg))
            entries = [e for e in entries if e[2] == closest_kg]

    # 등급 매칭
    grade_map = {e[0]: e for e in entries}

    # target_grade 우선 (단가시트의 품위)
    if target_grade in grade_map:
        g, u, _, mx, av = grade_map[target_grade]
        return mx, av, g, u

    # 없으면 등급 우선순위대로
    for g in ['특', '상', '보통', '하']:
        if g in grade_map:
            _, u, _, mx, av = grade_map[g]
            return mx, av, g, u

    return None, None, None, None


def fill_auction_prices(auction_file, master_file, sheet_name, target_date):
    auction_data = load_auction_data(auction_file, sheet_name)

    print(f"\nOpening master file: {master_file}")
    wb = openpyxl.load_workbook(master_file)
    ws = wb["단가"]

    target_rows = []
    for row_idx in range(2, ws.max_row + 1):
        date_cell = ws.cell(row=row_idx, column=2).value
        item_cell = ws.cell(row=row_idx, column=3).value

        is_match = False
        if date_cell:
            if str(date_cell).startswith(target_date):
                is_match = True
            elif hasattr(date_cell, 'strftime') and date_cell.strftime("%Y-%m-%d") == target_date:
                is_match = True

        if is_match and item_cell:
            category = str(ws.cell(row=row_idx, column=1).value or '')
            grade = str(ws.cell(row=row_idx, column=4).value or '').strip()
            unit = str(ws.cell(row=row_idx, column=5).value or '').strip()
            target_rows.append((row_idx, str(item_cell).strip(), category, grade, unit))

    print(f"Found {len(target_rows)} items for {target_date}")

    updates = 0
    skipped = 0
    no_match = []

    for row_idx, item_name, category, grade, unit in target_rows:
        # 소분/재고 rows -> skip
        if '소분' in category:
            skipped += 1
            continue

        # 매핑 확인
        if item_name not in ITEM_MAP:
            no_match.append((row_idx, item_name, "No mapping defined"))
            continue

        auction_key = ITEM_MAP[item_name]
        if auction_key is None:
            no_match.append((row_idx, item_name, "Explicitly excluded (no auction data)"))
            continue

        # 단가 시트 단위에서 kg 추출
        target_kg = parse_danga_unit_kg(unit)

        # 단가 시트 등급 정규화
        target_grade = grade.replace('or대', '').replace('or중', '').strip()
        if target_grade in ['상or대', '상OR대']:
            target_grade = '상'
        elif target_grade in ['보통or중', '보통OR중']:
            target_grade = '보통'

        max_price, avg_price, matched_grade, matched_unit = find_price(
            auction_data, auction_key, target_grade, target_kg
        )

        if max_price and avg_price:
            cell_f = ws.cell(row=row_idx, column=6)
            cell_g = ws.cell(row=row_idx, column=7)

            if not isinstance(cell_f, openpyxl.cell.cell.MergedCell):
                cell_f.value = max_price
            if not isinstance(cell_g, openpyxl.cell.cell.MergedCell):
                cell_g.value = avg_price

            print(f"  R{row_idx}: {item_name}({grade} {unit}) -> {auction_key}({matched_grade} {matched_unit}) 최고={max_price:,} 평균={avg_price:,}")
            updates += 1
        else:
            no_match.append((row_idx, item_name, f"No auction data for '{auction_key}'"))

    print(f"\n=== Results ===")
    print(f"Updated: {updates}")
    print(f"Skipped (소분/재고): {skipped}")
    print(f"No match: {len(no_match)}")
    for r, name, reason in no_match:
        print(f"  R{r}: {name} - {reason}")

    if updates > 0:
        print(f"\nSaving to {master_file}...")
        wb.save(master_file)
        print("Done.")
    else:
        print("\nNo updates made.")

    wb.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--auction", required=True, help="Path to 가락시장경매가.xlsx")
    parser.add_argument("--master", required=True, help="Path to 도크발주관리데이터.xlsx")
    parser.add_argument("--sheet", required=True, help="Sheet name in auction file (e.g. 260222)")
    parser.add_argument("--date", required=True, help="Target date (YYYY-MM-DD)")
    args = parser.parse_args()

    fill_auction_prices(args.auction, args.master, args.sheet, args.date)
