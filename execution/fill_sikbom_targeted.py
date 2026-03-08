"""
식봄(Foodspring) 가격을 특정 품목에 대해 검색하고 단가 시트에 채워넣는 스크립트.
사용자 피드백을 반영한 타겟 검색어 및 업체 지정.

Usage:
  python execution/fill_sikbom_targeted.py \
    --master "G:/내 드라이브/1. 도크_주문 명세서/0. 매입단가_자료/도크발주관리데이터.xlsx" \
    --date 2026-02-22
"""
import argparse
import re
import time
import sys
import os

# Windows cp949 인코딩 에러 방지 — 식봄 검색 결과에 이모지/특수문자 포함됨
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__))))
from fetch_foodspring_prices import FoodspringScraper

import openpyxl

# 다중 행 품목 — 등급/단위 조합별로 별도 행 매칭 필요
# 'grade+unit': 등급(Col D) + 단위(Col E) 복합키
# 'grade': 등급(Col D)만으로 구분
MULTI_ROW_ITEMS = {
    '깐마늘 대서': 'grade+unit',
    '애호박': 'grade',
}

# 타겟 검색 설정
# match_grade: MULTI_ROW_ITEMS 품목의 등급 매칭 (등급값에 이 문자열 포함 여부)
# match_unit: MULTI_ROW_ITEMS 품목의 단위 매칭 (대소문자 무시)
# unit_convert: 가격을 이 배수로 환산 (예: 2kg->10kg면 5)
SEARCH_CONFIG = [
    {
        'item': '양파',
        'search': '양파 15kg 국내산',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '15kg',
    },
    {
        'item': '감자',
        'search': '감자 왕왕',
        'suppliers': ['CJ프레시웨이'],
        'filter': '20kg',
    },
    {
        'item': '돌미나리',
        'search': '돌미나리',
        'suppliers': ['다봄푸드', 'CJ프레시웨이'],
        'filter': '4kg',
    },
    {
        'item': '취청오이',
        'search': '청오이',
        'suppliers': ['CJ프레시웨이'],
        'filter': '50입',
    },
    {
        'item': '배추',
        'search': '배추52',
        'suppliers': ['CJ프레시웨이', '세현F&B', '세현f&b'],
        'filter': '52',
    },
    {
        'item': '깐양배추45',
        'search': '깐양배추',
        'suppliers': ['CJ프레시웨이'],
        'filter': '2kg',
        'unit_convert': 5,  # 2kg -> 10kg: x5
    },
    {
        'item': '깐양배추42',
        'search': '깐양배추',
        'suppliers': ['CJ프레시웨이'],
        'filter': '2kg',
        'unit_convert': 4,  # 2kg -> 8kg: x4
    },
    {
        'item': '귤',
        'search': '귤 5kg',
        'suppliers': ['케이에프피(강남)', '케이에프피'],
        'filter': None,  # M사이즈 확인
    },
    {
        'item': '양상추(일반)',
        'search': '양상추',
        'suppliers': ['CJ프레시웨이'],
        'filter': '12입',
    },
    {
        'item': '라디치오',
        'search': '라디치오',
        'suppliers': ['농장에서바로', '쉐프의정원'],
        'filter': '4.5kg',
    },
    {
        'item': '간마늘',
        'search': '간마늘 1kg 국내산',
        'suppliers': ['다봄푸드', 'CJ프레시웨이'],
        'filter': '1kg',
        'exclude': '냉동',
    },
    {   # 깐마늘(대) 1K — "상or대" 등급 행
        'item': '깐마늘 대서',
        'search': '깐마늘(대) 1kg',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '1kg',
        'exclude': '냉동',
        'match_grade': '대',
        'match_unit': '1K',
    },
    {   # 깐마늘(중) 1K — "보통or중" 등급 행
        'item': '깐마늘 대서',
        'search': '깐마늘(중) 1kg',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '1kg',
        'exclude': '냉동',
        'match_grade': '중',
        'match_unit': '1K',
    },
    {   # 깐마늘(중) 20K — "보통or중" 등급, 20kg 환산
        'item': '깐마늘 대서',
        'search': '깐마늘(중) 1kg',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '1kg',
        'exclude': '냉동',
        'match_grade': '중',
        'match_unit': '20K',
        'unit_convert': 20,
    },
    {   # 애호박(특) — 인큐애호박(특,20개 내외,국내산)BOX 다봄푸드
        'item': '애호박',
        'search': '인큐 애호박',
        'suppliers': ['다봄푸드'],
        'filter': '인큐',
        'match_grade': '특',
        'unit_convert': 20,  # 개당 가격 × 20 = BOX 가격
    },
    {   # 애호박(상) — 애호박(국내산 상등급_20입/BOX) CJ프레시웨이
        'item': '애호박',
        'search': '애호박 20입',
        'suppliers': ['CJ프레시웨이'],
        'filter': '20입',
        'match_grade': '상',
    },
    {   # 아욱 — CJ프레시웨이 4KG/BOX
        'item': '아욱',
        'search': '아욱',
        'suppliers': ['CJ프레시웨이'],
        'filter': '4kg',
    },
    {
        'item': '근대',
        'search': '근대',
        'suppliers': ['CJ프레시웨이'],
        'filter': '4kg',
        'exclude': '적',  # 적근대 제외
    },
    {
        'item': '청양고추',
        'search': '청양고추 10kg',
        'suppliers': ['CJ프레시웨이'],
        'filter': '10kg',
    },
    {
        'item': '깻잎(4KG,큰잎)',
        'search': '찹큰',
        'suppliers': ['CJ프레시웨이'],
        'filter': '4kg',
    },
    {
        'item': '케일',
        'search': '케일',
        'suppliers': ['CJ프레시웨이'],
        'filter': '2kg',
    },
    {
        'item': '생표고 수입',
        'search': '표고버섯 2L',
        'suppliers': ['바름푸드', '세현F&B', '세현f&b'],
        'filter': None,
    },
    {
        'item': '깐대파',
        'search': '깐대파 국내산',
        'suppliers': ['다봄푸드'],
        'filter': '1단',
        'exclude': '중국',
    },
    {
        'item': '미나리',
        'search': '미나리 1단',
        'suppliers': ['세현F&B', '세현f&b'],
        'filter': '1단',
    },
    {
        'item': '레몬',
        'search': '레몬 1박스',
        'suppliers': ['다봄푸드', 'CJ프레시웨이'],
        'filter': '17',
    },
    # === 2026-02-25 추가: 기존 누락 품목 ===
    {
        'item': '청경채',
        'search': '청경채',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '4kg',
    },
    {
        'item': '쑥갓',
        'search': '쑥갓',
        'suppliers': ['CJ프레시웨이'],
        'filter': '4kg',
    },
    {
        'item': '깻잎(1KG)',
        'search': '깻잎 50속',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '50속',
        'exclude': '큰',
    },
    {
        'item': '새송이',
        'search': '새송이 2kg',
        'suppliers': ['다봄푸드', 'CJ프레시웨이'],
        'filter': '2kg',
        'exclude': '총알',
    },
    {
        'item': '맛느타리',
        'search': '맛느타리 BOX',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': 'BOX',
    },
    {
        'item': '치커리',
        'search': '치커리',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '2kg',
    },
    {
        'item': '양송이',
        'search': '양송이 2kg',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '2kg',
    },
    {
        'item': '알배기배추',
        'search': '알배기 12입',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': None,
    },
    {
        'item': '깐쪽파 1K',
        'search': '깐쪽파 10kg',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '10kg',
    },
    {
        'item': '꽈리고추',
        'search': '꽈리고추',
        'suppliers': ['CJ프레시웨이'],
        'filter': '4kg',
    },
    {
        'item': '당근 수입',
        'search': '수입당근 10kg',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '10kg',
    },
    {
        'item': '시금치',
        'search': '시금치',
        'suppliers': ['CJ프레시웨이'],
        'filter': '4kg',
    },
    {
        'item': '팽이버섯',
        'search': '팽이버섯',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '34입',
    },
    {
        'item': '흙대파',
        'search': '대파 10kg',
        'suppliers': ['CJ프레시웨이'],
        'filter': '10kg',
    },
    {
        'item': '부추(일반)',
        'search': '부추',
        'suppliers': ['CJ프레시웨이', '세현F&B'],
        'filter': '1단',
    },
    {
        'item': '깐양파',
        'search': '깐양파 1kg',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '1kg',
    },
    {
        'item': '깐쪽파',
        'search': '깐쪽파 1kg',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '1kg',
    },
    {
        'item': '팽이',
        'search': '팽이버섯 34입',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '34입',
    },
    {
        'item': '양배추42',
        'search': '양배추 8kg',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '8kg',
        'exclude': '깐',
    },
    {
        'item': '양배추45',
        'search': '양배추 10kg',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '10kg',
        'exclude': '깐',
    },
    {
        'item': '양배추40',
        'search': '양배추 8kg',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '8kg',
        'exclude': '깐',
    },
    {
        'item': '알배기배추',
        'search': '알배기 8kg',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '8kg',
    },
    {
        'item': '양상추',
        'search': '양상추 4kg',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '4kg',
    },
    {
        'item': '백다다기오이',
        'search': '백다다기 20kg',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '20kg',
    },
    {
        'item': '골드파인애플',
        'search': '골드파인애플',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '12kg',
    },
    {
        'item': '고수',
        'search': '고수',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': None,
        'exclude': '고수익',
    },
    {
        'item': '로메인(일반)',
        'search': '잎로메인',
        'suppliers': ['싱싱채소 그린팜'],
        'filter': '2kg',
    },
    {
        'item': '꽃느타리',
        'search': '꽃느타리',
        'suppliers': ['다봄푸드', '온국민 신선몰'],
        'filter': '2kg',
    },
    {
        'item': '무',
        'search': '무우 BOX',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': None,
    },
    {   # 토마토 완숙 — CJ프레시웨이/다봄푸드/세현F&B 5KG/BOX
        'item': '토마토 완숙',
        'search': '완숙토마토',
        'suppliers': ['CJ프레시웨이', '다봄푸드', '세현F&B'],
        'filter': '5kg',
    },
    {   # 적근대 — CJ프레시웨이 4KG/BOX
        'item': '적근대',
        'search': '적근대',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '4kg',
    },
    {   # 빨간양배추 국산 (적채) — 1통 가격 x6 = 1박스(6통)
        'item': '빨간양배추 국산',
        'search': '적채',
        'suppliers': ['CJ프레시웨이', '다봄푸드'],
        'filter': '국내산',
        'unit_convert': 6,  # 1통 -> 6통(1박스): x6
    },
]


# 품목별 최소 기대 가격 (이 이하면 비정상으로 간주하고 스킵)
# 박스/BOX 단위 거래 품목은 최소 2,000원 이상이어야 정상
MIN_EXPECTED_PRICE = {
    '깻잎(1KG)': 5000,
    '깻잎(4KG,큰잎)': 10000,
    '깐마늘 대서': 5000,
    '감자': 10000,
    '양파': 5000,
    '배추': 5000,
    '무': 5000,
}
# 기본 최소가: unit_convert 적용 전 단가 기준
DEFAULT_MIN_PRICE = 1000


class TargetedScraper(FoodspringScraper):
    """타겟 업체/상품 필터링이 가능한 스크래퍼"""

    def search_targeted(self, keyword, target_suppliers, product_filter=None, exclude_filter=None):
        """특정 업체/상품 필터로 검색"""
        from bs4 import BeautifulSoup

        try:
            url = "https://www.foodspring.co.kr/search"
            if self.page.url.split('?')[0] != url:
                self.page.goto(url, wait_until="networkidle")

            search_input = None
            try:
                loc = self.page.locator("input[type='search']")
                if loc.count() > 0:
                    search_input = loc.first
            except:
                pass
            if not search_input:
                try:
                    loc = self.page.locator("input[type='text']")
                    if loc.count() > 0:
                        search_input = loc.first
                except:
                    pass

            if not search_input:
                print(f"  -> Could not find search input", flush=True)
                return None, None, None

            search_input.fill(keyword)
            search_input.press("Enter")
            self.page.wait_for_timeout(4000)

            soup = BeautifulSoup(self.page.content(), 'html.parser')
            page_text = soup.get_text()

            # 모든 상품 카드/리스트 아이템 찾기
            results = []

            for supplier in target_suppliers:
                elements = soup.find_all(string=lambda t: t and supplier in t)
                for el in elements:
                    # 상위 컨테이너 찾기 (상품 카드)
                    container = el.find_parent(lambda tag: tag.name in ['div', 'li', 'a'])
                    if not container:
                        continue

                    text = container.get_text()

                    # 상품 필터 적용
                    if product_filter and product_filter.lower() not in text.lower():
                        continue

                    # 제외 필터 적용
                    if exclude_filter and exclude_filter in text:
                        continue

                    # 가격 추출
                    all_prices = re.findall(r'([\d,]+)원', text)
                    if not all_prices:
                        continue

                    # 할인가 패턴 (XX% 가격원)
                    discount_match = re.search(r'(\d{1,2})%\s*([\d,]+)원', text)

                    price = None
                    if discount_match:
                        price = int(discount_match.group(2).replace(',', ''))
                    elif len(all_prices) >= 2:
                        # 보통 두번째가 할인/회원가
                        price = int(all_prices[1].replace(',', ''))
                    else:
                        price = int(all_prices[0].replace(',', ''))

                    if price and price > 100:  # 너무 작은 숫자 제외
                        # 상품명 추출 (텍스트에서 처음 몇 줄)
                        product_name = text[:100].replace('\n', ' ').strip()
                        results.append({
                            'supplier': supplier,
                            'price': price,
                            'product': product_name,
                        })

            if results:
                # 업체 우선순위 적용
                for pref_supplier in target_suppliers:
                    matches = [r for r in results if r['supplier'] == pref_supplier]
                    if matches:
                        best = min(matches, key=lambda x: x['price'])  # 최저가 선택
                        return best['price'], best['supplier'], best['product']

                # 아무거나 반환
                best = results[0]
                return best['price'], best['supplier'], best['product']

            return None, None, None

        except Exception as e:
            print(f"  -> Search error: {e}", flush=True)
            return None, None, None


def fill_sikbom_prices(master_file, target_date):
    print(f"Opening master file: {master_file}")
    wb = openpyxl.load_workbook(master_file)
    ws = wb["단가"]

    # Find target rows (use FIRST occurrence only - topmost block is the active one)
    # MULTI_ROW_ITEMS는 등급/단위 복합키로 관리 (grade+unit, grade 등)
    row_map = {}  # {key: row_idx} — 일반 품목은 품목명, 다중행 품목은 복합키
    first_block_ended = set()  # 첫 번째 블록 끝 감지 (중복 블록 방지)
    for r in range(2, ws.max_row + 1):
        dc = ws.cell(row=r, column=2).value
        if not dc:
            continue
        is_match = False
        if str(dc).startswith(target_date):
            is_match = True
        elif hasattr(dc, 'strftime') and dc.strftime("%Y-%m-%d") == target_date:
            is_match = True
        if is_match:
            cat = str(ws.cell(row=r, column=1).value or '')
            if '소분' in cat:
                continue
            item = str(ws.cell(row=r, column=3).value or '').strip()
            if not item:
                continue
            if item in MULTI_ROW_ITEMS:
                mode = MULTI_ROW_ITEMS[item]
                grade_raw = str(ws.cell(row=r, column=4).value or '').strip()
                unit_raw = str(ws.cell(row=r, column=5).value or '').strip().upper()
                if mode == 'grade+unit':
                    compound_key = f"{item}|{grade_raw}|{unit_raw}"
                else:  # 'grade'
                    compound_key = f"{item}|{grade_raw}"
                if compound_key not in row_map:
                    row_map[compound_key] = r
                elif compound_key not in first_block_ended:
                    # 첫 블록 내 연속 행만 허용 (블록 간격 >5 시 중단)
                    if r - row_map[compound_key] > 5:
                        first_block_ended.add(compound_key)
            elif item not in row_map:  # Keep first occurrence only
                row_map[item] = r

    print(f"Found {len(row_map)} 직납/재고 items for {target_date}")
    print(f"Items: {list(row_map.keys())}\n")

    # Initialize scraper
    scraper = TargetedScraper()
    if not scraper.login():
        print("Login failed!")
        scraper.close()
        wb.close()
        return

    updates = 0
    results_log = []

    try:
        last_search = {}  # 동일 검색어 캐시 (중복 검색 방지)

        for cfg in SEARCH_CONFIG:
            item = cfg['item']
            match_grade = cfg.get('match_grade')
            match_unit = cfg.get('match_unit')

            # MULTI_ROW_ITEMS: 복합키로 row_map에서 매칭
            if item in MULTI_ROW_ITEMS:
                found_key = None
                for key in row_map:
                    if not key.startswith(f"{item}|"):
                        continue
                    parts = key.split('|')
                    grade_part = parts[1] if len(parts) > 1 else ''
                    unit_part = parts[2] if len(parts) > 2 else ''
                    if match_grade and match_grade not in grade_part:
                        continue
                    if match_unit and match_unit.upper() != unit_part.upper():
                        continue
                    found_key = key
                    break
                if not found_key:
                    label = f" (grade:{match_grade}, unit:{match_unit})"
                    print(f"[SKIP] {item}{label} not found in danga sheet for {target_date}")
                    continue
                row_idx = row_map[found_key]
            else:
                # 일반 품목: 품목명으로 직접 조회
                if item not in row_map:
                    print(f"[SKIP] {item} not found in danga sheet for {target_date}")
                    continue
                row_idx = row_map[item]

            search_kw = cfg['search']
            suppliers = cfg['suppliers']
            filt = cfg.get('filter')
            exclude = cfg.get('exclude')
            convert = cfg.get('unit_convert', 1)

            grade_label = f"({match_grade})" if match_grade else ""
            unit_label = f" {match_unit}" if match_unit else ""
            display_label = f"{item}{grade_label}{unit_label}"
            print(f"[{display_label}] Searching: '{search_kw}' (filter: {filt}, exclude: {exclude}, suppliers: {suppliers})")

            # 동일 검색어+필터 조합은 캐시 사용 (중복 검색 방지)
            cache_key = (search_kw, tuple(suppliers), filt, exclude)
            if cache_key in last_search:
                price, supplier, product = last_search[cache_key]
            else:
                try:
                    price, supplier, product = scraper.search_targeted(search_kw, suppliers, filt, exclude)
                    last_search[cache_key] = (price, supplier, product)
                except Exception as e:
                    print(f"  -> Search error: {e}")
                    results_log.append(f"{display_label}: ERROR ({e})")
                    continue

            if price:
                # 비정상 가격 검증: 품목별 최소가 이하면 스킵
                min_price = MIN_EXPECTED_PRICE.get(item, DEFAULT_MIN_PRICE)
                if price < min_price:
                    print(f"  -> R{row_idx}: {price:,} / {supplier} — ⚠️ 비정상 가격 (최소 {min_price:,}원 미만), SKIPPED")
                    results_log.append(f"{display_label}: {price:,} REJECTED (< {min_price:,})")
                    continue

                final_price = price * convert
                cell_h = ws.cell(row=row_idx, column=8)
                if not isinstance(cell_h, openpyxl.cell.cell.MergedCell):
                    cell_h.value = final_price

                print(f"  -> R{row_idx}: {price:,} x{convert} = {final_price:,} / {supplier}")
                results_log.append(f"{display_label}: {final_price:,} ({supplier})")
                updates += 1
            else:
                print(f"  -> No price found")
                results_log.append(f"{display_label}: NOT FOUND")

            time.sleep(1)

    finally:
        scraper.close()

        print(f"\n=== Results ===")
        print(f"Updated: {updates}/{len(SEARCH_CONFIG)}")
        for log in results_log:
            print(f"  {log}")

        # === 전파(Propagation): 첫 블록 가격을 모든 중복 행에 복사 ===
        if updates > 0:
            print(f"\n=== Propagating prices to duplicate rows ===")
            # 1) 첫 블록에서 채워진 값 수집
            first_values = {}  # {key: (price, supplier)}
            all_rows = {}      # {key: [row_indices]}

            for r in range(2, ws.max_row + 1):
                dc = ws.cell(row=r, column=2).value
                if not dc:
                    continue
                is_match = False
                if str(dc).startswith(target_date):
                    is_match = True
                elif hasattr(dc, 'strftime') and dc.strftime("%Y-%m-%d") == target_date:
                    is_match = True
                if not is_match:
                    continue

                cat = str(ws.cell(row=r, column=1).value or '')
                if '소분' in cat or '구분' in cat:
                    continue

                item_name = str(ws.cell(row=r, column=3).value or '').strip()
                if not item_name:
                    continue

                grade_raw = str(ws.cell(row=r, column=4).value or '').strip()
                unit_raw = str(ws.cell(row=r, column=5).value or '').strip()

                if item_name in MULTI_ROW_ITEMS:
                    mode = MULTI_ROW_ITEMS[item_name]
                    if mode == 'grade+unit':
                        key = f"{item_name}|{grade_raw}|{unit_raw}"
                    else:
                        key = f"{item_name}|{grade_raw}"
                else:
                    key = item_name

                if key not in all_rows:
                    all_rows[key] = []
                all_rows[key].append(r)

                price_val = ws.cell(row=r, column=8).value
                if key not in first_values and price_val is not None:
                    supplier_val = ws.cell(row=r, column=13).value
                    first_values[key] = (price_val, supplier_val)

            # 2) 중복 행에 전파
            propagated = 0
            for key, rows in all_rows.items():
                if key not in first_values:
                    continue
                price_val, supplier_val = first_values[key]
                for r in rows:
                    cell_h = ws.cell(row=r, column=8)
                    if cell_h.value is None or cell_h.value == '':
                        if not isinstance(cell_h, openpyxl.cell.cell.MergedCell):
                            cell_h.value = price_val
                            if supplier_val:
                                ws.cell(row=r, column=13).value = supplier_val
                            propagated += 1

            print(f"Propagated to {propagated} duplicate rows")

            print(f"\nSaving to {master_file}...")
            for attempt in range(1, 6):
                try:
                    wb.save(master_file)
                    print("Done.")
                    break
                except PermissionError:
                    if attempt < 5:
                        print(f"  File locked, retrying in 3s... (attempt {attempt}/5)")
                        time.sleep(3)
                    else:
                        # 임시 파일에 저장
                        tmp_path = os.path.join(os.environ.get('PUBLIC', 'C:\\Users\\Public'),
                                                'Documents', 'ESTsoft', 'CreatorTemp',
                                                os.path.basename(master_file).replace('.xlsx', '_sikbom.xlsx'))
                        os.makedirs(os.path.dirname(tmp_path), exist_ok=True)
                        wb.save(tmp_path)
                        print(f"  Saved to temp: {tmp_path}")
                        print(f"  Close the master file and copy manually.")
        else:
            print("\nNo updates made.")

    wb.close()


if __name__ == "__main__":
    import sys
    parser = argparse.ArgumentParser()
    parser.add_argument("--master", required=True)
    parser.add_argument("--date", required=True)
    args = parser.parse_args()

    try:
        fill_sikbom_prices(args.master, args.date)
    except SystemExit:
        raise
    except Exception:
        import traceback; traceback.print_exc()
        sys.exit(1)

    from _notify import notify
    notify("stage3_sikbom", date_str=args.date)
