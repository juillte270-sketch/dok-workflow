"""
DoKH 매입단가/판매단가 자동 입력 스크립트

발주시트에 매입단가(I), 판매단가(K), 공급업체(F)를 채우고
수식(J, L, M, N)을 입력한다.

사용법:
    python execution/fill_prices.py --date "2026-02-13" --master-file "path/to/master.xlsx"
"""

import os
import sys
import re
import json
import argparse
from datetime import datetime
from collections import Counter

import openpyxl
from openpyxl.styles import Font, Alignment

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(project_root)

# ==================== 설정 ====================

PRICE_FORMAT = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@'
RATE_FORMAT = '0.0'

# mappings.json에서 공급업체 매핑 로드
MAPPINGS_PATH = os.path.join(
    project_root, 'skills', 'order_processing', 'resources', 'mappings.json'
)
with open(MAPPINGS_PATH, 'r', encoding='utf-8') as f:
    _mappings = json.load(f)
ITEM_SUPPLIER_MAP = _mappings.get('item_supplier_map', {})

# fill_prices가 덮어쓰지 말아야 할 "구체적" 공급업체와 구별되는 generic 라벨
GENERIC_SUPPLIERS = {'직납/재고', '소분/재고', '재고', '동원1배치', '시장', ''}

# 소분/재고 per-unit 단가를 적용하면 안 되는 벌크 단위
BULK_UNITS = {'박스', '망'}

# 발주시트 품목명 → 단가시트 품목명 별칭
ITEM_ALIASES = {
    '표고버섯': '생표고 수입',
    '겉잎제거배추': '배추',
}

# 고정 매입단가 품목 (단가시트와 무관하게 항상 동일한 가격 적용)
FIXED_PRICES = {
    '숙주': 4500,
    '굵은숙주': 4700,
    '새싹(대)': 4000,
    '무순(대)': 800,
    '중란': 5600,
    '대란': 5800,
    '곱슬이콩나물': 2800,
    '일자콩나물': 2800,
    '아보카도': 1900,
    '중숙아보카도': 1900,
    '파김치': 19000,
    '맛김치': 16000,
    '냉동알마늘': 29000,
    '쌀': 59000,
    '매운고춧가루': 9500,
    '굵은고춧가루': 7800,
}


# ==================== 헬퍼 함수 ====================

def lookup_fixed_price(item_name):
    """고정 매입단가 품목 조회. 매칭되면 (가격, 품목명) 반환, 아니면 (0, None)."""
    item_str = item_name.strip()
    # 정확 매칭
    if item_str in FIXED_PRICES:
        return FIXED_PRICES[item_str], item_str
    # 부분 매칭 (긴 키워드 우선)
    for fixed_item in sorted(FIXED_PRICES.keys(), key=len, reverse=True):
        if fixed_item in item_str or item_str in fixed_item:
            return FIXED_PRICES[fixed_item], fixed_item
    return 0, None


def normalize_item(name):
    """품목명에서 수식어/규격을 제거해 퍼지 매칭용 기본 이름 추출"""
    name = re.sub(r'\([^)]*\)', '', name)       # (특,굵은), (4K,큰잎) 등 제거
    name = re.sub(r'[\d]+[A-Za-z]*$', '', name) # 뒤따르는 숫자+영문 제거 (45, 3L 등)
    name = name.replace('머스캣', '머스켓')       # 샤인머스캣↔샤인머스켓
    return name.strip()


def parse_price_unit(unit_str):
    """
    단가시트 단위(E열)를 파싱해 수량과 단위타입을 반환.
    예: "10K" → (10, 'kg'), "1팩" → (1, '팩'), "20K" → (20, 'kg')
        "0.5K" → (0.5, 'kg'), "12봉" → (12, '봉'), "1알" → (1, '알')
    """
    if not unit_str:
        return (1, '')

    unit_str = str(unit_str).strip()

    # "10K", "1K", "20K", "0.5K", "5kg", "1kg" 등 → (숫자, 'kg')
    m = re.match(r'^(\d+(?:\.\d+)?)\s*[Kk](?:[Gg])?$', unit_str)
    if m:
        return (float(m.group(1)), 'kg')

    # "5-6K" 등 범위 → 첫 번째 숫자 사용
    m = re.match(r'^(\d+(?:\.\d+)?)\s*-\s*\d+\s*[Kk](?:[Gg])?$', unit_str)
    if m:
        return (float(m.group(1)), 'kg')

    # "1팩", "1알", "12봉", "1개", "1통" 등 → (숫자, 단위)
    m = re.match(r'^(\d+(?:\.\d+)?)\s*(.+)$', unit_str)
    if m:
        return (float(m.group(1)), m.group(2).strip())

    return (1, unit_str)


def get_store_ranges(ws, target_date):
    """발주시트에서 대상 날짜의 거래처별 행 범위를 파악"""
    ranges = []
    current_store = None
    start_row = None
    last_row = None

    for r in range(2, ws.max_row + 1):
        dc = ws.cell(row=r, column=4).value  # D열: 날짜
        match = False

        if isinstance(dc, datetime):
            if dc.date() == target_date.date():
                match = True
        elif hasattr(dc, 'strftime'):
            try:
                if dc.strftime('%Y-%m-%d') == target_date.strftime('%Y-%m-%d'):
                    match = True
            except:
                pass

        if match:
            store = ws.cell(row=r, column=1).value  # A열: 매장명
            if store != current_store:
                if current_store:
                    ranges.append((current_store, start_row, last_row))
                current_store = store
                start_row = r
            last_row = r

    if current_store:
        ranges.append((current_store, start_row, last_row))

    return ranges


def read_price_sheet(ws_price, target_date):
    """
    단가시트에서 첫 번째 연속 블록의 당일 매입가를 읽어온다.
    중복 블록(같은 날짜 데이터가 여러 번 입력된 경우) 방지를 위해
    첫 번째 연속 블록만 읽고, row gap > 2 이면 중단한다.

    Returns: { item_name: [ {price, supplier, category, unit, bulk_qty, unit_type}, ... ] }
    각 품목에 여러 엔트리(직납/소분)가 있을 수 있으므로 리스트로 저장.
    """
    price_map = {}
    current_supplier = ""
    current_category = ""
    in_target = False
    gap_count = 0
    MAX_GAP = 2

    for r in range(2, ws_price.max_row + 1):
        # 카테고리 (A열) - rolling
        cat_val = ws_price.cell(row=r, column=1).value
        if cat_val:
            current_category = str(cat_val).strip()

        # 거래업체 (J열) - rolling
        supp_val = ws_price.cell(row=r, column=10).value
        if supp_val:
            current_supplier = str(supp_val).strip()

        # 날짜 확인 (B열)
        date_val = ws_price.cell(row=r, column=2).value
        date_match = False

        if isinstance(date_val, datetime):
            if date_val.date() == target_date.date():
                date_match = True
        elif hasattr(date_val, 'date'):
            try:
                if date_val.date() == target_date.date():
                    date_match = True
            except:
                pass

        if date_match:
            in_target = True
            gap_count = 0
        elif in_target:
            gap_count += 1
            if gap_count > MAX_GAP:
                break  # 첫 번째 연속 블록 종료
            continue
        else:
            continue

        # 품목 (C열)
        item_val = ws_price.cell(row=r, column=3).value
        if not item_val:
            continue
        item_name = str(item_val).strip()

        # 매입가 (I열)
        price_val = ws_price.cell(row=r, column=9).value
        unit_val = ws_price.cell(row=r, column=5).value

        price = 0
        if isinstance(price_val, (int, float)):
            price = int(price_val)

        bulk_qty, unit_type = parse_price_unit(unit_val)

        entry = {
            'price': price,
            'supplier': current_supplier,
            'category': current_category,
            'unit': str(unit_val).strip() if unit_val else '',
            'bulk_qty': bulk_qty,
            'unit_type': unit_type,
        }

        if item_name not in price_map:
            price_map[item_name] = []
        price_map[item_name].append(entry)

    return price_map


def lookup_purchase_price(price_map, item_name, order_unit):
    """
    발주시트 품목명으로 단가시트에서 매입가를 찾는다.
    price_map 값이 리스트이므로, 각 항목의 여러 엔트리(직납/소분) 중
    주문 단위에 가장 적합한 것을 선택한다.

    변환 로직:
    - 소분/재고: per-unit 가격 (1K, 1알 등). 벌크(박스/망) 주문에는 미적용.
    - 직납/재고: 벌크 가격 (10K=10kg 단위 등).
      * kg 주문 → price / bulk_qty 로 per-kg 변환
      * 박스/망 주문 → 벌크 가격 그대로 적용
      * kg↔개 단위 미스매치 → 미적용
    """
    # 단위 정규화: k, KG → kg
    unit_lower = order_unit.lower() if order_unit else ''
    is_bulk = order_unit in BULK_UNITS
    order_is_weight = unit_lower in ('kg', 'k')

    def _best_from_entries(entries):
        """
        여러 엔트리에서 주문 단위에 맞는 최적 가격을 찾는다.
        박스/망 주문 → 직납/재고 우선, kg/개 주문 → 소분/재고 우선.
        """
        # 카테고리 우선순위 정렬
        def sort_key(e):
            cat = e.get('category', '')
            if is_bulk:
                return 0 if '직납' in cat else 1
            else:
                return 0 if '소분' in cat else 1

        fallback_candidates = []
        for info in sorted(entries, key=sort_key):
            price = info['price']
            if price <= 0:
                continue

            bulk_qty = info.get('bulk_qty', 1) or 1
            unit_type = info.get('unit_type', '')
            category = info.get('category', '')
            price_is_weight = (unit_type == 'kg')

            if is_bulk:
                if '소분' in category:
                    continue  # 소분 per-unit → 박스에 적용 불가
                return price  # 직납 벌크 가격 그대로

            # kg/개/통/팩 등 소분 단위 주문
            if order_is_weight and price_is_weight:
                return int(price / bulk_qty)
            elif not order_is_weight and not price_is_weight:
                return int(price / bulk_qty)
            elif order_is_weight and not price_is_weight:
                fallback_candidates.append(int(price / bulk_qty))
                continue  # kg 주문에 개/알/통 가격 → 폴백 후보
            elif not order_is_weight and price_is_weight:
                fallback_candidates.append(int(price / bulk_qty))
                continue  # 개 주문에 kg 가격 → 폴백 후보

        # 단위 미스매치지만 유일한 가격이면 폴백 적용
        if fallback_candidates:
            return fallback_candidates[0]
        return 0

    def _has_any_price(entries):
        return any(e['price'] > 0 for e in entries)

    # 조회할 품목명 리스트 (원래 이름 + 별칭)
    lookup_names = [item_name]
    item_norm = normalize_item(item_name)
    for alias_from, alias_to in ITEM_ALIASES.items():
        if alias_from in item_name or alias_from == item_norm:
            lookup_names.append(alias_to)

    for name in lookup_names:
        # 1. 정확 매칭
        if name in price_map:
            result = _best_from_entries(price_map[name])
            if result > 0:
                return result

        # 2. 부분 매칭 (유사도 기반 최적 매칭 — 첫 매칭 반환 X)
        best_result = 0
        best_similarity = -1
        for price_item, entries in price_map.items():
            if not _has_any_price(entries):
                continue
            if name in price_item or price_item in name:
                result = _best_from_entries(entries)
                if result > 0:
                    sim = _name_similarity(name, price_item)
                    if sim > best_similarity:
                        best_similarity = sim
                        best_result = result
        if best_result > 0:
            return best_result

        # 3. 정규화 매칭 (괄호/숫자 제거 후 유사도)
        name_norm = normalize_item(name)
        best_result = 0
        best_similarity = -1
        for price_item, entries in price_map.items():
            if not _has_any_price(entries):
                continue
            price_norm = normalize_item(price_item)
            if name_norm == price_norm or name_norm in price_norm or price_norm in name_norm:
                result = _best_from_entries(entries)
                if result > 0:
                    sim = _name_similarity(name, price_item)
                    if sim > best_similarity:
                        best_similarity = sim
                        best_result = result
        if best_result > 0:
            return best_result

    return 0  # 매칭 실패


def _name_similarity(a, b):
    """두 문자열의 유사도 (0~1). 정규화 매칭 시 최적 매칭 선택에 사용."""
    from difflib import SequenceMatcher
    return SequenceMatcher(None, a, b).ratio()


def determine_supplier(item_name, existing_supplier):
    """
    공급업체를 결정한다.
    1) 기존 F열에 구체적 공급업체명이 있으면 유지
    2) generic 라벨(직납/재고 등)이면 item_supplier_map에서 조회
    """
    existing = str(existing_supplier).strip() if existing_supplier else ''

    # 이미 구체적 공급업체가 있으면 유지
    if existing and existing not in GENERIC_SUPPLIERS:
        return existing

    item_norm = normalize_item(item_name)

    # item_supplier_map에서 매칭
    # 점수: (match_type, keyword_length) 로 정렬해 가장 구체적인 매칭 사용
    candidates = []

    for supplier, keywords in ITEM_SUPPLIER_MAP.items():
        for keyword in keywords:
            kw_norm = normalize_item(keyword)

            if item_norm == kw_norm:
                candidates.append((supplier, 0, len(kw_norm)))  # exact = 최우선
            elif kw_norm in item_norm:
                candidates.append((supplier, 1, len(kw_norm)))  # keyword가 item에 포함
            elif item_norm in kw_norm:
                candidates.append((supplier, 2, len(kw_norm)))  # item이 keyword에 포함

    if candidates:
        # match_type 오름차순, keyword_length 내림차순 (더 구체적 = 긴 키워드 우선)
        candidates.sort(key=lambda c: (c[1], -c[2]))
        return candidates[0][0]

    return existing  # 매칭 실패 시 기존값 유지


def get_historical_prices(ws, target_date, scan_limit=5000):
    """
    발주시트에서 과거 거래 이력을 읽어 마진율 참고 데이터를 만든다.
    매입가(I열)와 판매가(K열) 모두 읽어 마진율을 계산한다.
    Returns: {
        (store_group, item_keyword): [
            {'purchase': int, 'sale': int, 'margin_rate': float, 'date': date}, ...
        ]
    }
    """
    history = {}

    max_row = ws.max_row
    start_scan = max(2, max_row - scan_limit)

    for r in range(start_scan, max_row + 1):
        date_val = ws.cell(row=r, column=4).value  # D열
        if not date_val:
            continue

        try:
            if isinstance(date_val, datetime):
                row_date = date_val.date()
            elif hasattr(date_val, 'date'):
                row_date = date_val.date()
            else:
                continue

            if row_date >= target_date.date():
                continue
        except:
            continue

        store = ws.cell(row=r, column=1).value  # A열
        item = ws.cell(row=r, column=5).value    # E열
        purchase_val = ws.cell(row=r, column=9).value   # I열: 매입단가
        sale_val = ws.cell(row=r, column=11).value      # K열: 판매단가

        if not store or not item:
            continue

        # 판매단가 검증
        if not sale_val or (isinstance(sale_val, (int, float)) and sale_val == 0):
            continue
        if isinstance(sale_val, str) and sale_val.startswith('='):
            continue

        store_str = str(store).strip()
        item_str = str(item).strip()
        sale_int = int(sale_val) if isinstance(sale_val, (int, float)) else 0
        purchase_int = int(purchase_val) if isinstance(purchase_val, (int, float)) else 0

        if sale_int <= 0:
            continue

        # 마진율 계산
        if purchase_int > 0:
            margin_rate = (sale_int - purchase_int) / sale_int
        else:
            margin_rate = None  # 매입가 없으면 마진율 계산 불가

        # 매장 그룹 결정
        is_sukju = ('숙주' in item_str)
        if '샤브야키' in store_str:
            if is_sukju:
                group = store_str
            else:
                group = 'shabu'
        elif '부엉이산장' in store_str:
            group = 'owl'
        else:
            group = store_str

        key = (group, item_str)
        if key not in history:
            history[key] = []
        history[key].append({
            'purchase': purchase_int,
            'sale': sale_int,
            'margin_rate': margin_rate,
            'date': row_date,
        })

    return history


def determine_sales_price(store_name, item_name, purchase_price, history):
    """
    판매단가 결정:
    1. 최근 5거래일 판매가가 모두 동일 → 고정가 유지 (매장별 고정 판매단가)
    2. 변동 → 최근 2거래일 마진율 평균으로 계산
    3. 이력 없음 → 기본 20% 마진
    """
    store_str = str(store_name).strip()
    item_str = str(item_name).strip()
    DEFAULT_MARGIN = 0.20

    is_sukju = ('숙주' in item_str)

    if '샤브야키' in store_str:
        group = store_str if is_sukju else 'shabu'
    elif '부엉이산장' in store_str:
        group = 'owl'
    else:
        group = store_str

    key = (group, item_name)
    records = history.get(key, [])

    if records:
        # 날짜 내림차순 정렬
        records_sorted = sorted(records, key=lambda x: x['date'], reverse=True)

        # 최근 5 거래일 추출
        recent_dates = []
        for rec in records_sorted:
            if rec['date'] not in recent_dates:
                recent_dates.append(rec['date'])
            if len(recent_dates) >= 5:
                break

        recent_records = [r for r in records_sorted if r['date'] in recent_dates]

        # 고정 판매가 감지: 최근 거래(3건 이상)의 판매가가 모두 동일하면 고정가
        recent_sales = [r['sale'] for r in recent_records if r['sale'] > 0]
        if len(recent_sales) >= 3 and len(set(recent_sales)) == 1:
            # 고정 판매가 → 매입가 변동과 무관하게 유지
            return recent_sales[0]

        # 변동 판매가 → 마진율 기반 계산
        recent_margins = [
            r['margin_rate'] for r in recent_records
            if r['date'] in recent_dates[:2] and r['margin_rate'] is not None
        ]

        if recent_margins and purchase_price and purchase_price > 0:
            avg_margin = sum(recent_margins) / len(recent_margins)
            # 마진율 합리적 범위 제한 (5% ~ 50%)
            avg_margin = max(0.05, min(0.50, avg_margin))
            new_price = int(round(purchase_price / (1 - avg_margin), -2))
            return new_price

        # 마진율 데이터 없으면 최근 판매가 그대로 (폴백)
        if records_sorted[0]['sale'] > 0:
            return records_sorted[0]['sale']

    # 이력 없음 → 기본 마진율
    if purchase_price and purchase_price > 0:
        return int(round(purchase_price / (1 - DEFAULT_MARGIN), -2))

    return 0


def validate_margins(ws, store_ranges, history):
    """마진율 검증 및 경고"""
    warnings = []

    for store_name, start_row, end_row in store_ranges:
        for r in range(start_row, end_row + 1):
            item = ws.cell(row=r, column=5).value
            purchase = ws.cell(row=r, column=9).value  # I열
            sale = ws.cell(row=r, column=11).value      # K열

            if not purchase or not sale:
                continue
            if not isinstance(purchase, (int, float)) or not isinstance(sale, (int, float)):
                continue
            if sale == 0:
                continue

            margin = (sale - purchase) / sale * 100
            item_str = str(item).strip() if item else ''

            if margin < 0:
                new_price = int(round(purchase / 0.80, -2))
                ws.cell(row=r, column=11).value = new_price
                warnings.append(
                    f"  [재계산] {store_name} / {item_str}: "
                    f"마진 {margin:.1f}% -> 20% 재계산 ({sale} -> {new_price})"
                )
            elif margin < 10:
                warnings.append(
                    f"  [저마진] {store_name} / {item_str}: "
                    f"마진 {margin:.1f}% (매입:{purchase}, 판매:{sale})"
                )
            elif margin > 70:
                if '부엉이산장' in str(store_name):
                    warnings.append(
                        f"  [고마진-유지] {store_name} / {item_str}: "
                        f"마진 {margin:.1f}% (부엉이산장 기존가 유지)"
                    )
                else:
                    warnings.append(
                        f"  [고마진] {store_name} / {item_str}: "
                        f"마진 {margin:.1f}% (매입:{purchase}, 판매:{sale})"
                    )

    return warnings


# ==================== 메인 ====================

def fill_prices(master_file, target_date_str):
    target_date = datetime.strptime(target_date_str, "%Y-%m-%d")

    print(f"=== 매입/판매단가 입력 ===")
    print(f"대상 날짜: {target_date_str}")
    print(f"마스터 파일: {master_file}")

    wb = openpyxl.load_workbook(master_file)
    ws_order = wb["발주"]
    ws_price = wb["단가"]

    # 1. 발주시트에서 대상 날짜 행 범위 파악
    print("\n[1/5] 대상 날짜 행 범위 파악...")
    store_ranges = get_store_ranges(ws_order, target_date)

    if not store_ranges:
        print(f"  {target_date_str} 날짜의 발주 데이터가 없습니다.")
        return

    total_rows = sum(end - start + 1 for _, start, end in store_ranges)
    print(f"  {len(store_ranges)}개 거래처, {total_rows}개 항목")
    for store, start, end in store_ranges:
        print(f"    {store}: 행 {start}~{end} ({end - start + 1}건)")

    # 2. 단가시트에서 매입가 조회
    print("\n[2/5] 단가시트 매입가 조회...")
    price_map = read_price_sheet(ws_price, target_date)
    price_items_with_value = [k for k, entries in price_map.items()
                              if any(e['price'] > 0 for e in entries)]
    total_entries = sum(len(v) for v in price_map.values())
    print(f"  단가시트에서 {len(price_map)}개 품목({total_entries}엔트리) 로드 (매입가 있음: {len(price_items_with_value)}개)")
    for item_name in price_items_with_value:
        for info in price_map[item_name]:
            if info['price'] <= 0:
                continue
            bulk_qty = info.get('bulk_qty', 1) or 1
            per_unit = int(info['price'] / bulk_qty) if bulk_qty > 0 else info['price']
            unit_label = f"{info['price']}원/{info['unit']}"
            if bulk_qty > 1:
                unit_label += f" (per-kg: {per_unit}원)"
            print(f"    [{info['category']}] {item_name}: {unit_label} (거래: {info['supplier']})")

    # 3. 과거 거래 이력 조회 (마진율 참고)
    print("\n[3/5] 과거 거래 이력 조회 (전거래일/전전거래일 마진율 참고)...")
    history = get_historical_prices(ws_order, target_date)
    print(f"  {len(history)}개 (매장그룹, 품목) 이력 로드")

    # 4. 값 & 수식 입력
    print("\n[4/5] 매입단가 / 판매단가 / 공급업체 / 수식 입력...")

    font_calibri = Font(name='Calibri', sz=11)
    font_malgun = Font(name='Malgun Gothic', sz=11)
    align_center = Alignment(horizontal='center', vertical='center')

    filled_count = 0
    purchase_filled = 0
    fixed_filled = 0
    sales_filled = 0
    supplier_changed = 0
    no_purchase_price = []
    no_sales_price = []

    for store_name, start_row, end_row in store_ranges:
        # N열: 이익률 (거래처 첫 행)
        profit_formula = (
            f"=IF(SUM(L{start_row}:L{end_row})>0,"
            f"(SUM(L{start_row}:L{end_row})-SUM(J{start_row}:J{end_row}))"
            f"/SUM(L{start_row}:L{end_row})*100,0)"
        )
        cell_n = ws_order.cell(row=start_row, column=14, value=profit_formula)
        cell_n.font = font_calibri
        cell_n.alignment = align_center
        cell_n.number_format = RATE_FORMAT

        for r in range(start_row, end_row + 1):
            item_name = ws_order.cell(row=r, column=5).value  # E열
            unit = ws_order.cell(row=r, column=7).value        # G열

            if not item_name:
                continue

            item_str = str(item_name).strip()
            unit_str = str(unit).strip() if unit else ''

            # --- 매입단가 ---
            # 1) 고정 매입단가 확인
            fixed_price, fixed_match = lookup_fixed_price(item_str)
            if fixed_price > 0:
                purchase_price = fixed_price
                fixed_filled += 1
            else:
                # 2) 단가시트에서 조회
                purchase_price = lookup_purchase_price(price_map, item_str, unit_str)

            if purchase_price > 0:
                purchase_filled += 1
            else:
                no_purchase_price.append(f"  {store_name} / {item_str} ({unit_str})")

            # --- 판매단가 ---
            sales_price = determine_sales_price(store_name, item_str, purchase_price, history)
            if sales_price > 0:
                sales_filled += 1
            else:
                no_sales_price.append(f"  {store_name} / {item_str}")

            # --- 공급업체 ---
            existing_supplier = ws_order.cell(row=r, column=6).value
            new_supplier = determine_supplier(item_str, existing_supplier)
            old_val = str(existing_supplier).strip() if existing_supplier else ''
            if new_supplier and new_supplier != old_val:
                supplier_changed += 1

            # I열: 매입단가
            # 기존에 이미 값이 있으면 보존 (수동 입력값 보호)
            cell_i = ws_order.cell(row=r, column=9)
            existing_buy = cell_i.value
            has_existing_buy = (isinstance(existing_buy, (int, float)) and existing_buy > 0)

            if purchase_price > 0:
                cell_i.value = purchase_price
            elif not has_existing_buy:
                cell_i.value = None
            # else: 기존 수동 입력값 유지
            cell_i.font = font_calibri
            cell_i.alignment = align_center
            cell_i.number_format = PRICE_FORMAT

            # K열: 판매단가
            cell_k = ws_order.cell(row=r, column=11)
            existing_sell = cell_k.value
            has_existing_sell = (isinstance(existing_sell, (int, float)) and existing_sell > 0)

            if sales_price > 0:
                cell_k.value = sales_price
            elif not has_existing_sell:
                cell_k.value = None
            # else: 기존 수동 입력값 유지
            cell_k.font = font_calibri
            cell_k.alignment = align_center
            cell_k.number_format = PRICE_FORMAT

            # F열: 공급업체
            if new_supplier:
                cell_f = ws_order.cell(row=r, column=6, value=new_supplier)
                cell_f.font = font_malgun
                cell_f.alignment = align_center

            # J열: 총매입금액 (수식)
            cell_j = ws_order.cell(row=r, column=10, value=f'=H{r}*I{r}')
            cell_j.font = font_calibri
            cell_j.alignment = align_center
            cell_j.number_format = PRICE_FORMAT

            # L열: 총판매금액 (수식)
            cell_l = ws_order.cell(row=r, column=12, value=f'=K{r}*H{r}')
            cell_l.font = font_calibri
            cell_l.alignment = align_center
            cell_l.number_format = PRICE_FORMAT

            # M열: 마진율 (수식)
            cell_m = ws_order.cell(row=r, column=13, value=f'=IF(L{r}>0,(L{r}-J{r})/L{r}*100,0)')
            cell_m.font = font_calibri
            cell_m.alignment = align_center
            cell_m.number_format = RATE_FORMAT

            filled_count += 1

    print(f"  총 {filled_count}건 처리")
    print(f"  매입단가 입력: {purchase_filled}건 / {filled_count}건 (고정단가: {fixed_filled}건)")
    print(f"  판매단가 입력: {sales_filled}건 / {filled_count}건")
    print(f"  공급업체 변경: {supplier_changed}건")

    if no_purchase_price:
        print(f"\n  매입단가 미확인 ({len(no_purchase_price)}건):")
        for line in no_purchase_price:
            print(line)

    if no_sales_price:
        print(f"\n  판매단가 미확인 ({len(no_sales_price)}건):")
        for line in no_sales_price:
            print(line)

    # 5. 검증
    print("\n[5/5] 마진율 검증...")
    warnings = validate_margins(ws_order, store_ranges, history)

    if warnings:
        print(f"\n  검증 결과 ({len(warnings)}건):")
        for w in warnings:
            print(w)
    else:
        print("  검증 통과")

    # 저장
    print(f"\n저장 중: {master_file}")
    wb.save(master_file)
    print("완료!")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="DoKH 매입/판매단가 자동 입력")
    parser.add_argument("--date", required=True, help="대상 날짜 (YYYY-MM-DD)")
    parser.add_argument("--master-file", required=True, help="마스터 엑셀 파일 경로")
    args = parser.parse_args()

    if not os.path.exists(args.master_file):
        print(f"마스터 파일을 찾을 수 없습니다: {args.master_file}")
        sys.exit(1)

    try:
        fill_prices(args.master_file, args.date)
    except SystemExit:
        raise
    except Exception:
        import traceback; traceback.print_exc()
        sys.exit(1)
    from _notify import notify
    notify("stage7_fill", date_str=args.date)
