"""
도크 비즈니스 분석 리포트 생성 -> 카카오톡 전송

사용법:
    # 일일 리포트
    python execution/generate_business_report.py --type daily --date "2026-02-13"

    # 주간 리포트 (2월 2주차)
    python execution/generate_business_report.py --type weekly --month 2026-02 --week 2

    # 월간 리포트
    python execution/generate_business_report.py --type monthly --month 2026-01

    # 카카오톡 전송 없이 미리보기만
    python execution/generate_business_report.py --type daily --date "2026-02-13" --preview

    # 친구에게 전송 (사전 friends scope 인증 필요)
    python execution/generate_business_report.py --type weekly --month 2026-02 --week 2 --to "친구UUID"

    # 카카오톡 친구 목록 조회
    python execution/generate_business_report.py --list-friends

    # 단톡방 전송 (카카오톡 PC앱에서 채팅방 열어둔 상태)
    python execution/generate_business_report.py --type weekly --month 2026-02 --week 2 --group "직납방"
"""
import os
import sys
import json
import time
import argparse
from datetime import datetime, timedelta
from collections import defaultdict

import openpyxl
import requests
from dotenv import load_dotenv

load_dotenv()

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 단위 값 목록 (공급업체 컬럼에 잘못 들어간 값 필터용)
UNIT_VALUES = {'박스', 'kg', 'KG', 'g', '망', '통', '팩', '개', '포', '판', '발', '단', '봉', '속', '근'}

# 과일류 키워드 (저마진 분석에서 별도 처리 - 구조적 저마진)
FRUIT_KEYWORDS = ['사과', '딸기', '블루베리', '망고', '오렌지', '레몬', '감귤', '한라봉',
                  '바나나', '키위', '멜론', '수박', '참외', '포도', '복숭아', '자두',
                  '체리', '파인애플', '아보카도', '라임', '자몽', '귤', '파파야',
                  '레드향', '천혜향', '황금향', '용과']

# 고정가격 품목 (공산품 - 매입단가 변동 분석에서 제외)
FIXED_PRICE_ITEMS = ['쌀']

# 품명 변경 매핑 (구명칭 → 현재 명칭으로 통합)
ITEM_ALIASES = {
    '꽃느타리': '맛느타리',
    '느타리': '맛느타리',
    '감자(왕왕)': '감자',
    '감자(왕특)': '감자',
}

# 품목별 알려진 컨텍스트 (인사이트 주석용)
ITEM_CONTEXT = {
    '모닝글로리': '공급업체변경',
}

# 설날 날짜 (연도별)
LUNAR_NEW_YEAR = {
    2026: datetime(2026, 2, 17),
    2027: datetime(2027, 2, 6),
}


def _is_fruit(name):
    """과일류 여부 판단"""
    for kw in FRUIT_KEYWORDS:
        if kw in str(name):
            return True
    return False


def _is_fixed_price(name):
    """고정가격 공산품 여부"""
    return str(name) in FIXED_PRICE_ITEMS


def _normalize_item(name):
    """품명 변경 매핑 적용 (구명칭 → 현재 명칭)"""
    return ITEM_ALIASES.get(str(name), name)


def _get_context(name):
    """품목별 알려진 컨텍스트 반환"""
    return ITEM_CONTEXT.get(str(name), None)


def _is_near_holiday(date_obj, year):
    """설날 등 명절 근접 여부 (전후 7일)"""
    if year in LUNAR_NEW_YEAR:
        holiday = LUNAR_NEW_YEAR[year].date() if hasattr(LUNAR_NEW_YEAR[year], 'date') else LUNAR_NEW_YEAR[year]
        if hasattr(date_obj, 'date'):
            date_obj = date_obj.date()
        diff = (holiday - date_obj).days
        if 0 <= diff <= 10:
            return '설날'
    return None


def _load_supplier_map():
    """mappings.json에서 품목->공급업체 매핑 로드"""
    mapping_path = os.path.join(PROJECT_ROOT, 'skills', 'order_processing', 'resources', 'mappings.json')
    item_to_supplier = {}
    if os.path.exists(mapping_path):
        with open(mapping_path, 'r', encoding='utf-8') as f:
            mappings = json.load(f)
        for supplier, items in mappings.get('item_supplier_map', {}).items():
            for item in items:
                item_to_supplier[item] = supplier
    return item_to_supplier


# 발주 시트 컬럼 인덱스 (1-based)
COL_STORE = 1       # 매장명
COL_YEAR = 2        # 기준년도
COL_MONTH = 3       # 기준월
COL_DATE = 4        # 기준일
COL_ITEM = 5        # 품목
COL_SUPPLIER = 6    # 공급업체
COL_UNIT = 7        # 단위
COL_QTY = 8         # 입고수량
COL_BUY_PRICE = 9   # 매입단가
COL_BUY_TOTAL = 10  # 총 매입금액
COL_SELL_PRICE = 11 # 판매단가
COL_SELL_TOTAL = 12 # 총 판매금액
COL_MARGIN = 13     # 마진율
COL_PROFIT = 14     # 이익률


def load_master_data(master_path=None):
    """마스터 파일에서 발주 데이터 로드"""
    if not master_path:
        master_path = os.path.join(PROJECT_ROOT, 'drive_master.xlsx')
        if not os.path.exists(master_path):
            gdrive = r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\도크발주관리데이터.xlsx"
            if os.path.exists(gdrive):
                master_path = gdrive
            else:
                print("Error: Master file not found. Download first.")
                return []

    print(f"Loading master data from: {os.path.basename(master_path)}...")
    wb = openpyxl.load_workbook(master_path, data_only=True)

    if '발주' not in wb.sheetnames:
        print("Error: '발주' sheet not found")
        return []

    ws = wb['발주']
    records = []
    supplier_map = _load_supplier_map()

    for row_idx in range(2, ws.max_row + 1):
        store = ws.cell(row=row_idx, column=COL_STORE).value
        date_val = ws.cell(row=row_idx, column=COL_DATE).value
        item = ws.cell(row=row_idx, column=COL_ITEM).value

        if not item:
            continue

        if isinstance(date_val, datetime):
            dt = date_val
        elif isinstance(date_val, str):
            try:
                dt = datetime.strptime(date_val[:10], "%Y-%m-%d")
            except ValueError:
                continue
        else:
            continue

        qty = _to_float(ws.cell(row=row_idx, column=COL_QTY).value)
        buy_price = _to_float(ws.cell(row=row_idx, column=COL_BUY_PRICE).value)
        buy_total = _to_float(ws.cell(row=row_idx, column=COL_BUY_TOTAL).value)
        sell_price = _to_float(ws.cell(row=row_idx, column=COL_SELL_PRICE).value)
        sell_total = _to_float(ws.cell(row=row_idx, column=COL_SELL_TOTAL).value)
        margin = _to_float(ws.cell(row=row_idx, column=COL_MARGIN).value)

        if not buy_total and qty and buy_price:
            buy_total = qty * buy_price
        if not sell_total and qty and sell_price:
            sell_total = qty * sell_price

        raw_supplier = ws.cell(row=row_idx, column=COL_SUPPLIER).value or ''
        if not raw_supplier or str(raw_supplier) in UNIT_VALUES:
            supplier = ''
            item_str = str(item)
            for mapped_item, mapped_sup in supplier_map.items():
                if mapped_item in item_str or item_str in mapped_item:
                    supplier = mapped_sup
                    break
        else:
            supplier = raw_supplier

        records.append({
            'store': store or '(미지정)',
            'date': dt,
            'item': item,
            'supplier': supplier,
            'unit': ws.cell(row=row_idx, column=COL_UNIT).value or '',
            'qty': qty,
            'buy_price': buy_price,
            'buy_total': buy_total,
            'sell_price': sell_price,
            'sell_total': sell_total,
            'margin': margin,
        })

    wb.close()
    print(f"Loaded {len(records)} records.")
    return records


def _to_float(val):
    if val is None:
        return 0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0


def _fmt_money(val):
    if abs(val) >= 10000:
        man = val / 10000
        if abs(man) >= 100:
            return f"{man:,.0f}만원"
        return f"{man:,.1f}만원"
    return f"{val:,.0f}원"


def _s(val):
    """짧은 금액 (만원 단위 표현, 1.5만 등 소수점 포함)"""
    if abs(val) < 10000:
        return f"{val:,.0f}원"
    man = val / 10000
    if man == int(man):
        return f"{int(man)}만"
    return f"{man:.1f}만"


def _week_number(day):
    return (day - 1) // 7 + 1


# ─── 공통 포맷팅 ─────────────────────────────────────────────

def _shorten(name):
    return name.replace("오레노이키루미치", "오레노").replace("선데이버거클럽", "선데이BC")


def _item_margin(d):
    if d['sell'] > 0:
        return (d['sell'] - d['buy']) / d['sell'] * 100
    return 0


def _change_arrow(pct):
    if pct > 5:
        return f"▲{pct:+.1f}%"
    elif pct < -5:
        return f"▼{pct:+.1f}%"
    elif pct > 0:
        return f"△{pct:+.1f}%"
    elif pct < 0:
        return f"▽{pct:+.1f}%"
    return "━ 0%"


def _bar(pct, width=8):
    filled = max(0, min(width, round(pct / 100 * width)))
    return '■' * filled + '□' * (width - filled)


# ─── 일일 리포트 ──────────────────────────────────────────────

def _is_no_sell_store(store_name):
    """판매단가 미입력 매장 (매입전용 매장) 판별"""
    no_sell = ['동원']
    return any(k in store_name for k in no_sell)


def daily_report(records, date_str):
    """일일 분석 리포트 (주간 리포트와 동일 간결 포맷)"""
    target = datetime.strptime(date_str, "%Y-%m-%d")
    day_data = [r for r in records if r['date'].date() == target.date()]

    if not day_data:
        return [f"📊 도크 일일분석\n{date_str}\n\n해당 날짜에 데이터가 없습니다."]

    # 비교 대상: 가장 최근의 금요일 (토요일 긴급입고 등 제외)
    # 현재 날짜 기준 과거 데이터 중 금요일인 날짜 찾기
    prev_dates = sorted(set(r['date'].date() for r in records if r['date'].date() < target.date()), reverse=True)
    comp_date = None
    for d in prev_dates:
        if d.weekday() == 4: # 금요일
            comp_date = d
            break
    
    # 금요일을 못 찾으면 직전 영업일 사용
    if not comp_date and prev_dates:
        comp_date = prev_dates[0]
    
    prev_data = [r for r in records if r['date'].date() == comp_date] if comp_date else []

    # 전체 합계 (동원 제외 — 매입전용 매장은 마진 계산에서 제외)
    total_buy = sum(r['buy_total'] for r in day_data if not _is_no_sell_store(r['store']))
    total_sell = sum(r['sell_total'] for r in day_data if not _is_no_sell_store(r['store']))
    total_profit = total_sell - total_buy
    avg_margin = (total_profit / total_sell * 100) if total_sell > 0 else 0

    # 매장별 집계 (동원 제외 — 판매단가 미입력 매장)
    stores = defaultdict(lambda: {'buy': 0, 'sell': 0, 'items': 0})
    for r in day_data:
        if _is_no_sell_store(r['store']):
            continue
        stores[r['store']]['buy'] += r['buy_total']
        stores[r['store']]['sell'] += r['sell_total']
        stores[r['store']]['items'] += 1

    # 품목별 집계 (동원 제외)
    items = defaultdict(lambda: {'qty': 0, 'buy': 0, 'sell': 0, 'stores': set(), 'buy_price': 0, 'sell_prices': []})
    for r in day_data:
        if _is_no_sell_store(r['store']):
            continue
        nm = _normalize_item(r['item'])
        items[nm]['qty'] += r['qty']
        items[nm]['buy'] += r['buy_total']
        items[nm]['sell'] += r['sell_total']
        items[nm]['stores'].add(r['store'])
        # 가중평균 매입단가 계산용
        items[nm]['buy_price'] = items[nm]['buy'] / items[nm]['qty'] if items[nm]['qty'] > 0 else 0
        
        # 판매단가가 있는 경우 리스트에 추가 (단순 평균용)
        if r['sell_price'] > 0:
            items[nm]['sell_prices'].append(r['sell_price'])

    # 공급업체별 집계 (동원 포함 — 매입금액에는 포함)
    suppliers = defaultdict(lambda: {'count': 0, 'buy': 0})
    for r in day_data:
        if r['supplier']:
            suppliers[r['supplier']]['count'] += 1
            suppliers[r['supplier']]['buy'] += r['buy_total']

    weekdays = ['월', '화', '수', '목', '금', '토', '일']
    wd = weekdays[target.weekday()]
    comp_label = "금요일 대비" if comp_date and comp_date.weekday() == 4 else "전일비"

    sections = []

    # ── Section 1: Header + 핵심 지표 ──
    L = []
    L.append("━━━━━━━━━━━━━━━━")
    L.append("📊 도크 일일분석 리포트")
    L.append(f"📅 {target.month}/{target.day}({wd})")
    L.append("━━━━━━━━━━━━━━━━")
    L.append("")
    L.append("💰 핵심 지표")
    L.append(f"  매출  {_fmt_money(total_sell)}")
    L.append(f"  매입  {_fmt_money(total_buy)}")
    L.append(f"  이익  {_fmt_money(total_profit)}")
    L.append(f"  마진  {avg_margin:.1f}% {_bar(avg_margin)}")
    L.append(f"  매장  {len(stores)}개 · 품목 {len(items)}개")

    if prev_data:
        prev_sell = sum(r['sell_total'] for r in prev_data)
        prev_buy = sum(r['buy_total'] for r in prev_data)
        prev_margin = ((prev_sell - prev_buy) / prev_sell * 100) if prev_sell > 0 else 0
        if prev_sell > 0:
            L.append(f"  {comp_label} {_change_arrow((total_sell - prev_sell) / prev_sell * 100)}")
            margin_diff = avg_margin - prev_margin
            if abs(margin_diff) >= 0.5:
                L.append(f"  마진변동 {margin_diff:+.1f}%p")
    sections.append("\n".join(L))

    # ── Section 2: 매장별 매출 (간결 포맷) ──
    store_list = sorted(stores.items(), key=lambda x: x[1]['sell'], reverse=True)
    L = []
    L.append(f"🏪 일일매출 ({len(store_list)}매장)")
    for i, (sname, sd) in enumerate(store_list, 1):
        short = _shorten(sname)
        margin = ((sd['sell'] - sd['buy']) / sd['sell'] * 100) if sd['sell'] > 0 else 0
        profit = sd['sell'] - sd['buy']
        if sd['sell'] > 0:
            L.append(f"")
            L.append(f"{i}.{short}/{_s(sd['sell'])} {margin:.0f}%({_s(profit)}),{sd['items']}품목")
        else:
            L.append(f"")
            L.append(f"{i}.{short}/{sd['items']}품목")
    sections.append("\n".join(L))

    # ── Section 3: 매입 TOP (간결 포맷) ──
    item_list = sorted(items.items(), key=lambda x: x[1]['buy'], reverse=True)
    L = []
    L.append("📦 일일매입 TOP10")
    for i, (name, d) in enumerate(item_list[:10], 1):
        margin = _item_margin(d)
        profit = d['sell'] - d['buy']
        L.append(f"")
        if d['sell'] > 0:
            L.append(f"{i}.{name}/{_s(d['buy'])} {margin:.0f}%({_s(profit)}),{len(d['stores'])}매장")
        else:
            L.append(f"{i}.{name}/{_s(d['buy'])},{len(d['stores'])}매장")
    sections.append("\n".join(L))

    # ── Section 4: 고마진/저마진 + 공급업체 (간결 포맷) ──
    # 비교용 금요일 데이터 품목별 집계
    prev_items_map = defaultdict(lambda: {'qty': 0, 'buy': 0, 'sell': 0})
    for r in prev_data:
        nm = _normalize_item(r['item'])
        prev_items_map[nm]['qty'] += r['qty']
        prev_items_map[nm]['buy'] += r['buy_total']
        prev_items_map[nm]['sell'] += r['sell_total']

    MIN_BUY = 30000
    margin_items = [(name, d, _item_margin(d)) for name, d in items.items()
                    if d['sell'] > 0 and d['buy'] >= MIN_BUY]
    margin_items.sort(key=lambda x: x[2], reverse=True)

    L = []
    high_margin = [x for x in margin_items if x[2] >= 15 and not _is_fruit(x[0])]
    if high_margin:
        L.append("💎 고마진 품목 TOP")
        for i, (name, d, m) in enumerate(high_margin[:6], 1):
            profit = d['sell'] - d['buy']
            buy_p = d['buy_price']
            
            # 판매가는 실질 판매 데이터가 있는 경우 그들의 평균 사용 ( diluition 방지 )
            if d['sell_prices']:
                sell_p = sum(d['sell_prices']) / len(d['sell_prices'])
            else:
                sell_p = d['sell'] / d['qty'] if d['qty'] > 0 else 0
            
            # 경매가(전일매입가 대용) 분석
            auction_info = ""
            if name in prev_items_map and prev_items_map[name]['qty'] > 0:
                prev_p = prev_items_map[name]['buy'] / prev_items_map[name]['qty']
                if buy_p < prev_p * 0.95:
                    auction_info = f"\n  ▸ {comp_label} 매입가 하락으로 마진 개선"
            
            L.append(f"")
            L.append(f"")
            L.append(f"{i}.{name}/{m:.0f}%({_s(profit)}),매입{_s(d['buy'])}")
            L.append(f"  [매입 {buy_p:,.0f}원 / 판매 {sell_p:,.0f}원]{auction_info}")

    low_margin_veg = [x for x in margin_items if x[2] < 15 and not _is_fruit(x[0]) and not _is_fixed_price(x[0])]
    low_margin_veg.sort(key=lambda x: x[2])
    if low_margin_veg:
        L.append(f"")
        L.append("⚠️ 저마진 주의")
        for name, d, m in low_margin_veg[:5]:
            profit = d['sell'] - d['buy']
            note = ""
            if "깻잎" in name:
                note = " (1.5kg 컴플레인 반품 반영)"
            L.append(f"  ▸ {name}/{m:.0f}%({_s(profit)}),매입{_s(d['buy'])}{note}")

    fruit_items = [x for x in margin_items if _is_fruit(x[0])]
    if fruit_items:
        fruit_margins = [m for _, _, m in fruit_items]
        avg_fruit_m = sum(fruit_margins) / len(fruit_margins)
        fruit_buy = sum(d['buy'] for _, d, _ in fruit_items)
        L.append(f"")
        L.append(f"🍎 과일류 {len(fruit_items)}건 평균{avg_fruit_m:.0f}% 매입{_s(fruit_buy)}")
        L.append(f"   (구조적 저마진 4~7% 유지중)")

    if sup_list := sorted(suppliers.items(), key=lambda x: x[1]['buy'], reverse=True):
        L.append(f"")
        L.append("🚚 공급업체 매입 (전체)")
        for sname, sd in sup_list: # 전체 업체 출력 (사용자 요청)
            pct = (sd['buy'] / total_buy * 100) if total_buy > 0 else 0
            L.append(f"  {sname}: {_s(sd['buy'])} ({pct:.0f}%)")

    if L:
        sections.append("\n".join(L))

    # ── Section 5: 일일 분석 + 인사이트 엔진 (하나로 통합) ──
    prev_items = defaultdict(lambda: {'qty': 0, 'buy': 0, 'sell': 0})
    prev_stores = defaultdict(lambda: {'sell': 0})

    if prev_data:
        for r in prev_data:
            if _is_no_sell_store(r['store']):
                continue
            prev_items[r['item']]['qty'] += r['qty']
            prev_items[r['item']]['buy'] += r['buy_total']
            prev_items[r['item']]['sell'] += r['sell_total']
            prev_stores[r['store']]['sell'] += r['sell_total']

    holiday_tag = _is_near_holiday(target, target.year)

    L = []
    L.append("💡 일일 분석")
    L.append("")

    # [수익성]
    L.append("[수익성]")
    L.append(f"  · 일마진 {avg_margin:.1f}%")
    if prev_data:
        prev_buy_t = sum(r['buy_total'] for r in prev_data)
        prev_sell_t = sum(r['sell_total'] for r in prev_data)
        prev_m = ((prev_sell_t - prev_buy_t) / prev_sell_t * 100) if prev_sell_t > 0 else 0
        diff = avg_margin - prev_m
        direction = "개선" if diff > 0 else "하락"
        L.append(f"  → {comp_label}({prev_m:.1f}%) 대비 {diff:+.1f}%p {direction}")

    if total_sell > 0 and len(store_list) >= 3:
        top3 = sum(sd['sell'] for _, sd in store_list[:3])
        concentration = top3 / total_sell * 100
        L.append(f"  · 매출집중도(TOP3): {concentration:.0f}%")
    L.append("")

    # [매장 분석]
    L.append("[매장 분석]")
    high_m_stores = [(s, sd) for s, sd in store_list if sd['sell'] > 0 and ((sd['sell'] - sd['buy']) / sd['sell'] * 100) > 25]
    low_m_stores = [(s, sd) for s, sd in store_list if sd['sell'] > 0 and 0 < ((sd['sell'] - sd['buy']) / sd['sell'] * 100) < 12]
    if high_m_stores:
        L.append(f"  · 고마진(25%+): {', '.join(_shorten(s) for s, _ in high_m_stores[:3])}")
    if low_m_stores:
        L.append(f"  · 저마진(12%-): {', '.join(_shorten(s) for s, _ in low_m_stores[:3])}")
        L.append("  → 해당 매장 납품 단가 재검토 필요")
    L.append("")

    # 인사이트 엔진 (리스크/개선) — 현재 비활성화 (인사이트 정확도 개선 후 재활성화)
    # insight_sections = _generate_insights(items, prev_items, stores, prev_stores, holiday_tag, comp_label)
    # for isec in insight_sections:
    #     L.append(isec)

    sections.append("\n".join(L))

    return sections


# ─── 주간 리포트 ──────────────────────────────────────────────

def weekly_report(records, year, month, week):
    """주간 분석 리포트 (섹션 리스트 반환)"""
    import calendar
    start_day = (week - 1) * 7 + 1
    end_day = week * 7
    last_day = calendar.monthrange(year, month)[1]
    if week == 4:
        end_day = last_day

    start_date = datetime(year, month, start_day)
    end_date = datetime(year, month, min(end_day, last_day))

    week_data = [r for r in records
                 if start_date.date() <= r['date'].date() <= end_date.date()]

    if not week_data:
        return [f"📊 도크 주간분석\n{year}년 {month}월 {week}주차\n\n해당 기간에 데이터가 없습니다."]

    # 전주 데이터
    if week > 1:
        prev_start = (week - 2) * 7 + 1
        prev_end = (week - 1) * 7
        prev_data = [r for r in records
                     if datetime(year, month, prev_start).date() <= r['date'].date() <= datetime(year, month, min(prev_end, last_day)).date()]
    else:
        prev_data = []

    # 전주 품목별 집계 (비교용, 품명 정규화 적용)
    prev_items = defaultdict(lambda: {'qty': 0, 'buy': 0, 'sell': 0})
    for r in prev_data:
        nm = _normalize_item(r['item'])
        prev_items[nm]['qty'] += r['qty']
        prev_items[nm]['buy'] += r['buy_total']
        prev_items[nm]['sell'] += r['sell_total']

    # 전주 매장별 집계 (비교용)
    prev_stores = defaultdict(lambda: {'buy': 0, 'sell': 0})
    for r in prev_data:
        prev_stores[r['store']]['buy'] += r['buy_total']
        prev_stores[r['store']]['sell'] += r['sell_total']

    # 일별 집계 (동원 제외)
    daily = defaultdict(lambda: {'buy': 0, 'sell': 0, 'stores': set(), 'items': set()})
    for r in week_data:
        if _is_no_sell_store(r['store']):
            continue
        d = r['date'].date()
        daily[d]['buy'] += r['buy_total']
        daily[d]['sell'] += r['sell_total']
        daily[d]['stores'].add(r['store'])
        daily[d]['items'].add(r['item'])

    total_buy = sum(r['buy_total'] for r in week_data if not _is_no_sell_store(r['store']))
    total_sell = sum(r['sell_total'] for r in week_data if not _is_no_sell_store(r['store']))
    total_profit = total_sell - total_buy
    avg_margin = (total_profit / total_sell * 100) if total_sell > 0 else 0
    real_days = [d for d, dd in daily.items() if dd['sell'] > 0]
    active_days = len(real_days)

    # 매장별 주간 집계 (동원 제외)
    stores = defaultdict(lambda: {'buy': 0, 'sell': 0, 'days': set(), 'items': 0})
    for r in week_data:
        if _is_no_sell_store(r['store']):
            continue
        stores[r['store']]['buy'] += r['buy_total']
        stores[r['store']]['sell'] += r['sell_total']
        stores[r['store']]['days'].add(r['date'].date())
        stores[r['store']]['items'] += 1

    # 품목별 주간 집계 (동원 제외, 품명 정규화 적용)
    items = defaultdict(lambda: {'qty': 0, 'buy': 0, 'sell': 0, 'stores': set()})
    for r in week_data:
        if _is_no_sell_store(r['store']):
            continue
        nm = _normalize_item(r['item'])
        items[nm]['qty'] += r['qty']
        items[nm]['buy'] += r['buy_total']
        items[nm]['sell'] += r['sell_total']
        items[nm]['stores'].add(r['store'])

    # 품목별 매장상세 (저마진 분석용)
    item_store_prices = defaultdict(lambda: defaultdict(lambda: {'buy': 0, 'sell': 0, 'qty': 0, 'buy_price': 0}))
    for r in week_data:
        nm = _normalize_item(r['item'])
        item_store_prices[nm][r['store']]['buy'] += r['buy_total']
        item_store_prices[nm][r['store']]['sell'] += r['sell_total']
        item_store_prices[nm][r['store']]['qty'] += r['qty']
        item_store_prices[nm][r['store']]['buy_price'] = r['buy_price']

    weekdays = ['월', '화', '수', '목', '금', '토', '일']

    sections = []

    # ── Section 1: Header + 핵심 지표 + 일별 추이 ──
    L = []
    L.append("━━━━━━━━━━━━━━━━")
    L.append("📊 도크 주간분석 리포트")
    L.append(f"📅 {year}년 {month}월 {week}주차")
    L.append(f"   ({start_date.month}/{start_date.day}~{end_date.month}/{end_date.day})")
    L.append("━━━━━━━━━━━━━━━━")
    L.append("")
    L.append("💰 주간 핵심 지표")
    L.append(f"  매출  {_fmt_money(total_sell)}")
    L.append(f"  매입  {_fmt_money(total_buy)}")
    L.append(f"  이익  {_fmt_money(total_profit)}")
    L.append(f"  마진  {avg_margin:.1f}% {_bar(avg_margin)}")
    L.append(f"  영업  {active_days}일 · {len(stores)}매장")
    if active_days > 0:
        L.append(f"  일평균 {_fmt_money(total_sell / active_days)}")

    if prev_data:
        prev_sell = sum(r['sell_total'] for r in prev_data)
        prev_buy = sum(r['buy_total'] for r in prev_data)
        prev_margin = ((prev_sell - prev_buy) / prev_sell * 100) if prev_sell > 0 else 0
        if prev_sell > 0:
            L.append(f"  전주비 {_change_arrow((total_sell - prev_sell) / prev_sell * 100)}")
            margin_diff = avg_margin - prev_margin
            if abs(margin_diff) >= 0.5:
                L.append(f"  마진변동 {margin_diff:+.1f}%p")
    L.append("")

    L.append("📈 일별 추이")
    max_sell = max((dd['sell'] for dd in daily.values()), default=1)
    for d in sorted(daily.keys()):
        dd = daily[d]
        wd = weekdays[d.weekday()]
        if dd['sell'] > 0:
            margin_d = (dd['sell'] - dd['buy']) / dd['sell'] * 100
            bar_len = round(dd['sell'] / max_sell * 6) if max_sell > 0 else 0
            bar = '▓' * bar_len + '░' * (6 - bar_len)
            # 토요일 긴급입고 표시
            if d.weekday() == 5:  # 토요일
                sat_stores = ', '.join(_shorten(s) for s in dd['stores'])
                L.append(f"  {d.day}({wd}) {bar} {_fmt_money(dd['sell'])} 마진{margin_d:.0f}% ({sat_stores} 긴급입고)")
            else:
                L.append(f"  {d.day}({wd}) {bar} {_fmt_money(dd['sell'])} 마진{margin_d:.0f}%")
        else:
            L.append(f"  {d.day}({wd}) ░░░░░░ (단가 미입력)")
    sections.append("\n".join(L))

    # ── Section 2: 매장별 주간매출 (간결 포맷) ──
    store_list = sorted(stores.items(), key=lambda x: x[1]['sell'], reverse=True)
    L = []
    L.append(f"🏪 주간매출 ({len(store_list)}매장)")
    for i, (sname, sd) in enumerate(store_list, 1):
        short = _shorten(sname)
        margin = ((sd['sell'] - sd['buy']) / sd['sell'] * 100) if sd['sell'] > 0 else 0
        profit = sd['sell'] - sd['buy']
        days = len(sd['days'])
        if sd['sell'] > 0:
            L.append(f"")
            L.append(f"{i}.{short}/{_s(sd['sell'])} {margin:.0f}%({_s(profit)}),{days}일")
        else:
            L.append(f"")
            L.append(f"{i}.{short}/{sd['items']}건,{days}일")
    sections.append("\n".join(L))

    # ── Section 3: 주간매입 TOP10 (간결 포맷) ──
    item_list = sorted(items.items(), key=lambda x: x[1]['buy'], reverse=True)
    L = []
    L.append("📦 주간매입 TOP10")
    for i, (name, d) in enumerate(item_list[:10], 1):
        margin = _item_margin(d)
        profit = d['sell'] - d['buy']
        L.append(f"")
        if d['sell'] > 0:
            L.append(f"{i}.{name}/{_s(d['buy'])} {margin:.0f}%({_s(profit)}),{len(d['stores'])}매장")
        else:
            L.append(f"{i}.{name}/{_s(d['buy'])},{len(d['stores'])}매장")
    sections.append("\n".join(L))

    # ── Section 4: 고마진/저마진 (간결 + 과일 별도) ──
    MIN_BUY = 100000
    margin_items = [(name, d, _item_margin(d)) for name, d in items.items()
                    if d['sell'] > 0 and d['buy'] >= MIN_BUY]
    margin_items.sort(key=lambda x: x[2], reverse=True)

    L = []
    high_margin = [x for x in margin_items if x[2] >= 15 and not _is_fruit(x[0])]
    if high_margin:
        L.append("💎 고마진 TOP")
        for i, (name, d, m) in enumerate(high_margin[:6], 1):
            profit = d['sell'] - d['buy']
            L.append(f"")
            L.append(f"{i}.{name}/{m:.0f}%({_s(profit)}),매입{_s(d['buy'])}")

    low_margin_veg = [x for x in margin_items if x[2] < 15 and not _is_fruit(x[0]) and not _is_fixed_price(x[0])]
    low_margin_veg.sort(key=lambda x: x[2])
    if low_margin_veg:
        L.append(f"")
        L.append("⚠️ 저마진 야채")
        for idx, (name, d, m) in enumerate(low_margin_veg[:5]):
            profit = d['sell'] - d['buy']
            L.append(f"  ▸ {name}/{m:.0f}%({_s(profit)}),매입{_s(d['buy'])}")
            # 최저마진 품목에 컨텍스트 추가 (매장별 가격차, 전주대비 급등)
            if idx == 0 and name in item_store_prices:
                sdetail = item_store_prices[name]
                if len(sdetail) >= 1:
                    dominant = max(sdetail.items(), key=lambda x: x[1]['buy'])
                    dom_name, dom_d = dominant
                    notes = []
                    notes.append(f"{_shorten(dom_name)} 매입가{dom_d['buy_price']:,.0f}원")
                    if name in prev_items and prev_items[name]['qty'] > 0:
                        prev_unit = prev_items[name]['buy'] / prev_items[name]['qty']
                        curr_unit = d['buy'] / d['qty']
                        pct = (curr_unit - prev_unit) / prev_unit * 100
                        if abs(pct) > 15:
                            notes.append(f"전주대비 {pct:+.0f}%")
                    L.append(f"    └ {', '.join(notes)}, 타품목 마진보전")

    fruit_items = [x for x in margin_items if _is_fruit(x[0])]
    if fruit_items:
        fruit_margins = [m for _, _, m in fruit_items]
        avg_fruit_m = sum(fruit_margins) / len(fruit_margins)
        fruit_buy = sum(d['buy'] for _, d, _ in fruit_items)
        L.append(f"")
        L.append(f"🍎 과일류 {len(fruit_items)}건 평균{avg_fruit_m:.0f}% 매입{_s(fruit_buy)}")
        L.append(f"   (구조적 저마진 4~7% 유지중)")

    if L:
        sections.append("\n".join(L))

    # ── Section 5: 주간 분석 (전주 대비 데이터 기반) ──
    L = []
    L.append("💡 주간 분석")
    L.append("")

    # 설날 근접 여부 확인
    holiday_tag = _is_near_holiday(start_date, year)

    # [수익성]
    L.append("[수익성]")
    L.append(f"  · 주간마진 {avg_margin:.1f}%")
    if prev_data:
        prev_sell_t = sum(r['sell_total'] for r in prev_data)
        prev_buy_t = sum(r['buy_total'] for r in prev_data)
        prev_m = ((prev_sell_t - prev_buy_t) / prev_sell_t * 100) if prev_sell_t > 0 else 0
        diff = avg_margin - prev_m
        direction = "개선" if diff > 0 else "하락"
        L.append(f"  → 전주({prev_m:.1f}%) 대비 {diff:+.1f}%p {direction}")

    if total_sell > 0 and len(store_list) >= 3:
        top3 = sum(sd['sell'] for _, sd in store_list[:3])
        concentration = top3 / total_sell * 100
        L.append(f"  · 매출집중도(TOP3): {concentration:.0f}%")
    L.append("")

    # [운영 효율]
    L.append("[운영 효율]")
    everyday = [s for s, d in stores.items() if len(d['days']) >= active_days and active_days >= 3]
    if everyday:
        names = ', '.join(_shorten(s) for s in everyday[:5])
        suffix = f" 외 {len(everyday)-5}개" if len(everyday) > 5 else ""
        L.append(f"  · 매일발주({active_days}일): {names}{suffix}")

    low_freq = [(s, len(d['days']), stores[s]['sell']) for s, d in stores.items() if len(d['days']) <= 2]
    low_freq.sort(key=lambda x: x[1])
    if low_freq:
        parts = [f"{_shorten(s)}({d}일)" for s, d, _ in low_freq]
        L.append(f"  · 비정기: {', '.join(parts)}")
        low_big = [(s, sv) for s, d, sv in low_freq if sv >= 500000]
        if low_big:
            L.append(f"  → {', '.join(_shorten(s) for s,_ in low_big)} 매출 규모 있음, 빈도 개선 가능")
    L.append("")

    # [전문 인사이트] (리스크/개선) — 현재 비활성화 (인사이트 정확도 개선 후 재활성화)
    # insight_sections = _generate_insights(items, prev_items, stores, prev_stores, holiday_tag)
    # sections.extend(insight_sections)

    sections.append("\n".join(L))

    return sections



def _generate_insights(items, prev_items, stores, prev_stores, holiday_tag=None, comp_label="전기 대비"):
    """리스크 및 개선 분석 엔진"""
    sections = []
    
    # [리스크]
    L = []
    L.append("[리스크]")
    has_risk = False

    # 매입단가 급등 (10%+, 과일/고정가격 제외)
    price_surged = []
    for name, d in items.items():
        if _is_fruit(name) or _is_fixed_price(name):
            continue
        if name in prev_items and d['qty'] >= 1 and prev_items[name]['qty'] >= 1 and d['buy'] >= 100000:
            curr_unit = d['buy'] / d['qty']
            prev_unit = prev_items[name]['buy'] / prev_items[name]['qty']
            if prev_unit > 0:
                change = (curr_unit - prev_unit) / prev_unit * 100
                if change > 15:
                    price_surged.append((name, change, curr_unit, prev_unit))
    price_surged.sort(key=lambda x: x[1], reverse=True)
    if price_surged:
        has_risk = True
        parts = [f"{n}({c:+.0f}%)" for n, c, _, _ in price_surged[:3]]
        L.append(f"  · 매입단가급등: {', '.join(parts)}")

    # 마진율 악화 (5%p+, 과일/고정가격 제외)
    margin_worse = []
    for name, d in items.items():
        if _is_fruit(name) or _is_fixed_price(name):
            continue
        if name in prev_items and d['sell'] > 0 and prev_items[name]['sell'] > 0 and d['buy'] >= 50000:
            curr_m = _item_margin(d)
            prev_m_i = _item_margin(prev_items[name])
            if curr_m < prev_m_i - 5:
                ctx = _get_context(name)
                margin_worse.append((name, curr_m, prev_m_i, curr_m - prev_m_i, ctx))
    margin_worse.sort(key=lambda x: x[3])
    if margin_worse:
        has_risk = True
        parts = []
        for n, cm, pm, d, ctx in margin_worse[:3]:
            tag = f"·{ctx}" if ctx else ""
            parts.append(f"{n}({d:+.0f}%p{tag})")
        L.append(f"  · 마진악화: {', '.join(parts)}")

    # 사용량 감소 (30%+, 과일/고정가격 제외)
    qty_down = []
    for name, d in items.items():
        if _is_fruit(name) or _is_fixed_price(name):
            continue
        if name in prev_items and prev_items[name]['qty'] > 0 and d['buy'] >= 50000:
            change = (d['qty'] - prev_items[name]['qty']) / prev_items[name]['qty'] * 100
            if change < -30:
                qty_down.append((name, change, d['qty'], prev_items[name]['qty']))
    qty_down.sort(key=lambda x: x[1])
    if qty_down:
        has_risk = True
        parts = [f"{n}({c:+.0f}%)" for n, c, _, _ in qty_down[:3]]
        L.append(f"  · 사용량감소({comp_label}): {', '.join(parts)}")

    # 전기 대비 매출 감소 매장 (25%+)
    store_down = []
    for sname, sd in stores.items():
        curr_sell = sd['sell'] if isinstance(sd, dict) else sum(r['sell_total'] for r in sd)
        
        if sname in prev_stores:
            prev_sell = prev_stores[sname]['sell'] if isinstance(prev_stores[sname], dict) else sum(r['sell_total'] for r in prev_stores[sname])
            
            if prev_sell > 0:
                change = (curr_sell - prev_sell) / prev_sell * 100
                if change < -25:
                    store_down.append((_shorten(sname), change))
    store_down.sort(key=lambda x: x[1])
    if store_down:
        has_risk = True
        parts = [f"{n}({c:+.0f}%)" for n, c in store_down[:3]]
        L.append(f"  · 매출감소({comp_label}): {', '.join(parts)}")

    if not has_risk:
        L.append("  · 특이사항 없음")
    L.append("")
    sections.append("\n".join(L))

    # [개선]
    L = []
    L.append("[개선]")
    has_opp = False

    # 사용량 증가 (30%+, 과일/고정가격 제외)
    qty_up = []
    for name, d in items.items():
        if _is_fruit(name) or _is_fixed_price(name):
            continue
        if name in prev_items and prev_items[name]['qty'] > 0 and d['buy'] >= 50000:
            change = (d['qty'] - prev_items[name]['qty']) / prev_items[name]['qty'] * 100
            if change > 30:
                qty_up.append((name, change, d['qty']))
    qty_up.sort(key=lambda x: x[1], reverse=True)
    if qty_up:
        has_opp = True
        parts = [f"{n}({c:+.0f}%)" for n, c, _ in qty_up[:3]]
        L.append(f"  · 사용량증가({comp_label}): {', '.join(parts)}")

    # 마진율 개선 (5%p+, 과일/고정가격 제외, 컨텍스트 표시)
    margin_better = []
    for name, d in items.items():
        if _is_fruit(name) or _is_fixed_price(name):
            continue
        if name in prev_items and d['sell'] > 0 and prev_items[name]['sell'] > 0 and d['buy'] >= 50000:
            curr_m = _item_margin(d)
            prev_m_i = _item_margin(prev_items[name])
            if curr_m > prev_m_i + 5:
                ctx = _get_context(name)
                margin_better.append((name, curr_m, prev_m_i, curr_m - prev_m_i, ctx))
    margin_better.sort(key=lambda x: x[3], reverse=True)
    if margin_better:
        has_opp = True
        parts = []
        for n, cm, pm, d, ctx in margin_better[:3]:
            tag = f"·{ctx}" if ctx else ""
            parts.append(f"{n}({d:+.0f}%p{tag})")
        L.append(f"  · 마진개선: {', '.join(parts)}")

    # 매입단가 하락 (10%+, 과일/고정가격 제외)
    price_down = []
    for name, d in items.items():
        if _is_fruit(name) or _is_fixed_price(name):
            continue
        if name in prev_items and d['qty'] > 0 and prev_items[name]['qty'] > 0 and d['buy'] >= 50000:
            curr_unit = d['buy'] / d['qty']
            prev_unit = prev_items[name]['buy'] / prev_items[name]['qty']
            if prev_unit > 0:
                change = (curr_unit - prev_unit) / prev_unit * 100
                if change < -10:
                    price_down.append((name, change))
    price_down.sort(key=lambda x: x[1])
    if price_down:
        has_opp = True
        parts = [f"{n}({c:+.0f}%)" for n, c in price_down[:3]]
        L.append(f"  · 매입단가하락: {', '.join(parts)}")
    
    # 고마진 품목 (고마진가공품 -> 고마진 품목으로 타이틀 변경 및 성격 조정)
    MIN_BUY = 30000 
    margin_items = [(name, d, _item_margin(d)) for name, d in items.items()
                    if d['sell'] > 0 and d['buy'] >= MIN_BUY]
    margin_items.sort(key=lambda x: x[2], reverse=True)
    high_margin = [x for x in margin_items if x[2] >= 15 and not _is_fruit(x[0])]

    if high_margin:
        has_opp = True
        # 주간 리포트 등에서는 기존 형식을 유지할 수 있으나, 여기서는 타이틀만 '고마진품목'으로 노출
        # 일일리포트에서는 이미 Section 4에서 상세 분석하므로 여기서는 요약만
        top_items = ', '.join(f"{n}({m:.0f}%)" for n, _, m in high_margin[:3])
        L.append(f"  · 고마진품목: {top_items}")

    if not has_opp:
        L.append("  · 특이사항 없음")

    sections.append("\n".join(L))
    return sections


# ─── 월간 리포트 ──────────────────────────────────────────────

def monthly_report(records, year, month):
    """월간 분석 리포트 (섹션 리스트 반환)"""
    import calendar
    last_day = calendar.monthrange(year, month)[1]
    start = datetime(year, month, 1)
    end = datetime(year, month, last_day)

    month_data = [r for r in records
                  if start.date() <= r['date'].date() <= end.date()]

    if not month_data:
        return [f"📊 도크 월간분석\n{year}년 {month}월\n\n해당 월에 데이터가 없습니다."]

    # 전월 데이터
    if month > 1:
        prev_month = month - 1
        prev_year = year
    else:
        prev_month = 12
        prev_year = year - 1
    prev_last = calendar.monthrange(prev_year, prev_month)[1]
    prev_data = [r for r in records
                 if datetime(prev_year, prev_month, 1).date() <= r['date'].date() <= datetime(prev_year, prev_month, prev_last).date()]

    total_buy = sum(r['buy_total'] for r in month_data if not _is_no_sell_store(r['store']))
    total_sell = sum(r['sell_total'] for r in month_data if not _is_no_sell_store(r['store']))
    total_profit = total_sell - total_buy
    avg_margin = (total_profit / total_sell * 100) if total_sell > 0 else 0

    active_dates = sorted(set(r['date'].date() for r in month_data))
    active_days = len(active_dates)

    # 주차별 집계 (동원 제외)
    weekly = defaultdict(lambda: {'buy': 0, 'sell': 0, 'days': set(), 'stores': set()})
    for r in month_data:
        if _is_no_sell_store(r['store']):
            continue
        w = _week_number(r['date'].day)
        weekly[w]['buy'] += r['buy_total']
        weekly[w]['sell'] += r['sell_total']
        weekly[w]['days'].add(r['date'].date())
        weekly[w]['stores'].add(r['store'])

    # 매장별 월간 집계 (동원 제외)
    stores = defaultdict(lambda: {'buy': 0, 'sell': 0, 'days': set(), 'items': set()})
    for r in month_data:
        if _is_no_sell_store(r['store']):
            continue
        stores[r['store']]['buy'] += r['buy_total']
        stores[r['store']]['sell'] += r['sell_total']
        stores[r['store']]['days'].add(r['date'].date())
        stores[r['store']]['items'].add(r['item'])

    # 품목별 월간 (동원 제외)
    items = defaultdict(lambda: {'qty': 0, 'buy': 0, 'sell': 0, 'stores': set()})
    for r in month_data:
        if _is_no_sell_store(r['store']):
            continue
        items[r['item']]['qty'] += r['qty']
        items[r['item']]['buy'] += r['buy_total']
        items[r['item']]['sell'] += r['sell_total']
        items[r['item']]['stores'].add(r['store'])

    # 공급업체별
    suppliers = defaultdict(lambda: {'count': 0, 'buy': 0})
    for r in month_data:
        if r['supplier']:
            suppliers[r['supplier']]['count'] += 1
            suppliers[r['supplier']]['buy'] += r['buy_total']

    sections = []

    # ── Section 1: Header + 핵심 지표 ──
    L = []
    L.append("━━━━━━━━━━━━━━━━")
    L.append("📊 도크 월간분석 리포트")
    L.append(f"📅 {year}년 {month}월")
    L.append("━━━━━━━━━━━━━━━━")
    L.append("")
    L.append("💰 월간 핵심 지표")
    L.append(f"  매출  {_fmt_money(total_sell)}")
    L.append(f"  매입  {_fmt_money(total_buy)}")
    L.append(f"  이익  {_fmt_money(total_profit)}")
    L.append(f"  마진  {avg_margin:.1f}% {_bar(avg_margin)}")
    L.append(f"  영업  {active_days}일 | {len(stores)}매장")
    if active_days > 0:
        L.append(f"  일평균 {_fmt_money(total_sell / active_days)}")
    if prev_data:
        prev_sell = sum(r['sell_total'] for r in prev_data)
        if prev_sell > 0:
            L.append(f"  전월비 {_change_arrow((total_sell - prev_sell) / prev_sell * 100)}")
    sections.append("\n".join(L))

    # ── Section 2: 주차별 추이 ──
    L = []
    L.append("📈 주차별 추이")
    max_w_sell = max(wd['sell'] for wd in weekly.values()) if weekly else 1
    for w in sorted(weekly.keys()):
        wd = weekly[w]
        margin = ((wd['sell'] - wd['buy']) / wd['sell'] * 100) if wd['sell'] > 0 else 0
        profit = wd['sell'] - wd['buy']
        bar_len = round(wd['sell'] / max_w_sell * 6) if max_w_sell > 0 else 0
        bar = '▓' * bar_len + '░' * (6 - bar_len)
        L.append(f"  {w}주 {bar} {_fmt_money(wd['sell'])}")
        L.append(f"     마진{margin:.0f}% | 이익 {_fmt_money(profit)} | {len(wd['days'])}일")
    sections.append("\n".join(L))

    # ── Section 3: 전체 매장 월매출 ──
    store_list = sorted(stores.items(), key=lambda x: x[1]['sell'], reverse=True)
    L = []
    L.append(f"🏪 매장별 월매출 현황 ({len(store_list)}매장)")
    for i, (sname, sd) in enumerate(store_list, 1):
        short = _shorten(sname)
        margin = ((sd['sell'] - sd['buy']) / sd['sell'] * 100) if sd['sell'] > 0 else 0
        profit = sd['sell'] - sd['buy']
        days = len(sd['days'])
        if sd['sell'] > 0:
            L.append(f"  {i}. {short}")
            L.append(f"     {_fmt_money(sd['sell'])} | {days}일 | 마진{margin:.0f}% | 이익{_fmt_money(profit)}")
        else:
            L.append(f"  {i}. {short} ({len(sd['items'])}품목, {days}일)")
    sections.append("\n".join(L))

    # ── Section 4: 매입 TOP + 마진 상세 ──
    item_list = sorted(items.items(), key=lambda x: x[1]['buy'], reverse=True)
    L = []
    L.append("📦 매입 TOP 품목")
    for i, (name, d) in enumerate(item_list[:10], 1):
        margin = _item_margin(d)
        profit = d['sell'] - d['buy']
        L.append(f"  {i}. {name}")
        L.append(f"     매입 {_fmt_money(d['buy'])} | {d['qty']:.0f}건 | {len(d['stores'])}매장")
        if d['sell'] > 0:
            L.append(f"     마진{margin:.0f}% | 이익 {_fmt_money(profit)}")
    sections.append("\n".join(L))

    # ── Section 5: 고마진/저마진 + 공급업체 ──
    MIN_BUY = 300000  # 월간: 30만원 이상
    margin_items = [(name, d, _item_margin(d)) for name, d in items.items()
                    if d['sell'] > 0 and d['buy'] >= MIN_BUY]
    margin_items.sort(key=lambda x: x[2], reverse=True)

    L = []
    high_margin = [x for x in margin_items if x[2] >= 15]
    if high_margin:
        L.append(f"💎 고마진 품목 TOP (매입{_fmt_money(MIN_BUY)}+)")
        for i, (name, d, m) in enumerate(high_margin[:6], 1):
            profit = d['sell'] - d['buy']
            L.append(f"  {i}. {name} - {m:.0f}%")
            L.append(f"     이익 {_fmt_money(profit)} | 매입 {_fmt_money(d['buy'])}")
        L.append("")

    low_margin = [x for x in margin_items if x[2] < 10]
    low_margin.sort(key=lambda x: x[2])
    if low_margin:
        L.append(f"⚠️ 저마진 주의 (매입{_fmt_money(MIN_BUY)}+, 10%미만)")
        for name, d, m in low_margin[:5]:
            profit = d['sell'] - d['buy']
            L.append(f"  ▸ {name} {m:.0f}%")
            L.append(f"    이익 {_fmt_money(profit)} | 매입 {_fmt_money(d['buy'])}")
        L.append("")

    sup_list = sorted(suppliers.items(), key=lambda x: x[1]['buy'], reverse=True)
    if sup_list:
        L.append("🚚 공급업체 월매입 TOP")
        for sname, sd in sup_list[:6]:
            pct = (sd['buy'] / total_buy * 100) if total_buy > 0 else 0
            bar = _bar(pct, 5)
            L.append(f"  {sname} {bar} {_fmt_money(sd['buy'])} ({pct:.0f}%)")

    if L:
        sections.append("\n".join(L))

    # ── Section 6: 전문 인사이트 ──
    L = []
    L.append("💡 월간 전문 분석")
    L.append("")

    # [수익성 진단]
    L.append("[수익성 진단]")
    L.append(f"  - 월마진 {avg_margin:.1f}%, 총이익 {_fmt_money(total_profit)}")
    if prev_data:
        prev_sell_t = sum(r['sell_total'] for r in prev_data)
        prev_buy_t = sum(r['buy_total'] for r in prev_data)
        prev_m = ((prev_sell_t - prev_buy_t) / prev_sell_t * 100) if prev_sell_t > 0 else 0
        diff = avg_margin - prev_m
        prev_profit = prev_sell_t - prev_buy_t
        profit_diff = total_profit - prev_profit
        L.append(f"  → 전월 대비 마진 {diff:+.1f}%p, 이익 {_fmt_money(profit_diff)}")

    if total_sell > 0 and len(store_list) >= 3:
        top3 = sum(sd['sell'] for _, sd in store_list[:3])
        concentration = top3 / total_sell * 100
        L.append(f"  - 매출집중도(TOP3): {concentration:.0f}%")
    L.append("")

    # [매장 효율]
    L.append("[매장 효율]")
    avg_days = sum(len(sd['days']) for _, sd in store_list) / len(store_list) if store_list else 0
    L.append(f"  - 매장당 평균발주: {avg_days:.1f}일/{active_days}일")

    store_margins = [(s, (d['sell']-d['buy'])/d['sell']*100) for s, d in stores.items() if d['sell'] > 0]
    if store_margins:
        avg_m = sum(m for _, m in store_margins) / len(store_margins)
        best = max(store_margins, key=lambda x: x[1])
        worst = min(store_margins, key=lambda x: x[1])
        L.append(f"  - 매장평균마진: {avg_m:.0f}%")
        L.append(f"  - 최고: {_shorten(best[0])} ({best[1]:.0f}%)")
        L.append(f"  - 최저: {_shorten(worst[0])} ({worst[1]:.0f}%)")
        gap = best[1] - worst[1]
        if gap > 20:
            L.append(f"  → 매장간 마진 격차 {gap:.0f}%p, 저마진매장 개선 필요")
    L.append("")

    # [추세 분석]
    L.append("[추세 분석]")
    week_sells = [(w, d['sell'], d['buy']) for w, d in sorted(weekly.items())]
    if len(week_sells) >= 2:
        first_w = week_sells[0][1]
        last_w = week_sells[-1][1]
        if first_w > 0:
            trend = (last_w - first_w) / first_w * 100
            direction = "성장" if trend > 0 else "감소"
            L.append(f"  - 월내추세: {direction} ({trend:+.0f}%)")

        # 마진 추세
        first_m = ((week_sells[0][1]-week_sells[0][2])/week_sells[0][1]*100) if week_sells[0][1] > 0 else 0
        last_m = ((week_sells[-1][1]-week_sells[-1][2])/week_sells[-1][1]*100) if week_sells[-1][1] > 0 else 0
        m_trend = last_m - first_m
        if abs(m_trend) >= 1:
            m_dir = "개선" if m_trend > 0 else "악화"
            L.append(f"  - 마진추세: {m_dir} ({m_trend:+.1f}%p)")
    L.append("")

    # [전략 제언]
    L.append("[전략 제언]")
    if low_margin:
        L.append(f"  1. 저마진 품목({len(low_margin)}건) 단가 재협상 우선")
    if store_margins:
        low_stores = [s for s, m in store_margins if m < 15]
        if low_stores:
            L.append(f"  2. 저마진매장({len(low_stores)}개) 품목 구성 최적화")
    if high_margin:
        L.append(f"  3. 고마진 품목({len(high_margin)}건) 공급 안정화 및 확대")

    sections.append("\n".join(L))

    return sections


# ─── 카카오톡 전송 ──────────────────────────────────────────

def _get_tokens():
    """카카오 토큰 로드 및 갱신"""
    sys.path.insert(0, os.path.join(PROJECT_ROOT, 'execution'))
    from send_daily_summary import load_and_refresh_token
    return load_and_refresh_token()


def _send_memo(text, tokens):
    """카카오톡 나에게 보내기 (단일 메시지)"""
    url = "https://kapi.kakao.com/v2/api/talk/memo/default/send"
    headers = {"Authorization": f"Bearer {tokens['access_token']}"}
    payload = {
        "template_object": json.dumps({
            "object_type": "text",
            "text": text,
            "link": {
                "web_url": "https://drive.google.com",
                "mobile_web_url": "https://drive.google.com"
            },
            "button_title": "Drive 열기"
        })
    }
    res = requests.post(url, headers=headers, data=payload)
    if res.status_code == 401:
        print("Token expired. Re-authenticate with kakao_auth.py")
        return False
    result = res.json()
    if result.get('result_code') == 0:
        return True
    print(f"  Send failed: {result}")
    return False


def _send_friend(text, tokens, friend_uuid):
    """카카오톡 친구에게 보내기"""
    url = "https://kapi.kakao.com/v1/api/talk/friends/message/default/send"
    headers = {"Authorization": f"Bearer {tokens['access_token']}"}
    payload = {
        "receiver_uuids": json.dumps([friend_uuid]),
        "template_object": json.dumps({
            "object_type": "text",
            "text": text,
            "link": {
                "web_url": "https://drive.google.com",
                "mobile_web_url": "https://drive.google.com"
            },
            "button_title": "Drive 열기"
        })
    }
    res = requests.post(url, headers=headers, data=payload)
    if res.status_code == 401:
        print("Token expired. Re-authenticate with kakao_auth.py")
        return False
    if res.status_code == 403:
        print("Friends scope required. Re-authenticate with: python execution/kakao_auth.py")
        return False
    result = res.json()
    if result.get('successful_receiver_uuids'):
        return True
    print(f"  Send failed: {result}")
    return False


def send_via_kakao(sections, to_friend=None):
    """섹션별로 카카오톡 전송 (각 섹션이 별도 메시지)"""
    tokens = _get_tokens()
    if not tokens:
        print("Cannot send: No valid Kakao token")
        return False

    # 문자열이면 리스트로 변환
    if isinstance(sections, str):
        sections = [sections]

    total = len(sections)
    print(f"Sending {total} section(s) via KakaoTalk...")

    for i, section in enumerate(sections):
        tag = f"\n({i+1}/{total})" if total > 1 else ""
        text = section + tag

        # 카카오톡 텍스트 최대 ~2000자, 넘으면 분할
        if len(text) > 1500:
            chunks = _split_section(text, 1200)
            for j, chunk in enumerate(chunks):
                chunk_tag = f"\n({i+1}-{j+1}/{total})" if len(chunks) > 1 else tag
                msg = chunk.rstrip() + chunk_tag
                if to_friend:
                    ok = _send_friend(msg, tokens, to_friend)
                else:
                    ok = _send_memo(msg, tokens)
                if not ok:
                    return False
                print(f"  Section {i+1}-{j+1} sent.")
                time.sleep(0.5)
        else:
            if to_friend:
                ok = _send_friend(text, tokens, to_friend)
            else:
                ok = _send_memo(text, tokens)
            if not ok:
                return False
            print(f"  Section {i+1}/{total} sent.")

        if i < total - 1:
            time.sleep(0.5)

    return True


def _split_section(text, max_size=1200):
    """줄 단위로 텍스트 분할"""
    chunks = []
    current = ""
    for line in text.split('\n'):
        if len(current) + len(line) + 1 > max_size and current:
            chunks.append(current.rstrip())
            current = line + "\n"
        else:
            current += line + "\n"
    if current.strip():
        chunks.append(current.rstrip())
    return chunks if chunks else [text[:max_size]]


def list_friends():
    """카카오톡 친구 목록 조회"""
    tokens = _get_tokens()
    if not tokens:
        print("No valid token.")
        return

    url = "https://kapi.kakao.com/v1/api/talk/friends"
    headers = {"Authorization": f"Bearer {tokens['access_token']}"}
    res = requests.get(url, headers=headers)

    if res.status_code == 403:
        print("Friends scope required.")
        print("1. kakao_auth.py를 다시 실행하세요 (friends scope 추가됨)")
        print("2. Kakao Developers > 동의항목에서 '카카오 서비스 내 친구목록' 활성화")
        return

    if res.status_code != 200:
        print(f"Error: {res.status_code} {res.text}")
        return

    data = res.json()
    friends = data.get('elements', [])
    print(f"\n카카오톡 친구 목록 ({len(friends)}명):")
    print("-" * 40)
    for f in friends:
        name = f.get('profile_nickname', '(이름없음)')
        uuid = f.get('uuid', '')
        fav = " ★" if f.get('favorite') else ""
        print(f"  {name}{fav}")
        print(f"    UUID: {uuid}")
    print("-" * 40)
    print("\n사용법: --to \"UUID\" 옵션으로 친구에게 전송")


def main():
    parser = argparse.ArgumentParser(description="도크 비즈니스 분석 리포트")
    parser.add_argument("--type", choices=["daily", "weekly", "monthly"],
                        help="리포트 유형")
    parser.add_argument("--date", help="일일 리포트 날짜 (YYYY-MM-DD)")
    parser.add_argument("--month", help="주간/월간 리포트 대상 월 (YYYY-MM)")
    parser.add_argument("--week", type=int, help="주간 리포트 주차 (1~4)")
    parser.add_argument("--file", dest="master", help="마스터 파일 경로 (선택)")
    parser.add_argument("--master", help="마스터 파일 경로 (구버전 호환)")
    parser.add_argument("--preview", action="store_true", help="미리보기만 (카카오톡 전송 안함)")
    parser.add_argument("--to", help="친구 UUID (친구에게 전송)")
    parser.add_argument("--group", help="카카오톡 PC 단톡방 이름 (예: 직납방)")
    parser.add_argument("--list-friends", action="store_true", help="카카오톡 친구 목록 조회")
    args = parser.parse_args()

    # 친구 목록 조회
    if args.list_friends:
        list_friends()
        return 0

    if not args.type:
        parser.print_help()
        return 1

    # 마스터 데이터 로드
    records = load_master_data(args.master)
    if not records:
        return 1

    # 리포트 생성 (섹션 리스트 반환)
    if args.type == "daily":
        if not args.date:
            args.date = datetime.now().strftime("%Y-%m-%d")
        sections = daily_report(records, args.date)

    elif args.type == "weekly":
        if not args.month or not args.week:
            print("Error: --month and --week required for weekly report")
            return 1
        year, month = map(int, args.month.split('-'))
        sections = weekly_report(records, year, month, args.week)

    elif args.type == "monthly":
        if not args.month:
            print("Error: --month required for monthly report")
            return 1
        year, month = map(int, args.month.split('-'))
        sections = monthly_report(records, year, month)

    # 출력 (모든 섹션)
    print("\n" + "=" * 50)
    for i, section in enumerate(sections):
        print(f"\n--- Section {i+1}/{len(sections)} ---")
        print(section)
    print("\n" + "=" * 50)

    if args.preview:
        print(f"\n(Preview mode - 카카오톡 전송 안함, {len(sections)} sections)")
        return 0

    # 카카오톡 전송
    print(f"\nSending {len(sections)} sections via KakaoTalk...")
    
    # 전송 대상 리스트 구성
    recipients = [None]  # None: 나에게 보내기
    
    # --to 옵션이 있으면 우선 적용, 없으면 .env의 KAKAO_TARGET_UUIDS 사용
    to_val = args.to or os.getenv("KAKAO_TARGET_UUIDS")
    if to_val:
        uuids = [u.strip() for u in to_val.split(",") if u.strip()]
        recipients.extend(uuids)
    
    success_count = 0
    for target in recipients:
        target_label = "Me" if target is None else f"Friend({target[:8]}...)"
        print(f"  Target: {target_label}")
        if send_via_kakao(sections, to_friend=target):
            success_count += 1
            print(f"  Sent to {target_label}!")
        else:
            print(f"  Failed sending to {target_label}.")

    if success_count > 0:
        print(f"\nSuccessfully sent to {success_count} recipients.")
        return 0
    else:
        fallback = os.path.join(PROJECT_ROOT, "data", "outputs",
                                f"report_{args.type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        with open(fallback, 'w', encoding='utf-8') as f:
            f.write("\n\n---\n\n".join(sections))
        print(f"Saved to: {fallback}")
        return 1

    # 단톡방 전송 (--group 옵션)
    if args.group:
        print(f"\n단톡방 '{args.group}' 전송 시작...")
        try:
            from send_kakao_group import send_to_group
        except ImportError:
            exec_dir = os.path.dirname(os.path.abspath(__file__))
            sys.path.insert(0, exec_dir)
            from send_kakao_group import send_to_group
        send_to_group(args.group, sections)

    return 0


if __name__ == "__main__":
    sys.exit(main())
