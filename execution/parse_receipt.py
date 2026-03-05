"""
공급처 영수증 파싱 엔진.

다양한 형태(ERP URL, POS 사진, 손글씨 사진, 카톡 텍스트)의
공급처 영수증에서 품목+단가를 추출한다.

Phase 1: URL 파싱 (건영농산 itanet, 오복상회 marketbom)
Phase 2: 이미지 OCR (Gemini Vision)
Phase 3: 손글씨/텍스트

Usage:
  # URL 파싱
  python execution/parse_receipt.py --url "https://www.itanet.co.kr/..." [--supplier 건영농산]

  # 이미지 파싱
  python execution/parse_receipt.py --image "/path/to/receipt.jpg" --supplier 경북상회

  # 이미지 URL 파싱
  python execution/parse_receipt.py --image-url "https://..." --supplier 가야웰빙

  # 텍스트 파싱
  python execution/parse_receipt.py --text "감자 62000\\n당근 35000" --supplier 명화농산
"""

import argparse
import json
import os
import re
import sys
from urllib.parse import urlparse

# ==================== 설정 ====================

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RESOURCES_DIR = os.path.join(PROJECT_ROOT, 'skills', 'order_processing', 'resources')
RECEIPT_MAP_PATH = os.path.join(RESOURCES_DIR, 'receipt_item_map.json')

# receipt_item_map.json 로드
with open(RECEIPT_MAP_PATH, 'r', encoding='utf-8') as f:
    RECEIPT_MAP = json.load(f)

SUPPLIER_DETECTION = RECEIPT_MAP['supplier_detection']
ITEM_MAP = RECEIPT_MAP['item_map']
PRICE_MULTIPLIER = RECEIPT_MAP.get('supplier_price_multiplier', {})
RECEIPT_TYPE = RECEIPT_MAP.get('supplier_receipt_type', {})


# ==================== 공급처 감지 ====================

def detect_supplier_from_url(url):
    """URL 도메인으로 공급처 식별."""
    parsed = urlparse(url)
    domain = parsed.netloc.lower()
    for pattern, supplier in SUPPLIER_DETECTION['url_patterns'].items():
        if pattern in domain:
            return supplier
    return None


def detect_supplier_from_text(text):
    """텍스트 내 상호명 키워드로 공급처 식별."""
    for keyword, supplier in SUPPLIER_DETECTION['text_keywords'].items():
        if keyword in text:
            return supplier
    return None


# ==================== URL 파싱 (Phase 1) ====================

def parse_url_receipt(url, supplier_hint=None):
    """ERP URL에서 영수증 데이터 추출.

    Returns:
        dict: {"supplier", "confidence", "items": [{"name", "qty", "unit_price", "total"}]}
    """
    supplier = supplier_hint or detect_supplier_from_url(url)
    if not supplier:
        return {"error": "공급처를 식별할 수 없습니다", "url": url}

    parsed = urlparse(url)
    domain = parsed.netloc.lower()

    if 'itanet.co.kr' in domain:
        return _parse_itanet_html(url, supplier)
    elif 'marketbom.com' in domain:
        return _parse_marketbom_html(url, supplier)
    else:
        return {"error": f"지원하지 않는 URL 도메인: {domain}", "url": url}


def _parse_itanet_html(url, supplier):
    """건영농산 (itanet.co.kr) ERP 영수증 HTML 파싱.

    itanet ERP는 테이블 형식으로 품목/수량/단가/금액을 표시한다.
    """
    import requests
    from bs4 import BeautifulSoup

    try:
        resp = requests.get(url, timeout=15)
        resp.encoding = resp.apparent_encoding or 'utf-8'
        soup = BeautifulSoup(resp.text, 'html.parser')

        items = []

        # 테이블 행에서 품목 데이터 추출
        tables = soup.find_all('table')
        for table in tables:
            rows = table.find_all('tr')
            for row in rows:
                cells = row.find_all(['td', 'th'])
                if len(cells) < 3:
                    continue

                texts = [c.get_text(strip=True) for c in cells]

                # 헤더 행 스킵
                if any(h in texts[0] for h in ['품목', '상품명', '번호', 'No', '#']):
                    continue

                # 금액이 포함된 행 찾기 (숫자 셀이 2개 이상)
                numeric_cells = []
                name_cell = None
                for i, t in enumerate(texts):
                    cleaned = t.replace(',', '').replace('-', '').strip()
                    if cleaned.isdigit() and int(cleaned) > 0:
                        numeric_cells.append((i, int(cleaned)))
                    elif t and not cleaned.isdigit() and name_cell is None:
                        name_cell = t

                if name_cell and len(numeric_cells) >= 2:
                    # 일반적으로: 품명, 수량, 단가, 금액
                    # 수량은 보통 작은 수, 금액은 큰 수
                    nums = sorted(numeric_cells, key=lambda x: x[1])
                    qty = nums[0][1] if nums[0][1] < 1000 else 1
                    total = nums[-1][1]
                    unit_price = total // qty if qty > 0 else total

                    # 수량이 1000 이상이면 단가일 가능성이 높음
                    if len(nums) >= 3:
                        # 3개 이상: 수량, 단가, 금액
                        qty = nums[0][1]
                        unit_price = nums[1][1]
                        total = nums[-1][1]
                    elif len(nums) == 2:
                        # 2개: 단가, 금액 또는 수량, 금액
                        if nums[0][1] < 100:
                            qty = nums[0][1]
                            total = nums[1][1]
                            unit_price = total // qty if qty > 0 else total
                        else:
                            unit_price = nums[0][1]
                            total = nums[1][1]
                            qty = total // unit_price if unit_price > 0 else 1

                    items.append({
                        "name": name_cell.strip(),
                        "qty": qty,
                        "unit_price": unit_price,
                        "total": total,
                    })

        if not items:
            # 테이블이 없으면 전체 텍스트에서 패턴 추출 시도
            text = soup.get_text()
            items = _extract_items_from_text(text)

        return {
            "supplier": supplier,
            "confidence": "high" if items else "low",
            "source": "itanet_url",
            "items": items,
        }

    except Exception as e:
        return {"error": f"itanet 파싱 실패: {e}", "supplier": supplier}


def _parse_marketbom_html(url, supplier):
    """오복상회 (marketbom.com) ERP 영수증 HTML 파싱.

    마켓봄 ERP는 거래명세서 형식으로 품목/수량/단가/금액을 표시한다.
    """
    import requests
    from bs4 import BeautifulSoup

    try:
        resp = requests.get(url, timeout=15)
        resp.encoding = resp.apparent_encoding or 'utf-8'
        soup = BeautifulSoup(resp.text, 'html.parser')

        items = []

        # 마켓봄은 보통 상품명/규격/수량/단가/공급가/합계 테이블
        tables = soup.find_all('table')
        for table in tables:
            rows = table.find_all('tr')
            header_idx = {}

            for row in rows:
                cells = row.find_all(['td', 'th'])
                texts = [c.get_text(strip=True) for c in cells]

                # 헤더 감지
                if any(h in ' '.join(texts) for h in ['상품명', '품명', '품목명']):
                    for i, t in enumerate(texts):
                        t_lower = t.strip()
                        if t_lower in ('상품명', '품명', '품목명', '품목'):
                            header_idx['name'] = i
                        elif t_lower in ('수량', 'QTY', 'qty'):
                            header_idx['qty'] = i
                        elif t_lower in ('단가', '매입가', '공급가'):
                            header_idx['price'] = i
                        elif t_lower in ('금액', '합계', '매입금액', '공급금액'):
                            header_idx['total'] = i
                    continue

                if not header_idx:
                    continue

                # 데이터 행
                if len(cells) <= max(header_idx.values(), default=0):
                    continue

                name_i = header_idx.get('name')
                if name_i is None or name_i >= len(texts):
                    continue

                name = texts[name_i].strip()
                if not name or name in ('합계', '총합계', '부가세', 'VAT'):
                    continue

                def parse_num(idx):
                    if idx is None or idx >= len(texts):
                        return 0
                    val = texts[idx].replace(',', '').replace('-', '').strip()
                    try:
                        return int(float(val))
                    except (ValueError, TypeError):
                        return 0

                qty = parse_num(header_idx.get('qty')) or 1
                unit_price = parse_num(header_idx.get('price'))
                total = parse_num(header_idx.get('total'))

                if not unit_price and total and qty:
                    unit_price = total // qty
                if not total and unit_price and qty:
                    total = unit_price * qty

                if unit_price > 0 or total > 0:
                    items.append({
                        "name": name,
                        "qty": qty,
                        "unit_price": unit_price,
                        "total": total,
                    })

        if not items:
            text = soup.get_text()
            items = _extract_items_from_text(text)

        return {
            "supplier": supplier,
            "confidence": "high" if items else "low",
            "source": "marketbom_url",
            "items": items,
        }

    except Exception as e:
        return {"error": f"marketbom 파싱 실패: {e}", "supplier": supplier}


# ==================== 이미지 OCR (Phase 2) ====================

def _get_gemini():
    """Gemini Vision 모델 로드 (lazy singleton)."""
    global _gemini_model
    if '_gemini_model' not in globals() or _gemini_model is None:
        from dotenv import load_dotenv
        load_dotenv(os.path.join(PROJECT_ROOT, '.env'))
        import google.generativeai as genai
        genai.configure(api_key=os.environ.get("GEMINI_API_KEY", ""))
        globals()['_gemini_model'] = genai.GenerativeModel("gemini-2.5-flash")
    return globals()['_gemini_model']

_gemini_model = None


def parse_image_receipt(image_source, supplier_hint=None):
    """이미지(파일경로 또는 URL)에서 영수증 OCR.

    Args:
        image_source: 로컬 파일 경로 또는 이미지 URL
        supplier_hint: 공급처명 힌트 (정확도 향상)

    Returns:
        dict: {"supplier", "confidence", "items": [...]}
    """
    import google.generativeai as genai

    model = _get_gemini()

    # 이미지 준비
    if image_source.startswith('http'):
        # URL → 다운로드
        import requests
        resp = requests.get(image_source, timeout=15)
        image_data = resp.content
        mime_type = resp.headers.get('content-type', 'image/jpeg')
        image_part = {"mime_type": mime_type, "data": image_data}
    else:
        # 로컬 파일
        import mimetypes
        mime_type = mimetypes.guess_type(image_source)[0] or 'image/jpeg'
        with open(image_source, 'rb') as f:
            image_data = f.read()
        image_part = {"mime_type": mime_type, "data": image_data}

    # 공급처별 힌트
    receipt_type = RECEIPT_MAP.get('supplier_receipt_type', {}).get(supplier_hint, 'unknown')

    supplier_context = ""
    if supplier_hint:
        # 해당 공급처가 취급하는 품목 목록을 힌트로 제공
        known_items = list(ITEM_MAP.get(supplier_hint, {}).keys())
        if known_items:
            supplier_context = f"\n이 공급처({supplier_hint})가 취급하는 품목: {', '.join(known_items)}"
            supplier_context += "\n품목명이 흐리거나 불분명해도 위 품목 리스트에서 가장 가까운 것을 선택해주세요."

    type_hint = ""
    if receipt_type == 'pos_print':
        type_hint = "이 이미지는 POS 영수증 인쇄물입니다. 사진이 회전되어 있을 수 있습니다."
    elif receipt_type == 'pos_screen':
        type_hint = "이 이미지는 POS 화면 캡처입니다."
    elif receipt_type == 'handwritten':
        type_hint = "이 이미지는 손글씨 메모/영수증입니다. 글씨를 정확히 읽어주세요."

    prompt = f"""이 영수증/거래명세서 이미지에서 품목명과 단가(금액)를 추출해주세요.
{type_hint}{supplier_context}

반드시 아래 JSON 형식으로만 응답하세요 (다른 텍스트 없이):
{{
  "supplier_detected": "영수증에서 확인된 상호명 (없으면 null)",
  "items": [
    {{"name": "품목명", "qty": 수량(숫자), "unit_price": 단가(숫자), "total": 금액(숫자)}}
  ]
}}

규칙:
- 품목명은 영수증에 적힌 그대로 (약어, 슬래시 포함)
- 단가와 금액은 원(₩) 단위 정수
- 수량이 명확히 표시되지 않으면 qty=1, unit_price=0, total=표시금액으로 설정
- POS 화면에서 품목명 옆에 금액만 있고 수량/단가 구분이 불명확한 경우: qty=1, unit_price=0, total=해당금액
- 합계/부가세/총액 행은 제외
- 금액이 0이거나 빈 행은 제외
- 중요: 손글씨에서 "13,-" 또는 "13,―" 또는 "13.-" 표기는 13,000원을 의미함 ("-"는 천원 이하 "000"의 약어)
- 예: "26,-" = 26,000원, "812,-" = 812,000원, "5,-" = 5,000원
- 거래명세서 양식의 컬럼: 월/일, 품목, 단위, 수량, 단가, 공급가액 — 단가 컬럼의 값을 unit_price로 사용
- 손글씨가 흐리거나 약어로 쓰여 있어도 최대한 추정하여 읽어주세요"""

    try:
        response = model.generate_content([prompt, image_part])
        text = response.text.strip()

        # JSON 추출 (마크다운 코드블록 제거)
        if '```' in text:
            match = re.search(r'```(?:json)?\s*(.*?)```', text, re.DOTALL)
            if match:
                text = match.group(1).strip()

        result = json.loads(text)

        items = result.get('items', [])
        supplier_detected = result.get('supplier_detected')

        # 공급처 결정
        supplier = supplier_hint
        if not supplier and supplier_detected:
            supplier = detect_supplier_from_text(supplier_detected)

        # 가격 배수 적용 (현진상회: POS 단가 × 10)
        multiplier = PRICE_MULTIPLIER.get(supplier, 1)
        if multiplier != 1:
            for item in items:
                item['unit_price'] = item.get('unit_price', 0) * multiplier
                item['total'] = item.get('total', 0) * multiplier

        confidence = "high" if receipt_type in ('pos_print', 'pos_screen') else "medium"
        if receipt_type == 'handwritten':
            confidence = "low"

        return {
            "supplier": supplier,
            "confidence": confidence,
            "source": f"gemini_ocr_{receipt_type}",
            "items": items,
        }

    except json.JSONDecodeError:
        return {
            "error": f"Gemini 응답 JSON 파싱 실패",
            "raw_response": text[:500] if 'text' in dir() else "no response",
            "supplier": supplier_hint,
        }
    except Exception as e:
        return {"error": f"이미지 OCR 실패: {e}", "supplier": supplier_hint}


# ==================== 텍스트 파싱 (Phase 3) ====================

def parse_text_receipt(text, supplier_hint=None):
    """카톡 텍스트 메시지에서 품목+단가 추출.

    패턴 예시:
      감자 62000
      당근 수입 35,000
      깐양파(중) 10kg  24,000

    Args:
        text: 영수증 텍스트
        supplier_hint: 공급처명

    Returns:
        dict: {"supplier", "confidence", "items": [...]}
    """
    supplier = supplier_hint or detect_supplier_from_text(text)

    items = _extract_items_from_text(text)

    # 가격 배수 적용
    multiplier = PRICE_MULTIPLIER.get(supplier, 1)
    if multiplier != 1:
        for item in items:
            item['unit_price'] = item.get('unit_price', 0) * multiplier
            item['total'] = item.get('total', 0) * multiplier

    return {
        "supplier": supplier,
        "confidence": "medium" if items else "low",
        "source": "text_regex",
        "items": items,
    }


def _extract_items_from_text(text):
    """텍스트에서 "품목 가격" 패턴 추출.

    다양한 패턴 지원:
    - "감자 62000"  "감자 62,000"
    - "감자/왕왕 62,000"
    - "감자 1 62,000"  (품목 수량 금액)
    - "감자  1박스  62,000원"
    """
    items = []
    lines = text.strip().split('\n')

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 합계/부가세 행 스킵
        if any(skip in line for skip in ['합계', '총합', '부가세', 'VAT', '카드', '현금', '거스름']):
            continue

        # 패턴 1: "품목명  단가" or "품목명  수량  단가" or "품목명  수량  단가  금액"
        # 품목명은 한글+영문+숫자+/+()+(공백)
        match = re.match(
            r'^([가-힣a-zA-Z()\/\s\d]+?)\s+'     # 품목명 (non-greedy)
            r'(\d{1,3})\s+'                         # 수량 (1~3자리)
            r'([\d,]+)\s*'                          # 단가
            r'(?:([\d,]+))?\s*원?\s*$',             # 금액(선택)
            line
        )
        if match:
            name = match.group(1).strip()
            qty = int(match.group(2))
            price = int(match.group(3).replace(',', ''))
            total = int(match.group(4).replace(',', '')) if match.group(4) else price * qty
            items.append({"name": name, "qty": qty, "unit_price": price, "total": total})
            continue

        # 패턴 2: "품목명  금액" (수량 없음)
        match = re.match(
            r'^([가-힣a-zA-Z0-9()\/\s]+?)\s+'   # 품목명 (숫자 포함: 세척당근2L 등)
            r'([\d,]+)\s*원?\s*$',                # 금액
            line
        )
        if match:
            name = match.group(1).strip()
            price = int(match.group(2).replace(',', ''))
            if price >= 100:  # 최소 100원 이상
                items.append({"name": name, "qty": 1, "unit_price": price, "total": price})
                continue

        # 패턴 3: 탭/콤마 구분
        parts = re.split(r'[\t,]+', line)
        if len(parts) >= 2:
            name = parts[0].strip()
            # 마지막 숫자 필드를 금액으로
            for p in reversed(parts[1:]):
                cleaned = p.strip().replace(',', '').replace('원', '')
                if cleaned.isdigit() and int(cleaned) >= 100:
                    price = int(cleaned)
                    items.append({"name": name, "qty": 1, "unit_price": price, "total": price})
                    break

    return items


# ==================== 통합 진입점 ====================

def parse_receipt(url=None, image_path=None, image_url=None, text=None, supplier_hint=None):
    """영수증 파싱 통합 진입점.

    유형을 자동 감지하고 적절한 파서를 호출한다.

    Args:
        url: ERP 영수증 URL (itanet/marketbom)
        image_path: 로컬 이미지 파일 경로
        image_url: 이미지 URL (kakaocdn 등)
        text: 텍스트 영수증 내용
        supplier_hint: 공급처명 힌트

    Returns:
        dict: {"supplier", "confidence", "source", "items": [...]}
    """
    if url:
        return parse_url_receipt(url, supplier_hint)
    elif image_path:
        return parse_image_receipt(image_path, supplier_hint)
    elif image_url:
        return parse_image_receipt(image_url, supplier_hint)
    elif text:
        return parse_text_receipt(text, supplier_hint)
    else:
        return {"error": "url, image_path, image_url, text 중 하나를 제공해야 합니다"}


# ==================== 발주리스트 교차 매칭 ====================

def load_supplier_order_items(supplier, date):
    """processed_orders 파일에서 해당 공급처의 발주 품목 목록을 순서대로 로드.

    손글씨 영수증의 OCR 품목명이 부정확할 때, 발주리스트의 품목 순서와
    영수증의 품목 순서가 동일하다는 점을 이용하여 교차 매칭한다.

    Args:
        supplier: 공급처명
        date: 날짜 (YYYY-MM-DD)

    Returns:
        list: [{"name": 품목명, "qty": 수량, "unit": 단위}, ...] 발주 순서대로
    """
    date_compact = date.replace('-', '')
    orders_path = os.path.join(
        PROJECT_ROOT, 'data', 'inputs', f'processed_orders_{date_compact}.txt'
    )

    if not os.path.exists(orders_path):
        return []

    with open(orders_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 공급처 섹션 찾기 (e.g., "-이모유통\n")
    pattern = rf'-\s*{re.escape(supplier)}\s*\n'
    match = re.search(pattern, content)
    if not match:
        return []

    items = []
    lines = content[match.end():].split('\n')
    for line in lines:
        stripped = line.strip()
        if not stripped:
            break
        if stripped.startswith('-'):
            break  # 다음 공급처 섹션 시작
        if stripped.startswith('('):
            continue  # 메타데이터 (금일, 차량정보 등)

        # "품목명 수량단위" 파싱 (e.g., "귤 10박스", "완숙토마토1호 2박스")
        item_match = re.match(
            r'^(.+?)\s+(\d+(?:\.\d+)?)\s*(박스|kg|KG|단|봉|개|망|팩|통|ea|EA|판)?',
            stripped
        )
        if item_match:
            items.append({
                "name": item_match.group(1).strip(),
                "qty": float(item_match.group(2)),
                "unit": item_match.group(3) or "",
            })
        else:
            # 수량 없는 품목 (드물지만 대비)
            items.append({"name": stripped, "qty": 0, "unit": ""})

    return items


def _map_single_item(name, mapping):
    """단일 품목명을 item_map으로 매핑.

    Returns:
        str or None: 매핑된 단가시트 품목명, 매핑 없으면 None
    """
    # 정확 매칭
    if name in mapping:
        return mapping[name]

    # 공백 제거 매칭
    name_nospace = name.replace(' ', '')
    for k, v in mapping.items():
        if k.replace(' ', '') == name_nospace:
            return v

    # 부분 매칭 (긴 키 우선)
    for k in sorted(mapping.keys(), key=len, reverse=True):
        if k in name or name in k:
            return mapping[k]

    return None


# ==================== 결과 → 단가시트 매핑 ====================

def map_to_danga_items(result, supplier=None, order_items=None):
    """파싱 결과를 단가시트 품목명으로 변환.

    receipt_item_map.json의 공급처별 매핑을 적용한다.
    order_items가 제공되면, 영수증 품목 순서와 발주리스트 품목 순서를
    교차 매칭하여 OCR 오인식을 보정한다.

    Args:
        result: parse_receipt() 결과
        supplier: 공급처명 (없으면 result에서 추출)
        order_items: 발주리스트 품목 리스트 (순서 교차매칭용, load_supplier_order_items() 결과)

    Returns:
        list: [{"receipt_name", "danga_name", "unit_price", "total", "skip_reason"}]
    """
    supplier = supplier or result.get('supplier')
    if not supplier:
        return []

    mapping = ITEM_MAP.get(supplier, {})
    receipt_items = result.get('items', [])
    mapped = []

    # 발주리스트 교차 매칭 모드 판별
    # 조건: order_items가 있고, 손글씨(handwritten) 공급처이며, 품목 수가 대응 가능
    use_crossref = False
    order_danga_names = []
    if order_items:
        receipt_type = RECEIPT_TYPE.get(supplier, 'unknown')
        if receipt_type == 'handwritten' and len(order_items) <= len(receipt_items):
            use_crossref = True
            # 발주리스트 품목명 → 단가시트 품목명 미리 매핑
            for oi in order_items:
                dn = _map_single_item(oi['name'], mapping)
                order_danga_names.append(dn or oi['name'])
            print(f"  [CROSSREF] 발주리스트 {len(order_items)}건 ↔ OCR {len(receipt_items)}건 교차매칭")
            for i, (oi, dn) in enumerate(zip(order_items, order_danga_names)):
                print(f"    #{i+1}: {oi['name']} → {dn}")

    # 교차매칭 시 발주 품목 소진 추적 (ITEM_MAP 매핑 우선, 실패 시 잔여 발주 품목 할당)
    remaining_order = list(range(len(order_danga_names)))  # 미사용 발주 인덱스

    for i, item in enumerate(receipt_items):
        receipt_name = item['name'].strip()
        unit_price = item.get('unit_price', 0)
        total = item.get('total', 0)

        if use_crossref:
            # 1단계: ITEM_MAP으로 매핑 시도 (OCR 품목명 기반)
            mapped_name = _map_single_item(receipt_name, mapping)

            if mapped_name and mapped_name in order_danga_names:
                # ITEM_MAP 매핑 성공 + 발주리스트에도 존재 → 확정
                danga_name = mapped_name
                # 발주 소진 처리 (같은 품목 여러 건 가능하므로 첫 번째 미사용 건 소진)
                for idx in remaining_order:
                    if order_danga_names[idx] == mapped_name:
                        remaining_order.remove(idx)
                        break
                print(f"  [CROSSREF] OCR '{receipt_name}' → ITEM_MAP '{danga_name}' (가격: {unit_price:,})")
            elif remaining_order:
                # ITEM_MAP 매핑 실패 → 잔여 발주 품목 중 첫 번째 할당
                fallback_idx = remaining_order.pop(0)
                danga_name = order_danga_names[fallback_idx]
                print(f"  [CROSSREF] OCR '{receipt_name}' → 발주 잔여 #{fallback_idx+1} '{danga_name}' (가격: {unit_price:,})")
            else:
                # 발주 품목 소진 → OCR 기반 폴백
                danga_name = mapped_name or receipt_name
                print(f"  [NOISE?] OCR line #{i+1} '{receipt_name}' → '{danga_name}' - 발주에 없는 추가 라인")
        else:
            # 기존 OCR 기반 매핑
            danga_name = _map_single_item(receipt_name, mapping)

            if danga_name is None:
                # 매핑 없음 → 원본 이름 그대로
                danga_name = receipt_name

        skip_reason = None
        if danga_name == '_FIXED':
            skip_reason = "고정가 품목 (스킵)"

        mapped.append({
            "receipt_name": receipt_name,
            "danga_name": danga_name,
            "unit_price": unit_price,
            "total": total,
            "qty": item.get('qty', 1),
            "skip_reason": skip_reason,
        })

    return mapped


# ==================== 경량 공급처 식별 ====================

def identify_supplier_from_image(image_path):
    """이미지에서 공급처(상호명)만 식별하는 경량 Gemini 호출.

    전체 파싱보다 빠르고 저렴. 대시보드 미리보기용.

    Returns:
        str: 식별된 공급처명 (매핑된 이름). 불명이면 "불명".
    """
    import mimetypes
    model = _get_gemini()

    mime_type = mimetypes.guess_type(image_path)[0] or 'image/jpeg'
    with open(image_path, 'rb') as f:
        image_data = f.read()
    image_part = {"mime_type": mime_type, "data": image_data}

    prompt = (
        "이 영수증/거래명세서 이미지에서 상호명(공급업체명)만 한 줄로 답하세요. "
        "상호명이 보이지 않으면 '불명'이라고 답하세요. "
        "다른 설명 없이 상호명만 출력하세요."
    )

    try:
        response = model.generate_content([prompt, image_part])
        raw_name = response.text.strip().strip('"').strip("'")
        # 알려진 공급처로 매핑
        mapped = detect_supplier_from_text(raw_name)
        return mapped or raw_name
    except Exception as e:
        return f"오류({e})"


def identify_suppliers_batch(image_dir, since_minutes=1440):
    """폴더 내 이미지들의 공급처를 일괄 식별.

    Returns:
        list: [{"file": filename, "supplier": supplier_name}, ...]
    """
    from datetime import datetime, timedelta

    cutoff = datetime.now() - timedelta(minutes=since_minutes)
    extensions = {'.jpg', '.jpeg', '.png', '.bmp'}
    results = []

    files = []
    for fname in os.listdir(image_dir):
        ext = os.path.splitext(fname)[1].lower()
        if ext not in extensions:
            continue
        fpath = os.path.join(image_dir, fname)
        mtime = datetime.fromtimestamp(os.path.getmtime(fpath))
        if mtime >= cutoff:
            files.append((fpath, fname, mtime))

    files.sort(key=lambda x: x[2], reverse=True)

    for fpath, fname, mtime in files:
        print(f"[ID] {fname}...", end=" ", flush=True)
        supplier = identify_supplier_from_image(fpath)
        print(supplier)
        results.append({"file": fname, "supplier": supplier})

    return results


# ==================== CLI ====================

def main():
    parser = argparse.ArgumentParser(description='공급처 영수증 파싱')
    parser.add_argument('--url', help='ERP 영수증 URL')
    parser.add_argument('--image', help='로컬 이미지 파일 경로')
    parser.add_argument('--image-url', help='이미지 URL')
    parser.add_argument('--text', help='텍스트 영수증 (줄바꿈은 \\n)')
    parser.add_argument('--supplier', help='공급처명 힌트')
    parser.add_argument('--map', action='store_true', help='단가시트 매핑 결과도 표시')
    parser.add_argument('--identify-dir', help='영수증 이미지 폴더 (공급처 식별만)')
    parser.add_argument('--since-minutes', type=int, default=1440, help='식별 시 최근 N분 (기본 1440)')
    args = parser.parse_args()

    # 공급처 식별 모드
    if args.identify_dir:
        results = identify_suppliers_batch(args.identify_dir, args.since_minutes)
        print(json.dumps(results, ensure_ascii=False))
        return

    # 텍스트의 리터럴 \n을 실제 줄바꿈으로
    text = args.text.replace('\\n', '\n') if args.text else None

    result = parse_receipt(
        url=args.url,
        image_path=args.image,
        image_url=args.image_url,
        text=text,
        supplier_hint=args.supplier,
    )

    print(json.dumps(result, ensure_ascii=False, indent=2))

    if args.map and 'items' in result:
        print("\n=== 단가시트 매핑 ===")
        mapped = map_to_danga_items(result)
        for m in mapped:
            skip = f"  [{m['skip_reason']}]" if m.get('skip_reason') else ""
            print(f"  {m['receipt_name']} → {m['danga_name']}  단가:{m['unit_price']:,}  합계:{m['total']:,}{skip}")


if __name__ == '__main__':
    main()
