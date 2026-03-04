"""
카카오톡 비즈니스 채널 발주 웹훅 서버

오픈빌더 챗봇이 발주 메시지를 POST로 전달하면,
매장별로 큐에 저장하고 마감 시 파이프라인을 실행한다.

이미지 발주: 백그라운드 Gemini Vision OCR → 다음 메시지에 결과 전달
영수증 처리: 공급처 영수증(URL/이미지) → 파싱 → 단가시트 매입가 입력

사용법:
    python execution/kakao_order_server.py
    python execution/kakao_order_server.py --port 5050
    python execution/kakao_order_server.py --debug
"""

import json
import os
import sys
import re
import subprocess
import threading
import time as _time
from datetime import datetime
from flask import Flask, request, jsonify
import requests as http_requests
from dotenv import load_dotenv

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(PROJECT_ROOT)
load_dotenv(os.path.join(PROJECT_ROOT, ".env"))

app = Flask(__name__)

# ── 이미지 → 텍스트 (Gemini Vision) ──────────────────────
_gemini_model = None
_store_products = None  # 매장별 품목 캐시

OCR_PROMPT_BASE = """이 사진은 식자재 발주서입니다.

사진이 인쇄된 표 형식(좌측 품목명, 우측 수량 기입란)인지 먼저 판단하세요.

■ 인쇄된 표 형식인 경우:
1단계: 표의 모든 행을 위에서 아래로 순서대로 읽으세요.
2단계: 각 행의 "발주수량" 칸에 손글씨 숫자가 적혀있는지 확인하세요.
3단계: 숫자가 적힌 행만 출력하세요. 빈칸 행은 건너뛰세요.

표 형식 주의사항:
- 각 숫자가 어느 행에 속하는지 세로 위치로 정확히 판별하세요
- 숫자의 세로 위치가 해당 품목행과 일치해야 합니다 (윗행이나 아랫행으로 밀리면 안 됩니다)
- 품목명 옆 괄호 안 텍스트(예: "5kg 단위", "1단 1kg")는 단위 설명이지 수량이 아닙니다
- 단위 설명을 수량에 포함하지 마세요 (깐양파 (5kg 단위)에 "5"가 적혀있으면 → "깐양파 5" 출력)
- 표 아래 여백에 추가 품목이 있으면 반드시 포함하세요

■ 손글씨(자유 형식)인 경우:
- 위에서 아래로 모든 줄을 빠짐없이 읽으세요
- 글씨가 지저분하거나 흐려도 품목 리스트를 참고해서 최대한 해석하세요
- 숫자를 줄 긋고 다시 쓴 경우, 최종 수정된 숫자를 사용하세요
- 낙서/번짐/얼룩이 있어도 그 줄의 품목과 수량은 반드시 읽으세요

■ 공통 출력 형식:
- 한 줄에 하나씩 "품목 수량" (예: 배추 2, 숙주 4, 양파 5)
- 품목명과 수량 사이에 공백 하나, 수량은 숫자만 (단위 생략)
- 품목명은 아래 품목 리스트에서 가장 가까운 것으로 매칭
- 다른 설명 없이 품목 리스트만 출력
- 누락은 업무 사고입니다. 숫자가 적힌 모든 품목을 포함하세요."""

OCR_FALLBACK_PRODUCTS = (
    "배추, 대파, 깻잎, 표고버섯, 느타리, 맛느타리, 꽃느타리, 새송이, 팽이, "
    "쪽파, 숙주, 청경채, 새싹, 무순, 미나리, 돌미나리, 양배추, 귤, 무, 계란, "
    "당근, 양파, 깐양파, 깐마늘, 고추, 감자, 콩나물, 곱슬이콩나물, 부추, "
    "청양고추, 케일, 로메인, 아보카도, 레몬, 오렌지, 딸기, 쑥갓, 근대, "
    "쌀, 알배기, 적채, 씻은김치, 포기김치, 파김치, 냉동알마늘, 맛두부, "
    "애호박, 깐대파, 깐쪽파, 영양부추"
)


def _load_store_products():
    """매장별 품목 리스트 로드 (캐시)"""
    global _store_products
    if _store_products is None:
        path = os.path.join(
            PROJECT_ROOT, "skills", "order_processing", "resources",
            "store_products_clean.json",
        )
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                _store_products = json.load(f)
        else:
            _store_products = {}
    return _store_products


def _get_store_product_hint(store_name):
    """매장에 맞는 품목 힌트 문자열 반환 (common + 해당 지점만)"""
    if not store_name:
        return OCR_FALLBACK_PRODUCTS

    # branch-specific 데이터 (store_products.json)
    sp_path = os.path.join(
        PROJECT_ROOT, "skills", "order_processing", "resources",
        "store_products.json",
    )
    if os.path.exists(sp_path):
        with open(sp_path, "r", encoding="utf-8") as f:
            sp_data = json.load(f)
    else:
        sp_data = {}

    group, branch = _split_store_name(store_name)
    group_data = sp_data.get(group, sp_data.get(store_name, {}))

    if isinstance(group_data, dict):
        # common + branch 품목 합치기
        combined = set(group_data.get("common", []))
        if branch:
            combined.update(group_data.get(branch, []))
        else:
            # 단일 매장 (지점 없음) — common만
            combined.update(group_data.get("common", []))
    elif isinstance(group_data, list):
        combined = set(group_data)
    else:
        return OCR_FALLBACK_PRODUCTS

    if not combined:
        return OCR_FALLBACK_PRODUCTS

    # 대표 품목명만 추출 (변형 통합, 괄호 등급/규격은 유지)
    seen_bases = {}  # base → best variant (괄호 있는 것 우선)
    for p in combined:
        # base: 괄호/숫자 제거한 이름 (중복 판별용)
        base = re.sub(r"[\(\(].*?[\)\)]", "", p)
        base = re.sub(r"\d+[KkGgLl]*$", "", base).strip()
        if not base:
            continue
        # 괄호가 있는 변형을 우선 (등급/규격 정보 보존)
        if base not in seen_bases or ("(" in p or "（" in p):
            seen_bases[base] = p
    clean = set(seen_bases.values())
    return ", ".join(sorted(clean))


def _build_ocr_prompt(store_name=None):
    """매장별 품목 힌트가 포함된 OCR 프롬프트 생성"""
    hint = _get_store_product_hint(store_name)

    prompt = f"{OCR_PROMPT_BASE}\n\n이 매장이 주문 가능한 품목 (이 중에서 매칭):\n{hint}"
    prompt += "\n\n주의: 보통 발주서에는 10~20개 품목이 있습니다. 8개 미만이면 누락을 의심하세요."
    return prompt


def _get_gemini():
    global _gemini_model
    if _gemini_model is None:
        import google.generativeai as genai
        genai.configure(api_key=os.environ.get("GEMINI_API_KEY", ""))
        _gemini_model = genai.GenerativeModel("gemini-2.5-flash")
    return _gemini_model


OCR_RETRY_PROMPT = """이전 인식에서 품목이 누락된 것 같습니다.
사진을 다시 한번 꼼꼼히 읽어주세요.

중요:
- 사진의 첫 줄부터 마지막 줄까지 모든 줄을 읽으세요
- 글씨가 흐리거나 지저분해도 건너뛰지 마세요
- 사진에 보이는 줄 수만큼 출력해야 합니다
- 이전에 {prev_count}개만 읽었는데, 사진에는 더 많은 품목이 있습니다

"""

# 이미지 OCR 최소 품목 수 (이보다 적으면 재시도)
OCR_MIN_ITEMS_RETRY = 5


def ocr_from_url(image_url, store_name=None):
    """이미지 URL → Gemini Vision → 발주 텍스트 추출 (매장별 품목 힌트 포함)

    품목 수가 너무 적으면 자동 재시도 (최대 1회).
    """
    resp = http_requests.get(image_url, timeout=10)
    resp.raise_for_status()
    image_bytes = resp.content

    # MIME type 추정
    content_type = resp.headers.get("Content-Type", "image/jpeg")
    if "png" in content_type:
        mime = "image/png"
    elif "webp" in content_type:
        mime = "image/webp"
    else:
        mime = "image/jpeg"

    image_part = {"mime_type": mime, "data": image_bytes}
    prompt = _build_ocr_prompt(store_name)
    model = _get_gemini()

    # 1차 시도
    response = model.generate_content([prompt, image_part])
    text = response.text.strip()

    # 품목 수 검증 → 너무 적으면 재시도
    line_count = len([l for l in text.split("\n") if l.strip()])
    if line_count < OCR_MIN_ITEMS_RETRY:
        # 재시도 없이 반환 (품목이 원래 적을 수 있음)
        return text

    # 재시도 판단: 이미지에서 기대 줄 수 대비 너무 적은지 확인
    # Heuristic: 손글씨 발주서는 보통 10-20줄, 8줄 미만이면 누락 의심
    if line_count < 8:
        retry_prompt = OCR_RETRY_PROMPT.format(prev_count=line_count) + prompt
        try:
            response2 = model.generate_content([retry_prompt, image_part])
            text2 = response2.text.strip()
            line_count2 = len([l for l in text2.split("\n") if l.strip()])
            # 재시도 결과가 더 많으면 채택
            if line_count2 > line_count:
                return text2
        except Exception:
            pass  # 재시도 실패 시 원본 사용

    return text


# ── 이미지 OCR 백그라운드 결과 캐시 ─────────────────────────
_ocr_results = {}  # {user_id: {"status": "processing"/"done"/"error", ...}}

# ── 영수증 파싱 결과 캐시 ─────────────────────────────────────
_receipt_results = {}  # {user_id: {"status": "processing"/"done"/"error", ...}}

# 영수증 트리거 단어
RECEIPT_TRIGGERS = {"영수증", "매입", "정산"}

# 영수증 URL 도메인 (자동 감지)
RECEIPT_URL_DOMAINS = {"itanet.co.kr", "marketbom.com"}


def _cleanup_caches():
    """2분 이상 된 캐시 엔트리 정리"""
    now = datetime.now().timestamp()
    expired = [k for k, v in _ocr_results.items() if now - v["time"] > 120]
    for k in expired:
        del _ocr_results[k]
    expired_r = [k for k, v in _receipt_results.items() if now - v["time"] > 120]
    for k in expired_r:
        del _receipt_results[k]


# ── 경로 설정 ──────────────────────────────────────────────
DATA_DIR = os.path.join(PROJECT_ROOT, "data", "inputs")
STORE_MAP_PATH = os.path.join(
    PROJECT_ROOT, "skills", "order_processing", "resources", "kakao_user_store_map.json"
)
MAPPINGS_PATH = os.path.join(
    PROJECT_ROOT, "skills", "order_processing", "resources", "mappings.json"
)


def _today_str():
    return datetime.now().strftime("%Y%m%d")


def _queue_path(date_str=None):
    ds = date_str or _today_str()
    return os.path.join(DATA_DIR, f"kakao_orders_{ds}.json")


def _load_queue(date_str=None):
    path = _queue_path(date_str)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"date": date_str or _today_str(), "orders": [], "closed": False}


def _save_queue(queue, date_str=None):
    path = _queue_path(date_str)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(queue, f, ensure_ascii=False, indent=2)


def _load_store_map():
    if os.path.exists(STORE_MAP_PATH):
        with open(STORE_MAP_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def _save_store_map(smap):
    os.makedirs(os.path.dirname(STORE_MAP_PATH), exist_ok=True)
    with open(STORE_MAP_PATH, "w", encoding="utf-8") as f:
        json.dump(smap, f, ensure_ascii=False, indent=2)


# ── 정규화 설정 로드 ──────────────────────────────────────
def _load_normalize_config():
    if os.path.exists(MAPPINGS_PATH):
        with open(MAPPINGS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("kakao_normalize", {})
    return {}


def _split_store_name(store_name):
    """'샤브야키 분당점' → ('샤브야키', '분당점'), '고른햇살' → ('고른햇살', None)"""
    if not store_name:
        return None, None
    parts = store_name.split(maxsplit=1)
    if len(parts) == 2:
        return parts[0], parts[1]
    return parts[0], None


def _eval_calc(expr, groups):
    """{calc:1*10} 계산식 평가. groups[1]=첫 번째 캡처 그룹 값."""
    import math
    m = re.match(r"^(\d+)([+\-*/])(\d+\.?\d*)(?::(\w+))?$", expr)
    if not m:
        return None
    idx = int(m.group(1))
    op = m.group(2)
    operand = float(m.group(3))
    rounding = m.group(4)
    val = float(groups[idx]) if idx < len(groups) else 0
    if op == "*":
        result = val * operand
    elif op == "/":
        result = val / operand if operand != 0 else 0
    elif op == "+":
        result = val + operand
    elif op == "-":
        result = val - operand
    else:
        result = val
    if rounding == "ceil":
        result = math.ceil(result)
    elif rounding == "floor":
        result = math.floor(result)
    elif rounding == "round":
        result = round(result)
    return str(int(result)) if result == int(result) else f"{result:.1f}"


def _apply_replacement(pattern_str, replacement, line):
    """정규식 패턴 매칭 + {1}, {calc:...} 치환. 매칭 실패 시 None."""
    m = re.match(pattern_str, line)
    if not m:
        return None
    groups = [m.group(0)] + list(m.groups())

    def calc_sub(cm):
        val = _eval_calc(cm.group(1), groups)
        return val if val else cm.group(0)

    def ref_sub(rm):
        idx = int(rm.group(1))
        return groups[idx] if idx < len(groups) and groups[idx] else ""

    result = re.sub(r"\{calc:([^}]+)\}", calc_sub, replacement)
    result = re.sub(r"\{(\d+)\}", ref_sub, result)
    return result.strip()


def _try_patterns(line, patterns_list):
    """패턴 리스트를 순서대로 시도, 첫 매칭 결과 반환."""
    for rule in patterns_list:
        result = _apply_replacement(rule["pattern"], rule["replacement"], line)
        if result is not None:
            return result
    return None


# ── 발주 파싱 + 정규화 ───────────────────────────────────
def parse_order_lines(text, store_name=None):
    """발주 텍스트 → 정규화된 [{item, qty, unit, raw, original}, ...]

    store_name이 주어지면 거래처별 규칙을 우선 적용:
    store-specific > group-wide > common
    """
    config = _load_normalize_config()

    # 계층 구조 로드
    common = config.get("common", {})
    common_items = common.get("items", {})
    common_patterns = common.get("patterns", [])
    by_customer = config.get("byCustomer", {})

    default_units = config.get("default_units", {})
    unit_aliases = config.get("unit_aliases", {})
    noise = config.get("noise_patterns", [])
    ocr_fixes = config.get("ocr_fixes", {})

    # 거래처 분리 & 규칙 로드
    group, store = _split_store_name(store_name)
    group_cfg = by_customer.get(group, {}) if group else {}
    group_all = group_cfg.get("all", {})
    store_cfg = group_cfg.get(store, {}) if store else {}

    store_items = store_cfg.get("items", {})
    store_patterns = store_cfg.get("patterns", [])
    group_items = group_all.get("items", {})
    group_patterns = group_all.get("patterns", [])

    # OCR 오인식 전처리 (파싱 전 텍스트 레벨 치환)
    for wrong, correct in sorted(ocr_fixes.items(), key=lambda x: -len(x[0])):
        text = text.replace(wrong, correct)

    # 단위 패턴
    unit_pat = r"박스|팩|kg|KG|Kg|k|K|봉|판|망|통|포|단|개|BOX|Box|box|EA|ea|Ea|단위"

    items = []
    for line in text.strip().split("\n"):
        line = line.strip().rstrip(".")
        if not line or line.startswith("(") or line.startswith("#"):
            continue
        line = re.sub(r"^\d+[\.\)]\s*", "", line)
        if not re.search(r"\d", line) and any(n in line for n in noise):
            continue

        # OCR 수량 깨짐 보정: "품목 X Ykg" → "품목 XYkg" (공백으로 분리된 숫자 합치기)
        line = re.sub(
            r"(\d+)\s+(\d+(?:\.\d+)?)\s*(kg|KG|Kg|k|K|박스|팩|판|망|봉|통|포|단)$",
            r"\1\2\3",
            line,
        )

        original = line

        # ── 1단계: 패턴 매칭 (store → group → common) ──
        pattern_result = (
            _try_patterns(line, store_patterns)
            or _try_patterns(line, group_patterns)
            or _try_patterns(line, common_patterns)
        )

        if pattern_result:
            pm = re.match(
                rf"^(.+?)\s+(\d+(?:\.\d+)?)\s*({unit_pat})?$",
                pattern_result,
            )
            if pm:
                item_name = pm.group(1).strip()
                qty = pm.group(2)
                unit = pm.group(3) or ""
                if unit in unit_aliases:
                    unit = unit_aliases[unit]
                if not unit:
                    unit = default_units.get(item_name, default_units.get("_default", "박스"))
                normalized = f"{item_name} {qty}{unit}"
                items.append({
                    "item": item_name, "qty": qty, "unit": unit,
                    "raw": normalized, "original": original,
                })
            else:
                items.append({
                    "item": pattern_result, "qty": "", "unit": "",
                    "raw": pattern_result, "original": original,
                })
            continue

        # ── 2단계: 일반 파싱 + 계층적 별칭 ──
        m = re.match(rf"^(.+?)\s+(\d+(?:\.\d+)?)\s*({unit_pat})?$", line)
        if not m:
            m = re.match(rf"^([가-힣]+?)(\d+(?:\.\d+)?)\s*({unit_pat})?$", line)

        if m:
            item_name = m.group(1).strip()
            qty = m.group(2)
            unit = m.group(3) or ""

            if unit in unit_aliases:
                unit = unit_aliases[unit]
            if unit.lower() == "k":
                unit = "kg"

            # 별칭: store → group → common (첫 매칭 사용)
            alias = (
                store_items.get(item_name)
                or group_items.get(item_name)
                or common_items.get(item_name)
            )
            if alias:
                item_name = alias

            if not unit:
                unit = default_units.get(item_name, default_units.get("_default", "박스"))

            normalized = f"{item_name} {qty}{unit}"
            items.append({
                "item": item_name, "qty": qty, "unit": unit,
                "raw": normalized, "original": original,
            })
        else:
            if any(n in line for n in noise):
                continue
            items.append({"item": line, "qty": "", "unit": "", "raw": line, "original": line})

    return items


# ── 운영자 카카오톡 알림 ─────────────────────────────────────
def _load_kakao_tokens():
    """kakao_token.json 로드 + access_token 갱신"""
    token_path = os.path.join(PROJECT_ROOT, "kakao_token.json")
    if not os.path.exists(token_path):
        return None
    with open(token_path, "r") as f:
        tokens = json.load(f)
    rest_api_key = os.environ.get("KAKAO_REST_API_KEY")
    if rest_api_key and tokens.get("refresh_token"):
        try:
            res = http_requests.post("https://kauth.kakao.com/oauth/token", data={
                "grant_type": "refresh_token",
                "client_id": rest_api_key,
                "refresh_token": tokens["refresh_token"],
            }, timeout=5)
            new_tokens = res.json()
            if "access_token" in new_tokens:
                tokens["access_token"] = new_tokens["access_token"]
                if "refresh_token" in new_tokens:
                    tokens["refresh_token"] = new_tokens["refresh_token"]
                with open(token_path, "w") as f:
                    json.dump(tokens, f)
        except Exception:
            pass
    return tokens


def _notify_operator(store_name, items, source="text"):
    """발주 접수 시 운영자에게 카카오톡 나에게 보내기"""
    def _send():
        try:
            tokens = _load_kakao_tokens()
            if not tokens or not tokens.get("access_token"):
                return
            item_lines = "\n".join(it["raw"] for it in items)
            now = datetime.now().strftime("%H:%M")
            if source == "ocr":
                label = " (사진)"
            elif source == "modify":
                label = " (수정)"
            else:
                label = ""
            msg = f"[발주 접수 알림]{label}\n\n-{store_name}\n{item_lines}\n\n총 {len(items)}건 ({now})"

            url = "https://kapi.kakao.com/v2/api/talk/memo/default/send"
            headers = {"Authorization": f"Bearer {tokens['access_token']}"}
            payload = {
                "template_object": json.dumps({
                    "object_type": "text",
                    "text": msg,
                    "link": {
                        "web_url": "https://drive.google.com",
                        "mobile_web_url": "https://drive.google.com",
                    },
                    "button_title": "현황 보기",
                })
            }
            http_requests.post(url, headers=headers, data=payload, timeout=5)
        except Exception as e:
            print(f"[알림 실패] {e}")

    threading.Thread(target=_send, daemon=True).start()


# ── 오픈빌더 스킬 응답 빌더 ────────────────────────────────
def _skill_response(text):
    """카카오 오픈빌더 simpleText 응답 규격"""
    return jsonify({
        "version": "2.0",
        "template": {
            "outputs": [{"simpleText": {"text": text}}]
        },
    })


# ── 공통: 텍스트 → 큐 저장 + 응답 ──────────────────────────
def _save_order_and_respond(order_text, user_id, is_ocr=True):
    """파싱된 텍스트 → 큐 저장 → 응답 반환 (텍스트/OCR 공용)"""
    store_map = _load_store_map()
    store_name = store_map.get(user_id)
    if not store_name:
        return _skill_response(
            "등록되지 않은 사용자입니다.\n\n매장등록 [매장명]으로 먼저 등록해 주세요."
        )

    queue = _load_queue()
    if queue.get("closed"):
        return _skill_response(f"오늘({_today_str()}) 발주가 이미 마감되었습니다.")

    items = parse_order_lines(order_text, store_name)
    order_entry = {
        "store": store_name,
        "user_id": user_id,
        "items": items,
        "raw": order_text,
        "received_at": datetime.now().isoformat(),
        "source": "ocr" if is_ocr else "text",
    }
    queue["orders"].append(order_entry)
    _save_queue(queue)

    # 운영자 카톡 알림
    _notify_operator(store_name, items, "ocr" if is_ocr else "text")

    item_lines = "\n".join(it["raw"] for it in items)
    label = " (사진 인식)" if is_ocr else ""
    return _skill_response(
        f"발주 접수 완료!{label}\n\n"
        f"-{store_name}\n{item_lines}\n\n"
        f"총 {len(items)}건"
    )


# ── 엔드포인트 ─────────────────────────────────────────────
def _extract_image_url(body):
    """오픈빌더 요청에서 이미지 URL 추출 (여러 위치 탐색)"""
    # 1) flow.trigger.type == IMAGE_UPLOAD 이면 utterance가 이미지 URL
    trigger_type = body.get("flow", {}).get("trigger", {}).get("type", "")
    if trigger_type == "IMAGE_UPLOAD":
        utt = body.get("userRequest", {}).get("utterance", "").strip()
        if utt.startswith("http"):
            return utt
    # 2) utterance가 카카오 CDN URL인 경우
    utt = body.get("userRequest", {}).get("utterance", "").strip()
    if "kakaocdn.net" in utt and utt.startswith("http"):
        return utt
    # 3) action.params 내 이미지
    params = body.get("action", {}).get("params", {})
    for v in params.values():
        if isinstance(v, str) and v.startswith("http") and "kakaocdn" in v:
            return v
    # 4) action.detailParams 내 이미지
    detail = body.get("action", {}).get("detailParams", {})
    for v in detail.values():
        if isinstance(v, dict):
            origin = v.get("origin", "")
            if isinstance(origin, str) and origin.startswith("http"):
                return origin
    # 5) userRequest.params.media
    media = body.get("userRequest", {}).get("params", {}).get("media", {})
    if isinstance(media, dict) and media.get("url"):
        return media["url"]
    return None


@app.route("/kakao/order", methods=["POST"])
def receive_order():
    """오픈빌더 스킬 웹훅 — 발주 접수
    텍스트: 즉시 처리
    이미지: 백그라운드 OCR → 다음 메시지 시 결과 전달
    """
    body = request.get_json(silent=True) or {}

    # 디버그: 전체 요청 로깅
    debug_path = os.path.join(PROJECT_ROOT, "data", "inputs", "kakao_debug_log.jsonl")
    with open(debug_path, "a", encoding="utf-8") as f:
        f.write(json.dumps(body, ensure_ascii=False) + "\n")

    user_req = body.get("userRequest", {})
    utterance = user_req.get("utterance", "").strip()
    user_id = user_req.get("user", {}).get("id", "unknown")

    _cleanup_caches()

    # ── 1) 대기 중인 OCR 결과 확인 ──
    if user_id in _ocr_results:
        result = _ocr_results.get(user_id)
        if result and result["status"] == "done" and result.get("text"):
            _ocr_results.pop(user_id, None)
            return _save_order_and_respond(result["text"], user_id, is_ocr=True)
        elif result and result["status"] == "error":
            err = result.get("error", "알 수 없는 오류")
            _ocr_results.pop(user_id, None)
            return _skill_response(
                f"이미지 인식 실패: {err}\n\n텍스트로 다시 입력해 주세요."
            )
        elif result and result["status"] == "processing":
            return _skill_response("아직 인식 중... 잠시 후 다시 보내주세요.")

    # ── 2) 이미지 감지 → 백그라운드 OCR ──
    image_url = _extract_image_url(body)
    if image_url:
        # 매장 조회 (OCR 품목 힌트용)
        store_map = _load_store_map()
        ocr_store = store_map.get(user_id)

        _ocr_results[user_id] = {
            "status": "processing",
            "time": datetime.now().timestamp(),
        }

        def _run_ocr():
            try:
                text = ocr_from_url(image_url, store_name=ocr_store)
                _ocr_results[user_id] = {
                    "status": "done",
                    "text": text.strip() if text else "",
                    "time": datetime.now().timestamp(),
                }
            except Exception as e:
                _ocr_results[user_id] = {
                    "status": "error",
                    "error": str(e),
                    "time": datetime.now().timestamp(),
                }

        threading.Thread(target=_run_ocr, daemon=True).start()

        return _skill_response("사진 접수! 인식 중...\n\n10초 후 '확인' 보내주세요.")

    # ── 3) 텍스트 발주 처리 ──
    if not utterance:
        return _skill_response("발주 내용이 비어 있습니다. 다시 입력해 주세요.")

    # 매장 매핑 확인
    store_map = _load_store_map()
    store_name = store_map.get(user_id)

    if not store_name:
        reg_match = re.match(r"^매장등록\s+(.+)$", utterance)
        if reg_match:
            new_store = reg_match.group(1).strip()
            store_map[user_id] = new_store
            _save_store_map(store_map)
            return _skill_response(
                f"매장 등록 완료!\n\n등록된 매장명: {new_store}\n\n이제 발주를 입력해 주세요."
            )
        return _skill_response(
            "등록되지 않은 사용자입니다.\n\n"
            '처음 이용 시 아래와 같이 매장명을 등록해 주세요:\n\n'
            '매장등록 샤브야키 도봉점'
        )

    queue = _load_queue()
    if queue.get("closed"):
        return _skill_response(
            f"오늘({_today_str()}) 발주가 이미 마감되었습니다.\n"
            "추가 발주가 필요하면 운영자에게 연락해 주세요."
        )

    items = parse_order_lines(utterance, store_name)
    order_entry = {
        "store": store_name,
        "user_id": user_id,
        "items": items,
        "raw": utterance,
        "received_at": datetime.now().isoformat(),
        "source": "text",
    }
    queue["orders"].append(order_entry)
    _save_queue(queue)

    # 운영자 카톡 알림
    _notify_operator(store_name, items, "text")

    item_lines = "\n".join(it["raw"] for it in items)
    return _skill_response(
        f"발주 접수 완료!\n\n"
        f"-{store_name}\n{item_lines}\n\n"
        f"총 {len(items)}건"
    )


@app.route("/kakao/modify", methods=["POST"])
def modify_order():
    """오픈빌더 스킬 — 발주 수정 (해당 매장의 마지막 발주를 교체)"""
    body = request.get_json(silent=True) or {}
    user_req = body.get("userRequest", {})
    utterance = user_req.get("utterance", "").strip()
    user_id = user_req.get("user", {}).get("id", "unknown")

    store_map = _load_store_map()
    store_name = store_map.get(user_id)
    if not store_name:
        return _skill_response("등록되지 않은 사용자입니다.")

    queue = _load_queue()
    if queue.get("closed"):
        return _skill_response("오늘 발주가 이미 마감되어 수정할 수 없습니다.")

    # 해당 매장의 마지막 발주 삭제
    removed = False
    for i in range(len(queue["orders"]) - 1, -1, -1):
        if queue["orders"][i]["store"] == store_name:
            queue["orders"].pop(i)
            removed = True
            break

    # 새 발주 추가
    items = parse_order_lines(utterance, store_name)
    queue["orders"].append({
        "store": store_name,
        "user_id": user_id,
        "items": items,
        "raw": utterance,
        "received_at": datetime.now().isoformat(),
    })
    _save_queue(queue)

    # 운영자 카톡 알림
    _notify_operator(store_name, items, "modify")

    item_lines = "\n".join(it["raw"] for it in items)
    prefix = "발주 수정 완료!" if removed else "기존 발주 없음 -- 신규 접수!"
    return _skill_response(
        f"{prefix}\n\n-{store_name}\n{item_lines}\n\n총 {len(items)}건"
    )


@app.route("/kakao/cancel", methods=["POST"])
def cancel_order():
    """오픈빌더 스킬 — 발주 취소"""
    body = request.get_json(silent=True) or {}
    user_id = body.get("userRequest", {}).get("user", {}).get("id", "unknown")

    store_map = _load_store_map()
    store_name = store_map.get(user_id)
    if not store_name:
        return _skill_response("등록되지 않은 사용자입니다.")

    queue = _load_queue()
    if queue.get("closed"):
        return _skill_response("오늘 발주가 이미 마감되어 취소할 수 없습니다.")

    new_orders = []
    removed_count = 0
    for o in queue["orders"]:
        if o["store"] == store_name:
            removed_count += 1
        else:
            new_orders.append(o)
    queue["orders"] = new_orders
    _save_queue(queue)

    if removed_count:
        return _skill_response(f"[{store_name}] 발주 {removed_count}건 취소되었습니다.")
    return _skill_response(f"[{store_name}] 취소할 발주가 없습니다.")


@app.route("/status", methods=["GET"])
def status():
    """운영자용 — 현재 접수 현황 JSON"""
    queue = _load_queue()
    stores = {}
    for o in queue["orders"]:
        stores.setdefault(o["store"], []).append({
            "items": o["items"],
            "received_at": o["received_at"],
        })
    return jsonify({
        "date": queue["date"],
        "closed": queue.get("closed", False),
        "store_count": len(stores),
        "total_orders": len(queue["orders"]),
        "stores": stores,
    })


@app.route("/trigger", methods=["POST"])
def trigger_pipeline():
    """운영자용 — 마감 + 파이프라인 실행"""
    queue = _load_queue()
    if not queue["orders"]:
        return jsonify({"ok": False, "error": "접수된 발주가 없습니다."}), 400

    # 마감 플래그
    queue["closed"] = True
    _save_queue(queue)

    # orders_today.txt 생성
    txt = export_queue_to_txt(queue)
    today = _today_str()
    orders_path = os.path.join(DATA_DIR, "orders_today.txt")
    archive_path = os.path.join(DATA_DIR, f"orders_{today}.txt")
    with open(orders_path, "w", encoding="utf-8") as f:
        f.write(txt)
    with open(archive_path, "w", encoding="utf-8") as f:
        f.write(txt)

    # 파이프라인을 별도 스레드에서 실행 (서버 블로킹 방지)
    date_formatted = f"{today[:4]}-{today[4:6]}-{today[6:]}"

    def _run():
        script = os.path.join(PROJECT_ROOT, "execution", "run_full_workflow.py")
        subprocess.run(
            [sys.executable, script, "--date", date_formatted, "--input", orders_path],
            cwd=PROJECT_ROOT,
        )

    t = threading.Thread(target=_run, daemon=True)
    t.start()

    return jsonify({
        "ok": True,
        "message": f"마감 완료. 파이프라인 실행 중 ({len(queue['orders'])}건, {date_formatted})",
        "orders_file": orders_path,
    })


@app.route("/reset", methods=["POST"])
def reset_queue():
    """운영자용 — 오늘 큐 초기화"""
    queue = {"date": _today_str(), "orders": [], "closed": False}
    _save_queue(queue)
    return jsonify({"ok": True, "message": "오늘 발주 큐 초기화 완료."})


@app.route("/register", methods=["POST"])
def register_store():
    """운영자용 — 매장 등록 (user_id → store_name)"""
    data = request.get_json(silent=True) or {}
    user_id = data.get("user_id", "").strip()
    store_name = data.get("store_name", "").strip()
    if not user_id or not store_name:
        return jsonify({"ok": False, "error": "user_id, store_name 필수"}), 400

    smap = _load_store_map()
    smap[user_id] = store_name
    _save_store_map(smap)
    return jsonify({"ok": True, "user_id": user_id, "store_name": store_name})


# ── 영수증 처리 ──────────────────────────────────────────────

def _is_receipt_url(text):
    """텍스트가 영수증 URL인지 확인."""
    if not text.startswith("http"):
        return False
    for domain in RECEIPT_URL_DOMAINS:
        if domain in text:
            return True
    return False


def _is_receipt_trigger(utterance):
    """발화에 영수증 트리거 키워드가 있는지 확인."""
    return any(t in utterance for t in RECEIPT_TRIGGERS)


def _notify_receipt_result(supplier, result_summary, warnings=None):
    """영수증 처리 결과를 운영자에게 카카오톡 알림."""
    def _send():
        try:
            tokens = _load_kakao_tokens()
            if not tokens or not tokens.get("access_token"):
                return
            now = datetime.now().strftime("%H:%M")
            msg = f"[영수증 입력 완료]\n\n공급처: {supplier}\n{result_summary}\n({now})"
            if warnings:
                msg += f"\n\n⚠ 경고: {'; '.join(warnings[:3])}"

            url = "https://kapi.kakao.com/v2/api/talk/memo/default/send"
            headers = {"Authorization": f"Bearer {tokens['access_token']}"}
            payload = {
                "template_object": json.dumps({
                    "object_type": "text",
                    "text": msg,
                    "link": {
                        "web_url": "https://drive.google.com",
                        "mobile_web_url": "https://drive.google.com",
                    },
                    "button_title": "현황 보기",
                })
            }
            http_requests.post(url, headers=headers, data=payload, timeout=5)
        except Exception as e:
            print(f"[영수증 알림 실패] {e}")

    threading.Thread(target=_send, daemon=True).start()


def _process_receipt_background(user_id, url=None, image_url=None, text=None,
                                supplier_hint=None):
    """백그라운드에서 영수증 파싱 + 단가시트 입력."""
    def _run():
        try:
            from parse_receipt import parse_receipt, map_to_danga_items
            from fill_receipt_prices import fill_receipt_prices

            # 1) 파싱
            result = parse_receipt(
                url=url, image_url=image_url, text=text,
                supplier_hint=supplier_hint,
            )

            if 'error' in result:
                _receipt_results[user_id] = {
                    "status": "error",
                    "error": result['error'],
                    "time": datetime.now().timestamp(),
                }
                return

            supplier = result.get('supplier', supplier_hint or '?')
            items = result.get('items', [])

            if not items:
                _receipt_results[user_id] = {
                    "status": "error",
                    "error": "영수증에서 품목을 찾지 못했습니다",
                    "time": datetime.now().timestamp(),
                }
                return

            # 2) 매핑 결과 미리보기 생성
            mapped = map_to_danga_items(result, supplier)
            preview_lines = []
            for m in mapped:
                if m.get('skip_reason'):
                    preview_lines.append(f"  {m['receipt_name']}: {m['skip_reason']}")
                else:
                    preview_lines.append(
                        f"  {m['receipt_name']} → {m['danga_name']}: {m['unit_price']:,}원"
                    )

            # 3) 단가시트 입력
            master_path = os.environ.get("MASTER_FILE_PATH", "")
            today = datetime.now().strftime("%Y-%m-%d")

            fill_result = None
            if master_path and os.path.exists(master_path):
                try:
                    fill_result = fill_receipt_prices(
                        master_file=master_path,
                        target_date=today,
                        receipts=[result],
                        dry_run=False,
                    )
                except Exception as e:
                    print(f"[영수증] 단가시트 입력 오류: {e}")
                    fill_result = {"updates": 0, "error": str(e)}

            # 4) 결과 저장
            summary = f"{len(items)}건 파싱"
            if fill_result:
                summary += f" / {fill_result.get('updates', 0)}건 입력"

            _receipt_results[user_id] = {
                "status": "done",
                "supplier": supplier,
                "items_count": len(items),
                "preview": "\n".join(preview_lines),
                "fill_result": fill_result,
                "summary": summary,
                "time": datetime.now().timestamp(),
            }

            # 운영자 알림
            _notify_receipt_result(
                supplier, summary,
                fill_result.get('warnings') if fill_result else None,
            )

        except Exception as e:
            _receipt_results[user_id] = {
                "status": "error",
                "error": str(e),
                "time": datetime.now().timestamp(),
            }

    threading.Thread(target=_run, daemon=True).start()


@app.route("/kakao/receipt", methods=["POST"])
def receive_receipt():
    """오픈빌더 스킬 웹훅 — 영수증 접수

    영수증 URL/이미지/텍스트 → 파싱 → 단가시트 매입가 입력
    """
    body = request.get_json(silent=True) or {}
    user_req = body.get("userRequest", {})
    utterance = user_req.get("utterance", "").strip()
    user_id = user_req.get("user", {}).get("id", "unknown")

    _cleanup_caches()

    # 1) 대기 중인 영수증 결과 확인
    if user_id in _receipt_results:
        result = _receipt_results.get(user_id)
        if result and result["status"] == "done":
            _receipt_results.pop(user_id, None)
            supplier = result.get('supplier', '?')
            summary = result.get('summary', '')
            preview = result.get('preview', '')
            fill = result.get('fill_result')
            warn_text = ""
            if fill and fill.get('warnings'):
                warn_text = "\n\n⚠ 경고:\n" + "\n".join(fill['warnings'][:3])
            return _skill_response(
                f"[영수증 처리 완료]\n\n"
                f"공급처: {supplier}\n{summary}\n\n"
                f"{preview}{warn_text}"
            )
        elif result and result["status"] == "error":
            err = result.get("error", "알 수 없는 오류")
            _receipt_results.pop(user_id, None)
            return _skill_response(f"영수증 처리 실패: {err}")
        elif result and result["status"] == "processing":
            return _skill_response("영수증 처리 중... 잠시 후 다시 보내주세요.")

    # 2) 영수증 URL 감지
    receipt_url = None
    for word in utterance.split():
        if _is_receipt_url(word):
            receipt_url = word
            break

    if receipt_url:
        _receipt_results[user_id] = {
            "status": "processing",
            "time": datetime.now().timestamp(),
        }
        _process_receipt_background(user_id, url=receipt_url)
        return _skill_response("영수증 URL 접수! 처리 중...\n\n10초 후 '확인' 보내주세요.")

    # 3) 이미지 감지
    image_url = _extract_image_url(body)
    if image_url:
        # 공급처 힌트 추출 (utterance에서)
        supplier_hint = None
        from parse_receipt import detect_supplier_from_text
        supplier_hint = detect_supplier_from_text(utterance)

        _receipt_results[user_id] = {
            "status": "processing",
            "time": datetime.now().timestamp(),
        }
        _process_receipt_background(user_id, image_url=image_url,
                                    supplier_hint=supplier_hint)
        return _skill_response(
            "영수증 사진 접수! 처리 중...\n\n"
            "10초 후 '확인' 보내주세요."
        )

    # 4) 텍스트 영수증
    if utterance and not _is_receipt_trigger(utterance):
        # "영수증" 키워드 없이 텍스트만 온 경우 → 텍스트 영수증으로 처리
        supplier_hint = None
        from parse_receipt import detect_supplier_from_text
        supplier_hint = detect_supplier_from_text(utterance)

        if supplier_hint:
            _receipt_results[user_id] = {
                "status": "processing",
                "time": datetime.now().timestamp(),
            }
            _process_receipt_background(user_id, text=utterance,
                                        supplier_hint=supplier_hint)
            return _skill_response("영수증 텍스트 접수! 처리 중...")

    return _skill_response(
        "영수증을 전달해 주세요.\n\n"
        "- URL 붙여넣기 (itanet/marketbom)\n"
        "- 영수증 사진 전송\n"
        "- 텍스트 직접 입력"
    )


@app.route("/receipt/parse", methods=["POST"])
def api_parse_receipt():
    """REST API — 영수증 파싱만 (단가시트 입력 없이)."""
    data = request.get_json(silent=True) or {}
    from parse_receipt import parse_receipt, map_to_danga_items

    result = parse_receipt(
        url=data.get('url'),
        image_url=data.get('image_url'),
        text=data.get('text'),
        supplier_hint=data.get('supplier'),
    )

    if 'error' not in result:
        result['mapped_items'] = map_to_danga_items(result)

    return jsonify(result)


@app.route("/receipt/fill", methods=["POST"])
def api_fill_receipt():
    """REST API — 영수증 파싱 + 단가시트 입력."""
    data = request.get_json(silent=True) or {}
    from parse_receipt import parse_receipt
    from fill_receipt_prices import fill_receipt_prices

    master_path = data.get('master') or os.environ.get("MASTER_FILE_PATH", "")
    target_date = data.get('date') or datetime.now().strftime("%Y-%m-%d")
    dry_run = data.get('dry_run', False)

    result = parse_receipt(
        url=data.get('url'),
        image_url=data.get('image_url'),
        text=data.get('text'),
        supplier_hint=data.get('supplier'),
    )

    if 'error' in result:
        return jsonify(result), 400

    fill_result = fill_receipt_prices(
        master_file=master_path,
        target_date=target_date,
        receipts=[result],
        dry_run=dry_run,
    )

    return jsonify({
        "parse_result": result,
        "fill_result": fill_result,
    })


# ── 큐 → orders_today.txt 변환 ────────────────────────────
def export_queue_to_txt(queue):
    """
    큐 JSON을 기존 orders_today.txt 형식으로 변환.
    동일 매장 주문이 여러 건이면 병합한다.
    """
    # 매장별 병합
    merged = {}
    for o in queue["orders"]:
        store = o["store"]
        if store not in merged:
            merged[store] = []
        merged[store].append(o)

    lines = []
    for store, entries in merged.items():
        lines.append(f"-{store}")
        for entry in entries:
            for it in entry["items"]:
                lines.append(it["raw"])
        lines.append("")  # 빈 줄 구분

    return "\n".join(lines)


# ── 메인 ───────────────────────────────────────────────────
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="카카오 발주 웹훅 서버")
    parser.add_argument("--port", type=int, default=5050)
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--host", default="0.0.0.0")
    args = parser.parse_args()

    # Gemini Vision 모델 미리 초기화
    print("Initializing Gemini Vision...")
    _get_gemini()
    print("Gemini Vision ready.")

    print(f"kakao order webhook server: http://{args.host}:{args.port}")
    print(f"  POST /kakao/order   - order receive (text + image OCR)")
    print(f"  POST /kakao/modify  - order modify")
    print(f"  POST /kakao/cancel  - order cancel")
    print(f"  POST /kakao/receipt - receipt receive (URL/image/text)")
    print(f"  GET  /status        - status")
    print(f"  POST /trigger       - close + pipeline")
    print(f"  POST /reset         - reset queue")
    print(f"  POST /register      - register store")
    print(f"  POST /receipt/parse - REST API: parse receipt only")
    print(f"  POST /receipt/fill  - REST API: parse + fill danga sheet")
    app.run(host=args.host, port=args.port, debug=args.debug)
