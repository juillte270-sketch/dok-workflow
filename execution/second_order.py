"""
2차 발주 메시지 자동화

공급처별 발주 메시지 템플릿에 수량을 채워서 카카오톡(나에게 보내기)으로 전송한다.
에이전트가 수량 JSON을 생성 → 이 스크립트가 메시지를 포맷 → 카카오톡 전송.

Usage:
  python execution/second_order.py --input data/inputs/second_order.json --send
  python execution/second_order.py --input data/inputs/second_order.json --dry-run
"""
import argparse
import json
import os
import sys
import time
from collections import OrderedDict
from pathlib import Path

import requests
from dotenv import load_dotenv

load_dotenv()
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

ROOT = Path(__file__).resolve().parent.parent
KAKAO_TOKEN_PATH = ROOT / 'kakao_token.json'


# ═══════════════════════════════════════════════════════════════
# 공급처 메시지 템플릿
# ═══════════════════════════════════════════════════════════════
# 각 item: (이름, prefix, colon_qty)
#   colon_qty=True  → "{prefix}{name} : {qty}"
#   colon_qty=False → "{prefix}{name}" (수량 없이 품목만 나열)
#
# 항목이 입력 JSON에 없으면 해당 줄 생략.
# 공급처에 항목이 하나도 없으면 해당 공급처 메시지 자체를 생략.
# ═══════════════════════════════════════════════════════════════

SUPPLIERS = OrderedDict([
    ('동원', {
        'header': None,
        'items': [
            ('아카시아청', ' - ', True),
            ('마카로니', ' - ', True),
        ],
    }),
    ('이모유통', {
        'header': '금일 2차발주입니다(지하2층i05기둥0457차량)',
        'items': [
            ('라임', '  *', True),
            ('바나나 1발', '  *', True),
            ('방토 2호 500g', '  *', True),
            ('수박(5kg 내외) 1통', '  *', True),
            ('사과(4다이 7개입)', '  *', True),
            ('파인애플(6수/개)', '  *', True),
        ],
    }),
    ('오복상회', {
        'header': '금일 2차발주입니다(지하2층i05기둥0457차량)',
        'items': [
            ('고사리', ' - ', True),
            ('숙주', ' - ', True),
            ('곱슬이 콩나물', ' - ', True),
        ],
    }),
    ('정운', {
        'header': '-금일 발주입니다',
        'items': [
            ('맛김치', '', True),
        ],
    }),
    ('세연농산', {
        'header': '금일 발주입니다',
        'items': [
            ('미나리(대)', ' - ', True),
            ('돌미나리(4KG)', ' - ', True),
            ('깐쪽파', ' - ', True),
            ('부추', ' - ', True),
        ],
    }),
    ('정복상회', {
        'display': '정복상회(랑희)',
        'header': '금일 발주입니다',
        'items': [
            ('무순50', '-', True),
            ('베이비50', '-', True),
            ('베이비250', '-', True),
            ('베이비500', '-', True),
        ],
    }),
    ('현진상회', {
        'header': '금일 발주입니다.',
        'items': [
            ('쑥갓', '-', False),
            ('치커리', '-', False),
            ('깻잎', '-', False),
            ('청상추', '-', False),
        ],
    }),
    ('경북상회', {
        'header': '금일 발주입니다.',
        'items': [
            ('청피망(상)300G', '-', False),
            ('가지', '-', False),
            ('쥬키니', '-', False),
        ],
    }),
    ('카드결제', {
        'header': None,
        'items': [
            ('애플민트', '-', False),
            ('루꼴라', '-', False),
        ],
    }),
])

# 품목명 약칭 → 정식 이름 매핑 (사용자가 축약형으로 입력해도 매칭)
ITEM_ALIASES = {
    '바나나': '바나나 1발',
    '방토': '방토 2호 500g',
    '수박': '수박(5kg 내외) 1통',
    '사과': '사과(4다이 7개입)',
    '파인애플': '파인애플(6수/개)',
    '곱슬이': '곱슬이 콩나물',
    '미나리': '미나리(대)',
    '돌미나리': '돌미나리(4KG)',
    '무순': '무순50',
    '청피망': '청피망(상)300G',
}


# ═══════════════════════════════════════════════════════════════
# 메시지 포맷
# ═══════════════════════════════════════════════════════════════

def normalize_item_name(name):
    """약칭을 정식 품목명으로 변환"""
    return ITEM_ALIASES.get(name, name)


def format_message(supplier_key, qty_map):
    """단일 공급처의 발주 메시지 생성.

    Args:
        supplier_key: SUPPLIERS 딕셔너리의 키
        qty_map: {"품목명": "수량"} 또는 {"품목명": true}

    Returns:
        (display_name, message_body) 또는 (None, None) 항목 없을 때
    """
    config = SUPPLIERS.get(supplier_key)
    if not config:
        return None, None

    display = config.get('display', supplier_key)
    lines = []

    if config.get('header'):
        lines.append(config['header'])

    item_count = 0
    for name, prefix, colon_qty in config['items']:
        # 약칭 매핑도 체크
        val = qty_map.get(name)
        if val is None:
            # alias로도 못 찾으면 skip
            for alias, full in ITEM_ALIASES.items():
                if full == name and alias in qty_map:
                    val = qty_map[alias]
                    break
        if val is None:
            continue

        item_count += 1

        if colon_qty:
            # "prefix + name + : + qty" 형식
            qty_str = '' if val is True else str(val)
            lines.append(f"{prefix}{name} : {qty_str}")
        else:
            # 품목명만 나열 (수량 있으면 뒤에 붙임)
            if val is True or val == '' or val is None:
                lines.append(f"{prefix}{name}")
            else:
                lines.append(f"{prefix}{name} {val}")

    if item_count == 0:
        return None, None

    return display, '\n'.join(lines)


def format_all_messages(quantities):
    """모든 공급처 메시지 생성.

    Args:
        quantities: {"공급처": {"품목": "수량", ...}, ...}

    Returns:
        list of (display_name, message_body)
    """
    messages = []
    for key in SUPPLIERS:
        if key not in quantities:
            continue
        name, body = format_message(key, quantities[key])
        if name and body:
            messages.append((name, body))
    return messages


# ═══════════════════════════════════════════════════════════════
# 카카오톡 전송
# ═══════════════════════════════════════════════════════════════

def load_and_refresh_token():
    """카카오 토큰 로드 및 갱신"""
    if not KAKAO_TOKEN_PATH.exists():
        print("ERROR: kakao_token.json 없음. kakao_auth.py 먼저 실행하세요.", flush=True)
        return None

    with open(KAKAO_TOKEN_PATH, 'r') as f:
        tokens = json.load(f)

    api_key = os.getenv('KAKAO_REST_API_KEY')
    if api_key and tokens.get('refresh_token'):
        try:
            res = requests.post("https://kauth.kakao.com/oauth/token", data={
                "grant_type": "refresh_token",
                "client_id": api_key,
                "refresh_token": tokens['refresh_token'],
            })
            new = res.json()
            if 'access_token' in new:
                tokens['access_token'] = new['access_token']
                if 'refresh_token' in new:
                    tokens['refresh_token'] = new['refresh_token']
                with open(KAKAO_TOKEN_PATH, 'w') as f:
                    json.dump(tokens, f)
        except Exception as e:
            print(f"토큰 갱신 경고: {e}", flush=True)

    return tokens


def send_kakao(text, tokens):
    """카카오톡 나에게 보내기 (단일 메시지)"""
    url = "https://kapi.kakao.com/v2/api/talk/memo/default/send"
    headers = {"Authorization": f"Bearer {tokens['access_token']}"}
    payload = {
        "template_object": json.dumps({
            "object_type": "text",
            "text": text,
            "link": {
                "web_url": "https://drive.google.com",
                "mobile_web_url": "https://drive.google.com",
            },
            "button_title": "Drive 열기",
        }, ensure_ascii=False),
    }
    res = requests.post(url, headers=headers, data=payload)
    if res.status_code == 401:
        print("  토큰 만료. kakao_auth.py 재실행 필요.", flush=True)
        return False
    return res.json().get('result_code') == 0


# ═══════════════════════════════════════════════════════════════
# 메인
# ═══════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description='2차 발주 메시지 생성 및 카카오톡 전송')
    parser.add_argument('--input', required=True, help='수량 JSON 파일 경로')
    parser.add_argument('--send', action='store_true', help='카카오톡 전송')
    parser.add_argument('--dry-run', action='store_true', help='메시지만 출력 (전송 안 함)')
    args = parser.parse_args()

    with open(args.input, 'r', encoding='utf-8') as f:
        quantities = json.load(f)

    messages = format_all_messages(quantities)

    if not messages:
        print("발주할 항목이 없습니다.", flush=True)
        return

    # --- 메시지 출력 ---
    print("=" * 55, flush=True)
    print(" 2차 발주 메시지 미리보기", flush=True)
    print("=" * 55, flush=True)

    for name, body in messages:
        print(f"\n[{name}]", flush=True)
        print(body, flush=True)

    print(f"\n{'=' * 55}", flush=True)
    print(f"총 {len(messages)}개 공급처 메시지", flush=True)

    # --- 카카오톡 전송 ---
    if args.send and not args.dry_run:
        print("\n카카오톡 전송 중...", flush=True)
        tokens = load_and_refresh_token()
        if not tokens:
            return

        ok = 0
        for name, body in messages:
            msg = f"[{name}]\n{body}"
            if send_kakao(msg, tokens):
                print(f"  [{name}] 전송 완료", flush=True)
                ok += 1
            else:
                print(f"  [{name}] 전송 실패", flush=True)
            time.sleep(0.5)

        print(f"\n전송 결과: {ok}/{len(messages)} 성공", flush=True)
        if ok > 0:
            from _notify import notify
            notify("stage4_second")
    elif args.dry_run:
        print("\n(dry-run: 전송 안 함)", flush=True)
    else:
        print("\n(--send 플래그 없음: 전송 안 함)", flush=True)


if __name__ == '__main__':
    main()
