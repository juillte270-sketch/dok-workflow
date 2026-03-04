"""
일일 워크플로우 결과 요약 → 카카오톡 전송

사용법:
    python execution/send_daily_summary.py --date "2026-02-15"
    python execution/send_daily_summary.py --date "2026-02-15" --results '{"stage1":true,"stage2":true}'
"""
import os
import sys
import json
import time
import argparse
from datetime import datetime
import requests
from dotenv import load_dotenv

load_dotenv()

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load_and_refresh_token():
    """카카오 토큰 로드 및 필요 시 갱신"""
    token_path = os.path.join(PROJECT_ROOT, 'kakao_token.json')
    if not os.path.exists(token_path):
        print("Error: kakao_token.json not found. Run kakao_auth.py first.")
        return None

    with open(token_path, 'r') as f:
        tokens = json.load(f)

    rest_api_key = os.getenv('KAKAO_REST_API_KEY')
    if rest_api_key and tokens.get('refresh_token'):
        try:
            res = requests.post("https://kauth.kakao.com/oauth/token", data={
                "grant_type": "refresh_token",
                "client_id": rest_api_key,
                "refresh_token": tokens['refresh_token']
            })
            new_tokens = res.json()
            if 'access_token' in new_tokens:
                tokens['access_token'] = new_tokens['access_token']
                if 'refresh_token' in new_tokens:
                    tokens['refresh_token'] = new_tokens['refresh_token']
                with open(token_path, 'w') as f:
                    json.dump(tokens, f)
                print("Token refreshed.")
        except Exception as e:
            print(f"Token refresh failed (using existing): {e}")

    return tokens


def send_kakao_message(text, tokens):
    """카카오톡 나에게 보내기 (긴 메시지 자동 분할)"""
    url = "https://kapi.kakao.com/v2/api/talk/memo/default/send"
    headers = {"Authorization": f"Bearer {tokens['access_token']}"}

    chunks = _split_text(text, max_size=800)
    print(f"Sending {len(chunks)} message(s)...")

    for i, chunk in enumerate(chunks):
        footer = f"\n({i+1}/{len(chunks)})" if len(chunks) > 1 else ""
        payload = {
            "template_object": json.dumps({
                "object_type": "text",
                "text": chunk + footer,
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

        if res.json().get('result_code') == 0:
            print(f"  Message {i+1}/{len(chunks)} sent.")
        else:
            print(f"  Failed: {res.json()}")
            return False

        if i < len(chunks) - 1:
            time.sleep(0.5)

    return True


def _split_text(text, max_size=800):
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


def collect_results(date_str):
    """날짜 기반 워크플로우 출력 파일 스캔"""
    date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    ymd = date_obj.strftime("%Y%m%d")

    outputs = os.path.join(PROJECT_ROOT, "data", "outputs")
    inputs = os.path.join(PROJECT_ROOT, "data", "inputs")

    r = {}

    # 배송리스트
    r["delivery_list"] = os.path.exists(os.path.join(outputs, f"배송리스트_{ymd}.docx"))

    # 처리된 발주
    r["processed_orders"] = os.path.exists(os.path.join(inputs, f"processed_orders_{ymd}.txt"))

    # 가명세서
    inv_dir = os.path.join(outputs, f"invoices_{ymd}")
    if os.path.exists(inv_dir):
        inv_files = [f for f in os.listdir(inv_dir) if f.endswith('.xlsx')]
        r["invoices"] = True
        r["invoice_count"] = len(inv_files)
        r["invoice_stores"] = [
            f.replace(f"{ymd}_", "").replace("_거래명세서.xlsx", "")
            for f in sorted(inv_files)
        ]
    else:
        r["invoices"] = False
        r["invoice_count"] = 0
        r["invoice_stores"] = []

    # ZIP
    r["invoice_zip"] = os.path.exists(os.path.join(outputs, f"거래명세서_{ymd}.zip"))

    # 최종 명세서
    final_dir = os.path.join(outputs, f"final_invoices_{ymd}")
    if os.path.exists(final_dir):
        final_files = [f for f in os.listdir(final_dir) if f.endswith(('.xlsx', '.pdf'))]
        r["final_invoices"] = True
        r["final_count"] = len(final_files)
    else:
        r["final_invoices"] = False
        r["final_count"] = 0

    # 발주 원본
    r["orders_input"] = os.path.exists(os.path.join(inputs, f"orders_{ymd}.txt"))

    return r


def build_message(date_str, file_results, stage_results=None):
    """카카오톡용 요약 메시지 작성"""
    date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    weekdays = ['월', '화', '수', '목', '금', '토', '일']
    wd = weekdays[date_obj.weekday()]

    lines = []
    lines.append(f"[도크 일일업무 리포트]")
    lines.append(f"{date_obj.month}/{date_obj.day}({wd})")
    lines.append("")

    # 워크플로우 단계별 결과
    if stage_results:
        lines.append("=== 실행 결과 ===")
        stage_map = [
            ("stage1", "1. 발주처리/배송리스트"),
            ("stage2", "2. 발주시트 입력"),
            ("stage3_1", "3-1. 단가시트"),
            ("stage3_2", "3-2. 네이버시세"),
            ("stage3_3", "3-3. 가명세서"),
        ]
        for key, label in stage_map:
            val = stage_results.get(key)
            if val is True:
                lines.append(f"  {label} .. OK")
            elif val is False:
                lines.append(f"  {label} .. FAIL")
        lines.append("")

    # 생성 파일 현황
    lines.append("=== 생성 파일 ===")

    if file_results.get("delivery_list"):
        lines.append("  배송리스트 .. OK")
    else:
        lines.append("  배송리스트 .. -")

    if file_results.get("processed_orders"):
        lines.append("  발주데이터 .. OK")
    else:
        lines.append("  발주데이터 .. -")

    if file_results.get("invoices"):
        cnt = file_results.get("invoice_count", 0)
        lines.append(f"  가명세서 .. {cnt}개 매장")
    else:
        lines.append("  가명세서 .. -")

    if file_results.get("invoice_zip"):
        lines.append("  ZIP파일 .. OK")

    if file_results.get("final_invoices"):
        cnt = file_results.get("final_count", 0)
        lines.append(f"  최종명세서 .. {cnt}개")

    lines.append("")

    # 처리 매장 목록
    if file_results.get("invoice_stores"):
        lines.append("=== 처리 매장 ===")
        for s in file_results["invoice_stores"]:
            lines.append(f"  - {s}")
        lines.append("")

    # 다음 단계 안내
    if not file_results.get("final_invoices") and file_results.get("invoices"):
        lines.append("=== 다음 단계 ===")
        lines.append("  단가 입력 후 최종명세서 생성 요청")

    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="일일 워크플로우 결과 → 카카오톡 전송")
    parser.add_argument("--date", help="대상 날짜 (YYYY-MM-DD)", default=None)
    parser.add_argument("--results", help="워크플로우 단계별 결과 JSON", default=None)
    args = parser.parse_args()

    date_str = args.date or datetime.now().strftime("%Y-%m-%d")

    stage_results = None
    if args.results:
        try:
            stage_results = json.loads(args.results)
        except json.JSONDecodeError:
            print("Warning: Could not parse --results JSON")

    print(f"Collecting results for {date_str}...")
    file_results = collect_results(date_str)

    message = build_message(date_str, file_results, stage_results)

    print("--- Preview ---")
    print(message)
    print("---")

    tokens = load_and_refresh_token()
    if not tokens:
        # 토큰 없으면 파일로 저장
        ymd = datetime.strptime(date_str, "%Y-%m-%d").strftime("%Y%m%d")
        fallback = os.path.join(PROJECT_ROOT, "data", "outputs", f"daily_summary_{ymd}.txt")
        with open(fallback, 'w', encoding='utf-8') as f:
            f.write(message)
        print(f"Saved to: {fallback}")
        return 1

    if send_kakao_message(message, tokens):
        print("Daily summary sent via KakaoTalk!")
        return 0
    else:
        print("Failed to send.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
