"""
카카오 발주 큐 관리 CLI 도구

사용법:
    python execution/kakao_order_manager.py --status          # 오늘 접수 현황
    python execution/kakao_order_manager.py --export          # 큐 → orders_today.txt 변환
    python execution/kakao_order_manager.py --trigger         # export + 전체 파이프라인
    python execution/kakao_order_manager.py --reset           # 오늘 큐 초기화
    python execution/kakao_order_manager.py --register USER_ID STORE_NAME
    python execution/kakao_order_manager.py --date 20260224   # 특정 날짜 조회
"""

import json
import os
import sys
import subprocess
import argparse
from datetime import datetime

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(PROJECT_ROOT)

DATA_DIR = os.path.join(PROJECT_ROOT, "data", "inputs")
STORE_MAP_PATH = os.path.join(
    PROJECT_ROOT, "skills", "order_processing", "resources", "kakao_user_store_map.json"
)


def _today_str():
    return datetime.now().strftime("%Y%m%d")


def _queue_path(date_str):
    return os.path.join(DATA_DIR, f"kakao_orders_{date_str}.json")


def _load_queue(date_str):
    path = _queue_path(date_str)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def _save_queue(queue, date_str):
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


def export_queue_to_txt(queue):
    """큐 JSON → orders_today.txt 형식 텍스트"""
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
        lines.append("")
    return "\n".join(lines)


def cmd_status(date_str):
    """접수 현황 출력"""
    queue = _load_queue(date_str)
    if not queue:
        print(f"[{date_str}] 접수된 발주가 없습니다.")
        return

    closed_mark = " (마감됨)" if queue.get("closed") else ""
    print(f"\n=== 카카오 발주 현황 [{date_str}]{closed_mark} ===\n")

    stores = {}
    for o in queue["orders"]:
        stores.setdefault(o["store"], []).append(o)

    for store, entries in stores.items():
        total_items = sum(len(e["items"]) for e in entries)
        print(f"  [{store}] — {total_items}건 ({len(entries)}회 접수)")
        for entry in entries:
            for it in entry["items"]:
                print(f"    {it['raw']}")
        print()

    print(f"총 {len(stores)}개 매장, {len(queue['orders'])}건 접수")


def cmd_export(date_str):
    """큐 → txt 파일 변환"""
    queue = _load_queue(date_str)
    if not queue or not queue["orders"]:
        print(f"[{date_str}] 접수된 발주가 없습니다.")
        return None

    txt = export_queue_to_txt(queue)
    orders_path = os.path.join(DATA_DIR, "orders_today.txt")
    archive_path = os.path.join(DATA_DIR, f"orders_{date_str}.txt")

    with open(orders_path, "w", encoding="utf-8") as f:
        f.write(txt)
    with open(archive_path, "w", encoding="utf-8") as f:
        f.write(txt)

    store_count = len(set(o["store"] for o in queue["orders"]))
    print(f"내보내기 완료: {store_count}개 매장, {len(queue['orders'])}건")
    print(f"  → {orders_path}")
    print(f"  → {archive_path}")
    return orders_path


def cmd_trigger(date_str):
    """내보내기 + 파이프라인 실행"""
    orders_path = cmd_export(date_str)
    if not orders_path:
        return

    # 마감 플래그
    queue = _load_queue(date_str)
    queue["closed"] = True
    _save_queue(queue, date_str)

    date_formatted = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
    script = os.path.join(PROJECT_ROOT, "execution", "run_full_workflow.py")

    print(f"\n파이프라인 실행: {date_formatted}")
    print("=" * 60)

    result = subprocess.run(
        [sys.executable, script, "--date", date_formatted, "--input", orders_path],
        cwd=PROJECT_ROOT,
    )
    return result.returncode


def cmd_reset(date_str):
    """큐 초기화"""
    queue = {"date": date_str, "orders": [], "closed": False}
    _save_queue(queue, date_str)
    print(f"[{date_str}] 발주 큐 초기화 완료.")


def cmd_register(user_id, store_name):
    """매장 등록"""
    smap = _load_store_map()
    smap[user_id] = store_name
    _save_store_map(smap)
    print(f"매장 등록 완료: {user_id} → {store_name}")


def cmd_list_stores():
    """등록 매장 목록"""
    smap = _load_store_map()
    if not smap:
        print("등록된 매장이 없습니다.")
        return
    print(f"\n=== 등록 매장 ({len(smap)}개) ===\n")
    for uid, store in smap.items():
        print(f"  {uid[:8]}... → {store}")


def main():
    parser = argparse.ArgumentParser(description="카카오 발주 큐 관리 도구")
    parser.add_argument("--status", action="store_true", help="접수 현황 출력")
    parser.add_argument("--export", action="store_true", help="큐 → orders_today.txt")
    parser.add_argument("--trigger", action="store_true", help="마감 + 파이프라인 실행")
    parser.add_argument("--reset", action="store_true", help="큐 초기화")
    parser.add_argument("--register", nargs=2, metavar=("USER_ID", "STORE_NAME"), help="매장 등록")
    parser.add_argument("--list-stores", action="store_true", help="등록 매장 목록")
    parser.add_argument("--date", default=_today_str(), help="대상 날짜 (YYYYMMDD)")
    args = parser.parse_args()

    if args.register:
        cmd_register(args.register[0], args.register[1])
    elif args.list_stores:
        cmd_list_stores()
    elif args.status:
        cmd_status(args.date)
    elif args.export:
        cmd_export(args.date)
    elif args.trigger:
        cmd_trigger(args.date)
    elif args.reset:
        cmd_reset(args.date)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
