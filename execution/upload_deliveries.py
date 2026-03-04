"""
upload_deliveries.py — processed_orders → Firebase 배송 데이터 업로드

기존 simple_docx_converter.py의 파싱 로직 재사용:
  - parse_sections(): 텍스트 → 섹션별 딕셔너리
  - load_page_config(): 배송 순서 (page_config)
  - extract_stock_market_items(): 재고/시장 색상 분류

Usage:
  python execution/upload_deliveries.py --date 2026-03-04
  python execution/upload_deliveries.py --date 2026-03-04 --dry-run
  python execution/upload_deliveries.py --date 2026-03-04 --json-only  # Firebase 없이 JSON 출력
"""

import argparse
import json
import re
import sys
import os
from pathlib import Path
from datetime import datetime

# ─── Paths ────────────────────────────────────────────────────

PROJECT_ROOT = Path(__file__).resolve().parent.parent
MAPPINGS_PATH = PROJECT_ROOT / "skills" / "order_processing" / "resources" / "mappings.json"
INPUTS_DIR = PROJECT_ROOT / "data" / "inputs"


# ─── Reuse parse logic from simple_docx_converter ────────────

def parse_sections(file_path):
    """Parse text file into named sections (from simple_docx_converter.py)"""
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = [line.rstrip() for line in f.readlines()]

    sections = {}
    current_name = None
    current_lines = []

    for line in lines:
        if not line:
            continue
        if line.startswith('-'):
            if current_name:
                sections[current_name] = current_lines
            current_name = line[1:]  # Remove leading '-'
            current_lines = [line]
        else:
            current_lines.append(line)

    if current_name:
        sections[current_name] = current_lines

    return sections


def load_mappings():
    """Load mappings.json"""
    with open(MAPPINGS_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_page_config(mappings, target_date=None):
    """Load page configuration (배송 순서)"""
    if target_date:
        dt = datetime.strptime(target_date, "%Y-%m-%d")
        weekday = dt.weekday()
    else:
        weekday = datetime.now().weekday()

    config_key = "friday" if weekday == 4 else "default"
    return mappings['page_config'][config_key]


# ─── Core Logic ──────────────────────────────────────────────

def is_supplier_section(name, supplier_names):
    """공급처 섹션인지 확인"""
    normalized = name.replace(" ", "").replace("_", "")
    for supplier in supplier_names:
        if supplier.replace(" ", "").replace("_", "") == normalized:
            return True
    return False


def is_special_section(name):
    """재고/시장 특수 섹션인지 확인"""
    lower = name.strip()
    return lower.startswith("재고") or lower.startswith("시장") or "창고소분" in lower or "시장소분" in lower


def extract_store_deliveries(sections, supplier_names):
    """매장 섹션만 추출 (공급처/재고/시장 제외)"""
    store_sections = {}
    for name, lines in sections.items():
        if is_supplier_section(name, supplier_names):
            continue
        if is_special_section(name):
            continue
        # 동원1차도 제외 (내부용)
        if "동원1차" in name:
            continue
        store_sections[name] = lines
    return store_sections


def extract_color_map(sections):
    """재고/시장 섹션에서 품목별 색상 매핑 추출"""
    # 패턴: "품목 수량단위(매장명)" 형태
    pattern = re.compile(r'(.+?)\s+([\d.]+\s*\S+)\s*\((.+?)\)')
    color_map = {}  # {(item_clean, store_clean): color}

    for name, lines in sections.items():
        if "재고" in name or "창고소분" in name:
            color = "red"
        elif "시장" in name or "시장소분" in name:
            color = "blue"
        else:
            continue

        for line in lines:
            if line.startswith('-'):
                continue
            match = pattern.match(line)
            if match:
                item = match.group(1).strip()
                store = match.group(3).strip()
                color_map[(item, store)] = color

    return color_map


def parse_items(lines, store_name, color_map):
    """섹션 라인 → 배송 품목 리스트"""
    items = []
    for line in lines:
        if line.startswith('-'):
            continue
        if line.startswith('('):
            continue  # 메모/주석 라인 건너뛰기

        # 품목명과 수량 분리
        # 예: "감자 10kg", "새싹(대) 2팩", "깐양파2호 10kg 2봉"
        parts = line.strip()
        if not parts:
            continue

        # 색상 결정: color_map에서 (품목, 매장) 조합 확인
        color = "black"
        for (item_key, store_key), c in color_map.items():
            # 매장명 부분 매칭
            store_norm = store_name.replace(" ", "")
            store_key_norm = store_key.replace(" ", "")
            if store_key_norm in store_norm or store_norm in store_key_norm:
                # 품목명 부분 매칭
                if item_key in parts:
                    color = c
                    break

        # 품목명/수량 분리 (마지막 공백 기준)
        # "감자 10kg" → name="감자", qty="10kg"
        # "깐양파2호 10kg 2봉" → name="깐양파2호", qty="10kg 2봉"
        match = re.match(r'(.+?)\s+([\d].+)$', parts)
        if match:
            name = match.group(1).strip()
            qty = match.group(2).strip()
        else:
            name = parts
            qty = ""

        items.append({
            "name": name,
            "qty": qty,
            "color": color,
        })

    return items


def make_store_id(store_name):
    """매장명 → Firestore 문서 ID (영문/숫자/언더스코어)"""
    # 공백 → 언더스코어, 한글 그대로 유지
    sid = store_name.strip().replace(" ", "_")
    # Firestore 금지 문자 제거 (/, ., ..)
    sid = re.sub(r'[/.\[\]*]', '', sid)
    return sid


def compute_delivery_order(store_name, page_config):
    """page_config 기반 배송 순서 계산"""
    order_idx = 0
    for page_key in ['page1', 'page2', 'page3']:
        page_stores = page_config.get(page_key, [])
        if isinstance(page_stores, list):
            for i, ps in enumerate(page_stores):
                order_idx += 1
                # 부분 매칭 (예: "압구정점" in "오레노이키루미치 압구정점")
                ps_norm = ps.replace(" ", "")
                name_norm = store_name.replace(" ", "")
                if ps_norm in name_norm or name_norm in ps_norm:
                    return order_idx
    return 999  # page_config에 없는 매장


def build_delivery_data(date_str, processed_orders_path):
    """
    메인 함수: processed_orders 파싱 → Firebase 업로드용 데이터 구조 생성

    Returns:
        dict: {
            "date": "YYYYMMDD",
            "status": "preparing",
            "deliveries": [ { id, storeName, order, items, status, ... }, ... ]
        }
    """
    # 1. Parse sections
    sections = parse_sections(processed_orders_path)

    # 2. Load mappings
    mappings = load_mappings()
    supplier_names = mappings.get('supplier_names', [])

    # 3. date → page config
    iso_date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
    page_config = load_page_config(mappings, iso_date)

    # 4. Extract store-only sections
    store_sections = extract_store_deliveries(sections, supplier_names)

    # 5. Color map (재고=red, 시장=blue)
    color_map = extract_color_map(sections)

    # 6. Build deliveries
    deliveries = []
    for store_name, lines in store_sections.items():
        store_id = make_store_id(store_name)
        items = parse_items(lines, store_name, color_map)
        order = compute_delivery_order(store_name, page_config)

        if not items:
            continue

        deliveries.append({
            "id": store_id,
            "storeName": store_name,
            "order": order,
            "assignedTo": None,
            "driverName": None,
            "status": "pending",
            "items": items,
            "completedAt": None,
            "photoUrl": None,
            "signatureUrl": None,
            "notes": "",
        })

    # Sort by delivery order
    deliveries.sort(key=lambda d: d['order'])

    # Re-number after sorting
    for i, d in enumerate(deliveries, 1):
        d['order'] = i

    return {
        "date": date_str,
        "status": "preparing",
        "totalDeliveries": len(deliveries),
        "completedCount": 0,
        "deliveries": deliveries,
    }


# ─── Firebase Upload ─────────────────────────────────────────

def upload_to_firestore(data):
    """Firebase Admin SDK로 Firestore에 업로드"""
    try:
        import firebase_admin
        from firebase_admin import credentials, firestore
    except ImportError:
        print("[ERROR] firebase-admin 패키지가 필요합니다: pip install firebase-admin")
        sys.exit(1)

    # Initialize Firebase Admin
    if not firebase_admin._apps:
        service_key_path = os.environ.get('FIREBASE_SERVICE_ACCOUNT_KEY')
        # Fallback: check project root
        if not service_key_path or not Path(service_key_path).exists():
            service_key_path = str(PROJECT_ROOT / 'serviceAccountKey.json')
        if not Path(service_key_path).exists():
            print("[ERROR] serviceAccountKey.json을 찾을 수 없습니다.")
            print(f"        확인 경로: {service_key_path}")
            sys.exit(1)
        cred = credentials.Certificate(service_key_path)
        firebase_admin.initialize_app(cred)

    db = firestore.client()
    date_str = data['date']

    # Create/update day document
    day_ref = db.collection('deliveryDays').document(date_str)
    day_ref.set({
        'status': data['status'],
        'totalDeliveries': data['totalDeliveries'],
        'completedCount': data['completedCount'],
        'createdAt': firestore.SERVER_TIMESTAMP,
    })

    # Upload each delivery as sub-document
    batch = db.batch()
    for delivery in data['deliveries']:
        doc_ref = day_ref.collection('deliveries').document(delivery['id'])
        doc_data = {k: v for k, v in delivery.items() if k != 'id'}
        batch.set(doc_ref, doc_data)

    batch.commit()
    print(f"  [OK] Firestore에 {len(data['deliveries'])}개 배송 업로드 완료 ({date_str})")


# ─── CLI ─────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='processed_orders → Firebase 배송 데이터 업로드')
    parser.add_argument('--date', required=True, help='날짜 (YYYYMMDD 또는 YYYY-MM-DD)')
    parser.add_argument('--file', help='processed_orders 파일 경로 (기본: data/inputs/processed_orders_YYYYMMDD.txt)')
    parser.add_argument('--dry-run', action='store_true', help='업로드 없이 미리보기')
    parser.add_argument('--json-only', action='store_true', help='JSON 출력만 (Firebase 없이)')
    args = parser.parse_args()

    # Normalize date
    date_str = args.date.replace('-', '')
    if len(date_str) != 8:
        print(f"[ERROR] 날짜 형식 오류: {args.date}")
        sys.exit(1)

    # Find processed_orders file
    if args.file:
        orders_path = Path(args.file)
    else:
        orders_path = INPUTS_DIR / f"processed_orders_{date_str}.txt"

    if not orders_path.exists():
        print(f"[ERROR] 파일을 찾을 수 없습니다: {orders_path}")
        sys.exit(1)

    print(f"[배송 업로드] {date_str}")
    print(f"  파일: {orders_path}")

    # Build data
    data = build_delivery_data(date_str, str(orders_path))

    # Print summary
    print(f"\n  총 {data['totalDeliveries']}개 매장 배송:")
    for d in data['deliveries']:
        color_info = ""
        red_count = sum(1 for i in d['items'] if i['color'] == 'red')
        blue_count = sum(1 for i in d['items'] if i['color'] == 'blue')
        if red_count or blue_count:
            parts = []
            if red_count:
                parts.append(f"재고{red_count}")
            if blue_count:
                parts.append(f"시장{blue_count}")
            color_info = f" ({', '.join(parts)})"
        print(f"  #{d['order']:2d} {d['storeName']:<25s} {len(d['items'])}개 품목{color_info}")

    if args.json_only:
        output_path = INPUTS_DIR.parent / "outputs" / f"deliveries_{date_str}.json"
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"\n  [OK] JSON 저장: {output_path}")
        return

    if args.dry_run:
        print("\n  [DRY-RUN] 업로드하지 않았습니다.")
        # Print sample JSON
        if data['deliveries']:
            sample = data['deliveries'][0]
            print(f"\n  샘플 (#{sample['order']} {sample['storeName']}):")
            print(json.dumps(sample, ensure_ascii=False, indent=4))
        return

    # Upload to Firebase
    upload_to_firestore(data)


if __name__ == '__main__':
    main()
