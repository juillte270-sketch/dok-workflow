"""
공급처 영수증에서 추출한 매입가를 단가시트(Col I)에 입력하는 스크립트.

parse_receipt.py로 파싱한 결과를 단가시트에 매칭하여 매입가(Col I = 9)를 채운다.

Usage:
  # JSON 파일에서 입력
  python execution/fill_receipt_prices.py \
    --master "G:/내 드라이브/1. 도크_주문 명세서/0. 매입단가_자료/도크발주관리데이터.xlsx" \
    --date 2026-03-01 \
    --receipt-json '[{"supplier":"건영농산","items":[{"name":"감자/왕왕","unit_price":62000}]}]'

  # dry-run으로 미리보기
  python execution/fill_receipt_prices.py \
    --master "..." --date 2026-03-01 --receipt-json '...' --dry-run

  # URL에서 직접 파싱 + 입력
  python execution/fill_receipt_prices.py \
    --master "..." --date 2026-03-01 --url "https://www.itanet.co.kr/..."
"""

import argparse
import json
import os
import re
import sys
import time

import openpyxl

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(os.path.join(PROJECT_ROOT, 'execution'))

from parse_receipt import (
    parse_receipt, map_to_danga_items, load_supplier_order_items,
    RECEIPT_MAP, ITEM_MAP, RECEIPT_TYPE
)

# ==================== 설정 ====================

COL_CATEGORY = 1     # 구분
COL_DATE = 2         # 날짜
COL_ITEM = 3         # 품목명
COL_GRADE = 4        # 등급
COL_UNIT = 5         # 단위
COL_AUCTION_MAX = 6  # 경매 최고가
COL_AUCTION_AVG = 7  # 경매 평균가
COL_SIKBOM = 8       # 식봄가
COL_PURCHASE = 9     # 매입가 (우리가 채울 컬럼)

# 고정가 품목 (fill_prices.py FIXED_PRICES와 동일) — 절대 덮어쓰지 않음
FIXED_PRICE_ITEMS = {
    '숙주', '굵은숙주', '새싹(대)', '무순(대)', '중란', '대란',
    '곱슬이콩나물', '일자콩나물', '아보카도', '중숙아보카도',
    '파김치', '맛김치', '냉동알마늘', '쌀', '매운고춧가루', '굵은고춧가루',
}

# 가격 이상치 검증: 경매가/식봄가 대비 허용 비율
PRICE_TOLERANCE = 0.5  # ±50%


# ==================== 단가시트 행 찾기 ====================

def find_target_rows(ws, target_date):
    """단가시트에서 target_date에 해당하는 첫 번째 연속 블록의 행을 수집.

    기존 fill_auction_from_api.py 패턴 재사용:
    - 날짜 매칭 (문자열/datetime 모두)
    - 첫 번째 연속 블록만 (gap > 2 시 중단)
    - 소분 행 제외

    Returns:
        dict: {품목명: {"row": row_idx, "grade": ..., "unit": ..., "category": ...}}
    """
    row_map = {}
    in_block = False
    last_match_row = 0

    for row_idx in range(2, ws.max_row + 1):
        date_cell = ws.cell(row=row_idx, column=COL_DATE).value
        item_cell = ws.cell(row=row_idx, column=COL_ITEM).value

        is_match = False
        if date_cell:
            if str(date_cell).startswith(target_date):
                is_match = True
            elif hasattr(date_cell, 'strftime') and date_cell.strftime("%Y-%m-%d") == target_date:
                is_match = True

        if is_match and item_cell:
            if in_block and (row_idx - last_match_row) > 2:
                break
            in_block = True
            last_match_row = row_idx

            item_name = str(item_cell).strip()
            category = str(ws.cell(row=row_idx, column=COL_CATEGORY).value or '')
            grade = str(ws.cell(row=row_idx, column=COL_GRADE).value or '').strip()
            unit = str(ws.cell(row=row_idx, column=COL_UNIT).value or '').strip()

            # 소분 행 스킵
            if '소분' in category:
                continue

            # 첫 번째 발견만 유지 (중복 방지)
            if item_name not in row_map:
                row_map[item_name] = {
                    "row": row_idx,
                    "grade": grade,
                    "unit": unit,
                    "category": category,
                }

    return row_map


# ==================== 가격 검증 ====================

def validate_price(item_name, new_price, ws, row_idx):
    """새 매입가가 기존 경매가/식봄가 대비 이상치인지 검증.

    Returns:
        tuple: (is_valid, warning_msg)
    """
    if new_price <= 0:
        return False, "가격이 0 이하"

    # 경매 평균가(Col G), 식봄가(Col H) 비교
    ref_prices = []
    for col in (COL_AUCTION_AVG, COL_SIKBOM):
        cell = ws.cell(row=row_idx, column=col)
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            continue
        val = cell.value
        if val and isinstance(val, (int, float)) and val > 0:
            ref_prices.append(val)

    if not ref_prices:
        return True, None  # 참고가 없으면 통과

    avg_ref = sum(ref_prices) / len(ref_prices)
    lower = avg_ref * (1 - PRICE_TOLERANCE)
    upper = avg_ref * (1 + PRICE_TOLERANCE)

    if new_price < lower or new_price > upper:
        return False, f"이상치 경고: {item_name} 매입가 {new_price:,}원 (참고가 범위 {lower:,.0f}~{upper:,.0f})"

    return True, None


# ==================== 매입가 입력 ====================

def fill_receipt_prices(master_file, target_date, receipts, dry_run=False):
    """파싱된 영수증 데이터를 단가시트 Col I(매입가)에 입력.

    Args:
        master_file: 마스터 파일 경로
        target_date: 대상 날짜 (YYYY-MM-DD)
        receipts: 파싱된 영수증 리스트 [{"supplier", "items": [...]}, ...]
        dry_run: True면 실제 입력 없이 결과만 표시

    Returns:
        dict: {"updates", "skipped", "warnings", "details"}
    """
    print(f"Opening master file: {master_file}")
    wb = openpyxl.load_workbook(master_file)
    ws = wb["단가"]

    # 단가시트 대상 행 수집
    row_map = find_target_rows(ws, target_date)
    print(f"Found {len(row_map)} items for {target_date}")
    print(f"Items: {list(row_map.keys())}\n")

    updates = 0
    skipped = 0
    warnings = []
    details = []

    for receipt in receipts:
        supplier = receipt.get('supplier', '?')
        print(f"\n--- {supplier} ---")

        # 손글씨 공급처면 발주리스트 교차매칭용 품목 로드
        order_items = None
        receipt_type = RECEIPT_TYPE.get(supplier, 'unknown')
        if receipt_type == 'handwritten':
            order_items = load_supplier_order_items(supplier, target_date)
            if order_items:
                print(f"  발주리스트 로드: {[oi['name'] for oi in order_items]}")

        # 영수증 → 단가시트 매핑 (order_items로 교차매칭)
        mapped_items = map_to_danga_items(receipt, supplier, order_items=order_items)

        for item in mapped_items:
            receipt_name = item['receipt_name']
            danga_name = item['danga_name']
            unit_price = item['unit_price']
            total = item['total']
            skip_reason = item.get('skip_reason')

            # 고정가 품목 스킵
            if skip_reason:
                print(f"  [SKIP] {receipt_name} → {skip_reason}")
                skipped += 1
                details.append(f"{supplier}/{receipt_name}: SKIP ({skip_reason})")
                continue

            # FIXED_PRICE_ITEMS에 해당하는 단가시트 품목도 스킵
            if danga_name in FIXED_PRICE_ITEMS:
                print(f"  [SKIP] {receipt_name} → {danga_name} (고정가 품목)")
                skipped += 1
                details.append(f"{supplier}/{receipt_name}: SKIP (고정가)")
                continue

            # 단가시트에서 품목 찾기
            row_info = row_map.get(danga_name)
            if not row_info:
                # 부분 매칭 시도 (긴 이름 우선)
                for key in sorted(row_map.keys(), key=len, reverse=True):
                    if danga_name in key or key in danga_name:
                        row_info = row_map[key]
                        danga_name = key
                        break

            if not row_info:
                print(f"  [MISS] {receipt_name} → {danga_name} (단가시트에 없음)")
                details.append(f"{supplier}/{receipt_name}: NOT FOUND in sheet")
                continue

            row_idx = row_info['row']

            # 사용할 가격 결정 (unit_price × qty가 total과 일치하는지 검증)
            qty = item.get('qty', 1) or 1
            if unit_price > 0 and total > 0:
                calculated_total = unit_price * qty
                if abs(calculated_total - total) > total * 0.1:  # 10% 이상 차이
                    print(f"  [WARN] {receipt_name}: unit_price {unit_price} × qty={qty} = {calculated_total} ≠ total {total} → total 사용")
                    price = total
                else:
                    price = unit_price
            else:
                price = unit_price if unit_price > 0 else total
            if price <= 0:
                print(f"  [SKIP] {receipt_name} → 가격 0")
                skipped += 1
                continue

            # 가격 검증
            is_valid, warning = validate_price(danga_name, price, ws, row_idx)
            if not is_valid and warning:
                print(f"  [WARN] {warning}")
                warnings.append(warning)
                # 경고는 하되 입력은 진행 (사용자가 나중에 확인)

            # 셀에 입력
            cell_i = ws.cell(row=row_idx, column=COL_PURCHASE)
            if isinstance(cell_i, openpyxl.cell.cell.MergedCell):
                print(f"  [SKIP] {receipt_name} R{row_idx} merged cell")
                skipped += 1
                continue

            old_val = cell_i.value
            if dry_run:
                print(f"  [DRY] {receipt_name} → {danga_name} R{row_idx}: {old_val} → {price:,}")
            else:
                cell_i.value = price
                cell_i.number_format = '#,##0'
                print(f"  [OK] {receipt_name} → {danga_name} R{row_idx}: {old_val} → {price:,}")

            updates += 1
            details.append(f"{supplier}/{receipt_name} → {danga_name} R{row_idx}: {price:,}")

    # 결과 요약
    print(f"\n=== 결과 ===")
    print(f"입력: {updates}건 / 스킵: {skipped}건 / 경고: {len(warnings)}건")
    if warnings:
        print("경고 목록:")
        for w in warnings:
            print(f"  [!] {w}")

    # 저장
    if updates > 0 and not dry_run:
        _save_workbook(wb, master_file)
    else:
        if dry_run:
            print("\n[DRY-RUN] 실제 파일 변경 없음")
        wb.close()

    return {
        "updates": updates,
        "skipped": skipped,
        "warnings": warnings,
        "details": details,
    }


def _save_workbook(wb, filepath):
    """워크북 저장 (잠금 시 재시도, fill_auction_from_api.py 패턴)."""
    saved = False
    for attempt in range(5):
        try:
            print(f"\nSaving to {filepath}... (attempt {attempt + 1})")
            wb.save(filepath)
            print("Done.")
            saved = True
            break
        except PermissionError:
            if attempt < 4:
                print(f"  File locked, retrying in 3s...")
                time.sleep(3)
            else:
                import tempfile
                tmp = os.path.join(tempfile.gettempdir(), "도크발주관리데이터_receipt.xlsx")
                wb.save(tmp)
                print(f"  Saved to temp: {tmp}")
                print(f"  Please close the file and copy manually.")
                saved = True
    wb.close()
    return saved


# ==================== 일괄 처리 (폴더 스캔) ====================

def _load_past_hashes(state_dir, target_date):
    """과거 날짜의 처리완료 해시를 모두 로드하여 set으로 반환."""
    import glob as glob_mod
    past_hashes = set()
    date_compact = target_date.replace('-', '')
    pattern = os.path.join(state_dir, 'receipt_processed_*.json')
    for fpath in glob_mod.glob(pattern):
        fname = os.path.basename(fpath)
        # receipt_processed_YYYYMMDD.json 에서 날짜 추출
        file_date = fname.replace('receipt_processed_', '').replace('.json', '')
        if file_date == date_compact:
            continue  # 오늘 날짜는 스킵 (과거만 로드)
        try:
            with open(fpath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            past_hashes.update(data.keys())
        except (json.JSONDecodeError, OSError):
            pass
    return past_hashes


def _save_today_processed(state_dir, target_date, today_processed):
    """오늘 날짜 처리완료 해시를 파일에 머지 저장."""
    if not today_processed:
        return
    os.makedirs(state_dir, exist_ok=True)
    date_compact = target_date.replace('-', '')
    fpath = os.path.join(state_dir, f'receipt_processed_{date_compact}.json')

    existing = {}
    if os.path.exists(fpath):
        try:
            with open(fpath, 'r', encoding='utf-8') as f:
                existing = json.load(f)
        except (json.JSONDecodeError, OSError):
            pass

    existing.update(today_processed)
    with open(fpath, 'w', encoding='utf-8') as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"[STATE] 처리 기록 저장: {fpath} ({len(today_processed)}건 추가, 총 {len(existing)}건)")


def batch_process_image_dir(image_dir, master_file, target_date, since_minutes=1440,
                            dry_run=False, state_dir=None):
    """폴더 내 영수증 이미지를 일괄 파싱 + 단가시트 입력.

    Args:
        image_dir: 이미지 폴더 경로 (카카오톡 다운로드 폴더)
        master_file: 마스터 파일 경로
        target_date: 대상 날짜 (YYYY-MM-DD)
        since_minutes: 최근 N분 이내 파일만 처리 (기본 24시간)
        dry_run: True면 실제 입력 없이 결과만 표시
        state_dir: 처리완료 기록 디렉토리 (기본: dashboard/state/)

    Returns:
        dict: {"total_files", "processed", "failed", "results": [...]}
    """
    import hashlib
    from datetime import datetime, timedelta

    if state_dir is None:
        state_dir = os.path.join(PROJECT_ROOT, 'dashboard', 'state')

    cutoff = datetime.now() - timedelta(minutes=since_minutes)
    extensions = {'.jpg', '.jpeg', '.png', '.bmp'}

    # 과거 날짜 처리완료 해시 로드
    past_hashes = _load_past_hashes(state_dir, target_date)
    if past_hashes:
        print(f"[STATE] 과거 날짜 처리완료 해시 {len(past_hashes)}건 로드")

    # 이미지 파일 스캔
    files = []
    for fname in os.listdir(image_dir):
        ext = os.path.splitext(fname)[1].lower()
        if ext not in extensions:
            continue
        fpath = os.path.join(image_dir, fname)
        mtime = datetime.fromtimestamp(os.path.getmtime(fpath))
        if mtime >= cutoff:
            files.append((fpath, fname, mtime))

    files.sort(key=lambda x: x[2])  # 시간순 정렬

    if not files:
        print(f"[BATCH] {image_dir} 에서 최근 {since_minutes}분 이내 이미지 없음")
        return {"total_files": 0, "processed": 0, "failed": 0, "results": []}

    print(f"[BATCH] {len(files)}개 이미지 발견 (최근 {since_minutes}분)")
    for fp, fn, mt in files:
        print(f"  {fn} ({mt.strftime('%H:%M')})")

    # 이미지별 파싱
    all_receipts = []
    file_results = []
    seen_hashes = set()
    today_processed = {}  # 오늘 처리 완료한 해시 기록

    for fpath, fname, mtime in files:
        # 중복 체크 (파일 해시)
        with open(fpath, 'rb') as f:
            file_hash = hashlib.md5(f.read()[:8192]).hexdigest()

        # 이전 날짜에 처리된 파일 스킵
        if file_hash in past_hashes:
            print(f"\n[SKIP] {fname} (이전 날짜에 처리됨)")
            file_results.append({"file": fname, "status": "skipped", "reason": "past_date"})
            continue

        if file_hash in seen_hashes:
            print(f"\n[SKIP] {fname} (중복)")
            file_results.append({"file": fname, "status": "skipped", "reason": "duplicate"})
            continue
        seen_hashes.add(file_hash)

        print(f"\n[PARSE] {fname}...")
        try:
            result = parse_receipt(image_path=fpath)
            if 'error' in result:
                print(f"  [ERROR] {result['error']}")
                file_results.append({"file": fname, "status": "error", "error": result['error']})
                continue

            supplier = result.get('supplier', '?')
            items = result.get('items', [])
            print(f"  [OK] {supplier}: {len(items)} items")
            for item in items:
                print(f"    {item.get('name', '?')} = {item.get('unit_price', 0):,}")

            if items:
                all_receipts.append(result)
                file_results.append({
                    "file": fname,
                    "status": "parsed",
                    "supplier": supplier,
                    "items": len(items),
                })
                # 처리 완료 기록
                today_processed[file_hash] = {
                    "file": fname,
                    "supplier": supplier,
                    "timestamp": datetime.now().isoformat(),
                }
            else:
                file_results.append({"file": fname, "status": "empty", "supplier": supplier})

        except Exception as e:
            print(f"  [ERROR] {e}")
            file_results.append({"file": fname, "status": "error", "error": str(e)})

    # 일괄 단가시트 입력
    if all_receipts:
        print(f"\n[FILL] {len(all_receipts)}개 영수증 → 단가시트 입력")
        fill_result = fill_receipt_prices(
            master_file=master_file,
            target_date=target_date,
            receipts=all_receipts,
            dry_run=dry_run,
        )
    else:
        fill_result = {"updates": 0, "skipped": 0, "warnings": [], "details": []}

    # 처리 완료 기록 저장 (dry-run이 아닐 때만)
    if today_processed and not dry_run:
        _save_today_processed(state_dir, target_date, today_processed)

    return {
        "total_files": len(files),
        "processed": sum(1 for r in file_results if r["status"] == "parsed"),
        "failed": sum(1 for r in file_results if r["status"] == "error"),
        "past_skipped": sum(1 for r in file_results if r.get("reason") == "past_date"),
        "file_results": file_results,
        **fill_result,
    }


# ==================== CLI ====================

def main():
    parser = argparse.ArgumentParser(description='공급처 영수증 매입가 → 단가시트 입력')
    parser.add_argument('--master', required=True, help='마스터 파일 경로')
    parser.add_argument('--date', required=True, help='대상 날짜 (YYYY-MM-DD)')
    parser.add_argument('--receipt-json', help='파싱된 영수증 JSON (리스트 또는 단일 객체)')
    parser.add_argument('--url', help='영수증 URL (직접 파싱 + 입력)')
    parser.add_argument('--image', help='영수증 이미지 파일 (직접 파싱 + 입력)')
    parser.add_argument('--image-url', help='영수증 이미지 URL (직접 파싱 + 입력)')
    parser.add_argument('--text', help='영수증 텍스트 (직접 파싱 + 입력)')
    parser.add_argument('--image-dir', help='영수증 이미지 폴더 (일괄 처리)')
    parser.add_argument('--since-minutes', type=int, default=1440,
                        help='일괄 처리 시 최근 N분 이내 파일만 (기본 1440=24시간)')
    parser.add_argument('--state-dir', default=None,
                        help='처리완료 기록 디렉토리 (기본: dashboard/state/)')
    parser.add_argument('--supplier', help='공급처명 힌트')
    parser.add_argument('--dry-run', action='store_true', help='실제 입력 없이 결과만 표시')
    args = parser.parse_args()

    # 영수증 데이터 준비
    receipts = []

    if args.image_dir:
        # 일괄 처리 모드
        result = batch_process_image_dir(
            image_dir=args.image_dir,
            master_file=args.master,
            target_date=args.date,
            since_minutes=args.since_minutes,
            dry_run=args.dry_run,
            state_dir=args.state_dir,
        )
        print(f"\n{json.dumps(result, ensure_ascii=False, indent=2)}")
        return

    if args.receipt_json:
        data = json.loads(args.receipt_json)
        if isinstance(data, list):
            receipts = data
        else:
            receipts = [data]
    elif args.url or args.image or args.image_url or args.text:
        text = args.text.replace('\\n', '\n') if args.text else None
        result = parse_receipt(
            url=args.url,
            image_path=args.image,
            image_url=args.image_url,
            text=text,
            supplier_hint=args.supplier,
        )
        if 'error' in result:
            print(f"파싱 오류: {result['error']}")
            sys.exit(1)
        print(f"파싱 결과: {json.dumps(result, ensure_ascii=False, indent=2)}\n")
        receipts = [result]
    else:
        print("--receipt-json, --url, --image, --image-url, --text, --image-dir 중 하나를 제공해야 합니다")
        sys.exit(1)

    # 매입가 입력
    result = fill_receipt_prices(
        master_file=args.master,
        target_date=args.date,
        receipts=receipts,
        dry_run=args.dry_run,
    )

    # JSON 결과 출력
    print(f"\n{json.dumps(result, ensure_ascii=False, indent=2)}")


if __name__ == '__main__':
    import sys
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        import traceback; traceback.print_exc()
        sys.exit(1)
