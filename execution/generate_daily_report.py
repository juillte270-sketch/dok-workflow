"""
일일업무보고서 생성: 전날 HWP 보고서를 복사 → 날짜만 교체하여 오늘 보고서 생성.

Usage:
    python execution/generate_daily_report.py --date 2026-03-03

Requires: olefile (pip install olefile)
"""
import argparse
import shutil
import sys
from datetime import datetime, timedelta
from pathlib import Path

# --------------- Config ---------------
REPORT_BASE = Path(r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\일일업무보고")
DAY_NAMES_KR = ["월", "화", "수", "목", "금", "토", "일"]


def get_day_kr(dt: datetime) -> str:
    return DAY_NAMES_KR[dt.weekday()]


def month_folder_name(dt: datetime) -> str:
    """e.g. '26년 3월'"""
    return f"{dt.year % 100}년 {dt.month}월"


def report_filename(dt: datetime) -> str:
    """e.g. '일일업무보고서(3.3 화).hwp'"""
    return f"일일업무보고서({dt.month}.{dt.day} {get_day_kr(dt)}).hwp"


def find_previous_report(target_date: datetime, max_lookback: int = 14):
    """Find the most recent existing report before target_date."""
    for i in range(1, max_lookback + 1):
        prev = target_date - timedelta(days=i)
        folder = REPORT_BASE / month_folder_name(prev)
        if not folder.exists():
            continue
        # Try various filename patterns (some have extra spaces)
        candidates = [
            report_filename(prev),
            f"일일업무보고서({prev.month}.{prev.day} {get_day_kr(prev)} ).hwp",
            f"일일업무보고서({prev.month}.{prev.day}  {get_day_kr(prev)}).hwp",
        ]
        for c in candidates:
            p = folder / c
            if p.exists():
                return p, prev
        # Also scan folder for matching date
        for f in folder.iterdir():
            if f.suffix.lower() == ".hwp" and f"({prev.month}.{prev.day}" in f.name:
                return f, prev
    return None, None


def _utf16le(s: str) -> bytes:
    return s.encode("utf-16-le")


def replace_dates_in_binary(data: bytes, old_date: datetime, new_date: datetime) -> bytes:
    """Replace date strings in UTF-16LE encoded binary data."""
    old_next = old_date + timedelta(days=1)
    new_next = new_date + timedelta(days=1)

    replacements = []

    # 1) Header: "YYYY년 M월 D일" → new
    old_header = f"{old_date.year}년 {old_date.month}월 {old_date.day}일"
    new_header = f"{new_date.year}년 {new_date.month}월 {new_date.day}일"
    replacements.append((_utf16le(old_header), _utf16le(new_header)))

    # 2) Day-of-week after header: "(요일)" right after "일"
    old_day_paren = f"({get_day_kr(old_date)})"
    new_day_paren = f"({get_day_kr(new_date)})"

    # 3) Short date: "M/D(요일)" for old_date → new_date
    old_short = f"{old_date.month}/{old_date.day}({get_day_kr(old_date)})"
    new_short = f"{new_date.month}/{new_date.day}({get_day_kr(new_date)})"
    replacements.append((_utf16le(old_short), _utf16le(new_short)))

    # 4) Short date: next day
    old_next_short = f"{old_next.month}/{old_next.day}({get_day_kr(old_next)})"
    new_next_short = f"{new_next.month}/{new_next.day}({get_day_kr(new_next)})"
    replacements.append((_utf16le(old_next_short), _utf16le(new_next_short)))

    # 5) Bare "M/D" without day-of-week (in case)
    old_bare = f"{old_date.month}/{old_date.day}"
    new_bare = f"{new_date.month}/{new_date.day}"

    # Apply replacements (specific first, then generic)
    for old_bytes, new_bytes in replacements:
        data = data.replace(old_bytes, new_bytes)

    return data


def replace_dates_in_text(text: str, old_date: datetime, new_date: datetime) -> str:
    """Replace date strings in plain text (PrvText)."""
    old_next = old_date + timedelta(days=1)
    new_next = new_date + timedelta(days=1)

    # Header date
    old_header = f"{old_date.year}년 {old_date.month}월 {old_date.day}일"
    new_header = f"{new_date.year}년 {new_date.month}월 {new_date.day}일"
    text = text.replace(old_header, new_header)

    # Short date with day of week: M/D(요일)
    old_short = f"{old_date.month}/{old_date.day}({get_day_kr(old_date)})"
    new_short = f"{new_date.month}/{new_date.day}({get_day_kr(new_date)})"
    text = text.replace(old_short, new_short)

    # Next day
    old_next_short = f"{old_next.month}/{old_next.day}({get_day_kr(old_next)})"
    new_next_short = f"{new_next.month}/{new_next.day}({get_day_kr(new_next)})"
    text = text.replace(old_next_short, new_next_short)

    # Day-of-week in parentheses after "일"
    old_day_header = f"({get_day_kr(old_date)})"
    new_day_header = f"({get_day_kr(new_date)})"
    # Only replace first occurrence (header)
    text = text.replace(old_day_header, new_day_header, 1)

    return text


def generate_report(target_date: datetime, dry_run: bool = False) -> str:
    """Generate daily report by copying and modifying previous day's report."""
    # Find previous report
    prev_path, prev_date = find_previous_report(target_date)
    if prev_path is None:
        raise FileNotFoundError(
            f"이전 보고서를 찾을 수 없습니다 (최근 14일 내). "
            f"기준폴더: {REPORT_BASE}"
        )

    print(f"이전 보고서: {prev_path.name} ({prev_date.strftime('%Y-%m-%d')})")

    # Target path
    target_folder = REPORT_BASE / month_folder_name(target_date)
    target_folder.mkdir(parents=True, exist_ok=True)
    target_path = target_folder / report_filename(target_date)

    if target_path.exists():
        print(f"⚠️ 이미 존재: {target_path.name}")
        return str(target_path)

    if dry_run:
        print(f"[DRY-RUN] 생성 예정: {target_path.name}")
        print(f"  폴더: {target_folder}")
        print(f"  날짜 변경: {prev_date.strftime('%m/%d')}({get_day_kr(prev_date)}) → "
              f"{target_date.strftime('%m/%d')}({get_day_kr(target_date)})")
        return str(target_path)

    # Step 1: Copy file
    shutil.copy2(prev_path, target_path)
    print(f"파일 복사 완료: {prev_path.name} → {target_path.name}")

    # Step 2: Modify dates inside HWP (PrvText via raw binary replacement)
    _modify_hwp_raw(target_path, prev_date, target_date)

    print(f"✅ 보고서 생성 완료: {target_path}")
    return str(target_path)


def _modify_hwp_raw(target_path: Path, prev_date: datetime, target_date: datetime):
    """Replace date strings directly in raw HWP binary (UTF-16LE).

    HWP body is zlib-compressed, but PrvText and some metadata are uncompressed.
    We decompress body, replace, recompress, and reassemble the OLE file.
    """
    old_next = prev_date + timedelta(days=1)
    new_next = target_date + timedelta(days=1)

    # Order matters: next-day FIRST (to avoid double-replacement),
    # then today's short date, then header.
    replacements = [
        # 1) Next-day short: "3/3(화)" → "3/4(수)" — MUST be first
        (f"{old_next.month}/{old_next.day}({get_day_kr(old_next)})",
         f"{new_next.month}/{new_next.day}({get_day_kr(new_next)})"),
        # 2) Today short: "3/2(월)" → "3/3(화)"
        (f"{prev_date.month}/{prev_date.day}({get_day_kr(prev_date)})",
         f"{target_date.month}/{target_date.day}({get_day_kr(target_date)})"),
        # 3) Header full: "2026년 3월 2일 (월)" → "2026년 3월 3일 (화)"
        (f"{prev_date.year}년 {prev_date.month}월 {prev_date.day}일 ({get_day_kr(prev_date)})",
         f"{target_date.year}년 {target_date.month}월 {target_date.day}일 ({get_day_kr(target_date)})"),
        # 4) Header date only (fallback): "2026년 3월 2일" → "2026년 3월 3일"
        (f"{prev_date.year}년 {prev_date.month}월 {prev_date.day}일",
         f"{target_date.year}년 {target_date.month}월 {target_date.day}일"),
    ]

    with open(target_path, "rb") as f:
        raw = f.read()

    count = 0
    for old_str, new_str in replacements:
        old_b = _utf16le(old_str)
        new_b = _utf16le(new_str)
        n = raw.count(old_b)
        if n > 0:
            raw = raw.replace(old_b, new_b)
            count += n

    with open(target_path, "wb") as f:
        f.write(raw)

    print(f"날짜 교체 완료: {count}건 치환 "
          f"({prev_date.month}/{prev_date.day}({get_day_kr(prev_date)}) → "
          f"{target_date.month}/{target_date.day}({get_day_kr(target_date)}))")


def main():
    parser = argparse.ArgumentParser(description="일일업무보고서 생성")
    parser.add_argument("--date", required=True, help="대상 날짜 (YYYY-MM-DD)")
    parser.add_argument("--dry-run", action="store_true", help="미리보기 (파일 생성 안 함)")
    args = parser.parse_args()

    target_date = datetime.strptime(args.date, "%Y-%m-%d")
    result = generate_report(target_date, dry_run=args.dry_run)

    # Notify dashboard
    try:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        from _notify import notify
        notify("stage_report", date_str=args.date)
    except Exception:
        pass

    print(f"\n결과: {result}")


if __name__ == "__main__":
    main()
