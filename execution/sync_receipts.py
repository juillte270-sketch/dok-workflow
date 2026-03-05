"""
카카오톡 받은 파일 → Google Drive 카톡 영수증 폴더 자동 동기화.

카카오톡 받은 파일 폴더를 감시하여 새로운 이미지/PDF 파일이 생기면
Google Drive 카톡 영수증 폴더로 자동 복사합니다.

Usage:
    python execution/sync_receipts.py              # 백그라운드 감시 시작
    python execution/sync_receipts.py --once       # 1회 동기화 후 종료
    python execution/sync_receipts.py --since 60   # 최근 60분 이내 파일만
"""
import os
import sys
import time
import shutil
import argparse
from pathlib import Path
from datetime import datetime, timedelta

# 경로 설정
KAKAO_DIRS = [
    Path(os.environ.get("USERPROFILE", "")) / "OneDrive" / "문서" / "카카오톡 받은 파일",
    Path(os.environ.get("USERPROFILE", "")) / "Documents" / "카카오톡 받은 파일",
    Path(os.environ.get("USERPROFILE", "")) / "문서" / "카카오톡 받은 파일",
]

DRIVE_RECEIPT_DIR = Path(r"G:\내 드라이브\1. 도크_주문 명세서\0. 매입단가_자료\카톡 영수증")

# 영수증으로 간주할 확장자
RECEIPT_EXTENSIONS = {".jpg", ".jpeg", ".png", ".pdf"}


def find_kakao_dir() -> Path:
    """카카오톡 받은 파일 폴더 탐지."""
    for d in KAKAO_DIRS:
        if d.is_dir():
            return d
    return None


def get_recent_files(src_dir: Path, since_minutes: int = None) -> list:
    """소스 폴더에서 영수증 파일 목록 반환."""
    files = []
    cutoff = None
    if since_minutes:
        cutoff = datetime.now() - timedelta(minutes=since_minutes)

    for f in src_dir.iterdir():
        if not f.is_file():
            continue
        if f.suffix.lower() not in RECEIPT_EXTENSIONS:
            continue
        if f.name.startswith("~$") or f.name == "desktop.ini":
            continue
        if cutoff:
            mtime = datetime.fromtimestamp(f.stat().st_mtime)
            if mtime < cutoff:
                continue
        files.append(f)

    return sorted(files, key=lambda f: f.stat().st_mtime, reverse=True)


def sync_files(src_dir: Path, dst_dir: Path, since_minutes: int = None) -> int:
    """새 파일을 dst로 복사. 이미 있으면 건너뜀. 복사 수 반환."""
    dst_dir.mkdir(parents=True, exist_ok=True)
    files = get_recent_files(src_dir, since_minutes)

    copied = 0
    for f in files:
        dst_path = dst_dir / f.name
        if dst_path.exists():
            # 크기 같으면 스킵
            if dst_path.stat().st_size == f.stat().st_size:
                continue
        shutil.copy2(f, dst_path)
        print(f"  복사: {f.name} ({f.stat().st_size / 1024:.0f}KB)", flush=True)
        copied += 1

    return copied


def watch_loop(src_dir: Path, dst_dir: Path, interval: int = 30):
    """주기적으로 새 파일 감시 + 복사."""
    print(f"감시 시작: {src_dir}", flush=True)
    print(f"복사 대상: {dst_dir}", flush=True)
    print(f"감시 주기: {interval}초", flush=True)
    print("Ctrl+C로 종료\n", flush=True)

    # 최초 실행: 최근 24시간 파일 동기화
    count = sync_files(src_dir, dst_dir, since_minutes=1440)
    if count:
        print(f"초기 동기화: {count}개 복사\n", flush=True)

    while True:
        try:
            time.sleep(interval)
            count = sync_files(src_dir, dst_dir, since_minutes=interval // 60 + 5)
            if count:
                ts = datetime.now().strftime("%H:%M:%S")
                print(f"[{ts}] {count}개 새 파일 복사", flush=True)
        except KeyboardInterrupt:
            print("\n감시 종료.", flush=True)
            break
        except Exception as e:
            print(f"오류: {e}", flush=True)
            time.sleep(10)


def main():
    parser = argparse.ArgumentParser(description="카카오톡 영수증 → Google Drive 자동 동기화")
    parser.add_argument("--once", action="store_true", help="1회 동기화 후 종료")
    parser.add_argument("--since", type=int, default=None, help="최근 N분 이내 파일만 (기본: 전체)")
    parser.add_argument("--interval", type=int, default=30, help="감시 주기 (초, 기본: 30)")
    args = parser.parse_args()

    kakao_dir = find_kakao_dir()
    if not kakao_dir:
        print("카카오톡 받은 파일 폴더를 찾을 수 없습니다.", flush=True)
        sys.exit(1)

    if not DRIVE_RECEIPT_DIR.parent.exists():
        print(f"Google Drive 경로 없음: {DRIVE_RECEIPT_DIR.parent}", flush=True)
        sys.exit(1)

    print(f"카카오톡 폴더: {kakao_dir}", flush=True)
    print(f"Drive 폴더: {DRIVE_RECEIPT_DIR}\n", flush=True)

    if args.once:
        count = sync_files(kakao_dir, DRIVE_RECEIPT_DIR, since_minutes=args.since)
        print(f"\n동기화 완료: {count}개 파일 복사", flush=True)
    else:
        watch_loop(kakao_dir, DRIVE_RECEIPT_DIR, interval=args.interval)


if __name__ == "__main__":
    main()
