"""
영수증 자동 감지 데몬.

카카오톡 다운로드 폴더 감시 + 클립보드 감시를 동시 실행하여
공급처 영수증을 자동으로 파싱 → 단가시트 Col I(매입가)에 입력한다.

트리거 1: 이미지 영수증 — 카카오톡 다운로드 폴더에 새 이미지 생성 시
트리거 2: URL/텍스트 영수증 — 클립보드에 itanet/marketbom URL 또는 다중행 텍스트 복사 시

Usage:
  python execution/receipt_watcher.py
  python execution/receipt_watcher.py --dry-run --console
  python execution/receipt_watcher.py --watch-dir "D:\\카카오톡"
  python execution/receipt_watcher.py --date 2026-03-01
"""

import argparse
import hashlib
import json
import os
import re
import sys
import threading
import time
import ctypes
import ctypes.wintypes
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse

from dotenv import load_dotenv

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(os.path.join(PROJECT_ROOT, 'execution'))
load_dotenv(os.path.join(PROJECT_ROOT, '.env'))

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

from parse_receipt import (
    parse_receipt, map_to_danga_items, load_supplier_order_items,
    detect_supplier_from_url, detect_supplier_from_text,
    RECEIPT_MAP, RECEIPT_TYPE
)
from fill_receipt_prices import fill_receipt_prices

# ==================== 설정 ====================

IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp'}
FILE_AGE_LIMIT = 300  # 최근 5분 이내 파일만 처리
DEBOUNCE_SEC = 1.0    # 파일 쓰기 완료 대기 시간

# 클립보드 URL 패턴 (ERP 영수증)
URL_PATTERNS = ['itanet.co.kr', 'marketbom.com']

# 클립보드 텍스트 영수증 최소 줄 수
MIN_TEXT_LINES = 3
# 줄에 가격이 있는지 판단하는 패턴
PRICE_LINE_RE = re.compile(r'[\d,]{3,}')

# Win32 상수
WM_CLIPBOARDUPDATE = 0x031D
CF_UNICODETEXT = 13


# ==================== 카카오톡 다운로드 폴더 탐지 ====================

def find_kakao_download_dir():
    """카카오톡 다운로드 폴더 자동 탐지."""
    user_profile = os.environ.get('USERPROFILE', '')
    candidates = [
        os.path.join(user_profile, 'Documents', '카카오톡 받은 파일'),
        os.path.join(user_profile, 'Documents', 'KakaoTalk Downloads'),
        os.path.join(user_profile, '문서', '카카오톡 받은 파일'),
        os.path.join(user_profile, 'OneDrive', '문서', '카카오톡 받은 파일'),
        os.path.join(user_profile, 'OneDrive', 'Documents', '카카오톡 받은 파일'),
    ]
    for path in candidates:
        if os.path.isdir(path):
            return path
    return None


# ==================== 중복 방지 ====================

class DuplicateGuard:
    """세션 내 중복 방지 (해시 기반)."""

    def __init__(self):
        self._hashes = set()
        self._lock = threading.Lock()

    def check_and_add(self, data: bytes) -> bool:
        """중복이면 True, 새로운 항목이면 False 반환 후 등록."""
        h = hashlib.md5(data).hexdigest()
        with self._lock:
            if h in self._hashes:
                return True
            self._hashes.add(h)
            return False

    def check_text(self, text: str) -> bool:
        return self.check_and_add(text.encode('utf-8'))

    def check_file(self, filepath: str) -> bool:
        try:
            with open(filepath, 'rb') as f:
                return self.check_and_add(f.read())
        except (OSError, IOError):
            return False


# ==================== 알림 ====================

def toast_notify(title, message):
    """Windows 토스트 알림."""
    try:
        from winotify import Notification
        toast = Notification(
            app_id="영수증 감시",
            title=title,
            msg=message[:200],
        )
        toast.show()
    except Exception as e:
        print(f"[TOAST ERROR] {e}")


def kakao_notify(text):
    """카카오톡 나에게보내기."""
    token_path = os.path.join(PROJECT_ROOT, 'kakao_token.json')
    if not os.path.exists(token_path):
        print("[KAKAO] kakao_token.json not found, skip notification")
        return False

    try:
        with open(token_path, 'r') as f:
            tokens = json.load(f)

        # 토큰 갱신
        rest_api_key = os.environ.get('KAKAO_REST_API_KEY')
        if rest_api_key and tokens.get('refresh_token'):
            import requests
            res = requests.post("https://kauth.kakao.com/oauth/token", data={
                "grant_type": "refresh_token",
                "client_id": rest_api_key,
                "refresh_token": tokens['refresh_token']
            }, timeout=10)
            new_tokens = res.json()
            if 'access_token' in new_tokens:
                tokens['access_token'] = new_tokens['access_token']
                if 'refresh_token' in new_tokens:
                    tokens['refresh_token'] = new_tokens['refresh_token']
                with open(token_path, 'w') as f:
                    json.dump(tokens, f)

        # 메시지 전송
        import requests
        url = "https://kapi.kakao.com/v2/api/talk/memo/default/send"
        headers = {"Authorization": f"Bearer {tokens['access_token']}"}
        payload = {
            "template_object": json.dumps({
                "object_type": "text",
                "text": text[:800],
                "link": {
                    "web_url": "https://drive.google.com",
                    "mobile_web_url": "https://drive.google.com"
                },
                "button_title": "Drive 열기"
            })
        }
        res = requests.post(url, headers=headers, data=payload, timeout=10)
        if res.json().get('result_code') == 0:
            print("[KAKAO] Message sent")
            return True
        else:
            print(f"[KAKAO] Send failed: {res.json()}")
            return False
    except Exception as e:
        print(f"[KAKAO ERROR] {e}")
        return False


# ==================== 메인 워처 클래스 ====================

class ReceiptWatcher:
    """영수증 자동 감시 데몬."""

    def __init__(self, master_file, date_str, dry_run=False,
                 watch_dir=None, console=False):
        self.master_file = master_file
        self.date_str = date_str
        self.dry_run = dry_run
        self.console = console
        self.guard = DuplicateGuard()
        self.running = False
        self._lock = threading.Lock()
        self._today_results = []

        # 카카오톡 다운로드 폴더
        self.watch_dir = watch_dir or os.environ.get('KAKAO_DOWNLOAD_DIR') or find_kakao_download_dir()

        print(f"[INIT] Master: {self.master_file}")
        print(f"[INIT] Date: {self.date_str}")
        print(f"[INIT] Dry-run: {self.dry_run}")
        print(f"[INIT] Watch dir: {self.watch_dir}")

    # -------------------- 폴더 감시 (watchdog) --------------------

    def start_folder_watcher(self):
        """카카오톡 다운로드 폴더에 새 이미지 파일 생성 감지."""
        if not self.watch_dir:
            print("[FOLDER] No KakaoTalk download directory found, folder watcher disabled")
            return

        if not os.path.isdir(self.watch_dir):
            print(f"[FOLDER] Directory not found: {self.watch_dir}, folder watcher disabled")
            return

        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler

        watcher = self

        class ImageHandler(FileSystemEventHandler):
            def on_created(self, event):
                if event.is_directory:
                    return
                filepath = event.src_path
                ext = os.path.splitext(filepath)[1].lower()
                if ext not in IMAGE_EXTENSIONS:
                    return
                # debounce: 파일 쓰기 완료 대기
                threading.Timer(DEBOUNCE_SEC, watcher._on_new_image, args=[filepath]).start()

        observer = Observer()
        observer.schedule(ImageHandler(), self.watch_dir, recursive=False)
        observer.daemon = True
        observer.start()
        print(f"[FOLDER] Watching: {self.watch_dir}")
        return observer

    def _on_new_image(self, filepath):
        """새 이미지 감지 시 처리."""
        try:
            if not os.path.exists(filepath):
                return

            # 최근 파일만 (오래된 파일 무시) — ctime=생성시간 사용 (복사 시 갱신됨)
            ctime = os.path.getctime(filepath)
            age = time.time() - ctime
            if age > FILE_AGE_LIMIT:
                print(f"[FOLDER] Skipped old file ({age:.0f}s): {os.path.basename(filepath)}")
                return

            # 중복 체크
            if self.guard.check_file(filepath):
                print(f"[FOLDER] Duplicate: {os.path.basename(filepath)}")
                return

            print(f"\n[FOLDER] New image: {os.path.basename(filepath)}")

            # Gemini Vision OCR (공급처 자동 식별)
            result = parse_receipt(image_path=filepath)

            if 'error' in result:
                print(f"[FOLDER] Parse error: {result['error']}")
                toast_notify("영수증 파싱 실패", result['error'])
                return

            supplier = result.get('supplier')
            items = result.get('items', [])
            confidence = result.get('confidence', 'unknown')

            if not supplier:
                print(f"[FOLDER] Supplier not identified, items: {len(items)}")
                toast_notify("공급처 미식별", f"이미지에서 {len(items)}개 품목 감지, 공급처 확인 필요")
                return

            if not items:
                print(f"[FOLDER] No items found from {supplier}")
                toast_notify(f"{supplier} 영수증", "품목을 찾지 못했습니다")
                return

            print(f"[FOLDER] Supplier: {supplier}, Items: {len(items)}, Confidence: {confidence}")
            for item in items:
                print(f"  {item.get('name')}: {item.get('unit_price', 0):,}원")

            # 단가시트 입력
            self._apply_receipt(result, f"image:{os.path.basename(filepath)}")

        except Exception as e:
            print(f"[FOLDER ERROR] {e}")
            import traceback
            traceback.print_exc()

    # -------------------- 클립보드 감시 (Win32 API) --------------------

    def start_clipboard_watcher(self):
        """클립보드 변경 실시간 감시 (Win32 메시지 루프)."""
        thread = threading.Thread(target=self._clipboard_loop, daemon=True)
        thread.start()
        print("[CLIPBOARD] Watcher started")
        return thread

    def _clipboard_loop(self):
        """Win32 불가시 윈도우로 클립보드 변경 이벤트 수신."""
        import win32gui
        import win32con
        import win32api

        # wndproc을 딕셔너리 형태로 전달 (pywin32 권장 방식)
        wndproc = {
            WM_CLIPBOARDUPDATE: self._clipboard_wndproc_handler,
        }

        wndclass = win32gui.WNDCLASS()
        wndclass.lpfnWndProc = wndproc
        wndclass.lpszClassName = f"ReceiptClipboardWatcher_{id(self)}"
        wndclass.hInstance = win32api.GetModuleHandle(None)

        atom = win32gui.RegisterClass(wndclass)
        hwnd = win32gui.CreateWindow(
            atom, "ReceiptClipboardWatcher",
            0, 0, 0, 0, 0, 0, 0, wndclass.hInstance, None
        )

        # AddClipboardFormatListener
        user32 = ctypes.windll.user32
        if not user32.AddClipboardFormatListener(hwnd):
            print("[CLIPBOARD] AddClipboardFormatListener failed")
            return

        print(f"[CLIPBOARD] Listening (hwnd={hwnd})")

        # 메시지 루프 (win32gui.PumpMessages)
        win32gui.PumpMessages()

    def _clipboard_wndproc_handler(self, hwnd, msg, wparam, lparam):
        """WM_CLIPBOARDUPDATE 핸들러."""
        # 별도 스레드에서 처리 (메시지 루프 블로킹 방지)
        threading.Thread(target=self._on_clipboard_change, daemon=True).start()
        return 0

    def _on_clipboard_change(self):
        """클립보드 변경 시 내용 분석."""
        time.sleep(0.3)  # 클립보드 안정화 대기 (다른 앱과 경합 방지)

        import win32clipboard

        # OpenClipboard 재시도 (다른 앱이 잡고 있을 수 있음)
        opened = False
        for attempt in range(3):
            try:
                win32clipboard.OpenClipboard()
                opened = True
                break
            except Exception:
                time.sleep(0.2)

        if not opened:
            return  # 클립보드 접근 불가 — 무시

        try:
            # 텍스트 확인
            if win32clipboard.IsClipboardFormatAvailable(CF_UNICODETEXT):
                text = win32clipboard.GetClipboardData(CF_UNICODETEXT)
                win32clipboard.CloseClipboard()
                opened = False
                if text:
                    self._on_clipboard_text(text.strip())
                return

            # 이미지 확인 (CF_DIB)
            CF_DIB = 8
            if win32clipboard.IsClipboardFormatAvailable(CF_DIB):
                win32clipboard.CloseClipboard()
                opened = False
                self._on_clipboard_image()
                return
        except Exception as e:
            if '1418' not in str(e):
                print(f"[CLIPBOARD ERROR] {e}")
        finally:
            if opened:
                try:
                    win32clipboard.CloseClipboard()
                except Exception:
                    pass

    def _on_clipboard_text(self, text):
        """클립보드 텍스트 분석."""
        # 빈 텍스트/짧은 텍스트 무시
        if len(text) < 10:
            return

        # 1) URL 패턴 (itanet/marketbom)
        url_match = re.search(r'https?://[^\s]+', text)
        if url_match:
            url = url_match.group(0)
            parsed = urlparse(url)
            domain = parsed.netloc.lower()
            if any(p in domain for p in URL_PATTERNS):
                if self.guard.check_text(url):
                    return
                print(f"\n[CLIPBOARD] URL detected: {url}")
                result = parse_receipt(url=url)
                if 'error' not in result:
                    self._apply_receipt(result, f"url:{domain}")
                else:
                    print(f"[CLIPBOARD] URL parse error: {result['error']}")
                    toast_notify("URL 파싱 실패", result['error'])
                return

        # 2) 텍스트 영수증 (여러 줄 + 가격 패턴)
        lines = text.strip().split('\n')
        if len(lines) < MIN_TEXT_LINES:
            return

        # 가격이 포함된 줄이 절반 이상이면 영수증으로 판단
        price_lines = sum(1 for l in lines if PRICE_LINE_RE.search(l))
        if price_lines < len(lines) * 0.4:
            return

        if self.guard.check_text(text):
            return

        # 공급처 식별 시도
        supplier = detect_supplier_from_text(text)
        if not supplier:
            # 텍스트 영수증인데 공급처 불명 → 무시 (너무 많은 false positive)
            return

        print(f"\n[CLIPBOARD] Text receipt detected ({supplier}, {len(lines)} lines)")
        result = parse_receipt(text=text, supplier_hint=supplier)
        if 'error' not in result:
            self._apply_receipt(result, f"text:{supplier}")
        else:
            print(f"[CLIPBOARD] Text parse error: {result['error']}")

    def _on_clipboard_image(self):
        """클립보드 이미지 분석."""
        try:
            from PIL import ImageGrab
            img = ImageGrab.grabclipboard()
            if img is None:
                return

            # 이미지를 임시 파일로 저장
            import tempfile
            tmp = os.path.join(tempfile.gettempdir(), 'receipt_clipboard.png')
            img.save(tmp, 'PNG')

            if self.guard.check_file(tmp):
                return

            print(f"\n[CLIPBOARD] Image detected (clipboard)")

            result = parse_receipt(image_path=tmp)
            if 'error' not in result and result.get('items'):
                self._apply_receipt(result, "clipboard_image")
            elif 'error' in result:
                print(f"[CLIPBOARD] Image parse error: {result['error']}")

        except Exception as e:
            print(f"[CLIPBOARD IMAGE ERROR] {e}")

    # -------------------- 결과 적용 --------------------

    def _apply_receipt(self, result, source):
        """파싱 결과를 단가시트에 입력 + 알림."""
        supplier = result.get('supplier', '?')
        items = result.get('items', [])
        confidence = result.get('confidence', 'unknown')

        if not items:
            print(f"[APPLY] {supplier}: no items")
            return

        # 손글씨 공급처면 발주리스트 교차매칭
        order_items = None
        if RECEIPT_TYPE.get(supplier) == 'handwritten':
            order_items = load_supplier_order_items(supplier, self.date_str)

        # 매핑 미리보기
        mapped = map_to_danga_items(result, supplier, order_items=order_items)
        mapped_count = sum(1 for m in mapped if not m.get('skip_reason'))

        print(f"[APPLY] {supplier}: {len(items)} parsed → {mapped_count} mappable")

        # 단가시트 입력
        try:
            fill_result = fill_receipt_prices(
                master_file=self.master_file,
                target_date=self.date_str,
                receipts=[result],
                dry_run=self.dry_run,
            )
        except Exception as e:
            print(f"[APPLY ERROR] {e}")
            import traceback
            traceback.print_exc()
            toast_notify(f"{supplier} 입력 실패", str(e)[:150])
            return

        updates = fill_result.get('updates', 0)
        warnings = fill_result.get('warnings', [])
        details = fill_result.get('details', [])

        # 결과 기록
        entry = {
            "time": datetime.now().strftime("%H:%M:%S"),
            "supplier": supplier,
            "source": source,
            "confidence": confidence,
            "updates": updates,
            "warnings": len(warnings),
            "dry_run": self.dry_run,
        }
        with self._lock:
            self._today_results.append(entry)

        # 알림
        mode = "[DRY] " if self.dry_run else ""
        summary = f"{mode}{supplier} 매입가 {updates}건 입력"
        if warnings:
            summary += f" (경고 {len(warnings)}건)"
        if confidence == 'low':
            summary += " ⚠확인필요"

        detail_text = "\n".join(details[:5])
        if len(details) > 5:
            detail_text += f"\n... 외 {len(details)-5}건"

        toast_notify(summary, detail_text or "처리 완료")

        # 카카오톡 알림 (dry-run이 아닐 때만)
        if not self.dry_run and updates > 0:
            kakao_text = f"[영수증 자동입력]\n{summary}\n{detail_text}"
            kakao_notify(kakao_text)

        print(f"[DONE] {summary}")

    # -------------------- 시스템 트레이 --------------------

    def start_tray(self):
        """시스템 트레이 아이콘."""
        import pystray
        from PIL import Image, ImageDraw

        # 간단한 아이콘 생성 (녹색 원)
        size = 64
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.ellipse([4, 4, size-4, size-4], fill='#4CAF50' if not self.dry_run else '#FF9800')
        draw.text((size//4, size//4), "R", fill='white')

        def get_status(item):
            count = len(self._today_results)
            return f"오늘 처리: {count}건"

        def toggle_dry_run(icon, item):
            self.dry_run = not self.dry_run
            mode = "DRY-RUN" if self.dry_run else "LIVE"
            print(f"[TRAY] Mode: {mode}")
            toast_notify("모드 변경", f"현재: {mode}")

        def show_results(icon, item):
            if not self._today_results:
                toast_notify("처리 내역", "오늘 처리된 영수증이 없습니다")
                return
            lines = []
            for r in self._today_results[-10:]:
                prefix = "[DRY] " if r['dry_run'] else ""
                lines.append(f"{r['time']} {prefix}{r['supplier']}: {r['updates']}건")
            toast_notify(f"오늘 처리 내역 ({len(self._today_results)}건)", "\n".join(lines))

        def quit_app(icon, item):
            print("[TRAY] Quit requested")
            self.running = False
            icon.stop()

        menu = pystray.Menu(
            pystray.MenuItem(get_status, None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("처리 내역 보기", show_results),
            pystray.MenuItem(
                lambda item: f"{'☑' if self.dry_run else '☐'} Dry-run 모드",
                toggle_dry_run,
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("종료", quit_app),
        )

        title = "영수증 감시 (DRY)" if self.dry_run else "영수증 감시"
        icon = pystray.Icon("receipt_watcher", img, title, menu)
        icon.run_detached()
        print("[TRAY] System tray icon started")
        return icon

    # -------------------- 메인 실행 --------------------

    def run(self):
        """감시 시작."""
        self.running = True

        # 폴더 감시 시작
        observer = self.start_folder_watcher()

        # 클립보드 감시 시작
        clip_thread = self.start_clipboard_watcher()

        # 트레이 (콘솔 모드가 아닐 때)
        tray_icon = None
        if not self.console:
            try:
                tray_icon = self.start_tray()
            except Exception as e:
                print(f"[TRAY ERROR] {e}, running in console mode")

        mode = "DRY-RUN" if self.dry_run else "LIVE"
        print(f"\n{'='*50}")
        print(f"  영수증 자동 감시 실행 중 ({mode})")
        print(f"  날짜: {self.date_str}")
        print(f"  마스터: {os.path.basename(self.master_file)}")
        if self.watch_dir:
            print(f"  폴더 감시: {self.watch_dir}")
        print(f"  클립보드 감시: 활성")
        print(f"  종료: Ctrl+C")
        print(f"{'='*50}\n")

        toast_notify("영수증 감시 시작",
                     f"{mode} | 폴더: {'O' if self.watch_dir else 'X'} | 클립보드: O")

        try:
            while self.running:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n[STOP] Shutting down...")
        finally:
            self.running = False
            if observer:
                observer.stop()
                observer.join(timeout=3)
            if tray_icon:
                try:
                    tray_icon.stop()
                except Exception:
                    pass

        # 종료 요약
        count = len(self._today_results)
        if count > 0:
            print(f"\n[SUMMARY] 오늘 총 {count}건 처리:")
            for r in self._today_results:
                prefix = "[DRY] " if r['dry_run'] else ""
                print(f"  {r['time']} {prefix}{r['supplier']}: {r['updates']}건")


# ==================== CLI ====================

def main():
    parser = argparse.ArgumentParser(description='영수증 자동 감지 데몬')
    parser.add_argument('--master', help='마스터 파일 경로 (기본: .env MASTER_FILE_PATH)')
    parser.add_argument('--date', help='대상 날짜 YYYY-MM-DD (기본: 오늘)')
    parser.add_argument('--dry-run', action='store_true', help='실제 입력 없이 결과만 표시')
    parser.add_argument('--console', action='store_true', help='시스템 트레이 없이 콘솔 모드')
    parser.add_argument('--watch-dir', help='카카오톡 다운로드 폴더 경로 (기본: 자동 탐지)')
    args = parser.parse_args()

    # 마스터 파일 경로
    master = args.master or os.environ.get('MASTER_FILE_PATH')
    if not master:
        master = "G:/내 드라이브/1. 도크_주문 명세서/0. 매입단가_자료/도크발주관리데이터.xlsx"
    if not os.path.exists(master):
        print(f"Error: Master file not found: {master}")
        print("Set MASTER_FILE_PATH in .env or use --master")
        sys.exit(1)

    # 날짜
    date_str = args.date or datetime.now().strftime('%Y-%m-%d')

    watcher = ReceiptWatcher(
        master_file=master,
        date_str=date_str,
        dry_run=args.dry_run,
        watch_dir=args.watch_dir,
        console=args.console,
    )
    watcher.run()


if __name__ == '__main__':
    main()
