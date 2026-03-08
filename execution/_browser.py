"""Playwright 브라우저 공통 팩토리.

모든 스크래핑 스크립트가 이 모듈을 통해 브라우저를 생성.
첫 실행 시 Chromium 자동 설치.
Streamlit Cloud 환경 자동 감지 및 대응.
"""
import os
import subprocess
import sys


def _is_streamlit_cloud() -> bool:
    """Streamlit Cloud 환경 감지."""
    return os.path.exists("/mount/src") or os.environ.get("STREAMLIT_SHARING_MODE") == "1"


def _setup_browser_path():
    """Streamlit Cloud에서 브라우저 경로를 쓰기 가능한 위치로 설정."""
    if _is_streamlit_cloud():
        browser_path = "/home/adminuser/.cache/ms-playwright"
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = browser_path
        os.makedirs(browser_path, exist_ok=True)


def ensure_browser():
    """Playwright Chromium이 없으면 설치. 앱 최초 실행 시 1회."""
    _setup_browser_path()

    try:
        from playwright.sync_api import sync_playwright
        pw = sync_playwright().start()
        try:
            browser = pw.chromium.launch(headless=True)
            browser.close()
        except Exception:
            # Chromium not installed yet
            _install_chromium()
        finally:
            pw.stop()
    except ImportError:
        raise RuntimeError(
            "playwright 패키지가 설치되지 않았습니다. "
            "pip install playwright 후 재시도하세요."
        )
    except Exception:
        _install_chromium()


def _install_chromium():
    """Chromium 브라우저 설치 (with-deps 옵션으로 시스템 의존성도 함께)."""
    cmd = [sys.executable, "-m", "playwright", "install", "chromium"]
    if _is_streamlit_cloud():
        # Streamlit Cloud: 시스템 의존성도 함께 설치 시도
        cmd = [sys.executable, "-m", "playwright", "install", "--with-deps", "chromium"]

    try:
        result = subprocess.run(
            cmd, check=True,
            capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
        print(f"[Browser] Chromium 설치 완료", flush=True)
    except subprocess.CalledProcessError as e:
        stderr = e.stderr or ""
        raise RuntimeError(
            f"Chromium 브라우저 설치 실패.\n"
            f"환경: {'Streamlit Cloud' if _is_streamlit_cloud() else 'Local'}\n"
            f"오류: {stderr[-500:]}"
        )


def create_browser(headless=True, mobile=False):
    """Playwright browser + context + page 생성.

    Args:
        headless: headless 모드 (Cloud=True, 로컬 디버깅=False)
        mobile: 모바일 에뮬레이션 (식봄용)

    Returns:
        (pw, browser, page) 튜플. 종료 시 browser.close(); pw.stop() 필수.
    """
    _setup_browser_path()
    from playwright.sync_api import sync_playwright

    pw = sync_playwright().start()

    # Streamlit Cloud: 메모리 절약 args
    launch_args = []
    if _is_streamlit_cloud():
        launch_args = [
            "--disable-dev-shm-usage",
            "--no-sandbox",
            "--disable-gpu",
            "--single-process",
        ]

    browser = pw.chromium.launch(headless=headless, args=launch_args)

    if mobile:
        context = browser.new_context(
            viewport={"width": 375, "height": 812},
            device_scale_factor=3.0,
            user_agent=(
                "Mozilla/5.0 (iPhone; CPU iPhone OS 13_2_3 like Mac OS X) "
                "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0.3 "
                "Mobile/15E148 Safari/604.1"
            ),
        )
    else:
        context = browser.new_context(viewport={"width": 1920, "height": 1080})

    page = context.new_page()
    return pw, browser, page
