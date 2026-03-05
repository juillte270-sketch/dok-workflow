"""Playwright 브라우저 공통 팩토리.

모든 스크래핑 스크립트가 이 모듈을 통해 브라우저를 생성.
첫 실행 시 Chromium 자동 설치.
"""
import subprocess
import sys


def ensure_browser():
    """Playwright Chromium이 없으면 설치. 앱 최초 실행 시 1회."""
    try:
        from playwright.sync_api import sync_playwright
        pw = sync_playwright().start()
        try:
            browser = pw.chromium.launch(headless=True)
            browser.close()
        except Exception:
            # Chromium not installed yet
            subprocess.run(
                [sys.executable, "-m", "playwright", "install", "chromium"],
                check=True,
            )
        finally:
            pw.stop()
    except Exception:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=True,
        )


def create_browser(headless=True, mobile=False):
    """Playwright browser + context + page 생성.

    Args:
        headless: headless 모드 (Cloud=True, 로컬 디버깅=False)
        mobile: 모바일 에뮬레이션 (식봄용)

    Returns:
        (pw, browser, page) 튜플. 종료 시 browser.close(); pw.stop() 필수.
    """
    from playwright.sync_api import sync_playwright

    pw = sync_playwright().start()
    browser = pw.chromium.launch(headless=headless)

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
