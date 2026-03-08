from bs4 import BeautifulSoup
import sys
import os
import argparse

# Windows cp949 인코딩 에러 방지 — 식봄 검색 결과에 이모지/특수문자 포함됨
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# Configuration
USERNAME = "juillte@naver.com"
PASSWORD = "Tfe7c1p4!!"

# Priorities
SUPPLIER_PRIORITY = [
    "BJ프레시웨이",
    "다봄푸드",
    "세현F&B",
    "온국민신선몰"
]
TARGET_SUPPLIERS = ["CJ프레시웨이", "다봄푸드", "세현F&B", "온국민신선몰"]


class FoodspringScraper:
    def __init__(self):
        print("Initializing Scraper...", flush=True)
        # Ensure Chromium is installed
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from _browser import ensure_browser, create_browser
        ensure_browser()
        self._pw, self._browser, self.page = create_browser(headless=True, mobile=True)

    def login(self):
        try:
            print("Logging in...", flush=True)
            self.page.goto("https://www.foodspring.co.kr/login", wait_until="networkidle")

            # Click Email Login
            self.page.locator("button[aria-label='이메일_로그인_버튼']").wait_for(state="visible", timeout=10000)
            self.page.locator("button[aria-label='이메일_로그인_버튼']").click()
            self.page.wait_for_timeout(1000)

            # Find email input
            self.page.wait_for_timeout(2000)

            email_input = None
            try:
                loc = self.page.locator("input[type='email']")
                if loc.count() > 0:
                    email_input = loc
                    print("Found email input (type=email).", flush=True)
            except Exception:
                pass

            if not email_input:
                try:
                    loc = self.page.locator("input[type='text']")
                    if loc.count() > 0:
                        email_input = loc.first
                        print("Found email input (type=text).", flush=True)
                except Exception:
                    pass

            if not email_input:
                raise Exception("Could not find email input")

            email_input.fill(USERNAME)

            self.page.wait_for_timeout(1000)
            pw_input = self.page.locator("input[type='password']")
            pw_input.wait_for(state="visible", timeout=10000)
            pw_input.fill(PASSWORD)

            # Submit
            pw_input.press("Enter")

            self.page.wait_for_timeout(5000)
            print(f"Login complete. Current URL: {self.page.url}", flush=True)
            self.page.screenshot(path="debug_scraper_login.png")
            return True
        except Exception as e:
            print(f"Login failed: {e}", flush=True)
            self.page.screenshot(path="debug_scraper_login_fail.png")
            return False

    def search_item(self, keyword):
        try:
            print(f"Searching for: {keyword}", flush=True)
            url = "https://www.foodspring.co.kr/search"

            # Avoid reloading if already on search page
            if self.page.url.split('?')[0] != url:
                self.page.goto(url, wait_until="networkidle")

            # Find search input
            search_input = None
            try:
                loc = self.page.locator("input[type='search']")
                loc.wait_for(state="visible", timeout=10000)
                search_input = loc
            except Exception:
                try:
                    loc = self.page.locator("input[type='text']")
                    if loc.count() > 0:
                        search_input = loc.first
                except Exception:
                    pass

            if search_input:
                search_input.fill(keyword)
                search_input.press("Enter")
                self.page.wait_for_timeout(500)
            else:
                print("Could not find search input.", flush=True)
                return None, None

            # Wait for results
            self.page.wait_for_timeout(3500)

            soup = BeautifulSoup(self.page.content(), 'html.parser')

            best_match = None
            best_priority = 999
            best_price = None
            best_supplier = None

            for supplier in TARGET_SUPPLIERS:
                elements = soup.find_all(string=lambda t: t and supplier in t)
                if elements:
                    for el in elements:
                        container = el.find_parent(lambda tag: tag.name == 'div' or tag.name == 'li')
                        if container:
                            text = container.get_text()
                            import re

                            # Try to find discounted price first (format: "XX% YY,YYY원" or "회원가 YY,YYY원")
                            discount_match = re.search(r'(\d{1,2})%\s*([\d,]+)원', text)

                            # Pattern 2: Last price in the text (usually member price)
                            all_prices = re.findall(r'([\d,]+)원', text)

                            price = None

                            if discount_match:
                                price_str = discount_match.group(2).replace(',', '')
                                price = int(price_str)
                            elif all_prices and len(all_prices) >= 2:
                                price_str = all_prices[1].replace(',', '') if len(all_prices) > 1 else all_prices[0].replace(',', '')
                                price = int(price_str)
                            elif all_prices:
                                price_str = all_prices[0].replace(',', '')
                                price = int(price_str)

                            if price:
                                try:
                                    prio = TARGET_SUPPLIERS.index(supplier)
                                    if prio < best_priority:
                                        best_priority = prio
                                        best_price = price
                                        best_supplier = supplier
                                except Exception:
                                    pass

            return best_price, best_supplier

        except Exception as e:
            print(f"Search failed for {keyword}: {e}", flush=True)
            return None, None

    def close(self):
        self._browser.close()
        self._pw.stop()

if __name__ == "__main__":
    scraper = FoodspringScraper()
    if scraper.login():
        # Test search
        test_item = "양파"
        p, s = scraper.search_item(test_item)
        print(f"Final Result for {test_item}: {p} / {s}")
    scraper.close()
