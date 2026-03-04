from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import time
import sys
import argparse

# Configuration
USERNAME = "juillte@naver.com"
PASSWORD = "Tfe7c1p4!!"

# Priorities
SUPPLIER_PRIORITY = [
    "BJ프레시웨이", # Note: User said CJ, but let's be robust. Map CJ -> BJ? Or is it BJ?
    "다봄푸드",
    "세현F&B",
    "온국민신선몰"
]
# Fixing typo for priority check if needed, but assuming user meant CJ Freshway.
# Actually in previous logs I saw 'CJ프레시웨이'. I will use exact strings.
TARGET_SUPPLIERS = ["CJ프레시웨이", "다봄푸드", "세현F&B", "온국민신선몰"] 

# Configure Chrome options
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Mobile Emulation
mobile_emulation = {
    "deviceMetrics": { "width": 375, "height": 812, "pixelRatio": 3.0 },
    "userAgent": "Mozilla/5.0 (iPhone; CPU iPhone OS 13_2_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0.3 Mobile/15E148 Safari/604.1"
}
chrome_options.add_experimental_option("mobileEmulation", mobile_emulation)

class FoodspringScraper:
    def __init__(self):
        print("Initializing Scraper...")
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 10)

    def login(self):
        try:
            print("Logging in...", flush=True)
            self.driver.get("https://www.foodspring.co.kr/login")
            
            # Click Email Login
            btn = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[aria-label='이메일_로그인_버튼']")))
            btn.click()
            time.sleep(1)
            
            # Inputs
            time.sleep(2)
            
            email_input = None
            try:
                email_input = self.driver.find_element(By.CSS_SELECTOR, "input[type='email']")
                print("Found email input (type=email).", flush=True)
            except:
                pass
                
            if not email_input:
                try:
                    inputs = self.driver.find_elements(By.CSS_SELECTOR, "input[type='text']")
                    if inputs:
                        email_input = inputs[0]
                        print("Found email input (type=text).", flush=True)
                except:
                    pass
            
            if not email_input:
                raise Exception("Could not find email input")

            email_input.clear()
            email_input.send_keys(USERNAME)
            
            time.sleep(1)
            pw_input = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']")))
            pw_input.clear()
            pw_input.send_keys(PASSWORD)
            
            # Submit
            pw_input.send_keys(u'\ue007')
            
            time.sleep(5)
            print(f"Login complete. Current URL: {self.driver.current_url}", flush=True)
            self.driver.save_screenshot("debug_scraper_login.png")
            return True
        except Exception as e:
            print(f"Login failed: {e}", flush=True)
            self.driver.save_screenshot("debug_scraper_login_fail.png")
            return False

    def search_item(self, keyword):
        try:
            print(f"Searching for: {keyword}", flush=True)
            url = "https://www.foodspring.co.kr/search"
            
            # Avoid reloading if already on search page
            if self.driver.current_url.split('?')[0] != url:
                self.driver.get(url)
                # Replacing sleep with wait for input
                # time.sleep(3) 
            
            # Find search input - use explicit wait
            search_input = None
            try:
                search_input = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='search']")))
            except:
                try:
                    search_input = self.driver.find_element(By.CSS_SELECTOR, "input[type='text']")
                except:
                    pass
            
            if search_input:
                # print("Found search input. Typing keyword...", flush=True)
                search_input.clear()
                search_input.send_keys(keyword)
                search_input.send_keys(u'\ue007')
                # minimal wait for enter
                time.sleep(0.5) 
            else:
                print("Could not find search input.", flush=True)
                return None, None
            
            # Reduce wait time for results (AJAX)
            # A more robust way would be to wait for a change in results, 
            # but since we don't have a reliable "loading" indicator, we'll just reduce the sleep
            # and rely on the fact that we clear the input.
            # Actually, waiting 3.5s is safer than 5.
            time.sleep(3.5) 
            
            # print(f"  Saving debug screenshot for {keyword}...", flush=True)
            # self.driver.save_screenshot(f"debug_search_{keyword}.png") 
            
            # Dump source
            # with open(f"debug_search_{keyword}.html", "w", encoding="utf-8") as f:
            #    f.write(self.driver.page_source)
            
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            # page_text = soup.get_text()
            # print(f"Page Text Len: {len(page_text)}", flush=True)
            
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
                            # Pattern 1: "XX% YY,YYY원" - discount percentage followed by price
                            discount_match = re.search(r'(\d{1,2})%\s*([\d,]+)원', text)
                            
                            # Pattern 2: Last price in the text (usually member price)
                            all_prices = re.findall(r'([\d,]+)원', text)
                            
                            price = None
                            
                            if discount_match:
                                # Found discounted price
                                price_str = discount_match.group(2).replace(',', '')
                                price = int(price_str)
                            elif all_prices and len(all_prices) >= 2:
                                # If multiple prices, take the second one (discounted) or middle one
                                # Usually format is: OriginalPrice DiscountedPrice CouponPrice
                                # We want the middle one (discounted/member price)
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
                                except:
                                    pass
            
            return best_price, best_supplier

        except Exception as e:
            print(f"Search failed for {keyword}: {e}", flush=True)
            return None, None


    def close(self):
        self.driver.quit()

if __name__ == "__main__":
    scraper = FoodspringScraper()
    if scraper.login():
        # Test search
        test_item = "양파"
        p, s = scraper.search_item(test_item)
        print(f"Final Result for {test_item}: {p} / {s}")
    scraper.close()
