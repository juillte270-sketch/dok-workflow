import sys
import json
import datetime
from playwright.sync_api import sync_playwright

def fetch_posts(username):
    url = f"https://www.threads.net/@{username}"
    print(f"Fetching posts from {url}...")
    
    data = []
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        
        try:
            page.goto(url, wait_until="domcontentloaded")
            # Wait for some content to load
            try:
                page.wait_for_selector("div[data-pressable-container='true']", timeout=10000)
            except:
                print("Timeout waiting for posts container.")
            
            # Scroll down a few times to load more posts
            for _ in range(3):
                page.mouse.wheel(0, 1000)
                page.wait_for_timeout(1000)
            
            # Extract posts
            posts = page.query_selector_all("div[data-pressable-container='true']")
            print(f"Found {len(posts)} potential post containers.")
            
            for post in posts:
                try:
                    # Extract Timestamp
                    time_el = post.query_selector("time")
                    if not time_el:
                        continue
                    
                    iso_time = time_el.get_attribute("datetime")
                    if not iso_time:
                        continue
                        
                    # Extract Text
                    # Join all span[dir='auto'] texts
                    text_els = post.query_selector_all("span[dir='auto']")
                    text = "\n".join([t.inner_text() for t in text_els]).strip()
                    
                    if not text:
                         # Fallback for text in div
                         text_div = post.query_selector("div[dir='auto']")
                         if text_div:
                             text = text_div.inner_text().strip()
                    
                    # Convert to unix timestamp for compatibility
                    # ISO format: 2024-01-27T06:00:00.000Z
                    # Python 3.11 supports formisoformat with Z, but older might not.
                    # safer: strict parsing or dateutil
                    
                    dt = datetime.datetime.fromisoformat(iso_time.replace('Z', '+00:00'))
                    taken_at = int(dt.timestamp())
                    
                    data.append({
                        "taken_at": taken_at,
                        "iso_time": iso_time,
                        "caption": {"text": text}, # Match previous structure roughly
                        "description": text
                    })
                    
                except Exception as e:
                    print(f"Error parsing post: {e}")
                    continue
                    
        except Exception as e:
            print(f"Error accessing page: {e}")
        finally:
            browser.close()
            
    # Save Data
    import os
    os.makedirs('.tmp', exist_ok=True)
    output_path = os.path.join('.tmp', 'threads_posts.json')
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"Saved {len(data)} posts to {output_path}")

if __name__ == "__main__":
    target_user = sys.argv[1] if len(sys.argv) > 1 else "choi.openai"
    # Remove @ if present
    target_user = target_user.replace("@", "")
    fetch_posts(target_user)
