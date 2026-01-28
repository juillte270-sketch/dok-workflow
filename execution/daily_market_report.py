import os
import json
import datetime
import requests
import feedparser
import yfinance as yf
import google.generativeai as genai
from newspaper import Article
from dotenv import load_dotenv

# Load Env
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
KAKAO_REST_API_KEY = os.getenv("KAKAO_REST_API_KEY")

genai.configure(api_key=GEMINI_API_KEY)

def get_market_data():
    """Fetches comprehensive market data."""
    tickers = {
        "S&P 500": "^GSPC",
        "Nasdaq": "^IXIC",
        "KOSPI": "^KS11",
        "KOSDAQ": "^KQ11",
        "VIX": "^VIX",         # Volatility
        "USD/KRW": "KRW=X",    # Exchange Rate
        "Gold": "GC=F",        # Commodity
        "10Y Bond": "^TNX"     # Interest Rate
    }
    
    data_summary = []
    regime_signals = {} # To help AI decide market regime
    
    for name, ticker in tickers.items():
        try:
            tok = yf.Ticker(ticker)
            hist = tok.history(period="5d") # Look back 5 days for trend
            
            if len(hist) > 0:
                current = hist['Close'].iloc[-1]
                prev = hist['Close'].iloc[-2] if len(hist) > 1 else current
                change_pct = ((current - prev) / prev) * 100
                
                # Simple moving average for trend context
                ma5 = hist['Close'].mean()
                trend = "Bullish" if current > ma5 else "Bearish"
                
                # Emojis
                icon = "⚪"
                if change_pct > 1.0: icon = "🔥" if name in ["VIX", "USD/KRW"] else "🚀" 
                elif change_pct < -1.0: icon = "💧" if name in ["VIX", "USD/KRW"] else "📉"
                elif change_pct > 0: icon = "🔺"
                elif change_pct < 0: icon = "🔻"

                data_str = f"{name}: {current:,.2f} ({change_pct:+.2f}%) {trend} {icon}"
                data_summary.append(data_str)
                regime_signals[name] = {"trend": trend, "change": change_pct}
        except Exception as e:
            data_summary.append(f"{name}: N/A")
            
    return "\n".join(data_summary), regime_signals

def get_detailed_news():
    """Fetches news and scrapes content for depth."""
    # Google News - Business - KR
    rss_url = "https://news.google.com/rss/topics/CAAqIggKIhxDQkFTRDgwL0ZRbndibU50Y3pKeWJtVjVibWNvQUFQAQ?hl=ko&gl=KR&ceid=KR%3Ako"
    
    feed = feedparser.parse(rss_url)
    
    news_content = []
    print(f"Found {len(feed.entries)} articles. Scraping top 5...")
    
    count = 0
    for entry in feed.entries:
        if count >= 5: break
        try:
            # Newspaper3k scraping
            article = Article(entry.link)
            article.download()
            article.parse()
            
            # Limit text length to avoid token limits
            text_snippet = article.text[:800].replace('\n', ' ') 
            
            news_item = f"""
            [TITLE]: {entry.title}
            [SOURCE]: {entry.source.title}
            [CONTENT]: {text_snippet}...
            """
            news_content.append(news_item)
            count += 1
        except Exception as e:
            print(f"Failed to scrape {entry.link}: {e}")
            continue
            
    return "\n".join(news_content)

import time

def send_kakao(text):
    """Sends a message using KakaoTalk API, splitting into chunks if too long."""
    try:
        if not os.path.exists('kakao_token.json'):
            print("Error: kakao_token.json not found.")
            return False

        with open('kakao_token.json', 'r') as f:
            tokens = json.load(f)
        
        url = "https://kapi.kakao.com/v2/api/talk/memo/default/send"
        headers = {"Authorization": f"Bearer {tokens['access_token']}"}
        
        # Split text into safe chunks (approx 800 chars to be safe for display)
        # We try to split by newlines to keep readability
        max_chunk_size = 800
        chunks = []
        current_chunk = ""
        
        for line in text.split('\n'):
            if len(current_chunk) + len(line) + 1 > max_chunk_size:
                chunks.append(current_chunk)
                current_chunk = line + "\n"
            else:
                current_chunk += line + "\n"
        
        if current_chunk:
            chunks.append(current_chunk)
            
        print(f"Sending {len(chunks)} messages...")
        
        for i, chunk in enumerate(chunks):
            # Add page number footer
            footer = f"\n({i+1}/{len(chunks)})"
            final_text = chunk + footer
            
            payload = {
                "template_object": json.dumps({
                    "object_type": "text",
                    "text": final_text,
                    "link": {
                        "web_url": "https://m.stock.naver.com",
                        "mobile_web_url": "https://m.stock.naver.com"
                    },
                    "button_title": "증시 더보기"
                })
            }
            
            res = requests.post(url, headers=headers, data=payload)
            
            # Simple Token Refresh Logic if needed (check logic in daily_market_report if copied fully)
            if res.status_code == 401:
                 print("Token expired during sending.")
                 # (Refresh logic omitted for brevity in chunking update, assuming fresh start usually)
                 return False

            if res.json().get('result_code') == 0:
                print(f"Message {i+1} sent.")
            else:
                print(f"Failed to send message {i+1}: {res.json()}")
            
            # Sleep to ensure order
            time.sleep(0.5)
            
        return True
    except Exception as e:
        print(f"KakaoTalk error: {e}")
        return False

def generate_report():
    print("1. Fetching Market Data...")
    market_data_str, regime_signals = get_market_data()
    
    print("2. Fetching & Scraping News (This may take a moment)...")
    detailed_news = get_detailed_news()
    
    print("3. analyzing with AI...")
    today_str = datetime.date.today().strftime("%Y년 %m월 %d일")
    
    # Advanced Prompt
    prompt = f"""
    You are a top-tier financial analyst (like 'Alpha-Prime OS'). 
    Analyze the provided MARKET DATA and DETAILED NEWS to create a high-quality "Morning Market Brief" for **이형기 (Hyung-ki Lee)**.
    Use a professional yet engaging tone (expert voice).
    
    TARGET FORMAT:
    
    # [Market Brief] {{Write a Provocative Title based on the biggest news}}
    **발행일**: {today_str}
    
    ## 📌 개요 (Today's Snapshot)
    - Synthesize the correlation between indices (e.g. "Nasdaq fell while Yields rose").
    - Mention the "Fear/Greed" sentiment inferred from VIX and News.
    
    ## 🚨 핵심 이슈 (Key Issues)
    - Select top 3 events.
    - Deep dive into ONE main event: explain the 'Why' and 'Impact'.
    
    ## 📊 Market Regime (판단)
    - **통화/금리**: (e.g. "Strong Dollar warning")
    - **변동성(VIX)**: (e.g. "Stable" or "Fear spike")
    - **종합 의견**: (e.g. "Neutral", "Buy the Dip", "Caution")
    
    ## 💎 오늘의 옥석 (Top Picks Idea)
    - Based on the news trends (e.g. AI, Energy, Bio), suggest 2-3 sectors or themes to watch.
    - (Do NOT recommend specific stock buy/sell actions, but highlight "Attention" stocks mentioned in news).
    
    ## 🎯 투자 전략 (Action Plan)
    - Suggest a stance: (e.g. "Hold Cash 30%", "Focus on Defensive", "Aggressive Tech").
    
    ---
    MARKET DATA:
    {market_data_str}
    
    NEWS CONTENT:
    {detailed_news}
    """
    
    # Try using the best available model
    try:
        models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        models.sort(key=lambda x: '1.5' in x, reverse=True) # Prefer 1.5
        best_model = models[0] if models else 'gemini-pro'
        print(f"Using AI Model: {best_model}")
        
        model = genai.GenerativeModel(best_model)
        response = model.generate_content(prompt)
        report = response.text
    except Exception as e:
        print(f"AI Error: {e}")
        report = "Analysis Failed."

    print("4. Sending Report...")
    send_kakao(report)

if __name__ == "__main__":
    generate_report()
