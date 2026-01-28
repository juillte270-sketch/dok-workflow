import os
import json
import sys
import datetime
import base64
from email.mime.text import MIMEText
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import google.generativeai as genai
from dotenv import load_dotenv
import requests
import json

# Load environment variables
load_dotenv()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
EMAIL_RECIPIENT = os.getenv("EMAIL_RECIPIENT")
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

if not GEMINI_API_KEY:
    print("Error: GEMINI_API_KEY not found in .env")
    sys.exit(1)

genai.configure(api_key=GEMINI_API_KEY)

def get_gmail_service():
    """Gets authenticated Gmail service."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            # Revert to local server
            creds = flow.run_local_server(port=8080)
        
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)

def send_email(subject, body, to):
    """Sends an email using Gmail API. Returns True if successful."""
    try:
        service = get_gmail_service()
        message = MIMEText(body)
        message['to'] = to
        message['subject'] = subject
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
        service.users().messages().send(userId='me', body={'raw': raw}).execute()
        print(f"Email sent to {to}")
        return True
    except Exception as e:
        print(f"Failed to send email (Auth or Network Error): {e}")
        with open('last_error.txt', 'w', encoding='utf-8') as f:
            f.write(str(e))
            import traceback
            traceback.print_exc(file=f)
        return False
        return False

def send_kakao(text):
    """Sends a message to self using KakaoTalk API."""
    try:
        with open('kakao_token.json', 'r') as f:
            tokens = json.load(f)
        
        url = "https://kapi.kakao.com/v2/api/talk/memo/default/send"
        headers = {
            "Authorization": f"Bearer {tokens['access_token']}"
        }
        
        # Simple text template
        # Kakao 'text' object type supports up to 200 characters for mobile display in some cases,
        # but the actual API limit is higher. We target around 1000 chars.
        
        safe_text = text[:1000] + ("..." if len(text) > 1000 else "")
        
        payload = {
            "template_object": json.dumps({
                "object_type": "text",
                "text": safe_text,
                "link": {
                    "web_url": "https://mail.google.com",
                    "mobile_web_url": "https://mail.google.com"
                },
                "button_title": "이메일 전체보기"
            })
        }
        
        res = requests.post(url, headers=headers, data=payload)
        if res.json().get('result_code') == 0:
            print("KakaoTalk sent successfully.")
            return True
        else:
            print(f"KakaoTalk failed: {res.json()}")
            return False
            
    except Exception as e:
        print(f"KakaoTalk error: {e}")
        return False

def save_report_locally(subject, content):
    """Saves the report to a markdown file."""
    filename = f"Threads_Report_{datetime.date.today()}.md"
    filepath = os.path.join(os.getcwd(), filename)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(f"# {subject}\n\n{content}")
    print(f"Report saved locally to: {filepath}")
    return filepath

def analyze_posts():
    """Reads posts, filters by 24h, summarizes, and sends email."""
    input_path = os.path.join('.tmp', 'threads_posts.json')
    if not os.path.exists(input_path):
        print(f"Error: File not found {input_path}")
        return

    with open(input_path, 'r', encoding='utf-8') as f:
        threads_data = json.load(f)

    # Filter posts from last 24 hours
    # Note: Threads data structure varies. Assuming standard 'taken_at' or similar timestamp.
    # We will look for keys like 'taken_at', 'created_at', or 'timestamp'.
    # If using threads-api, we need to inspect the structure.
    # For now, let's grab all text and let LLM filter if timestamp is ambiguous, 
    # OR try to parse known fields.
    
    # As a robust fallback, we'll extract ALL text and ask Gemini to "Identify posts from last 24h" 
    # but that consumes tokens. Better to filter if possible.
    # Inspecting common `threads-api` output: `thread_items` -> `post` -> `taken_at` (unix ts)
    
    filtered_posts = []
    now = datetime.datetime.now().timestamp()
    one_day_ago = now - 24 * 3600
    
    # Flatten structure if needed. 
    # Usually `threads_data` is a list of threads. Each thread has `thread_items`.
    # Make sure we handle list vs dict
    
    if isinstance(threads_data, dict) and 'threads' in threads_data:
         # api might return a dict with 'threads' key
         items_list = threads_data['threads']
    elif isinstance(threads_data, list):
         items_list = threads_data
    else:
         items_list = []

    post_texts = []
    
    for item in items_list:
        # Structure assumption: item might be a Post object or a Thread object containing posts
        # We try to find the timestamp. 
        # Let's interactively inspect via print in a real scenario, but here we assume 'taken_at' exists in post
        
        # Deep search for posts
        # This is a simplification.
        posts = []
        if 'thread_items' in item:
            for thread_item in item['thread_items']:
                 if 'post' in thread_item:
                     posts.append(thread_item['post'])
        elif 'post' in item: # direct post
             posts.append(item['post'])
        elif 'taken_at' in item: # it is the post
             posts.append(item)
             
        for post in posts:
            taken_at = post.get('taken_at')
            # taken_at is likely unix timestamp
            if taken_at and taken_at >= one_day_ago:
                text = post.get('caption', {}).get('text', '') if isinstance(post.get('caption'), dict) else post.get('caption')
                # Sometimes text is in 'caption' -> 'text'
                if not text:
                     # Check other fields
                     text = post.get('description', '') 
                
                if text:
                    post_texts.append(f"- [Time: {datetime.datetime.fromtimestamp(taken_at)}] {text}")

    if not post_texts:
        print("No posts found in the last 24 hours.")
        # Optional: Send "No updates" email
        return

    # Prepare Prompt
    posts_content = "\n".join(post_texts)
    prompt = f"""
    The following are Threads posts from the user @choi.openai from the last 24 hours.
    Please read them and provide a structured report in Korean.
    
    REPORT FORMAT (Optimized for Chat Message):
    - Use clear headings with emojis (e.g. 📌 요약, 💡 주요 내용).
    - Add an empty line between every section and bullet point for readability.
    - Keep sentences concise.
    
    1. **📌 요약 (Summary)**: 
       - 3-5 bullet points.
    
    2. **💡 주요 내용 (Key Details)**: 
       - Detail the most important posts.
    
    3. **🚀 인사이트 (Insights)**: 
       - What can we learn?
    
    POSTS:
    {posts_content}
    """

    print("Generating report with Gemini...")
    
    # Check if report already exists locally to speed up auth testing
    filename = f"Threads_Report_{datetime.date.today()}.md"
    if os.path.exists(filename):
        print(f"Found existing local report: {filename}. Using it.")
        with open(filename, 'r', encoding='utf-8') as f:
            # Skip header line
            lines = f.readlines()
            report_content = "".join(lines)
    else:
        # Dynamic model selection
        try:
            all_models = list(genai.list_models())
            # Filter for models that likely support text generation
            candidates = [m for m in all_models if 'generateContent' in m.supported_generation_methods]
            
            # Sort to prefer 1.5 flash or pro
            candidates.sort(key=lambda m: ('flash' in m.name, '1.5' in m.name), reverse=True)
            if not candidates:
                 candidates = all_models

            # Try models
            model_names = [m.name for m in candidates] + ['gemini-1.5-flash', 'gemini-pro']
            unique_names = []
            [unique_names.append(x) for x in model_names if x not in unique_names]

            for model_name in unique_names:
                try:
                    print(f"Trying model: {model_name}...")
                    model = genai.GenerativeModel(model_name)
                    response = model.generate_content(prompt)
                    report_content = response.text
                    print(f"Successfully generated content with {model_name}")
                    break
                except Exception as e:
                    print(f"Failed with {model_name}: {e}")
        except:
             report_content = "AI 분석 실패. 로그 확인 바람."

    if not report_content:
        report_content = "분석된 내용이 없습니다."

    print("Report generated.")
    print("="*50)
    print(report_content)
    print("="*50)
    
    subject = f"Threads Daily Report: @choi.openai ({datetime.date.today()})"
    
    email_success = False
    if EMAIL_RECIPIENT:
        print(f"Attempting to send email to {EMAIL_RECIPIENT}...")
        try:
             email_success = send_email(subject, report_content, EMAIL_RECIPIENT)
        except Exception as e:
             print(f"Email send failed: {e}")
    
    if not email_success:
        print("Email could not be sent. Saving report to file instead.")
        save_report_locally(subject, report_content)
    
    # Send KakaoTalk
    print("Attempting to send KakaoTalk...")
    
    # Send FULL content (safe sliced to avoid overflow error, approx 1900 chars)
    # Adding title header
    full_text = f"[{datetime.date.today()}] Threads Daily Report\n@choi.openai\n\n{report_content}"
    
    send_kakao(full_text)

if __name__ == "__main__":
    analyze_posts()
