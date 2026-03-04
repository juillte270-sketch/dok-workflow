from flask import Flask, request, redirect
import requests
import json
import os
import webbrowser
from dotenv import load_dotenv

# Load Env
load_dotenv()
KAKAO_REST_API_KEY = os.getenv('KAKAO_REST_API_KEY')
REDIRECT_URI = "http://127.0.0.1:5000/"

app = Flask(__name__)

@app.route('/')
def home():
    code = request.args.get('code')
    
    if code:
        # If code is present in query string, it's the callback
        token_url = "https://kauth.kakao.com/oauth/token"
        payload = {
            "grant_type": "authorization_code",
            "client_id": KAKAO_REST_API_KEY,
            "redirect_uri": REDIRECT_URI,
            "code": code
        }
        
        try:
            response = requests.post(token_url, data=payload)
            tokens = response.json()
            
            if 'access_token' in tokens:
                with open('kakao_token.json', 'w') as f:
                    json.dump(tokens, f)
                return "<h1>Login Successful!</h1><p>You can close this window now. Token saved to kakao_token.json</p>"
            else:
                return f"<h1>Error</h1><p>{tokens}</p>"
        except Exception as e:
            return f"<h1>Error</h1><p>{e}</p>"
    else:
        # No code, redirect to login
        kakao_login_url = (
            f"https://kauth.kakao.com/oauth/authorize?"
            f"client_id={KAKAO_REST_API_KEY}&"
            f"redirect_uri={REDIRECT_URI}&"
            f"response_type=code&"
            f"scope=talk_message,friends"
        )
        return redirect(kakao_login_url)

if __name__ == '__main__':
    print(f"Opening browser for Kakao Login... {REDIRECT_URI}")
    webbrowser.open(REDIRECT_URI)
    app.run(port=5000)
