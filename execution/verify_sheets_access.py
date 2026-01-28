from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os.path
import pickle

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

SPREADSHEET_ID = '1Bv1iLDh_dpD5DllYUaW8KnejKBx9vajm' # From user provided link

def main():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Token expired, refreshing...")
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            # Authenticate using the new credentials
            creds = flow.run_local_server(port=0)
            
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        drive_service = build('drive', 'v3', credentials=creds)
        
        # Get file metadata
        file_metadata = drive_service.files().get(fileId=SPREADSHEET_ID, fields="name, mimeType").execute()
        
        print(f"SUCCESS: Found file on Drive!")
        print(f"Name: {file_metadata.get('name')}")
        print(f"MimeType: {file_metadata.get('mimeType')}")
        
        if file_metadata.get('mimeType') == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            print("Note: This is an Excel (.xlsx) file, not a native Google Sheet.")
            print("We will handle this by downloading/uploading via Drive API.")
        
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        print(f"ERROR: Could not access spreadsheet.")
        print("-" * 40)
        print(error_msg)
        print("-" * 40)
        with open('last_error.txt', 'w', encoding='utf-8') as f:
            f.write(error_msg)

if __name__ == '__main__':
    main()
