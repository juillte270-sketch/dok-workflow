from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import os.path
import json

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def main():
    if not os.path.exists('token.json'):
        print("token.json not found.")
        return
        
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    service = build('sheets', 'v4', credentials=creds)

    # Load ID from .env
    spreadsheet_id = None
    if os.path.exists('.env'):
        with open('.env', 'r', encoding='utf-8') as f:
            for line in f:
                if line.startswith('MASTER_DATA_FILE_ID='):
                    spreadsheet_id = line.split('=')[1].strip()
                    break
    
    if not spreadsheet_id:
        print("SPREADSHEET_ID not found in .env")
        return

    try:
        print(f"Accessing Spreadsheet: {spreadsheet_id}")
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        
        for sheet in sheets:
            title = sheet.get('properties', {}).get('title')
            print(f"\nSheet Found: {title}")
            
            # Read first row
            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{title}'!A1:Z2"
            ).execute()
            values = result.get('values', [])
            if values:
                print(f"Header: {values[0]}")
            else:
                print("Sheet is empty.")
                
    except Exception as e:
        print(f"ERROR: {e}")

if __name__ == '__main__':
    main()
