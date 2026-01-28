from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import os.path

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1Bv1iLDh_dpD5DllYUaW8KnejKBx9vajm'

def main():
    if not os.path.exists('token.json'):
        print("token.json not found. Authenticate first.")
        return
        
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    service = build('sheets', 'v4', credentials=creds)

    # Get sheet names
    spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = spreadsheet.get('sheets', [])
    
    for sheet in sheets:
        title = sheet.get('properties', {}).get('title')
        print(f"\nSheet: {title}")
        
        # Read the first 5 rows to see headers
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{title}'!A1:Z5"
        ).execute()
        values = result.get('values', [])
        
        for i, row in enumerate(values):
            print(f"Row {i+1}: {row}")

if __name__ == '__main__':
    main()
