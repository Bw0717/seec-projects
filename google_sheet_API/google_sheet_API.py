from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = r'G:\MES自動化\credentials.json'
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=creds)
SPREADSHEET_ID = '1SbBtTxpFqQ_HoAiw2yJM4OhHR7dUMRPbL_RJFFV4UeQ'
RANGE_NAME = '工作表1!A1:B2'
result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
values = result.get('values', [])
print("變更前資料：")
for row in values:
    print(row)
new_values = [
    ["產品", "價格"],
    ["筆電", "500"]
]

body = {'values': new_values}
response = service.spreadsheets().values().update(
    spreadsheetId=SPREADSHEET_ID,
    range=RANGE_NAME,
    valueInputOption="RAW",
    body=body
).execute()
print(f"更新範圍: {response.get('updatedRange')}")
print("資料已更新：")
update_result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
update_values = update_result.get('values', [])
for row in update_values:
    print(row)