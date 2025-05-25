from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import os

def upload_to_drive(file_path, folder_id=None, file_name=None):
    try:
        cred_path =r'G:\MES自動化\credentials.json'
        SCOPES = ['https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_file(cred_path, scopes=SCOPES)
        service = build('drive', 'v3', credentials=creds)
        if file_name is None:
            file_name = os.path.basename(file_path)
        if not os.path.exists(file_path):
            print(f"找不到上傳的檔案：{file_path}")
            return
        file_metadata = {'name': file_name}
        if folder_id:
            file_metadata['parents'] = [folder_id]
        media = MediaFileUpload(file_path, mimetype='application/pdf')
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        file_id = file.get('id')
        print(f"檔案上傳成功，ID: {file_id}")

        permission = {
            'type': 'anyone',
            'role': 'reader'
        }
        service.permissions().create(fileId=file_id, body=permission).execute()
        print('任何人都可瀏覽')
        file = service.files().get(fileId=file_id, fields='webViewLink').execute()
        share_url = file.get('webViewLink')
        print(f'連結：{share_url}')
        return share_url
    except Exception as e:
        print(f'上傳失敗：{e}')

