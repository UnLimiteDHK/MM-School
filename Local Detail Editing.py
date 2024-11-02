import logging
from googleapiclient.discovery import build
from google.oauth2 import service_account

# ログの設定
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# 定数の設定
SERVICE_ACCOUNT_FILE = r'C:\Users\kanchi\Desktop\プログラミング\MMスクール\出品算出シート001\mmschool-unlimi-001-dc6603fc2808.json'
SPREADSHEET_ID = '1oNSqWAQZd-Tqg5QUsY-M-hjx1Pf9WdFgGxPIipW0sEE'

class GoogleSheetService:
    def __init__(self, service_account_file, spreadsheet_id):
        self.scopes = ['https://www.googleapis.com/auth/spreadsheets']
        self.credentials = service_account.Credentials.from_service_account_file(
            service_account_file, scopes=self.scopes)
        self.service = build('sheets', 'v4', credentials=self.credentials)
        self.spreadsheet_id = spreadsheet_id
        logging.debug(f"Initialized GoogleSheetService with spreadsheet ID: {spreadsheet_id}")

    def get_values(self, range_name):
        try:
            logging.debug(f"Fetching values from range: {range_name}")
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id, range=range_name).execute()
            values = result.get('values', [])
            logging.debug(f"Fetched values: {values}")
            return values
        except Exception as e:
            logging.error(f"Error fetching values from range {range_name}: {e}")
            return []

    def update_values(self, range_name, values):
        try:
            body = {
                'values': values
            }
            result = self.service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id, range=range_name,
                valueInputOption='RAW', body=body).execute()
            logging.debug(f"Updated range {range_name} with values: {values}")
            return result
        except Exception as e:
            logging.error(f"Error updating values in range {range_name}: {e}")
            return None

    def batch_update_values(self, data):
        try:
            body = {
                'valueInputOption': 'RAW',
                'data': data
            }
            result = self.service.spreadsheets().values().batchUpdate(
                spreadsheetId=self.spreadsheet_id, body=body).execute()
            logging.debug(f"Batch update result: {result}")
            return result
        except Exception as e:
            logging.error(f"Error during batch update: {e}")
            return None

# 使用例
service_account_file = SERVICE_ACCOUNT_FILE
spreadsheet_id = SPREADSHEET_ID

gs_service = GoogleSheetService(service_account_file, spreadsheet_id)

# データの取得
range_name = 'シート1!A1:B2'
values = gs_service.get_values(range_name)
print(values)

# データの更新
update_range_name = 'シート1!A1:B2'
update_values = [
    ['新しい値1', '新しい値2'],
    ['新しい値3', '新しい値4']
]
gs_service.update_values(update_range_name, update_values)

# バッチ更新
batch_data = [
    {
        'range': 'シート1!A1:B2',
        'values': [
            ['バッチ値1', 'バッチ値2'],
            ['バッチ値3', 'バッチ値4']
        ]
    }
]
gs_service.batch_update_values(batch_data)