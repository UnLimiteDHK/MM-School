import openai
import requests
from googleapiclient.discovery import build
from google.oauth2 import service_account
import threading
import aiohttp
import logging

# ログの設定
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

class GoogleSheetService:
    def __init__(self, service_account_file, spreadsheet_id):
        self.scopes = ['https://www.googleapis.com/auth/spreadsheets']
        self.credentials = service_account.Credentials.from_service_account_file(
            service_account_file, scopes=self.scopes)
        self.service = build('sheets', 'v4', credentials=self.credentials)
        self.spreadsheet_id = spreadsheet_id
        logging.debug(f"Initialized GoogleSheetService with spreadsheet ID: {spreadsheet_id}")

    def get_sheet_id(self, sheet_name):
        try:
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
            for sheet in spreadsheet.get('sheets', []):
                if sheet.get("properties", {}).get("title") == sheet_name:
                    return sheet.get("properties", {}).get("sheetId")
            logging.error(f"Sheet name {sheet_name} not found.")
            return None
        except Exception as e:
            logging.error(f"Error fetching sheet ID for {sheet_name}: {e}")
            return None

    def get_values(self, range_name):
        try:
            logging.debug(f"Fetching values from range: {range_name}")
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id, range=range_name).execute()
            logging.debug(f"Fetched values: {result.get('values', [])}")
            return result.get('values', [])
        except Exception as e:
            logging.error(f"Error fetching values from range {range_name}: {e}")
            return []

    def update_values(self, range_name, values):
        try:
            logging.debug(f"Updating values in range: {range_name} with data: {values}")
            body = {
                'values': values
            }
            result = self.service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id, range=range_name,
                valueInputOption='RAW', body=body).execute()
            logging.info("Update successful")
            return result
        except Exception as e:
            logging.error(f"Error updating values in range {range_name}: {e}")

def prepare_batch_data(sheet_name, start_col, values):
    data = []
    for i, value in enumerate(values):
        if isinstance(value, list):
            for j, val in enumerate(value):
                data.append({
                    'range': f'{sheet_name}!{chr(start_col + j)}{i + 1}',
                    'values': [[val]]
                })
        else:
            data.append({
                'range': f'{sheet_name}!{chr(start_col)}{i + 1}',
                'values': [[value]]
            })
    #logging.debug(f"Prepared batch data: {data}")
    return data

def main():
    SERVICE_ACCOUNT_FILE = r'C:\Users\kanchi\Desktop\プログラミング\MMスクール\出品算出シート001\mmschool-unlimi-001-dc6603fc2808.json'
    SPREADSHEET_ID = '1oNSqWAQZd-Tqg5QUsY-M-hjx1Pf9WdFgGxPIipW0sEE'

    sheet_service = GoogleSheetService(SERVICE_ACCOUNT_FILE, SPREADSHEET_ID)
    
    # スプレッドシートから値を取得
    ss_range_listing_csv = ['出品用CSV!AD2:AD', '出品用CSV!AE2:AE', '出品用CSV!B2:B', '出品用CSV!H2:H']
    values_list = list(map(sheet_service.get_values, ss_range_listing_csv))
    #logging.debug(f"Fetched values: {values_list}")
    # AI-memoシートの全てのデータとセルの色をクリア
    try:
        # シート全体の範囲を指定
        clear_range = 'AI-memo'
        clear_body = {
            'ranges': [clear_range]
        }
        # データをクリア
        sheet_service.service.spreadsheets().values().batchClear(
            spreadsheetId=sheet_service.spreadsheet_id,
            body=clear_body
        ).execute()
        
        # セルの色をクリアするためのリクエストを作成
        # AI-memoシートのIDを取得
        sheet_metadata = sheet_service.service.spreadsheets().get(
            spreadsheetId=sheet_service.spreadsheet_id
        ).execute()
        sheets = sheet_metadata.get('sheets', '')
        ai_memo_sheet_id = None
        for sheet in sheets:
            if sheet.get("properties", {}).get("title", "") == "AI-memo":
                ai_memo_sheet_id = sheet.get("properties", {}).get("sheetId", None)
                break

        if ai_memo_sheet_id is not None:
            requests = [{
                'repeatCell': {
                    'range': {
                        'sheetId': ai_memo_sheet_id
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'backgroundColor': {
                                'red': 1,
                                'green': 1,
                                'blue': 1
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.backgroundColor'
                }
            }]
            
            # バッチ更新リクエストを送信
            sheet_service.service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_service.spreadsheet_id,
                body={'requests': requests}
            ).execute()
        else:
            logging.error("AI-memoシートが見つかりませんでした。")
        
        logging.info("Cleared all data and cell colors from AI-memo sheet")
    except Exception as e:
        logging.error(f"Error clearing data and cell colors from AI-memo sheet: {e}")
    # AI-memoシートの1行目に指定されたヘッダーを追加
    headers = [
        "日本語タイトル", "日本語説明", "SKU", "画像-01", "画像-02", "画像-03", "画像-04", "画像-05",
        "画像-06", "画像-07", "画像-08", "画像-09", "画像-10", "画像-11", "画像-12", "画像-13",
        "画像-14", "画像-15", "画像-16", "画像-17", "画像-18", "画像-19", "画像-20", "画像-21",
        "画像-22", "画像-23", "画像-24", "New-Titel", "New-Discription"
    ]
    header_range = 'AI-memo!A1:1'  # 1行目の範囲を指定
    sheet_service.update_values(header_range, [headers])

    # 出品用CSVシートの1行目のAF列以降の値を取得
    csv_header_range = '出品用CSV!AF1:1'  # AF列以降の1行目の範囲を指定
    csv_headers = sheet_service.get_values(csv_header_range)

    # AI-memoシートの1行目の指定されたヘッダーの後に出品用CSVシートのヘッダーを追加
    if csv_headers and len(csv_headers) > 0:
        combined_headers = headers + csv_headers[0]
        combined_header_range = 'AI-memo!A1:1'  # 1行目の範囲を指定
        sheet_service.update_values(combined_header_range, [combined_headers])
    # AI-memoシートのA列に最初のデータを2行目から貼り付け
    if values_list and values_list[0]:
        range_to_update = 'AI-memo!A2:A'  # A列の2行目からの範囲を指定
        sheet_service.update_values(range_to_update, values_list[0])
    
    # B列に2番目のデータを2行目から貼り付け
    if len(values_list) > 1 and values_list[1]:
        range_to_update = 'AI-memo!B2:B'  # B列の2行目からの範囲を指定
        sheet_service.update_values(range_to_update, values_list[1])
    
    # C列に3番目のデータを2行目から貼り付け
    if len(values_list) > 2 and values_list[2]:
        range_to_update = 'AI-memo!C2:C'  # C列の2行目からの範囲を指定
        sheet_service.update_values(range_to_update, values_list[2])
    
    # D列以降に4番目のデータを2行目から貼り付け
    if len(values_list) > 3 and values_list[3]:
        # 各要素がリストであることを考慮
        split_values = [value[0].split('|') for value in values_list[3] if value and isinstance(value, list) and len(value) > 0]
        
        # バッチデータを準備
        batch_data = []
        format_requests = []
        for row_index, row_values in enumerate(split_values, start=2):  # 2行目から開始
            for col_index, val in enumerate(row_values, start=3):  # D列は3番目のインデックス
                range_to_update = f'AI-memo!{chr(65 + col_index)}{row_index}:{chr(65 + col_index)}{row_index}'
                batch_data.append({
                    'range': range_to_update,
                    'values': [[val]]
                })
                # 値があるセルに色をつけるリクエストを追加
                format_requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_service.spreadsheet_id,
                            "startRowIndex": row_index - 1,
                            "endRowIndex": row_index,
                            "startColumnIndex": col_index,
                            "endColumnIndex": col_index + 1
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": 1.0,
                                    "green": 1.0,
                                    "blue": 0.0
                                }
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })
        
        # バッチで更新
        if batch_data:
            # Google Sheets APIのバッチ更新を実行
            body = {
                'valueInputOption': 'RAW',
                'data': batch_data
            }
            sheet_service.service.spreadsheets().values().batchUpdate(
                spreadsheetId=sheet_service.spreadsheet_id,
                body=body
            ).execute()
        
        # セルの色を変更するリクエストを実行
        if format_requests:
            # シートIDを整数として取得
            sheet_id = sheet_service.get_sheet_id('AI-memo')
            if sheet_id is not None:
                # シートIDをリクエストに追加
                for request in format_requests:
                    request['repeatCell']['range']['sheetId'] = sheet_id

                # バッチ更新を実行
                sheet_service.service.spreadsheets().batchUpdate(
                    spreadsheetId=sheet_service.spreadsheet_id,
                    body={"requests": format_requests}
                ).execute()
            else:
                logging.error("Failed to retrieve sheet ID for formatting requests.")
    
    # シートIDを取得
    sheet_id = sheet_service.get_sheet_id('AI-memo')
    if sheet_id is None:
        logging.error("Failed to retrieve sheet ID.")
        exit()

    # D列2行目からAA列までのデータを取得
    range_to_fetch = 'AI-memo!D2:AA'
    existing_values = sheet_service.get_values(range_to_fetch)
    logging.debug(f"Existing values: {existing_values}")

    # SettingシートのB2列の画像URLを取得
    setting_values = sheet_service.get_values('Setting!B2')
    
    if setting_values and len(setting_values) > 0 and len(setting_values[0]) > 0:
        # 取得した値をGoogle DriveのURLから画像IDに変換
        drive_url = setting_values[0][0]  # 取得した値の最初の要素を使用
        try:
            image_id = drive_url.split('/d/')[1].split('/')[0]  # URLから画像IDを抽出
            image_url = f"https://lh3.googleusercontent.com/d/{image_id}"
            logging.debug(f"Image URL: {image_url}")
            
            # 各行の要素数が24になるまで画像URLを追加
            updated_values = False
            for row_index, row in enumerate(existing_values):
                while len(row) < 24:
                    row.append(image_url)
                    updated_values = True
            
            # 更新するデータを準備
            if updated_values:
                logging.debug(f"Updated values: {existing_values}")
                try:
                    sheet_service.update_values(range_to_fetch, existing_values)
                except Exception as e:
                    logging.error(f"Error updating sheet with image URLs: {e}")
        except IndexError:
            logging.error("Invalid Google Drive URL format.")
    else:
        logging.error("No valid image URL found in Setting!B2.")

if __name__ == "__main__":
    main()
