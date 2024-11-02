from fastapi import FastAPI, HTTPException
import openai
import requests
from googleapiclient.discovery import build
from google.oauth2 import service_account
import logging
import asyncio
from concurrent.futures import ThreadPoolExecutor
import uvicorn
# ログの設定
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# FastAPIのインスタンスを作成
app = FastAPI()

executor = ThreadPoolExecutor()

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
            raise HTTPException(status_code=500, detail=str(e))

    def batch_clear_values(self, ranges):
        try:
            logging.debug(f"Clearing values in ranges: {ranges}")
            clear_body = {
                'ranges': ranges
            }
            result = self.service.spreadsheets().values().batchClear(
                spreadsheetId=self.spreadsheet_id,
                body=clear_body
            ).execute()
            logging.info("Batch clear successful")
            return result
        except Exception as e:
            logging.error(f"Error clearing values in ranges {ranges}: {e}")
            raise HTTPException(status_code=500, detail=str(e))

    def batch_update_cell_colors(self, sheet_id, requests):
        try:
            logging.debug(f"Updating cell colors with requests: {requests}")
            result = self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={'requests': requests}
            ).execute()
            logging.info("Batch update cell colors successful")
            return result
        except Exception as e:
            logging.error(f"Error updating cell colors: {e}")
            raise HTTPException(status_code=500, detail=str(e))

# Google Sheetsサービスのインスタンスを初期化
SERVICE_ACCOUNT_FILE = r'C:\Users\kanchi\Desktop\プログラミング\MMスクール\出品算出シート001\mmschool-unlimi-001-dc6603fc2808.json'
SPREADSHEET_ID = '1oNSqWAQZd-Tqg5QUsY-M-hjx1Pf9WdFgGxPIipW0sEE'
sheet_service = GoogleSheetService(SERVICE_ACCOUNT_FILE, SPREADSHEET_ID)

@app.get("/get-values/{range_name}")
async def get_values(range_name: str):
    """
    指定された範囲からGoogle Sheetsの値を取得します。
    """
    try:
        loop = asyncio.get_event_loop()
        values = await loop.run_in_executor(executor, sheet_service.get_values, range_name)
        if not values:
            raise HTTPException(status_code=404, detail="No values found in the specified range.")
        return {"values": values}
    except Exception as e:
        logging.error(f"Error in get_values endpoint: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/update-values/{range_name}")
async def update_values(range_name: str, values: list):
    """
    指定された範囲にGoogle Sheetsの値を更新します。
    """
    try:
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(executor, sheet_service.update_values, range_name, values)
        return {"updatedRange": result.get('updatedRange'), "updatedRows": result.get('updatedRows')}
    except Exception as e:
        logging.error(f"Error in update_values endpoint: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/batch-clear-values")
async def batch_clear_values(ranges: list):
    """
    指定された範囲のGoogle Sheetsの値をクリアします。
    """
    try:
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(executor, sheet_service.batch_clear_values, ranges)
        return {"clearedRanges": result.get('clearedRanges')}
    except Exception as e:
        logging.error(f"Error in batch_clear_values endpoint: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/update-cell-colors/{sheet_name}")
async def update_cell_colors(sheet_name: str, color_requests: list):
    """
    指定されたシート内のセルの色を更新します。
    """
    try:
        loop = asyncio.get_event_loop()
        sheet_id = await loop.run_in_executor(executor, sheet_service.get_sheet_id, sheet_name)
        if sheet_id is None:
            raise HTTPException(status_code=404, detail="Sheet not found.")
        # シートIDをリクエストに追加
        for request in color_requests:
            request['repeatCell']['range']['sheetId'] = sheet_id
        result = await loop.run_in_executor(executor, sheet_service.batch_update_cell_colors, sheet_id, color_requests)
        return {"updatedCells": result.get('replies')}
    except Exception as e:
        logging.error(f"Error in update_cell_colors endpoint: {e}")
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
