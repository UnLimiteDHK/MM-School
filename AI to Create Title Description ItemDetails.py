import logging
import requests
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseUpload
import io
import string
from concurrent.futures import ThreadPoolExecutor, as_completed
import openai
import base64
import time
import json
from PIL import Image
from datetime import datetime
from googleapiclient.errors import HttpError
from requests.adapters import HTTPAdapter

# ログの設定
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# 定数の設定
SERVICE_ACCOUNT_FILE = r'C:\Users\kanchi\Desktop\プログラミング\MMスクール\出品算出シート001\mmschool-unlimi-001-dc6603fc2808.json'
SPREADSHEET_ID = '1oNSqWAQZd-Tqg5QUsY-M-hjx1Pf9WdFgGxPIipW0sEE'
MAX_RETRIES = 10
BATCH_SIZE = 20

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
        except Exception as e:
            logging.error(f"Error during batch update: {e}")

class OpenAIService:
    def __init__(self, api_keys):
        self.api_keys = api_keys

    def generate_summary(self, description, index):
        endpoint = "https://api.openai.com/v1/chat/completions"
        messages = [
            {
                "role": "user",
                "content": f"次のテキストを要約/修正してください: {description}\n"
                           f"**商品の説明とは関係のない文言は省いてください。"
            }
        ]

        payload = {
            "model": "gpt-4o-mini",
            "messages": messages,
            "max_tokens": 500,
            "temperature": 0.5
        }

        headers = {
            'Authorization': f'Bearer {self.api_keys[index % len(self.api_keys)]}',
            'Content-Type': 'application/json'
        }

        try:
            response = requests.post(endpoint, headers=headers, json=payload)
            response.raise_for_status()
            json_response = response.json()
            summary = json_response['choices'][0]['message']['content']
            return summary
        except requests.exceptions.HTTPError as e:
            logging.error(f"HTTP error: {e}")
            return "Error generating summary"
        except Exception as e:
            logging.error(f"Error generating summary: {e}")
            return "Error generating summary"

    def send_to_openai(self, image_url, title, description, item_specifics_headers, index):
        endpoint = "https://api.openai.com/v1/chat/completions"

        item_specifics_schema = {header: {"type": "string"} for header in item_specifics_headers}

        schema = {
            "type": "object",
            "properties": {
                "NewTitle": {"type": "string"},
                "NewDescription": {"type": "string"},
                "ItemSpecifics": {
                    "type": "object",
                    "properties": item_specifics_schema
                }
            }
        }

        messages = [
            {
                "role": "user",
                "content": f"画像を見て、次の参考情報を元に商品タイトルと説明を英語で作成し、次の参考情報に基づいて商品情報を埋めてください。\n"
                           f"参考商品タイトル-{title}\n"
                           f"参考商品説明-{description}\n"
                           f"商品情報: {', '.join(item_specifics_headers)}\n"
                           f"出力形式はJSON形式で NewTitle, NewDescription, ItemSpecifics を含むようにしてください。\n"
                           f"**作成するタイトルは英語で可能な限り80文字に近い文字数で作成してください。半角スペースは1文字でカウントします。できるだけ80字に近くなるように作成してください。**\n"
                           f"**※作成するタイトルは最大80文字とし、できるだけ文字数を活用してください。**\n"
                           f"**作成するタイトルをカウントして70文字以下のようなに明らかに少なければ再度作成してください。**\n"
                           f"**作成する商品説明は送料関する項目や発送方法などの余分な説明は省いてください。**\n"
                           f"**省く文言例：発送に関しての説明、梱包に関しての説明、購入する際の注意点など**\n"
                           f"**残す文言例：商品のサイズ、商品に関する傷や汚れなどの注意事項**\n"
                           f"**商品情報に記載するサイズがセンチメートルの場合はインチに変換して記載してください。**\n"
                           f"**情報が不明な場合は「N/A」と記載してください。**\n"
            }
        ]
        if image_url:
            base64_image = ImageService.encode_image_from_url(image_url)
            if base64_image:
                messages.append({
                    "role": "user",
                    "content": f"data:image/jpeg;base64,{base64_image}"
                })

        payload = {
            "model": "gpt-4o-mini",
            "messages": messages,
            "max_tokens": 4000,
            "temperature": 0.0,
            "functions": [
                {
                    "name": "generate_product_info",
                    "parameters": schema
                }
            ]
        }

        headers = {
            'Authorization': f'Bearer {self.api_keys[index % len(self.api_keys)]}',
            'Content-Type': 'application/json'
        }

        retry_attempts = 5
        for attempt in range(retry_attempts):
            try:
                logging.info(f'Using API Key: {self.api_keys[index % len(self.api_keys)]}')
                response = requests.post(endpoint, headers=headers, json=payload)
                response.raise_for_status()

                json_response = response.json()
                extracted_data = json_response['choices'][0]['message']['function_call']['arguments']
                logging.info(f'API response received for title: {title}')

                try:
                    extracted_json = json.loads(extracted_data)
                    new_title = extracted_json.get("NewTitle", "No Title Found")
                    new_description = extracted_json.get("NewDescription", "No Description Found")
                    item_specifics = extracted_json.get("ItemSpecifics", {})
                    return new_title, new_description, item_specifics
                except json.JSONDecodeError as e:
                    logging.error(f"JSON decode error: {e}")
                    return "JSON Decode Error"

            except requests.exceptions.HTTPError as e:
                if response.status_code == 429:
                    wait_time = 10 * (2 ** attempt)
                    logging.warning(f"Rate limit reached, retrying in {wait_time} seconds...")
                    time.sleep(wait_time)
                else:
                    logging.error(f'Error during API request: {e}')
                    logging.error(f'Response content: {response.content.decode()}')
                    return "Request Error"
            except requests.RequestException as e:
                logging.error(f'Error during API request: {e}')
                return "Request Error"

        logging.error("Max retry attempts reached.")
        return "Max Retry Error"

class ImageService:
    @staticmethod
    def encode_image_from_url(url, max_size=(150, 150)):
        try:
            response = requests.get(url)
            response.raise_for_status()

            image = Image.open(io.BytesIO(response.content))
            image.thumbnail(max_size)

            buffered = io.BytesIO()
            image.save(buffered, format="JPEG", quality=85)

            return base64.b64encode(buffered.getvalue()).decode('utf-8')
        except requests.RequestException as e:
            logging.error(f'Error encoding image from URL: {e}')
            return None

class BatchUpdater:
    @staticmethod
    def batch_update_values(sheet_service, data, max_retries=MAX_RETRIES, batch_size=BATCH_SIZE):
        for i in range(0, len(data), batch_size):
            batch_data = data[i:i + batch_size]
            for attempt in range(max_retries):
                try:
                    body = {
                        'valueInputOption': 'RAW',
                        'data': batch_data
                    }
                    result = sheet_service.service.spreadsheets().values().batchUpdate(
                        spreadsheetId=sheet_service.spreadsheet_id, body=body).execute()
                    logging.debug(f"Batch update result: {result}")
                    break
                except HttpError as e:
                    if e.resp.status in [403, 429]:
                        wait_time = 60 ** attempt
                        logging.warning(f"Rate limit reached, retrying in {wait_time} seconds...")
                        time.sleep(wait_time)
                    else:
                        logging.error(f"Error during batch update: {e}")
                        break
                except Exception as e:
                    logging.error(f"Unexpected error during batch update: {e}")
                    break

class Utils:
    @staticmethod
    def get_column_letter(index):
        letters = ""
        while index >= 0:
            letters = string.ascii_uppercase[index % 26] + letters
            index = index // 26 - 1
        return letters

def get_openai_api_keys(sheet_service):
    try:
        api_key_values = sheet_service.get_values('Setting!F1:F')
        api_keys = [row[0] for row in api_key_values if row]
        if not api_keys:
            raise ValueError("No API keys found in the specified range.")
        logging.debug(f"Fetched OpenAI API keys: {api_keys}")
        return api_keys
    except Exception as e:
        logging.error(f"Error fetching OpenAI API keys: {e}")
        return []

def main():
    start_time = datetime.now()
    
    sheet_service = GoogleSheetService(SERVICE_ACCOUNT_FILE, SPREADSHEET_ID)
    openai_api_keys = get_openai_api_keys(sheet_service)
    openai_service = OpenAIService(openai_api_keys)

    if not openai_api_keys:
        logging.error("OpenAI APIキーが見つかりませんでした。")
        return

    descriptions = sheet_service.get_values('AI-memo!B2:B')
    row_indices = [i + 2 for i in range(len(descriptions))]

    # 説明を要約
    #summaries = []
    #for i, description in enumerate(descriptions):
    #    summary = openai_service.generate_summary(description[0], i)
    #    summaries.append(summary)

    #data = [{'range': f'AI-memo!B{row_indices[i]}:B{row_indices[i]}', 'values': [[summary]]} for i, summary in enumerate(summaries) if summary]
    #BatchUpdater.batch_update_values(sheet_service, data)

    # 商品タイトルと説明を更新
    item_specifics_headers = sheet_service.get_values('AI-memo!AD1:1')[0]
    jp_titles = sheet_service.get_values('AI-memo!A2:A')
    img_urls = sheet_service.get_values('AI-memo!D2:D')
    min_length = min(len(jp_titles), len(descriptions), len(img_urls))
    
    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = {
            executor.submit(
                openai_service.send_to_openai,
                img_urls[i][0],
                jp_titles[i][0],
                descriptions[i][0],
                item_specifics_headers,
                i
            ): i for i in range(min_length)
        }

        titles = [None] * min_length
        new_descriptions = [None] * min_length
        specifics = {header: [None] * min_length for header in item_specifics_headers}

        for future in as_completed(futures):
            i = futures[future]
            try:
                new_title, new_description, item_specifics = future.result()
                titles[i] = new_title
                new_descriptions[i] = new_description
                for key, value in item_specifics.items():
                    if key in specifics:
                        specifics[key][i] = value
            except Exception as e:
                logging.error(f"Error in thread: {e}")

    # タイトルと説明をシートに挿入
    data = []
    for i, (title, description) in enumerate(zip(titles, new_descriptions)):
        if title is not None:
            title_range = f'AI-memo!AB{row_indices[i]}:AB{row_indices[i]}'
            data.append({'range': title_range, 'values': [[title]]})

        if description is not None:
            description_range = f'AI-memo!AC{row_indices[i]}:AC{row_indices[i]}'
            data.append({'range': description_range, 'values': [[description]]})

    BatchUpdater.batch_update_values(sheet_service, data)

    # 商品情報 (ItemSpecifics) をシートに挿入
    item_data = []
    for key, values in specifics.items():
        if key in item_specifics_headers:
            col_index = item_specifics_headers.index(key) + 29  # AD列は29番目
            col_letter = Utils.get_column_letter(col_index)  # 列の文字を取得
            for i, value in enumerate(values):
                if value is not None:
                    specifics_range = f'AI-memo!{col_letter}{row_indices[i]}:{col_letter}{row_indices[i]}'
                    item_data.append({'range': specifics_range, 'values': [[value]]})

    BatchUpdater.batch_update_values(sheet_service, item_data)
    
    end_time = datetime.now()
    elapsed_time = end_time - start_time
    logging.info(f"プログラム開始時間: {start_time}")
    logging.info(f"プログラム終了時間: {end_time}")
    logging.info(f"プログラムにかかった時間: {elapsed_time}")

if __name__ == "__main__":
    main()
