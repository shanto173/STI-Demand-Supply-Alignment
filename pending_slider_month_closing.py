import os
import sys
import logging
import time
import io
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import pytz
import pandas as pd
import json
import requests
from google.oauth2 import service_account
import gspread
from gspread_dataframe import set_with_dataframe
from dotenv import load_dotenv

# ===== Setup Logging =====
logging.basicConfig(stream=sys.stdout, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
log = logging.getLogger()

# ===== Load environment variables =====
load_dotenv()
ODOO_URL = os.getenv("ODOO_URL")
DB = os.getenv("ODOO_DB")
USERNAME = os.getenv("ODOO_USERNAME", "supply.chain3@texzipperbd.com")
PASSWORD = os.getenv("ODOO_PASSWORD", "@Shanto@86")

# Google Sheets info - Replace with your actual IDs
SPREADSHEET_ID = "1PECBN0MiOpOeVa3XvSpit5Xb5fLT91HxbXNvEUpbqM8" # Update if needed
WORKSHEET_NAME = "Pending_Slider" # Update if needed
GCP_CREDENTIALS_FILE = "rugged-feat-456510-q7-381d4ee0fb45.json"

session = requests.Session()
USER_ID = None

def check_first_day():
    dhaka_tz = pytz.timezone('Asia/Dhaka')
    now = datetime.now(dhaka_tz)
    
    # We strictly run only on the 1st of the month.
    if now.day != 1:
        log.warning(f"Skipping execution. Expected run on 1st of the month BD time. Current BD time: {now}")
        sys.exit(0)
        
    return now

def get_previous_month_dates(now):
    """Returns the start and end dates of the previous month."""
    # Since this runs on the 1st, we want data for the month that just ended.
    first_day_of_current_month = now.replace(day=1)
    last_day_of_prev_month = first_day_of_current_month - relativedelta(days=1)
    first_day_of_prev_month = last_day_of_prev_month.replace(day=1)
    
    return first_day_of_prev_month.strftime("%Y-%m-%d"), last_day_of_prev_month.strftime("%Y-%m-%d")

def login():
    global USER_ID
    log.info("Attempting to login to Odoo...")
    payload = {
        "jsonrpc": "2.0",
        "params": {
            "db": DB,
            "login": USERNAME,
            "password": PASSWORD
        }
    }
    r = session.post(f"{ODOO_URL}/web/session/authenticate", json=payload)
    r.raise_for_status()
    result = r.json().get("result", {})
    if "uid" in result:
        USER_ID = result["uid"]
        log.info(f"✅ Logged in successfully (uid={USER_ID})")
        return result
    else:
        raise Exception("❌ Login failed! Check credentials or URL.")

def generate_report(date_from, date_to):
    """Creates the ppc.report record and returns the report ID."""
    log.info(f"Creating ppc.report from {date_from} to {date_to}...")
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "ppc.report",
            "method": "web_save",
            "args": [
                [],
                {
                    "report_type": "pslc",
                    "date_from": date_from,
                    "date_to": date_to,
                    "order_filter": "all",
                    "all_buyer_list": [],
                    "all_Customer": []
                }
            ],
            "kwargs": {
                "context": {"tz": "Asia/Dhaka", "uid": USER_ID, "allowed_company_ids": [1]},
                "specification": {
                    "report_type": {},
                    "date_from": {},
                    "date_to": {},
                    "order_filter": {},
                    "all_buyer_list": {"fields": {"display_name": {}}},
                    "all_Customer": {"fields": {"display_name": {}}}
                }
            }
        }
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw/ppc.report/web_save", json=payload)
    r.raise_for_status()
    resp = r.json()
    result = resp.get("result")
    
    if not result:
        raise Exception(f"❌ Failed to create report record. Response: {resp}")
        
    report_id = result[0].get("id")
    log.info(f"✅ Report record created with ID: {report_id}")
    return report_id

def download_excel(report_id):
    """Triggers the action to generate XLSX and downloads it."""
    # Step 1: Call action_generate_xlsx_report to build the report
    log.info("Triggering Excel generation action...")
    action_payload = {
        "id": 133,
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "args": [[report_id]],
            "kwargs": {
                "context": {"tz": "Asia/Dhaka", "uid": USER_ID, "allowed_company_ids": [1]}
            },
            "method": "action_generate_xlsx_report",
            "model": "ppc.report"
        }
    }
    session.post(f"{ODOO_URL}/web/dataset/call_button", json=action_payload)
    
    # Step 2: Download the file via /report/download
    log.info("Downloading Excel file...")
    download_url = f"{ODOO_URL}/report/download"
    data_val = f'["/report/xlsx/taps_manufacturing.pending_slider_count/{report_id}?context=%7B%22tz%22%3A%22Asia%2FDhaka%22%2C%22uid%22%3A{USER_ID}%2C%22allowed_company_ids%22%3A%5B1%5D%7D","xlsx"]'
    context_val = json.dumps({"tz": "Asia/Dhaka", "uid": USER_ID, "allowed_company_ids": [1]})
    
    # Using specific form-data boundary as expected by Odoo
    files = {
        "data": (None, data_val),
        "context": (None, context_val),
        "token": (None, "dummy-because-api-expects-one"),
        "csrf_token": (None, "dummy_token") # Odoo doesn't strictly check csrf if proper session
    }
    
    r = session.post(download_url, files=files)
    r.raise_for_status()
    
    log.info(f"✅ Excel downloaded. Size: {len(r.content)} bytes.")
    return r.content

def upload_to_google_sheet(excel_content):
    log.info("Reading downloaded Excel into Pandas...")
    df = pd.read_excel(io.BytesIO(excel_content))
    
    if df.empty:
        log.warning("⚠️ DataFrame is empty. Skipping upload.")
        return
        
    df = df.replace(False, "")
    df = df.fillna("")
    
    if not os.path.exists(GCP_CREDENTIALS_FILE):
        raise FileNotFoundError(f"❌ GCP Credentials file '{GCP_CREDENTIALS_FILE}' not found!")

    log.info("Connecting to Google Sheets...")
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_file(GCP_CREDENTIALS_FILE, scopes=scope)
    client = gspread.authorize(creds)
    
    sheet = client.open_by_key(SPREADSHEET_ID)
    
    # Check if the worksheet exists, otherwise get first or create:
    try:
        worksheet = sheet.worksheet(WORKSHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        log.warning(f"Worksheet '{WORKSHEET_NAME}' not found. Selecting the first sheet.")
        worksheet = sheet.get_worksheet(0)
    
    log.info(f"Clearing existing data on {worksheet.title}...")
    worksheet.clear()
    
    log.info("Pasting new data...")
    set_with_dataframe(worksheet, df, include_index=False, include_column_header=True)
    
    log.info(f"✅ Data successfully updated in Google Sheets!")

def main():
    try:
        now = check_first_day()
        date_from, date_to = get_previous_month_dates(now)
        
        login()
        report_id = generate_report(date_from, date_to)
        excel_bytes = download_excel(report_id)
        upload_to_google_sheet(excel_bytes)
    except Exception as e:
        log.exception(f"❌ Critical Error: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()
