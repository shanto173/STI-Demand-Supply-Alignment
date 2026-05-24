import os
import sys
import logging
import io
import json
import re
from datetime import datetime
import pytz
import pandas as pd
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
USERNAME = os.getenv("ODOO_USERNAME", "")
PASSWORD = os.getenv("ODOO_PASSWORD", "")

# Google Sheets info
SPREADSHEET_ID = "1PECBN0MiOpOeVa3XvSpit5Xb5fLT91HxbXNvEUpbqM8"
WORKSHEET_NAME = "Pending_pi"
GCP_CREDENTIALS_FILE = "rugged-feat-456510-q7-381d4ee0fb45.json"

session = requests.Session()
USER_ID = None

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

def fetch_pending_pi_data():
    log.info("Fetching CSRF Token...")
    r_web = session.get(f"{ODOO_URL}/web")
    csrf_match = re.search(r'csrf_token:\s*"([^"]+)"', r_web.text)
    csrf_token = csrf_match.group(1) if csrf_match else "dummy_token"
    
    log.info("Fetching Pending PI from Odoo via Excel Export...")
    
    export_url = f"{ODOO_URL}/web/export/xlsx"
    
    export_data = {
        "import_compat": False,
        "context": {
            "lang": "en_US",
            "tz": "Asia/Dhaka",
            "uid": USER_ID,
            "allowed_company_ids": [1, 2, 3, 4]
        },
        "domain": [
            "&", "&", "&", "&", "&", 
            ["company_id", "=", 1], 
            ["sales_type", "=", "sale"], 
            "|", 
            ["oa_count", "=", False], 
            ["oa_count", "=", 0], 
            ["is_active", "=", True], 
            ["pi_type", "=", "regular"], 
            ["state", "!=", "cancel"]
        ],
        "fields": [
            {"name": "order_line/order_id/name", "label": "Order Lines/Order Reference/Order Reference"},
            {"name": "order_line/order_id/buyer_name", "label": "Order Lines/Order Reference/Buyer"},
            {"name": "order_line/order_id/buyer_name/brand", "label": "Order Lines/Order Reference/Buyer/Brand Group"},
            {"name": "order_line/order_id/buying_house", "label": "Order Lines/Order Reference/Buying House"},
            {"name": "order_line/order_id/company_id", "label": "Order Lines/Order Reference/Company"},
            {"name": "order_line/order_id/partner_id", "label": "Order Lines/Order Reference/Customer"},
            {"name": "order_line/order_id/partner_id/group", "label": "Order Lines/Order Reference/Customer/Group"},
            {"name": "order_line/order_id/pi_date", "label": "Order Lines/Order Reference/PI Date"},
            {"name": "order_line/order_id/team_id", "label": "Order Lines/Order Reference/Sales Team"},
            {"name": "order_line/order_id/user_id", "label": "Order Lines/Order Reference/Salesperson"},
            {"name": "order_line/product_template_id/fg_categ_type", "label": "Order Lines/Product Template/FG Category"},
            {"name": "order_line/product_uom_qty", "label": "Order Lines/Quantity"},
            {"name": "order_line/price_total", "label": "Order Lines/Total"},
            {"name": "order_line/slidercodesfg", "label": "Order Lines/Slider Code (SFG)"},
            {"name": "order_line/order_id/lc_number", "label": "Order Lines/Order Reference/LC Number"},
            {"name": "order_line/order_id/payment_term_id", "label": "Order Lines/Order Reference/Payment Terms"},
            {"name": "order_line/order_id/state", "label": "Order Lines/Order Reference/Status"},
            {"name": "order_line/order_id/pi_type", "label": "Order Lines/Order Reference/Type"}
        ],
        "groupby": [],
        "ids": False,
        "model": "sale.order"
    }

    files = {
        "data": (None, json.dumps(export_data)),
        "token": (None, "dummy-because-api-expects-one"),
        "csrf_token": (None, csrf_token) 
    }
    
    r = session.post(export_url, files=files)
    r.raise_for_status()
    log.info(f"✅ Downloaded Excel data, size: {len(r.content)} bytes.")
    
    return r.content

def upload_to_google_sheet(excel_content):
    log.info("Reading Excel directly into Pandas...")
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
    
    try:
        worksheet = sheet.worksheet(WORKSHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        log.warning(f"Worksheet '{WORKSHEET_NAME}' not found. Selecting the first sheet instead.")
        worksheet = sheet.get_worksheet(0)
    
    log.info(f"Clearing old data from columns A:S on {worksheet.title}...")
    worksheet.batch_clear(["A:S"])
    
    log.info("Pasting data into the sheet...")
    set_with_dataframe(worksheet, df, include_index=False, include_column_header=True)
    
    dhaka_tz = pytz.timezone('Asia/Dhaka')
    now_str = datetime.now(dhaka_tz).strftime("%Y-%m-%d %I:%M %p")
    
    log.info(f"Setting Timestamp in S1: {now_str}")
    worksheet.update_acell("S1", f"Updated: {now_str}")
    
    log.info(f"✅ Data successfully updated in Google Sheets!")

def main():
    try:
        login()
        excel_bytes = fetch_pending_pi_data()
        upload_to_google_sheet(excel_bytes)
    except Exception as e:
        log.exception(f"❌ Critical Error: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()
