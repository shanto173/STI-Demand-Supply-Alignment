import os
import sys
import logging
import time
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pytz
import pandas as pd
import requests
from google.oauth2 import service_account
import gspread
from gspread_dataframe import set_with_dataframe
from dotenv import load_dotenv
load_dotenv()
# ===== Setup Logging =====
logging.basicConfig(stream=sys.stdout, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
log = logging.getLogger()

# ===== Load environment variables =====
load_dotenv()
ODOO_URL = os.getenv("ODOO_URL")
DB = os.getenv("ODOO_DB")
USERNAME = os.getenv("ODOO_USERNAME", "supply.chain3@texzipperbd.com")
PASSWORD = os.getenv("ODOO_PASSWORD", "@Shanto@86")

# Google Sheets info
SPREADSHEET_ID = "1PECBN0MiOpOeVa3XvSpit5Xb5fLT91HxbXNvEUpbqM8"
WORKSHEET_NAME = "Forecast_raw_data"
GCP_CREDENTIALS_FILE = "rugged-feat-456510-q7-381d4ee0fb45.json"

session = requests.Session()
USER_ID = None

def get_target_months():
    """Returns the current month and the next 2 months in 'YYYY-MM' format."""
    now = datetime.now()
    months = []
    for i in range(3):
        target = now + relativedelta(months=i)
        months.append(target.strftime("%Y-%m"))
    return months

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

def fetch_forecast_data():
    months = get_target_months()
    log.info(f"Fetching data for months: {months}")
    
    # We use web_search_read with the new ORM specification format mapped from HAR.
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "rolling.forecast.line",
            "method": "web_search_read",
            "args": [],
            "kwargs": {
                "specification": {
                    "buyer": {"fields": {"display_name": {}}},
                    "brand_group": {"fields": {"display_name": {}}},
                    "classification": {},
                    "customer_name": {"fields": {"display_name": {}}},
                    "customer_group": {"fields": {"display_name": {}}},
                    "forecast_product_id": {"fields": {"display_name": {}}},
                    "item_category": {"fields": {"display_name": {}}},
                    "next_month": {},
                    "projection_number": {},
                    "qty": {},
                    "sales_team_id": {"fields": {"display_name": {}}},
                    "salesperson_id": {"fields": {"display_name": {}}},
                    "segments": {},
                    "state": {},
                    "total_price": {},
                    "company_id": {"fields": {"display_name": {}}},
                    "avg_price": {}
                },
                "offset": 0,
                "limit": 100000,
                "context": {
                    "tz": "Asia/Dhaka",
                    "bin_size": True
                },
                "domain": [
                    "&", ["state", "!=", "cancel"],
                    ["next_month", "in", months]
                ]
            }
        }
    }
    
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw/rolling.forecast.line/web_search_read", json=payload)
    r.raise_for_status()
    response_json = r.json()
    
    if 'error' in response_json:
        raise Exception(f"❌ API Error: {response_json['error']}")
        
    records = response_json.get("result", {}).get("records", [])
    log.info(f"✅ Fetched {len(records)} records from Odoo.")
    return records

def process_data(records):
    if not records:
        return pd.DataFrame()
        
    df = pd.DataFrame(records)
    
    def extract_display_name(val):
        if isinstance(val, dict):
            return val.get('display_name', '')
        elif isinstance(val, list) and len(val) == 2:
            return val[1]
        elif pd.isna(val) or val is False:
            return ""
        return val

    # Fields that need M2O parsing
    m2o_fields = [
        "buyer", "brand_group", "customer_name", "customer_group",
        "forecast_product_id", "item_category", "sales_team_id", 
        "salesperson_id", "company_id"
    ]
    
    for col in m2o_fields:
        if col in df.columns:
            df[col] = df[col].apply(extract_display_name)
    
    # Rename columns to match Labels requested
    rename_mapping = {
        "buyer": "Brand",
        "brand_group": "Brand Group",
        "classification": "Classification",
        "customer_name": "Customer",
        "customer_group": "Customer Group",
        "forecast_product_id": "Forecast Product",
        "item_category": "Item",
        "next_month": "Next Month ",
        "projection_number": "Projection Number",
        "qty": "Qty",
        "sales_team_id": "Sales Team",
        "salesperson_id": "Salesperson",
        "segments": "Segments",
        "state": "Status",
        "total_price": "Total Price",
        "company_id": "Unit",
        "avg_price": "Unit Price"
    }
    
    df = df.rename(columns=rename_mapping)
    
    # Ensure correct column order if possible, otherwise just keep the mapped ones
    expected_cols = list(rename_mapping.values())
    actual_cols = [c for c in expected_cols if c in df.columns]
    
    df = df[actual_cols]
    df = df.replace(False, "")
    df = df.fillna("")
    
    return df

def upload_to_google_sheet(df):
    if df.empty:
        log.warning("⚠️ DataFrame is empty. Skipping upload.")
        return
        
    if not os.path.exists(GCP_CREDENTIALS_FILE):
        raise FileNotFoundError(f"❌ GCP Credentials file '{GCP_CREDENTIALS_FILE}' not found!")

    log.info("Connecting to Google Sheets...")
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_file(GCP_CREDENTIALS_FILE, scopes=scope)
    client = gspread.authorize(creds)
    
    sheet = client.open_by_key(SPREADSHEET_ID)
    worksheet = sheet.worksheet(WORKSHEET_NAME)
    
    log.info("Clearing existing data...")
    worksheet.clear()
    
    log.info("Pasting new data...")
    set_with_dataframe(worksheet, df, include_index=False, include_column_header=True)
    
    # Optional timestamp update in a separate cell if space permits (e.g. AA2)
    local_tz = pytz.timezone('Asia/Dhaka')
    local_time = datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
    # For now, appending it at the very right
    # worksheet.update("AA2", [[f"Last Updated: {local_time}"]])
    log.info(f"✅ Data successfully updated in Google Sheets!")

def main():
    try:
        login()
        records = fetch_forecast_data()
        df = process_data(records)
        upload_to_google_sheet(df)
    except Exception as e:
        log.exception(f"❌ Critical Error: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()
