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
from gspread_dataframe import set_with_dataframe, get_as_dataframe
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

# Google Sheets info
SPREADSHEET_ID = "1acV7UrmC8ogC54byMrKRTaD9i1b1Cf9QZ-H1qHU5ZZc"
SNAPSHOT_SHEET = "pending_pi_data"
DAYWISE_SHEET = "pending_pi_data_day_wise"
GCP_CREDENTIALS_FILE = "rugged-feat-456510-q7-381d4ee0fb45.json"

DHAKA_TZ = pytz.timezone('Asia/Dhaka')

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

    log.info("Fetching pi.pending.view export from Odoo (xlsx)...")
    export_url = f"{ODOO_URL}/web/export/xlsx"

    export_data = {
        "import_compat": False,
        "context": {
            "lang": "en_US",
            "tz": "Asia/Dhaka",
            "uid": USER_ID,
            "allowed_company_ids": [1, 3, 2, 4]
        },
        "domain": [],
        "fields": [
            {"name": "buyer", "label": "Buyer"},
            {"name": "company_id", "label": "Company"},
            {"name": "customer_name", "label": "Customer"},
            {"name": "item", "label": "Item"},
            {"name": "pi_amount_usd", "label": "Pending Amount"},
            {"name": "pi_quantity", "label": "Pending Quantity"},
            {"name": "pi_id", "label": "PI"},
            {"name": "pi_date", "label": "PI Date"},
            {"name": "pt_id", "label": "Product"},
            {"name": "salesperson_name", "label": "Salesperson"},
            {"name": "segment", "label": "Segment"},
            {"name": "pi_id/payment_term_id", "label": "PI/Payment Terms"},
            {"name": "pi_id/team_id", "label": "PI/Sales Team"},
            {"name": "pi_id/sales_type", "label": "PI/Sales Type"},
            {"name": "pi_id/user_id", "label": "PI/Salesperson"},
            {"name": "pt_id/fg_categ_type", "label": "Product/FG Category"},
            {"name": "pi_id/brand_group", "label": "PI/Brand Group"}
        ],
        "groupby": [],
        "ids": False,
        "model": "pi.pending.view"
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


def get_gspread_client():
    if not os.path.exists(GCP_CREDENTIALS_FILE):
        raise FileNotFoundError(f"❌ GCP Credentials file '{GCP_CREDENTIALS_FILE}' not found!")
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_file(GCP_CREDENTIALS_FILE, scopes=scope)
    return gspread.authorize(creds)


def get_or_create_worksheet(sheet, name, rows=2000, cols=30):
    try:
        return sheet.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        log.warning(f"Worksheet '{name}' not found. Creating it...")
        return sheet.add_worksheet(title=name, rows=str(rows), cols=str(cols))


def upload_snapshot(sheet, df, today_str):
    """Sheet 1: clear and replace with latest snapshot.

    Row 1 is reserved for user-maintained subtotal formulas, so data is
    written starting at A2 (header on row 2, rows from row 3 onward) and
    row 1 is left untouched.
    """
    worksheet = get_or_create_worksheet(sheet, SNAPSHOT_SHEET)
    log.info(f"Clearing old data on '{SNAPSHOT_SHEET}' from row 2 down (row 1 preserved)...")
    worksheet.batch_clear(["A2:ZZ"])

    log.info(f"Pasting {len(df)} rows into '{SNAPSHOT_SHEET}' starting at A2...")
    set_with_dataframe(worksheet, df, include_index=False, include_column_header=True, row=2, col=1)

    now_str = datetime.now(DHAKA_TZ).strftime("%Y-%m-%d %I:%M %p")
    timestamp_col = chr(ord('A') + len(df.columns))  # column right after data
    worksheet.update_acell(f"{timestamp_col}2", f"Updated: {now_str}")
    log.info(f"✅ '{SNAPSHOT_SHEET}' updated at {now_str}")


def upload_daywise(sheet, df, today_str):
    """Sheet 2: append today's rows with a Date column.

    - If today's date already exists in the sheet, remove those rows first
      (re-run during the same day overwrites that day's snapshot).
    - Previous days' rows are preserved.
    """
    worksheet = get_or_create_worksheet(sheet, DAYWISE_SHEET)

    # Build today's frame with a leading Date column
    today_df = df.copy()
    today_df.insert(0, "Date", today_str)

    # Read existing data
    try:
        existing = get_as_dataframe(worksheet, evaluate_formulas=True, header=0)
    except Exception as e:
        log.warning(f"Could not read existing day-wise data ({e}); treating as empty.")
        existing = pd.DataFrame()

    if existing is not None and not existing.empty:
        # Drop fully empty rows that gspread_dataframe may pad in
        existing = existing.dropna(how='all')
        # Drop fully empty columns (gspread_dataframe sometimes adds trailing NaN cols)
        existing = existing.loc[:, ~existing.columns.astype(str).str.startswith("Unnamed")]
        # Drop any leftover "Updated: ..." timestamp columns from older runs
        existing = existing.loc[:, ~existing.columns.astype(str).str.startswith("Updated")]
        existing = existing.fillna("")

        if "Date" in existing.columns:
            existing["Date"] = existing["Date"].astype(str)
            before = len(existing)
            existing = existing[existing["Date"] != today_str]
            removed = before - len(existing)
            if removed > 0:
                log.info(f"Removed {removed} existing rows for today ({today_str}) before re-appending.")

        # Align columns: union, preserving today's column order first
        all_cols = list(today_df.columns) + [c for c in existing.columns if c not in today_df.columns]
        existing = existing.reindex(columns=all_cols, fill_value="")
        today_df = today_df.reindex(columns=all_cols, fill_value="")
        combined = pd.concat([existing, today_df], ignore_index=True)
    else:
        combined = today_df

    combined = combined.fillna("")

    log.info(f"Writing {len(combined)} total rows ({len(today_df)} for {today_str}) to '{DAYWISE_SHEET}'...")
    worksheet.clear()
    set_with_dataframe(worksheet, combined, include_index=False, include_column_header=True)
    log.info(f"✅ '{DAYWISE_SHEET}' updated.")


def upload_to_google_sheets(excel_content):
    log.info("Reading Excel directly into Pandas...")
    df = pd.read_excel(io.BytesIO(excel_content))

    if df.empty:
        log.warning("⚠️ DataFrame is empty. Skipping upload.")
        return

    df = df.replace(False, "")
    df = df.fillna("")

    log.info("Connecting to Google Sheets...")
    client = get_gspread_client()
    sheet = client.open_by_key(SPREADSHEET_ID)

    today_str = datetime.now(DHAKA_TZ).strftime("%Y-%m-%d")

    upload_snapshot(sheet, df, today_str)
    upload_daywise(sheet, df, today_str)


def main():
    try:
        login()
        excel_bytes = fetch_pending_pi_data()
        upload_to_google_sheets(excel_bytes)
    except Exception as e:
        log.exception(f"❌ Critical Error: {str(e)}")
        sys.exit(1)


if __name__ == '__main__':
    main()
