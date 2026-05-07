"""
metabase_to_sheets.py
Fetches results from Metabase card/question URLs and writes each to a
separate tab in a Google Sheet.

Environment variables (set as GitHub Actions secrets):
  METABASE_URL        e.g. https://metabase.yourcompany.co
  METABASE_USERNAME   your login email
  METABASE_PASSWORD   your login password
  GOOGLE_SHEET_ID     the long ID from your sheet URL
  GCP_SERVICE_ACCOUNT_JSON  full contents of your service account JSON
"""

import os
import json
import re
import sys
import time
from datetime import datetime

import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ── Config ────────────────────────────────────────────────────────────────────

METABASE_URL      = os.environ["METABASE_URL"].rstrip("/")
METABASE_USERNAME = os.environ["METABASE_USERNAME"]
METABASE_PASSWORD = os.environ["METABASE_PASSWORD"]
GOOGLE_SHEET_ID   = os.environ["GOOGLE_SHEET_ID"]
GCP_SA_JSON       = os.environ["GCP_SERVICE_ACCOUNT_JSON"]

# ── List your Metabase card/question URLs here ────────────────────────────────
# Format: ("Tab name in Sheet", "full Metabase URL to card or question")
# Supports both /question/NNN and /card/NNN URLs.

QUERIES = [
    ("All Time RFD Details",   "https://metabase-lierhfgoeiwhr.newtonschool.co/question/10317-rfd-details-excl-cancellation-all-time"),
    ("Perf RFDs",              "https://metabase-lierhfgoeiwhr.newtonschool.co/question/10293-overall-perf-rfds-all-time"),
    ("Overall Inbound RFDs",   "https://metabase-lierhfgoeiwhr.newtonschool.co/question/10296-overall-inbound-rfds-all-time"),
    ("Perf Inbound RFDs",      "https://metabase-lierhfgoeiwhr.newtonschool.co/question/10297-perf-inbound-rfds-all-time"),
    ("Organic",                "https://metabase-lierhfgoeiwhr.newtonschool.co/question/10298-organic-inbound-rfds-all-time"),
    ("Masterclass",            "https://metabase-lierhfgoeiwhr.newtonschool.co/question/10288-masterclass-rfds-all-time"),
    ("Referral",               "https://metabase-lierhfgoeiwhr.newtonschool.co/question/10287-referral-rfds-all-time"),
    ("Open Funnel",            "https://metabase-lierhfgoeiwhr.newtonschool.co/question/10289-open-funnel-rfds-all-time"),
    ("Reapplied",              "https://metabase-lierhfgoeiwhr.newtonschool.co/question/10290-reapplied-rfds-all-time"),
    ("Reactivation",           "https://metabase-lierhfgoeiwhr.newtonschool.co/question/10292-reactivation-rfds-all-time"),
]

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ── Helpers ───────────────────────────────────────────────────────────────────

def metabase_session_token() -> str:
    """Authenticate and return a Metabase session token."""
    resp = requests.post(
        f"{METABASE_URL}/api/session",
        json={"username": METABASE_USERNAME, "password": METABASE_PASSWORD},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["id"]


def extract_card_id(url: str) -> int:
    """Pull the numeric card/question ID out of a Metabase URL."""
    match = re.search(r"/(?:question|card)/(\d+)", url)
    if not match:
        raise ValueError(f"Cannot extract card ID from URL: {url}")
    return int(match.group(1))


def fetch_card_data(card_id: int, session_token: str) -> tuple[list[str], list[list]]:
    """
    Run a saved Metabase card and return (column_names, rows).
    Uses the /api/card/:id/query/json endpoint which returns up to 10 000 rows.
    For larger result sets the CSV export endpoint is used instead.
    """
    headers = {"X-Metabase-Session": session_token}

    # Try JSON first (fast, 10 000 row limit)
    resp = requests.post(
        f"{METABASE_URL}/api/card/{card_id}/query/json",
        headers=headers,
        timeout=300,
    )

    if resp.status_code == 200:
        data = resp.json()
        if isinstance(data, list) and data:
            cols = list(data[0].keys())
            rows = [[str(row.get(c, "")) for c in cols] for row in data]
            return cols, rows
        return [], []

    # Fallback: CSV export (handles > 10 000 rows)
    resp = requests.post(
        f"{METABASE_URL}/api/card/{card_id}/query/csv",
        headers=headers,
        timeout=180,
    )
    resp.raise_for_status()

    import csv, io
    reader = csv.reader(io.StringIO(resp.text))
    rows_raw = list(reader)
    if not rows_raw:
        return [], []
    return rows_raw[0], rows_raw[1:]


def sheets_service():
    """Build and return an authenticated Google Sheets service."""
    creds_info = json.loads(GCP_SA_JSON)
    creds = service_account.Credentials.from_service_account_info(
        creds_info, scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def ensure_sheet_tab(service, spreadsheet_id: str, tab_name: str) -> int:
    """
    Make sure a tab with `tab_name` exists.
    Returns its sheetId (needed for clearValues).
    Creates it if missing.
    """
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in meta["sheets"]:
        if sheet["properties"]["title"] == tab_name:
            return sheet["properties"]["sheetId"]

    # Create it
    body = {"requests": [{"addSheet": {"properties": {"title": tab_name}}}]}
    resp = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body=body
    ).execute()
    return resp["replies"][0]["addSheet"]["properties"]["sheetId"]


def write_tab(service, spreadsheet_id: str, tab_name: str, cols: list, rows: list):
    """Clear the tab and write headers + data rows."""
    # Clear existing content
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=f"'{tab_name}'!A:T",
    ).execute()

    if not cols:
        print(f"  ⚠️  No data returned — tab '{tab_name}' cleared.")
        return

    # Timestamp row so you know when it last ran
    timestamp_row = [f"Last updated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}"]
    values = [timestamp_row, cols] + rows

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'{tab_name}'!A1",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

    print(f"  ✅  '{tab_name}' → {len(rows)} rows written.")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("🔐 Authenticating with Metabase...")
    token = metabase_session_token()
    print("✅ Metabase session OK")

    print("🔐 Building Google Sheets client...")
    svc = sheets_service()
    print("✅ Sheets client OK")

    errors = []
    for tab_name, url in QUERIES:
        try:
            card_id = extract_card_id(url)
            print(f"\n📊 Fetching card {card_id} → '{tab_name}'")
            cols, rows = fetch_card_data(card_id, token)
            ensure_sheet_tab(svc, GOOGLE_SHEET_ID, tab_name)
            write_tab(svc, GOOGLE_SHEET_ID, tab_name, cols, rows)
            time.sleep(0.5)  # gentle rate limiting
        except Exception as e:
            msg = f"❌ '{tab_name}' (card {url}): {e}"
            print(msg)
            errors.append(msg)

    if errors:
        print("\n\n⚠️  Completed with errors:")
        for e in errors:
            print(" ", e)
        sys.exit(1)
    else:
        print("\n\n🎉 All queries written successfully.")


if __name__ == "__main__":
    main()
