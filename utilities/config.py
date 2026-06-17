"""
Shared configuration and helpers for LOTS LARP local utilities.
Secrets are loaded from a .env file in the utilities/ directory.
Copy .env-example to .env and fill in your values before running any script.
"""

import os
import re
import sys
import time
import gspread
from gspread.exceptions import APIError
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))

def _require(key):
    val = os.getenv(key)
    if not val:
        print(f"ERROR: {key} is not set. Copy utilities/.env-example to utilities/.env and fill it in.")
        sys.exit(1)
    return val

# ── Configuration (loaded from .env) ──────────────────────────────────────────

SPREADSHEET_ID       = _require('SPREADSHEET_ID')
SERVICE_ACCOUNT_JSON = _require('SERVICE_ACCOUNT_JSON')
BOT_TOKEN            = _require('BOT_TOKEN')
GUILD_ID             = _require('GUILD_ID')

# ── Constants (mirror GAS Constants.js) ───────────────────────────────────────

TIMESTAMP_COL        = 1   # all 1-indexed, matching GAS
STATUS_COL           = 2
SEND_DISCORD_COL     = 3
CHARACTER_NAME_COL   = 5

CHAR_SHEET_NAME_COL         = 1
CHAR_SHEET_WEBHOOK_COL      = 24  # Column X
CHAR_SHEET_CHANNEL_ID_COL   = 25  # Column Y
CHAR_SHEET_CHANNEL_NAME_COL = 26  # Column Z

CHARACTERS_SHEET   = "Characters"
MONTH_YEAR_PATTERN = re.compile(r'^\w+ \d{4}$')
MAX_MESSAGE_LEN    = 1800
DISCORD_API        = "https://discord.com/api/v10"

# ──────────────────────────────────────────────────────────────────────────────


def sheets_client():
    return gspread.service_account(
        filename=SERVICE_ACCOUNT_JSON,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )


def bot_headers():
    return {"Authorization": f"Bot {BOT_TOKEN}", "Content-Type": "application/json"}


def sheets_retry(fn, *args, max_retries=6, base_delay=2.0, notify=None, **kwargs):
    """
    Calls a gspread function, retrying on HTTP 429 (quota exceeded) and 5xx
    errors with exponential backoff (capped at 64s).

    notify: optional callable(message_str) used to report each wait — pass your
            rich console.print or a plain print. Re-raises on non-retryable
            errors or once retries are exhausted.
    """
    delay = base_delay
    for attempt in range(max_retries + 1):
        try:
            return fn(*args, **kwargs)
        except APIError as e:
            status = getattr(e.response, 'status_code', None)
            retryable = status == 429 or (status is not None and 500 <= status < 600)
            if retryable and attempt < max_retries:
                if notify:
                    reason = "quota exceeded" if status == 429 else f"server error {status}"
                    notify(f"Sheets {reason} — backing off {delay:.0f}s "
                           f"(attempt {attempt + 1}/{max_retries})…")
                time.sleep(delay)
                delay = min(delay * 2, 64.0)
                continue
            raise
