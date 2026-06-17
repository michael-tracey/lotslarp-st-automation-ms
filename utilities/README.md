# LOTS LARP — Local Utilities

These scripts run on your local machine and talk to Discord's bot API directly,
bypassing the IP restrictions that prevent the Google Apps Script from making
authenticated bot calls.

## Setup

**1. Install dependencies**
```bash
cd utilities
pip install -r requirements.txt
```

**2. Create your `.env` file**
```bash
cp .env-example .env
```
Then edit `.env` and fill in:

| Variable | Where to find it |
|---|---|
| `SPREADSHEET_ID` | The long ID in the Google Sheet URL |
| `SERVICE_ACCOUNT_JSON` | Path to your downloaded service account key file |
| `BOT_TOKEN` | Discord Developer Portal → your bot → Token |
| `GUILD_ID` | Right-click your Discord server icon → Copy Server ID |
| `TEST_CHANNEL_ID` | *(Optional)* Channel ID for dry-run level 2 |

**3. Add your service account key**

Download the JSON key for your Google service account and place it in this
directory (the default expected filename is `service_account.json`). Make sure
the service account has been granted **Editor** access to the spreadsheet.

Both `.env` and `service_account.json` are gitignored and will never be committed.

---

## Scripts

### `send_downtimes.py` — Send downtime results to players via Discord

Reads the active downtime sheet, finds all unsent response rows, and posts each
player's results to their dedicated Discord channel. Marks rows as sent in the
sheet when done.

```bash
# Interactive sheet picker (recommends most recent)
python send_downtimes.py

# Skip the picker and use a specific sheet
python send_downtimes.py "June 2026"

# Dry run — prompts you to choose a level
python send_downtimes.py --dry-run
python send_downtimes.py "June 2026" --dry-run

# Include rows already marked sent (override the skip)
python send_downtimes.py --ignore-sent
```

**Send mode** (prompted on every run, dry or live):

| Mode | Behaviour |
|---|---|
| Automated | Sends everything; use the `P`/`S` keys to pause or stop mid-run |
| Confirm each | Prompts `Y/n/a/q` before every character |

In confirm mode: `Y`/Enter sends, `n` skips, `a` sends this and all remaining
(switches to automated), `q` quits.

**Dry-run levels:**

| Level | Behaviour | Sheet updated? |
|---|---|---|
| 1 | Prints each message to console in a bordered panel | No |
| 2 | Sends every message to `TEST_CHANNEL_ID` instead of player channels | No |
| 3 | Creates one thread in `TEST_CHANNEL_ID` (named with the sheet + current time); posts all messages there with a separator between each character | No |
| 4 | Creates one thread per character name in `TEST_CHANNEL_ID`; posts each message into its own thread | No |

Levels 2, 3, and 4 all require `TEST_CHANNEL_ID` to be set in `.env`.

**Controls during an automated send:**

| Key | Action |
|---|---|
| `P` or `Space` | Pause / resume |
| `S` or `Q` | Stop early — finishes the current character cleanly, then shows results |

The script skips rows already marked sent — either the **Status** column reads
`sent` or the **Send Discord** checkbox is ticked. Each skip is logged. Stopping
early and re-running is safe; already-sent rows won't be double-delivered. Use
`--ignore-sent` to deliberately re-send rows that are already marked sent.

---

### `populate_channels.py` — Fill in Discord channel IDs and names

One-time (and occasional maintenance) script. Reads the Characters sheet, calls
the Discord guild API to match each character's webhook URL to a channel, then
writes the channel ID (col Y) and channel name (col Z) back to the sheet.

Only touches rows that are missing an ID or have a bad/placeholder name.
Rows that already have both are skipped.

```bash
python populate_channels.py
```

Run this before your first `send_downtimes.py` run, and again whenever new
players are added or webhooks are changed.
