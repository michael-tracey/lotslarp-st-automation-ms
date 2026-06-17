#!/usr/bin/env python3
"""
LOTS LARP — Discord Channel Info Populator
Reads the Characters sheet, queries the Discord guild API, and fills in
Discord Channel ID (col Y) and Discord Channel Name (col Z) by matching
each character's webhook URL to a guild webhook, then looking up the channel.

Setup:
  pip install -r requirements.txt
  cp .env-example .env  # fill in your values

Usage:
  python populate_channels.py
"""

import os
import re
import sys

import requests
from gspread.utils import rowcol_to_a1
from rich.console import Console
from rich.table import Table
from rich import box

sys.path.insert(0, os.path.dirname(__file__))
from config import (
    SPREADSHEET_ID, DISCORD_API, GUILD_ID,
    CHAR_SHEET_NAME_COL, CHAR_SHEET_WEBHOOK_COL,
    CHAR_SHEET_CHANNEL_ID_COL, CHAR_SHEET_CHANNEL_NAME_COL,
    CHARACTERS_SHEET,
    sheets_client, bot_headers, sheets_retry,
)

console = Console()


def populate_channel_info(spreadsheet):
    """
    Fills channel IDs and names into the Characters sheet using the guild API.
    Two API calls total: guild webhooks + guild channels.
    Writes col Y (channel ID) and col Z (channel name) for each matched character.
    Skips rows that already have both a valid ID and a good name.
    """
    console.print("Fetching all guild webhooks…")
    wh_resp = requests.get(f"{DISCORD_API}/guilds/{GUILD_ID}/webhooks", headers=bot_headers())
    if wh_resp.status_code != 200:
        console.print(f"  [red]Guild webhooks fetch failed:[/red] {wh_resp.status_code} {wh_resp.text}")
        sys.exit(1)

    webhook_to_channel = {}  # webhook_id → channel_id
    for wh in wh_resp.json():
        if wh.get('id') and wh.get('channel_id'):
            webhook_to_channel[wh['id']] = wh['channel_id']
    console.print(f"  [dim]Found {len(webhook_to_channel)} webhooks in guild.[/dim]")

    console.print("Fetching all guild channels…")
    ch_resp = requests.get(f"{DISCORD_API}/guilds/{GUILD_ID}/channels", headers=bot_headers())
    channel_to_name = {}
    if ch_resp.status_code == 200:
        for ch in ch_resp.json():
            if ch.get('id') and ch.get('name'):
                channel_to_name[ch['id']] = ch['name']
        console.print(f"  [dim]Found {len(channel_to_name)} channels in guild.[/dim]")
    else:
        console.print(f"  [yellow]Channel fetch failed ({ch_resp.status_code}) — names won't be populated.[/yellow]")

    ws   = spreadsheet.worksheet(CHARACTERS_SHEET)
    rows = ws.get_all_values()

    ids_added   = 0
    names_added = 0
    not_found   = 0
    skipped     = 0
    updates     = []  # batched cell writes: {'range': 'Y2', 'values': [[val]]}

    console.print(f"\nProcessing {len(rows) - 1} character rows…")

    for i, row in enumerate(rows[1:], start=2):  # start=2 → 1-indexed sheet row
        while len(row) < CHAR_SHEET_CHANNEL_NAME_COL:
            row.append('')

        name        = row[CHAR_SHEET_NAME_COL - 1].strip()
        webhook_url = row[CHAR_SHEET_WEBHOOK_COL - 1].strip()
        existing_id = row[CHAR_SHEET_CHANNEL_ID_COL - 1].strip()
        existing_nm = row[CHAR_SHEET_CHANNEL_NAME_COL - 1].strip()
        name_is_good = existing_nm and not existing_nm.startswith('(name')

        if not name or not webhook_url:
            continue

        if existing_id and name_is_good:
            skipped += 1
            continue

        # Extract webhook ID from URL: https://discord.com/api/webhooks/{id}/{token}
        m = re.search(r'/webhooks/(\d+)/', webhook_url)
        if not m:
            console.print(f"  [yellow]Row {i} ({name}):[/yellow] unrecognised webhook URL — skipping.")
            continue

        channel_id = webhook_to_channel.get(m.group(1))
        if not channel_id:
            console.print(f"  [yellow]Row {i} ({name}):[/yellow] webhook not found in guild — skipping.")
            not_found += 1
            continue

        if not existing_id:
            updates.append({
                'range':  rowcol_to_a1(i, CHAR_SHEET_CHANNEL_ID_COL),
                'values': [[channel_id]],
            })
            ids_added += 1
            console.print(f"  [green]ID[/green]   row {i} ({name}): {channel_id}")

        if not name_is_good and channel_to_name.get(channel_id):
            updates.append({
                'range':  rowcol_to_a1(i, CHAR_SHEET_CHANNEL_NAME_COL),
                'values': [[channel_to_name[channel_id]]],
            })
            names_added += 1
            console.print(f"  [green]Name[/green] row {i} ({name}): {channel_to_name[channel_id]}")

    # ── Write all updates in one batched call (retries on quota errors) ─────────
    if updates:
        console.print(f"\nWriting {len(updates)} cells to the sheet in one batch…")
        sheets_retry(ws.batch_update, updates, notify=lambda msg: console.print(f"  [yellow]{msg}[/yellow]"))
        console.print("  [dim]Done.[/dim]")

    # ── Summary ───────────────────────────────────────────────────────────────
    console.print()
    table = Table(title="Populate complete", show_header=False, box=box.ROUNDED, min_width=36)
    table.add_column("", style="bold", no_wrap=True)
    table.add_column("", justify="right")
    if ids_added:
        table.add_row("[green]IDs written[/green]",   str(ids_added))
    if names_added:
        table.add_row("[green]Names written[/green]", str(names_added))
    if skipped:
        table.add_row("[dim]Already complete[/dim]",  str(skipped))
    if not_found:
        table.add_row("[yellow]Not in guild[/yellow]", str(not_found))
    console.print(table)


def main():
    console.print("Connecting to Google Sheets…")
    client      = sheets_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    populate_channel_info(spreadsheet)


if __name__ == "__main__":
    main()
