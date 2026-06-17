#!/usr/bin/env python3
"""
LOTS LARP — Local Downtime Discord Sender
Reads pending downtime results from Google Sheets and sends them via the
Discord bot API from your local machine (bypasses GAS IP restrictions).

Setup:
  pip install -r requirements.txt
  cp .env-example .env  # fill in your values

Usage:
  python send_downtimes.py                  # interactive sheet picker, live send
  python send_downtimes.py "June 2026"     # skip the sheet picker, live send
  python send_downtimes.py --dry-run       # interactive sheet + dry-run level picker
  python send_downtimes.py --ignore-sent   # include rows already marked sent

Every run also prompts for a send mode:
  Automated      send everything, with P/S controls during the run
  Confirm each   Y/N prompt before each character

Dry-run levels (prompted when --dry-run is passed):
  1   Print messages to console only
  2   Send all to TEST_CHANNEL_ID
  3   Create one thread in TEST_CHANNEL_ID; post all messages there with separators
  4   Create a thread per character in TEST_CHANNEL_ID

Controls during an automated send:
  P / Space   pause / resume
  S / Q       stop early (shows results so far)

Confirm-each prompt:
  Y / Enter   send this character
  n           skip this character
  a           send this and all remaining (switches to automated)
  q           quit (shows results so far)
"""

import argparse
import os
import sys
import termios
import threading
import time
import tty
from datetime import datetime, timezone

import requests
from rich import box
from rich.console import Console
from rich.panel import Panel
from rich.progress import (
    BarColumn,
    MofNCompleteColumn,
    Progress,
    SpinnerColumn,
    TextColumn,
    TimeElapsedColumn,
)
from rich.table import Table

from gspread.utils import rowcol_to_a1

sys.path.insert(0, os.path.dirname(__file__))
from config import (
    SPREADSHEET_ID, DISCORD_API, MAX_MESSAGE_LEN,
    TIMESTAMP_COL, STATUS_COL, SEND_DISCORD_COL, CHARACTER_NAME_COL,
    CHAR_SHEET_NAME_COL, CHAR_SHEET_CHANNEL_ID_COL, CHAR_SHEET_CHANNEL_NAME_COL,
    CHARACTERS_SHEET, MONTH_YEAR_PATTERN,
    sheets_client, bot_headers, sheets_retry,
)

console = Console()


# ── Keyboard listener ─────────────────────────────────────────────────────────

class KeyboardListener:
    """Background thread — reads single keypresses without Enter.
    Only active when stdin is a real TTY; silently disabled otherwise."""

    _PAUSE = {'p', 'P', ' '}
    _STOP  = {'s', 'S', 'q', 'Q', '\x03'}  # \x03 = Ctrl-C in raw mode

    def __init__(self):
        self.paused  = False
        self.stopped = False
        self._fd     = None
        self._saved  = None
        self._active = False
        self._thread = threading.Thread(target=self._run, daemon=True)

    def start(self):
        if not sys.stdin.isatty():
            return
        try:
            self._fd    = sys.stdin.fileno()
            self._saved = termios.tcgetattr(self._fd)
            tty.setraw(self._fd)
            self._active = True
            self._thread.start()
        except Exception:
            pass

    def stop(self):
        self.stopped = True
        if self._active and self._saved is not None:
            try:
                termios.tcsetattr(self._fd, termios.TCSADRAIN, self._saved)
            except Exception:
                pass
        self._active = False

    def _run(self):
        try:
            while not self.stopped:
                ch = sys.stdin.read(1)
                if not ch:
                    break
                if ch in self._PAUSE:
                    self.paused = not self.paused
                elif ch in self._STOP:
                    self.stopped = True
                    break
        except Exception:
            pass


# ── Helpers ───────────────────────────────────────────────────────────────────

def chunk_message(message):
    if len(message) <= MAX_MESSAGE_LEN:
        return [message]
    chunks = []
    remaining = message
    while remaining:
        if len(remaining) <= MAX_MESSAGE_LEN:
            chunks.append(remaining)
            break
        split = remaining.rfind('\n', 0, MAX_MESSAGE_LEN)
        if split <= 0:
            split = remaining.rfind(' ', 0, MAX_MESSAGE_LEN)
        if split <= 0:
            split = MAX_MESSAGE_LEN
        chunks.append(remaining[:split].strip())
        remaining = remaining[split:].strip()
    return chunks


def get_channel_map(spreadsheet):
    ws   = spreadsheet.worksheet(CHARACTERS_SHEET)
    rows = ws.get_all_values()
    result = {}
    for row in rows[1:]:
        while len(row) < CHAR_SHEET_CHANNEL_NAME_COL:
            row.append('')
        name       = row[CHAR_SHEET_NAME_COL - 1].strip()
        channel_id = row[CHAR_SHEET_CHANNEL_ID_COL - 1].strip()
        ch_name    = row[CHAR_SHEET_CHANNEL_NAME_COL - 1].strip()
        if name:
            result[name.lower()] = (channel_id, ch_name)
    return result


def list_month_sheets(spreadsheet):
    candidates = [ws for ws in spreadsheet.worksheets()
                  if MONTH_YEAR_PATTERN.match(ws.title)]
    def sort_key(ws):
        try:
            return datetime.strptime(ws.title, "%B %Y")
        except ValueError:
            return datetime.min
    return sorted(candidates, key=sort_key, reverse=True)


def build_message(headers, submission_row, response_row, month, year):
    character_name = response_row[CHARACTER_NAME_COL - 1].strip()
    title = f"**Downtime Results for {character_name} ({month}, {year})**\n\n"
    body  = ""
    for j in range(CHARACTER_NAME_COL, len(headers)):
        header    = headers[j].strip() if j < len(headers) else ""
        sub_text  = submission_row[j].strip() if j < len(submission_row) else ""
        resp_text = response_row[j].strip()   if j < len(response_row)   else ""
        if not header or not resp_text:
            continue
        body += f"**{header}**\n"
        body += f"*Your Action:* {sub_text or '(No submission text found)'}\n"
        body += f"*Result:* {resp_text}\n\n"
    return (title + body) if body.strip() else None


# ── Discord thread helper ─────────────────────────────────────────────────────

def create_discord_thread(channel_id, name):
    """Creates a standalone public thread in a text channel. Returns thread_id or None."""
    url = f"{DISCORD_API}/channels/{channel_id}/threads"
    for _attempt in range(3):
        resp = requests.post(url, json={
            "name": name[:100],  # Discord max is 100 chars
            "type": 11,          # PUBLIC_THREAD
            "auto_archive_duration": 1440,
        }, headers=bot_headers())
        if resp.status_code in (200, 201):
            return resp.json()['id']
        elif resp.status_code == 429:
            time.sleep(resp.json().get('retry_after', 5) + 0.5)
        else:
            return None
    return None


# ── Interactive pickers ───────────────────────────────────────────────────────

def pick_sheet(spreadsheet):
    sheets = list_month_sheets(spreadsheet)
    if not sheets:
        console.print("[red]No 'Month Year' sheets found in the spreadsheet.[/red]")
        sys.exit(1)

    console.print("\n[bold]Available downtime sheets:[/bold]")
    for i, ws in enumerate(sheets):
        tag = "  [dim]← most recent (default)[/dim]" if i == 0 else ""
        console.print(f"  [cyan]{i + 1}.[/cyan] {ws.title}{tag}")

    while True:
        try:
            raw = input("\nSelect sheet [1]: ").strip()
        except (EOFError, KeyboardInterrupt):
            console.print()
            sys.exit(0)
        if not raw:
            return sheets[0].title
        try:
            idx = int(raw) - 1
            if 0 <= idx < len(sheets):
                return sheets[idx].title
        except ValueError:
            pass
        console.print(f"  [yellow]Enter a number between 1 and {len(sheets)}.[/yellow]")


def pick_dry_run_level():
    console.print("\n[bold]Dry-run level:[/bold]")
    console.print("  [cyan]1.[/cyan] Print messages to console only — no Discord calls, sheet unchanged")
    console.print("  [cyan]2.[/cyan] Send all to TEST_CHANNEL_ID — sheet unchanged")
    console.print("  [cyan]3.[/cyan] Create one thread in TEST_CHANNEL_ID; post all messages there with separators")
    console.print("  [cyan]4.[/cyan] Create a thread per character in TEST_CHANNEL_ID")
    while True:
        try:
            raw = input("\nSelect level [1]: ").strip()
        except (EOFError, KeyboardInterrupt):
            console.print()
            sys.exit(0)
        if not raw:
            return 1
        if raw in ('1', '2', '3', '4'):
            return int(raw)
        console.print("  [yellow]Enter 1, 2, 3, or 4.[/yellow]")


def pick_send_mode():
    """Returns confirm_each (bool): False = automated, True = prompt per character."""
    console.print("\n[bold]Send mode:[/bold]")
    console.print("  [cyan]1.[/cyan] Automated — send everything ([bold]P[/bold] pause / [bold]S[/bold] stop during run)")
    console.print("  [cyan]2.[/cyan] Confirm each — Y/N prompt before every character")
    while True:
        try:
            raw = input("\nSelect mode [1]: ").strip()
        except (EOFError, KeyboardInterrupt):
            console.print()
            sys.exit(0)
        if not raw or raw == '1':
            return False
        if raw == '2':
            return True
        console.print("  [yellow]Enter 1 or 2.[/yellow]")


def confirm_send(index, total, char_name, display, was_sent):
    """Per-character prompt. Returns 'y' (send), 'n' (skip), 'a' (all rest), 'q' (quit)."""
    tag = " [yellow](already sent — resending)[/yellow]" if was_sent else ""
    while True:
        try:
            raw = console.input(
                f"[cyan]\\[{index}/{total}][/cyan] Send to [bold]{char_name}[/bold] "
                f"→ #{display}{tag}?  [[bold]Y[/bold]/n/a/q] "
            ).strip().lower()
        except (EOFError, KeyboardInterrupt):
            return 'q'
        if raw in ('', 'y', 'yes'):
            return 'y'
        if raw in ('n', 'no'):
            return 'n'
        if raw in ('a', 'all'):
            return 'a'
        if raw in ('q', 'quit', 's', 'stop'):
            return 'q'
        console.print("  [yellow]Y = send, n = skip, a = send all remaining, q = quit[/yellow]")


# ── Core send logic ───────────────────────────────────────────────────────────

def send_downtimes(spreadsheet, sheet_name, dry_run=0, confirm_each=False, ignore_sent=False):
    test_channel_id = None
    if dry_run in (2, 3, 4):
        test_channel_id = os.getenv('TEST_CHANNEL_ID', '').strip()
        if not test_channel_id:
            console.print(f"[red]ERROR:[/red] TEST_CHANNEL_ID not set in .env — required for dry-run level {dry_run}")
            sys.exit(1)

    if dry_run == 1:
        console.print("[bold yellow]DRY RUN (level 1)[/bold yellow] — console output only, no Discord calls, sheet unchanged.\n")
    elif dry_run == 2:
        console.print(f"[bold yellow]DRY RUN (level 2)[/bold yellow] — sending to test channel [bold]{test_channel_id}[/bold], sheet unchanged.\n")
    elif dry_run == 3:
        console.print(f"[bold yellow]DRY RUN (level 3)[/bold yellow] — posting all to one thread in [bold]{test_channel_id}[/bold], sheet unchanged.\n")
    elif dry_run == 4:
        console.print(f"[bold yellow]DRY RUN (level 4)[/bold yellow] — creating one thread per character in [bold]{test_channel_id}[/bold], sheet unchanged.\n")

    console.print("Loading character channel IDs…")
    channel_map = get_channel_map(spreadsheet)
    mapped = sum(1 for cid, _ in channel_map.values() if cid)
    console.print(f"  [dim]{mapped} characters have channel IDs.[/dim]")

    sheet = spreadsheet.worksheet(sheet_name)
    parts = sheet.title.split()
    month, year = (parts[0], parts[1]) if len(parts) == 2 else ("?", "?")

    all_values = sheet.get_all_values()
    headers    = all_values[0]

    # ── First pass: collect pending work ──────────────────────────────────────
    pending      = []
    already_sent = []
    no_channel   = []
    no_responses = []

    for i in range(2, len(all_values), 2):
        response_row   = all_values[i]
        submission_row = all_values[i - 1]
        sheet_row      = i + 1

        while len(response_row) < max(STATUS_COL, SEND_DISCORD_COL, CHARACTER_NAME_COL):
            response_row.append('')

        character_name = response_row[CHARACTER_NAME_COL - 1].strip()
        status         = response_row[STATUS_COL - 1].strip().lower()
        discord_sent   = response_row[SEND_DISCORD_COL - 1].strip().upper() == 'TRUE'

        if not character_name:
            continue
        # Respect both the Status column and the Send Discord checkbox — either
        # one being set means this character was already delivered. Honoured in
        # every mode, including dry runs, so a dry run mirrors a real run.
        # --ignore-sent overrides this and includes them anyway.
        was_sent = (status == 'sent' or discord_sent)
        if was_sent and not ignore_sent:
            already_sent.append(character_name)
            flag = 'status=sent' if status == 'sent' else 'checkbox=TRUE'
            console.print(f"  [dim]SKIP {character_name} — already sent ({flag})[/dim]")
            continue

        channel_id, ch_name = channel_map.get(character_name.lower(), ('', ''))
        if not channel_id:
            no_channel.append(character_name)
            continue

        message = build_message(headers, submission_row, response_row, month, year)
        if not message:
            no_responses.append(character_name)
            continue

        if was_sent:  # only reachable when ignore_sent is set
            console.print(f"  [yellow]INCLUDE {character_name} — already sent, overridden by --ignore-sent[/yellow]")

        pending.append({
            'sheet_row':      sheet_row,
            'character_name': character_name,
            'channel_id':     channel_id,
            'ch_name':        ch_name,
            'message':        message,
            'was_sent':       was_sent,
        })

    summary_parts = []
    if already_sent:  summary_parts.append(f"{len(already_sent)} already sent")
    if no_channel:    summary_parts.append(f"{len(no_channel)} missing channel ID")
    if no_responses:  summary_parts.append(f"{len(no_responses)} no responses yet")
    if summary_parts:
        console.print(f"  [dim]{', '.join(summary_parts)}[/dim]")
    console.print(f"  [bold]{len(pending)} to send[/bold]")

    if not pending:
        console.print("\n[yellow]Nothing to send.[/yellow]")
        return

    if confirm_each:
        console.print("\n[dim]Confirm mode:  [bold]Y[/bold] send   [bold]n[/bold] skip   "
                      "[bold]a[/bold] all remaining   [bold]q[/bold] quit[/dim]\n")
    else:
        console.print("\n[dim]Controls:  [bold]P[/bold] pause/resume   [bold]S[/bold] stop early[/dim]\n")

    # Level 3: create the shared thread before the loop starts
    level3_thread_id = None
    if dry_run == 3:
        now_str     = datetime.now().strftime("%b %d %Y %H:%M")
        thread_name = f"Dry Run — {sheet.title} — {now_str}"
        console.print(f"Creating thread [bold]{thread_name}[/bold]…")
        level3_thread_id = create_discord_thread(test_channel_id, thread_name)
        if not level3_thread_id:
            console.print("[red]ERROR:[/red] Failed to create thread in test channel.")
            return
        console.print(f"  [dim]Thread ready.[/dim]\n")

    # ── Send loop ──────────────────────────────────────────────────────────────
    sent            = 0
    failed          = 0
    skipped_by_user = 0
    stopped_early   = False
    failed_names    = []
    total           = len(pending)
    main_desc       = f"[bold cyan]Sending {sheet.title}[/bold cyan]"
    pause_desc      = "[bold yellow]⏸  PAUSED — press P to resume, S to stop[/bold yellow]"

    kb = KeyboardListener()
    if not confirm_each:
        kb.start()  # raw-mode key listener would break the confirm prompt
    try:
        with Progress(
            SpinnerColumn(),
            TextColumn("{task.description}"),
            BarColumn(bar_width=None),
            MofNCompleteColumn(),
            TextColumn("•"),
            TimeElapsedColumn(),
            console=console,
            expand=True,
        ) as progress:
            main_task = progress.add_task(main_desc, total=total)

            for idx, item in enumerate(pending, start=1):
                char_name  = item['character_name']
                channel_id = item['channel_id']
                ch_name    = item['ch_name']
                message    = item['message']
                sheet_row  = item['sheet_row']
                was_sent   = item.get('was_sent', False)
                display    = ch_name or channel_id
                chunks     = chunk_message(message)
                n_chunks   = len(chunks)
                chunk_note = f" [dim]({n_chunks} chunks)[/dim]" if n_chunks > 1 else ""

                # ── Confirm-each prompt ───────────────────────────────────────
                if confirm_each:
                    progress.stop()  # pause the live display so input is clean
                    choice = confirm_send(idx, total, char_name, display, was_sent)
                    progress.start()
                    if choice == 'q':
                        stopped_early = True
                        console.log("[yellow]⏹  Stopped by user[/yellow]")
                        break
                    if choice == 'n':
                        console.log(f"[dim]⤳ skipped {char_name} (user)[/dim]")
                        skipped_by_user += 1
                        progress.advance(main_task)
                        continue
                    if choice == 'a':
                        confirm_each = False
                        console.print("[dim]Switching to automated for the rest…[/dim]")
                        kb.start()  # enable P/S controls for the remainder

                # ── Pause ─────────────────────────────────────────────────────
                if kb.paused:
                    progress.update(main_task, description=pause_desc)
                    while kb.paused and not kb.stopped:
                        time.sleep(0.1)
                    progress.update(main_task, description=main_desc)

                # ── Stop ──────────────────────────────────────────────────────
                if kb.stopped:
                    stopped_early = True
                    console.log("[yellow]⏹  Stopped — wrapping up[/yellow]")
                    break

                # ── Dry-run level 1: print to console ─────────────────────────
                if dry_run == 1:
                    console.print(Panel(
                        message,
                        title=f"[bold]{char_name}[/bold]  →  [cyan]#{display}[/cyan]",
                        subtitle=f"[dim]{n_chunks} chunk{'s' if n_chunks > 1 else ''}[/dim]",
                        border_style="dim",
                        expand=False,
                    ))
                    sent += 1
                    progress.advance(main_task)
                    continue

                # ── Live send (levels 0, 2, 3, 4) ──────────────────────────────
                # dest_note records where the message ACTUALLY went, so dry-run
                # logs make the test destination obvious instead of the player's.
                dest_note = ""
                if dry_run == 0:
                    target = channel_id
                elif dry_run == 2:
                    target = test_channel_id
                    dest_note = f" [yellow](dry: test channel {test_channel_id})[/yellow]"
                elif dry_run == 3:
                    # Post the separator header to the shared thread first
                    sep = f"{'─' * 40}\n**{char_name}**  →  #{display}"
                    requests.post(
                        f"{DISCORD_API}/channels/{level3_thread_id}/messages",
                        json={"content": sep},
                        headers=bot_headers(),
                    )
                    time.sleep(0.5)
                    target = level3_thread_id
                    dest_note = f" [yellow](dry: shared thread {level3_thread_id})[/yellow]"
                else:  # dry_run == 4
                    thread_id = create_discord_thread(test_channel_id, char_name)
                    if not thread_id:
                        console.log(f"[red]✗[/red] {char_name}  →  [dim]failed to create thread[/dim]")
                        failed += 1
                        failed_names.append(char_name)
                        progress.advance(main_task)
                        continue
                    time.sleep(0.5)  # pace thread creation
                    target = thread_id
                    dest_note = f" [yellow](dry: thread '{char_name}' {thread_id})[/yellow]"
                url = f"{DISCORD_API}/channels/{target}/messages"

                chunk_task = None
                if n_chunks > 1:
                    chunk_task = progress.add_task(
                        f"  [dim]chunk 1/{n_chunks}[/dim]", total=n_chunks
                    )

                ok           = True
                error_detail = None

                for idx, chunk in enumerate(chunks):
                    if chunk_task is not None:
                        progress.update(chunk_task,
                            description=f"  [dim]chunk {idx + 1}/{n_chunks}[/dim]",
                            completed=idx,
                        )

                    for _attempt in range(3):
                        resp = requests.post(url, json={"content": chunk}, headers=bot_headers())
                        if resp.status_code in (200, 201):
                            break
                        elif resp.status_code == 429:
                            wait     = resp.json().get('retry_after', 5)
                            wait_end = time.monotonic() + wait + 0.5
                            while time.monotonic() < wait_end and not kb.stopped:
                                remaining = max(0.0, wait_end - time.monotonic())
                                progress.update(main_task,
                                    description=f"[yellow]⏳ rate limited — {remaining:.0f}s remaining[/yellow]")
                                time.sleep(min(0.5, remaining + 0.01))
                            progress.update(main_task, description=main_desc)
                            if kb.stopped:
                                ok           = False
                                error_detail = "stopped during rate-limit wait — not retried"
                                break
                        else:
                            ok           = False
                            error_detail = f"HTTP {resp.status_code}"
                            break
                    else:
                        # Loop exhausted without a break — all attempts were 429s
                        ok           = False
                        error_detail = "max retries exceeded (rate limited)"

                    if not ok:
                        break

                    if chunk_task is not None:
                        progress.update(chunk_task, completed=idx + 1)

                    if idx < n_chunks - 1:
                        time.sleep(1)

                if chunk_task is not None:
                    progress.remove_task(chunk_task)

                if ok:
                    console.log(f"[green]✓[/green] {char_name}  →  [bold]#{display}[/bold]{chunk_note}{dest_note}")
                    sent += 1
                    if dry_run == 0:
                        now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                        # One batched write for all three cells (avoids quota churn)
                        sheets_retry(
                            sheet.batch_update,
                            [
                                {'range': rowcol_to_a1(sheet_row, TIMESTAMP_COL),    'values': [[now]]},
                                {'range': rowcol_to_a1(sheet_row, STATUS_COL),       'values': [['sent']]},
                                {'range': rowcol_to_a1(sheet_row, SEND_DISCORD_COL), 'values': [['TRUE']]},
                            ],
                            # USER_ENTERED so 'TRUE' ticks the checkbox (matches old update_cell)
                            value_input_option='USER_ENTERED',
                            notify=lambda msg: progress.console.print(f"  [yellow]{msg}[/yellow]"),
                        )
                        time.sleep(0.5)
                else:
                    console.log(f"[red]✗[/red] {char_name}  →  [bold]#{display}[/bold]  [dim]({error_detail})[/dim]")
                    failed += 1
                    failed_names.append(char_name)

                progress.advance(main_task)

    finally:
        kb.stop()

    # ── Summary table ──────────────────────────────────────────────────────────
    console.print()
    dry_tag     = " [yellow](dry run)[/yellow]" if dry_run else ""
    stopped_tag = " [yellow](stopped early)[/yellow]" if stopped_early else ""
    table = Table(
        title=f"{sheet.title}{dry_tag}{stopped_tag}",
        show_header=False,
        box=box.ROUNDED,
        min_width=40,
    )
    table.add_column("", style="bold", no_wrap=True)
    table.add_column("", justify="right")

    if sent:
        label = "Would send" if dry_run else "Sent"
        table.add_row(f"[green]{label}[/green]", str(sent))
    if failed:
        table.add_row("[red]Failed[/red]", str(failed))
    if skipped_by_user:
        table.add_row("[yellow]Skipped (user)[/yellow]", str(skipped_by_user))
    if stopped_early:
        remaining = total - sent - failed - skipped_by_user
        table.add_row("[yellow]Not reached[/yellow]", str(remaining))
    if already_sent:
        table.add_row("[dim]Already sent[/dim]", str(len(already_sent)))
    if no_channel:
        table.add_row("[yellow]No channel ID[/yellow]", str(len(no_channel)))
    if no_responses:
        table.add_row("[dim]No responses yet[/dim]", str(len(no_responses)))

    console.print(table)

    if failed_names:
        console.print(f"\n[red]Failed:[/red] {', '.join(failed_names)}")
    if no_channel:
        console.print(f"[yellow]Missing channel IDs:[/yellow] {', '.join(no_channel)}")


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="LOTS LARP Downtime Discord Sender")
    parser.add_argument('sheet', nargs='?',
                        help='Sheet name e.g. "June 2026" — skips the interactive picker')
    parser.add_argument('--dry-run', action='store_true',
                        help='Interactively choose a dry-run level instead of sending live')
    parser.add_argument('--ignore-sent', action='store_true',
                        help='Include rows already marked sent (overrides the skip)')
    args = parser.parse_args()

    console.print("Connecting to Google Sheets…")
    client      = sheets_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    sheet_name = args.sheet if args.sheet else pick_sheet(spreadsheet)
    console.print(f"Selected: [bold]{sheet_name}[/bold]")

    dry_run      = pick_dry_run_level() if args.dry_run else 0
    confirm_each = pick_send_mode()
    console.print()

    send_downtimes(spreadsheet, sheet_name, dry_run=dry_run,
                   confirm_each=confirm_each, ignore_sent=args.ignore_sent)


if __name__ == "__main__":
    main()
