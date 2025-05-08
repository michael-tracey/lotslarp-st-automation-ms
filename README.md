# LARP Downtime Management Script - User Guide

## 1. Introduction

This script is designed to help Storytellers manage and process Live Action Role-Playing (LARP) game downtimes and influence actions submitted by players. It integrates with Google Sheets and Discord to streamline data entry, tracking, reporting, and communication.

Key features include:
* Automated processing of Google Form submissions for downtimes.
* A custom menu in Google Sheets for various Storyteller actions.
* Dialogs for editing downtime responses and viewing progress.
* Automated and manual sending of downtime results and reports to Discord.
* Character name validation with visual feedback.
* Functions to fill cells with randomized data for common actions.
* A system for sending scheduled announcements to Discord.
* An audit log to track script actions.

## 2. Setup & Configuration

### 2.1. Initial Script Properties Setup

Before using the script, several properties must be configured in the Apps Script editor:

1.  Open your Google Sheet.
2.  Go to **Extensions > Apps Script**.
3.  In the Apps Script editor, click on **Project Settings** (the gear icon ⚙️ on the left).
4.  Scroll down to the **Script Properties** section.
5.  Click **Add script property**.
6.  Add or verify the following properties (replace placeholder URLs with your actual webhook URLs):

    * `PROP_ST_WEBHOOK`: Your main Storyteller Discord webhook URL (e.g., for reports).
        * Example: `https://discord.com/api/webhooks/.../...`
    * `PROP_ANNOUNCEMENT_WEBHOOK`: Webhook URL for the Discord channel where general scheduled announcements should go.
        * Example: `https://discord.com/api/webhooks/.../...`
    * `PROP_IC_CHAT_WEBHOOK`: Webhook URL for the Discord channel intended for IC (In-Character) chat messages from scheduled announcements.
        * Example: `https://discord.com/api/webhooks/.../...`
    * `PROP_IC_NEWS_FEED_WEBHOOK`: Webhook URL for the Discord channel intended for IC news feed style messages from scheduled announcements.
        * Example: `https://discord.com/api/webhooks/.../...`
    * `PROP_DOWNTIME_FORM_ID`: The ID of the Google Form used for downtime submissions.
        * To find this: Open your Google Form for editing. The ID is the long string of characters in the URL between `/d/` and `/edit`.
        * Example: `1B7zV_uvkIdYroWkv64ljoI80Ubplgm0eEMAboFbmJDk`
    * `PROP_DOWNTIME_YEAR`: The current game year (e.g., `2025`). This is used for naming downtime sheets and in report titles.
    * `PROP_DOWNTIME_MONTH`: The current game month (e.g., `April`). This is used for naming downtime sheets and in report titles.
    * `PROP_PDF_FOLDER_ID`: (Currently a placeholder, but intended for PDF export features if added later) The ID of the Google Drive folder where PDFs might be saved.
        * Example: `YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE`
    * `PROP_DISCORD_TEST_MODE`: Set to `false` for normal operation. Can be toggled via the menu.
    * `PROP_TEST_WEBHOOK`: A Discord webhook URL for a private test channel. Used when "Discord Test Mode" is active.
        * Example: `https://discord.com/api/webhooks/.../...`

    **Important:** If you see alerts about unconfigured properties when opening the sheet, please ensure these are correctly set.

### 2.2. Sheet Naming Conventions

The script relies on specific sheet names for its operations:

* **Downtime Sheets:** Named in the format "Month Year" (e.g., `April 2025`). The script uses the `PROP_DOWNTIME_MONTH` and `PROP_DOWNTIME_YEAR` properties to determine the current active downtime sheet.
* **`Characters`:** Stores player character information, including names (Column A), individual Discord webhooks (Column X), and avatar URLs (Column Y).
* **`NPCs`:** Stores Non-Player Character information, including names (Column A) and avatar URLs (Column Y) for the scheduled message sender override.
* **`Influences`:** Lists character specializations.
    * Column A: Elite Character Name
    * Column B: Elite Specialization Name
    * Column D: UW Character Name
    * Column E: UW Specialization Name
* **`Elite Infl.`:** Details for Elite influence actions. Expected headers: `Date` (A), `Age` (B), `Character` (C), `Specialization` (D), `Action` (E), `Details` (F), `Any blocks?` (G), `Output` (H).
* **`UW Infl.`:** Details for Underworld influence actions. Expected headers: `Date` (A), `Age` (B), `Character` (C), `Specialization` (D), `Action` (E), `Details` (F), `Any blocks?` (G), `Output` (H).
* **`Scheduled Messages`:** Used for the scheduled announcement feature.
    * Column A: `Sent` (Checkbox)
    * Column B: `Send After Timestamp` (Date time format)
    * Column C: `Target Channel` (Dropdown: `#announcements`, `#ic-chat`, `#ic-news-feed`)
    * Column D: `Sender (Optional)` (Character/NPC name from `NPCs` sheet)
    * Column E: `Message Body`
    * Column F: `Send Log`
* **`Log`:** An audit log sheet automatically created by the script to record actions.

## 3. Main Menu: "Storyteller Menu"

This custom menu appears in your Google Sheet when the script is active.

### 3.1. Dialogs Group

* **`Downtime Editor popup`**:
    * **Usage:** Select a cell in a downtime sheet (response row, not header or protected columns) that you want to edit. Click this menu item.
    * **Function:** Opens a dialog box showing the original player submission (prompt) and the current response. You can edit the response text. Markdown formatting buttons are provided for Discord compatibility.
    * **Note:** Saves changes directly to the sheet and logs the edit.

* **`Show Downtime Progress`**:
    * **Usage:** Click this when on an active downtime sheet ("Month Year").
    * **Function:** Displays a dialog summarizing the completion status of downtime responses, including overall percentage, character submission count, word count statistics, and a breakdown by keywords found in submissions.

* **`Show Missing Downtime Responses`**:
    * **Usage:** Click this when on an active downtime sheet.
    * **Function:** Shows a dialog listing all downtime actions that have been submitted by players but do not yet have a corresponding ST response filled in. Includes character name, the action header, and the cell needing a response.

* **`Show Influences Progress`**:
    * **Usage:** Click this when on an active downtime sheet.
    * **Function:** Displays a summary of influence actions (identified by headers containing "influence") from the *current downtime sheet*. It shows submitted vs. completed counts for categories like "elite" and "underworld" based on header keywords.

* **`Show Resources Progress`**:
    * **Usage:** Click this when on an active downtime sheet.
    * **Function:** Similar to "Influences Progress," but for resource actions (identified by headers containing "resources") from the *current downtime sheet*.

* **`Show Detailed Influence Summary`**:
    * **Usage:** Click this menu item.
    * **Function:** Opens a dialog that summarizes influence actions from the dedicated `Elite Infl.` and `UW Infl.` sheets.
        * Displays a table of total Elite and Underworld actions per month, including percentage change from the previous month.
        * For each month (sorted most recent first), it groups actions by Level (the number prefix in the action name, e.g., "1: Action Name").
        * For each Level, it shows Elite and Underworld action names side-by-side, with a list of specializations that used that action and their counts.
        * Also shows the percentage of total Elite/Underworld actions for that month that each specific action represents.
        * Includes "Most Used," "Least Used (Used > 0)," and separate "Unused Elite/Underworld Specializations" lists for each month.

### 3.2. Fill Cell Group

These functions help populate cells with randomized or calculated data. Generally, you select the cell you want to fill and then click the menu item.

* **`Fill cell with Feed Data`**: Fills the active cell with random data from the `feed list` sheet, based on the "Feed" column and replacing placeholders.
* **`Fill cell with Herd Data`**: Similar to Feed Data, but uses the "Herd" column.
* **`Fill cell with Patrol Data`**: Similar to Feed Data, but uses the "Patrol" column.
* **`Fill Cell with Elite Influences`**:
    * **Usage:** Select a cell in a response row (odd-numbered row) where you want the Elite influence result to appear.
    * **Function:**
        1.  Determines the character name from Column E of the selected row.
        2.  Looks up the character's Elite and Underworld specializations from the `Influences` sheet.
        3.  Opens a dialog displaying the character's name, their Elite/UW specs, and total points in each.
        4.  Prompts you to select an "Action Power Level" (1-10) from a dropdown.
        5.  Provides two buttons: "Gossip and Insider Trading" (for Elite) and "Word on the Street" (for Underworld). These are greyed out if the character has no points in that category.
        6.  When a button is clicked, it filters the `Elite Infl.` or `UW Infl.` sheet based on:
            * Character's specializations for that type (Elite/UW).
            * Age <= 95 (Column B).
            * Blocks <= Selected Power Level + 1 (Column G).
            * Output (Column H) must contain a colon (`:`) and must NOT start with the number `2`.
        7.  The matching output(s) are joined by newlines and placed in the selected cell.
        8.  A report dialog shows what was filled and what was skipped (and why).
* **`Fill Cell with Underworld Influences`**: Works identically to "Fill Cell with Elite Influences" but targets Underworld specializations and uses the `UW Infl.` sheet.

### 3.3. Discord Actions Group

* **`Send Downtime Report to Discord`**:
    * **Usage:** Click when on an active downtime sheet.
    * **Function:** Generates a summary report (similar to "Show Downtime Progress") and sends it to the Discord webhook defined in `PROP_ST_WEBHOOK`. Respects Test Mode.

* **`Send Missing Downtime Responses to Discord`**:
    * **Usage:** Click when on an active downtime sheet.
    * **Function:** Generates a list of missing responses (similar to "Show Missing Downtime Responses") and sends it to the `PROP_ST_WEBHOOK`. Respects Test Mode.

* **`Test Discord to ST Channel`**:
    * **Usage:** Click to send a test message.
    * **Function:** Sends a predefined test message to the `PROP_ST_WEBHOOK`. This function *also respects the Test Mode toggle*. If Test Mode is ON and `PROP_TEST_WEBHOOK` is configured, the message goes there instead.

### 3.4. Maintenance Submenu

* **`Reinstall Form Trigger`**: Deletes and recreates the trigger that automatically processes new Google Form submissions. Use this if form submissions stop being added to the sheet.
* **`Reinstall Edit Trigger`**: Deletes and recreates the `onEdit` trigger that handles checkbox actions and name validation. Use this if those features stop working.
* **`Setup Scheduled Message Trigger`**: Deletes any existing triggers for `sendScheduledMessages_` and creates a new time-driven trigger to run that function hourly. Use this to enable or reset the scheduled announcement feature.
* **`Manually Test Usernames`**:
    * **Usage:** Click when on an active downtime sheet.
    * **Function:** Iterates through all response rows (odd-numbered rows) on the current sheet. For each character name in Column E:
        * Validates the name against the `Characters` sheet.
        * Checks for a corresponding Discord webhook in the `Characters` sheet.
        * Applies background color to the name cell:
            * **Orange:** Name not found in `Characters` sheet.
            * **Red:** Name found, but no Discord webhook URL.
            * **Green:** Name found, and webhook URL is present.
        * Shows an alert summarizing any issues found.

* **`Check Character Count`**: Reads the `Characters` sheet, counts how many characters have "approved" in their approval status column (Column H), and sends this count to the `PROP_ST_WEBHOOK`. Respects Test Mode.

* **`Manually Send Discord for Selected Row`**:
    * **Usage:** Select any cell in a *response row* (odd-numbered row) for which you want to send Discord results.
    * **Function:** Triggers the `handleSendDiscord_` function for that row. It will prompt for confirmation, especially if on an older sheet. It respects the "sent" status (will not resend if already marked as sent).

* **`Manually Send Email for Selected Row`**:
    * **Usage:** Select any cell in a *response row* (odd-numbered row) for which you want to send email results.
    * **Function:** Triggers the `handleSendEmail_` function for that row. It will prompt for an email address and confirmation (especially if on an older sheet). This manual version *ignores* the "sent" status in the status column.

* **`Fetch Downtimes (Placeholder)`**: Currently a placeholder, does nothing.

### 3.5. Toggle Discord Test Mode

* **Usage:** Click to toggle the Discord test mode ON or OFF. The menu item will show a checkmark (✅) when Test Mode is ON.
* **Function:**
    * When **ON**: All Discord messages sent by the script (player downtime results, ST reports, scheduled messages) will be redirected to the webhook URL specified in the `PROP_TEST_WEBHOOK` script property. Messages will be prefixed with `🧪 **[TEST MODE]** 🧪`.
    * When **OFF**: Messages go to their intended webhooks (`PROP_ST_WEBHOOK`, player-specific webhooks, or selected announcement webhooks).
    * **Important:** Ensure `PROP_TEST_WEBHOOK` is correctly configured in Script Properties if you plan to use Test Mode.

## 4. Automated Features

### 4.1. Form Submission Processing (`onFormSubmitHandler_`)

* When a player submits the linked Google Form:
    1.  An email notification is sent to the ST email address (`avllotslarp@gmail.com`).
    2.  A new "submission row" (even-numbered) is added to the current downtime sheet with the raw player input.
    3.  A new "response row" (odd-numbered, below the submission) is added as a template for STs.
    4.  Checkboxes for "Send Discord" and "Send Email" are inserted into the response row.
    5.  The character name from the form is placed in Column E of the response row.
    6.  The character name in Column E of the response row is automatically validated (see "Manually Test Usernames" for color logic).
    7.  The action is logged in the `Log` sheet.

### 4.2. Checkbox & Edit Triggers (`onEditHandler_`)

This function runs automatically whenever a cell is edited on a sheet named "Month Year".

* **"Send Discord" Checkbox (Column C, Odd Rows):**
    * If checked AND the status in Column B of that row is NOT "sent":
        * It will first check if the current sheet is an older month's sheet. If so, it will ask for "Yes/No" confirmation before proceeding (unless it's within the first 7 days of the next month, which is a grace period).
        * If confirmed (or not an old sheet/in grace period), it calls `handleSendDiscord_` to construct and send the message to the character's webhook (or test webhook if Test Mode is ON).
        * On success, it updates the status to "sent", adds a timestamp, and colors the checkbox cell green.
        * On failure (after retries), it alerts the user and unchecks the box.
    * If checked AND the status IS "sent": An alert pops up explaining that it cannot resend, and the box is unchecked.

* **"Send Email" Checkbox (Column D, Odd Rows):**
    * If checked:
        * It will first check if the current sheet is an older month's sheet. If so, it will ask for "Yes/No" confirmation before proceeding (unless it's within the first 7 days of the next month).
        * If confirmed (or not an old sheet/in grace period), it calls `handleSendEmail_` which prompts for the recipient's email address, then sends the formatted results.
        * On success, it updates the status to "sent", adds a timestamp, and colors the checkbox cell green.
        * On failure, it alerts the user and unchecks the box.
    * **Note:** This trigger *ignores* the status in Column B.

* **Character Name Edit (Column E, Odd Rows):**
    * If a name in Column E of an odd-numbered (response) row is edited, `validateCharacterNameCell_` is called to re-validate the name and update its background color.

* **Status Change (Column B, Odd Rows):**
    * If the value in the Status column (B) of an odd-numbered (response) row is manually changed, the change (old and new value) is logged in the `Log` sheet.

### 4.3. Scheduled Messages (`sendScheduledMessages_`)

* This function is designed to be run by a time-driven trigger (typically hourly, set up via the "Setup Scheduled Message Trigger" menu item).
* It reads the `Scheduled Messages` sheet.
* For each row:
    1.  It skips the row if the "Sent" checkbox (Column A) is checked.
    2.  It skips the row if the "Send After Timestamp" (Column B) is empty or invalid.
    3.  It skips the row if the "Message Body" (Column E) is empty.
    4.  It reads the "Target Channel" (Column C) and attempts to find a corresponding webhook URL from Script Properties (`PROP_ANNOUNCEMENT_WEBHOOK`, `PROP_IC_CHAT_WEBHOOK`, `PROP_IC_NEWS_FEED_WEBHOOK`). If the channel is invalid or its webhook isn't configured, it logs an error in the "Send Log" (Column F) and skips.
    5.  If the current time is at or after the "Send After Timestamp":
        * It checks the "Sender (Optional)" column (D). If a name is provided:
            * It looks up this name in the `NPCs` sheet (Column A).
            * If found, it uses the NPC's name and their Avatar URL (from `NPCs` sheet, Column Y) to override the webhook's default appearance. If no avatar URL is found for the NPC, it sends with the NPC name but the webhook's default avatar.
            * If the name isn't found, it sends with the webhook's default appearance.
        * It sends the "Message Body" to the selected channel's webhook (respecting Test Mode).
        * On success, it checks the "Sent" checkbox and updates the "Send Log".
        * On failure, it updates the "Send Log" with an error.
        * All send attempts and outcomes are logged to the `Log` sheet.

## 5. Data Sheets Overview

* **`Characters` Sheet:**
    * **Column A (CHAR\_SHEET\_NAME\_COL):** Character Name (used for validation and lookups).
    * **Column H (CHAR\_SHEET\_APPROVAL\_COL):** Approval Status (e.g., "approved", used by "Check Character Count").
    * **Column X (CHAR\_SHEET\_WEBHOOK\_COL):** Individual Discord Webhook URL for this character (used for sending downtime results).
    * **Column Y (CHAR\_SHEET\_AVATAR\_COL):** URL to an image for the character's avatar (used by scheduled messages if this character is a sender).

* **`NPCs` Sheet:**
    * **Column A (CHAR\_SHEET\_NAME\_COL):** NPC Name.
    * **Column Y (CHAR\_SHEET\_AVATAR\_COL):** URL to an image for the NPC's avatar (used by scheduled messages if this NPC is a sender).

* **`Influences` Sheet:**
    * **Column A:** Elite Character Name
    * **Column B:** Elite Specialization Name
    * **Column D:** UW Character Name
    * **Column E:** UW Specialization Name
    * (This sheet is used by the "Fill Cell with Elite/Underworld Influences" functions to determine which specializations a character possesses).

* **`Elite Infl.` / `UW Infl.` Sheets:**
    * These sheets store the predefined influence actions available.
    * **Column A (INFL\_DATE\_COL\_IDX\_):** Date (used for monthly summary).
    * **Column B:** Age (used for filtering, e.g., `<= 95`).
    * **Column D (INFL\_SPEC\_COL\_IDX\_):** Specialization name (must match one of the character's specs from the `Influences` sheet for the action to be considered).
    * **Column E (INFL\_ACTION\_COL\_IDX\_):** Action name (e.g., "1: Free Travel", "10: Regional Influence"). The number prefix is used for sorting.
    * **Column G:** Blocks (numeric value, compared against `Power Level + 1`).
    * **Column H:** Output (the text to be filled into the cell; filtered to require a colon `:` and not start with `2`).

* **`Scheduled Messages` Sheet:**
    * **Column A (SCHEDULED\_MSG\_SENT\_COL):** Checkbox, marked `TRUE` by script after sending.
    * **Column B (SCHEDULED\_MSG\_TIME\_COL):** Date and Time after which the message should be sent.
    * **Column C (SCHEDULED\_MSG\_CHANNEL\_COL):** Dropdown to select target channel (`#announcements`, `#ic-chat`, `#ic-news-feed`).
    * **Column D (SCHEDULED\_MSG\_SENDER_COL):** Optional. Name of an NPC from the `NPCs` sheet to send the message as.
    * **Column E (SCHEDULED\_MSG\_BODY\_COL):** The content of the message (supports Discord Markdown).
    * **Column F (SCHEDULED\_MSG\_LOG\_COL):** Log of send attempts/results for that row.

* **`Log` Sheet:**
    * Automatically created if it doesn't exist.
    * Records various script actions with a timestamp, user, action type, sheet name, and details. Useful for troubleshooting and auditing.

## 6. Troubleshooting Tips

* **Function Not Found / Menu Item Error:**
    * Ensure all script files (`Main.gs`, `Constants.gs`, `Triggers.gs`, etc.) are present in your Apps Script project and have been saved.
    * Verify that function names in menu item definitions exactly match the function names in the script files.
* **Checkboxes Not Working:**
    * Ensure the `onEdit` trigger is installed and enabled (see Maintenance menu).
    * Check the Apps Script execution logs (Extensions > Apps Script > Executions) for errors after clicking a checkbox.
    * Make sure you are on a sheet named "Month Year" and are editing a response row (odd-numbered).
* **Form Submissions Not Processing:**
    * Verify `PROP_DOWNTIME_FORM_ID` in Script Properties is correct.
    * Ensure the Form Submit trigger is installed (see Maintenance menu).
    * Check Execution Logs for errors related to `onFormSubmitHandler_`.
* **Discord Messages Not Sending:**
    * Verify all relevant webhook URLs (`PROP_ST_WEBHOOK`, `PROP_ANNOUNCEMENT_WEBHOOK`, etc., and individual character webhooks) are correctly set in Script Properties and on the `Characters` sheet.
    * Check if Discord Test Mode is ON and if `PROP_TEST_WEBHOOK` is set.
    * Review the `sendDiscordWebhookMessage_` function logs for detailed error messages from Discord (e.g., rate limits, invalid webhook).
* **Incorrect Data/Summaries:**
    * Ensure your data sheets (`Characters`, `NPCs`, `Influences`, `Elite Infl.`, `UW Infl.`) are correctly formatted with the expected headers and data types in the correct columns.
    * Check the `PROP_DOWNTIME_MONTH` and `PROP_DOWNTIME_YEAR` script properties if summaries seem to be for the wrong period.

Always check the **Execution Logs** in the Apps Script editor first when encountering issues. They often provide specific error messages that can help pinpoint the problem.
