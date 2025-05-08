/**
 * @OnlyCurrentDoc
 *
 * Functions for handling scheduled Discord messages.
 */

// --- Constants for this feature ---
// Note: These constants are defined in Constants.gs
// const SCHEDULED_MSG_SHEET_NAME = 'Scheduled Messages';
// const SCHEDULED_MSG_SENT_COL = 1;
// const SCHEDULED_MSG_TIME_COL = 2;
// const SCHEDULED_MSG_CHANNEL_COL = 3;
// const SCHEDULED_MSG_SENDER_COL = 4;
// const SCHEDULED_MSG_BODY_COL = 5;
// const SCHEDULED_MSG_LOG_COL = 6;
// const NPC_SHEET_NAME = 'NPC'; // Defined in Constants.gs
// const CHAR_SHEET_NAME_COL = 1; // Defined in Constants.gs
// const CHAR_SHEET_AVATAR_COL = 25; // Defined in Constants.gs

/**
 * Function intended to be run by a time-driven trigger (e.g., hourly).
 * Reads the SCHEDULED_MSG_SHEET_NAME sheet and sends pending messages
 * to the selected Discord channel webhook, optionally using a character/NPC sender override
 * by looking up the sender name in the NPC_SHEET_NAME.
 * Skips rows without a valid 'Send After Timestamp' or invalid channel selection.
 */
function sendScheduledMessages_() {
    const functionName = 'sendScheduledMessages_';
    Logger.log(`Starting ${functionName}...`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SCHEDULED_MSG_SHEET_NAME); // Constant from Constants.gs

    if (!sheet) {
        Logger.log(`${functionName}: Sheet "${SCHEDULED_MSG_SHEET_NAME}" not found. Aborting.`);
        // Optionally notify admin/STs
        // sendDiscordWebhookMessage_(PropertiesService.getScriptProperties().getProperty(PROP_ST_WEBHOOK), `🚨 ERROR: Scheduled Messages sheet ("${SCHEDULED_MSG_SHEET_NAME}") is missing!`, 'Scheduled Message Error');
        return;
    }

    // Get all webhook URLs at once
    const scriptProperties = PropertiesService.getScriptProperties();
    const webhooks = {
        '#announcements': scriptProperties.getProperty(PROP_ANNOUNCEMENT_WEBHOOK), // In Constants.gs
        '#ic-chat': scriptProperties.getProperty(PROP_IC_CHAT_WEBHOOK),             // In Constants.gs
        '#ic-news-feed': scriptProperties.getProperty(PROP_IC_NEWS_FEED_WEBHOOK)     // In Constants.gs
    };

    // Check if at least one webhook is configured (optional, but good practice)
    const configuredWebhooks = Object.values(webhooks).filter(url => url && !url.includes('YOUR_'));
    if (configuredWebhooks.length === 0) {
         Logger.log(`${functionName}: No valid Discord Webhook URLs configured in Script Properties for announcements, ic-chat, or ic-news-feed. Aborting.`);
         return;
    }

    const dataRange = sheet.getDataRange();
    // Check if there's more than just the header row
    if (dataRange.getNumRows() <= 1) {
        Logger.log(`${functionName}: No messages found in "${SCHEDULED_MSG_SHEET_NAME}".`);
        return;
    }

    const data = dataRange.getValues();
    const now = new Date(); // Current time
    let npcDataCache = null; // *** Cache NPC data ***

    Logger.log(`${functionName}: Checking ${data.length - 1} potential messages.`);

    // Iterate through rows, skipping header (index 0)
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const currentRowNum = i + 1; // Sheet row number (1-based)

        // Get data using constants defined in Constants.gs
        const sendAfterTimestamp = row[SCHEDULED_MSG_TIME_COL - 1];
        const isSent = row[SCHEDULED_MSG_SENT_COL - 1];
        const targetChannel = String(row[SCHEDULED_MSG_CHANNEL_COL - 1] || '').trim();
        const senderName = String(row[SCHEDULED_MSG_SENDER_COL - 1] || '').trim(); // Get Sender Name
        const messageBody = row[SCHEDULED_MSG_BODY_COL - 1];
        const currentLog = row[SCHEDULED_MSG_LOG_COL - 1];

        // --- Validation ---
        // 1. Check for valid timestamp FIRST
        if (!(sendAfterTimestamp instanceof Date) || isNaN(sendAfterTimestamp.getTime())) {
             // Logger.log(`Row ${currentRowNum}: Skipping, invalid or empty date in column ${SCHEDULED_MSG_TIME_COL}. Value: ${sendAfterTimestamp}`);
             continue;
        }

        // 2. Check if already sent
        if (isSent === true) {
            continue;
        }

        // 3. Check for message body
        if (!messageBody || String(messageBody).trim() === '') {
             Logger.log(`Row ${currentRowNum}: Skipping, message body is empty.`);
             continue;
        }

        // 4. Check for valid channel and get webhook
        let targetWebhookUrl = null;
        if (targetChannel === '#ic-chat') {
            targetWebhookUrl = webhooks['#ic-chat'];
        } else if (targetChannel === '#ic-news-feed') {
            targetWebhookUrl = webhooks['#ic-news-feed'];
        } else if (targetChannel === '#announcements') {
            targetWebhookUrl = webhooks['#announcements'];
        } else {
            Logger.log(`Row ${currentRowNum}: Skipping, invalid or empty Target Channel selected: "${targetChannel}".`);
            sheet.getRange(currentRowNum, SCHEDULED_MSG_LOG_COL).setValue(`Error: Invalid channel "${targetChannel}" (${new Date().toLocaleString()})`);
            continue;
        }

        if (!targetWebhookUrl || targetWebhookUrl.includes('YOUR_')) {
             Logger.log(`Row ${currentRowNum}: Skipping, Webhook URL for channel "${targetChannel}" is not configured in Script Properties.`);
             sheet.getRange(currentRowNum, SCHEDULED_MSG_LOG_COL).setValue(`Error: Webhook not configured for ${targetChannel} (${new Date().toLocaleString()})`);
             continue;
        }

        // Check if it's time to send
        if (now >= sendAfterTimestamp) {
            Logger.log(`Row ${currentRowNum}: Timestamp reached (${sendAfterTimestamp.toLocaleString()}). Attempting to send message to ${targetChannel}.`);

            // *** Look up Sender Info if specified (from NPCs sheet) ***
            let senderUsername = null;
            let senderAvatarUrl = null;
            if (senderName) {
                Logger.log(`Row ${currentRowNum}: Sender specified: "${senderName}". Looking up info in NPCs sheet...`);
                if (!npcDataCache) { // Lazy load NPC data
                     Logger.log('Loading NPC data cache...');
                     // *** Use NPC_SHEET_NAME ***
                     const npcSheet = ss.getSheetByName(NPC_SHEET_NAME); // Constant from Constants.gs
                     if (npcSheet) {
                         npcDataCache = npcSheet.getDataRange().getValues();
                     } else {
                         Logger.log(`Warning: NPC sheet "${NPC_SHEET_NAME}" not found. Cannot look up sender info.`);
                     }
                }

                if (npcDataCache) {
                    // Find NPC row (case-insensitive)
                    const lowerSenderName = senderName.toLowerCase();
                    Logger.log(`Row ${currentRowNum}: Searching for lowercase name: "${lowerSenderName}" in NPCs sheet.`);
                    for (let j = 1; j < npcDataCache.length; j++) { // Start from 1 to skip header
                        const npcRow = npcDataCache[j];
                        // *** Use CHAR_SHEET constants assuming structure is same ***
                        const nameInSheet = String(npcRow[CHAR_SHEET_NAME_COL - 1] || '').trim();
                        const lowerNameInSheet = nameInSheet.toLowerCase();

                        // Log first/last few comparisons for brevity
                        if (j < 10 || j > npcDataCache.length - 10) {
                          Logger.log(`Row ${currentRowNum}: Comparing with NPC Sheet Row ${j+1}: "${lowerNameInSheet}"`);
                        }

                        if (lowerNameInSheet === lowerSenderName) { // Case-insensitive comparison
                            senderUsername = nameInSheet; // Use the exact name from sheet
                            senderAvatarUrl = String(npcRow[NPC_SHEET_AVATAR_COL_SHEET_AVATAR_COL - 1] || '').trim(); // Constant from Constants.gs
                            Logger.log(`Row ${currentRowNum}: MATCH FOUND at NPCs sheet row ${j+1}. Name: "${senderUsername}", Avatar Col Y: "${senderAvatarUrl}"`);
                            if (!senderAvatarUrl) {
                                Logger.log(`Row ${currentRowNum}: Avatar URL in column Y is empty for NPC "${senderUsername}".`);
                                senderAvatarUrl = null; // Ensure it's null if empty
                            }
                            break; // Stop searching once found
                        }
                    }
                    if (!senderUsername) {
                         Logger.log(`Row ${currentRowNum}: Sender "${senderName}" not found in NPCs sheet after checking ${npcDataCache.length - 1} rows. Sending with default appearance.`);
                    }
                }
            } else {
                 Logger.log(`Row ${currentRowNum}: No sender specified. Sending with default appearance.`);
            }
            // --- End Sender Info Lookup ---

            logAudit_('Scheduled Message Send Attempt', sheet.getName(), `Row: ${currentRowNum}, Target Channel: ${targetChannel}, Sender: ${senderUsername || 'Default'}, Target Time: ${sendAfterTimestamp.toLocaleString()}`); // In Utilities.gs

            let success = false;
            let errorMessage = '';
            try {
                // *** Pass sender info to the webhook function ***
                success = sendDiscordWebhookMessage_(
                    targetWebhookUrl,
                    String(messageBody),
                    `Scheduled Message to ${targetChannel}`,
                    senderUsername, // Pass username (or null)
                    senderAvatarUrl // Pass avatar URL (or null)
                 ); // In Discord.gs
            } catch (error) {
                Logger.log(`Error during sendDiscordWebhookMessage_ for row ${currentRowNum}: ${error}`);
                errorMessage = `Error: ${error.message}`;
                success = false; // Ensure success is false on exception
            }

            // Update log and checkbox based on success/failure
            const logTimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
            if (success) {
                sheet.getRange(currentRowNum, SCHEDULED_MSG_SENT_COL).setValue(true); // Check the box
                sheet.getRange(currentRowNum, SCHEDULED_MSG_LOG_COL).setValue(`${logTimestamp} - Sent successfully to ${targetChannel} ${senderUsername ? 'as ' + senderUsername : ''}.`);
                logAudit_('Scheduled Message Sent', sheet.getName(), `Row: ${currentRowNum}, Channel: ${targetChannel}, Sender: ${senderUsername || 'Default'}`); // In Utilities.gs
                Logger.log(`Row ${currentRowNum}: Message sent successfully to ${targetChannel}.`);
            } else {
                sheet.getRange(currentRowNum, SCHEDULED_MSG_LOG_COL).setValue(`${logTimestamp} - SEND FAILED to ${targetChannel}. ${errorMessage}`.trim());
                logAudit_('Scheduled Message Send FAILED', sheet.getName(), `Row: ${currentRowNum}, Channel: ${targetChannel}, Sender: ${senderUsername || 'Default'}, Error: ${errorMessage || 'Check execution logs'}`); // In Utilities.gs
                Logger.log(`Row ${currentRowNum}: Message send FAILED to ${targetChannel}. Error: ${errorMessage || 'Check execution logs'}`);
            }
             // Add a small delay between sending messages if processing many rows
             Utilities.sleep(500); // 0.5 second delay
        } else {
             // Logger.log(`Row ${currentRowNum}: Skipping, send time (${sendAfterTimestamp.toLocaleString()}) not yet reached.`);
        }
    } // End loop through rows

    Logger.log(`Finished ${functionName}.`);
}
