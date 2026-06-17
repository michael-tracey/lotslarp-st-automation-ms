/**
 * @OnlyCurrentDoc
 *
 * Functions related to Discord communication (sending messages, reports).
 */

/**
 * Shows a non-blocking toast on the spreadsheet and writes to the script cache
 * so bulk-send dialogs can surface retry activity in real time.
 */
function notifyDiscordRetry_(message, durationSeconds = 20) {
    try {
        SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Discord Retry', durationSeconds);
    } catch (e) {}
    try {
        const cache = CacheService.getScriptCache();
        cache.put('discordRetryActivity', message, 60);
        const existing = cache.get('discordRetryLog');
        const log = existing ? JSON.parse(existing) : [];
        log.push(message);
        cache.put('discordRetryLog', JSON.stringify(log), 600);
    } catch (e) {}
}

/** Returns the most recent Discord retry message for dialog polling. */
function getDiscordRetryActivity() {
    return CacheService.getScriptCache().get('discordRetryActivity') || '';
}

/** Returns all accumulated Discord retry/rate-limit log entries for this send run. */
function getDiscordRetryLog() {
    const data = CacheService.getScriptCache().get('discordRetryLog');
    return data ? JSON.parse(data) : [];
}

function clearDiscordRetryLog_() {
    CacheService.getScriptCache().remove('discordRetryLog');
}

/**
 * Sends a message to a Discord channel via the bot API.
 * Handles chunking, test mode override, and retries (429 / 5xx).
 * @param {string} channelId The target Discord channel ID.
 * @param {string} message The message content to send.
 * @param {string} [context="General"] Optional context for logging.
 * @returns {boolean} True if all chunks sent successfully, false otherwise.
 */
function sendDiscordBotMessage_(channelId, message, context = "General") {
  const botToken = PropertiesService.getScriptProperties().getProperty(PROP_BOT_TOKEN);
  if (!botToken || botToken === 'YOUR_BOT_TOKEN_HERE') {
    Logger.log(`[${context}] Bot send skipped: BOT_TOKEN not configured.`);
    return false;
  }
  if (!channelId || !message) {
    Logger.log(`[${context}] Bot send skipped: channelId or message is empty.`);
    return false;
  }

  let targetChannelId = channelId;
  let testModeActive = isDiscordTestMode_();

  if (testModeActive) {
    const testChannelId = PropertiesService.getScriptProperties().getProperty(PROP_TEST_CHANNEL_ID);
    if (testChannelId && testChannelId !== 'YOUR_TEST_CHANNEL_ID_HERE') {
      Logger.log(`[${context}] Test Mode ACTIVE. Using test channel ${testChannelId} instead of ${channelId}.`);
      targetChannelId = testChannelId;
    } else {
      Logger.log(`[${context}] Test Mode ACTIVE but TEST_CHANNEL_ID not configured. Using original channel.`);
      testModeActive = false;
    }
  }

  const finalMessage = testModeActive ? `🧪 **[TEST MODE]** 🧪\n${message}` : message;
  const chunks = getDiscordChunks_(finalMessage);
  const apiUrl = `https://discord.com/api/v10/channels/${targetChannelId}/messages`;

  Logger.log(`[${context}] Sending ${chunks.length} chunk(s) via bot to channel ${targetChannelId}.`);

  let allChunksSent = true;
  for (let chunkIndex = 0; chunkIndex < chunks.length; chunkIndex++) {
    const chunk = chunks[chunkIndex];
    if (!chunk.trim()) continue;

    let chunkSent = false;
    let errorAttempts = 0;
    let rateLimitWaits = 0;

    while (!chunkSent && errorAttempts <= MAX_RETRIES && rateLimitWaits <= MAX_RETRIES * 2) {
      Logger.log(`[${context}] Sending chunk ${chunkIndex + 1}/${chunks.length} (errors: ${errorAttempts}/${MAX_RETRIES}, rate-limit waits: ${rateLimitWaits}/${MAX_RETRIES * 2})...`);

      const params = {
        method: 'POST',
        headers: { 'Authorization': `Bot ${botToken}` },
        contentType: 'application/json',
        payload: JSON.stringify({ content: chunk }),
        muteHttpExceptions: true,
      };

      try {
        const response = UrlFetchApp.fetch(apiUrl, params);
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();
        const headers = response.getHeaders();

        if (responseCode >= 200 && responseCode < 300) {
          Logger.log(`[${context}] Chunk ${chunkIndex + 1} sent successfully (Code: ${responseCode}).`);
          chunkSent = true;
        } else if (responseCode === 429) {
          rateLimitWaits++;
          const retryAfterHeader = headers['Retry-After'] || headers['retry-after'];
          let waitSeconds = 0;
          if (retryAfterHeader) {
            waitSeconds = parseFloat(retryAfterHeader);
          } else {
            try {
              const jsonBody = JSON.parse(responseBody);
              if (jsonBody.retry_after) waitSeconds = parseFloat(jsonBody.retry_after);
            } catch (e) { /* ignore */ }
          }
          const delayMs = waitSeconds > 0
            ? Math.ceil(waitSeconds * 1000) + 500
            : Math.min(INITIAL_BACKOFF_MS * Math.pow(2, rateLimitWaits) + Math.random() * 1000, MAX_BACKOFF_MS);

          const GAS_MAX_SLEEP_MS = 300000;
          if (delayMs > GAS_MAX_SLEEP_MS) {
            const humanTime = waitSeconds >= 60
              ? `${(waitSeconds / 60).toFixed(1)} minutes`
              : `${waitSeconds.toFixed(0)} seconds`;
            const abortMsg = `Chunk ${chunkIndex + 1}/${chunks.length}: rate limit is ${humanTime} — too long for GAS to wait. Try again later.`;
            Logger.log(`[${context}] ${abortMsg}`);
            notifyDiscordRetry_(abortMsg);
            break;
          }

          const rlMsg = `Chunk ${chunkIndex + 1}/${chunks.length}: rate limited — waiting ${(delayMs / 1000).toFixed(1)}s (retry ${rateLimitWaits}/${MAX_RETRIES * 2}).`;
          Logger.log(`[${context}] ${rlMsg}`);
          notifyDiscordRetry_(rlMsg, Math.ceil(delayMs / 1000) + 3);
          Utilities.sleep(delayMs);
        } else if (responseCode >= 500 && responseCode < 600) {
          errorAttempts++;
          const delayMs = Math.min(INITIAL_BACKOFF_MS * Math.pow(2, errorAttempts) + Math.random() * 1000, MAX_BACKOFF_MS);
          const seMsg = `Chunk ${chunkIndex + 1}/${chunks.length}: server error ${responseCode} — retrying (${errorAttempts}/${MAX_RETRIES}).`;
          Logger.log(`[${context}] ${seMsg}`);
          notifyDiscordRetry_(seMsg);
          if (errorAttempts <= MAX_RETRIES) Utilities.sleep(delayMs);
        } else {
          Logger.log(`[${context}] Client error (${responseCode}) on chunk ${chunkIndex + 1}. Body: ${responseBody}`);
          break;
        }
      } catch (fetchError) {
        errorAttempts++;
        const delayMs = Math.min(INITIAL_BACKOFF_MS * Math.pow(2, errorAttempts) + Math.random() * 1000, MAX_BACKOFF_MS);
        const feMsg = `Chunk ${chunkIndex + 1}/${chunks.length}: network error — retrying (${errorAttempts}/${MAX_RETRIES}).`;
        Logger.log(`[${context}] ${feMsg} Error: ${fetchError}.`);
        notifyDiscordRetry_(feMsg);
        if (errorAttempts <= MAX_RETRIES) Utilities.sleep(delayMs);
      }
    }

    if (!chunkSent) {
      Logger.log(`[${context}] Failed to send chunk ${chunkIndex + 1} after ${errorAttempts} errors and ${rateLimitWaits} rate-limit waits.`);
      allChunksSent = false;
      break;
    }

    if (chunks.length > 1 && chunkIndex < chunks.length - 1) {
      Utilities.sleep(1000);
    }
  }

  if (allChunksSent) {
    Logger.log(`[${context}] Successfully sent all chunks to Discord via bot (channel ${targetChannelId}).`);
  } else {
    Logger.log(`[${context}] Failed to send one or more chunks via bot (channel ${targetChannelId}).`);
  }
  return allChunksSent;
}

/**
 * Sends a message to a specified Discord webhook URL. Handles message chunking,
 * test mode override, and retries on failure (429 or 5xx).
 * Allows overriding the webhook's default username and avatar.
 * @param {string} originalWebhookUrl The intended Discord webhook URL (player or ST).
 * @param {string} message The message content to send.
 * @param {string} [context="General"] Optional context for logging.
 * @param {string} [overrideUsername=null] Optional username to override the webhook default.
 * @param {string} [overrideAvatarUrl=null] Optional avatar URL to override the webhook default. Must be a direct image link.
 * @returns {boolean} True if the message (all chunks) was sent successfully, false otherwise.
 */
function sendDiscordWebhookMessage_(originalWebhookUrl, message, context = "General", overrideUsername = null, overrideAvatarUrl = null) {
  let targetWebhookUrl = originalWebhookUrl;
  let testModeActive = isDiscordTestMode_(); // In Utilities.gs
  let testWebhookUrl = null;

  // Determine target URL based on test mode
  if (testModeActive) {
      testWebhookUrl = getTestWebhookUrl_(); // In Utilities.gs
      if (testWebhookUrl && testWebhookUrl !== 'YOUR_TEST_DISCORD_WEBHOOK_URL_HERE') {
          targetWebhookUrl = testWebhookUrl;
          Logger.log(`[${context}] Discord Test Mode ACTIVE. Overriding target webhook. Original: ${originalWebhookUrl}, Using Test: ${targetWebhookUrl}`);
      } else {
          Logger.log(`[${context}] Discord Test Mode ACTIVE, but Test Webhook URL (${PROP_TEST_WEBHOOK}) is invalid or not configured. Sending to original URL: ${originalWebhookUrl}`);
          testModeActive = false; // Treat as test mode off for notification purposes if URL is bad
      }
  }

  // Validate final URL and message
  if (!targetWebhookUrl || !message) {
    Logger.log(`[${context}] Discord message skipped: Target Webhook URL or message is empty. URL: ${targetWebhookUrl}`);
    if (!targetWebhookUrl) {
        Logger.log(`[${context}] Error: Target Discord webhook URL is missing.`);
    }
    return false; // Indicate failure
  }
  if (targetWebhookUrl.includes('?')) {
    targetWebhookUrl += '&wait=true';
  } else {
    targetWebhookUrl += '?wait=true';
  }
  // Prepend Test Mode indicator to the message if active
  const finalMessage = testModeActive ? `🧪 **[TEST MODE]** 🧪\n${message}` : message;
  Logger.log(`[${context}] Preparing to send message (length: ${finalMessage.length}) via ${targetWebhookUrl}. Override User: ${overrideUsername || 'None'}, Override Avatar: ${overrideAvatarUrl ? 'Yes' : 'No'}`);

  // --- Chunking Logic ---
  const chunks = [];
  if (finalMessage.length <= MAX_DISCORD_MESSAGE_LENGTH) {
    chunks.push(finalMessage);
  } else {
    Logger.log(`[${context}] Message exceeds max length (${MAX_DISCORD_MESSAGE_LENGTH}). Chunking...`);
    let currentChunk = "";
    const prefix = testModeActive ? `🧪 **[TEST MODE]** 🧪\n` : "";
    const messageWithoutPrefix = testModeActive ? message : finalMessage;
    let remainingMessage = messageWithoutPrefix;
    let isFirstChunk = true;

    while (remainingMessage.length > 0) {
        let chunkContent;
        let chunkPrefix = isFirstChunk ? prefix : `\n`;
        let availableLength = MAX_DISCORD_MESSAGE_LENGTH - chunkPrefix.length;

        if (remainingMessage.length <= availableLength) {
            chunkContent = chunkPrefix + remainingMessage;
            remainingMessage = "";
        } else {
            let splitIndex = remainingMessage.lastIndexOf('\n', availableLength);
            if (splitIndex <= 0) splitIndex = remainingMessage.lastIndexOf(' ', availableLength);
            if (splitIndex <= 0) splitIndex = availableLength;
            chunkContent = chunkPrefix + remainingMessage.substring(0, splitIndex);
            remainingMessage = remainingMessage.substring(splitIndex).trimStart();
        }
        chunks.push(chunkContent.trim());
        isFirstChunk = false;
    }
     Logger.log(`[${context}] Message split into ${chunks.length} chunks.`);
  }

  // --- Sending Logic with Retries ---
  let allChunksSent = true;
  for (let chunkIndex = 0; chunkIndex < chunks.length; chunkIndex++) {
      const chunk = chunks[chunkIndex];
      if (chunk.trim() === "") continue;

      let chunkSent = false;
      // Two independent counters so a run of 429 rate-limit waits never burns
      // the error-retry budget, and vice versa.
      let errorAttempts = 0;    // 5xx responses and network/fetch exceptions
      let rateLimitWaits = 0;   // 429 responses

      while (!chunkSent && errorAttempts <= MAX_RETRIES && rateLimitWaits <= MAX_RETRIES * 2) {
          Logger.log(`[${context}] Sending chunk ${chunkIndex + 1}/${chunks.length} (errors: ${errorAttempts}/${MAX_RETRIES}, rate-limit waits: ${rateLimitWaits}/${MAX_RETRIES * 2})...`);

          const payloadObject = { content: chunk };
          if (overrideUsername) {
              payloadObject.username = overrideUsername;
          }
          if (overrideAvatarUrl) {
              if (overrideAvatarUrl.startsWith('http://') || overrideAvatarUrl.startsWith('https://')) {
                  payloadObject.avatar_url = overrideAvatarUrl;
              } else {
                  Logger.log(`[${context}] Warning: Invalid avatar_url provided, skipping: ${overrideAvatarUrl}`);
              }
          }

          const params = {
              method: 'POST',
              contentType: 'application/json',
              payload: JSON.stringify(payloadObject),
              muteHttpExceptions: true,
          };

          try {
              const response = UrlFetchApp.fetch(targetWebhookUrl, params);
              const responseCode = response.getResponseCode();
              const responseBody = response.getContentText();
              const headers = response.getHeaders();

              if (responseCode >= 200 && responseCode < 300) {
                  Logger.log(`[${context}] Chunk ${chunkIndex + 1} sent successfully (Code: ${responseCode}).`);
                  chunkSent = true;
              } else if (responseCode === 429) {
                  rateLimitWaits++;
                  const retryAfterHeader = headers['Retry-After'] || headers['retry-after'];
                  let waitSeconds = 0;
                  if (retryAfterHeader) {
                      waitSeconds = parseFloat(retryAfterHeader);
                  } else {
                      try {
                          const jsonBody = JSON.parse(responseBody);
                          if (jsonBody.retry_after) waitSeconds = parseFloat(jsonBody.retry_after);
                      } catch (parseError) { /* Ignore */ }
                  }
                  const delayMs = waitSeconds > 0
                      ? Math.ceil(waitSeconds * 1000) + 500
                      : Math.min(INITIAL_BACKOFF_MS * Math.pow(2, rateLimitWaits) + Math.random() * 1000, MAX_BACKOFF_MS);

                  // GAS Utilities.sleep() max is ~300s. If Discord's retry_after exceeds that,
                  // we can't wait it out in a single execution — bail with a clear message.
                  const GAS_MAX_SLEEP_MS = 300000;
                  if (delayMs > GAS_MAX_SLEEP_MS) {
                      const humanTime = waitSeconds >= 60
                          ? `${(waitSeconds / 60).toFixed(1)} minutes`
                          : `${waitSeconds.toFixed(0)} seconds`;
                      const abortMsg = `Chunk ${chunkIndex + 1}/${chunks.length}: Discord rate limit is ${humanTime} — too long for GAS to wait. Try again later.`;
                      Logger.log(`[${context}] ${abortMsg}`);
                      notifyDiscordRetry_(abortMsg);
                      break; // exit while loop; outer for loop will see !chunkSent and stop
                  }

                  const rlMsg = `Chunk ${chunkIndex + 1}/${chunks.length}: rate limited — waiting ${(delayMs / 1000).toFixed(1)}s (retry ${rateLimitWaits}/${MAX_RETRIES * 2}).`;
                  Logger.log(`[${context}] ${rlMsg}`);
                  notifyDiscordRetry_(rlMsg, Math.ceil(delayMs / 1000) + 3);
                  if (rateLimitWaits <= MAX_RETRIES * 2) {
                      Utilities.sleep(delayMs);
                  }
              } else if (responseCode >= 500 && responseCode < 600) {
                  errorAttempts++;
                  const delayMs = Math.min(INITIAL_BACKOFF_MS * Math.pow(2, errorAttempts) + Math.random() * 1000, MAX_BACKOFF_MS);
                  const seMsg = `Chunk ${chunkIndex + 1}/${chunks.length}: server error ${responseCode} — retrying (${errorAttempts}/${MAX_RETRIES}).`;
                  Logger.log(`[${context}] ${seMsg} Sleeping ${delayMs}ms.`);
                  notifyDiscordRetry_(seMsg);
                  if (errorAttempts <= MAX_RETRIES) {
                      Utilities.sleep(delayMs);
                  }
              } else {
                  // 4xx (except 429) — bad request, invalid webhook, etc. — don't retry
                  Logger.log(`[${context}] Client error (${responseCode}) on chunk ${chunkIndex + 1}. Not retrying. Body: ${responseBody}`);
                  break;
              }

          } catch (fetchError) {
              errorAttempts++;
              const delayMs = Math.min(INITIAL_BACKOFF_MS * Math.pow(2, errorAttempts) + Math.random() * 1000, MAX_BACKOFF_MS);
              const feMsg = `Chunk ${chunkIndex + 1}/${chunks.length}: network error — retrying (${errorAttempts}/${MAX_RETRIES}).`;
              Logger.log(`[${context}] ${feMsg} Error: ${fetchError}. Sleeping ${delayMs}ms.`);
              notifyDiscordRetry_(feMsg);
              if (errorAttempts <= MAX_RETRIES) {
                  Utilities.sleep(delayMs);
              }
          }
      } // End retry while loop for this chunk

      if (!chunkSent) {
          Logger.log(`[${context}] Failed to send chunk ${chunkIndex + 1} after ${errorAttempts} errors and ${rateLimitWaits} rate-limit waits.`);
          allChunksSent = false;
          break;
      }

      // 1 s between chunks keeps us well within Discord's 30-req/min webhook rate limit
      if (chunks.length > 1 && chunkIndex < chunks.length - 1) {
          Utilities.sleep(1000);
      }

  } // End loop through chunks

  if (allChunksSent) {
      Logger.log(`[${context}] Successfully sent all chunks to Discord via ${targetWebhookUrl}.`);
      return true;
  } else {
      Logger.log(`[${context}] Failed to send one or more chunks to Discord via ${targetWebhookUrl}.`);
      // Error is already logged, calling function can decide to alert UI
      return false;
  }
}

/**
 * Handles sending downtime results via Discord when the checkbox is checked or manually triggered.
 * Uses test mode if enabled. Shows alert on final failure.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet.
 * @param {number} responseRowIndex The row index of the response row (where checkbox was checked).
 */
function handleSendDiscord_(sheet, responseRowIndex, isBulkSend = false, skipConfirmDialog = false) {
    const submissionRowIndex = responseRowIndex - 1;
    if (submissionRowIndex < 1) {
        Logger.log(`Cannot send Discord for row ${responseRowIndex}: No corresponding submission row found.`);
        return;
    }

    const responseRowData = sheet.getRange(responseRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const submissionRowData = sheet.getRange(submissionRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const characterName = responseRowData[CHARACTER_NAME_COL - 1];
    const playerChannelId = getCharacterChannelId_(characterName); // In SheetData.gs — col Y
    const playerWebhook = getCharacterWebhook_(characterName);     // In SheetData.gs — col X
    const useBotSend = !!playerChannelId;
    const testModeActive = isDiscordTestMode_(); // In Utilities.gs
    const sheetName = sheet.getName();

    // Determine description shown in the confirmation dialog
    let targetDescription = characterName;

    if (testModeActive) {
        const testChannelId = PropertiesService.getScriptProperties().getProperty(PROP_TEST_CHANNEL_ID);
        const testWebhookUrl = getTestWebhookUrl_(); // In Utilities.gs
        if (useBotSend && testChannelId && testChannelId !== 'YOUR_TEST_CHANNEL_ID_HERE') {
            targetDescription = `TEST CHANNEL (for ${characterName})`;
        } else if (testWebhookUrl && testWebhookUrl !== 'YOUR_TEST_DISCORD_WEBHOOK_URL_HERE') {
            targetDescription = `TEST WEBHOOK (for ${characterName})`;
        } else {
            const msg = `Cannot send in Test Mode: neither TEST_CHANNEL_ID (${PROP_TEST_CHANNEL_ID}) nor TEST_WEBHOOK (${PROP_TEST_WEBHOOK}) is configured.`;
            Logger.log(`Discord send cancelled for ${characterName} (Row ${responseRowIndex}): Test mode ON but no test target configured.`);
            try { sheet.getRange(responseRowIndex, SEND_DISCORD_COL).setValue(false); } catch(e){}
            if (isBulkSend) throw new Error(msg);
            SpreadsheetApp.getUi().alert(msg);
            return;
        }
    } else if (!playerChannelId && !playerWebhook) {
        const msg = `No Discord channel ID (col Y) or webhook (col X) found for '${characterName}'. Check the '${CHARACTER_SHEET_NAME}' sheet.`;
        Logger.log(`Discord send cancelled for ${characterName} (Row ${responseRowIndex}): No channel ID or webhook found.`);
        try { sheet.getRange(responseRowIndex, SEND_DISCORD_COL).setValue(false); } catch(e){}
        if (isBulkSend) throw new Error(msg);
        SpreadsheetApp.getUi().alert(msg);
        return;
    } else {
        targetDescription = useBotSend ? `${characterName} (Bot)` : `${characterName} (Webhook)`;
    }

    // --- Construct Message Body ---
    let messageBody = "";
    const year = PropertiesService.getScriptProperties().getProperty(PROP_DOWNTIME_YEAR);
    const month = PropertiesService.getScriptProperties().getProperty(PROP_DOWNTIME_MONTH);
    const title = `**Downtime Results for ${characterName} (${month}, ${year})**\n\n`;

    // Start message construction after basic info columns
    for (let j = CHARACTER_NAME_COL; j < headers.length; j++) {
        const header = headers[j];
        const submissionText = submissionRowData[j] ? String(submissionRowData[j]).trim() : "";
        const responseText = responseRowData[j] ? String(responseRowData[j]).trim() : "";

        // Only include sections where a response was actually provided
        if (responseText !== "") {
            messageBody += `**${header}**\n`;
            if (submissionText !== "") {
                messageBody += `*Your Action:* ${submissionText}\n`;
            } else {
                 messageBody += `*Your Action:* (No submission text found)\n`;
            }
            messageBody += `*Result:* ${responseText}\n\n`;
        } else if (submissionText !== "") {
             // Optionally include submitted actions even if no response yet (maybe for confirmation?)
             // messageBody += `**${header}**\n*Your Action:* ${submissionText}\n*Result:* (Pending)\n\n`;
        }
    }

    if (messageBody.trim() === "") {
        const msg = `No completed downtime results found for ${characterName} to send.`;
        Logger.log(`No downtime results found to send for ${characterName} (Row ${responseRowIndex}).`);
        try { sheet.getRange(responseRowIndex, SEND_DISCORD_COL).setValue(false); } catch(e){}
        if (isBulkSend) throw new Error(msg);
        SpreadsheetApp.getUi().alert(msg);
        return;
    }

    // --- Show Confirmation Dialog (non-bulk only) ---
    if (!isBulkSend && !skipConfirmDialog) {
        const fullMessage = title + messageBody;
        const previewMessage = testModeActive ? `🧪 **[TEST MODE]** 🧪\n${fullMessage}` : fullMessage;
        const chunks = getDiscordChunks_(previewMessage);

        const template = HtmlService.createTemplateFromFile('DiscordConfirmDialog');
        template.targetDescription = targetDescription;
        template.useBotSend = useBotSend;
        template.chunks = chunks;
        template.sheetName = sheetName;
        template.responseRowIndex = responseRowIndex;

        const height = Math.min(200 + chunks.length * 220, 700);
        SpreadsheetApp.getUi().showModalDialog(
            template.evaluate().setWidth(620).setHeight(height),
            'Confirm Discord Send'
        );
        return; // Sending happens via performDiscordSend callback if user confirms
    }

    // --- Send Message ---
    let success = false;
    const logDetailsBase = `Row: ${responseRowIndex}, Char: ${characterName}, Target: ${targetDescription}`;
    try {
        success = useBotSend
            ? sendDiscordBotMessage_(playerChannelId, title + messageBody, `Downtime Send - ${characterName}`)
            : sendDiscordWebhookMessage_(playerWebhook, title + messageBody, `Downtime Send - ${characterName}`);

        if (success) {
            // --- Update Sheet on Success ---
            const statusCell = sheet.getRange(responseRowIndex, STATUS_COL);
            const timestampCell = sheet.getRange(responseRowIndex, TIMESTAMP_COL); // Use response row timestamp col
            const checkboxCell = sheet.getRange(responseRowIndex, SEND_DISCORD_COL);

            timestampCell.setValue(new Date()).setBackgroundRGB(229, 255, 204); // Green background for timestamp
            statusCell.setValue('sent');
            checkboxCell.setValue(true);
            checkboxCell.setBackgroundRGB(229, 255, 204); // Green background for checkbox

            Logger.log(`Successfully initiated Discord send for ${targetDescription}.`);
            logAudit_('Sent Discord', sheetName, logDetailsBase); // In Utilities.gs
             // Optionally show success alert
             // ui.alert(`Discord message sent successfully to ${targetDescription}.`);
        } else {
             // Failure after retries
             Logger.log(`Failed to send Discord message for ${targetDescription} (Row ${responseRowIndex}) after retries.`);
             logAudit_('Sent Discord FAILED', sheetName, `${logDetailsBase}, Check Logs`); // In Utilities.gs
             if (!isBulkSend) {
               SpreadsheetApp.getUi().alert(`Failed to send Discord message for ${targetDescription} after multiple attempts. Please check the script logs and try again later.`);
             }
             try { sheet.getRange(responseRowIndex, SEND_DISCORD_COL).setValue(false); } catch(e){}
        }

    } catch (error) {
        // Catch errors from sendDiscordWebhookMessage_ (e.g., invalid URL if not caught earlier)
        Logger.log(`Error during Discord send process for ${targetDescription} (Row ${responseRowIndex}): ${error}`);
        logAudit_('Sent Discord ERROR', sheetName, `${logDetailsBase}, Error: ${error.message}`); // In Utilities.gs
        try { sheet.getRange(responseRowIndex, SEND_DISCORD_COL).setValue(false); } catch(e){}
        if (isBulkSend) throw error;
        SpreadsheetApp.getUi().alert(`An unexpected error occurred sending the Discord message for ${targetDescription}: ${error.message}`);
    }
    return success;
}

/** Splits a message into Discord-sized chunks using the same algorithm as sendDiscordWebhookMessage_. */
function getDiscordChunks_(message) {
    if (message.length <= MAX_DISCORD_MESSAGE_LENGTH) {
        return [message];
    }
    const chunks = [];
    let remaining = message;
    while (remaining.length > 0) {
        if (remaining.length <= MAX_DISCORD_MESSAGE_LENGTH) {
            chunks.push(remaining);
            break;
        }
        let splitIndex = remaining.lastIndexOf('\n', MAX_DISCORD_MESSAGE_LENGTH);
        if (splitIndex <= 0) splitIndex = remaining.lastIndexOf(' ', MAX_DISCORD_MESSAGE_LENGTH);
        if (splitIndex <= 0) splitIndex = MAX_DISCORD_MESSAGE_LENGTH;
        chunks.push(remaining.substring(0, splitIndex).trim());
        remaining = remaining.substring(splitIndex).trimStart();
    }
    return chunks;
}

/** Called from DiscordConfirmDialog when the user clicks Send. */
function performDiscordSend(sheetName, responseRowIndex) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
        Logger.log(`performDiscordSend: Sheet '${sheetName}' not found.`);
        return;
    }
    handleSendDiscord_(sheet, responseRowIndex, false, true);
}

/** Called from DiscordConfirmDialog when the user clicks Manually Sent. Marks the row as sent without using the webhook. */
function markDiscordManuallySent(sheetName, responseRowIndex) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
        Logger.log(`markDiscordManuallySent: Sheet '${sheetName}' not found.`);
        return;
    }
    const statusCell = sheet.getRange(responseRowIndex, STATUS_COL);
    const timestampCell = sheet.getRange(responseRowIndex, TIMESTAMP_COL);
    const checkboxCell = sheet.getRange(responseRowIndex, SEND_DISCORD_COL);

    timestampCell.setValue(new Date()).setBackgroundRGB(229, 255, 204);
    statusCell.setValue('sent (manual)');
    checkboxCell.setValue(true);
    checkboxCell.setBackgroundRGB(229, 255, 204);

    Logger.log(`Row ${responseRowIndex} on '${sheetName}' marked as manually sent.`);
    logAudit_('Sent Discord (Manual)', sheetName, `Row: ${responseRowIndex}`);
}

/** Called from DiscordConfirmDialog when the user clicks Cancel. */
function cancelDiscordSend(sheetName, responseRowIndex) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet) {
        try { sheet.getRange(responseRowIndex, SEND_DISCORD_COL).setValue(false); } catch(e){}
    }
    Logger.log(`Discord send cancelled by user for row ${responseRowIndex} on sheet '${sheetName}'.`);
}


/** Sends the main downtime progress report to the ST Discord channel. */
function sendDowntimeReportToDiscord() {
  Logger.log('Sending Downtime Report to Discord...');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const results = getDowntimeCompletionData_(sheet); // In SheetData.gs
  const scriptProps = PropertiesService.getScriptProperties();
  const stChannelId = scriptProps.getProperty(PROP_ST_CHANNEL_ID);
  const stWebhook = scriptProps.getProperty(PROP_ST_WEBHOOK);
  const useBot = !!(stChannelId && stChannelId !== 'YOUR_ST_CHANNEL_ID_HERE');

  if (!useBot && (!stWebhook || stWebhook === 'YOUR_STORYTELLER_DISCORD_WEBHOOK_URL_HERE')) {
    SpreadsheetApp.getUi().alert(`Error: Neither ST_CHANNEL_ID (${PROP_ST_CHANNEL_ID}) nor ST_WEBHOOK (${PROP_ST_WEBHOOK}) is configured in Script Properties.`);
    Logger.log(`Discord report failed: no ST channel ID or webhook configured.`);
    return;
  }
   if (!results) {
    SpreadsheetApp.getUi().alert('Could not retrieve downtime data to send report. Is this the correct sheet?');
    Logger.log('Failed to get downtime completion data for Discord report.');
    return;
  }


  const completionPercentage = results.totalDowntimeCells > 0
      ? (results.completedDowntimeCells / results.totalDowntimeCells) * 100
      : 0;

  let discordMessage = `📊 **Downtime Report (${sheet.getName()})** 📊\n\n`;
  discordMessage += `**Overall Completion:** ${completionPercentage.toFixed(1)}% (${results.completedDowntimeCells}/${results.totalDowntimeCells})\n`;
  discordMessage += `**Characters Submitted:** ${results.characterCount}\n\n`;
  discordMessage += `**Word Count Stats (Submissions):**\n`;
  discordMessage += `  Avg: ${results.averageWords.toFixed(1)}, Median: ${results.medianWords}, Min: ${results.minWords} (${results.minWordCell}), Max: ${results.maxWords} (${results.maxWordCell})\n`;
  discordMessage += `**Word Count Stats (Responses):**\n`;
  discordMessage += `  Avg: ${results.averageResponseWords.toFixed(1)}, Median: ${results.medianResponseWords}, Min: ${results.minResponseWords} (${results.minResponseWordCell}), Max: ${results.maxResponseWords} (${results.maxResponseWordCell})\n\n`;
  discordMessage += `**Keyword Breakdown:**\n`;

  const keywordKeys = Object.keys(DOWNTIME_KEYWORDS); // In Constants.gs
  if (keywordKeys.length > 0) {
    keywordKeys.forEach(key => {
      const total = results.keywordCounts[key] || 0;
      const completed = results.keywordCompletedCounts[key] || 0;
      if (total > 0) { // Only show keywords that were actually submitted
          const keywordPercentage = (completed / total) * 100;
          const overallPercentage = results.totalDowntimeCells > 0 ? (total / results.totalDowntimeCells) * 100 : 0;
          discordMessage += `  • **${key}:** ${keywordPercentage.toFixed(1)}% (${completed}/${total}) _(${overallPercentage.toFixed(1)}% of total)_\n`;
      }
    });
  } else {
    discordMessage += "  _(No keyword data available)_\n";
  }

  try {
    const success = useBot
      ? sendDiscordBotMessage_(stChannelId, discordMessage, 'Downtime Report')
      : sendDiscordWebhookMessage_(stWebhook, discordMessage, 'Downtime Report');
    if (success) {
        SpreadsheetApp.getUi().alert('Downtime report sent to Discord!');
        Logger.log('Downtime report successfully sent to Discord.');
    } else {
        SpreadsheetApp.getUi().alert('Failed to send Downtime report to Discord after retries. Check Logs.');
        Logger.log('Failed to send Downtime report to Discord after retries.');
    }
  } catch (error) {
    Logger.log(`Error sending downtime report to Discord: ${error}\nStack: ${error.stack}`);
    SpreadsheetApp.getUi().alert(`Error sending report to Discord: ${error.message}`);
  }
}

/** Sends the list of missing downtime responses to the ST Discord channel. */
function sendMissingDowntimeResponsesToDiscord() {
  Logger.log('Sending Missing Downtime Responses Report to Discord...');
  const scriptProps = PropertiesService.getScriptProperties();
  const stChannelId = scriptProps.getProperty(PROP_ST_CHANNEL_ID);
  const stWebhook = scriptProps.getProperty(PROP_ST_WEBHOOK);
  const useBot = !!(stChannelId && stChannelId !== 'YOUR_ST_CHANNEL_ID_HERE');

  if (!useBot && (!stWebhook || stWebhook === 'YOUR_STORYTELLER_DISCORD_WEBHOOK_URL_HERE')) {
    SpreadsheetApp.getUi().alert(`Error: Neither ST_CHANNEL_ID (${PROP_ST_CHANNEL_ID}) nor ST_WEBHOOK (${PROP_ST_WEBHOOK}) is configured in Script Properties.`);
    Logger.log(`Missing responses report failed: no ST channel ID or webhook configured.`);
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const missingResponses = getMissingDowntimeData_(sheet); // In SheetData.gs

  if (!missingResponses) {
      SpreadsheetApp.getUi().alert('Could not retrieve missing downtime data. Is this the correct sheet?');
      Logger.log('Failed to get missing downtime data for Discord report.');
      return;
  }

  let discordMessage;
  if (missingResponses.length === 0) {
    discordMessage = `✅ All downtime responses filled out for sheet: ${sheet.getName()}!`;
  } else {
    discordMessage = `📝 **Missing Downtime Responses (${sheet.getName()})** 📝\n\n`;
    missingResponses.forEach(entry => {
      // Truncate long text for Discord message
      const truncatedText = entry.text.length > 150 ? entry.text.substring(0, 147) + "..." : entry.text;
      discordMessage += `• **${entry.characterName}** (${entry.header} - Cell: ${entry.cell}): "${truncatedText}"\n`;
    });
     discordMessage += `\nSheet Link: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`; // Add link to sheet
  }

  try {
    const success = useBot
      ? sendDiscordBotMessage_(stChannelId, discordMessage, 'Missing Responses Report')
      : sendDiscordWebhookMessage_(stWebhook, discordMessage, 'Missing Responses Report');
    if (success) {
        SpreadsheetApp.getUi().alert('Missing downtime responses report sent to Discord!');
        Logger.log('Missing downtime responses report successfully sent to Discord.');
    } else {
        SpreadsheetApp.getUi().alert('Failed to send Missing Responses report to Discord after retries. Check Logs.');
        Logger.log('Failed to send Missing Responses report to Discord after retries.');
    }
  } catch (error) {
    Logger.log(`Error sending missing downtime responses report to Discord: ${error}\nStack: ${error.stack}`);
    SpreadsheetApp.getUi().alert(`Error sending report to Discord: ${error.message}`);
  }
}

/**
 * Sends a test message to the configured Storyteller Discord channel.
 * Prefers the bot channel ID; falls back to the ST webhook.
 */
function sendTestMessageToStorytellers_() {
  Logger.log('Sending test message to ST Discord channel...');
  const scriptProps = PropertiesService.getScriptProperties();
  const stChannelId = scriptProps.getProperty(PROP_ST_CHANNEL_ID);
  const stWebhook = scriptProps.getProperty(PROP_ST_WEBHOOK);
  const useBot = !!(stChannelId && stChannelId !== 'YOUR_ST_CHANNEL_ID_HERE');
  const sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();

  if (!useBot && (!stWebhook || stWebhook === 'YOUR_STORYTELLER_DISCORD_WEBHOOK_URL_HERE')) {
    SpreadsheetApp.getUi().alert(`Error: Neither ST_CHANNEL_ID (${PROP_ST_CHANNEL_ID}) nor ST_WEBHOOK (${PROP_ST_WEBHOOK}) is configured in Script Properties.`);
    Logger.log(`Test message failed: no ST channel ID or webhook configured.`);
    return;
  }

  const method = useBot ? 'Bot' : 'Webhook';
  const message = `👋 This is a test message from the Google Sheet script (${SpreadsheetApp.getActiveSpreadsheet().getName()}) at ${new Date().toLocaleString()}. If you see this, the ${method} configuration is working correctly!`;
  const context = 'ST Test Message';

  try {
    const success = useBot
      ? sendDiscordBotMessage_(stChannelId, message, context)
      : sendDiscordWebhookMessage_(stWebhook, message, context);

    if (success) {
        const finalDestination = isDiscordTestMode_() ? `Test ${method}` : `ST Channel (${method})`;
        logAudit_('Sent Test Message', sheetName, `Destination: ${finalDestination}`);
        SpreadsheetApp.getUi().alert(`Test message sent successfully to the ${finalDestination}!`);
        Logger.log(`Test message sent successfully to ${finalDestination}.`);
    } else {
        logAudit_('Sent Test Message FAILED', sheetName, `Check Logs`);
        SpreadsheetApp.getUi().alert(`Failed to send Test message to Discord after retries. Check Logs.`);
        Logger.log(`Failed to send Test message to Discord after retries.`);
    }
  } catch (error) {
    Logger.log(`Error sending test message: ${error}\nStack: ${error.stack}`);
    logAudit_('Sent Test Message ERROR', sheetName, `Error: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error sending test message: ${error.message}`);
  }
}

/** Counts approved characters and sends the count to the ST Discord channel. */
function checkCharacterCount_() {
    Logger.log('Checking approved character count...');
    const sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    const characterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHARACTER_SHEET_NAME);

    if (!characterSheet) {
        SpreadsheetApp.getUi().alert(`Error: Cannot find the '${CHARACTER_SHEET_NAME}' sheet.`);
        Logger.log(`Character count check failed: Sheet '${CHARACTER_SHEET_NAME}' not found.`);
        return;
    }

    const scriptProps = PropertiesService.getScriptProperties();
    const stChannelId = scriptProps.getProperty(PROP_ST_CHANNEL_ID);
    const stWebhook = scriptProps.getProperty(PROP_ST_WEBHOOK);
    const useBot = !!(stChannelId && stChannelId !== 'YOUR_ST_CHANNEL_ID_HERE');

    if (!useBot && (!stWebhook || stWebhook === 'YOUR_STORYTELLER_DISCORD_WEBHOOK_URL_HERE')) {
        SpreadsheetApp.getUi().alert(`Error: Neither ST_CHANNEL_ID (${PROP_ST_CHANNEL_ID}) nor ST_WEBHOOK (${PROP_ST_WEBHOOK}) is configured in Script Properties.`);
        Logger.log(`Character count check failed: no ST channel ID or webhook configured.`);
        return;
    }

    try {
        const data = characterSheet.getDataRange().getValues();
        let approvedCount = 0;
        // Start from row 1 to skip header
        for (let i = 1; i < data.length; i++) {
            // Check the approval status column (adjust index if needed)
            if (data[i][CHAR_SHEET_APPROVAL_COL - 1] && String(data[i][CHAR_SHEET_APPROVAL_COL - 1]).trim().toLowerCase() === "approved") { // In Constants.gs
                approvedCount++;
            }
        }

        const message = `ℹ️ Character Count Check: There are currently **${approvedCount}** approved characters in the '${CHARACTER_SHEET_NAME}' sheet.`;
        const success = useBot
            ? sendDiscordBotMessage_(stChannelId, message, 'Character Count')
            : sendDiscordWebhookMessage_(stWebhook, message, 'Character Count');
        if(success) {
            logAudit_('Character Count Check', sheetName, `Result: ${approvedCount} approved.`); // In Utilities.gs
            SpreadsheetApp.getUi().alert(`Character count (${approvedCount}) sent to Discord.`);
            Logger.log(`Approved character count: ${approvedCount}. Report sent to Discord.`);
        } else {
             logAudit_('Character Count Check FAILED', sheetName, `Send failed. Count was ${approvedCount}.`); // In Utilities.gs
             SpreadsheetApp.getUi().alert(`Failed to send Character Count report to Discord after retries. Check Logs.`);
             Logger.log(`Failed to send Character Count report to Discord after retries.`);
        }

    } catch (error) {
        Logger.log(`Error checking character count: ${error}\nStack: ${error.stack}`);
        logAudit_('Character Count Check ERROR', sheetName, `Error: ${error.message}`);
        SpreadsheetApp.getUi().alert(`Error checking character count: ${error.message}`);
    }
}

/**
 * Maintenance function: reads each character's existing webhook URL (col X), calls the
 * Discord webhook GET endpoint to retrieve the channel_id, and writes it to col Y.
 * If BOT_TOKEN is configured, also fetches the channel name from the API and writes it to col Z.
 * Skips rows that already have a channel ID in col Y.
 */
function populateChannelInfoFromWebhooks_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const characterSheet = ss.getSheetByName(CHARACTER_SHEET_NAME);

    if (!characterSheet) {
        ui.alert(`Cannot find '${CHARACTER_SHEET_NAME}' sheet.`);
        return;
    }

    const scriptProps = PropertiesService.getScriptProperties();
    const botToken = scriptProps.getProperty(PROP_BOT_TOKEN);
    const guildId  = scriptProps.getProperty(PROP_GUILD_ID);
    const hasBotToken = !!(botToken && botToken !== 'YOUR_BOT_TOKEN_HERE');
    const hasGuildId  = !!(guildId  && guildId  !== 'YOUR_GUILD_ID_HERE');

    const lastRow = characterSheet.getLastRow();
    if (lastRow < 2) {
        ui.alert('No character rows found in the Characters sheet.');
        return;
    }
    const numRows = lastRow - 1;

    // Counters for the summary alert
    let p1IdsAdded = 0, p1NamesAdded = 0;                 // Pass 1: guild query
    let p2IdsAdded = 0, p2IdsErrors = 0;                  // Pass 2: individual webhook GETs (fallback)
    let p3NamesAdded = 0, p3NamesErrors = 0;               // Pass 3: individual channel name fetches (fallback)
    let rateLimitAbortAt = null;
    let firstNameError   = null;

    // -------------------------------------------------------------------------
    // PASS 1 — Guild-level query: two bot API calls give us every webhook→channel
    //          mapping and every channel name in the server simultaneously.
    // -------------------------------------------------------------------------
    if (hasBotToken && hasGuildId) {
        ss.toast('Pass 1 of 3: querying guild webhooks & channels…', 'Populate Channel Info', 30);
        Logger.log('populateChannelInfoFromWebhooks_: starting Pass 1 (guild query).');

        try {
            // Fetch all webhooks in the server
            const whResp = UrlFetchApp.fetch(
                `https://discord.com/api/v10/guilds/${guildId}/webhooks`,
                { headers: { 'Authorization': `Bot ${botToken}` }, muteHttpExceptions: true }
            );

            if (whResp.getResponseCode() !== 200) {
                Logger.log(`Pass 1: guild webhooks fetch failed (${whResp.getResponseCode()}: ${whResp.getContentText()}) — skipping pass.`);
            } else {
                // webhookId → channelId
                const webhookToChannel = {};
                JSON.parse(whResp.getContentText()).forEach(wh => {
                    if (wh.id && wh.channel_id) webhookToChannel[wh.id] = wh.channel_id;
                });

                // Fetch all channels in the server
                const chResp = UrlFetchApp.fetch(
                    `https://discord.com/api/v10/guilds/${guildId}/channels`,
                    { headers: { 'Authorization': `Bot ${botToken}` }, muteHttpExceptions: true }
                );

                // channelId → channelName
                const channelToName = {};
                if (chResp.getResponseCode() === 200) {
                    JSON.parse(chResp.getContentText()).forEach(ch => {
                        if (ch.id && ch.name) channelToName[ch.id] = ch.name;
                    });
                } else {
                    Logger.log(`Pass 1: guild channels fetch failed (${chResp.getResponseCode()}) — names won't be filled by this pass.`);
                }

                // Apply to sheet
                const sheetData = characterSheet.getRange(2, 1, numRows, CHAR_SHEET_CHANNEL_NAME_COL).getValues();

                for (let i = 0; i < sheetData.length; i++) {
                    const row          = i + 2;
                    const characterName = String(sheetData[i][CHAR_SHEET_NAME_COL - 1] || '').trim();
                    const webhookUrl    = String(sheetData[i][CHAR_SHEET_WEBHOOK_COL - 1] || '').trim();
                    const existingId    = String(sheetData[i][CHAR_SHEET_CHANNEL_ID_COL - 1] || '').trim();
                    const existingName  = String(sheetData[i][CHAR_SHEET_CHANNEL_NAME_COL - 1] || '').trim();
                    const nameIsGood    = existingName !== '' && !existingName.startsWith('(name');

                    if (!characterName || !webhookUrl) continue;

                    // Extract webhook ID from the URL (format: /webhooks/{id}/{token})
                    const m = webhookUrl.match(/\/webhooks\/(\d+)\//);
                    if (!m) continue;
                    const channelId = webhookToChannel[m[1]];
                    if (!channelId) continue; // webhook not (or no longer) in this server

                    if (!existingId) {
                        characterSheet.getRange(row, CHAR_SHEET_CHANNEL_ID_COL).setValue(channelId);
                        p1IdsAdded++;
                    }
                    if (!nameIsGood && channelToName[channelId]) {
                        characterSheet.getRange(row, CHAR_SHEET_CHANNEL_NAME_COL).setValue(channelToName[channelId]);
                        p1NamesAdded++;
                    }
                }
            }
        } catch (e) {
            Logger.log(`Pass 1: error — ${e}`);
        }

        Logger.log(`Pass 1 complete. IDs added: ${p1IdsAdded}, names added: ${p1NamesAdded}.`);
    } else {
        Logger.log('Pass 1 skipped: BOT_TOKEN or GUILD_ID not configured.');
    }

    // -------------------------------------------------------------------------
    // PASS 2 — Individual webhook GETs for any rows still missing a channel ID
    //          (fallback: deleted webhooks, or no GUILD_ID configured)
    // -------------------------------------------------------------------------
    ss.toast('Pass 2 of 3: individual webhook lookups for remaining rows…', 'Populate Channel Info', 60);
    Logger.log('populateChannelInfoFromWebhooks_: starting Pass 2 (webhook GET fallback).');

    const pass2Data  = characterSheet.getRange(2, 1, numRows, CHAR_SHEET_WEBHOOK_COL).getValues();
    const afterPass1Ids = characterSheet.getRange(2, CHAR_SHEET_CHANNEL_ID_COL, numRows, 1).getValues();
    const WEBHOOK_MAX_WAIT_MS = 30000;

    for (let i = 0; i < pass2Data.length; i++) {
        const row           = i + 2;
        const characterName = String(pass2Data[i][CHAR_SHEET_NAME_COL - 1] || '').trim();
        const webhookUrl    = String(pass2Data[i][CHAR_SHEET_WEBHOOK_COL - 1] || '').trim();
        const existingId    = String(afterPass1Ids[i][0] || '').trim();

        if (!characterName || existingId) continue; // blank row or already resolved

        if (!webhookUrl || !webhookUrl.startsWith('https://discord.com/api/webhooks/')) {
            continue; // no webhook, nothing to try
        }

        ss.toast(`Pass 2: ${characterName}…`, 'Populate Channel Info', 10);

        try {
            let webhookCode = 0;
            let webhookData = null;

            for (let attempt = 0; attempt <= 1; attempt++) {
                const resp  = UrlFetchApp.fetch(webhookUrl, { muteHttpExceptions: true });
                webhookCode = resp.getResponseCode();

                if (webhookCode === 200) {
                    webhookData = JSON.parse(resp.getContentText());
                    break;
                } else if (webhookCode === 429 && attempt === 0) {
                    const headers       = resp.getHeaders();
                    const retryAfterSec = parseFloat(headers['Retry-After'] || headers['retry-after'] || '0');
                    const waitMs        = retryAfterSec > 0 ? Math.ceil(retryAfterSec * 1000) + 500 : 5000;
                    if (waitMs > WEBHOOK_MAX_WAIT_MS) {
                        rateLimitAbortAt = retryAfterSec > 0
                            ? `~${Math.ceil(retryAfterSec / 60)} min (retry after ${new Date(Date.now() + waitMs).toLocaleTimeString()})`
                            : 'unknown duration';
                        Logger.log(`Pass 2 Row ${row} (${characterName}): long 429 (${rateLimitAbortAt}) — aborting pass.`);
                        break;
                    }
                    Logger.log(`Pass 2 Row ${row} (${characterName}): 429, waiting ${(waitMs / 1000).toFixed(1)}s…`);
                    Utilities.sleep(waitMs);
                } else {
                    break;
                }
            }

            if (rateLimitAbortAt) break;

            if (webhookCode !== 200 || !webhookData) {
                Logger.log(`Pass 2 Row ${row} (${characterName}): webhook GET failed (${webhookCode}) — skipping.`);
                p2IdsErrors++;
                continue;
            }

            const channelId = String(webhookData.channel_id || '').trim();
            if (!channelId) {
                Logger.log(`Pass 2 Row ${row} (${characterName}): channel_id missing — skipping.`);
                p2IdsErrors++;
                continue;
            }

            characterSheet.getRange(row, CHAR_SHEET_CHANNEL_ID_COL).setValue(channelId);
            Logger.log(`Pass 2 Row ${row} (${characterName}): channel_id=${channelId}.`);
            p2IdsAdded++;
            Utilities.sleep(300);
        } catch (e) {
            Logger.log(`Pass 2 Row ${row} (${characterName}): error — ${e}`);
            p2IdsErrors++;
        }
    }

    Logger.log(`Pass 2 complete. IDs added: ${p2IdsAdded}, errors: ${p2IdsErrors}, aborted: ${!!rateLimitAbortAt}`);

    // -------------------------------------------------------------------------
    // PASS 3 — Individual bot channel name fetches for rows that have an ID
    //          but still no name (fallback for anything Pass 1 didn't cover)
    // -------------------------------------------------------------------------
    if (hasBotToken) {
        ss.toast('Pass 3 of 3: filling remaining channel names…', 'Populate Channel Info', 60);
        Logger.log('populateChannelInfoFromWebhooks_: starting Pass 3 (channel name fallback).');

        const pass3Data = characterSheet.getRange(2, 1, numRows, CHAR_SHEET_CHANNEL_NAME_COL).getValues();

        for (let i = 0; i < pass3Data.length; i++) {
            const row           = i + 2;
            const characterName = String(pass3Data[i][CHAR_SHEET_NAME_COL - 1] || '').trim();
            const channelId     = String(pass3Data[i][CHAR_SHEET_CHANNEL_ID_COL - 1] || '').trim();
            const channelName   = String(pass3Data[i][CHAR_SHEET_CHANNEL_NAME_COL - 1] || '').trim();
            const nameIsGood    = channelName !== '' && !channelName.startsWith('(name');

            if (!characterName || !channelId || nameIsGood) continue;

            ss.toast(`Pass 3: ${characterName}…`, 'Populate Channel Info', 10);

            try {
                Utilities.sleep(300);
                const resp = UrlFetchApp.fetch(
                    `https://discord.com/api/v10/channels/${channelId}`,
                    { headers: { 'Authorization': `Bot ${botToken}` }, muteHttpExceptions: true }
                );
                const code = resp.getResponseCode();
                const body = resp.getContentText();

                if (code === 200) {
                    const name = String(JSON.parse(body).name || '').trim();
                    characterSheet.getRange(row, CHAR_SHEET_CHANNEL_NAME_COL).setValue(name);
                    Logger.log(`Pass 3 Row ${row} (${characterName}): name="${name}".`);
                    p3NamesAdded++;
                } else {
                    Logger.log(`Pass 3 Row ${row} (${characterName}): name fetch failed — HTTP ${code}: ${body}`);
                    characterSheet.getRange(row, CHAR_SHEET_CHANNEL_NAME_COL).setValue('(name unavailable)');
                    if (!firstNameError) firstNameError = `HTTP ${code}: ${body.substring(0, 200)}`;
                    p3NamesErrors++;
                }
            } catch (e) {
                Logger.log(`Pass 3 Row ${row} (${characterName}): error — ${e}`);
                p3NamesErrors++;
            }
        }

        Logger.log(`Pass 3 complete. Names added: ${p3NamesAdded}, errors: ${p3NamesErrors}`);
    }

    // -------------------------------------------------------------------------
    // Summary
    // -------------------------------------------------------------------------
    logAudit_('Populate Channel Info', CHARACTER_SHEET_NAME,
        `P1 IDs:${p1IdsAdded} names:${p1NamesAdded} | P2 IDs:${p2IdsAdded} err:${p2IdsErrors} | P3 names:${p3NamesAdded} err:${p3NamesErrors}`);

    let summaryMsg =
        `Pass 1 (guild query):    +${p1IdsAdded} IDs, +${p1NamesAdded} names.\n` +
        `Pass 2 (webhook GETs):   +${p2IdsAdded} IDs, ${p2IdsErrors} errors.\n` +
        `Pass 3 (name fallback):  +${p3NamesAdded} names, ${p3NamesErrors} errors.`;

    if (!hasBotToken || !hasGuildId) {
        summaryMsg += `\n\n⚠️ Pass 1 skipped — set both BOT_TOKEN and GUILD_ID in Script Properties for the fast path.`;
    }
    if (rateLimitAbortAt) {
        summaryMsg += `\n\n⚠️ Pass 2 aborted early: webhook rate limit hit.\nRetry-After: ${rateLimitAbortAt}\nRun again after that time to fetch remaining IDs.`;
    }
    if (firstNameError) {
        summaryMsg += `\n\nFirst name-fetch error:\n${firstNameError}`;
    }

    ui.alert('Populate Channel Info Complete', summaryMsg, ui.ButtonSet.OK);
}