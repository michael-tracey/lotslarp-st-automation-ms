/**
 * @OnlyCurrentDoc
 *
 * Functions related to Discord communication (sending messages, reports).
 */

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
  let allChunksSent = true; // Assume success initially
  for (let chunkIndex = 0; chunkIndex < chunks.length; chunkIndex++) {
      const chunk = chunks[chunkIndex];
      if (chunk.trim() === "") continue; // Skip empty chunks

      let chunkSent = false;
      for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
          Logger.log(`[${context}] Sending chunk ${chunkIndex + 1}/${chunks.length}, Attempt ${attempt + 1}/${MAX_RETRIES + 1}...`);

          // *** Construct payload with optional overrides ***
          const payloadObject = { content: chunk };
          if (overrideUsername) {
              payloadObject.username = overrideUsername;
          }
          if (overrideAvatarUrl) {
              // Basic URL validation (optional but recommended)
              if (overrideAvatarUrl.startsWith('http://') || overrideAvatarUrl.startsWith('https://')) {
                 payloadObject.avatar_url = overrideAvatarUrl;
              } else {
                 Logger.log(`[${context}] Warning: Invalid avatar_url provided, skipping: ${overrideAvatarUrl}`);
              }
          }
          const payload = JSON.stringify(payloadObject);
          // *** End payload construction ***

          const params = {
              method: 'POST',
              contentType: 'application/json',
              payload: payload,
              muteHttpExceptions: true, // Capture errors like 429
          };

          try {
              const response = UrlFetchApp.fetch(targetWebhookUrl, params);
              const responseCode = response.getResponseCode();
              const responseBody = response.getContentText();
              const headers = response.getHeaders();

              if (responseCode >= 200 && responseCode < 300) {
                  Logger.log(`[${context}] Chunk ${chunkIndex + 1} sent successfully (Code: ${responseCode}).`);
                  chunkSent = true;
                  break; // Exit retry loop for this chunk
              } else if (responseCode === 429) {
                  // Rate Limited
                  const retryAfterHeader = headers['Retry-After'] || headers['retry-after']; // Case-insensitive check
                  let waitSeconds = 0;
                  if (retryAfterHeader) {
                      waitSeconds = parseFloat(retryAfterHeader);
                  } else {
                      // Fallback: check JSON body (less reliable for headers)
                      try {
                          const jsonBody = JSON.parse(responseBody);
                          if (jsonBody.retry_after) {
                              waitSeconds = parseFloat(jsonBody.retry_after);
                          }
                      } catch (parseError) { /* Ignore parsing error */ }
                  }

                  // If no specific wait time, use exponential backoff
                  let delayMs;
                  if (waitSeconds > 0) {
                      delayMs = Math.ceil(waitSeconds * 1000) + 500; // Use header value + buffer
                      Logger.log(`[${context}] Rate limited (429). Retrying after header value: ${waitSeconds}s. Waiting ${delayMs}ms.`);
                  } else {
                      delayMs = INITIAL_BACKOFF_MS * Math.pow(2, attempt) + Math.random() * 1000; // Exponential backoff + jitter
                      delayMs = Math.min(delayMs, MAX_BACKOFF_MS); // Cap delay
                      Logger.log(`[${context}] Rate limited (429). No Retry-After header. Using exponential backoff. Waiting ${delayMs}ms.`);
                  }

                  if (attempt < MAX_RETRIES) {
                      Utilities.sleep(delayMs);
                  } else {
                      Logger.log(`[${context}] Rate limited (429). Max retries reached for chunk ${chunkIndex + 1}.`);
                  }
              } else if (responseCode >= 500 && responseCode < 600) {
                  // Server Error - Retry with exponential backoff
                  const delayMs = INITIAL_BACKOFF_MS * Math.pow(2, attempt) + Math.random() * 1000;
                  const cappedDelayMs = Math.min(delayMs, MAX_BACKOFF_MS);
                  Logger.log(`[${context}] Server error (${responseCode}). Attempt ${attempt + 1}. Waiting ${cappedDelayMs}ms before retry.`);
                   if (attempt < MAX_RETRIES) {
                      Utilities.sleep(cappedDelayMs);
                  } else {
                      Logger.log(`[${context}] Server error (${responseCode}). Max retries reached for chunk ${chunkIndex + 1}.`);
                  }
              } else {
                  // Other Client Error (4xx - e.g., 400 Bad Request, 404 Not Found, 401 Unauthorized, 403 Forbidden) - Don't retry these by default
                  Logger.log(`[${context}] Client error (${responseCode}) sending chunk ${chunkIndex + 1}. Not retrying. Body: ${responseBody}`);
                  chunkSent = false; // Mark as failed
                  break; // Exit retry loop for this chunk
              }

          } catch (fetchError) {
              // Network error or other UrlFetchApp issue
              Logger.log(`[${context}] Network or fetch error on attempt ${attempt + 1} for chunk ${chunkIndex + 1}: ${fetchError}`);
              if (attempt < MAX_RETRIES) {
                  const delayMs = INITIAL_BACKOFF_MS * Math.pow(2, attempt) + Math.random() * 1000;
                  const cappedDelayMs = Math.min(delayMs, MAX_BACKOFF_MS);
                  Logger.log(`[${context}] Waiting ${cappedDelayMs}ms before retry due to fetch error.`);
                  Utilities.sleep(cappedDelayMs);
              } else {
                  Logger.log(`[${context}] Max retries reached for chunk ${chunkIndex + 1} after fetch error.`);
              }
          }
      } // End retry loop for one chunk

      if (!chunkSent) {
          Logger.log(`[${context}] Failed to send chunk ${chunkIndex + 1} after all retries.`);
          allChunksSent = false; // Mark overall message as failed
          break; // Stop trying to send remaining chunks for this message
      }

      // Optional delay between chunks if needed
      if (chunks.length > 1 && chunkIndex < chunks.length - 1) {
         Utilities.sleep(200); // Short delay between chunks
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
function handleSendDiscord_(sheet, responseRowIndex, isBulkSend = false) {
    const submissionRowIndex = responseRowIndex - 1;
    if (submissionRowIndex < 1) {
        Logger.log(`Cannot send Discord for row ${responseRowIndex}: No corresponding submission row found.`);
        return;
    }

    const responseRowData = sheet.getRange(responseRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const submissionRowData = sheet.getRange(submissionRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const characterName = responseRowData[CHARACTER_NAME_COL - 1];
    const playerWebhook = getCharacterWebhook_(characterName); // In SheetData.gs
    const testModeActive = isDiscordTestMode_(); // In Utilities.gs
    const sheetName = sheet.getName(); // Get sheet name for logging

    // Determine the target for confirmation message
    let targetDescription = characterName;
    let targetWebhookForSend = playerWebhook; // Default to player webhook

    if (testModeActive) {
        const testWebhookUrl = getTestWebhookUrl_(); // In Utilities.gs
        if (testWebhookUrl && testWebhookUrl !== 'YOUR_TEST_DISCORD_WEBHOOK_URL_HERE') {
            targetDescription = `TEST WEBHOOK (for ${characterName})`;
            targetWebhookForSend = testWebhookUrl; // Use test webhook for sending
        } else {
            targetDescription = `INVALID TEST WEBHOOK (intended for ${characterName})`;
             SpreadsheetApp.getUi().alert(`Cannot send in Test Mode: Test Webhook URL is not set correctly in Script Properties (${PROP_TEST_WEBHOOK}).`);
             Logger.log(`Discord send cancelled for ${characterName} (Row ${responseRowIndex}): Test mode ON but test webhook invalid.`);
             try { sheet.getRange(responseRowIndex, SEND_DISCORD_COL).setValue(false); } catch(e){} // Uncheck box
             return;
        }
    } else if (!playerWebhook) {
         targetDescription = `${characterName} (NO WEBHOOK FOUND)`;
         SpreadsheetApp.getUi().alert(`No valid Discord webhook found for '${characterName}'. Please check the '${CHARACTER_SHEET_NAME}' sheet (Column ${getColumnLetter(CHAR_SHEET_WEBHOOK_COL)}) and add the webhook value.`); // In Utilities.gs
         Logger.log(`Discord send cancelled for ${characterName} (Row ${responseRowIndex}): No player webhook found and test mode OFF.`);
         try { sheet.getRange(responseRowIndex, SEND_DISCORD_COL).setValue(false); } catch(e){} // Uncheck box
         return;
    }

    if (!isBulkSend) {
      const ui = SpreadsheetApp.getUi();
      const confirmMessage = `Send Downtimes to ${targetDescription} via Discord?`;
      const confirm = ui.alert(confirmMessage, ui.ButtonSet.YES_NO);
      if (confirm !== ui.Button.YES) {
          Logger.log(`Discord send cancelled by user for ${targetDescription} (Row ${responseRowIndex}).`);
          try { sheet.getRange(responseRowIndex, SEND_DISCORD_COL).setValue(false); } catch(e){} // Uncheck box
          return false;
      }
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
        Logger.log(`No downtime results found to send for ${characterName} (Row ${responseRowIndex}).`);
        SpreadsheetApp.getUi().alert(`No completed downtime results found for ${characterName} to send.`);
        try { sheet.getRange(responseRowIndex, SEND_DISCORD_COL).setValue(false); } catch(e){} // Uncheck box
        return;
    }

    // --- Send Message ---
    let success = false;
    const logDetailsBase = `Row: ${responseRowIndex}, Char: ${characterName}, Target: ${targetDescription}`;
    try {
        // Pass the determined target webhook (player or test) to the sending function
        success = sendDiscordWebhookMessage_(targetWebhookForSend, title + messageBody, `Downtime Send - ${characterName}`);

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
             // Failure after retries - Alert the user
             Logger.log(`Failed to send Discord message for ${targetDescription} (Row ${responseRowIndex}) after retries.`);
             logAudit_('Sent Discord FAILED', sheetName, `${logDetailsBase}, Check Logs`); // In Utilities.gs
             SpreadsheetApp.getUi().alert(`Failed to send Discord message for ${targetDescription} after multiple attempts. Please check the script logs and try again later.`);
             try { sheet.getRange(responseRowIndex, SEND_DISCORD_COL).setValue(false); } catch(e){} // Uncheck box on final failure
        }

    } catch (error) {
        // Catch errors from sendDiscordWebhookMessage_ (e.g., invalid URL if not caught earlier)
        Logger.log(`Error during Discord send process for ${targetDescription} (Row ${responseRowIndex}): ${error}`);
        logAudit_('Sent Discord ERROR', sheetName, `${logDetailsBase}, Error: ${error.message}`); // In Utilities.gs
        SpreadsheetApp.getUi().alert(`An unexpected error occurred sending the Discord message for ${targetDescription}: ${error.message}`);
        try { sheet.getRange(responseRowIndex, SEND_DISCORD_COL).setValue(false); } catch(e){} // Uncheck box on error
    }
    return success;
}


/** Sends the main downtime progress report to the ST Discord channel. */
function sendDowntimeReportToDiscord() {
  Logger.log('Sending Downtime Report to Discord...');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const results = getDowntimeCompletionData_(sheet); // In SheetData.gs
  const stWebhook = PropertiesService.getScriptProperties().getProperty(PROP_ST_WEBHOOK); // In Constants.gs

  if (!stWebhook || stWebhook === 'YOUR_STORYTELLER_DISCORD_WEBHOOK_URL_HERE') {
    SpreadsheetApp.getUi().alert(`Error: Storyteller Discord Webhook URL is not set in Script Properties (${PROP_ST_WEBHOOK}).`);
    Logger.log(`Discord report failed: ${PROP_ST_WEBHOOK} not set.`);
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
    // This report goes to the ST webhook, but will use Test if enabled
    const success = sendDiscordWebhookMessage_(stWebhook, discordMessage, 'Downtime Report');
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
  const stWebhook = PropertiesService.getScriptProperties().getProperty(PROP_ST_WEBHOOK); // In Constants.gs

  if (!stWebhook || stWebhook === 'YOUR_STORYTELLER_DISCORD_WEBHOOK_URL_HERE') {
    SpreadsheetApp.getUi().alert(`Error: Storyteller Discord Webhook URL is not set in Script Properties (${PROP_ST_WEBHOOK}).`);
    Logger.log(`Missing responses report failed: ${PROP_ST_WEBHOOK} not set.`);
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
    // This report goes to the ST webhook, but will use Test if enabled
    const success = sendDiscordWebhookMessage_(stWebhook, discordMessage, 'Missing Responses Report');
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
 * Sends a test message to the configured Storyteller Discord webhook.
 * This message WILL be sent to the Test Webhook if Test Mode is active.
 */
function sendTestMessageToStorytellers_() {
  Logger.log('Sending test message to ST Discord channel (via helper)...');
  const stWebhook = PropertiesService.getScriptProperties().getProperty(PROP_ST_WEBHOOK); // In Constants.gs
  const sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName(); // Get current sheet for logging context

  if (!stWebhook || stWebhook === 'YOUR_STORYTELLER_DISCORD_WEBHOOK_URL_HERE') {
    SpreadsheetApp.getUi().alert(`Error: Storyteller Discord Webhook URL is not set in Script Properties (${PROP_ST_WEBHOOK}).`);
    Logger.log(`Test message failed: ${PROP_ST_WEBHOOK} not set.`);
    return;
  }

  const message = `👋 This is a test message from the Google Sheet script (${SpreadsheetApp.getActiveSpreadsheet().getName()}) at ${new Date().toLocaleString()}. If you see this, the webhook configuration is being read correctly!`;
  const context = 'ST Test Message';

  try {
    // Use the helper function, passing the ST webhook as the original target
    const success = sendDiscordWebhookMessage_(stWebhook, message, context);

    if (success) {
        // Determine where it was actually sent for the alert message
        const finalDestination = isDiscordTestMode_() ? "Test Webhook" : "ST Channel"; // In Utilities.gs
        logAudit_('Sent Test Message', sheetName, `Destination: ${finalDestination}`); // Log success // In Utilities.gs
        SpreadsheetApp.getUi().alert(`Test message sent successfully to the ${finalDestination}!`);
        Logger.log(`Test message sent successfully via helper to ${finalDestination}.`);
    } else {
         logAudit_('Sent Test Message FAILED', sheetName, `Check Logs`); // Log failure // In Utilities.gs
         SpreadsheetApp.getUi().alert(`Failed to send Test message to Discord after retries. Check Logs.`);
         Logger.log(`Failed to send Test message to Discord after retries.`);
    }
  } catch (error) {
    Logger.log(`Error sending test message via helper: ${error}\nStack: ${error.stack}`);
    logAudit_('Sent Test Message ERROR', sheetName, `Error: ${error.message}`); // Log error // In Utilities.gs
    SpreadsheetApp.getUi().alert(`Error sending test message: ${error.message}`);
  }
}

/** Counts approved characters and sends the count to the ST Discord channel. */
function checkCharacterCount_() {
    Logger.log('Checking approved character count...');
    const sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName(); // Get current sheet for logging context
    const characterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHARACTER_SHEET_NAME); // In Constants.gs

    if (!characterSheet) {
        SpreadsheetApp.getUi().alert(`Error: Cannot find the '${CHARACTER_SHEET_NAME}' sheet.`);
        Logger.log(`Character count check failed: Sheet '${CHARACTER_SHEET_NAME}' not found.`);
        return;
    }

    const stWebhook = PropertiesService.getScriptProperties().getProperty(PROP_ST_WEBHOOK); // In Constants.gs
     if (!stWebhook || stWebhook === 'YOUR_STORYTELLER_DISCORD_WEBHOOK_URL_HERE') {
        SpreadsheetApp.getUi().alert(`Error: Storyteller Discord Webhook URL is not set in Script Properties (${PROP_ST_WEBHOOK}). Cannot send report.`);
        Logger.log(`Character count check failed: ${PROP_ST_WEBHOOK} not set.`);
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
        // Send to ST webhook, allowing test mode override
        const success = sendDiscordWebhookMessage_(stWebhook, message, 'Character Count');
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
        logAudit_('Character Count Check ERROR', sheetName, `Error: ${error.message}`); // In Utilities.gs
        SpreadsheetApp.getUi().alert(`Error checking character count: ${error.message}`);
    }
}