/**
 * @OnlyCurrentDoc
 *
 * Functions related to sending emails.
 */

/**
 * Sends an email notification about a new form submission.
 * @param {string} email Submitter's email.
 * @param {Date} timestamp Submission timestamp.
 * @param {Array<GoogleAppsScript.Forms.ItemResponse>} itemResponses Array of responses.
 */
function sendFormSubmissionEmail_(email, timestamp, itemResponses) {
  const subject = `New Downtime Form Submitted - ${email}`;
  let htmlBody = `<h2>New Downtime Submission</h2>
                  <p><b>Submitted by:</b> ${email}</p>
                  <p><b>Timestamp:</b> ${timestamp.toLocaleString()}</p>
                  <h3>Responses:</h3><ul>`;

  itemResponses.forEach(r => {
    htmlBody += `<li><b>${escapeHtml_(r.getItem().getTitle())}</b>: ${escapeHtml_(r.getResponse() || '(empty)')}</li>\n`;
  });
  htmlBody += "</ul>";

  try {
    GmailApp.sendEmail(getNotificationEmail_(), subject, '', { htmlBody: htmlBody }); // Recipient configurable via NOTIFICATION_EMAIL property
    Logger.log(`Sent form submission notification email for ${email}.`);
  } catch (error) {
      Logger.log(`Error sending form submission email for ${email}: ${error}`);
      // Optionally notify STs via Discord if email fails
  }
}

/**
 * Handles sending downtime results via Email when the checkbox is checked or manually triggered.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet.
 * @param {number} responseRowIndex The row index of the response row (where checkbox was checked).
 */
function handleSendEmail_(sheet, responseRowIndex) {
    const submissionRowIndex = responseRowIndex - 1;
     if (submissionRowIndex < 1) {
        Logger.log(`Cannot send Email for row ${responseRowIndex}: No corresponding submission row found.`);
        return;
    }

    const responseRowData = sheet.getRange(responseRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const submissionRowData = sheet.getRange(submissionRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const characterName = responseRowData[CHARACTER_NAME_COL - 1];
    const sheetName = sheet.getName(); // Get sheet name for logging

    // --- Look up the recipient's email from the Characters sheet (col AA) ---
    // Pre-fills the preview dialog's "To:" field. If no email is on file it
    // starts blank; either way the address stays editable before sending.
    const recipientEmail = getCharacterEmail_(characterName); // In SheetData.js
    if (recipientEmail) {
        Logger.log(`Pre-filled email for ${characterName}: ${recipientEmail}`);
    } else {
        Logger.log(`No email on file for ${characterName}; preview "To:" field will start blank.`);
    }

    // --- Construct Message Body ---
    const year = PropertiesService.getScriptProperties().getProperty(PROP_DOWNTIME_YEAR);
    const month = PropertiesService.getScriptProperties().getProperty(PROP_DOWNTIME_MONTH);
    const subject = `Downtime Results for ${characterName} (${month}, ${year})`;
    let htmlBody = `<h2>${subject}</h2><hr>`;

    let contentFound = false;
    for (let j = CHARACTER_NAME_COL; j < headers.length; j++) {
        const header = headers[j];
        const submissionText = submissionRowData[j] ? String(submissionRowData[j]).trim() : "";
        const responseText = responseRowData[j] ? String(responseRowData[j]).trim() : "";

        if (responseText !== "") {
            contentFound = true;
            htmlBody += `<h3>${escapeHtml_(header)}</h3>`;
            if (submissionText !== "") {
                // Basic newline to <br> conversion for readability
                htmlBody += `<p><b>Your Action:</b><br>${escapeHtml_(submissionText).replace(/\n/g, '<br>')}</p>`;
            } else {
                 htmlBody += `<p><b>Your Action:</b><br>(No submission text found)</p>`;
            }
            htmlBody += `<p><b>Result:</b><br>${escapeHtml_(responseText).replace(/\n/g, '<br>')}</p><hr>`;
        } else if (submissionText !== "") {
            // Optionally include pending actions
            // contentFound = true;
            // htmlBody += `<h3>${escapeHtml_(header)}</h3><p><b>Your Action:</b><br>${escapeHtml_(submissionText).replace(/\n/g, '<br>')}</p><p><b>Result:</b><br>(Pending)</p><hr>`;
        }
    }

     if (!contentFound) {
        Logger.log(`No downtime results found to email for ${characterName} (Row ${responseRowIndex}).`);
        SpreadsheetApp.getUi().alert(`No completed downtime results found for ${characterName} to send via email.`);
        try {
             const checkboxCell = sheet.getRange(responseRowIndex, SEND_EMAIL_COL);
             if (checkboxCell.isChecked()) { checkboxCell.setValue(false); }
        } catch(err) { Logger.log(`Could not uncheck email box on no content: ${err}`);}
        return;
    }

    // Show preview dialog for last-minute editing before sending
    const template = HtmlService.createTemplateFromFile('EmailPreviewDialog');
    template.rowIndex = responseRowIndex;
    template.recipientEmail = recipientEmail;
    template.subject = subject;
    template.htmlBody = htmlBody;
    const html = template.evaluate().setWidth(700).setHeight(560);
    SpreadsheetApp.getUi().showModalDialog(html, `Preview: ${characterName}'s Downtime Email`);
}

/**
 * Called by EmailPreviewDialog when the user confirms. Sends the (possibly edited) email.
 */
function sendPreparedEmail_(rowIndex, recipientEmail, subject, htmlBody) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const responseRowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const characterName = responseRowData[CHARACTER_NAME_COL - 1];
    const sheetName = sheet.getName();
    const logDetailsBase = `Row: ${rowIndex}, Char: ${characterName}, To: ${recipientEmail}`;

    try {
        GmailApp.sendEmail(recipientEmail, subject, '', { htmlBody: htmlBody });

        const statusCell = sheet.getRange(rowIndex, STATUS_COL);
        const timestampCell = sheet.getRange(rowIndex, TIMESTAMP_COL);
        const checkboxCell = sheet.getRange(rowIndex, SEND_EMAIL_COL);

        timestampCell.setValue(new Date()).setBackgroundRGB(229, 255, 204);
        statusCell.setValue('sent');
        checkboxCell.setBackgroundRGB(229, 255, 204);

        Logger.log(`Successfully sent downtime results for ${characterName} to ${recipientEmail}.`);
        logAudit_('Sent Email', sheetName, logDetailsBase);
    } catch (error) {
        Logger.log(`Error sending email for ${characterName} (Row ${rowIndex}) to ${recipientEmail}: ${error}`);
        logAudit_('Sent Email FAILED', sheetName, `${logDetailsBase}, Error: ${error.message}`);
        throw error;
    }
}

/**
 * Called by EmailPreviewDialog when the user cancels. Unchecks the Send Email checkbox.
 */
function cancelEmailSend_(rowIndex) {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const checkboxCell = sheet.getRange(rowIndex, SEND_EMAIL_COL);
        if (checkboxCell.isChecked()) { checkboxCell.setValue(false); }
    } catch(err) {
        Logger.log(`Could not uncheck email box on cancel: ${err}`);
    }
}
