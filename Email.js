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
    // Consider making the recipient email a script property
    GmailApp.sendEmail('avllotslarp@gmail.com', subject, '', { htmlBody: htmlBody });
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

    // --- Prompt for Email Address ---
    const ui = SpreadsheetApp.getUi();
    const emailPrompt = ui.prompt(`Enter email address to send ${characterName}'s downtime results to:`, ui.ButtonSet.OK_CANCEL);

    if (emailPrompt.getSelectedButton() !== ui.Button.OK) {
        Logger.log(`Email send cancelled by user for ${characterName} (Row ${responseRowIndex}).`);
        // Only uncheck if triggered by checkbox (manual doesn't have 'e')
        try {
             const checkboxCell = sheet.getRange(responseRowIndex, SEND_EMAIL_COL);
             if (checkboxCell.isChecked()) { // Check current state before unchecking
                 checkboxCell.setValue(false);
             }
        } catch(err) { Logger.log(`Could not uncheck email box on cancel: ${err}`);}
        return;
    }

    const recipientEmail = emailPrompt.getResponseText().trim();
    // Basic email validation
    if (!recipientEmail || !/^\S+@\S+\.\S+$/.test(recipientEmail)) {
        ui.alert(`Invalid email address provided: "${recipientEmail}". Please try again.`);
        Logger.log(`Email send cancelled for ${characterName} (Row ${responseRowIndex}): Invalid email "${recipientEmail}".`);
         try {
             const checkboxCell = sheet.getRange(responseRowIndex, SEND_EMAIL_COL);
             if (checkboxCell.isChecked()) { checkboxCell.setValue(false); }
        } catch(err) { Logger.log(`Could not uncheck email box on invalid address: ${err}`);}
        return;
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

    const logDetailsBase = `Row: ${responseRowIndex}, Char: ${characterName}, To: ${recipientEmail}`;

    // --- Send Email ---
    try {
        GmailApp.sendEmail(recipientEmail, subject, '', { htmlBody: htmlBody });

        // --- Update Sheet on Success ---
        const statusCell = sheet.getRange(responseRowIndex, STATUS_COL);
        const timestampCell = sheet.getRange(responseRowIndex, TIMESTAMP_COL); // Use response row timestamp col
        const checkboxCell = sheet.getRange(responseRowIndex, SEND_EMAIL_COL);

        timestampCell.setValue(new Date()).setBackgroundRGB(229, 255, 204);
        statusCell.setValue('sent');
        checkboxCell.setBackgroundRGB(229, 255, 204);

        Logger.log(`Successfully sent downtime results for ${characterName} to ${recipientEmail}.`);
        logAudit_('Sent Email', sheetName, logDetailsBase); // Log success
        ui.alert(`Email sent successfully to ${recipientEmail}.`);

    } catch (error) {
        Logger.log(`Error sending email for ${characterName} (Row ${responseRowIndex}) to ${recipientEmail}: ${error}`);
        logAudit_('Sent Email FAILED', sheetName, `${logDetailsBase}, Error: ${error.message}`); // Log failure
        SpreadsheetApp.getUi().alert(`Failed to send email to ${recipientEmail}: ${error.message}`);
         try {
             const checkboxCell = sheet.getRange(responseRowIndex, SEND_EMAIL_COL);
             if (checkboxCell.isChecked()) { checkboxCell.setValue(false); }
        } catch(err) { Logger.log(`Could not uncheck email box on failure: ${err}`);}
    }
}
