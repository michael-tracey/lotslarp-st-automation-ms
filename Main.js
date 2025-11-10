/**
 * @OnlyCurrentDoc
 *
 * Main script file containing onOpen, onEdit, onFormSubmit handlers,
 * and top-level menu item functions.
 */

// ============================================================================ 
// === SCRIPT INITIALIZATION & MENU ========================================= 
// ============================================================================ 

/**
 * Runs when the spreadsheet is opened. Sets up script properties (if needed)
 * and creates the custom menu with reorganized items.
 */
function onOpen() {
  // --- Set default properties ---
  const scriptProperties = PropertiesService.getScriptProperties();
  const propertiesToInitialize = {
    [PROP_ST_WEBHOOK]: 'YOUR_STORYTELLER_DISCORD_WEBHOOK_URL_HERE', // Replace if needed
    [PROP_DOWNTIME_FORM_ID]: 'YOUR_GOOGLE_FORM_ID_HERE', // Replace if needed
    [PROP_DOWNTIME_YEAR]: new Date().getFullYear().toString(),
    [PROP_DOWNTIME_MONTH]: new Date().toLocaleString('default', { month: 'long' }),
    [PROP_PDF_FOLDER_ID]: 'YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE', // Replace if needed
    [PROP_DISCORD_TEST_MODE]: 'false', // Default test mode to off
    [PROP_TEST_WEBHOOK]: 'YOUR_TEST_DISCORD_WEBHOOK_URL_HERE', // Replace if needed
    [PROP_IC_NEWS_FEED_WEBHOOK]: 'PROP_IC_NEWS_FEED_WEBHOOK',
    [PROP_LARP_NAME]: 'My Awesome LARP',
    [PROP_TASK_COLOR_HEX]: '#FFFF00' // Default color for 'Any Narrator or Storyteller'
  };

  let propertiesUpdated = false;
  for (const key in propertiesToInitialize) {
    if (!scriptProperties.getProperty(key)) {
      scriptProperties.setProperty(key, propertiesToInitialize[key]);
      Logger.log(`Property ${key} was not set. Initialized with default/placeholder.`);
      propertiesUpdated = true;
    }
  }
  if (propertiesUpdated) {
      Logger.log('One or more script properties were initialized. Please verify placeholder values in Project Settings > Script properties.');
      // Optionally alert the user if critical placeholders were set
      if (!scriptProperties.getProperty(PROP_ST_WEBHOOK) || scriptProperties.getProperty(PROP_ST_WEBHOOK).includes('YOUR_')) {
          SpreadsheetApp.getUi().alert('Please configure the Storyteller Discord Webhook URL in Project Settings > Script properties.');
      }
       if (!scriptProperties.getProperty(PROP_ANNOUNCEMENT_WEBHOOK) || scriptProperties.getProperty(PROP_ANNOUNCEMENT_WEBHOOK).includes('YOUR_')) {
          SpreadsheetApp.getUi().alert('Please configure the Announcement Discord Webhook URL in Project Settings > Script properties.');
      }
       if (!scriptProperties.getProperty(PROP_IC_CHAT_WEBHOOK) || scriptProperties.getProperty(PROP_IC_CHAT_WEBHOOK).includes('YOUR_')) {
          SpreadsheetApp.getUi().alert('Please configure the IC Chat Discord Webhook URL in Project Settings > Script properties.');
      }
       if (!scriptProperties.getProperty(PROP_IC_NEWS_FEED_WEBHOOK) || scriptProperties.getProperty(PROP_IC_NEWS_FEED_WEBHOOK).includes('YOUR_')) {
          SpreadsheetApp.getUi().alert('Please configure the IC News Feed Discord Webhook URL in Project Settings > Script properties.');
      }
       if (!scriptProperties.getProperty(PROP_TEST_WEBHOOK) || scriptProperties.getProperty(PROP_TEST_WEBHOOK).includes('YOUR_')) {
          SpreadsheetApp.getUi().alert('Please configure the Test Discord Webhook URL in Project Settings > Script properties if you plan to use test mode.');
      }
       if (!scriptProperties.getProperty(PROP_TASK_COLOR_HEX) || scriptProperties.getProperty(PROP_TASK_COLOR_HEX).includes('#')) {
          SpreadsheetApp.getUi().alert('Please configure the TASK_COLOR_HEX in Project Settings > Script properties if you want to change the default color for "Any Staff".');
      }
  }


  // --- Create Menu (Reorganized) ---
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Storyteller Menu');

  menu.addItem('Downtime Editor popup', 'openEditCellDialog'); // In Dialogs.gs
  menu.addItem('Show Downtime Progress', 'showDowntimeProgressDialog'); // In Dialogs.gs
  menu.addItem('Show Downtimes Progress by Staff', 'showDowntimeProgressByStaffDialog'); // In Dialogs.gs
  menu.addSeparator();

  // --- Dynamic Narrator Assignment Menu ---
  const narrators = getNarrators_(); // From SheetData.js
  if (narrators && narrators.length > 0) {
    const assignToMenu = ui.createMenu('Assign to');
    const MAX_NARRATORS_IN_MENU = 50;
    const taskColorHex = scriptProperties.getProperty(PROP_TASK_COLOR_HEX) || '#FFFF00'; // Default to yellow if not set

    // Add 'Any Staff' option
    assignToMenu.addItem(`Any Staff [${taskColorHex}]`, 'assignAnyStaff_');

    // Sort narrators alphabetically
    narrators.sort((a, b) => a.name.localeCompare(b.name));

    narrators.forEach((narrator, index) => {
      if (index < MAX_NARRATORS_IN_MENU) {
        assignToMenu.addItem(`${narrator.name} [${narrator.color}]`, `assignNarrator${index + 1}`);
      }
    });

    if (narrators.length > MAX_NARRATORS_IN_MENU) {
      ui.alert(`Warning: Only the first ${MAX_NARRATORS_IN_MENU} narrators are shown in the 'Assign to' menu. Please increase MAX_NARRATORS_IN_MENU in Main.js if needed.`);
      Logger.log(`Warning: More narrators (${narrators.length}) than available menu slots (${MAX_NARRATORS_IN_MENU}).`);
    }
    menu.addSubMenu(assignToMenu);
  } else {
    Logger.log('No narrators found or "Narrators" sheet is empty/missing. "Assign to" menu not created.');
  }

  const downtimeMenu = ui.createMenu('Downtime Management');
  downtimeMenu.addItem('Bulk Send Downtimes to Discord', 'showBulkSendDowntimesDialog'); // In BulkSendDowntimes.js
  downtimeMenu.addItem('Show Missing Downtime Responses', 'showMissingDowntimeDialog'); // In Dialogs.gs
  downtimeMenu.addItem('Find Pending Responses by Color', 'showFindPendingByColorDialog'); // In Dialogs.gs
  menu.addSubMenu(downtimeMenu);

  const influenceMenu = ui.createMenu('Influence & Resources');
  influenceMenu.addItem('Show Detailed Influence Summary', 'showDetailedInfluenceSummaryDialog_');
  influenceMenu.addItem('Show Influences Progress', 'showInfluencesProgressDialog'); // In Dialogs.gs
  influenceMenu.addItem('Show Resources Progress', 'showResourcesProgressDialog'); // In Dialogs.gs
  influenceMenu.addItem('Fill Cell with Influences (G&IT, WotS)', 'openInfluencePowerDialog_');
  menu.addSubMenu(influenceMenu);

  const fillCellsMenu = ui.createMenu('Fill Cells');
  fillCellsMenu.addItem('Fill cell with Feed Data', 'fillCellWithFeedData_');
  fillCellsMenu.addItem('Fill cell with Herd Data', 'fillCellWithHerdData_');
  fillCellsMenu.addItem('Fill cell with Patrol Data', 'fillCellWithPatrolData_'); 
  fillCellsMenu.addItem('Fill cell with Discipline Data','fillCellWithDisciplineData_')
  menu.addSubMenu(fillCellsMenu);

  const discordMenu = ui.createMenu('Discord Actions');
  discordMenu.addItem('Send Downtime Report to Discord', 'sendDowntimeReportToDiscord'); // In Discord.gs
  discordMenu.addItem('Send Missing Downtime Responses to Discord', 'sendMissingDowntimeResponsesToDiscord'); // In Discord.gs
  discordMenu.addItem('Test Discord to ST Channel', 'sendTestMessageToStorytellers_'); // In Discord.gs
  menu.addSubMenu(discordMenu);

  const maintenanceMenu = ui.createMenu('Maintenance');
  maintenanceMenu.addItem('Open Web App', 'openWebApp_');
  maintenanceMenu.addItem('Encrypt Narrator Password', 'showEncryptPasswordDialog_');
  maintenanceMenu.addItem('Manage Script Properties', 'openScriptPropertiesDialog_'); // In ScriptPropertiesManager.gs
  maintenanceMenu.addItem('Re-import Downtime from Form', 'showReimportDowntimeDialog_');
  maintenanceMenu.addItem('Manually Send Discord for Selected Row', 'manuallySendDiscord_'); // Wrapper below
  maintenanceMenu.addItem('Manually Send Email for Selected Row', 'manuallySendEmail_'); // Wrapper below
  maintenanceMenu.addItem('Manually Test Usernames', 'manuallyTestCharacterNames_'); // Wrapper below
  maintenanceMenu.addItem('Check Character Count', 'checkCharacterCount_'); // In Discord.gs
  maintenanceMenu.addItem('Run Permission Test', 'testAllFeaturesAndPermissions'); // In PermissionTester.js
  const triggerMenu = ui.createMenu('Reinstall Triggers');
  triggerMenu.addItem('Reinstall Form Trigger', 'reinstallFormTrigger_'); // In Triggers.gs
  triggerMenu.addItem('Reinstall Edit Trigger', 'reinstallOnEditTrigger_'); // In Triggers.gs
  triggerMenu.addItem('Reinstall Scheduled Message Trigger', 'setupScheduledMessagesTrigger_')
  maintenanceMenu.addSubMenu(triggerMenu);
  menu.addSubMenu(maintenanceMenu);

  menu.addSeparator();
  const isTestMode = isDiscordTestMode_(); // In Utilities.gs
  const testModeLabel = `${isTestMode ? '✅ ' : ''}Toggle Discord Test Mode`;
  menu.addItem(testModeLabel, 'toggleDiscordTestMode_'); // Wrapper below

  // Finalize Menu
  menu.addToUi();

  Logger.log(`Custom menu created. Discord Test Mode is currently ${isTestMode ? 'ON' : 'OFF'}.`);
}

// ============================================================================ 
// === EVENT HANDLERS (onEdit, onFormSubmit) ================================ 
// ============================================================================ 
// These remain in the main file as they are directly called by triggers.

/**
 * Handles the 'onEdit' event trigger for the spreadsheet.
 * Monitors changes to specific columns (Send Discord, Send Email, Character Name, Status).
 * Adds confirmation for actions on older sheets.
 * Email send ignores status check, Discord send respects it.
 * Alerts user if trying to send Discord when status is 'sent'.
 * Validates character name on ODD rows (response rows).
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object.
 */
function onEditHandler_(e) {
  // Check if the event object and range exist.
  if (!e || !e.range) {
    return; // Exit quietly for non-user edits
  }

  const sheet = e.range.getSheet();
  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();
  const sheetName = sheet.getName();

  // --- Basic Checks ---
  if (editedRow <= 1) return; // Is it the header row?
  const isMonthYearSheet = MONTH_YEAR_SHEET_REGEX.test(sheetName);
  if (!isMonthYearSheet) return; // Is it a sheet named like "Month Year"?

  // Log the entry point of the edit
  Logger.log(`onEdit: Sheet='${sheetName}', Cell='${e.range.getA1Notation()}', User='${e.user ? e.user.getEmail() : 'N/A'}', NewValue='${e.value}'`);

  try {
    // --- Get Status Value (often needed) ---
    // Status is on the response row (odd). If editing an even row, check status of the row below.
    const statusRowIndex = (editedRow % 2 !== 0) ? editedRow : editedRow + 1;
    const statusValue = (statusRowIndex <= sheet.getLastRow()) ? sheet.getRange(statusRowIndex, STATUS_COL).getValue() : '';

    // ======================================================================== 
    // === ACTION HANDLERS ==================================================== 
    // ======================================================================== 

    // --- Action: Send via Discord Checkbox ---
    if (editedCol === SEND_DISCORD_COL && editedRow % 2 !== 0) {
      const isChecked = e.range.isChecked();
      Logger.log(`Discord Checkbox edit on row ${editedRow}. Checked: ${isChecked}, Status: '${statusValue}'`);
      if (isChecked === true && statusValue !== 'sent') {
        // Confirmation for Old Sheets
        if (sheetName !== getDowntimeSheetName_()) {
          const ui = SpreadsheetApp.getUi();
          const response = ui.alert('Confirmation Needed', `You are triggering Discord send on an OLDER month's sheet ('${sheetName}').\n\nAre you sure?`, ui.ButtonSet.YES_NO);
          if (response !== ui.Button.YES) {
            e.range.setValue(false); // Uncheck on cancel
            return;
          }
        }
        handleSendDiscord_(sheet, editedRow); // In Discord.gs
      } else if (isChecked === true && statusValue === 'sent') {
        SpreadsheetApp.getUi().alert('Action Blocked', 'Cannot send to Discord because the status for this row is already "sent".\n\nTo resend, please change the status back to "unprocessed" or use the "Manually Send Discord for Selected Row" option under the Maintenance menu.', SpreadsheetApp.getUi().ButtonSet.OK);
        e.range.setValue(false); // Uncheck the box
      }
    }

    // --- Action: Send via Email Checkbox ---
    else if (editedCol === SEND_EMAIL_COL && editedRow % 2 !== 0) {
      const isChecked = e.range.isChecked();
      Logger.log(`Email Checkbox edit on row ${editedRow}. Checked: ${isChecked}`);
      if (isChecked === true) {
        if (sheetName !== getDowntimeSheetName_()) {
          const ui = SpreadsheetApp.getUi();
          const response = ui.alert('Confirmation Needed', `You are triggering Email send on an OLDER month's sheet ('${sheetName}').\n\nAre you sure?`, ui.ButtonSet.YES_NO);
          if (response !== ui.Button.YES) {
            e.range.setValue(false); // Uncheck on cancel
            return;
          }
        }
        handleSendEmail_(sheet, editedRow); // In Email.gs
      }
    }

    // --- Action: Character Name Validation ---
    else if (editedCol === CHARACTER_NAME_COL && editedRow % 2 !== 0) {
      Logger.log(`Character name edited in cell ${e.range.getA1Notation()}. Validating.`);
      validateCharacterNameCell_(e.range); // In SheetData.gs
    }

    // --- Action: Log Manual Status Changes ---
    else if (editedCol === STATUS_COL && editedRow % 2 !== 0 && e.oldValue !== e.value) {
      Logger.log(`Status manually changed in cell ${e.range.getA1Notation()}.`);
      logAudit_('Manual Status Change', sheetName, `Cell: ${e.range.getA1Notation()}, Old: ${e.oldValue}, New: ${e.value}`); // In Utilities.gs
    }

    // --- Action: Add Note to Edited Response Cell ---
    else {
      const isResponseRow = editedRow % 2 !== 0;
      const isResponseColumn = editedCol > CHARACTER_NAME_COL;

      // Check if it's a meaningful edit in a response cell
      if (isResponseRow && isResponseColumn && e.value !== e.oldValue) {
        const userEmail = e.user ? e.user.getEmail() : null;
        if (userEmail) {
          const userName = getNarratorNameByEmail_(userEmail) || userEmail;
          const note = `Last edited by: ${userName} on ${new Date().toLocaleString()}`;
          e.range.setNote(note);
          Logger.log(`Set note on ${e.range.getA1Notation()} for user ${userName}`);
        }
      }
    }

  } catch (error) {
    Logger.log(`Error in onEditHandler_ for range ${e.range.getA1Notation()}: ${error} 
Stack: ${error.stack}`);
    // Try to uncheck box on error if it was a checkbox action that was proceeding
    if ((editedCol === SEND_DISCORD_COL || editedCol === SEND_EMAIL_COL)) {
      try { e.range.setValue(false); } catch(err) {}
    }
  }
}


/**
 * Handles the 'onFormSubmit' event trigger for the Google Form.
 * Processes the form response and adds it to the downtime sheet.
 *
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e The event object.
 */
function onFormSubmitHandler_(e) {
  // Add log message at the very beginning
  Logger.log(`onFormSubmitHandler_ started. Event Source: ${e ? e.source : 'N/A'}, Response ID: ${e && e.response ? e.response.getId() : 'N/A'}`);

  if (!e || !e.response) {
    Logger.log('onFormSubmitHandler_ called without valid event object or response.');
    return; // Exit if event object is not valid
  }

  const formResponse = e.response;
  const itemResponses = formResponse.getItemResponses();
  const email = formResponse.getRespondentEmail();
  const timestamp = formResponse.getTimestamp();
  const formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");

  Logger.log(`Form submitted by '${email}' at ${timestamp}. Processing ${itemResponses.length} responses.`);

  try {
    // --- Send Email Notification to STs ---
    sendFormSubmissionEmail_(email, timestamp, itemResponses); // In Email.gs

    // --- Add Data to Sheet ---
    const sheet = getOrCreateDowntimeSheet_(formResponse); // In Utilities.gs
    const characterNameResponse = itemResponses.length > 0 ? itemResponses[0].getResponse() : 'Unknown Character'; // Assuming first question is Character Name

    // Log Form Submission
    logAudit_('Form Submission', sheet.getName(), `Submitter: ${email}, Character: ${characterNameResponse}`); // In Utilities.gs


    // --- Append Submission Row ---
    const submissionRowData = [formattedTimestamp, 'input', '', '']; // Timestamp, Status, Send Discord, Send Email
    itemResponses.forEach(response => submissionRowData.push(response.getResponse()));
    sheet.appendRow(submissionRowData);
    const submissionRowIndex = sheet.getLastRow();
    Logger.log(`Appended submission data to row ${submissionRowIndex}.`);

    // --- Append Response Row ---
    const responseRowData = ['', 'unprocessed', false, false, characterNameResponse]; // Blank timestamp/status, checkboxes, character name
    // Add blank cells for other response columns
    for (let j = CHARACTER_NAME_COL; j < submissionRowData.length; j++) { // Start from CHARACTER_NAME_COL
        responseRowData.push('');
    }
    sheet.appendRow(responseRowData);
    const responseRowIndex = sheet.getLastRow();
    Logger.log(`Appended response template row ${responseRowIndex}.`);

    // --- Format Response Row ---
    sheet.getRange(responseRowIndex, SEND_DISCORD_COL).insertCheckboxes();
    sheet.getRange(responseRowIndex, SEND_EMAIL_COL).insertCheckboxes();
    // Set default background for empty response cells (excluding name column)
    const firstResponseCol = CHARACTER_NAME_COL + 1;
    if (sheet.getLastColumn() >= firstResponseCol) {
        sheet.getRange(responseRowIndex, firstResponseCol, 1, sheet.getLastColumn() - CHARACTER_NAME_COL).setBackgroundRGB(255, 204, 204); // Light red
    }

    // --- Validate Character Name and Set Background ---
    const characterNameCell = sheet.getRange(responseRowIndex, CHARACTER_NAME_COL);
    Logger.log(`Validating character name cell: ${characterNameCell.getA1Notation()}`); // Added log
    validateCharacterNameCell_(characterNameCell); // In SheetData.gs

    // --- Apply Conditional Formatting for Responses ---
     if (sheet.getLastColumn() >= firstResponseCol) {
        const responseValueRange = sheet.getRange(responseRowIndex, firstResponseCol, 1, sheet.getLastColumn() - CHARACTER_NAME_COL);
        applyConditionalFormatting_(sheet, responseValueRange); // In SheetData.gs
     }

    Logger.log(`Successfully processed form submission for ${characterNameResponse}.`);

  } catch (error) {
    Logger.log(`Error processing form submission from ${email}: ${error} 
Stack: ${error.stack}`);
    // Notify STs about the failure
    try {
        const stWebhook = PropertiesService.getScriptProperties().getProperty(PROP_ST_WEBHOOK);
        if (stWebhook && stWebhook !== 'YOUR_STORYTELLER_DISCORD_WEBHOOK_URL_HERE') {
            // Use sendDiscordWebhookMessage_ helper for potential test mode override if desired for error reports too
            const errorSent = sendDiscordWebhookMessage_(stWebhook, `🚨 **Error processing form submission!**\nSubmitter: ${email}\nError: ${error.message}\nCheck script logs for details.`, 'Form Submit Error'); // In Discord.gs
            if (!errorSent) { Logger.log("Failed to send form submission error report to Discord."); }
        }
        GmailApp.sendEmail('avllotslarp@gmail.com', `Error Processing Downtime Submission - ${email}`, `An error occurred while processing the downtime submission from ${email}:\n\n${error}\n\n${error.stack}`);
    } catch (notifyError) {
        Logger.log(`Failed to send error notification: ${notifyError}`);
    }
  }
}


// ============================================================================ 
// === MENU ITEM WRAPPERS =================================================== 
// ============================================================================ 
// These functions simply call the real logic located in other files.

function openWebApp_() {
  const url = ScriptApp.getService().getUrl();
  const html = `<script>window.open('${url}', '_blank');google.script.host.close();</script>`;
  const userInterface = HtmlService.createHtmlOutput(html).setHeight(10).setWidth(100);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Opening Web App...');
}

// --- Fill Cell Wrappers ---
function fillCellWithFeedData_() { fillCellWithData_('feed'); } // In Actions.gs
function fillCellWithHerdData_() { fillCellWithData_('herd'); } // In Actions.gs
function fillCellWithPatrolData_() { fillCellWithData_('patrol'); } // In Actions.gs
function fillCellWithDisciplineData_() { fillCellWithData_('discipline'); } // In Actions.gs
function fillCellWithEliteInfluenceData_() { fillCellWithInfluenceData_(ELITE_INFLUENCES_SHEET_NAME, 'Elite'); } // In Actions.gs
function fillCellWithUnderworldInfluenceData_() { fillCellWithInfluenceData_(UW_INFLUENCES_SHEET_NAME, 'Underworld'); } // In Actions.gs

/**
 * Assigns 'Any Staff' to the active cell and sets its background color.
 */
function assignAnyStaff_() {
  const ui = SpreadsheetApp.getUi();
  const activeRange = SpreadsheetApp.getActiveRange();
  if (!activeRange) {
    ui.alert('Please select a cell to assign.');
    return;
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  const taskColorHex = scriptProperties.getProperty(PROP_TASK_COLOR_HEX) || '#FFFF00'; // Default to yellow if not set

  activeRange.setBackground(taskColorHex);
  if (isColorDark(taskColorHex)) {
    activeRange.setFontColor('#FFFFFF'); // White font for dark backgrounds
  } else {
    activeRange.setFontColor('#000000'); // Black font for light backgrounds
  }
  Logger.log(`Assigned 'Any Staff' to ${activeRange.getA1Notation()} with color ${taskColorHex}`);
}

// --- Maintenance Wrappers ---
function showReimportDowntimeDialog_() { showReimportDowntimeDialog(); } // In ReimportDowntime.js
function manuallyTestCharacterNames_() { _manuallyTestCharacterNamesInternal_(); } // In SheetData.gs (renamed internal)
function manuallySendDiscord_() { _manuallySendDiscordInternal_(); } // Wrapper below
function manuallySendEmail_() { _manuallySendEmailInternal_(); } // Wrapper below
function fetchDowntimes_() { _fetchDowntimesInternal_(); } // Wrapper below
function toggleDiscordTestMode_() { _toggleDiscordTestModeInternal_(); } // Wrapper below

// --- Placeholder/Simple Maintenance Actions ---
function _fetchDowntimesInternal_() {
  Logger.log('Placeholder function fetchDowntimes_ called.');
  SpreadsheetApp.getUi().alert('Fetch Downtimes function is not yet implemented.');
  // TODO: Implement logic if this is a required feature.
}

/**
 * Toggles the Discord Test Mode setting in script properties and rebuilds the menu.
 */
function _toggleDiscordTestModeInternal_() {
    const scriptProperties = PropertiesService.getScriptProperties();
    const currentState = isDiscordTestMode_(); // In Utilities.gs
    const newState = !currentState;
    scriptProperties.setProperty(PROP_DISCORD_TEST_MODE, newState.toString()); // Store as string 'true' or 'false'
    Logger.log(`Discord Test Mode toggled from ${currentState} to ${newState}.`);
    logAudit_('Toggled Discord Test Mode', 'N/A', `New State: ${newState ? 'ON' : 'OFF'}`); // In Utilities.gs

    const testWebhook = getTestWebhookUrl_(); // In Utilities.gs
     if (newState && (!testWebhook || testWebhook === 'YOUR_TEST_DISCORD_WEBHOOK_URL_HERE')) {
         SpreadsheetApp.getUi().alert(`Discord Test Mode is ON, but the Test Webhook URL is not set correctly in Script Properties (${PROP_TEST_WEBHOOK}). Messages may fail.`);
         Logger.log(`Warning: Test mode enabled but ${PROP_TEST_WEBHOOK} is not configured.`);
     } else {
        SpreadsheetApp.getUi().alert(`Discord Test Mode is now ${newState ? 'ON' : 'OFF'}.`);
     }

    // Rebuild the menu to reflect the change
    onOpen();
}

/**
 * Manually triggers the Send Discord action for the selected row.
 */
function _manuallySendDiscordInternal_() {
    const range = SpreadsheetApp.getActiveRange();
    if (!range) {
        SpreadsheetApp.getUi().alert('Please select a cell in the row you want to send.');
        return;
    }
    const sheet = range.getSheet();
    const row = range.getRow();
    const sheetName = sheet.getName();
    Logger.log(`Manual Send Discord triggered for row ${row} on sheet "${sheetName}".`);

    // Basic validation: Is it a "Month Year" sheet and not the header?
    if (row <= 1 || !MONTH_YEAR_SHEET_REGEX.test(sheetName)) {
        SpreadsheetApp.getUi().alert('Please select a cell within a valid downtime data row (not the header) on a "Month Year" sheet.');
        return;
    }

    // Target the response row (odd number)
    const responseRowIndex = (row % 2 !== 0) ? row : (row > 2 ? row - 1 : 3); // If even, use row above (unless it's row 2)
     if (responseRowIndex <= 1) {
         SpreadsheetApp.getUi().alert('Cannot determine a valid response row from selection.');
         return;
     }
    Logger.log(`Targeting response row: ${responseRowIndex}`);

    // Check status - Manual send *should* still check status to avoid accidental resends
    const statusValue = sheet.getRange(responseRowIndex, STATUS_COL).getValue();
     if (statusValue === 'sent') {
         SpreadsheetApp.getUi().alert(`Action blocked: Status for row ${responseRowIndex} is already 'sent'.`);
         Logger.log(`Manual Send Discord blocked for row ${responseRowIndex}: Status is 'sent'.`);
         return;
     }


    // Confirmation for Old Sheets
    const currentTargetSheet = getDowntimeSheetName_(); // In Utilities.gs
    let proceed = true;
    if (sheet.getName() !== currentTargetSheet) {
        Logger.log(`Manual Send: Confirmation needed for old sheet "${sheetName}".`);
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert(
            'Confirmation Needed',
            `You are manually triggering send for an OLDER month's sheet ('${sheetName}').\n\nAre you sure?`,
            ui.ButtonSet.YES_NO
        );
        if (response !== ui.Button.YES) {
            Logger.log('User cancelled manual action on old sheet.');
            proceed = false;
        } else {
             Logger.log('User confirmed manual action on old sheet.');
        }
    }

    if (proceed) {
        Logger.log(`Proceeding with manual Send Discord for row ${responseRowIndex}.`);
        logAudit_('Manual Send Discord Triggered', sheetName, `Row: ${responseRowIndex}`); // In Utilities.gs
        // Directly call the handler function
        handleSendDiscord_(sheet, responseRowIndex); // In Discord.gs
        // Note: handleSendDiscord_ shows its own alerts on failure/webhook issues,
        // and logs success/failure to Audit Log.
    }
}

/**
 * Manually triggers the Send Email action for the selected row.
 * This version SKIPS the status check.
 */
function _manuallySendEmailInternal_() {
    const range = SpreadsheetApp.getActiveRange();
    if (!range) {
        SpreadsheetApp.getUi().alert('Please select a cell in the row you want to send.');
        return;
    }
    const sheet = range.getSheet();
    const row = range.getRow();
    const sheetName = sheet.getName();
    Logger.log(`Manual Send Email triggered for row ${row} on sheet "${sheetName}".`);

    // Basic validation
    if (row <= 1 || !MONTH_YEAR_SHEET_REGEX.test(sheetName)) {
        SpreadsheetApp.getUi().alert('Please select a cell within a valid downtime data row (not the header) on a "Month Year" sheet.');
        return;
    }

    // Target the response row (odd number)
    const responseRowIndex = (row % 2 !== 0) ? row : (row > 2 ? row - 1 : 3); // If even, use row above (unless it's row 2)
     if (responseRowIndex <= 1) {
         SpreadsheetApp.getUi().alert('Cannot determine a valid response row from selection.');
         return;
     }
    Logger.log(`Targeting response row: ${responseRowIndex}`);

     // --- Status Check REMOVED for manual send ---
     Logger.log(`Manual Send Email skipping status check for row ${responseRowIndex}.`);

    // Confirmation for Old Sheets
    const currentTargetSheet = getDowntimeSheetName_(); // In Utilities.gs
    let proceed = true;
    if (sheet.getName() !== currentTargetSheet) {
         Logger.log(`Manual Send: Confirmation needed for old sheet "${sheetName}".`);
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert(
            'Confirmation Needed',
            `You are manually triggering send for an OLDER month's sheet ('${sheetName}').\n\nAre you sure?`,
            ui.ButtonSet.YES_NO
        );
        if (response !== ui.Button.YES) {
            Logger.log('User cancelled manual action on old sheet.');
            proceed = false;
        } else {
             Logger.log('User confirmed manual action on old sheet.');
        }
    }

    if (proceed) {
         Logger.log(`Proceeding with manual Send Email for row ${responseRowIndex}.`);
         logAudit_('Manual Send Email Triggered', sheetName, `Row: ${responseRowIndex}`); // In Utilities.gs
        // Directly call the handler function
        handleSendEmail_(sheet, responseRowIndex); // In Email.gs
         // Note: handleSendEmail_ shows its own alerts on success/failure,
         // and logs success/failure to Audit Log.
    }
}
