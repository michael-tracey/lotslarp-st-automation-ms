/**
 * @OnlyCurrentDoc
 *
 * Functions related to creating and displaying HTML Dialogs/Sidebars.
 */

/** Displays the downtime progress summary dialog. */
function showDowntimeProgressDialog() {
  Logger.log('Showing Downtime Progress Dialog...');
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = getDowntimeCompletionData_(sheet); // In SheetData.gs
    if (!data) {
        SpreadsheetApp.getUi().alert('Could not retrieve downtime data. Is this the correct sheet?');
        Logger.log('Failed to get downtime completion data.');
        return;
    }

    const completionPercentage = data.totalDowntimeCells > 0
        ? (data.completedDowntimeCells / data.totalDowntimeCells) * 100
        : 0;

    const template = HtmlService.createTemplateFromFile('DowntimeProgressDialog');
    template.data = data; // Pass data to the template
    template.completionPercentage = completionPercentage;
    template.keywordData = Object.keys(DOWNTIME_KEYWORDS).map(keyword => {
        const total = data.keywordCounts[keyword] || 0;
        const completed = data.keywordCompletedCounts[keyword] || 0;
        const keywordPercentage = total > 0 ? (completed / total) * 100 : 0;
        const overallPercentage = data.totalDowntimeCells > 0 ? (total / data.totalDowntimeCells) * 100 : 0;
        return {
            keyword: keyword,
            keywordPercentage: keywordPercentage.toFixed(1), // Use 1 decimal place
            completed: completed,
            total: total,
            overallPercentage: overallPercentage.toFixed(1) // Use 1 decimal place
        };
    });


    const htmlOutput = template.evaluate()
      .setWidth(800)
      .setHeight(550); // Adjusted height

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Downtime Summary');
    Logger.log('Downtime Progress Dialog displayed.');
  } catch (error) {
    Logger.log(`Error showing downtime progress dialog: ${error}\nStack: ${error.stack}`);
    SpreadsheetApp.getUi().alert(`Error showing dialog: ${error.message}`);
  }
}

/** Displays the influences progress summary dialog. */
function showInfluencesProgressDialog() {
    Logger.log('Showing Influences Progress Dialog...');
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const data = getInfluencesCompletionData_(sheet); // In SheetData.gs
        if (!data) {
            SpreadsheetApp.getUi().alert('Could not retrieve influence data. Is this the correct sheet?');
            Logger.log('Failed to get influence completion data.');
            return;
        }

        const template = HtmlService.createTemplateFromFile('InfluencesProgressDialog');
        template.data = data; // Pass data to the template

        const htmlOutput = template.evaluate()
            .setWidth(800)
            .setHeight(450); // Adjusted height

        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Influences Summary');
        Logger.log('Influences Progress Dialog displayed.');
    } catch (error) {
        Logger.log(`Error showing influences progress dialog: ${error}\nStack: ${error.stack}`);
        SpreadsheetApp.getUi().alert(`Error showing dialog: ${error.message}`);
    }
}

/** Displays the resources progress summary dialog. */
function showResourcesProgressDialog() {
    Logger.log('Showing Resources Progress Dialog...');
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const data = getResourcesCompletionData_(sheet); // In SheetData.gs
         if (!data) {
            SpreadsheetApp.getUi().alert('Could not retrieve resource data. Is this the correct sheet?');
            Logger.log('Failed to get resource completion data.');
            return;
        }

        const template = HtmlService.createTemplateFromFile('ResourcesProgressDialog');
        template.data = data; // Pass data to the template

        const htmlOutput = template.evaluate()
            .setWidth(800)
            .setHeight(450); // Adjusted height

        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Resources Summary');
        Logger.log('Resources Progress Dialog displayed.');
    } catch (error) {
        Logger.log(`Error showing resources progress dialog: ${error}\nStack: ${error.stack}`);
        SpreadsheetApp.getUi().alert(`Error showing dialog: ${error.message}`);
    }
}

/** Opens the cell editor dialog for the currently active cell. */
function openEditCellDialog() {
  Logger.log('Opening Edit Cell Dialog...');
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const row = cell.getRow();
  const col = cell.getColumn();

  // Basic check: Don't allow editing header row or timestamp/status/checkbox columns
  if (row <= 1 || col <= SEND_EMAIL_COL) {
      SpreadsheetApp.getUi().alert('Cannot edit this cell using the Downtime Editor.');
      Logger.log(`Edit dialog cancelled: Attempted edit on protected cell ${cell.getA1Notation()}.`);
      return;
  }

  // Determine if it's a submission or response row (response rows are ODD)
  const isResponseRow = row % 2 !== 0;
  const promptRow = isResponseRow ? row - 1 : row; // If it's a response row, prompt is above; otherwise, it's the row itself
  const responseRowForName = isResponseRow ? row : row + 1; // Name is always on the response row

  // Ensure promptRow is valid
  if (promptRow < 1) {
       SpreadsheetApp.getUi().alert('Cannot determine prompt row for this cell.');
       Logger.log(`Edit dialog cancelled: Invalid prompt row calculated (${promptRow}) for cell ${cell.getA1Notation()}.`);
       return;
  }


  const currentValue = cell.getValue();
  const promptValue = sheet.getRange(promptRow, col).getValue();
  const columnName = sheet.getRange(1, col).getValue();
  // Character name is always in the response row's character column
  const characterName = sheet.getRange(responseRowForName, CHARACTER_NAME_COL).getValue();

  Logger.log(`Editing Cell: ${cell.getA1Notation()}, Character: ${characterName}, Column: ${columnName}`);

  try {
    const template = HtmlService.createTemplateFromFile('EditCellDialog');
    template.currentValue = currentValue;
    template.promptValue = promptValue;
    template.columnName = columnName;
    template.characterName = characterName;
    template.row = row;
    template.col = col;
    template.markdownButtons = DISCORD_MARKDOWN_STYLES.map(style =>
      `<button class="md-button" data-prefix="${escapeHtml_(style.prefix)}" data-suffix="${escapeHtml_(style.suffix)}">${escapeHtml_(style.name)}</button>` // In Utilities.gs
    ).join('');

    const htmlOutput = template.evaluate()
      .setWidth(900)
      .setHeight(600);

    const dialogTitle = `Edit ${columnName} for ${characterName}`;
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, dialogTitle);
    Logger.log(`Edit Cell Dialog displayed for ${cell.getA1Notation()}.`);
  } catch (error) {
    Logger.log(`Error opening edit cell dialog: ${error}\nStack: ${error.stack}`);
    SpreadsheetApp.getUi().alert(`Error opening editor: ${error.message}`);
  }
}

/**
 * Updates a cell's value. Called from the client-side JavaScript of the editor dialog.
 * Includes logging before and after the update. Also logs to Audit sheet.
 * @param {string} newValue The new value for the cell.
 * @param {number} row The row number of the cell.
 * @param {number} col The column number of the cell.
 */
function updateCellValue(newValue, row, col) {
  const user = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  const sheet = SpreadsheetApp.getActiveSheet(); // Assumes the active sheet is the correct one
  const cell = sheet.getRange(row, col);
  const oldValue = cell.getValue();
  const sheetName = sheet.getName();
  const cellA1 = cell.getA1Notation();

  Logger.log(`Attempting update: User='${user}', Cell='${cellA1}', OldValue='${oldValue}', NewValue='${newValue}'`);

  try {
    cell.setValue(newValue);
    SpreadsheetApp.flush(); // Ensure the change is written
    const updatedValue = cell.getValue(); // Verify the update
    Logger.log(`Update successful: Cell='${cellA1}', UpdatedValue='${updatedValue}'`);

    // Log to Audit Sheet
    logAudit_('Cell Edit via Dialog', sheetName, `Cell: ${cellA1}, Old: ${oldValue}, New: ${newValue}`); // In Utilities.gs


    // Optionally change background color if cell is now non-empty
    if (newValue && String(newValue).trim() !== '' && String(oldValue).trim() === '') {
        cell.setBackground(null); // Clear background (removes red)
    } else if (String(newValue).trim() === '' && String(oldValue).trim() !== '') {
         // Optional: If cleared, maybe set back to red? Depends on workflow.
         // cell.setBackgroundRGB(255, 204, 204);
    }
  } catch (error) {
    Logger.log(`Error updating cell ${cellA1}: ${error}\nStack: ${error.stack}`);
    // Log failure to audit?
    logAudit_('Cell Edit via Dialog FAILED', sheetName, `Cell: ${cellA1}, Error: ${error.message}`); // In Utilities.gs
    // Rethrow or handle as needed - client side might need notification
    throw new Error(`Failed to update cell value: ${error.message}`);
  }
}

/**
 * Activates a specific cell in the active sheet. Called from client-side HTML dialogs.
 * @param {string} cellAddress A1 notation of the cell (e.g., "C5").
 */
function jumpToCell(cellAddress) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getRange(cellAddress);
    sheet.setActiveRange(range); // More reliable way to change selection
    Logger.log(`Jumped to cell: ${cellAddress}`);
  } catch (error) {
      Logger.log(`Error jumping to cell ${cellAddress}: ${error}`);
      // Optionally show an alert if jumping fails
      // SpreadsheetApp.getUi().alert(`Could not jump to cell ${cellAddress}: ${error.message}`);
  }
}

/**
 * Displays a generic dialog for showing missing items (Downtimes, Influences, Resources).
 * @param {string} itemType Title for the dialog (e.g., "Downtime", "Influence").
 * @param {function} getDataFunction Function to call to get the array of missing items (e.g., getMissingDowntimeData_).
 */
function showMissingItemsDialog_(itemType, getDataFunction) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const missingItems = getDataFunction(sheet); // Assumes getDataFunction is in SheetData.gs

    if (!missingItems) {
        SpreadsheetApp.getUi().alert(`Could not retrieve missing ${itemType} data. Is this the correct sheet?`);
        Logger.log(`Failed to get missing ${itemType} data.`);
        return;
    }

    if (missingItems.length === 0) {
        SpreadsheetApp.getUi().alert(`All ${itemType} responses have been filled out!`);
        Logger.log(`No missing ${itemType} responses found.`);
        return;
    }

    try {
        const template = HtmlService.createTemplateFromFile('MissingItemsDialog');
        template.itemType = itemType;
        template.missingItems = missingItems;

        const htmlOutput = template.evaluate()
            .setWidth(800)
            .setHeight(500);

        const dialogTitle = `Missing ${itemType} Responses`;
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, dialogTitle);
        Logger.log(`Missing ${itemType} Dialog displayed with ${missingItems.length} items.`);
    } catch (error) {
        Logger.log(`Error showing missing ${itemType} dialog: ${error}\nStack: ${error.stack}`);
        SpreadsheetApp.getUi().alert(`Error showing dialog: ${error.message}`);
    }
}

// Wrapper function for the menu item, ensuring it calls the correct internal function
function showMissingDowntimeDialog() {
    Logger.log('Menu item "Show Missing Downtime Responses" clicked.');
    showMissingItemsDialog_('Downtime', getMissingDowntimeData_); // In SheetData.gs
}

/**
 * Gets unique colors from the sheet and displays the "Find by Color" dialog.
 */
function showFindPendingByColorDialog() {
    Logger.log('Showing Find Pending by Color Dialog...');
    const ui = SpreadsheetApp.getUi();
    try {
        const uniqueColors = getUniqueBackgroundColors_(); // In SheetData.gs
        if (!uniqueColors) { // Null on error from function
            ui.alert("Could not retrieve background colors from the sheet. Check logs for details.");
            return;
        }

        const template = HtmlService.createTemplateFromFile('FindPendingByColorDialog');
        template.colors = uniqueColors; // Pass the array of color strings

        const htmlOutput = template.evaluate()
            .setWidth(850)
            .setHeight(500);

        ui.showModalDialog(htmlOutput, 'Find Pending Responses by Color');
        Logger.log('Find Pending by Color Dialog displayed.');

    } catch (error) {
        Logger.log(`Error showing Find Pending by Color Dialog: ${error}\nStack: ${error.stack}`);
        ui.alert(`Error showing dialog: ${error.message}`);
    }
}
