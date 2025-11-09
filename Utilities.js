/**
 * @OnlyCurrentDoc
 *
 * General utility functions for the LARP Downtime Management Script.
 */

/**
 * Gets the display name of a narrator from their email address.
 * Caches the email-to-name mapping to reduce sheet lookups.
 * @param {string} email The email address of the user.
 * @returns {string|null} The narrator's name or null if not found.
 */
function getNarratorNameByEmail_(email) {
  if (!email) return null;

  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'narrator_email_map';
  let emailMap = cache.get(CACHE_KEY);

  if (emailMap === null) {
    try {
      const narratorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('narrators');
      if (!narratorSheet) {
        Logger.log('getNarratorNameByEmail_: "narrators" sheet not found.');
        return null; // No sheet, no names
      }
      const data = narratorSheet.getDataRange().getValues();
      const headers = data[0].map(h => String(h).toLowerCase());
      const nameCol = headers.indexOf('name');
      const emailCol = headers.indexOf('email');

      if (nameCol === -1 || emailCol === -1) {
        Logger.log('getNarratorNameByEmail_: Could not find "Name" or "Email" column in "narrators" sheet.');
        return null;
      }

      const tempMap = {};
      for (let i = 1; i < data.length; i++) {
        const name = data[i][nameCol];
        const emailAddress = data[i][emailCol];
        if (name && emailAddress) {
          tempMap[String(emailAddress).toLowerCase().trim()] = String(name);
        }
      }
      
      emailMap = tempMap;
      // Cache the map for 6 hours
      cache.put(CACHE_KEY, JSON.stringify(emailMap), 21600);

    } catch (e) {
      Logger.log(`getNarratorNameByEmail_: Error building narrator map: ${e}`);
      return null; // Return null on error
    }
  } else {
    // If map was in cache, it's a JSON string
    emailMap = JSON.parse(emailMap);
  }

  return emailMap[email.toLowerCase().trim()] || null;
}


// --- Audit Log Helper ---

/**
 * Logs an action to the 'Log' sheet. Creates the sheet if it doesn't exist.
 * @param {string} action - Description of the action performed.
 * @param {string} sheetName - Name of the sheet where the action occurred.
 * @param {string} [details=''] - Optional additional details about the action.
 */
function logAudit_(action, sheetName, details = '') {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let logSheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);

        // Create sheet and header if it doesn't exist
        if (!logSheet) {
            logSheet = ss.insertSheet(AUDIT_LOG_SHEET_NAME, ss.getNumSheets()); // Insert at the end
            const headers = ['Timestamp', 'User', 'Action', 'Sheet Name', 'Details'];
            logSheet.appendRow(headers);
            logSheet.setFrozenRows(1);
            logSheet.getRange("A1:E1").setFontWeight("bold");
            // Optional: Set column widths
            logSheet.setColumnWidth(1, 150); // Timestamp
            logSheet.setColumnWidth(2, 180); // User
            logSheet.setColumnWidth(3, 180); // Action
            logSheet.setColumnWidth(4, 120); // Sheet Name
            logSheet.setColumnWidth(5, 400); // Details
            Logger.log(`Created Audit Log sheet: "${AUDIT_LOG_SHEET_NAME}"`);
        }

        const timestamp = new Date();
        // Attempt to get user email; fallback if restricted
        let user = 'Unknown User';
        try {
            user = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'Unknown User';
        } catch (e) {
             Logger.log('Could not get user email for audit log: ' + e);
        }
        


        // Append the log entry
        logSheet.appendRow([timestamp, user, action, sheetName, details]);

    } catch (error) {
        Logger.log(`CRITICAL ERROR writing to Audit Log: ${error} - Action: ${action}, Details: ${details}`);
        // Alert if logging fails, as it's important
        SpreadsheetApp.getUi().alert('CRITICAL ERROR: Failed to write to the Audit Log sheet. Please check script permissions or contact support.');
    }
}


/**
 * Attempts an action that requires user authorization,
 * aiming to trigger the Google OAuth consent screen if permissions are missing.
 * This can help ensure the script has the necessary permissions.
 *
 * To use this:
 * 1. Add this function to one of your .gs files (e.g., Utilities.js).
 * 2. Open the Apps Script editor (Extensions > Apps Script).
 * 3. In the editor, select "forceReauthorizationCheck_" from the function dropdown.
 * 4. Click the "Run" button (play icon).
 * 5. If prompted, review and grant the permissions.
 */
function forceReauthorizationCheck_() {
  try {
    // Attempt to access a service that requires specific authorization.
    // Accessing the user's email is a common one that requires explicit consent.
    const email = Session.getActiveUser().getEmail();
    
    // If the above line executes without error, permissions are likely in place.
    SpreadsheetApp.getUi().alert(
      'Authorization Check Successful',
      'The script appears to have the necessary permissions to access your email address: ' + email,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    Logger.log('forceReauthorizationCheck_: Permissions appear to be granted. User email retrieved.');

  } catch (e) {
    // An error here might mean the authorization prompt was shown and is being processed,
    // or that there's another issue.
    Logger.log('forceReauthorizationCheck_: An error occurred. This might be due to the authorization flow. Error: ' + e.toString());
    SpreadsheetApp.getUi().alert(
      'Authorization Status',
      'If you were just prompted to authorize the script, please try your intended action again now that permissions should be granted. If no prompt appeared, an error occurred: ' + e.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// --- Test Mode Helpers ---

/**
 * Checks if Discord Test Mode is currently enabled.
 * @returns {boolean} True if test mode is on, false otherwise.
 */
function isDiscordTestMode_() {
    const testMode = PropertiesService.getScriptProperties().getProperty(PROP_DISCORD_TEST_MODE);
    return testMode === 'true'; // Check against string 'true'
}

/**
 * Gets the configured Test Discord Webhook URL.
 * @returns {string|null} The URL string or null if not set.
 */
function getTestWebhookUrl_() {
    return PropertiesService.getScriptProperties().getProperty(PROP_TEST_WEBHOOK);
}

// --- Sheet/Date Utilities ---

/**
 * Gets the expected name of the current downtime sheet based on script properties.
 * @returns {string} The sheet name (e.g., "March 2025").
 */
function getDowntimeSheetName_() {
  const year = PropertiesService.getScriptProperties().getProperty(PROP_DOWNTIME_YEAR);
  const month = PropertiesService.getScriptProperties().getProperty(PROP_DOWNTIME_MONTH);
  Logger.log(`Retrieved downtime properties: Year='${year}', Month='${month}'`);
  if (!year || !month) {
      Logger.log(`Warning: Downtime year (${year}) or month (${month}) not set in properties. Using default name.`);
      return "Current Downtime"; // Fallback name
  }
  const sheetName = `${month} ${year}`;
  Logger.log(`Constructed target sheet name: "${sheetName}"`);
  return sheetName;
}

/**
 * Gets the active downtime sheet, creating it if it doesn't exist.
 * Uses the form response to determine headers if creating a new sheet.
 * @param {GoogleAppsScript.Forms.FormResponse} [formResponse] Optional form response to use for headers if creating the sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The downtime sheet.
 */
function getOrCreateDowntimeSheet_(formResponse) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getDowntimeSheetName_();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`Sheet '${sheetName}' not found. Creating new sheet.`);
    sheet = ss.insertSheet(sheetName);
    // Build header row
    const header = ["Timestamp", "Status", "Send Discord", "Send Email"];
    if (formResponse) {
        const itemResponses = formResponse.getItemResponses();
        itemResponses.forEach(itemResponse => {
            header.push(itemResponse.getItem().getTitle());
        });
    } else {
        // Add default/placeholder headers if no form response provided
        header.push("Character Name", "Downtime Action 1", "Downtime Action 2", "Influence Action", "Resource Action");
        Logger.log('Created sheet with default headers as no formResponse was provided.');
    }

    sheet.appendRow(header);
    sheet.setFrozenRows(1);
    // Apply status dropdown validation
    const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['input', 'unprocessed', 'processed/pending', 'process/hold', 'sent']).build();
    sheet.getRange(`B2:B`).setDataValidation(statusRule); // Apply to column B starting from row 2
    Logger.log(`Created and formatted new sheet: '${sheetName}'.`);
  }
  return sheet;
}

/**
 * Determines if a given hex color is considered 'dark' for readability purposes.
 * @param {string} hexColor The hex color string (e.g., '#RRGGBB' or 'RRGGBB').
 * @returns {boolean} True if the color is dark, false otherwise.
 */   
function isColorDark(hexColor) {
  // Remove '#' if present
  const hex = hexColor.startsWith('#') ? hexColor.slice(1) : hexColor;

  // Convert hex to RGB
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
          
  // Calculate luminance (perceived brightness) using the W3C recommended formula
  const luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
      
  // A common threshold for dark colors is 0.5. Adjust as needed.
  return luminance < 0.5;
}


// --- Calculation/Formatting Utilities ---

/**
 * Calculates statistics (average, median, min, max) for an array of numbers.
 * @param {number[]} numbers - Array of numbers.
 * @returns {{average: number, median: number, min: number, max: number}} Calculated statistics. Returns 0 for all if array is empty or contains no valid numbers.
 */
function calculateStats_(numbers) {
    if (!numbers || numbers.length === 0) {
        return { average: 0, median: 0, min: 0, max: 0 };
    }

    const validNumbers = numbers.filter(n => typeof n === 'number' && !isNaN(n));
    if (validNumbers.length === 0) {
         return { average: 0, median: 0, min: 0, max: 0 };
    }

    const sortedNumbers = [...validNumbers].sort((a, b) => a - b);
    const sum = validNumbers.reduce((acc, val) => acc + val, 0);
    const average = sum / validNumbers.length;
    const min = sortedNumbers[0];
    const max = sortedNumbers[sortedNumbers.length - 1];

    let median;
    const mid = Math.floor(validNumbers.length / 2);
    if (validNumbers.length % 2 === 0) {
        median = (sortedNumbers[mid - 1] + sortedNumbers[mid]) / 2;
    } else {
        median = sortedNumbers[mid];
    }

    return { average, median, min, max };
}

/**
 * Converts a column number (1-based) to its corresponding letter (A, B, AA, etc.).
 * @param {number} columnNumber The 1-based column number.
 * @returns {string} The column letter.
 */
function getColumnLetter(columnNumber) {
  let columnLetter = '';
  let num = columnNumber;
  while (num > 0) {
    let remainder = (num - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    num = Math.floor((num - 1) / 26);
  }
  return columnLetter;
}

/**
 * Simple HTML escaping function.
 * @param {string} str The string to escape.
 * @returns {string} The escaped string.
 */
function escapeHtml_(str) {
    if (str === null || str === undefined) return '';
    return String(str).replace(/&/g, '&amp;')
                      .replace(/</g, '&lt;')
                      .replace(/>/g, '&gt;')
                      .replace(/"/g, '&quot;')
                      .replace(/'/g, '&#039;');
}

/**
 * Escapes characters that have special meaning in regular expressions.
 * @param {string} str The input string.
 * @returns {string} The escaped string.
 */
function escapeRegex_(str) {
    return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
}


/**
 * Checks if a character is whitespace.
 * @param {string} char The character to check.
 * @returns {boolean} True if the character is whitespace, false otherwise.
 */
function isWhitespace_(char) {
  return /\s/.test(char);
}

/**
 * Parses a sheet name expected to be in "Month Year" format.
 * @param {string} sheetName The sheet name (e.g., "April 2025").
 * @returns {{month: number, year: number}|null} An object with zero-based month and year, or null if parsing fails.
 */
function parseSheetName_(sheetName) {
    if (!sheetName) return null;
    const parts = sheetName.trim().split(' ');
    if (parts.length !== 2) return null;

    const monthName = parts[0];
    const year = parseInt(parts[1], 10);

    if (isNaN(year)) return null;

    // Simple month name to number mapping (case-insensitive)
    const monthMap = {
        january: 0, february: 1, march: 2, april: 3, may: 4, june: 5,
        july: 6, august: 7, september: 8, october: 9, november: 10, december: 11
    };
    const month = monthMap[monthName.toLowerCase()];

    if (month === undefined) {
        Logger.log(`Could not parse month name: "${monthName}"`);
        return null; // Month name not recognized
    }

    return { month, year };
}
