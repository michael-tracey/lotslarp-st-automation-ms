/**
 * Main entry point for the web app. Handles session management and serves the appropriate page.
 */
function doGet(e) {
  const larpName = PropertiesService.getScriptProperties().getProperty('LARP_NAME') || 'Downtime Portal';

  // If the action is to logout, clear the session and show the login page.
  if (e.parameter.action === 'logout') {
    logout();
    const template = HtmlService.createTemplateFromFile('Login.html');
    template.appUrl = ScriptApp.getService().getUrl();
    template.larpName = larpName;
    return template.evaluate().setTitle(larpName + ' Login');
  }

  const user = getActiveUser_();

  if (user) {
    // If the user is logged in, show the main interface.
    const template = HtmlService.createTemplateFromFile('WebAppInterface.html');
    template.user = user;
    template.larpName = larpName;
    return template.evaluate().setTitle(larpName);
  } else {
    // If the user is not logged in, show the login page.
    const template = HtmlService.createTemplateFromFile('Login.html');
    template.appUrl = ScriptApp.getService().getUrl();
    template.larpName = larpName;
    return template.evaluate().setTitle(larpName + ' Login');
  }
}

/**
 * Logs in a user by verifying their credentials against the 'narrators' sheet.
 *
 * @param {string} username The user's username.
 * @param {string} password The user's password.
 * @return {object} An object with a 'success' property and a 'message'.
 */
function login(username, password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('narrators');
    if (!sheet) {
      return { success: false, message: "'narrators' sheet not found." };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const usernameCol = headers.indexOf('Username');
    const passwordCol = headers.indexOf('Password');
    const colorCol = headers.indexOf('Hex Color');
    const nameCol = headers.indexOf('Name');
    const roleCol = headers.indexOf('Role');

    if (usernameCol === -1 || passwordCol === -1 || colorCol === -1 || nameCol === -1 || roleCol === -1) {
      return { success: false, message: "Required columns not found in 'narrators' sheet." };
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[usernameCol] === username) {
        const encryptedPassword = row[passwordCol];
        const decryptedPassword = Utilities.newBlob(Utilities.base64Decode(encryptedPassword)).getDataAsString("UTF-8");

        if (decryptedPassword === password) {
          const user = {
            username: row[usernameCol],
            name: row[nameCol],
            color: row[colorCol],
            role: row[roleCol]
          };
          CacheService.getUserCache().put('user', JSON.stringify(user), 3600);
          logAudit_('User Login', 'WebApp', `User: ${user.username}`);
          return { success: true, message: 'Login successful.' };
        }
      }
    }

    return { success: false, message: 'Invalid username or password.' };
  } catch (e) {
    Logger.log('Error during login: ' + e.toString());
    return { success: false, message: 'An error occurred during login.' };
  }
}

/**
 * Logs out the current user by clearing their session from the cache.
 */
function logout() {
  CacheService.getUserCache().remove('user');
}

/**
 * Gets the active user from the cache.
 *
 * @return {object|null} The user object if logged in, otherwise null.
 */
function getActiveUser_() {
  const userJson = CacheService.getUserCache().get('user');
  if (userJson) {
    return JSON.parse(userJson);
  }
  return null;
}

/**
 * Fetches downtime tasks for the logged-in user.
 *
 * @return {Object} An object containing the list of tasks and the sheet name.
 */
function getDowntimeTasks() {
  const user = getActiveUser_();
  if (!user) {
    throw new Error('You are not logged in.');
  }

  const properties = PropertiesService.getScriptProperties();
  const sheetName = properties.getProperty('DOWNTIME_SHEET_NAME');
  const generalHexColor = properties.getProperty('TASK_COLOR_HEX');

  if (!sheetName) {
    throw new Error('Web app is not configured. Please set DOWNTIME_SHEET_NAME script property.');
  }

  // --- User Tasks ---
  const userHexColor = user.color;
  const myPendingFound = findMissingResponsesByColor_(userHexColor, sheetName);
  const myCompletedFound = findCompletedResponsesByColor_(userHexColor, sheetName);

  const myPending = myPendingFound.map(item => ({
    taskID: item.cell,
    characterName: (item.characterName || '').replace(/\n/g, ' '),
    summary: (item.text || '').replace(/\n/g, ' ')
  }));
  const myCompleted = myCompletedFound.map(item => ({
    taskID: item.cell,
    characterName: (item.characterName || '').replace(/\n/g, ' '),
    summary: (item.text || '').replace(/\n/g, ' '),
    response: (item.response || '').replace(/\n/g, ' '),
    completedBy: item.completedBy
  }));

  const myTotal = myPending.length + myCompleted.length;
  const myCompletionRate = myTotal > 0 ? Math.round((myCompleted.length / myTotal) * 100) : 0;

  // --- General Tasks ---
  let generalPending = [];
  let generalCompleted = [];
  let generalTotalTasks = 0;

  if (generalHexColor) {
    const generalPendingFound = findMissingResponsesByColor_(generalHexColor, sheetName);
    const generalCompletedFound = findCompletedResponsesByColor_(generalHexColor, sheetName);

    generalPending = generalPendingFound.map(item => ({
      taskID: item.cell,
      characterName: (item.characterName || '').replace(/\n/g, ' '),
      summary: (item.text || '').replace(/\n/g, ' ')
    }));
    generalCompleted = generalCompletedFound.map(item => ({
      taskID: item.cell,
      characterName: (item.characterName || '').replace(/\n/g, ' '),
      summary: (item.text || '').replace(/\n/g, ' '),
      response: (item.response || '').replace(/\n/g, ' '),
      completedBy: item.completedBy
    }));

    generalTotalTasks = generalPending.length + generalCompleted.length;
  }

  return {
    myPending: myPending,
    myCompleted: myCompleted,
    myCompletionRate: myCompletionRate,
    generalPending: generalPending,
    generalCompleted: generalCompleted,
    generalTotalTasks: generalTotalTasks,
    generalHexColor: generalHexColor,
    sheetName: sheetName,
    user: user
  };
}

/**
 * Writes the response back to the sheet.
 *
 * @param {string} taskID The fully qualified A1 notation of the cell to edit.
 * @param {string} responseText The text to write into the cell.
 * @return {Object} A status object.
 */
function submitDowntimeResponse(taskID, responseText) {
  const user = getActiveUser_();
  if (!user) {
    throw new Error('You are not logged in.');
  }

  try {
    const range = SpreadsheetApp.getActiveSpreadsheet().getRange(taskID);
    range.setValue(responseText);
    const note = `Completed by: ${user.name}`;
    range.setNote(note);
    return { status: 'success', message: `Response for ${taskID} submitted!` };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

/**
 * Assigns a task to a narrator by changing the background color of the task cell.
 *
 * @param {string} taskID The fully qualified A1 notation of the cell to edit.
 * @param {string} color The hex color to set as the background.
 * @return {object} A status object.
 */
function assignTaskColor(taskID, color) {
  const user = getActiveUser_();
  if (!user) {
    throw new Error('You are not logged in.');
  }

  try {
    // The taskID is in the format 'Sheet Name'!A1. We need to get the submission cell, which is one row above the response cell.
    const parts = taskID.split('!');
    const sheetName = parts[0].replace(/'/g, '');
    const rangeA1 = parts[1];
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const responseRange = sheet.getRange(rangeA1);
    const submissionRange = responseRange.offset(-1, 0);

    submissionRange.setBackground(color);
    
    logAudit_('Assign Task', 'WebApp', `Task ${taskID} assigned to color ${color} by ${user.name}`);
    return { status: 'success', message: `Task ${taskID} assigned.` };
  } catch (error) {
    Logger.log(`Error assigning task color: ${error.toString()}`);
    return { status: 'error', message: error.message };
  }
}

function findMissingResponsesByColor_(hexColor, sheetName) {
  if (!hexColor || !/^#([0-9A-F]{3}){1,2}$/i.test(hexColor)) {
    throw new Error("Invalid hex color format provided. Make sure it starts with #.");
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet with name '${sheetName}' was not found.`);
  }

  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const dataRange = sheet.getDataRange();
    const backgrounds = dataRange.getBackgrounds();
    const values = dataRange.getValues();
    const foundItems = [];

    for (let i = 1; i < values.length; i += 2) {
      if (i + 1 >= values.length) continue;

      const submissionRowValues = values[i];
      const submissionRowBgs = backgrounds[i];
      const responseRowValues = values[i + 1];
      const characterName = responseRowValues[CHARACTER_NAME_COL - 1] || 'Unknown';

      for (let j = 0; j < submissionRowValues.length; j++) {
        const submissionValue = String(submissionRowValues[j] || '').trim();
        const cellBg = submissionRowBgs[j];

        if (submissionValue && cellBg.toLowerCase() === hexColor.toLowerCase()) {
          const responseValue = String(responseRowValues[j] || '').trim();
          if (responseValue === '') {
            const header = headers[j] || `Column ${j + 1}`;
            const a1 = sheet.getRange(i + 2, j + 1).getA1Notation();
            const responseCellA1 = `'${sheetName}'!${a1}`;

            foundItems.push({
              characterName: characterName,
              header: header,
              text: submissionValue,
              cell: responseCellA1
            });
          }
        }
      }
    }

    return foundItems;
  } catch (error) {
    Logger.log(`Error getting missing responses by color: ${error}\nStack: ${error.stack}`);
    throw new Error(`Failed to retrieve data from the sheet: ${error.message}`);
  }
}

/**
 * Finds all downtime submissions with a specific background color that HAVE a response.
 * @param {string} hexColor The hex color string to search for.
 * @param {string} sheetName The name of the sheet to search in.
 * @returns {Array<{characterName: string, header: string, text: string, cell: string, response: string}>} An array of objects for each completed item.
 */
function findCompletedResponsesByColor_(hexColor, sheetName) {
  if (!hexColor || !/^#([0-9A-F]{3}){1,2}$/i.test(hexColor)) {
    throw new Error("Invalid hex color format provided. Make sure it starts with #.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet with name '${sheetName}' was not found.`);
  }

  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const dataRange = sheet.getDataRange();
    const backgrounds = dataRange.getBackgrounds();
    const values = dataRange.getValues();
    const notes = dataRange.getNotes(); // Get all notes
    const foundItems = [];

    for (let i = 1; i < values.length; i += 2) {
      if (i + 1 >= values.length) continue;

      const submissionRowValues = values[i];
      const submissionRowBgs = backgrounds[i];
      const responseRowValues = values[i + 1];
      const responseRowNotes = notes[i + 1]; // Get notes for the response row
      const characterName = responseRowValues[CHARACTER_NAME_COL - 1] || 'Unknown';

      for (let j = 0; j < submissionRowValues.length; j++) {
        const submissionValue = String(submissionRowValues[j] || '').trim();
        const cellBg = submissionRowBgs[j];

        if (submissionValue && cellBg.toLowerCase() === hexColor.toLowerCase()) {
          const responseValue = String(responseRowValues[j] || '').trim();
          if (responseValue !== '') { // Find completed tasks
            const header = headers[j] || `Column ${j + 1}`;
            const a1 = sheet.getRange(i + 2, j + 1).getA1Notation();
            const responseCellA1 = `'${sheetName}'!${a1}`;
            const note = responseRowNotes[j]; // Get the note for the specific cell

            // Parse narrator name from note
            let completedBy = null;
            if (note && note.startsWith("Completed by: ")) {
              completedBy = note.substring("Completed by: ".length);
            }

            foundItems.push({
              characterName: characterName,
              header: header,
              text: submissionValue,
              cell: responseCellA1,
              response: responseValue, // Include the response
              completedBy: completedBy // Add the completer's name
            });
          }
        }
      }
    }

    return foundItems;
  } catch (error) {
    Logger.log(`Error getting completed responses by color: ${error}\nStack: ${error.stack}`);
    throw new Error(`Failed to retrieve data from the sheet: ${error.message}`);
  }
}

/**
 * Gets a sorted list of narrator names and their assigned hex colors.
 * @returns {Array<{name: string, color: string}>} A sorted array of objects, each with a name and color.
 */
function getNarratorColors() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('narrators');
    if (!sheet) {
      return []; // Return empty array if sheet not found
    }
    const data = sheet.getDataRange().getValues();
    const narrators = [];
    // Start from 1 to skip header
    for (let i = 1; i < data.length; i++) {
      const name = data[i][0]; // Column A: Name
      const color = data[i][2]; // Column C: Hex Color
      if (name && color) {
        narrators.push({ name: name, color: color });
      }
    }

    // Sort narrators alphabetically by name
    narrators.sort((a, b) => a.name.localeCompare(b.name));

    // Add the "Any Narrator" option at the top
    const generalColor = PropertiesService.getScriptProperties().getProperty('TASK_COLOR_HEX');
    if (generalColor) {
      narrators.unshift({ name: 'Any Narrator', color: generalColor });
    }

    return narrators;
  } catch (e) {
    Logger.log('Error in getNarratorColors: ' + e.toString());
    return [];
  }
}
