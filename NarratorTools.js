/**
 * @OnlyCurrentDoc
 *
 * Functions related to narrator tools.
 */

/**
 * Shows the dialog to encrypt a narrator's password.
 */
function showEncryptPasswordDialog_() {
  const html = HtmlService.createHtmlOutputFromFile('EncryptPasswordDialog')
      .setWidth(400)
      .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Encrypt Narrator Password');
}

/**
 * Encrypts the given password and updates the selected row in the 'narrators' sheet.
 *
 * @param {string} password The password to encrypt.
 * @return {string} A success or error message.
 */
function encryptPassword(password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('narrators');
    if (!sheet) {
      return "Error: 'narrators' sheet not found.";
    }

    const range = SpreadsheetApp.getActiveRange();
    if (!range) {
      return "Error: No active range selected.";
    }

    if (range.getSheet().getName() !== 'narrators') {
      return "Error: Please select a cell in the 'narrators' sheet.";
    }

    const row = range.getRow();
    if (row <= 1) {
      return "Error: Please select a row other than the header.";
    }

    const encryptedPassword = Utilities.base64Encode(password);
    sheet.getRange(row, 5).setValue(encryptedPassword); // Password is in the 5th column

    // Get the background color of the Hex Color cell and set its value
    const colorCell = sheet.getRange(row, 3); // Hex Color is in the 3rd column
    const backgroundColor = colorCell.getBackground();
    colorCell.setValue(backgroundColor);

    return "Password encrypted and hex code value set successfully.";
  } catch (e) {
    Logger.log('Error encrypting password: ' + e.toString());
    return "Error: " + e.toString();
  }
}