/**
 * @OnlyCurrentDoc
 *
 * Functions for managing Script Properties via a dialog.
 */

/**
 * Opens a dialog to view and edit script properties.
 * Intended to be called from a menu item.
 */
function openScriptPropertiesDialog_() {
  try {
    const htmlOutput = HtmlService.createTemplateFromFile('ScriptPropertiesDialog')
      .evaluate()
      .setWidth(700)
      .setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Manage Script Properties');
    Logger.log('Script Properties dialog opened.');
  } catch (e) {
    Logger.log(`Error opening script properties dialog: ${e.toString()}\nStack: ${e.stack}`);
    SpreadsheetApp.getUi().alert(`Error opening dialog: ${e.message}`);
  }
}

/**
 * Fetches all script properties for the dialog.
 * @returns {Object} A plain JavaScript object الإعلامية (key-value pairs) of all script properties.
 */
function getScriptPropertiesForDialog() {
  try {
    const properties = PropertiesService.getScriptProperties().getProperties();
    Logger.log(`Fetched ${Object.keys(properties).length} script properties for dialog.`);
    return properties;
  } catch (e) {
    Logger.log(`Error fetching script properties: ${e.toString()}`);
    throw new Error(`Failed to fetch script properties: ${e.message}`);
  }
}

/**
 * Updates script properties based on input from the dialog.
 * @param {Object} newProperties A plain JavaScript object where keys are property names and values are the new property values.
 * @returns {string} A status message indicating success or failure.
 */
function updateScriptPropertiesFromDialog(newProperties) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    let updatedCount = 0;
    let details = [];

    for (const key in newProperties) {
      if (newProperties.hasOwnProperty(key)) {
        const oldValue = scriptProperties.getProperty(key);
        const newValue = newProperties[key];
        if (oldValue !== newValue) {
          scriptProperties.setProperty(key, newValue);
          details.push(`'${key}' changed`);
          updatedCount++;
        }
      }
    }
    const message = `Successfully updated ${updatedCount} script properties.`;
    Logger.log(`${message} Details: ${details.join(', ')}`);
    logAudit_('Script Properties Updated', 'N/A', `${updatedCount} properties updated via dialog. ${details.join('; ')}`);
    return message;
  } catch (e) {
    Logger.log(`Error updating script properties: ${e.toString()}\nStack: ${e.stack}`);
    logAudit_('Script Properties Update FAILED', 'N/A', `Error: ${e.message}`);
    throw new Error(`Failed to update script properties: ${e.message}`);
  }
}