/**
 * @OnlyCurrentDoc
 *
 * Functions for the Re-import Downtime dialog.
 */

/**
 * Shows the Re-import Downtime dialog.
 */
function showReimportDowntimeDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ReimportDowntimeDialog.html')
      .setWidth(400)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Re-import Downtime');
}

/**
 * Gets a list of all form responses from the downtime form.
 * @returns {Array<{id: string, characterName: string}>} An array of objects containing the response ID and character name.
 */
function getFormResponses() {
  try {
    const formId = PropertiesService.getScriptProperties().getProperty(PROP_DOWNTIME_FORM_ID);
    if (!formId || formId === 'YOUR_GOOGLE_FORM_ID_HERE') {
      throw new Error(`Downtime Form ID is not set correctly in Script Properties (${PROP_DOWNTIME_FORM_ID}).`);
    }

    const form = FormApp.openById(formId);
    const formResponses = form.getResponses();
    const responses = [];

    formResponses.forEach(formResponse => {
      const itemResponses = formResponse.getItemResponses();
      // Assuming the first question is always the character name.
      const characterName = itemResponses.length > 0 ? itemResponses[0].getResponse() : 'Unknown Character';
      responses.push({
        id: formResponse.getId(),
        characterName: `${characterName} (${formResponse.getTimestamp().toLocaleString()})`
      });
    });

    return responses.reverse(); // Show newest first
  } catch (error) {
    Logger.log(`Error in getFormResponses: ${error.message}`);
    // Re-throw the error so it can be caught by the .withFailureHandler on the client-side
    throw new Error(`Failed to fetch form responses: ${error.message}`);
  }
}

/**
 * Re-imports a single form response by its ID.
 * @param {string} responseId The ID of the form response to re-import.
 * @returns {string} A success message.
 */
function reimportFormResponse(responseId) {
  try {
    const formId = PropertiesService.getScriptProperties().getProperty(PROP_DOWNTIME_FORM_ID);
    if (!formId || formId === 'YOUR_GOOGLE_FORM_ID_HERE') {
      throw new Error(`Downtime Form ID is not set correctly in Script Properties (${PROP_DOWNTIME_FORM_ID}).`);
    }

    const form = FormApp.openById(formId);
    const formResponse = form.getResponse(responseId);

    if (!formResponse) {
      throw new Error(`Form response with ID ${responseId} not found.`);
    }

    // The onFormSubmitHandler_ expects an event object 'e'.
    // We need to construct a mock event object that has the same structure.
    const mockEvent = {
      response: formResponse,
      source: form,
      authMode: ScriptApp.AuthMode.FULL, // Or whatever is appropriate
      triggerUid: 'manual-reimport' // Can be any string
    };

    Logger.log(`Manually re-importing form response ID: ${responseId}`);
    
    // Call the existing handler with the mock event
    onFormSubmitHandler_(mockEvent);

    const characterNameResponse = formResponse.getItemResponses()[0].getResponse();
    const successMessage = `Successfully re-imported downtime for '${characterNameResponse}'.`;
    Logger.log(successMessage);
    logAudit_('Manual Re-import', 'N/A', `Response ID: ${responseId}, Character: ${characterNameResponse}`);

    return successMessage;
  } catch (error) {
    Logger.log(`Error in reimportFormResponse: ${error.message} 
Stack: ${error.stack}`);
    // Re-throw for the client-side handler
    throw new Error(`Failed to re-import: ${error.message}`);
  }
}
