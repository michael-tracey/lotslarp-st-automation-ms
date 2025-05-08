/**
 * @OnlyCurrentDoc
 *
 * Functions for managing script triggers (onFormSubmit, onEdit).
 */

/**
 * Reinstalls ONLY the Form Submit trigger.
 * Deletes existing Form Submit triggers first to prevent duplicates.
 */
function reinstallFormTrigger_() {
  const formId = PropertiesService.getScriptProperties().getProperty(PROP_DOWNTIME_FORM_ID);

  if (!formId || formId === 'YOUR_GOOGLE_FORM_ID_HERE') {
      SpreadsheetApp.getUi().alert(`Error: Downtime Form ID is not set correctly in Script Properties (${PROP_DOWNTIME_FORM_ID}). Cannot install trigger.`);
      Logger.log(`Form trigger installation failed: ${PROP_DOWNTIME_FORM_ID} not set.`);
      return;
  }

  try {
      const form = FormApp.openById(formId); // Check if form exists first
      if (!form) {
          throw new Error(`Form with ID ${formId} not found.`);
      }

      // Delete existing Form Submit triggers for this script
      const triggers = ScriptApp.getProjectTriggers();
      let deletedCount = 0;
      triggers.forEach(trigger => {
          // Ensure trigger source is the form before attempting to get source ID
          if (trigger.getTriggerSource() === ScriptApp.TriggerSource.FORMS &&
              trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT &&
              trigger.getTriggerSourceId() === form.getId()) { // Use form.getId() for comparison
              try {
                  ScriptApp.deleteTrigger(trigger);
                  Logger.log(`Deleted existing Form Submit trigger with ID: ${trigger.getUniqueId()}`);
                  deletedCount++;
              } catch (err) {
                  Logger.log(`Could not delete Form Submit trigger ${trigger.getUniqueId()}: ${err}`);
              }
          }
      });
      Logger.log(`Finished deleting ${deletedCount} existing Form Submit trigger(s).`);

      // Install new Form Submit trigger
      ScriptApp.newTrigger('onFormSubmitHandler_') // In Main.gs
          .forForm(form)
          .onFormSubmit()
          .create();
      Logger.log(`Created new onFormSubmit trigger for form ID: ${formId}`);

      SpreadsheetApp.getUi().alert('Form Submit trigger reinstalled successfully.');
      logAudit_('Reinstall Form Trigger', 'N/A', 'Success'); // In Utilities.gs

  } catch (error) {
      Logger.log(`Error during Form Submit trigger reinstallation: ${error} \nStack: ${error.stack}`);
      logAudit_('Reinstall Form Trigger FAILED', 'N/A', `Error: ${error.message}`); // In Utilities.gs
      SpreadsheetApp.getUi().alert(`Error reinstalling Form Submit trigger: ${error.message}. Check Logs.`);
  }
}

/**
 * Reinstalls ONLY the onEdit trigger for the active spreadsheet.
 * Deletes existing onEdit triggers first to prevent duplicates.
 */
function reinstallOnEditTrigger_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssId = ss.getId();
    const handlerFunctionName = 'onEditHandler_'; // The name of our handler function in Main.gs

    Logger.log(`Attempting to reinstall onEdit trigger for spreadsheet ID: ${ssId}`);

    try {
        // Delete existing onEdit triggers for this script targeting this spreadsheet
        const triggers = ScriptApp.getProjectTriggers();
        let deletedCount = 0;
        triggers.forEach(trigger => {
            // Check event type, source (spreadsheet), and handler function name
            if (trigger.getEventType() === ScriptApp.EventType.ON_EDIT &&
                trigger.getTriggerSource() === ScriptApp.TriggerSource.SPREADSHEETS &&
                trigger.getTriggerSourceId() === ssId &&
                trigger.getHandlerFunction() === handlerFunctionName) {
                try {
                    ScriptApp.deleteTrigger(trigger);
                    Logger.log(`Deleted existing onEdit trigger with ID: ${trigger.getUniqueId()}`);
                    deletedCount++;
                } catch (err) {
                    Logger.log(`Could not delete onEdit trigger ${trigger.getUniqueId()}: ${err}`);
                }
            }
        });
        Logger.log(`Finished deleting ${deletedCount} existing onEdit trigger(s).`);

        // Install new onEdit trigger
        ScriptApp.newTrigger(handlerFunctionName)
            .forSpreadsheet(ss)
            .onEdit()
            .create();
        Logger.log(`Created new onEdit trigger for spreadsheet ID: ${ssId}`);

        SpreadsheetApp.getUi().alert('onEdit trigger reinstalled successfully.');
        logAudit_('Reinstall Edit Trigger', ss.getActiveSheet().getName(), 'Success'); // In Utilities.gs

    } catch (error) {
        Logger.log(`Error during onEdit trigger reinstallation: ${error} \nStack: ${error.stack}`);
        logAudit_('Reinstall Edit Trigger FAILED', ss.getActiveSheet().getName(), `Error: ${error.message}`); // In Utilities.gs
        SpreadsheetApp.getUi().alert(`Error reinstalling onEdit trigger: ${error.message}. Check Logs.`);
    }
}



/**
 * Sets up the time-driven trigger for sendScheduledMessages_.
 * Deletes existing triggers for this function first.
 * This function is called from the Maintenance menu.
 */
function setupScheduledMessagesTrigger_() {
    const handlerFunctionName = 'sendScheduledMessages_'; // In ScheduledMessages.gs
    Logger.log(`Setting up time-driven trigger for ${handlerFunctionName}...`);
    const ui = SpreadsheetApp.getUi();

    try {
        // Delete existing triggers for this function
        const triggers = ScriptApp.getProjectTriggers();
        let deletedCount = 0;
        triggers.forEach(trigger => {
            if (trigger.getHandlerFunction() === handlerFunctionName) {
                try {
                    ScriptApp.deleteTrigger(trigger);
                    Logger.log(`Deleted existing trigger for ${handlerFunctionName} (ID: ${trigger.getUniqueId()})`);
                    deletedCount++;
                } catch (err) {
                     Logger.log(`Could not delete trigger ${trigger.getUniqueId()}: ${err}`);
                }
            }
        });
         Logger.log(`Deleted ${deletedCount} existing trigger(s) for ${handlerFunctionName}.`);

        // Create a new trigger to run every hour
        ScriptApp.newTrigger(handlerFunctionName)
            .timeBased()
            .everyHours(1)
            .create();

        Logger.log(`Successfully created hourly trigger for ${handlerFunctionName}.`);
        ui.alert('Success!', `Hourly trigger set up to run the '${handlerFunctionName}' function.`, ui.ButtonSet.OK);
        logAudit_('Setup Scheduled Message Trigger', 'N/A', 'Success'); // In Utilities.gs

    } catch (error) {
        Logger.log(`Error setting up trigger for ${handlerFunctionName}: ${error}\n${error.stack}`);
         logAudit_('Setup Scheduled Message Trigger FAILED', 'N/A', `Error: ${error.message}`); // In Utilities.gs
        ui.alert('Error', `Could not set up the hourly trigger: ${error.message}. Please check script permissions or set it up manually via Edit > Current project's triggers.`, ui.ButtonSet.OK);
    }
}
