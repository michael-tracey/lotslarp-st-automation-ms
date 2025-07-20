/**
 * @OnlyCurrentDoc
 *
 * This file contains a comprehensive function to test various script functionalities
 * and report on required permissions.
 */

/**
 * Tests various script functionalities to check for necessary permissions.
 * Displays a summary of results in a UI alert.
 */
function testAllFeaturesAndPermissions() {
  const ui = SpreadsheetApp.getUi();
  let results = [];
  let overallStatus = '✅ ALL TESTED FEATURES PASSED PERMISSION CHECKS ✅';
  let hasFailures = false;

  results.push('--- Apps Script Permission Test Results ---');
  results.push('');

  // --- 1. Test Logger.log() (Basic Logging) ---
  try {
    Logger.log('Permission Test: Attempting a basic log entry.');
    results.push('✅ Logger.log(): OK');
  } catch (e) {
    results.push(`❌ Logger.log(): FAILED - ${e.message}. Check 'Script Editor > Executions' for details.`);
    hasFailures = true;
  }

  // --- 2. Test SpreadsheetApp (Basic Read/Write) ---
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const tempRange = sheet.getRange('A1'); // Just try to get a range
    // Attempt a non-destructive read
    const value = tempRange.getValue();
    results.push('✅ SpreadsheetApp (Read): OK (Accessed A1)');
  } catch (e) {
    results.push(`❌ SpreadsheetApp (Read): FAILED - ${e.message}. Check 'https://www.googleapis.com/auth/spreadsheets' scope.`);
    hasFailures = true;
  }

  // --- 3. Test GmailApp (Sending Email) ---
  try {
    // Attempt to send a dummy email to the active user
    const userEmail = Session.getActiveUser().getEmail();
    if (userEmail && userEmail !== 'unknown@example.com') {
      GmailApp.sendEmail(userEmail, 'Apps Script Permission Test Email', 'This is a test email from your Apps Script to check email sending permissions. You can ignore this.');
      results.push(`✅ GmailApp (Send Email): OK (Test email sent to ${userEmail})`);
    } else {
      results.push('⚠️ GmailApp (Send Email): SKIPPED - Could not determine active user email for test.');
    }
  } catch (e) {
    results.push(`❌ GmailApp (Send Email): FAILED - ${e.message}. Check 'https://www.googleapis.com/auth/script.send_mail' or 'https://www.googleapis.com/auth/gmail.send' scope.`);
    hasFailures = true;
  }

  // --- 4. Test UrlFetchApp (Discord Webhooks) ---
  try {
    // This will call sendTestMessageToStorytellers_ which uses UrlFetchApp
    // It will also log its own success/failure.
    sendTestMessageToStorytellers_(); // This function is in Discord.js
    results.push('✅ UrlFetchApp (Discord Webhook): OK (Test message attempt logged by Discord.js)');
  } catch (e) {
    results.push(`❌ UrlFetchApp (Discord Webhook): FAILED - ${e.message}. Check 'https://www.googleapis.com/auth/script.external_request' scope and webhook configuration.`);
    hasFailures = true;
  }

  // --- 5. Test PropertiesService (Script Properties) ---
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('permissionTestKey', 'testValue');
    const testValue = scriptProperties.getProperty('permissionTestKey');
    scriptProperties.deleteProperty('permissionTestKey');
    results.push('✅ PropertiesService: OK (Read/Write/Delete property)');
  } catch (e) {
    results.push(`❌ PropertiesService: FAILED - ${e.message}. Check 'https://www.googleapis.com/auth/script.properties' scope.`);
    hasFailures = true;
  }

  // --- 6. Test ScriptApp (Trigger Management) ---
  try {
    // Attempt to list triggers (read-only check)
    const triggers = ScriptApp.getProjectTriggers();
    results.push(`✅ ScriptApp (Triggers Read): OK (Found ${triggers.length} triggers)`);
  } catch (e) {
    results.push(`❌ ScriptApp (Triggers Read): FAILED - ${e.message}. Check 'https://www.googleapis.com/auth/script.script.projects' scope.`);
    hasFailures = true;
  }

  // --- 7. Test FormApp (if form ID is configured) ---
  try {
    const formId = PropertiesService.getScriptProperties().getProperty(PROP_DOWNTIME_FORM_ID);
    if (formId && formId !== 'YOUR_GOOGLE_FORM_ID_HERE') {
      // Attempt to open the form (read-only check)
      const form = FormApp.openById(formId);
      results.push(`✅ FormApp (Read): OK (Accessed form ID: ${form.getId()})`);
    } else {
      results.push('⚠️ FormApp (Read): SKIPPED - Form ID not configured in Script Properties.');
    }
  } catch (e) {
    results.push(`❌ FormApp (Read): FAILED - ${e.message}. Check 'https://www.googleapis.com/auth/forms' scope and if the Form ID is correct.`);
    hasFailures = true;
  }

  results.push('');
  if (hasFailures) {
    overallStatus = '❌ SOME PERMISSION CHECKS FAILED ❌';
    results.push('If any checks failed, you likely need to re-authorize the script with the necessary permissions.');
    results.push('Go to "Project Settings" (gear icon) > "OAuth consent screen" to ensure scopes are listed.');
    results.push('Then, run a function that requires the missing scope (e.g., "onOpen" or "reinstallFormTrigger_") to trigger the authorization prompt.');
  } else {
    results.push(overallStatus);
  }

  ui.alert('Permission Test Summary', results.join('\n'), ui.ButtonSet.OK);
  Logger.log('Permission Test Summary:\n' + results.join('\n'));
}
