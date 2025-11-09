/**
 * @OnlyCurrentDoc
 *
 * Functions related to project setup and initialization.
 */

/**
 * Creates all necessary sheets for the project to function correctly.
 * This function is intended to be called from the "Initialise Project" menu item.
 */
function initialiseProject_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    // --- Confirmation ---
    const confirmation = ui.alert(
        'Initialise Project',
        'This will create all the necessary sheets for the project. This includes sheets for characters, NPCs, influences, and more. Are you sure you want to continue?',
        ui.ButtonSet.YES_NO
    );

    if (confirmation !== ui.Button.YES) {
        ui.alert('Initialisation cancelled.');
        return;
    }

    // --- Sheet Creation ---
    const sheetsToCreate = [
        { name: CHARACTER_SHEET_NAME, headers: ['Character Name', 'Player Name', 'Chronicle', 'Ancilla', 'Email', 'User ID', 'XP', 'Discord Handle', 'Short Form Link', 'Status', 'Image URL', 'Long Form Link', 'Approval Status', 'Approval Timestamp', 'Approval Notes', 'Influence 1', 'Influence 2', 'Influence 3', 'Influence 4', 'Influence 5', 'Influence 6', 'Influence 7', 'Influence 8', 'Influence 9', 'Influence 10', 'Webhook'] },
        { name: NPC_SHEET_NAME, headers: ['NPC Name', 'Chronicle', 'Status', 'Notes'] },
        { name: ELITE_INFLUENCES_SHEET_NAME, headers: ['Influence Name', 'Description'] },
        { name: UNDERWORLD_INFLUENCES_SHEET_NAME, headers: ['Influence Name', 'Description'] },
        { name: RESOURCES_SHEET_NAME, headers: ['Resource Name', 'Description'] },
        { name: DOWNTIME_ACTIONS_SHEET_NAME, headers: ['Action Name', 'Description'] },
        { name: CHARACTER_LOG_SHEET_NAME, headers: ['Timestamp', 'Character Name', 'Action', 'Notes'] },
        { name: AUDIT_LOG_SHEET_NAME, headers: ['Timestamp', 'User', 'Action', 'Sheet Name', 'Details'] },
        { name: SCHEDULED_MSG_SHEET_NAME, headers: ['ID', 'Message', 'Channel', 'Timestamp', 'Status'] },
        { name: WEBHOOKS_SHEET_NAME, headers: ['Name', 'URL'] },
        { name: 'narrators', headers: ['Name', 'Email', 'Hex Color', 'Username', 'Password'] }
    ];

    let sheetsCreated = 0;
    let sheetsExist = 0;

    sheetsToCreate.forEach(sheetInfo => {
        let sheet = ss.getSheetByName(sheetInfo.name);
        if (!sheet) {
            sheet = ss.insertSheet(sheetInfo.name);
            const headers = sheetInfo.headers || [];
            if (headers.length > 0) {
                sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
            }
            sheetsCreated++;
        } else {
            sheetsExist++;
        }
    });

    // --- Final Report ---
    let summary = '';
    if (sheetsCreated > 0) {
        summary += `Created ${sheetsCreated} new sheets. `;
    }
    if (sheetsExist > 0) {
        summary += `${sheetsExist} sheets already existed. `;
    }

    ui.alert('Initialisation Complete', summary, ui.ButtonSet.OK);
}
