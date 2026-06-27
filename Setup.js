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

    // --- Guard: refuse to run on an already-initialised spreadsheet ---
    // The Characters sheet is the master sheet. If it already exists, this
    // spreadsheet is already set up and may hold live data, so we abort rather
    // than risk overwriting anything.
    if (ss.getSheetByName(CHARACTER_SHEET_NAME)) {
        ui.alert(
            'Initialisation Blocked',
            `A master "${CHARACTER_SHEET_NAME}" sheet already exists, so this spreadsheet ` +
            `appears to already be initialised.\n\n` +
            `Initialisation has been aborted to protect existing data. If you really do want ` +
            `to start over, manually delete the "${CHARACTER_SHEET_NAME}" sheet first and run ` +
            `this again.`,
            ui.ButtonSet.OK
        );
        Logger.log(`initialiseProject_ aborted: master "${CHARACTER_SHEET_NAME}" sheet already exists.`);
        return;
    }

    // --- Warning / Confirmation ---
    const confirmation = ui.alert(
        'Initialise Project',
        '⚠️ WARNING: This is intended for a fresh, empty spreadsheet.\n\n' +
        'It will create all the sheets the project needs (Characters, NPCs, Influences, ' +
        'Narrators, Log, and more). Any existing sheet that happens to share one of these ' +
        'names will keep its data, but its layout may no longer match what the script ' +
        'expects.\n\n' +
        'Are you sure you want to continue?',
        ui.ButtonSet.YES_NO
    );

    if (confirmation !== ui.Button.YES) {
        ui.alert('Initialisation cancelled.');
        return;
    }

    // --- Sheet Creation ---
    // Header layouts must stay consistent with the column constants in Constants.js
    // (e.g. the Characters sheet reads Approval Status at col 8, Webhook at col 24,
    // Discord Channel ID/Name at cols 25/26).
    const fillReplace = getFillReplaceSampleData_(); // In FillReplaceData.js — headers + sample rows
    const sheetsToCreate = [
        { name: CHARACTER_SHEET_NAME, headers: ['Character Name', "Who's Who Clan", 'Clan/Bloodline', 'Social Age', 'Court Position', 'Player Name', 'BG Age', 'Discord Name', 'Short form', 'Approved', 'Yorick', 'Long form', 'Character Pronons', 'Player Pronouns', 'Public LV', 'Unrepresented', 'Caucuses with', 'Public Notes', 'Discord Channel', 'Long Form', 'Location/Equipment', 'Lines and Veils', 'Playlist', 'Webhook', 'Channel ID', 'Channel Name', 'Email'] },
        { name: NPC_SHEET_NAME, headers: ['Character Name', "Who's Who Clan", 'Clan', 'Social Age', 'Court Position', 'Player Name', 'BG Age', 'Discord Name', 'Short form', 'Approved', 'Yorick', 'Long form', 'Character Pronons', 'Player Pronouns', "Who's Who LV", 'Unrepresented', 'Caucuses with', 'Public Notes', 'Discord Channel', 'Next step', 'Location/Equipment', 'Lines and Veils', 'Playlist', 'Webhook', 'avatar_url'] },
        { name: INFLUENCES_SHEET_NAME, headers: ['Elite Character', 'Elite Spec', '', 'UW Character', 'UW Spec', '', 'Source'] },
        { name: ELITE_INFLUENCES_SHEET_NAME, headers: ['Date', 'Age', 'Character', 'Specialization', 'Action', 'Details', 'Any blocks?', 'Output'] },
        { name: UW_INFLUENCES_SHEET_NAME, headers: ['Date', 'Age', 'Character', 'Specialization', 'Action', 'Details', 'Any blocks?', 'Output'] },
        { name: FILL_REPLACE_SHEET_NAME, headers: fillReplace.headers, seedRows: fillReplace.rows },
        { name: AUDIT_LOG_SHEET_NAME, headers: ['Timestamp', 'User', 'Action', 'Sheet Name', 'Details'] },
        { name: SCHEDULED_MSG_SHEET_NAME, headers: ['Sent', 'Send After Timestamp', 'Target Channel', 'Sender (Optional)', 'Message Body', 'Send Log'] },
        { name: 'narrators', headers: ['Name', 'Email', 'Hex Color', 'Username', 'Password', 'Role'] }
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
                sheet.setFrozenRows(1);
                sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
            }
            // Seed sample data rows if this sheet provides any (e.g. fill-replace-data).
            if (sheetInfo.seedRows && sheetInfo.seedRows.length > 0) {
                sheet.getRange(2, 1, sheetInfo.seedRows.length, headers.length).setValues(sheetInfo.seedRows);
            }
            sheetsCreated++;
        } else {
            sheetsExist++;
        }
    });

    logAudit_('Initialise Project', 'N/A', `Created ${sheetsCreated} sheet(s); ${sheetsExist} already existed.`);

    // --- Final Report ---
    let summary = '';
    if (sheetsCreated > 0) {
        summary += `Created ${sheetsCreated} new sheets. `;
    }
    if (sheetsExist > 0) {
        summary += `${sheetsExist} sheets already existed (left untouched). `;
    }

    ui.alert('Initialisation Complete', summary, ui.ButtonSet.OK);
}
