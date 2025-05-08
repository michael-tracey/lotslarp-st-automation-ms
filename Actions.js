/**
 * @OnlyCurrentDoc
 *
 * Functions that perform specific data filling actions in cells.
 */

/**
 * Fills the active cell with data based on the specified type ('herd', 'feed', 'patrol').
 * Replaces placeholders within the selected text. Logs the action.
 * @param {string} type - The type of data ('herd', 'feed', 'patrol').
 */
function fillCellWithData_(type) {
  Logger.log(`Attempting to fill cell with data type: ${type}`);
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const row = cell.getRow();
  const col = cell.getColumn();
  const cellA1 = cell.getA1Notation();
  const sheetName = sheet.getName();

  try {
    const feedData = getFeedListData_(); // In SheetData.gs
    if (!feedData) return; // Error handled in getFeedListData_

    const { headers, values } = feedData;
    let columnIndex;
    let headerName;

    // Determine column index based on type (case-insensitive header matching)
    const lowerCaseHeaders = headers.map(h => String(h).toLowerCase());
    if (type === 'herd') {
      columnIndex = lowerCaseHeaders.indexOf('herd');
      headerName = 'Herd';
    } else if (type === 'feed') {
      columnIndex = lowerCaseHeaders.indexOf('feed');
      headerName = 'Feed';
    } else if (type === 'patrol') {
      columnIndex = lowerCaseHeaders.indexOf('patrol');
      headerName = 'Patrol';
    } else {
      const errorMessage = 'Error: Invalid data type specified for fillCellWithData_.';
      SpreadsheetApp.getUi().alert(errorMessage);
      Logger.log(errorMessage);
      return;
    }

    if (columnIndex === -1) {
        const errorMessage = `Error: Column header "${headerName}" not found (case-insensitive) in '${FEED_LIST_SHEET_NAME}'.`;
        SpreadsheetApp.getUi().alert(errorMessage);
        Logger.log(errorMessage);
        return;
    }

    // Get random value for the specified type
    const columnValues = values.slice(1) // Skip header row
                           .map(row => row[columnIndex])
                           .filter(value => value && String(value).trim() !== ''); // Filter empty values

    if (columnValues.length === 0) {
        const errorMessage = `Error: No data found in the "${headerName}" column of '${FEED_LIST_SHEET_NAME}'.`;
        SpreadsheetApp.getUi().alert(errorMessage);
        Logger.log(errorMessage);
        return;
    }

    let textToInsert = columnValues[Math.floor(Math.random() * columnValues.length)];
    Logger.log(`Selected random value for ${type}: "${textToInsert}"`);

    // Replace other placeholders like [Location], [Target], etc.
    // Iterate through all headers to find potential placeholders
    for (let i = 0; i < headers.length; i++) {
        // Skip the column we already used for the base text
        if (i === columnIndex) continue;

        const header = headers[i];
        if (!header || String(header).trim() === '') continue; // Skip empty headers

        const placeholder = `[${header}]`;
        // Use case-insensitive check for placeholder presence
        if (String(textToInsert).toLowerCase().includes(placeholder.toLowerCase())) {
            Logger.log(`Found potential placeholder: ${placeholder}`);
            const placeholderColValues = values.slice(1)
                                        .map(row => row[i])
                                        .filter(value => value && String(value).trim() !== '');

            if (placeholderColValues.length > 0) {
                const replacement = placeholderColValues[Math.floor(Math.random() * placeholderColValues.length)];
                Logger.log(`Replacing ${placeholder} with random value: "${replacement}"`);
                // Use regex for global, case-insensitive replacement
                const regex = new RegExp(`\\[${escapeRegex_(header)}\\]`, 'gi'); // In Utilities.gs
                textToInsert = String(textToInsert).replace(regex, replacement);
                Logger.log(`Text after replacement: "${textToInsert}"`);
            } else {
                Logger.log(`No replacement values found for placeholder ${placeholder}. Leaving it as is.`);
            }
        }
    }

    cell.setValue(textToInsert);
    Logger.log(`Successfully filled cell ${cellA1} with ${type} data.`);
    logAudit_(`Fill ${type} Data`, sheetName, `Cell: ${cellA1}`); // Log success // In Utilities.gs

  } catch (error) {
    Logger.log(`Error filling cell with ${type} data: ${error}\nStack: ${error.stack}`);
    logAudit_(`Fill ${type} Data FAILED`, sheetName, `Cell: ${cellA1}, Error: ${error.message}`); // Log error // In Utilities.gs
    SpreadsheetApp.getUi().alert(`Error generating ${type} data: ${error.message}`);
  }
}

/**
 * Shows a dialog to get the action power level, displaying character info first.
 * This function is now the primary entry point called by menu items.
 * It gathers context and shows the dialog. The actual fill logic is in executeInfluenceFill.
 */
function openInfluencePowerDialog_() {
    Logger.log(`Opening Influence power dialog...`);
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const cell = sheet.getActiveCell();
    if (!cell) {
        ui.alert("Please select a target cell first.");
        return;
    }
    // ** Ensure operating on the response row (ODD row) for character name lookup **
    const activeRow = cell.getRow();
    const responseRowForName = (activeRow % 2 !== 0) ? activeRow : (activeRow > 1 ? activeRow - 1 : null); // If even, use row above, unless it's row 1

     if (!responseRowForName) {
         ui.alert('Cannot determine character row. Please select a cell within a downtime entry response row (usually odd-numbered > 1).');
         return;
     }

    const cellA1 = cell.getA1Notation(); // Cell to *fill* is still the active cell
    const sheetName = sheet.getName();

    // --- Get Character Name (from Col E of the determined response row) ---
    const characterNameCell = sheet.getRange(responseRowForName, CHARACTER_NAME_COL); // In Constants.gs
    const characterName = characterNameCell.getValue();

    if (!characterName || typeof characterName !== 'string' || characterName.trim() === '') {
        ui.alert(`Cannot determine character name from cell ${characterNameCell.getA1Notation()}. Please ensure a valid character name is present in the response row.`);
        Logger.log(`Influence fill failed: Character name cell ${characterNameCell.getA1Notation()} is empty or invalid.`);
        return;
    }
    const cleanCharacterName = characterName.trim();
    Logger.log(`Character name found: "${cleanCharacterName}" from cell ${characterNameCell.getA1Notation()}`);

    // --- Lookup Specs and Calculate Totals from 'Influences' Sheet ---
    let eliteSpecs = [];
    let uwSpecs = [];
    let totalEliteActions = 0;
    let totalUwActions = 0;
    try {
        const influencesSheet = ss.getSheetByName(INFLUENCES_SHEET_NAME); // In Constants.gs
        if (!influencesSheet) {
            throw new Error(`Sheet "${INFLUENCES_SHEET_NAME}" not found.`);
        }
        const influencesData = influencesSheet.getDataRange().getValues();
        // Col A (idx 0): Elite Char, B(1): Elite Spec, D(3): UW Char, E(4): UW Spec
        influencesData.slice(1).forEach(row => {
            const eliteChar = String(row[0] || '').trim();
            const eliteSpec = String(row[1] || '').trim();
            const uwChar = String(row[3] || '').trim(); // Use column D (index 3)
            const uwSpec = String(row[4] || '').trim(); // Use column E (index 4)

            if (eliteChar.toLowerCase() === cleanCharacterName.toLowerCase() && eliteSpec) {
                eliteSpecs.push(eliteSpec);
            }
            if (uwChar.toLowerCase() === cleanCharacterName.toLowerCase() && uwSpec) {
                uwSpecs.push(uwSpec);
            }
        });
        totalEliteActions = eliteSpecs.length;
        totalUwActions = uwSpecs.length;
        Logger.log(`Found Specs - Elite (${totalEliteActions}): [${eliteSpecs.join(', ')}], UW (${totalUwActions}): [${uwSpecs.join(', ')}]`);

    } catch (error) {
         Logger.log(`Error looking up influences for ${cleanCharacterName}: ${error}`);
         ui.alert(`Could not look up character influences: ${error.message}`);
         return; // Stop if we can't get spec info
    }


    // Pass necessary context to the dialog's client-side script
    const context = {
        sheetName: sheetName,
        cellA1: cellA1, // The cell where the result will be placed
        characterNameRow: responseRowForName, // Pass the row where the name was found
        // Pass sheet names for Elite/UW actions
        eliteInfluenceSheetName: ELITE_INFLUENCES_SHEET_NAME, // In Constants.gs
        uwInfluenceSheetName: UW_INFLUENCES_SHEET_NAME,     // In Constants.gs
        // Pass character info for display
        characterName: cleanCharacterName,
        eliteSpecs: eliteSpecs.sort(),
        uwSpecs: uwSpecs.sort(),
        totalEliteActions: totalEliteActions,
        totalUwActions: totalUwActions
    };

    try {
        const template = HtmlService.createTemplateFromFile('InfluencePowerDialog');
        template.context = context; // Make context available in HTML scriptlets

        const htmlOutput = template.evaluate()
            .setWidth(650) // Keep wider width
            .setHeight(450); // Keep increased height

        ui.showModalDialog(htmlOutput, `Select Influence Action Power for ${cleanCharacterName}`);

    } catch (error) {
        Logger.log(`Error showing influence power dialog: ${error}\nStack: ${error.stack}`);
        ui.alert(`Error opening dialog: ${error.message}`);
    }
}


/**
 * Executes the influence data filling after power level is selected from the dialog.
 * Tracks skipped items and returns a *summary* report object to the client-side.
 * @param {number} actionPower The power level selected by the user.
 * @param {object} context Context object passed from the dialog.
 * @returns {object} Report object: { outputValue: string, skippedTooOld: number, skippedBlocksCount: number, skippedNoColonCount: number, skippedStartsWithTwo: number }
 */
function executeInfluenceFill(actionPower, context) {
    // Destructure context, including the character name we already found
    const { sheetName, cellA1, characterNameRow, influenceSheetName, influenceTypeLabel, characterName } = context;
    Logger.log(`Executing ${influenceTypeLabel} Influence fill for ${cellA1} on sheet ${sheetName} with Power: ${actionPower} for Character: ${characterName}`); // Log entry point
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
         Logger.log(`Error in executeInfluenceFill: Sheet "${sheetName}" not found.`);
         throw new Error(`Sheet "${sheetName}" could not be found.`); // Throw error back to dialog
    }
    const cell = sheet.getRange(cellA1);

    // Character name is already known from context
    if (!characterName) {
         Logger.log(`${influenceTypeLabel} Influence fill failed: Character name missing from context.`);
         throw new Error('Character name was missing in context.');
    }

    // Initialize report variables
    let acceptedResults = [];
    let skippedTooOld = 0;
    let skippedBlocks = []; // Collect details for logging, return count
    let skippedNoColon = {}; // Collect details for logging, return count
    let skippedStartsWithTwo = 0;
    let skippedSpecMismatch = 0; // Add counter for spec mismatch
    let reportObject = null; // Initialize report object

    try {
        // --- Get Character's Influences (needed again for filtering) ---
        const influencesSheet = ss.getSheetByName(INFLUENCES_SHEET_NAME); // In Constants.gs
        if (!influencesSheet) {
             throw new Error(`Sheet "${INFLUENCES_SHEET_NAME}" not found.`);
        }
        const influencesData = influencesSheet.getDataRange().getValues();
        // Col A (idx 0): Elite Char, B(1): Elite Spec, D(3): UW Char, E(4): UW Spec
        const characterInfluences = influencesData
            .slice(1)
            .filter(row => {
                 const eliteChar = String(row[0] || '').trim().toLowerCase();
                 const uwChar = String(row[3] || '').trim().toLowerCase(); // Check Col D (idx 3)
                 const cleanCharNameLower = characterName.toLowerCase();
                 // Check if character matches in the correct column for the type
                 return (influenceTypeLabel === 'Elite' && eliteChar === cleanCharNameLower && row[1]) ||
                        (influenceTypeLabel === 'Underworld' && uwChar === cleanCharNameLower && row[4]); // Check Col E (idx 4) for spec
            })
            .map(row => String(influenceTypeLabel === 'Elite' ? row[1] : row[4]).trim()); // Get spec from Col B or E


        if (characterInfluences.length === 0) {
            Logger.log(`No relevant ${influenceTypeLabel} influences found for character "${characterName}" in sheet "${INFLUENCES_SHEET_NAME}". Setting fallback text.`);
            cell.setValue(INFLUENCE_FALLBACK_TEXT);
            logAudit_(`Fill ${influenceTypeLabel} Influence`, sheetName, `Cell: ${cellA1}, Power: ${actionPower}, Result: Fallback (No Char Influences)`); // In Utilities.gs
            // Return empty report structure
            reportObject = { outputValue: INFLUENCE_FALLBACK_TEXT, skippedTooOld: 0, skippedBlocksCount: 0, skippedNoColonCount: 0, skippedStartsWithTwo: 0, skippedSpecMismatch: 0 };
            Logger.log('Returning report data (No Char Influences):', JSON.stringify(reportObject));
            return reportObject;
        }
        Logger.log(`Found relevant ${influenceTypeLabel} influences for ${characterName}: [${characterInfluences.join(', ')}]`);

        // --- Build Regex Pattern ---
        const escapedInfluences = characterInfluences.map(inf => escapeRegex_(inf)); // In Utilities.gs
        const influenceRegexPattern = escapedInfluences.join('|');
        // Make regex case-insensitive for matching spec names
        const influenceRegex = new RegExp(`^(?:${influenceRegexPattern})$`, 'i');
        Logger.log(`Built influence regex: ${influenceRegex}`);

        // --- Get and Filter Specific Influence Data ---
        const specificInfluenceSheet = ss.getSheetByName(influenceSheetName);
        if (!specificInfluenceSheet) {
             throw new Error(`Sheet "${influenceSheetName}" not found.`);
        }
        const specificInfluenceData = specificInfluenceSheet.getDataRange().getValues();

        // Indices based on assumed format: Date(A), Age(B), Character(C), Specialization(D), Action(E), Details(F), Blocks?(G), Output(H)
        const colB_Index = 1; // Age <= 95
        const colD_Index = 3; // Specialization MATCHES influence
        const colG_Index = 6; // Blocks <= actionPower + 1
        const colH_Index = 7; // Output (Result)

        Logger.log(`Starting filter loop for ${influenceSheetName}. Checking ${specificInfluenceData.length - 1} rows.`);
        for (let i = 1; i < specificInfluenceData.length; i++) {
            const row = specificInfluenceData[i];
            const rowNum = i + 1; // For logging

            try {
                const valueB_Age = parseFloat(row[colB_Index]);
                const valueG_Blocks = parseFloat(row[colG_Index]);
                const valueD_Spec = String(row[colD_Index]).trim();
                const valueH_Output = row[colH_Index];
                const actionName = String(row[INFL_ACTION_COL_IDX_] || 'Unknown Action').trim(); // Get action name for logging skipped blocks

                // Logger.log(`--- Row ${rowNum} ---`);
                // Logger.log(`Raw Values: Age='${row[colB_Index]}', Spec='${valueD_Spec}', Blocks='${row[colG_Index]}', Output='${valueH_Output}'`);

                // 1. Check Age
                const conditionB = !isNaN(valueB_Age) && valueB_Age <= 95;
                // Logger.log(`Check Age (<= 95): Value=${valueB_Age}, Result=${conditionB}`);
                if (!conditionB) {
                    skippedTooOld++;
                    // Logger.log(`Row ${rowNum}: SKIPPED: Age check failed.`);
                    continue; // Go to next row
                }

                // 2. Check Blocks
                const neededBlocks = actionPower + 1;
                const conditionG = !isNaN(valueG_Blocks) && valueG_Blocks <= neededBlocks;
                // Logger.log(`Check Blocks (<= ${neededBlocks}): Value=${valueG_Blocks}, Result=${conditionG}`);
                if (!conditionG) {
                    skippedBlocks.push({ name: actionName, blocks: valueG_Blocks, needed: neededBlocks });
                    // Logger.log(`Row ${rowNum}: SKIPPED: Blocks check failed.`);
                    continue; // Go to next row
                }

                // 3. Check Output Format (if output exists)
                let outputString = null;
                let containsColon = false;
                let startsWithTwo = false;
                if (valueH_Output && typeof valueH_Output === 'string') {
                    outputString = String(valueH_Output).trim();
                    if (outputString) { // Only check non-empty strings
                       containsColon = outputString.includes(':');
                       startsWithTwo = outputString.charAt(0) === '2';
                    }
                }
                // Logger.log(`Check Output Format: Value='${outputString}', ContainsColon=${containsColon}, StartsWithTwo=${startsWithTwo}`);
                if (!outputString || !containsColon) {
                     if(outputString){ // Only count if there was something to check
                       skippedNoColon[outputString] = (skippedNoColon[outputString] || 0) + 1;
                       // Logger.log(`Row ${rowNum}: SKIPPED: Output missing colon.`);
                     } else {
                        // Logger.log(`Row ${rowNum}: SKIPPED: Output column is empty or not a string.`);
                     }
                     continue; // Go to next row
                }
                if (startsWithTwo) {
                    skippedStartsWithTwo++;
                    // Logger.log(`Row ${rowNum}: SKIPPED: Output starts with '2'.`);
                    continue; // Go to next row
                }

                // 4. Check Specialization Match
                const conditionD = influenceRegex.test(valueD_Spec);
                // Logger.log(`Check Spec Match (${influenceRegex}): Value='${valueD_Spec}', Result=${conditionD}`);
                if (!conditionD) {
                    skippedSpecMismatch++; // Increment new counter
                    // Logger.log(`Row ${rowNum}: SKIPPED: Specialization mismatch.`);
                    continue; // Go to next row
                }

                // If all checks passed:
                acceptedResults.push(outputString);
                Logger.log(`Row ${rowNum}: ACCEPTED. Added value: "${outputString}"`);

            } catch(rowError) {
                Logger.log(`Skipping row ${rowNum} in ${influenceSheetName} due to processing error: ${rowError}`);
            }
        } // End for loop
        Logger.log(`Finished filter loop for ${influenceSheetName}. Found ${acceptedResults.length} accepted results.`);

        // --- Format and Set Output ---
        let outputValue;
        let logDetails;
        if (acceptedResults.length === 0) {
            outputValue = INFLUENCE_FALLBACK_TEXT;
            logDetails = `Cell: ${cellA1}, Power: ${actionPower}, Result: Fallback (No Matches in ${influenceSheetName} after filters)`;
            Logger.log(`No matching ${influenceTypeLabel} influence results found after filtering. Setting fallback text.`);
        } else {
            outputValue = acceptedResults.join('\n');
            logDetails = `Cell: ${cellA1}, Power: ${actionPower}, Result Count: ${acceptedResults.length}`;
            Logger.log(`Found ${acceptedResults.length} ${influenceTypeLabel} results after filtering. Joined output:\n${outputValue}`);
        }

        cell.setValue(outputValue);
        logAudit_(`Fill ${influenceTypeLabel} Influence`, sheetName, logDetails); // Log success
        Logger.log(`Successfully set value for cell ${cellA1}.`);

        // --- Construct Simplified Report Object ---
        reportObject = {
            outputValue: outputValue,
            skippedTooOld: skippedTooOld,
            skippedBlocks: skippedBlocks,
            skippedNoColon: skippedNoColon,
            skippedStartsWithTwo: skippedStartsWithTwo,
            skippedSpecMismatch: skippedSpecMismatch
        };
        Logger.log(reportObject);
        // Logger.log('Returning detailed report data:', JSON.stringify(reportObject)); // Log before returning
        return reportObject;

    } catch (error) {
        Logger.log(`Error executing ${influenceTypeLabel} Influence fill: ${error}\nStack: ${error.stack}`);
        logAudit_(`Fill ${influenceTypeLabel} Influence FAILED`, sheetName, `Cell: ${cellA1}, Power: ${actionPower}, Error: ${error.message}`); // Log error
        // Throw error back to client-side to be displayed
        throw new Error(`An error occurred: ${error.message}`);
    }
}
