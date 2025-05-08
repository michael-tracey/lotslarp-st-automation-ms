/**
 * @OnlyCurrentDoc
 *
 * Functions related to reading, processing, and validating data from various sheets.
 */

/**
 * Fetches data from the 'feed list' sheet.
 * @returns {{headers: Array<string>, values: Array<Array<string>>}|null} Object with headers and values, or null if sheet not found.
 */
function getFeedListData_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const feedListSheet = ss.getSheetByName(FEED_LIST_SHEET_NAME);

    if (!feedListSheet) {
        const errorMessage = `Error: The sheet "${FEED_LIST_SHEET_NAME}" was not found.`;
        SpreadsheetApp.getUi().alert(errorMessage);
        Logger.log(errorMessage);
        return null;
    }

    try {
        const headers = feedListSheet.getRange(1, 1, 1, feedListSheet.getLastColumn()).getValues()[0];
        const values = feedListSheet.getDataRange().getValues(); // Includes header row
        Logger.log(`Retrieved data from '${FEED_LIST_SHEET_NAME}'. Headers: [${headers.join(', ')}]. Found ${values.length - 1} data rows.`);
        return { headers, values };
    } catch (error) {
        const errorMessage = `Error reading data from "${FEED_LIST_SHEET_NAME}": ${error.message}`;
        SpreadsheetApp.getUi().alert(errorMessage);
        Logger.log(errorMessage);
        return null;
    }
}

/**
 * Gets unique character names from the specified sheet ('Characters' or 'NPC').
 * @param {string} type - 'Characters' or 'NPCs'.
 * @returns {Array<string>|null} Sorted array of unique names, or null if sheet/type is invalid.
 */
function getUniqueCharacterNames_(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetName;
  if (type === 'Characters') {
    sheetName = CHARACTER_SHEET_NAME;
  } else if (type === 'NPCs') {
    sheetName = NPC_SHEET_NAME;
  } else {
    Logger.log(`Invalid type "${type}" specified for getUniqueCharacterNames_.`);
    return null;
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet "${sheetName}" not found for getUniqueCharacterNames_.`);
    return null;
  }

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return []; // Empty sheet
    // Assumes names are in the first column (A)
    const values = sheet.getRange(1, CHAR_SHEET_NAME_COL, lastRow, 1).getValues();
    const uniqueNames = [...new Set(values.flat())] // Use Set for uniqueness
                         .filter(name => name && String(name).trim() !== '') // Filter out empty/whitespace names
                         .sort();
    Logger.log(`Retrieved ${uniqueNames.length} unique names from sheet '${sheetName}'.`);
    return uniqueNames;
  } catch (error) {
      Logger.log(`Error retrieving unique names from '${sheetName}': ${error}`);
      return null;
  }
}

/**
 * Retrieves the Discord webhook URL for a specific character name.
 * @param {string} characterName The name of the character.
 * @returns {string|false} The webhook URL string if found, otherwise false.
 */
function getCharacterWebhook_(characterName) {
  if (!characterName || String(characterName).trim() === '') return false;

  const characterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHARACTER_SHEET_NAME);
  if (!characterSheet) {
    Logger.log(`Cannot find character sheet: ${CHARACTER_SHEET_NAME}`);
    return false;
  }

  try {
    const lastRow = characterSheet.getLastRow();
    if (lastRow < 1) return false; // Empty sheet

    // Get names and webhooks in one go
    const data = characterSheet.getRange(1, CHAR_SHEET_NAME_COL, lastRow, CHAR_SHEET_WEBHOOK_COL - CHAR_SHEET_NAME_COL + 1).getValues();
    const nameColIndex = 0; // Index within the fetched data array
    const webhookColIndex = CHAR_SHEET_WEBHOOK_COL - CHAR_SHEET_NAME_COL; // Index within the fetched data array

    const cleanTargetName = String(characterName).toLowerCase().trim();

    for (let i = 0; i < data.length; i++) {
        const currentName = data[i][nameColIndex];
        if (currentName && String(currentName).toLowerCase().trim() === cleanTargetName) {
            const webhook = data[i][webhookColIndex];
            if (webhook && String(webhook).trim().startsWith('https://discord')) {
                Logger.log(`Webhook found for character '${characterName}'.`);
                return String(webhook).trim();
            } else {
                Logger.log(`Webhook entry found for '${characterName}' but it is invalid or empty.`);
                return false; // Found name but webhook is invalid/empty
            }
        }
    }

    Logger.log(`Webhook not found for character '${characterName}'.`);
    return false; // Name not found
  } catch (error) {
      Logger.log(`Error retrieving character webhook for "${characterName}": ${error}`);
      return false;
  }
}


/**
 * Validates a character name against the 'Characters' sheet and checks for a webhook.
 * @param {string} name The character name to test.
 * @returns {{isMatch: boolean, hasWebhook: boolean}} Object indicating if the name matches and if a webhook exists.
 */
function validateCharacterName_(name) {
  if (!name || String(name).trim() === '') {
      return { isMatch: false, hasWebhook: false }; // No name provided
  }

  const cleanName = String(name).toLowerCase().trim();
  // Cache unique names? For frequent calls, this could speed things up, but might get stale.
  // For now, fetch fresh each time.
  const uniqueNames = getUniqueCharacterNames_('Characters');

  if (!uniqueNames) {
      Logger.log("Validation failed: Could not retrieve unique character names.");
      // Treat as no match if the list isn't available
      return { isMatch: false, hasWebhook: false };
  }

  const isMatch = uniqueNames.some(uniqueName => String(uniqueName).toLowerCase().trim() === cleanName);
  let hasWebhook = false;
  if (isMatch) {
      hasWebhook = getCharacterWebhook_(name) !== false;
  }

  Logger.log(`Validation for name "${name}": Match=${isMatch}, Webhook=${hasWebhook}`);
  return { isMatch, hasWebhook };
}

/**
 * Validates the character name in a given cell and updates its background color.
 * @param {GoogleAppsScript.Spreadsheet.Range} cell The cell containing the character name.
 */
function validateCharacterNameCell_(cell) {
    const characterName = cell.getValue();
    const validationResult = validateCharacterName_(characterName);
    applyValidationHighlighting_(cell, validationResult);
}

/**
 * Applies background highlighting to a cell based on validation results using the new color scheme.
 * Orange = No Match; Red = Match, No Webhook; Green = Match + Webhook.
 * @param {GoogleAppsScript.Spreadsheet.Range} cell The cell to highlight.
 * @param {{isMatch: boolean, hasWebhook: boolean}} validationResult The result from validateCharacterName_.
 */
function applyValidationHighlighting_(cell, validationResult) {
    if (validationResult.isMatch && validationResult.hasWebhook) {
        cell.setBackground(COLOR_VALID); // Green: Match + Webhook
    } else if (validationResult.isMatch && !validationResult.hasWebhook) {
        cell.setBackground(COLOR_NO_WEBHOOK); // Red: Match, No Webhook
    } else { // Includes !validationResult.isMatch and cases where name is empty/invalid
        cell.setBackground(COLOR_NO_MATCH); // Orange: No Match
    }
}

/**
 * Applies conditional formatting rule to highlight non-empty cells in response columns (F onwards).
 * NOTE: This does NOT handle the name validation coloring in Column E.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to apply formatting to.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The range to apply formatting to (typically F#:LastCol#).
 */
function applyConditionalFormatting_(sheet, range) {
  try {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenCellNotEmpty()
      .setBackground("#CCE5FF") // Light blue for non-empty responses
      .setRanges([range])
      .build();
    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
    Logger.log(`Applied conditional formatting to range ${range.getA1Notation()}.`);
  } catch (error) {
      Logger.log(`Error applying conditional formatting to ${range.getA1Notation()}: ${error}`);
  }
}


/**
 * Gets downtime completion statistics and keyword counts from the sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to analyze.
 * @returns {object|null} An object with statistics, or null if data is invalid.
 */
function getDowntimeCompletionData_(sheet) {
    if (!sheet) {
        Logger.log('getDowntimeCompletionData_ called with invalid sheet object.');
        return null;
    }
    Logger.log(`Calculating downtime completion data for sheet: ${sheet.getName()}`);

    try {
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const lastRow = sheet.getLastRow();
        // Basic return structure for empty/invalid sheet
        const defaultResults = {
            characterCount: 0, completedDowntimeCells: 0, totalDowntimeCells: 0,
            keywordCounts: {}, keywordCompletedCounts: {},
            averageWords: 0, medianWords: 0, minWords: 0, maxWords: 0,
            minWordCell: "", maxWordCell: "",
            averageResponseWords: 0, medianResponseWords: 0, minResponseWords: 0, maxResponseWords: 0,
            minResponseWordCell: "", maxResponseWordCell: ""
        };
        Object.keys(DOWNTIME_KEYWORDS).forEach(key => {
            defaultResults.keywordCounts[key] = 0;
            defaultResults.keywordCompletedCounts[key] = 0;
        });

        if (lastRow <= 1) return defaultResults;

        let characterCount = 0;
        let completedDowntimeCells = 0;
        let totalDowntimeCells = 0;
        const keywordCounts = Object.keys(DOWNTIME_KEYWORDS).reduce((acc, key) => { acc[key] = 0; return acc; }, {});
        const keywordCompletedCounts = Object.keys(DOWNTIME_KEYWORDS).reduce((acc, key) => { acc[key] = 0; return acc; }, {});
        // Optional: Track uncategorized
        // keywordCounts['uncategorized'] = 0;
        // keywordCompletedCounts['uncategorized'] = 0;


        const submissionWordCounts = [];
        const responseWordCounts = [];
        let minWords = Infinity, maxWords = 0, minWordCell = "", maxWordCell = "";
        let minResponseWords = Infinity, maxResponseWords = 0, minResponseWordCell = "", maxResponseWordCell = "";

        const data = sheet.getDataRange().getValues(); // Get all data at once

        // Iterate through data rows (skip header)
        // Row pairs: Submission (even index in data array, odd row number), Response (odd index in data array, even row number)
        for (let i = 1; i < data.length; i += 2) { // Step by 2
            const submissionRow = data[i]; // Even index = Submission Row
            // Check if it's a valid submission row (has timestamp)
            if (submissionRow[TIMESTAMP_COL - 1]) {
                characterCount++;
                // Check if there's a corresponding response row
                if (i + 1 < data.length) {
                    const responseRow = data[i + 1]; // Odd index = Response Row
                    const characterName = responseRow[CHARACTER_NAME_COL - 1] || 'Unknown'; // Get character name from RESPONSE row

                    // Iterate through columns identified as downtime columns
                    headers.forEach((header, index) => {
                        if (DOWNTIME_HEADER_REGEX.test(header)) {
                            const colIndex = index; // 0-based index for data array
                            const submissionValue = submissionRow[colIndex] ? String(submissionRow[colIndex]).trim() : "";
                            const responseValue = responseRow[colIndex] ? String(responseRow[colIndex]).trim() : "";
                            const currentSubmissionCell = getColumnLetter(colIndex + 1) + (i + 1 + 1); // +1 for 1-based sheet row, +1 because i is 0-based data index
                            const currentResponseCell = getColumnLetter(colIndex + 1) + (i + 2 + 1); // +1 for 1-based sheet row, +2 because i is 0-based data index

                            if (submissionValue !== "") {
                                totalDowntimeCells++;

                                // Submission Word Count Stats
                                const submissionWordCount = submissionValue.split(/\s+/).filter(Boolean).length;
                                submissionWordCounts.push(submissionWordCount);
                                if (submissionWordCount < minWords) { minWords = submissionWordCount; minWordCell = currentSubmissionCell; }
                                if (submissionWordCount > maxWords) { maxWords = submissionWordCount; maxWordCell = currentSubmissionCell; }

                                // Keyword Matching
                                const lowerSubmissionValue = submissionValue.toLowerCase();
                                let keywordFound = false;
                                for (const keywordCategory in DOWNTIME_KEYWORDS) {
                                    if (DOWNTIME_KEYWORDS[keywordCategory].some(term => lowerSubmissionValue.includes(term))) {
                                        keywordCounts[keywordCategory]++;
                                        if (responseValue !== "") {
                                            keywordCompletedCounts[keywordCategory]++;
                                        }
                                        keywordFound = true;
                                        // break; // Count only the first match? Or allow multiple? Current: Allows multiple.
                                    }
                                }
                                // Optional: Track uncategorized
                                // if (!keywordFound) {
                                //     keywordCounts['uncategorized']++;
                                //     if (responseValue !== "") keywordCompletedCounts['uncategorized']++;
                                // }


                                // Response Processing
                                if (responseValue !== "") {
                                    completedDowntimeCells++;

                                    // Response Word Count Stats
                                    const responseWordCount = responseValue.split(/\s+/).filter(Boolean).length;
                                    responseWordCounts.push(responseWordCount);
                                    if (responseWordCount < minResponseWords) { minResponseWords = responseWordCount; minResponseWordCell = currentResponseCell; }
                                    if (responseWordCount > maxResponseWords) { maxResponseWords = responseWordCount; maxResponseWordCell = currentResponseCell; }
                                }
                            }
                        }
                    });
                } else {
                    Logger.log(`Warning: Submission row ${i + 1 + 1} has no corresponding response row.`); // Adjust row number for logging
                }
                // No i++ here as the main loop steps by 2
            }
        }

        const submissionStats = calculateStats_(submissionWordCounts);
        const responseStats = calculateStats_(responseWordCounts);

        // Handle cases where no words were found
        minWords = minWords === Infinity ? 0 : minWords;
        minResponseWords = minResponseWords === Infinity ? 0 : minResponseWords;

        const results = {
            characterCount: characterCount,
            completedDowntimeCells: completedDowntimeCells,
            totalDowntimeCells: totalDowntimeCells,
            keywordCounts: keywordCounts,
            keywordCompletedCounts: keywordCompletedCounts,
            averageWords: submissionStats.average,
            medianWords: submissionStats.median,
            minWords: minWords,
            maxWords: maxWords,
            minWordCell: minWordCell,
            maxWordCell: maxWordCell,
            averageResponseWords: responseStats.average,
            medianResponseWords: responseStats.median,
            minResponseWords: minResponseWords,
            maxResponseWords: maxResponseWords,
            minResponseWordCell: minResponseWordCell,
            maxResponseWordCell: maxResponseWordCell,
        };
        Logger.log(`Downtime completion data calculated. Total Cells: ${results.totalDowntimeCells}, Completed: ${results.completedDowntimeCells}`);
        return results;

    } catch (error) {
        Logger.log(`Error calculating downtime completion data: ${error}\nStack: ${error.stack}`);
        return null;
    }
}


/**
 * Gets data for missing downtime responses.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to analyze.
 * @returns {Array<{characterName: string, header: string, text: string, cell: string}>|null} Array of missing items or null on error.
 */
function getMissingDowntimeData_(sheet) {
    return getMissingItemsData_(sheet, DOWNTIME_HEADER_REGEX, 'Downtime');
}
/**
 * Gets data for missing influence responses.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to analyze.
 * @returns {Array<{characterName: string, header: string, text: string, cell: string}>|null} Array of missing items or null on error.
 */
function getMissingInfluencesData_(sheet) {
    return getMissingItemsData_(sheet, INFLUENCE_HEADER_REGEX, 'Influence');
}

/**
 * Gets data for missing resource responses.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to analyze.
 * @returns {Array<{characterName: string, header: string, text: string, cell: string}>|null} Array of missing items or null on error.
 */
function getMissingResourcesData_(sheet) {
    return getMissingItemsData_(sheet, RESOURCES_HEADER_REGEX, 'Resource');
}

/**
 * Generic function to get data for missing responses based on a header regex.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to analyze.
 * @param {RegExp} headerRegex Regex to identify relevant columns.
 * @param {string} itemType Used for logging (e.g., 'Downtime', 'Influence').
 * @returns {Array<{characterName: string, header: string, text: string, cell: string}>|null} Array of missing items or null on error.
 */
function getMissingItemsData_(sheet, headerRegex, itemType) {
    if (!sheet || !headerRegex) {
        Logger.log(`getMissingItemsData_ called with invalid sheet or regex for type ${itemType}.`);
        return null;
    }
    Logger.log(`Getting missing ${itemType} data for sheet: ${sheet.getName()}`);

    try {
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return []; // No data rows

        const missingItems = [];
        const data = sheet.getDataRange().getValues(); // Get all data at once

        // Iterate through data rows (skip header)
        // Row pairs: Submission (even index in data array, odd row number), Response (odd index in data array, even row number)
        for (let i = 1; i < data.length; i += 2) { // Step by 2
            const submissionRow = data[i]; // Even index = Submission Row
            // Check if it's a valid submission row (has timestamp)
            if (submissionRow[TIMESTAMP_COL - 1]) {
                // Check if there's a corresponding response row
                if (i + 1 < data.length) {
                    const responseRow = data[i + 1]; // Odd index = Response Row
                    const characterName = responseRow[CHARACTER_NAME_COL - 1] || 'Unknown'; // Get character name from RESPONSE row

                    // Iterate through columns matching the regex
                    headers.forEach((header, index) => {
                        if (headerRegex.test(header)) {
                            const colIndex = index; // 0-based index for data array
                            const submissionValue = submissionRow[colIndex] ? String(submissionRow[colIndex]).trim() : "";
                            const responseValue = responseRow[colIndex] ? String(responseRow[colIndex]).trim() : "";
                            const responseCellA1 = getColumnLetter(colIndex + 1) + (i + 1 + 1); // +1 for 1-based sheet row, +1 because i is 0-based data index

                            // Check if submission exists but response is empty
                            if (submissionValue !== "" && responseValue === "") {
                                const truncatedText = submissionValue.length > 100 ? submissionValue.substring(0, 97) + "..." : submissionValue;
                                missingItems.push({
                                    characterName: characterName,
                                    header: header,
                                    text: truncatedText, // Show the submitted text (truncated)
                                    cell: responseCellA1 // The cell needing the response
                                });
                            }
                        }
                    });
                } else {
                     Logger.log(`Warning: Submission row ${i + 1 + 1} has no corresponding response row while checking for missing ${itemType}.`); // Adjust row number
                }
                // No i++ needed here as main loop steps by 2
            }
        }
        Logger.log(`Found ${missingItems.length} missing ${itemType} responses.`);
        return missingItems;
    } catch (error) {
        Logger.log(`Error getting missing ${itemType} data: ${error}\nStack: ${error.stack}`);
        SpreadsheetApp.getUi().alert(`Error calculating missing ${itemType} items: ${error.message}`);
        return null;
    }
}

/**
 * Gets influence completion statistics.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to analyze.
 * @returns {object|null} An object with statistics, or null if data is invalid.
 */
function getInfluencesCompletionData_(sheet) {
    if (!sheet) {
        Logger.log('getInfluencesCompletionData_ called with invalid sheet object.');
        return null;
    }
    Logger.log(`Calculating influence completion data for sheet: ${sheet.getName()}`);
    return getCategoryCompletionData_(sheet, INFLUENCE_HEADER_REGEX, ['elite', 'underworld'], 'Influence');
}

/**
 * Gets resource completion statistics.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to analyze.
 * @returns {object|null} An object with statistics, or null if data is invalid.
 */
function getResourcesCompletionData_(sheet) {
    if (!sheet) {
        Logger.log('getResourcesCompletionData_ called with invalid sheet object.');
        return null;
    }
    Logger.log(`Calculating resource completion data for sheet: ${sheet.getName()}`);
    // Assuming only one category "resources" for this type
    return getCategoryCompletionData_(sheet, RESOURCES_HEADER_REGEX, ['resources'], 'Resource');
}

/**
 * Generic function to calculate completion data for categorized items (Influences, Resources).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to analyze.
 * @param {RegExp} headerRegex Regex to identify relevant columns.
 * @param {string[]} categories Array of category names (lowercase) to look for in headers (e.g., ['elite', 'underworld']).
 * @param {string} itemType Used for logging (e.g., 'Influence', 'Resource').
 * @returns {object|null} An object containing counts per category and a list of missing items, or null on error.
 */
function getCategoryCompletionData_(sheet, headerRegex, categories, itemType) {
    Logger.log(`Calculating ${itemType} completion data for sheet: ${sheet.getName()}, Categories: [${categories.join(', ')}]`);
    // Default structure
     const defaultResults = {
        characterCount: 0,
        missingItems: [],
        totalCounts: categories.reduce((acc, cat) => { acc[cat] = 0; return acc; }, {}),
        completedCounts: categories.reduce((acc, cat) => { acc[cat] = 0; return acc; }, {})
    };

    try {
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return defaultResults;

        const results = {
            characterCount: 0,
            missingItems: [],
            totalCounts: categories.reduce((acc, cat) => { acc[cat] = 0; return acc; }, {}),
            completedCounts: categories.reduce((acc, cat) => { acc[cat] = 0; return acc; }, {})
        };

        const data = sheet.getDataRange().getValues(); // Get all data at once

        // Iterate through data rows (skip header)
        // Row pairs: Submission (even index in data array, odd row number), Response (odd index in data array, even row number)
        for (let i = 1; i < data.length; i += 2) { // Step by 2
            const submissionRow = data[i]; // Even index = Submission Row
            if (submissionRow[TIMESTAMP_COL - 1]) { // Is a submission row
                results.characterCount++;
                if (i + 1 < data.length) {
                    const responseRow = data[i + 1]; // Odd index = Response Row
                    const characterName = responseRow[CHARACTER_NAME_COL - 1] || 'Unknown'; // Name from response row

                    headers.forEach((header, index) => {
                        if (headerRegex.test(header)) {
                            const colIndex = index;
                            const submissionValue = submissionRow[colIndex] ? String(submissionRow[colIndex]).trim() : "";
                            const responseValue = responseRow[colIndex] ? String(responseRow[colIndex]).trim() : "";
                            const responseCellA1 = getColumnLetter(colIndex + 1) + (i + 1 + 1); // Response row number
                            const lowerHeader = header.toLowerCase();

                            if (submissionValue !== "") {
                                // Find which category this header belongs to
                                const matchedCategory = categories.find(cat => lowerHeader.includes(cat));

                                if (matchedCategory) {
                                    results.totalCounts[matchedCategory]++;
                                    if (responseValue !== "") {
                                        results.completedCounts[matchedCategory]++;
                                    } else {
                                        // Add to missing items list
                                        const truncatedText = submissionValue.length > 100 ? submissionValue.substring(0, 97) + "..." : submissionValue;
                                        results.missingItems.push({
                                            characterName: characterName,
                                            header: header,
                                            text: truncatedText,
                                            cell: responseCellA1,
                                            category: matchedCategory // Store category if needed later
                                        });
                                    }
                                } else {
                                     Logger.log(`Header "${header}" matched regex but not any specified category for ${itemType}.`);
                                }
                            }
                        }
                    });
                } else {
                     Logger.log(`Warning: Submission row ${i + 1 + 1} has no corresponding response row while checking ${itemType}.`); // Adjust row number
                }
                // No i++ needed here as main loop steps by 2
            }
        }

        Logger.log(`${itemType} completion data calculated. Total Counts: ${JSON.stringify(results.totalCounts)}, Completed Counts: ${JSON.stringify(results.completedCounts)}, Missing: ${results.missingItems.length}`);
        return results;

    } catch (error) {
        Logger.log(`Error calculating ${itemType} completion data: ${error}\nStack: ${error.stack}`);
        SpreadsheetApp.getUi().alert(`Error calculating ${itemType} completion: ${error.message}`);
        return null; // Return null on error
    }
}
