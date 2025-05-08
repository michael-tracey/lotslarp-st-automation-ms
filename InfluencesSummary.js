/**
 * @OnlyCurrentDoc
 *
 * Functions to generate and display a detailed summary of Influence actions
 * grouped by month, level, type, action, and specialization.
 * Also includes function to list differing action names per level and top/bottom/unused spec usage.
 */

// ALL MOVED TO CONSTANTS.GS
/*
const ELITE_INFL_SHEET_NAME_ = 'Elite Infl.';
const UW_INFL_SHEET_NAME_ = 'UW Infl.';

const INFL_DATE_COL_IDX_ = 0; // Column A
const INFL_SPEC_COL_IDX_ = 3; // Column D
const INFL_ACTION_COL_IDX_ = 4; // Column E
const TOP_N_SPECS = 3;
const ALL_SPECIALIZATIONS = [
    "Academics", "Activist Organizations", "Arts", "Commercial Industry",
    "Criminal Organizations", "Emergency Services", "Entertainment", "Finance",
    "High Society", "Infrastructure", "Manufacturing", "Law Enforcement",
    "Legal", "Local Government", "Science/Medical", "Social Media",
    "News Media", "Occult", "Religious Groups/Organizations", "Transportation"
].sort(); // Sort master list alphabetically
*/

/**
 * Formats a Date object into "MonthName Year" string.
 * Handles invalid dates gracefully.
 * @param {Date | string | number} dateInput The date to format (Date object, timestamp number, or parsable string).
 * @returns {string} Formatted date string or 'Invalid Date'.
 */
function formatDateAsMonthYear_(dateInput) {
    let dateObject;
    if (dateInput instanceof Date) {
        dateObject = dateInput;
    } else if (typeof dateInput === 'number' || typeof dateInput === 'string') {
        // Attempt to parse common string formats or timestamps
        dateObject = new Date(dateInput);
    } else {
        Logger.log(`Invalid date input type: ${typeof dateInput}`);
        return 'Invalid Date Input';
    }

    // Check if the date is valid after parsing/creation
    if (isNaN(dateObject.getTime())) {
         Logger.log(`Could not parse date input: ${dateInput}`);
        return 'Invalid Date';
    }
    const options = { year: 'numeric', month: 'long' };
    return dateObject.toLocaleDateString('en-US', options); // Adjust locale if needed
}


/**
 * Parses an action string (e.g., "3: Action Name") into level and name.
 * @param {string} actionString The action string.
 * @returns {{level: number, name: string}} Parsed level and name. Returns level Infinity if no number found.
 */
function parseActionString_(actionString) {
    if (!actionString) return { level: Infinity, name: 'Unknown Action' };
    const trimmedAction = String(actionString).trim();
    const match = trimmedAction.match(/^(\d+)\s*:\s*(.*)/);
    if (match && match[1] && match[2]) {
        // Found "Number: Name" format
        return { level: parseInt(match[1], 10), name: match[2].trim() };
    } else if (trimmedAction) {
         // If no "Number:", treat level as Infinity and use the whole string as name
         return { level: Infinity, name: trimmedAction };
    } else {
        // If empty after trimming
        return { level: Infinity, name: 'Unknown Action' };
    }
}


/**
 * Gathers and counts influence actions, grouping by Month -> Level -> Type -> Action -> Spec.
 * Also calculates totals per month per type.
 * @returns {object|null} A nested object containing the counts and totals, or null if sheets are missing.
 * Structure: { "Month Year": { totalElite: X, totalUW: Y, levels: { levelNum: { Elite: { "ActionName": { specs: {"SpecName": count} } }, UW: { ... } } } }, ... }
 */
function getInfluenceActionDetails_() {
    Logger.log("Starting detailed influence action data gathering (Level-first)...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summaryData = {}; // { "Month Year": { totalElite: X, totalUW: Y, levels: {} } }
    const ui = SpreadsheetApp.getUi();
    let eliteSheetFound = false;
    let uwSheetFound = false;

    const processSheet = (sheetName, type) => {
        let sheetFound = false;
        try {
            const sheet = ss.getSheetByName(sheetName);
            if (sheet) {
                sheetFound = true;
                const rawData = sheet.getDataRange().getValues();
                Logger.log(`Processing ${rawData.length - 1} rows from ${sheetName} as ${type}.`);

                for (let i = 1; i < rawData.length; i++) { // Skip header
                    const row = rawData[i];
                    const actionString = String(row[INFL_ACTION_COL_IDX_] || '').trim();
                    if (!actionString) continue; // Skip if no action specified

                    const dateVal = row[INFL_DATE_COL_IDX_];
                    const monthYear = formatDateAsMonthYear_(dateVal);
                    const specialization = String(row[INFL_SPEC_COL_IDX_] || 'Unknown Spec').trim();
                    const parsedAction = parseActionString_(actionString); // {level, name}

                    if (monthYear === 'Invalid Date' || monthYear === 'Invalid Date Input' || parsedAction.name === 'Unknown Action') {
                        Logger.log(`Skipping row ${i+1} in ${sheetName}: Invalid date or action format (Date: ${dateVal}, Action: ${actionString}).`);
                        continue;
                    }

                    // Initialize structures
                    summaryData[monthYear] = summaryData[monthYear] || { totalElite: 0, totalUW: 0, levels: {} };
                    summaryData[monthYear].levels[parsedAction.level] = summaryData[monthYear].levels[parsedAction.level] || { Elite: {}, Underworld: {} };
                    summaryData[monthYear].levels[parsedAction.level][type] = summaryData[monthYear].levels[parsedAction.level][type] || {};
                    summaryData[monthYear].levels[parsedAction.level][type][parsedAction.name] = summaryData[monthYear].levels[parsedAction.level][type][parsedAction.name] || { specs: {} };

                    // Increment counts
                    if (type === 'Elite') {
                        summaryData[monthYear].totalElite++;
                    } else {
                        summaryData[monthYear].totalUW++;
                    }
                    summaryData[monthYear].levels[parsedAction.level][type][parsedAction.name].specs[specialization] = (summaryData[monthYear].levels[parsedAction.level][type][parsedAction.name].specs[specialization] || 0) + 1;
                }
            } else {
                Logger.log(`Sheet "${sheetName}" not found.`);
            }
        } catch (error) {
             Logger.log(`Error processing sheet "${sheetName}": ${error}\n${error.stack}`);
             ui.alert(`Error reading ${sheetName}: ${error.message}`);
        }
        return sheetFound;
    };

    // Process both sheets
    eliteSheetFound = processSheet(ELITE_INFL_SHEET_NAME_, 'Elite');
    uwSheetFound = processSheet(UW_INFL_SHEET_NAME_, 'Underworld');

    if (!eliteSheetFound && !uwSheetFound) {
        ui.alert(`Could not find sheets "${ELITE_INFL_SHEET_NAME_}" or "${UW_INFL_SHEET_NAME_}" to generate summary.`);
        return null;
    }

    Logger.log(`Finished influence data gathering. Summary keys: ${Object.keys(summaryData)}`);
    return summaryData;
}


/**
 * Helper to calculate and format percentage change.
 * @param {number} current The current value.
 * @param {number} previous The previous value.
 * @returns {string} Formatted percentage change string (e.g., "+10.5%", "-5.0%", "N/A", "+Inf%").
 */
function formatPercentageChange_(current, previous) {
    if (previous === 0) {
        return (current > 0) ? '+Inf%' : 'N/A'; // Or just 'N/A' if preferred for 0 to >0 change
    }
    const change = ((current - previous) / previous) * 100;
    const sign = change >= 0 ? '+' : '';
    return `${sign}${change.toFixed(1)}%`;
}


/**
 * Calculates percentages, top/bottom/unused specs, and restructures summary data for the template.
 * @param {object} summaryData The raw summary data from getInfluenceActionDetails_.
 * @returns {{monthTotals: Array<object>, details: object}} The processed data structure for the HTML template.
 * Structure: { monthTotals: [...], details: { "Month Year": { levels: [...], mostUsed: [...], leastUsed: [...], unusedElite: [], unusedUw: [] } } }
 */
function processInfluenceSummaryForTemplate_(summaryData) {
    const monthTotals = [];
    const details = {};

    // Helper function to extract number prefix for sorting
    const extractActionNumber_ = (actionName) => {
        const match = actionName.match(/^(\d+):/);
        return match ? parseInt(match[1], 10) : Infinity; // Sort actions without numbers last
    };

    // Get sorted list of months (keys) based on date parsing
    const sortedMonths = Object.keys(summaryData).sort((a, b) => new Date(a) - new Date(b));

    let prevEliteTotal = null;
    let prevUwTotal = null;

    sortedMonths.forEach(monthYear => {
        const monthData = summaryData[monthYear];
        const totalElite = monthData.totalElite || 0;
        const totalUW = monthData.totalUW || 0;

        // Calculate % change from previous month
        const eliteChangePercent = (prevEliteTotal !== null) ? formatPercentageChange_(totalElite, prevEliteTotal) : 'N/A';
        const uwChangePercent = (prevUwTotal !== null) ? formatPercentageChange_(totalUW, prevUwTotal) : 'N/A';

        // Add to month totals summary
        monthTotals.push({
             monthYear: monthYear,
             eliteTotal: totalElite,
             uwTotal: totalUW,
             eliteChangePercent: eliteChangePercent, // Add change %
             uwChangePercent: uwChangePercent // Add change %
        });
        // Update previous totals for next iteration
        prevEliteTotal = totalElite;
        prevUwTotal = totalUW;


        const levelsForMonth = [];
        const specTotalsForMonth = {}; // { "SpecName": count } - Aggregate counts here
        const usedEliteSpecs = new Set(); // Track specs used in Elite
        const usedUwSpecs = new Set(); // Track specs used in UW

        // Get level numbers, convert to number, sort numerically
        const sortedLevels = Object.keys(monthData.levels).map(Number).sort((a, b) => a - b);

        sortedLevels.forEach(level => {
            const levelData = monthData.levels[level];
            let eliteActionDetails = null;
            let uwActionDetails = null;

            // Process Elite actions for this level
            if (levelData.Elite && Object.keys(levelData.Elite).length > 0) {
                 const actionName = Object.keys(levelData.Elite)[0];
                 const specData = levelData.Elite[actionName].specs;
                 let eliteActionCount = 0;
                 const specs = Object.entries(specData).map(([specName, count]) => {
                    specTotalsForMonth[specName] = (specTotalsForMonth[specName] || 0) + count; // Aggregate total spec count
                    usedEliteSpecs.add(specName); // Track used spec
                    eliteActionCount += count;
                    return { name: specName, count: count };
                 }).sort((a, b) => a.name.localeCompare(b.name));

                 const percentage = totalElite > 0 ? (eliteActionCount / totalElite * 100).toFixed(1) : '0.0';
                 eliteActionDetails = { name: actionName, specs: specs, percentage: percentage };
            }

             // Process Underworld actions for this level
            if (levelData.Underworld && Object.keys(levelData.Underworld).length > 0) {
                 const actionName = Object.keys(levelData.Underworld)[0];
                 const specData = levelData.Underworld[actionName].specs;
                 let uwActionCount = 0;
                 const specs = Object.entries(specData).map(([specName, count]) => {
                     specTotalsForMonth[specName] = (specTotalsForMonth[specName] || 0) + count; // Aggregate total spec count
                     usedUwSpecs.add(specName); // Track used spec
                     uwActionCount += count;
                     return { name: specName, count: count };
                 }).sort((a, b) => a.name.localeCompare(b.name));

                 const percentage = totalUW > 0 ? (uwActionCount / totalUW * 100).toFixed(1) : '0.0';
                 uwActionDetails = { name: actionName, specs: specs, percentage: percentage };
            }

            // Only add level if at least one type has data
            if (eliteActionDetails || uwActionDetails) {
                levelsForMonth.push({
                    level: level,
                    eliteAction: eliteActionDetails,
                    uwAction: uwActionDetails
                });
            }
        });

        // Calculate Top/Bottom/Unused Specializations for the month using the full list
        const specTotalsArray = ALL_SPECIALIZATIONS.map(specName => ({
            name: specName,
            totalCount: specTotalsForMonth[specName] || 0 // Default to 0 if not found
        }));

        const mostUsedSpecs = [...specTotalsArray]
                                .sort((a, b) => b.totalCount - a.totalCount) // Sort descending
                                .slice(0, TOP_N_SPECS);

        // For least used, filter out those with 0 count before sorting and slicing
        const leastUsedSpecs = specTotalsArray
                                .filter(spec => spec.totalCount > 0) // Only consider specs used at least once
                                .sort((a, b) => a.totalCount - b.totalCount) // Sort ascending
                                .slice(0, TOP_N_SPECS);

        // For unused, filter based on the sets tracked earlier
        const unusedElite = ALL_SPECIALIZATIONS.filter(specName => !usedEliteSpecs.has(specName));
        const unusedUw = ALL_SPECIALIZATIONS.filter(specName => !usedUwSpecs.has(specName));


        details[monthYear] = {
            levels: levelsForMonth,
            mostUsed: mostUsedSpecs,
            leastUsed: leastUsedSpecs,
            unusedElite: unusedElite, // Add separate unused list
            unusedUw: unusedUw       // Add separate unused list
         };
    });

    return { monthTotals, details };
}


/**
 * Displays the detailed influence action summary dialog.
 * This is intended to be called from the main script's menu.
 */
function showDetailedInfluenceSummaryDialog_() {
    Logger.log('Showing Detailed Influence Summary Dialog...');
    const ui = SpreadsheetApp.getUi();
    try {
        const rawSummaryData = getInfluenceActionDetails_();
        if (!rawSummaryData || Object.keys(rawSummaryData).length === 0) {
            Logger.log('No influence data found to display.');
             if (Object.keys(rawSummaryData || {}).length === 0 && rawSummaryData !== null) {
                 ui.alert('No influence actions found in the "' + ELITE_INFL_SHEET_NAME_ + '" or "' + UW_INFL_SHEET_NAME_ + '" sheets to summarize.');
            }
            return;
        }

        // Process data for template (calculate percentages, restructure/sort, get top/bottom/unused)
        const processedData = processInfluenceSummaryForTemplate_(rawSummaryData);

        const template = HtmlService.createTemplateFromFile('DetailedInfluenceSummaryDialog');
        template.summary = processedData; // Pass processed data {monthTotals, details}
        template.TOP_N_SPECS = TOP_N_SPECS; // Pass constant for display

        const htmlOutput = template.evaluate()
            .setWidth(950) // Keep wider width
            .setHeight(600); // Increase height slightly for new tables

        ui.showModalDialog(htmlOutput, 'Detailed Influence Action Summary (by Level)'); // Update title
        Logger.log('Detailed Influence Summary Dialog displayed.');
    } catch (error) {
        Logger.log(`Error showing detailed influence summary dialog: ${error}\nStack: ${error.stack}`);
        ui.alert(`Error showing influence summary: ${error.message}`);
    }
}

/**
 * Finds and lists influence actions where the name differs between Elite and UW sheets for the same level.
 */
function listDifferingInfluenceActions_() {
    Logger.log("Finding differing influence action names per level...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const eliteActionsByLevel = {}; // { level: Set<string> }
    const uwActionsByLevel = {};    // { level: Set<string> }
    let eliteSheetFound = false;
    let uwSheetFound = false;

    // --- Process Elite Sheet ---
    try {
        const eliteSheet = ss.getSheetByName(ELITE_INFL_SHEET_NAME_);
        if (eliteSheet) {
            eliteSheetFound = true;
            const eliteData = eliteSheet.getDataRange().getValues();
            Logger.log(`Processing ${eliteData.length - 1} rows from ${ELITE_INFL_SHEET_NAME_} for differences.`);
            for (let i = 1; i < eliteData.length; i++) {
                const actionString = String(eliteData[i][INFL_ACTION_COL_IDX_] || '').trim();
                if (!actionString) continue;
                const parsed = parseActionString_(actionString);
                if (!eliteActionsByLevel[parsed.level]) {
                    eliteActionsByLevel[parsed.level] = new Set();
                }
                eliteActionsByLevel[parsed.level].add(parsed.name);
            }
        } else {
             Logger.log(`Sheet "${ELITE_INFL_SHEET_NAME_}" not found for difference check.`);
        }
    } catch(e) { Logger.log(`Error reading ${ELITE_INFL_SHEET_NAME_}: ${e}`); }

     // --- Process Underworld Sheet ---
     try {
        const uwSheet = ss.getSheetByName(UW_INFL_SHEET_NAME_);
        if (uwSheet) {
            uwSheetFound = true;
            const uwData = uwSheet.getDataRange().getValues();
             Logger.log(`Processing ${uwData.length - 1} rows from ${UW_INFL_SHEET_NAME_} for differences.`);
            for (let i = 1; i < uwData.length; i++) {
                const actionString = String(uwData[i][INFL_ACTION_COL_IDX_] || '').trim(); // Use same index
                 if (!actionString) continue;
                 const parsed = parseActionString_(actionString);
                 if (!uwActionsByLevel[parsed.level]) {
                     uwActionsByLevel[parsed.level] = new Set();
                 }
                 uwActionsByLevel[parsed.level].add(parsed.name);
            }
        } else {
             Logger.log(`Sheet "${UW_INFLUENCES_SHEET_NAME_}" not found for difference check.`);
        }
    } catch(e) { Logger.log(`Error reading ${UW_INFLUENCES_SHEET_NAME_}: ${e}`); }

     if (!eliteSheetFound && !uwSheetFound) {
        ui.alert(`Could not find sheets "${ELITE_INFLUENCES_SHEET_NAME_}" or "${UW_INFLUENCES_SHEET_NAME_}" to check differences.`);
        return;
    }

    // --- Compare levels ---
    const differingLevels = [];
    const allLevels = new Set([...Object.keys(eliteActionsByLevel), ...Object.keys(uwActionsByLevel)]);

    allLevels.forEach(levelStr => {
        // Skip potential 'Infinity' key from parsing errors if it got added directly
        if (levelStr === 'Infinity') return;
        const level = parseInt(levelStr, 10);
        if (isNaN(level)) return;

        const eliteNames = eliteActionsByLevel[level] || new Set();
        const uwNames = uwActionsByLevel[level] || new Set();

        // Convert sets to sorted arrays for reliable comparison
        const eliteArr = [...eliteNames].sort();
        const uwArr = [...uwNames].sort();

        // Check if the stringified sorted arrays are different
        if (JSON.stringify(eliteArr) !== JSON.stringify(uwArr)) {
            differingLevels.push({
                level: level,
                elite: eliteArr.length > 0 ? eliteArr.join(', ') : 'N/A',
                uw: uwArr.length > 0 ? uwArr.join(', ') : 'N/A'
            });
        }
    });

    // --- Display Results ---
    let message = "Influence Action Name Differences by Level:\n\n";
    if (differingLevels.length === 0) {
        message = "No differences found in action names for the same level between Elite and Underworld sheets.";
    } else {
        // Sort results by level numerically
        differingLevels.sort((a, b) => a.level - b.level);
        differingLevels.forEach(diff => {
            message += `Level ${diff.level}:\n  Elite: ${diff.elite}\n  Underworld: ${diff.uw}\n\n`;
        });
    }

    ui.alert(message);
    logAudit_('Checked Influence Name Differences', 'N/A', `Found ${differingLevels.length} levels with differences.`); // In Utilities.gs

}
