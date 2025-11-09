/**
 * Functions callable from the web app for actions.
 */

/**
 * Generates data for a given type ('herd', 'feed', 'discipline', 'patrol') and returns it.
 * This is a web-app-callable version of the logic in Actions.js.
 * @param {string} type - The type of data to generate.
 * @returns {string} The generated text.
 */
function getFillData(type) {
  try {
    const fillReplaceData = getFillReplaceListData_(); // In SheetData.gs
    if (!fillReplaceData) {
      throw new Error('Could not retrieve fill/replace list data.');
    }

    const { headers, values } = fillReplaceData;
    let columnIndex;
    let headerName;

    const lowerCaseHeaders = headers.map(h => String(h).toLowerCase());
    if (type === 'herd') {
      columnIndex = lowerCaseHeaders.indexOf('herd');
      headerName = 'Herd';
    } else if (type === 'feed') {
      columnIndex = lowerCaseHeaders.indexOf('feed');
      headerName = 'Feed';
    } else if (type === 'discipline') {
      columnIndex = lowerCaseHeaders.indexOf('discipline');
      headerName = 'Discipline';
    } else if (type === 'patrol') {
      columnIndex = lowerCaseHeaders.indexOf('patrol');
      headerName = 'Patrol';
    } else {
      throw new Error('Invalid data type specified.');
    }

    if (columnIndex === -1) {
      throw new Error(`Column header "${headerName}" not found in '${FILL_REPLACE_SHEET_NAME}'.`);
    }

    const columnValues = values.slice(1)
                           .map(row => row[columnIndex])
                           .filter(value => value && String(value).trim() !== '');

    if (columnValues.length === 0) {
      throw new Error(`No data found in the "${headerName}" column of '${FILL_REPLACE_SHEET_NAME}'.`);
    }

    let textToInsert = columnValues[Math.floor(Math.random() * columnValues.length)];

    for (let i = 0; i < headers.length; i++) {
      if (i === columnIndex) continue;
      const header = headers[i];
      if (!header || String(header).trim() === '') continue;

      const placeholder = `[${header}]`;
      if (String(textToInsert).toLowerCase().includes(placeholder.toLowerCase())) {
        const placeholderColValues = values.slice(1)
                                      .map(row => row[i])
                                      .filter(value => value && String(value).trim() !== '');
        if (placeholderColValues.length > 0) {
          const replacement = placeholderColValues[Math.floor(Math.random() * placeholderColValues.length)];
          const regex = new RegExp(`\\[${escapeRegex_(header)}\\]`, 'gi'); // In Utilities.gs
          textToInsert = String(textToInsert).replace(regex, replacement);
        }
      }
    }

    return textToInsert;

  } catch (error) {
    Logger.log(`Error in getFillData for type ${type}: ${error.stack}`);
    // Re-throw the error so the client-side failure handler catches it.
    throw new Error(`Error generating ${type} data: ${error.message}`);
  }
}
