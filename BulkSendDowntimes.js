/**
 * @OnlyCurrentDoc
 */

let isPaused = false;
let isStopped = false;
const CACHE_KEY = 'bulkSendStatus';
const SUMMARY_CACHE_KEY = 'bulkSendSummary';

function showBulkSendDowntimesDialog() {
  const html = HtmlService.createHtmlOutputFromFile('BulkSendDowntimesDialog')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Bulk Send Downtimes');
}

function getAllDowntimes() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();

    if (!data || data.length < 2) {
      throw new Error('The sheet is empty or has no data rows. Please make sure the sheet has a header row and at least one row of data.');
    }

    const headers = data[0];
    if (!headers || headers.length === 0) {
      throw new Error('The header row is empty. Please make sure the first row of the sheet contains the column headers.');
    }

    const characterNameIndex = headers.indexOf('What is your character\'s name?');
    const discordSentIndex = headers.indexOf('Send Discord');

    if (characterNameIndex === -1 || discordSentIndex === -1) {
      throw new Error('Could not find required columns. Please check the sheet headers.');
    }

    const allDowntimes = [];
    for (let i = 1; i < data.length; i++) {
      // Only process response rows (even indices in the data array, which correspond to odd rows in the sheet)
      if (i > 0 && i % 2 === 0) {
        const responseRow = data[i];
        const submissionRow = data[i - 1];
        let responseCount = 0;
        let stWordCount = 0;
        let playerWordCount = 0;

        for (let j = characterNameIndex + 1; j < headers.length; j++) {
          if (responseRow[j] && responseRow[j].toString().trim() !== '') {
            responseCount++;
            stWordCount += responseRow[j].toString().split(' ').length;
          }
        }

        for (let j = 0; j < headers.length; j++) {
          if (headers[j].toLowerCase().includes('downtime')) {
            if (submissionRow[j] && submissionRow[j].toString().trim() !== '') {
              playerWordCount += submissionRow[j].toString().split(' ').length;
            }
          }
        }
        
        const isSent = responseRow[discordSentIndex] !== '' && responseRow[discordSentIndex] !== false;

        allDowntimes.push({
          rowIndex: i + 1,
          characterName: responseRow[characterNameIndex],
          responseCount: responseCount,
          stWordCount: stWordCount,
          playerWordCount: playerWordCount,
          isSent: isSent,
          status: isSent ? 'sent' : 'pending'
        });
      }
    }
    return allDowntimes;
  } catch (e) {
    Logger.log(`Error in getAllDowntimes: ${e.message}`);
    throw e;
  }
}

function getDebugState() {
  const debugState = PropertiesService.getScriptProperties().getProperty(PROP_DISCORD_TEST_MODE);
  return debugState === 'true' ? 'Debug' : 'Live';
}

function startSending() {
  isStopped = false;
  isPaused = false;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const allDowntimes = getAllDowntimes();
  const cache = CacheService.getScriptCache();
  cache.put(CACHE_KEY, JSON.stringify(allDowntimes));

  let successCount = 0;
  let failureCount = 0;

  for (let i = 0; i < allDowntimes.length; i++) {
    const downtime = allDowntimes[i];
    if (downtime.isSent) {
      continue;
    }

    if (isStopped) {
      break;
    }
    while (isPaused) {
      Utilities.sleep(1000);
    }

    try {
      downtime.status = 'in-progress';
      cache.put(CACHE_KEY, JSON.stringify(allDowntimes));

      const success = handleSendDiscord_(sheet, downtime.rowIndex, true);
      if (success) {
        downtime.status = 'sent';
        downtime.isSent = true;
        successCount++;
      } else {
        downtime.status = 'error';
        downtime.error = 'Failed to send.';
        failureCount++;
      }
    } catch (e) {
      downtime.status = 'error';
      downtime.error = e.message;
      failureCount++;
    }
    cache.put(CACHE_KEY, JSON.stringify(allDowntimes));
  }

  const summary = { successCount, failureCount };
  cache.put(SUMMARY_CACHE_KEY, JSON.stringify(summary));
}

function getSendStatus() {
  const cache = CacheService.getScriptCache();
  const status = cache.get(CACHE_KEY);
  return status ? JSON.parse(status) : [];
}

function getSendSummary() {
  const cache = CacheService.getScriptCache();
  const summary = cache.get(SUMMARY_CACHE_KEY);
  return summary ? JSON.parse(summary) : { successCount: 0, failureCount: 0 };
}

function stopSending() {
  isStopped = true;
}

function pauseSending() {
  isPaused = true;
}

function resumeSending() {
  isPaused = false;
}
