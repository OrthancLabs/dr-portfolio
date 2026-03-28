/**
 * Sample Apps Script for syncing sheet rows to an external CRM API.
 */

const CRM_SYNC_CONFIG = {
  sheetName: 'CRM Sync',
  headerRow: 1,
  apiUrl: 'https://api.example.com/crm/records',
  maxRetries: 3,
};

/**
 * Run the sync for all rows that are new or need another attempt.
 */
function syncRowsToCrm() {
  const sheet = getCrmSyncSheet();
  const headers = getCrmHeaders(sheet);
  const records = getCrmRecords(sheet, headers);

  records.forEach((record) => {
    const syncStatus = String(record['Sync Status'] || '').toLowerCase();

    if (syncStatus === 'synced') {
      return;
    }

    const payload = buildCrmPayload(record);
    const result = sendWithRetry(payload, CRM_SYNC_CONFIG.maxRetries);

    if (result.success) {
      updateCrmSyncStatus(sheet, headers, record._rowNumber, {
        'Sync Status': 'Synced',
        'Last Sync At': new Date(),
        'Sync Notes': 'Record pushed successfully',
      });
    } else {
      updateCrmSyncStatus(sheet, headers, record._rowNumber, {
        'Sync Status': 'Failed',
        'Last Sync At': new Date(),
        'Sync Notes': result.message,
      });
    }
  });
}

/**
 * Optional onEdit hook if you want to sync as rows are updated.
 */
function onEdit(e) {
  const sheet = e.range.getSheet();

  if (sheet.getName() !== CRM_SYNC_CONFIG.sheetName || e.range.getRow() <= CRM_SYNC_CONFIG.headerRow) {
    return;
  }

  syncRowsToCrm();
}

/**
 * Build a simple payload from one sheet row.
 */
function buildCrmPayload(record) {
  return {
    firstName: record['First Name'] || '',
    lastName: record['Last Name'] || '',
    email: record.Email || '',
    company: record.Company || '',
    status: record.Status || '',
  };
}

/**
 * Send the payload and retry a few times before giving up.
 */
function sendWithRetry(payload, maxRetries) {
  let attempt = 0;

  while (attempt < maxRetries) {
    attempt += 1;

    try {
      const response = UrlFetchApp.fetch(CRM_SYNC_CONFIG.apiUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });

      const responseCode = response.getResponseCode();

      if (responseCode >= 200 && responseCode < 300) {
        return { success: true, message: 'OK' };
      }

      logSyncEvent(`Attempt ${attempt} failed with status ${responseCode}`);
    } catch (error) {
      logSyncEvent(`Attempt ${attempt} failed with error: ${error.message}`);
    }
  }

  return { success: false, message: `Failed after ${maxRetries} attempts` };
}

/**
 * Create the sheet and headers if they do not exist yet.
 */
function getCrmSyncSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CRM_SYNC_CONFIG.sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CRM_SYNC_CONFIG.sheetName);
    sheet.appendRow([
      'First Name',
      'Last Name',
      'Email',
      'Company',
      'Status',
      'Sync Status',
      'Last Sync At',
      'Sync Notes',
    ]);
  }

  return sheet;
}

/**
 * Read the header row.
 */
function getCrmHeaders(sheet) {
  return sheet.getRange(CRM_SYNC_CONFIG.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * Convert sheet rows into objects and keep the row number for updates.
 */
function getCrmRecords(sheet, headers) {
  const lastRow = sheet.getLastRow();

  if (lastRow <= CRM_SYNC_CONFIG.headerRow) {
    return [];
  }

  const rows = sheet
    .getRange(CRM_SYNC_CONFIG.headerRow + 1, 1, lastRow - CRM_SYNC_CONFIG.headerRow, headers.length)
    .getValues();

  return rows.map((row, index) => {
    const record = { _rowNumber: CRM_SYNC_CONFIG.headerRow + 1 + index };

    headers.forEach((header, columnIndex) => {
      record[header] = row[columnIndex];
    });

    return record;
  });
}

/**
 * Update sync fields for one row.
 */
function updateCrmSyncStatus(sheet, headers, rowNumber, updates) {
  Object.keys(updates).forEach((columnName) => {
    const columnIndex = headers.indexOf(columnName);

    if (columnIndex === -1) {
      return;
    }

    sheet.getRange(rowNumber, columnIndex + 1).setValue(updates[columnName]);
  });
}

/**
 * Basic logging helper for sync attempts.
 */
function logSyncEvent(message) {
  Logger.log(`[CRM Sync] ${message}`);
}

/**
 * Create a scheduled trigger for the sync job.
 */
function createCrmSyncTrigger() {
  ScriptApp.newTrigger('syncRowsToCrm')
    .timeBased()
    .everyHours(1)
    .create();
}
