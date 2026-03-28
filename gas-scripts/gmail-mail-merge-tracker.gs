/**
 * Sample Apps Script for a Gmail mail merge tied to a sheet.
 */

const MAIL_MERGE_CONFIG = {
  sheetName: 'Outreach',
  headerRow: 1,
  senderName: 'Sample Team',
  subjectTemplate: 'Hello {{First Name}}, a quick follow-up',
};

/**
 * Send emails for rows that have not been sent yet.
 */
function sendMailMergeEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIL_MERGE_CONFIG.sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${MAIL_MERGE_CONFIG.sheetName}" was not found.`);
  }

  const headers = getMailMergeHeaders(sheet);
  const records = getMailMergeRecords(sheet, headers);

  records.forEach((record) => {
    if (String(record['Sent Status'] || '').toLowerCase() === 'sent') {
      return;
    }

    if (!record.Email) {
      return;
    }

    const subject = applyTemplate(MAIL_MERGE_CONFIG.subjectTemplate, record);
    const body = buildPersonalizedEmail(record);

    // Send the email for this row.
    GmailApp.sendEmail(record.Email, subject, body, {
      name: MAIL_MERGE_CONFIG.senderName,
    });

    // Mark the row as sent.
    updateMailMergeStatus(sheet, headers, record._rowNumber, {
      'Sent Status': 'Sent',
      'Sent At': new Date(),
      'Reply Status': 'Awaiting Reply',
    });
  });
}

/**
 * Check Gmail for replies and update the sheet.
 */
function updateReplyStatuses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIL_MERGE_CONFIG.sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${MAIL_MERGE_CONFIG.sheetName}" was not found.`);
  }

  const headers = getMailMergeHeaders(sheet);
  const records = getMailMergeRecords(sheet, headers);

  records.forEach((record) => {
    const sentStatus = String(record['Sent Status'] || '').toLowerCase();

    if (sentStatus !== 'sent' || !record.Email) {
      return;
    }

    // Look for threads involving this contact.
    const threads = GmailApp.search(`to:${record.Email} OR from:${record.Email}`, 0, 10);

    const hasReply = threads.some((thread) => {
      const messages = thread.getMessages();

      return messages.some((message) => {
        const fromAddress = message.getFrom().toLowerCase();
        return fromAddress.indexOf(String(record.Email).toLowerCase()) !== -1;
      });
    });

    if (hasReply) {
      updateMailMergeStatus(sheet, headers, record._rowNumber, {
        'Reply Status': 'Responded',
        'Last Reply Check': new Date(),
      });
    } else {
      updateMailMergeStatus(sheet, headers, record._rowNumber, {
        'Last Reply Check': new Date(),
      });
    }
  });
}

/**
 * Build the email body from row data.
 */
function buildPersonalizedEmail(record) {
  return [
    `Hi ${record['First Name'] || 'there'},`,
    '',
    `I wanted to reach out regarding ${record['Company'] || 'your team'}.`,
    `Based on your domain (${record['Domain'] || 'n/a'}), this seemed relevant.`,
    '',
    'This is a sample Gmail mail merge generated from Google Sheets.',
    '',
    'Best,',
    MAIL_MERGE_CONFIG.senderName,
  ].join('\n');
}

/**
 * Replace template placeholders with row values.
 */
function applyTemplate(template, record) {
  return template.replace(/\{\{([^}]+)\}\}/g, (_, key) => {
    return record[key.trim()] || '';
  });
}

/**
 * Read the header row.
 */
function getMailMergeHeaders(sheet) {
  return sheet.getRange(MAIL_MERGE_CONFIG.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * Convert sheet rows into objects and keep the source row number.
 */
function getMailMergeRecords(sheet, headers) {
  const lastRow = sheet.getLastRow();

  if (lastRow <= MAIL_MERGE_CONFIG.headerRow) {
    return [];
  }

  const rows = sheet
    .getRange(MAIL_MERGE_CONFIG.headerRow + 1, 1, lastRow - MAIL_MERGE_CONFIG.headerRow, headers.length)
    .getValues();

  return rows.map((row, index) => {
    const record = { _rowNumber: MAIL_MERGE_CONFIG.headerRow + 1 + index };

    headers.forEach((header, columnIndex) => {
      record[header] = row[columnIndex];
    });

    return record;
  });
}

/**
 * Update status fields for one row.
 */
function updateMailMergeStatus(sheet, headers, rowNumber, updates) {
  Object.keys(updates).forEach((columnName) => {
    const columnIndex = headers.indexOf(columnName);

    if (columnIndex === -1) {
      return;
    }

    sheet.getRange(rowNumber, columnIndex + 1).setValue(updates[columnName]);
  });
}
