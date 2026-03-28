/**
 * Sample Apps Script that reads a sheet, builds a summary,
 * exports it as PDF, and emails it once a day.
 */

const CONFIG = {
  sheetName: 'Daily Metrics',
  recipientEmail: 'client@example.com',
  emailSubject: 'Daily Sheet Summary Report',
  pdfFileNamePrefix: 'Daily_Sheet_Summary',
};

/**
 * Main entry point for the daily job.
 */
function sendDailySummaryEmail() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${CONFIG.sheetName}" was not found.`);
  }

  // Read the whole sheet.
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    throw new Error('The sheet does not contain any data rows.');
  }

  const headers = values[0];
  const rows = values.slice(1);

  // Map each row to an object keyed by the headers.
  const records = rows.map((row) => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = row[index];
    });
    return record;
  });

  // Build the summary text.
  const summary = buildDailySummary(records);

  // Build a temporary doc and export it as PDF.
  const pdfBlob = createSummaryPdf(summary);

  // Email the PDF.
  MailApp.sendEmail({
    to: CONFIG.recipientEmail,
    subject: CONFIG.emailSubject,
    body: summary,
    attachments: [pdfBlob],
  });
}

/**
 * Build a simple summary from the sheet data.
 */
function buildDailySummary(records) {
  const totalRows = records.length;

  const completedCount = records.filter((record) => {
    return String(record.Status || '').toLowerCase() === 'completed';
  }).length;

  const totalAmount = records.reduce((sum, record) => {
    return sum + (Number(record.Amount) || 0);
  }, 0);

  return [
    'Daily Summary Report',
    `Generated: ${new Date().toLocaleString()}`,
    '',
    `Total rows reviewed: ${totalRows}`,
    `Completed items: ${completedCount}`,
    `Total amount: $${totalAmount.toFixed(2)}`,
    '',
    'This PDF was generated automatically by Google Apps Script.',
  ].join('\n');
}

/**
 * Create a temporary doc and export it as PDF.
 */
function createSummaryPdf(summaryText) {
  const doc = DocumentApp.create(`${CONFIG.pdfFileNamePrefix}_${formatDate(new Date())}`);
  doc.getBody().appendParagraph(summaryText);
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  const pdfBlob = file.getAs(MimeType.PDF).setName(`${file.getName()}.pdf`);

  // Clean up the temp doc.
  file.setTrashed(true);

  return pdfBlob;
}

/**
 * Create the daily trigger.
 */
function createDailyTrigger() {
  ScriptApp.newTrigger('sendDailySummaryEmail')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}

/**
 * Remove existing triggers while testing.
 */
function deleteExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'sendDailySummaryEmail') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
