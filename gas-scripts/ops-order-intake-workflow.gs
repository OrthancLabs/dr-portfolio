/**
 * Sample Apps Script for order intake from Gmail into Sheets.
 */

const ORDER_INTAKE_CONFIG = {
  inboxQuery: 'label:inbox subject:"Order" newer_than:7d',
  sheetName: 'Orders',
  summaryRecipient: 'ops@example.com',
  summarySubject: 'Daily Order Intake Summary',
};

/**
 * Main job: read order emails, parse them, and write them to the sheet.
 */
function processOrderEmails() {
  const sheet = getOrderSheet();
  const existingMessageIds = getExistingMessageIds(sheet);
  const threads = GmailApp.search(ORDER_INTAKE_CONFIG.inboxQuery, 0, 20);
  let newOrders = 0;

  threads.forEach((thread) => {
    thread.getMessages().forEach((message) => {
      const messageId = message.getId();

      if (existingMessageIds.has(messageId)) {
        return;
      }

      const order = parseOrderEmail(message);

      if (!order) {
        return;
      }

      sheet.appendRow([
        new Date(),
        messageId,
        order.customerName,
        order.company,
        order.product,
        order.quantity,
        order.requestedDate,
        order.notes,
      ]);

      newOrders += 1;
    });
  });

  return newOrders;
}

/**
 * Pull a few basic order fields from the email body.
 * This is sample parsing, so it assumes the email follows a loose pattern.
 */
function parseOrderEmail(message) {
  const body = message.getPlainBody();

  return {
    customerName: matchField(body, 'Customer'),
    company: matchField(body, 'Company'),
    product: matchField(body, 'Product'),
    quantity: matchField(body, 'Quantity'),
    requestedDate: matchField(body, 'Requested Date'),
    notes: matchField(body, 'Notes'),
  };
}

/**
 * Find a "Field: value" pattern in the email body.
 */
function matchField(text, fieldName) {
  const regex = new RegExp(`${fieldName}:\\s*(.*)`, 'i');
  const match = text.match(regex);
  return match ? match[1].trim() : '';
}

/**
 * Send a short summary email for the day's intake run.
 */
function sendDailyOrderSummary() {
  const sheet = getOrderSheet();
  const lastRow = sheet.getLastRow();
  const ordersToday = Math.max(lastRow - 1, 0);
  const newOrders = processOrderEmails();

  const summary = [
    'Daily Order Intake Summary',
    `Run time: ${new Date().toLocaleString()}`,
    `New orders added this run: ${newOrders}`,
    `Total orders in sheet: ${ordersToday + newOrders}`,
  ].join('\n');

  MailApp.sendEmail({
    to: ORDER_INTAKE_CONFIG.summaryRecipient,
    subject: ORDER_INTAKE_CONFIG.summarySubject,
    body: summary,
  });
}

/**
 * Create the target sheet and headers if they do not exist yet.
 */
function getOrderSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(ORDER_INTAKE_CONFIG.sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(ORDER_INTAKE_CONFIG.sheetName);
    sheet.appendRow([
      'Processed At',
      'Message ID',
      'Customer',
      'Company',
      'Product',
      'Quantity',
      'Requested Date',
      'Notes',
    ]);
  }

  return sheet;
}

/**
 * Keep track of processed emails so they are not imported twice.
 */
function getExistingMessageIds(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return new Set();
  }

  const messageIds = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
  return new Set(messageIds.filter(Boolean));
}

/**
 * Create a daily trigger for the summary job.
 */
function createOrderIntakeTrigger() {
  ScriptApp.newTrigger('sendDailyOrderSummary')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
}
