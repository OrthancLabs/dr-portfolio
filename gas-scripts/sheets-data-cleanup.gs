/**
 * Sample Apps Script for sheet cleanup, validation, and enrichment.
 */

const CLEANUP_CONFIG = {
  sheetName: 'Leads',
  headerRow: 1,
  domainColumnName: 'Domain',
  companyColumnName: 'Company Name',
  industryColumnName: 'Industry',
  statusColumnName: 'Status',
  allowedStatuses: ['New', 'Qualified', 'Rejected'],
};

/**
 * Runs on edit.
 */
function onEdit(e) {
  const sheet = e.range.getSheet();

  if (sheet.getName() !== CLEANUP_CONFIG.sheetName) {
    return;
  }

  if (e.range.getRow() <= CLEANUP_CONFIG.headerRow) {
    return;
  }

  const headers = getHeaders(sheet);
  const rowNumber = e.range.getRow();
  const rowData = getRowAsObject(sheet, headers, rowNumber);

  // Format the row first.
  formatRow(sheet, headers, rowNumber, rowData);

  // Block duplicate domains.
  enforceUniqueDomain(sheet, headers, rowNumber, rowData);

  // Validate status values.
  validateStatus(rowData);

  // Fill in company data from the domain.
  enrichCompanyData(sheet, headers, rowNumber, rowData);
}

/**
 * Read the header row.
 */
function getHeaders(sheet) {
  const lastColumn = sheet.getLastColumn();
  return sheet.getRange(CLEANUP_CONFIG.headerRow, 1, 1, lastColumn).getValues()[0];
}

/**
 * Convert one row into an object.
 */
function getRowAsObject(sheet, headers, rowNumber) {
  const rowValues = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const rowObject = {};

  headers.forEach((header, index) => {
    rowObject[header] = rowValues[index];
  });

  return rowObject;
}

/**
 * Apply simple row formatting.
 */
function formatRow(sheet, headers, rowNumber, rowData) {
  const statusColumnIndex = headers.indexOf(CLEANUP_CONFIG.statusColumnName);

  if (statusColumnIndex === -1) {
    return;
  }

  const rowRange = sheet.getRange(rowNumber, 1, 1, headers.length);
  const statusValue = String(rowData[CLEANUP_CONFIG.statusColumnName] || '').trim();

  // Reset the row style before applying status colors.
  rowRange.setBackground('#ffffff').setFontWeight('normal');

  if (statusValue === 'Qualified') {
    rowRange.setBackground('#e7f4e4').setFontWeight('bold');
  } else if (statusValue === 'Rejected') {
    rowRange.setBackground('#fbe4e4');
  }
}

/**
 * Prevent duplicate domains.
 */
function enforceUniqueDomain(sheet, headers, rowNumber, rowData) {
  const domainColumnIndex = headers.indexOf(CLEANUP_CONFIG.domainColumnName);

  if (domainColumnIndex === -1) {
    return;
  }

  const domain = String(rowData[CLEANUP_CONFIG.domainColumnName] || '').trim().toLowerCase();

  if (!domain) {
    return;
  }

  const domainValues = sheet
    .getRange(CLEANUP_CONFIG.headerRow + 1, domainColumnIndex + 1, sheet.getLastRow() - CLEANUP_CONFIG.headerRow, 1)
    .getValues()
    .flat()
    .map((value) => String(value || '').trim().toLowerCase());

  const duplicateCount = domainValues.filter((value) => value === domain).length;

  if (duplicateCount > 1) {
    sheet.getRange(rowNumber, domainColumnIndex + 1).clearContent();
    SpreadsheetApp.getUi().alert(`Duplicate domain blocked: ${domain}`);
    throw new Error(`Duplicate domain detected: ${domain}`);
  }
}

/**
 * Validate the Status value.
 */
function validateStatus(rowData) {
  const status = String(rowData[CLEANUP_CONFIG.statusColumnName] || '').trim();

  if (!status) {
    return;
  }

  if (!CLEANUP_CONFIG.allowedStatuses.includes(status)) {
    throw new Error(`Invalid status value: ${status}`);
  }
}

/**
 * Fetch company data from a sample API.
 */
function enrichCompanyData(sheet, headers, rowNumber, rowData) {
  const domain = String(rowData[CLEANUP_CONFIG.domainColumnName] || '').trim().toLowerCase();

  if (!domain) {
    return;
  }

  const companyColumnIndex = headers.indexOf(CLEANUP_CONFIG.companyColumnName);
  const industryColumnIndex = headers.indexOf(CLEANUP_CONFIG.industryColumnName);

  if (companyColumnIndex === -1 || industryColumnIndex === -1) {
    return;
  }

  // Skip rows that already have company data.
  if (rowData[CLEANUP_CONFIG.companyColumnName]) {
    return;
  }

  const url = `https://autocomplete.clearbit.com/v1/companies/suggest?query=${encodeURIComponent(domain)}`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

  if (response.getResponseCode() !== 200) {
    return;
  }

  const companies = JSON.parse(response.getContentText());

  if (!companies.length) {
    return;
  }

  const company = companies[0];

  // Write the results back to the sheet.
  sheet.getRange(rowNumber, companyColumnIndex + 1).setValue(company.name || '');
  sheet.getRange(rowNumber, industryColumnIndex + 1).setValue(company.category && company.category.industry || '');
}
