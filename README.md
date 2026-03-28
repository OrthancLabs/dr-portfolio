# dr-portfolio

Portfolio repo with sample scripts for client demos and workflow examples. This is a collection of standalone examples across Python, Node.js, HubSpot custom code, and Google Apps Script.

## Files

### Root Scripts

`bank-client-fraud-script`

Sample Python script that loads mock transaction data, builds a few simple fraud signals, and asks OpenAI for a fraud decision with a short explanation.

`custom-hubspot-lambda-function`

Sample HubSpot custom code action that looks up line items on a deal, groups ACV margin by product type, and returns totals back to HubSpot.

`allergy-matching`

Sample Node.js script that checks ingredient text for allergens using a synonym list, regex matching, and a simple fuzzy match fallback.

### Google Apps Script Samples

`gas-scripts/daily-sheet-summary.gs`

Sample Apps Script that reads data from a Google Sheet, builds a summary, exports that summary as a PDF, and emails it on a daily trigger.

`gas-scripts/sheets-data-cleanup.gs`

Sample Apps Script for Sheets that formats rows on edit, blocks duplicate entries, validates business rules, and enriches rows from a domain-based API lookup.

`gas-scripts/gmail-mail-merge-tracker.gs`

Sample Apps Script mail merge that reads recipients from a Google Sheet, sends personalized Gmail messages, and tracks sent and reply status back in the sheet.

## Notes

- These scripts are intentionally simple and sample-grade.
- Most scripts assume you will replace sample config values before using them.
- The Google Apps Script samples are meant to be copied into an Apps Script project attached to a Sheet.
