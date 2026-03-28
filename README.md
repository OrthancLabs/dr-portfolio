# Dominic Romito Portfolio

A collection of production-inspired automation workflows and scripts demonstrating how to connect data, APIs, and business logic across systems. These examples focus on real operational use cases—eliminating manual work, structuring messy inputs, and delivering actionable outputs inside Google Workspace and adjacent systems.

## Patterns Demonstrated

- Event-driven automation (onEdit, onFormSubmit, time triggers)
- Cross-system integrations via APIs and webhooks
- Email automation and mail merge workflows
- Data validation, deduplication, and enrichment
- Document generation (PDF reports via Drive)
- Stateful workflows (tracking status across systems)
- Lightweight ETL pipelines using Google Sheets as a data layer
- Reusable component-based script design

## Workflow Design Approach

1. Identify repetitive manual processes
2. Define a system of record (typically Google Sheets)
3. Automate intake (Forms, Gmail parsing, APIs)
4. Apply validation and enrichment
5. Trigger outputs (emails, reports, notifications)
6. Structure for reuse across clients and workflows

## Projects

### fraud-detection-openai (Python)

**Use Case:**  
Evaluate transaction data and flag potential fraud using lightweight heuristics and AI-assisted reasoning.

**Inputs:**  
Mock transaction dataset

**Outputs:**  
Fraud classification + explanation

**Tech:**  
Python, OpenAI API

**Notes:**  
Demonstrates combining deterministic signals with LLM reasoning for decision support.

### hubspot-margin-aggregation (Node / HubSpot)

**Use Case:**  
Aggregate deal margin by product type inside HubSpot workflows.

**Trigger:**  
HubSpot custom code action

**Inputs:**  
Deal + associated line items

**Outputs:**  
Margin totals written back to HubSpot

**Tech:**  
Node.js, HubSpot API

### ingredient-allergy-matcher (Node)

**Use Case:**  
Detect allergens in unstructured ingredient text.

**Inputs:**  
Free-text ingredient strings

**Outputs:**  
Matched allergens with fuzzy fallback

**Tech:**  
Node.js, regex, fuzzy matching

**Notes:**  
Demonstrates handling messy, real-world text inputs.

## Google Apps Script Workflows

### daily-sheet-summary

**Use Case:**  
Automated daily reporting for operations teams

**Trigger:**  
Time-driven (daily)

**Inputs:**  
Google Sheet (operational data)

**Outputs:**  
PDF summary emailed to stakeholders

**Services Used:**  
Sheets, Drive, Gmail

**Before:**  
- Manual reporting  
- No consistent daily visibility  

**After:**  
- Automated reporting pipeline  
- Consistent, scheduled delivery  

### sheets-data-cleanup

**Use Case:**  
Real-time data validation and enrichment inside Sheets

**Trigger:**  
onEdit (event-driven)

**Inputs:**  
User-edited spreadsheet rows

**Outputs:**  
- Formatted rows  
- Duplicate prevention  
- Enriched data via API  

**Services Used:**  
Sheets, external API

**Notes:**  
Demonstrates enforcing business rules at the data entry layer.

### gmail-mail-merge-tracker

**Use Case:**  
Outbound communication + response tracking

**Trigger:**  
Manual / scheduled

**Inputs:**  
Recipient list in Google Sheets

**Outputs:**  
- Personalized Gmail messages  
- Send + reply status tracked in Sheet  

**Services Used:**  
Gmail, Sheets

**Notes:**  
Shows stateful workflow design across communication systems.

### ops-order-intake-workflow

**Use Case:**  
Automate order intake from unstructured email into a structured system

**Trigger:**  
Gmail-based ingestion + time-driven processing

**Inputs:**  
Inbound emails with order details

**Outputs:**  
- Structured records in Google Sheets  
- Daily ops summary email  

**Services Used:**  
Gmail, Sheets

**Before:**  
- Manual data entry from email  
- Inconsistent formatting  
- Limited visibility  

**After:**  
- Automated parsing and ingestion  
- Standardized dataset  
- Daily reporting  

**Notes:**  
Demonstrates transforming messy operational inputs into structured workflows.

### sheets-to-crm-sync

**Use Case:**  
Sync structured data from Google Sheets to external systems

**Trigger:**  
onEdit or scheduled sync

**Inputs:**  
New or updated rows in Sheets

**Outputs:**  
- Records pushed to external API (CRM or similar)  
- Sync status logged in Sheet  

**Services Used:**  
Sheets, external API (via UrlFetchApp)

**Notes:**  
Includes basic retry handling and logging pattern.

## AppSheet (In Progress)

### appsheet-field-ops-app

**Use Case:**  
Mobile-friendly tool for field teams to submit job updates

**Features:**  
- Form-based data entry  
- Status tracking  
- Photo uploads  
- Sync to Google Sheets backend  

**Purpose:**  
Demonstrates extending Workspace workflows to non-technical end users.

## Notes

- These workflows are intentionally lightweight but production-oriented
- Designed to reflect real operational use cases, not isolated code samples
- Most scripts are structured for reuse across multiple scenarios
- Google Apps Script examples are intended to run inside a Sheet-bound project