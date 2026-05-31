# Lumepay Invoice Automation

A Google Apps Script that automatically scans incoming invoices from Gmail, extracts structured data using Claude AI, and logs them to a Google Sheets payment tracker — with PDF archiving to Google Drive.

---

## Features

- Scans Gmail for emails to/from `expense@lumepay.com`
- Extracts invoice fields (supplier, amount, currency, dates, billed-to entity) via Claude AI
- Reads PDF attachments directly using the Anthropic Files API for higher accuracy
- Saves and names PDF invoices to a structured Google Drive folder hierarchy
- Writes extracted data to a Google Sheets payment tracker
- Calculates payment run dates (next Tuesday or Thursday before due date)
- Detects and skips duplicate invoices
- Flags low-confidence results for manual review
- Labels processed emails to prevent reprocessing
- Sends an end-of-day digest email summarising results

---

## Prerequisites

- A Google account with access to Gmail, Google Drive, and Google Sheets
- [Node.js](https://nodejs.org/) installed locally
- [clasp](https://github.com/google/clasp) (Google Apps Script CLI) v3+
- An [Anthropic API key](https://console.anthropic.com/)

---

## Project Structure

```
Invoice Automation/
├── Code.js              # Main script (all logic)
├── appsscript.json      # Apps Script manifest (scopes, runtime, timezone)
├── .clasp.json          # clasp config (links to remote Apps Script project)
└── README.md
```

---

## Setup

### 1. Clone and configure clasp

```bash
npm install -g @google/clasp
clasp login
```

Ensure `.clasp.json` points to the correct Apps Script project ID.

### 2. Push the script

```bash
clasp push
```

### 3. Store your Anthropic API key

In the Apps Script editor, open **Project Settings → Script Properties** and add:

| Property | Value |
|---|---|
| `ANTHROPIC_API_KEY` | Your Anthropic API key |

### 4. Configure `CONFIG` in `Code.js`

| Key | Description |
|---|---|
| `SPREADSHEET_ID` | ID of the Google Sheet used as the payment tracker |
| `TARGET_SHEET` | Sheet tab name (default: `Payment run - Claude`) |
| `DRIVE_ROOT_FOLDER_ID` | ID of the root Google Drive folder for PDF storage |
| `DAYS_TO_SCAN` | How many days back to scan on each run (default: `1`) |
| `NOTIFY_EMAIL` | Email address for fatal error alerts |
| `DIGEST_RECIPIENTS` | Array of emails for the daily digest |
| `CONFIDENCE_THRESHOLD` | Minimum AI confidence to auto-save (default: `0.80`) |
| `PAYMENT_DAYS` | Days of week for payment runs — `2` = Tuesday, `4` = Thursday |

### 5. Run initial setup check

In the Apps Script editor, run `checkSetup` to verify Gmail access, API key, and Drive folder configuration.

### 6. Create triggers

Run `createDailyTrigger` once to set up automated time-based triggers. This schedules:
- `runDailyInvoiceScan` at 06:00, 08:00, 10:00, 12:00, 15:00, 17:00, 19:00 (Africa/Johannesburg)
- `sendEndOfDayDigest` at ~16:45 (Africa/Johannesburg)

---

## Usage

| Function | Description |
|---|---|
| `runDailyInvoiceScan` | Production scan — processes the last 1 day of emails |
| `testRun` | Test scan — processes the last 5 days of emails |
| `checkSetup` | Validates configuration and connectivity |
| `createDailyTrigger` | Creates all time-based triggers (run once) |
| `sendEndOfDayDigest` | Manually sends the daily summary email |

---

## Google Sheets Output

The script writes one row per invoice to the configured sheet:

| Column | Description |
|---|---|
| Invoice Number | Extracted invoice/tax invoice number |
| Customer Name | Supplier issuing the invoice |
| Entity Invoiced | The Lumepay entity being billed |
| Invoice Currency | ZAR, USD, EUR, GBP, etc. |
| Total incl. VAT | Total amount payable |
| Invoice Date | Date on the invoice |
| Invoice Due Date | Payment due date |
| Payment Run Date | Calculated Tuesday/Thursday before due date |
| Status | `Saved` or `⚠️ Flagged for review` |
| AI Confidence | Claude's confidence score (0–100%) |
| Source Email | Link to the original Gmail message |
| Source Invoice | Link to the PDF in Google Drive |
| Date Email Received | When the email arrived |
| Date Added | When the row was written |

Rows flagged for review are highlighted in yellow.

---

## Google Drive Structure

PDFs are saved under the configured root folder, organised by year and month:

```
Invoice Automation (root)/
└── 1. 2025/
    └── 4. April/
        └── Supplier Name - 25Apr25 - INV-001.pdf
```

---

## Duplicate Detection

An invoice is considered a duplicate if any of the following match an existing row:

1. Same Gmail message link
2. Same invoice number + supplier + billed-to entity
3. Same supplier + billed-to + amount + invoice date
4. Same supplier + invoice number + amount (when billed-to is absent on both)

---

## Development

To push local changes to Apps Script:

```bash
clasp push
```

To pull the remote version locally:

```bash
clasp pull
```

Enable `DEBUG_MODE: true` in `CONFIG` to log full API responses to the Apps Script execution log.

---

## License

Private — internal Lumepay tooling. Not for public distribution.
