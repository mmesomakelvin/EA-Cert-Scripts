# EA-Cert-Scripts

Google Apps Scripts for automating certificate distribution across EduBridge Academy programs. Each script reads student data from a Google Sheet, attaches certificate PDFs from Google Drive, and sends personalized emails via Gmail.

## Programs

### Data School Program (`data-school-program/`)

Certificate and feedback distribution for **Data School Program** cohorts.

**What it does:**
- Reads student records from a Google Sheet (name, email, scores, certificate links)
- Generates a personalized **feedback PDF** for each student with performance metrics:
  - Attendance Score, Punctuality Score, Assessment Score, Individual Classwork, Presentation Score
- Determines which certificates each student earned (Attendance, Proficiency, both, or none)
- Sends one of four email templates based on certificate eligibility
- Attaches the feedback PDF and any earned certificate PDFs from Google Drive

**Key functions:**
- `sendCertificateEmails()` — Main entry point. Iterates through the sheet and sends all emails
- `createFeedbackPdf()` — Generates a temporary Google Doc with scores, converts to PDF
- `getGoogleDrivePdf()` — Retrieves certificate files from Google Drive links

**Sheet columns required:**
`NAME`, `EMAIL ADDRESS`, `Certificate of Attendance`, `Certificate of Proficiency`, `Attendance Score`, `Punctuality Score`, `Assessment Score`, `Individual Classwork`, `Presentation Score`, `Percentage %`

---

### Analyst Program (`analyst-program/`)

Certificate distribution for **Analyst Program** cohorts with multi-track support and a built-in test UI.

**What it does:**
- Reads student records from a Google Sheet with flexible header matching
- Supports **multiple tracks** (Finance, Management Consulting) with per-track email sending
- Provides a **Test Dialog** (HTML sidebar) for previewing and testing emails before live sends
- Adds a custom **Google Sheets menu** ("Analyst Certificates") for easy access to all actions
- Includes a test sheet workflow: copies a real row into a Test sheet for safe experimentation

**Key functions:**
- `onOpen()` — Creates the Analyst Certificates menu in Google Sheets
- `sendCertificateEmails()` — Sends certificates for the active sheet
- `sendFinanceSheetEmails()` / `sendManagementConsultingSheetEmails()` — Track-specific sends
- `showTestDialog()` — Opens the test UI for previewing/sending test emails
- `prepareTestSheet()` — Copies a data row to a Test sheet for safe testing
- `previewActiveRowEmail()` — Shows a preview of what would be sent for the selected row
- `processCertificateRows_()` — Core engine that handles live, test, and preview modes

**Sheet columns required:**
`Email Address`, `Name`, `Track`, `Attendance Certificate`, `Completion`, `Proficiency`, `Proficiency Certificate`

**Configuration** (top of `Code.js`):
```javascript
const CONFIG = {
  cohortName: "January 2026 Analyst Program",
  nextCohortDate: "April 11, 2026",
  registrationLink: "bit.ly/PhysicalGP",
  contactNumber: "07030146818"
};
```

---

## Setup

### Prerequisites
- A Google Sheet with student data (see column requirements above)
- Certificate PDFs stored in Google Drive
- [clasp](https://github.com/niccokunzmann/clasp) CLI installed (optional, for local development)

### Deploying a Script

1. Navigate to the program folder:
   ```bash
   cd analyst-program   # or data-school-program
   ```

2. Push to Google Apps Script:
   ```bash
   clasp push
   ```

3. Open in the Apps Script editor:
   ```bash
   clasp open
   ```

4. Run the main function from the script editor or (for Analyst Program) use the custom menu in Google Sheets.

### Local Development

Each program folder contains its own `.clasp.json` with the Apps Script project ID. To work on a specific program:

```bash
cd data-school-program
clasp pull    # Pull latest from Apps Script
# Edit Code.js
clasp push    # Deploy changes
```

## Project Structure

```
EA-Cert-Scripts/
├── README.md
├── data-school-program/
│   ├── .clasp.json          # Apps Script project ID for Data School
│   ├── Code.js              # Main script (email + feedback PDF generation)
│   └── appsscript.json      # Apps Script manifest
└── analyst-program/
    ├── .clasp.json          # Apps Script project ID for Analyst Program
    ├── Code.js              # Main script (multi-track, test modes, config)
    ├── TestDialog.html      # Test UI for previewing/sending test emails
    └── appsscript.json      # Apps Script manifest
```

## Google APIs Used

- **SpreadsheetApp** — Read student data from Google Sheets
- **DocumentApp** — Create feedback documents (Data School only)
- **DriveApp** — Retrieve certificate PDFs from Google Drive
- **GmailApp** — Send personalized emails with attachments
- **HtmlService** — Render the test dialog UI (Analyst Program only)
