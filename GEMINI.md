# Gemini Codebase Analysis

This document provides an overview of the Google Apps Script project for sending personalized certificates and feedback to students.

## Project Overview

The project is designed to automate the distribution of certificates and personalized feedback to participants of the Data School Program. It reads student data from a Google Sheet, generates feedback PDFs, and sends customized emails with the appropriate attachments.

## File Descriptions

*   **`Source_Certificate/Code.js`**: This is the main script file containing all the logic for the application. It reads data from a spreadsheet, creates PDFs, and sends emails.
*   **`Source_Certificate/Aug 25 Cert Script/Aug_25_Cert_DS.js`**: This file is an identical copy of `Source_Certificate/Code.js`, likely a backup or a different version of the script.
*   **`.clasp.json`**: This is the project manifest file for the Google Apps Script CLI (`clasp`). It contains project settings and metadata.
*   **`appsscript.json`**: These files, located in both `Source_Certificate` and `Source_Certificate/Aug 25 Cert Script`, are the manifests for the Google Apps Script projects. They define the project's dependencies and required API scopes.

## Core Functionality

*   **Data Processing:** The script reads student information, including names, email addresses, and certificate links, from a Google Sheet.
*   **PDF Generation:** It dynamically creates a personalized feedback PDF for each student, including their scores and a summary of their performance.
*   **Email Automation:** The script sends emails to students with different content depending on whether they receive a "Certificate of Attendance," a "Certificate of Proficiency," both, or neither.
*   **Attachment Handling:** It attaches the generated feedback PDF and any earned certificates (retrieved from Google Drive) to the emails.

## Dependencies

The script utilizes the following Google Apps Script services:

*   **`SpreadsheetApp`**: To read data from Google Sheets.
*   **`DocumentApp`**: To create the feedback document that is then converted to a PDF.
*   **`DriveApp`**: To retrieve certificate files from Google Drive and to manage the temporary feedback documents.
*   **`GmailApp`**: To send the personalized emails.
*   **`Utilities`**: To introduce a delay between sending emails to avoid exceeding Google's quota limits.

## Setup and Usage

1.  **Prepare the Google Sheet:** Create a Google Sheet with the student data, including columns for "NAME", "EMAIL ADDRESS", "Certificate of Attendance" (with a link to the certificate in Google Drive), "Certificate of Proficiency" (with a link to the certificate in Google Drive), and the various score columns.
2.  **Open the Script Editor:** Open the Google Sheet and go to "Extensions" > "Apps Script".
3.  **Paste the Code:** Copy the code from either `Source_Certificate/Code.js` or `Source_Certificate/Aug 25 Cert Script/Aug_25_Cert_DS.js` into the script editor.
4.  **Run the Script:** Run the `sendCertificateEmails` function from the script editor. You will be prompted to authorize the script to access your Google account (Sheets, Docs, Drive, and Gmail).
5.  **Execution:** The script will then iterate through the rows of the spreadsheet, generate the feedback PDFs, and send the emails with the appropriate certificates and feedback attached.
