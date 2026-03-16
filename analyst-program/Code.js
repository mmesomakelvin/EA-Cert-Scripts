const CONFIG = {
  cohortName: "January 2026 Analyst Program",
  nextCohortDate: "April 11, 2026",
  registrationLink: "bit.ly/PhysicalGP",
  contactNumber: "07030146818",
  liveSendPauseMs: 750,
  testEmail: ""
};

const TEST_SHEET_NAME = "Test";
const FIRST_DATA_ROW = 2;
const SHEET_NAMES = {
  finance: "FINANCE",
  managementConsulting: "MANAGEMENT CONSULTING"
};

const REQUIRED_HEADERS = {
  email: ["email address", "email"],
  fullName: ["name", "full name"],
  track: ["track"],
  attendanceLink: ["attendance certificate"],
  completionStatus: ["completion"],
  proficiencyStatus: ["proficiency"],
  proficiencyLink: ["proficiency certificate"]
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Analyst Certificates")
    .addItem("Open Test Dialog", "showTestDialog")
    .addItem("Prepare Test Sheet", "prepareTestSheet")
    .addItem("Preview Test Sheet Email", "previewTestSheetEmail")
    .addItem("Send Test Sheet Email", "promptAndSendTestSheetEmail")
    .addSeparator()
    .addItem("Preview Active Row", "previewActiveRowEmail")
    .addItem("Send Test For Active Row", "sendTestEmailForActiveRow")
    .addSeparator()
    .addItem("Send Active Sheet Emails", "sendCertificateEmails")
    .addItem("Send FINANCE Emails", "sendFinanceSheetEmails")
    .addItem("Send MANAGEMENT CONSULTING Emails", "sendManagementConsultingSheetEmails")
    .addToUi();
}

function showTestDialog() {
  const html = HtmlService.createHtmlOutputFromFile("TestDialog")
    .setWidth(460)
    .setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, "Analyst Certificate Test");
}

function sendCertificateEmails() {
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const result = processCertificateRows_({
    mode: "live",
    sheet: activeSheet
  });
  showResultAlert_(`${activeSheet.getName()} Certificate Emails`, result);
}

function sendFinanceSheetEmails() {
  confirmAndSendNamedSheetEmails_(SHEET_NAMES.finance);
}

function sendManagementConsultingSheetEmails() {
  confirmAndSendNamedSheetEmails_(SHEET_NAMES.managementConsulting);
}

function sendTestEmailForActiveRow() {
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const rowNumber = activeSheet.getActiveRange().getRow();

  if (rowNumber <= 1) {
    throw new Error("Select a data row before running the test email.");
  }

  const result = processCertificateRows_({
    mode: "test",
    rowNumbers: [rowNumber]
  });
  showResultAlert_("Test Certificate Email", result);
}

function previewActiveRowEmail() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const rowNumber = sheet.getActiveRange().getRow();

  if (rowNumber <= 1) {
    throw new Error("Select a data row before previewing the email.");
  }

  const sheetData = getSheetData_(sheet);
  const row = sheetData.values[rowNumber - 1];
  const recipient = buildRecipientRecord_(row, rowNumber, sheetData.headerMap);
  const emailPayload = buildEmailPayload_(recipient, {
    mode: "preview",
    overrideEmail: resolveTestEmail_()
  });

  const attachmentNames = emailPayload.attachmentNames.length
    ? emailPayload.attachmentNames.join(", ")
    : "None";

  const previewMessage = [
    `Row: ${rowNumber}`,
    `Recipient: ${recipient.fullName || "(missing name)"}`,
    `Send To: ${emailPayload.to}`,
    `Subject: ${emailPayload.subject}`,
    `Certificate Type: ${recipient.certificateType}`,
    `Attachments: ${attachmentNames}`
  ].join("\n");

  SpreadsheetApp.getUi().alert("Analyst Certificate Preview", previewMessage, SpreadsheetApp.getUi().ButtonSet.OK);
}

function prepareTestSheet() {
  const result = prepareTestSheet_();
  SpreadsheetApp.getUi().alert("Test Sheet Ready", result.message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function prepareTestSheetForDialog() {
  return prepareTestSheet_();
}

function previewTestSheetEmail() {
  const recipientEmail = promptForRecipientEmail_();
  const preview = previewTestSheetEmailForDialog(recipientEmail);

  SpreadsheetApp.getUi().alert("Test Email Preview", preview.message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function promptAndSendTestSheetEmail() {
  const recipientEmail = promptForRecipientEmail_();
  const result = sendTestSheetEmailForDialog(recipientEmail);

  SpreadsheetApp.getUi().alert("Test Email Sent", result.message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function previewTestSheetEmailForDialog(recipientEmail) {
  const payload = buildTestSheetEmailPayload_(recipientEmail);
  const message = [
    `Sheet: ${TEST_SHEET_NAME}`,
    `Row: ${FIRST_DATA_ROW}`,
    `Recipient: ${payload.recipient.fullName || "(missing name)"}`,
    `Send To: ${payload.emailPayload.to}`,
    `Subject: ${payload.emailPayload.subject}`,
    `Certificate Type: ${payload.recipient.certificateType}`,
    `Attachments: ${payload.emailPayload.attachmentNames.length ? payload.emailPayload.attachmentNames.join(", ") : "None"}`
  ].join("\n");

  return { message: message };
}

function sendTestSheetEmailForDialog(recipientEmail) {
  const testSheet = getTestSheet_();
  const result = processCertificateRows_({
    mode: "test",
    sheet: testSheet,
    rowNumbers: [FIRST_DATA_ROW],
    overrideEmail: resolveRecipientEmail_(recipientEmail)
  });

  return {
    message: [
      `Mode: ${result.mode}`,
      `Sent: ${result.sentCount}`,
      `Skipped: ${result.skippedCount}`
    ].concat(result.logs).join("\n")
  };
}

function processCertificateRows_(options) {
  const mode = options.mode || "live";
  const sheet = options.sheet || SpreadsheetApp.getActiveSheet();
  const sheetData = getSheetData_(sheet);
  const rowNumbers = options.rowNumbers || buildAllRowNumbers_(sheetData.values.length);
  const logs = [];
  let sentCount = 0;
  let skippedCount = 0;

  rowNumbers.forEach(function(rowNumber) {
    const row = sheetData.values[rowNumber - 1];

    try {
      if (!row) {
        skippedCount += 1;
        logs.push(`Row ${rowNumber}: skipped because the row does not exist.`);
        return;
      }

      const recipient = buildRecipientRecord_(row, rowNumber, sheetData.headerMap);

      if (!recipient.email) {
        skippedCount += 1;
        logs.push(`Row ${rowNumber}: skipped because email address is missing.`);
        return;
      }

      if (recipient.certificateType === "none") {
        skippedCount += 1;
        logs.push(`Row ${rowNumber}: skipped because no certificate was found for ${recipient.fullName || "this row"}.`);
        return;
      }

      const sendOptions = {
        mode: mode,
        overrideEmail: options.overrideEmail || (mode === "test" ? resolveTestEmail_() : "")
      };
      const emailPayload = buildEmailPayload_(recipient, sendOptions);

      if (mode === "preview") {
        logs.push(`Row ${rowNumber}: previewed ${recipient.certificateType} email for ${recipient.fullName}.`);
        return;
      }

      GmailApp.sendEmail(emailPayload.to, emailPayload.subject, emailPayload.plainBody, {
        htmlBody: emailPayload.htmlBody,
        attachments: emailPayload.attachments
      });

      sentCount += 1;
      logs.push(`Row ${rowNumber}: ${mode} email sent to ${emailPayload.to} for ${recipient.fullName}.`);

      if (mode === "live") {
        Utilities.sleep(CONFIG.liveSendPauseMs);
      }
    } catch (error) {
      skippedCount += 1;
      logs.push(`Row ${rowNumber}: ${error.message}`);
    }
  });

  Logger.log(logs.join("\n"));

  return {
    mode: mode,
    sentCount: sentCount,
    skippedCount: skippedCount,
    logs: logs
  };
}

function getSheetData_(sheet) {
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    throw new Error("The active sheet does not have any data rows.");
  }

  const headerMap = mapHeaders_(values[0]);
  return {
    values: values,
    headerMap: headerMap
  };
}

function prepareTestSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = getSourceSheetForTest_(spreadsheet);
  const testSheet = spreadsheet.getSheetByName(TEST_SHEET_NAME) || spreadsheet.insertSheet(TEST_SHEET_NAME);
  const lastColumn = sourceSheet.getLastColumn();

  if (sourceSheet.getLastRow() < FIRST_DATA_ROW) {
    throw new Error(`Source sheet "${sourceSheet.getName()}" does not have a data row to copy.`);
  }

  testSheet.clear();
  sourceSheet.getRange(1, 1, FIRST_DATA_ROW, lastColumn).copyTo(testSheet.getRange(1, 1));
  testSheet.setFrozenRows(1);

  const headerMap = mapHeaders_(sourceSheet.getRange(1, 1, 1, lastColumn).getValues()[0]);
  const completionCell = testSheet.getRange(FIRST_DATA_ROW, headerMap.completionStatus + 1);
  const proficiencyCell = testSheet.getRange(FIRST_DATA_ROW, headerMap.proficiencyStatus + 1);

  completionCell.setBackground("#fff2cc").setNote('Set to "YES" to attach the Attendance Certificate or "NO" to skip it.');
  proficiencyCell.setBackground("#fff2cc").setNote('Set to "YES" to attach the Proficiency Certificate or "NO" to skip it.');

  testSheet.getRange("A4").setValue("Instructions");
  testSheet.getRange("A5").setValue("Row 2 is your test record copied from the source sheet.");
  testSheet.getRange("A6").setValue('Change Completion and Proficiency on row 2 to YES or NO, then use the test dialog buttons to preview or send.');
  testSheet.getRange("A7").setValue("The test email will only go to the address you enter in the popup.");
  testSheet.autoResizeColumns(1, lastColumn);
  spreadsheet.setActiveSheet(testSheet);

  return {
    sheetName: TEST_SHEET_NAME,
    sourceSheetName: sourceSheet.getName(),
    message: `Copied header row and the first data row from "${sourceSheet.getName()}" into "${TEST_SHEET_NAME}". Edit Completion and Proficiency on row 2, then preview or send the test email.`
  };
}

function confirmAndSendNamedSheetEmails_(sheetName) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Confirm Live Send",
    `Send certificate emails for "${sheetName}" only?`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response !== ui.Button.OK) {
    ui.alert("Live send cancelled.");
    return;
  }

  const sheet = getRequiredSheetByName_(sheetName);
  const result = processCertificateRows_({
    mode: "live",
    sheet: sheet
  });

  showResultAlert_(`${sheetName} Certificate Emails`, result);
}

function mapHeaders_(headerRow) {
  const normalizedHeaders = headerRow.map(function(header) {
    return normalizeHeader_(header);
  });

  const headerMap = {};

  Object.keys(REQUIRED_HEADERS).forEach(function(key) {
    const matches = REQUIRED_HEADERS[key];
    const columnIndex = normalizedHeaders.findIndex(function(header) {
      return matches.indexOf(header) !== -1;
    });

    if (columnIndex === -1) {
      throw new Error(`Missing required column: ${matches[0]}`);
    }

    headerMap[key] = columnIndex;
  });

  return headerMap;
}

function buildAllRowNumbers_(rowCount) {
  const rowNumbers = [];

  for (let rowNumber = 2; rowNumber <= rowCount; rowNumber += 1) {
    rowNumbers.push(rowNumber);
  }

  return rowNumbers;
}

function buildRecipientRecord_(row, rowNumber, headerMap) {
  const fullName = cleanString_(row[headerMap.fullName]);
  const email = cleanString_(row[headerMap.email]);
  const track = cleanString_(row[headerMap.track]);
  const attendanceLink = cleanString_(row[headerMap.attendanceLink]);
  const proficiencyLink = cleanString_(row[headerMap.proficiencyLink]);
  const completionStatus = isYes_(row[headerMap.completionStatus]);
  const proficiencyStatus = isYes_(row[headerMap.proficiencyStatus]);

  const wantsAttendance = completionStatus;
  const wantsProficiency = proficiencyStatus;
  const attachments = [];

  if (wantsAttendance) {
    if (!attendanceLink) {
      throw new Error(`missing Attendance Certificate link for ${fullName || "row " + rowNumber}`);
    }

    attachments.push(getCertificateAttachment_(attendanceLink, `${fullName || "Recipient"} - Attendance Certificate.pdf`));
  }

  if (wantsProficiency) {
    if (!proficiencyLink) {
      throw new Error(`missing Proficiency Certificate link for ${fullName || "row " + rowNumber}`);
    }

    attachments.push(getCertificateAttachment_(proficiencyLink, `${fullName || "Recipient"} - Proficiency Certificate.pdf`));
  }

  return {
    rowNumber: rowNumber,
    fullName: fullName,
    firstName: getFirstName_(fullName),
    email: email,
    track: track,
    attachments: attachments,
    certificateType: getCertificateType_(wantsAttendance, wantsProficiency)
  };
}

function buildTestSheetEmailPayload_(recipientEmail) {
  const testSheet = getTestSheet_();
  const sheetData = getSheetData_(testSheet);
  const row = sheetData.values[FIRST_DATA_ROW - 1];

  if (!row) {
    throw new Error(`The ${TEST_SHEET_NAME} sheet does not have a test row yet. Run "Prepare Test Sheet" first.`);
  }

  const recipient = buildRecipientRecord_(row, FIRST_DATA_ROW, sheetData.headerMap);

  if (recipient.certificateType === "none") {
    throw new Error('The test row has neither Completion nor Proficiency set to "YES".');
  }

  return {
    recipient: recipient,
    emailPayload: buildEmailPayload_(recipient, {
      mode: "test",
      overrideEmail: resolveRecipientEmail_(recipientEmail)
    })
  };
}

function buildEmailPayload_(recipient, options) {
  const mode = options.mode || "live";
  const to = options.overrideEmail || recipient.email;
  const content = getEmailContent_(recipient);
  const attachmentBlobs = recipient.attachments.map(function(item) {
    return item.blob;
  });
  const attachmentNames = recipient.attachments.map(function(item) {
    return item.name;
  });
  const subjectPrefix = mode === "test" ? "[TEST] " : "";

  return {
    to: to,
    subject: subjectPrefix + content.subject,
    plainBody: content.plainBody,
    htmlBody: content.htmlBody,
    attachments: attachmentBlobs,
    attachmentNames: attachmentNames
  };
}

function getEmailContent_(recipient) {
  const trackReference = formatTrackReference_(recipient.track);
  const greetingName = recipient.firstName || recipient.fullName || "there";
  let subject = "Your January 2026 Analyst Program Certificate";
  let recognitionLine = "";
  let attachmentLine = "";

  if (recipient.certificateType === "both") {
    subject = "Your January 2026 Analyst Program Certificates";
    recognitionLine = `We are pleased to provide you with both your Attendance Certificate and Proficiency Certificate, which recognize your full participation and mastery of ${trackReference}. These certificates reflect your dedication to professional growth and the skills you have developed during the program.`;
    attachmentLine = "Your certificates are attached to this email. Feel free to share them on your professional profiles and with your network.";
  } else if (recipient.certificateType === "attendance") {
    subject = "Your January 2026 Analyst Program Attendance Certificate";
    recognitionLine = `We are pleased to provide you with your Attendance Certificate, which recognizes your full participation in ${trackReference}. This certificate reflects your dedication to professional growth and the skills you have developed during the program.`;
    attachmentLine = "Your certificate is attached to this email. Feel free to share it on your professional profiles and with your network.";
  } else if (recipient.certificateType === "proficiency") {
    subject = "Your January 2026 Analyst Program Proficiency Certificate";
    recognitionLine = `We are pleased to provide you with your Proficiency Certificate, which recognizes your mastery of ${trackReference}. This certificate reflects your dedication to professional growth and the skills you have developed during the program.`;
    attachmentLine = "Your certificate is attached to this email. Feel free to share it on your professional profiles and with your network.";
  } else {
    throw new Error("No certificate available for this recipient.");
  }

  const paragraphs = [
    `Dear ${greetingName},`,
    `Congratulations on your outstanding achievement in the ${CONFIG.cohortName}! We are thrilled to recognize your commitment and excellence throughout the program.`,
    recognitionLine,
    attachmentLine,
    "Thank you for your active participation and remarkable performance. We hope the knowledge and skills you have gained will contribute significantly to your professional journey.",
    `As we prepare for the next cohort of the program starting on ${CONFIG.nextCohortDate}, we would truly appreciate your support in recommending friends, colleagues, or anyone in your network who may benefit from the program.`,
    `Registration Link: ${CONFIG.registrationLink}`,
    `Contact: ${CONFIG.contactNumber}`,
    "Thank you once again for being a part of the Analyst Program community.",
    "Best regards,\nThe Analyst Program Team"
  ];

  return {
    subject: subject,
    plainBody: paragraphs.join("\n\n"),
    htmlBody: buildHtmlBody_(paragraphs)
  };
}

function buildHtmlBody_(paragraphs) {
  const htmlParagraphs = paragraphs.map(function(paragraph) {
    return `<p>${escapeHtml_(paragraph).replace(/\n/g, "<br>")}</p>`;
  });

  return [
    '<div style="font-family: Arial, sans-serif; max-width: 640px; margin: 0 auto; line-height: 1.6;">',
    htmlParagraphs.join(""),
    "</div>"
  ].join("");
}

function getCertificateAttachment_(driveLink, fileName) {
  const fileId = extractDriveFileId_(driveLink);

  if (!fileId) {
    throw new Error(`invalid Google Drive link: ${driveLink}`);
  }

  const blob = DriveApp.getFileById(fileId).getAs(MimeType.PDF).setName(fileName);

  return {
    name: fileName,
    blob: blob
  };
}

function extractDriveFileId_(driveLink) {
  const match = String(driveLink || "").match(/[-\w]{25,}/);
  return match ? match[0] : "";
}

function resolveTestEmail_() {
  const email = CONFIG.testEmail || Session.getEffectiveUser().getEmail();

  if (!email) {
    throw new Error("Set CONFIG.testEmail before running a test email.");
  }

  return email;
}

function resolveRecipientEmail_(recipientEmail) {
  const email = cleanString_(recipientEmail);
  return email || resolveTestEmail_();
}

function getCertificateType_(hasAttendance, hasProficiency) {
  if (hasAttendance && hasProficiency) {
    return "both";
  }

  if (hasAttendance) {
    return "attendance";
  }

  if (hasProficiency) {
    return "proficiency";
  }

  return "none";
}

function getFirstName_(fullName) {
  const parts = cleanString_(fullName).split(/\s+/).filter(Boolean);
  return parts.length ? parts[0] : "";
}

function formatTrackReference_(track) {
  if (!track) {
    return "the program";
  }

  const trimmedTrack = track.replace(/\.$/, "");
  return /^the\s/i.test(trimmedTrack) ? trimmedTrack : `the ${trimmedTrack}`;
}

function normalizeHeader_(value) {
  return cleanString_(value).toLowerCase();
}

function cleanString_(value) {
  return String(value == null ? "" : value).trim();
}

function isYes_(value) {
  return cleanString_(value).toUpperCase() === "YES";
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function showResultAlert_(title, result) {
  const lines = [
    `Mode: ${result.mode}`,
    `Sent: ${result.sentCount}`,
    `Skipped: ${result.skippedCount}`
  ].concat(result.logs.slice(0, 8));

  SpreadsheetApp.getUi().alert(title, lines.join("\n"), SpreadsheetApp.getUi().ButtonSet.OK);
}

function getTestSheet_() {
  const testSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEST_SHEET_NAME);

  if (!testSheet) {
    throw new Error(`The ${TEST_SHEET_NAME} sheet does not exist yet. Run "Prepare Test Sheet" first.`);
  }

  return testSheet;
}

function getSourceSheetForTest_(spreadsheet) {
  const activeSheet = spreadsheet.getActiveSheet();

  if (activeSheet && activeSheet.getName() !== TEST_SHEET_NAME) {
    return activeSheet;
  }

  const sourceSheet = spreadsheet.getSheets().find(function(sheet) {
    return sheet.getName() !== TEST_SHEET_NAME && sheet.getLastRow() >= FIRST_DATA_ROW;
  });

  if (!sourceSheet) {
    throw new Error("No source sheet with data was found for test preparation.");
  }

  return sourceSheet;
}

function getRequiredSheetByName_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" was not found.`);
  }

  return sheet;
}

function promptForRecipientEmail_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Test Email Recipient",
    "Enter the email address that should receive this test message.",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    throw new Error("Test email cancelled.");
  }

  return resolveRecipientEmail_(response.getResponseText());
}
