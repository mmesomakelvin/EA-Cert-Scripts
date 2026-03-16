/**
 * Google Apps Script to send personalized certificate emails to students
 * Certificate distribution for May 2025 Data School Program cohort
 * Updated to use streamlined column structure
 */

function sendCertificateEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find column indices based on updated column names
  const nameIndex = headers.indexOf("NAME");
  const emailIndex = headers.indexOf("EMAIL ADDRESS");
  const attendanceCertIndex = headers.indexOf("Certificate of Attendance");
  const proficiencyCertIndex = headers.indexOf("Certificate of Proficiency");
  
  // Score columns for feedback PDF (streamlined)
  const attendanceScoreIndex = headers.indexOf("Attendance Score");
  const punctualityScoreIndex = headers.indexOf("Punctuality Score");
  const assessmentScoreIndex = headers.indexOf("Assessment Score");
  const individualClassworkIndex = headers.indexOf("Individual Classwork");
  const presentationScoreIndex = headers.indexOf("Presentation Score");
  const percentageIndex = headers.indexOf("Percentage %");
  
  // Start from row 1 (skip header row)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[nameIndex];
    const email = row[emailIndex];
    
    // Check if certificates exist
    const hasAttendanceCert = row[attendanceCertIndex] ? true : false;
    const hasProficiencyCert = row[proficiencyCertIndex] ? true : false;
    
    // Get certificate links
    const attendanceCertLink = row[attendanceCertIndex];
    const proficiencyCertLink = row[proficiencyCertIndex];
    
    // Skip if no email address
    if (!email) continue;
    
    // Create feedback PDF with scores
    const feedbackPdfBlob = createFeedbackPdf(
      name,
      row[attendanceScoreIndex],
      row[punctualityScoreIndex],
      row[assessmentScoreIndex],
      row[individualClassworkIndex],
      row[presentationScoreIndex],
      row[percentageIndex]
    );
    
    // Get certificate PDFs
    const attachments = [];
    let attendanceCertPdf = null;
    let proficiencyCertPdf = null;
    
    // Always include feedback PDF
    attachments.push({
      fileName: `${name} - Data School Program Feedback.pdf`,
      content: feedbackPdfBlob.getBytes(),
      mimeType: "application/pdf"
    });
    
    if (hasAttendanceCert) {
      attendanceCertPdf = getGoogleDrivePdf(attendanceCertLink);
      if (attendanceCertPdf) {
        attachments.push({
          fileName: `${name} - Data School Program Attendance Certificate.pdf`,
          content: attendanceCertPdf.getBlob().getBytes(),
          mimeType: "application/pdf"
        });
      }
    }
    
    if (hasProficiencyCert) {
      proficiencyCertPdf = getGoogleDrivePdf(proficiencyCertLink);
      if (proficiencyCertPdf) {
        attachments.push({
          fileName: `${name} - Data School Program Proficiency Certificate.pdf`,
          content: proficiencyCertPdf.getBlob().getBytes(),
          mimeType: "application/pdf"
        });
      }
    }
    
    // Send appropriate email based on the three criteria
    if (hasAttendanceCert && hasProficiencyCert) {
      // Criteria 3: Both certificates
      sendBothCertificatesEmail(name, email, attachments);
    } else if (hasAttendanceCert && !hasProficiencyCert) {
      // Criteria 1: Only attendance certificate
      sendAttendanceOnlyEmail(name, email, attachments);
    } else if (!hasAttendanceCert && hasProficiencyCert) {
      // Criteria 2: Only proficiency certificate
      sendProficiencyOnlyEmail(name, email, attachments);
    } else {
      // No certificates, but still send feedback
      sendFeedbackOnlyEmail(name, email, attachments);
    }
    
    // Add small delay to avoid quota limits
    Utilities.sleep(1000);
  }
}

/**
 * Create PDF with student's scores for feedback (streamlined version)
 */
function createFeedbackPdf(name, attendanceScore, punctualityScore, assessmentScore, individualClasswork, presentationScore, percentage) {
  // Create a temporary Google Doc with the score information
  const doc = DocumentApp.create(`${name} - Data School Program Feedback`);
  const body = doc.getBody();
  
  // Add content to the document
  body.appendParagraph("DATA SCHOOL PROGRAM - STUDENT FEEDBACK").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Name: ${name}`).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(`Date: July 27, 2025`);
  body.appendParagraph("").appendHorizontalRule();
  
  body.appendParagraph("PERFORMANCE METRICS").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  // Create a table for scores
  const table = body.appendTable([
    ["Metric", "Score"],
    ["Attendance", attendanceScore || "N/A"],
    ["Punctuality", punctualityScore || "N/A"],
    ["Assessment Score", assessmentScore || "N/A"],
    ["Individual Classwork", individualClasswork || "N/A"],
    ["Presentation Score", presentationScore || "N/A"],
    ["Overall Percentage", percentage ? `${parseFloat(percentage).toFixed(1)}%` : "N/A"]
  ]);
  
  // Format table
  table.setColumnWidth(0, 200);
  table.setColumnWidth(1, 100);
  
  body.appendParagraph("").appendHorizontalRule();
  body.appendParagraph("FEEDBACK SUMMARY").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  // Add general feedback based on overall percentage
  if (percentage >= 80) {
    body.appendParagraph("You have demonstrated excellent performance throughout the program. Your strong engagement, quality submissions, and collaborative efforts have been exemplary.");
  } else if (percentage >= 70) {
    body.appendParagraph("You have shown very good performance throughout the program. Your consistent engagement and quality work have been noted.");
  } else if (percentage >= 60) {
    body.appendParagraph("You have performed well throughout the program. With additional practice and engagement, you can further enhance your skills.");
  } else {
    body.appendParagraph("Thank you for your participation in the program. We recommend continued practice and engagement with the material to strengthen your skills.");
  }
  
  // Additional general guidance
  body.appendParagraph("\nRecommendations for further growth:");
  body.appendParagraph("1. Continue to apply the data analysis techniques learned in real-world scenarios");
  body.appendParagraph("2. Join industry communities to stay updated with the latest trends");
  body.appendParagraph("3. Consider pursuing advanced certifications to build on your current knowledge");
  
  // Add footer
  body.appendParagraph("").appendHorizontalRule();
  body.appendParagraph("The Data School Program Team\nMay 2025 Cohort");
  
  // Save and close the document
  doc.saveAndClose();
  
  // Get the PDF
  const pdf = DriveApp.getFileById(doc.getId()).getAs("application/pdf");
  
  // Delete the temporary document
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  
  return pdf;
}

/**
 * Get PDF file from Google Drive link
 */
function getGoogleDrivePdf(driveLink) {
  try {
    // Extract file ID from Drive link
    const fileId = driveLink.match(/[-\w]{25,}/);
    if (!fileId) return null;
    
    // Get file by ID
    return DriveApp.getFileById(fileId[0]);
  } catch (error) {
    Logger.log(`Error getting PDF: ${error.toString()}`);
    return null;
  }
}

/**
 * Send email for students receiving attendance certificate only (with feedback)
 */
function sendAttendanceOnlyEmail(name, email, attachments) {
  const subject = "Your Data School Program Attendance Certificate and Feedback";
  
  const htmlBody = `
  <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
    <p>Dear ${name},</p>
    
    <p>Thank you for your participation in the May 2025 Data School Program! We appreciate your dedication and engagement throughout the program.</p>
    
    <p>We are pleased to provide you with your Certificate of Attendance, recognizing your commitment to professional development and active participation in the program.</p>
    
    <p>Additionally, we've included your personalized feedback document, which provides insights into your performance metrics and areas where you've excelled, as well as recommendations for future growth.</p>
    
    <p>Your certificate and feedback are attached to this email. We encourage you to share your certificate on your professional profiles.</p>
    
    <p>We hope the insights and knowledge you've gained during the program will be valuable in your professional journey.</p>
    
    <p>Best regards,<br>
    The Data School Program Team</p>
  </div>`;
  
  const plainBody = `Dear ${name},

Thank you for your participation in the May 2025 Data School Program! We appreciate your dedication and engagement throughout the program.

We are pleased to provide you with your Certificate of Attendance, recognizing your commitment to professional development and active participation in the program.

Additionally, we've included your personalized feedback document, which provides insights into your performance metrics and areas where you've excelled, as well as recommendations for future growth.

Your certificate and feedback are attached to this email. We encourage you to share your certificate on your professional profiles.

We hope the insights and knowledge you've gained during the program will be valuable in your professional journey.

Best regards,
The Data School Program Team`;

  sendEmailWithHtml(email, subject, htmlBody, plainBody, attachments);
}

/**
 * Send email for students receiving proficiency certificate only (with feedback)
 */
function sendProficiencyOnlyEmail(name, email, attachments) {
  const subject = "Congratulations on Your Data School Program Proficiency Achievement";
  
  const htmlBody = `
  <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
    <p>Dear ${name},</p>
    
    <p>Congratulations on achieving proficiency in the May 2025 Data School Program! Your performance has been exceptional.</p>
    
    <p>We are pleased to provide you with your Certificate of Proficiency, which recognizes your mastery of the program content and successful demonstration of the required skills.</p>
    
    <p>Additionally, we've included your personalized feedback document, which provides insights into your performance metrics and areas where you've excelled, as well as recommendations for future growth.</p>
    
    <p>Your certificate and feedback are attached to this email. We encourage you to showcase your certificate on your professional profiles as a testament to your expertise.</p>
    
    <p>We hope the specialized skills you've developed will enhance your professional capabilities and open new opportunities for you.</p>
    
    <p>Best regards,<br>
    The Data School Program Team</p>
  </div>`;
  
  const plainBody = `Dear ${name},

Congratulations on achieving proficiency in the May 2025 Data School Program! Your performance has been exceptional.

We are pleased to provide you with your Certificate of Proficiency, which recognizes your mastery of the program content and successful demonstration of the required skills.

Additionally, we've included your personalized feedback document, which provides insights into your performance metrics and areas where you've excelled, as well as recommendations for future growth.

Your certificate and feedback are attached to this email. We encourage you to showcase your certificate on your professional profiles as a testament to your expertise.

We hope the specialized skills you've developed will enhance your professional capabilities and open new opportunities for you.

Best regards,
The Data School Program Team`;

  sendEmailWithHtml(email, subject, htmlBody, plainBody, attachments);
}

/**
 * Send email for students receiving both attendance and proficiency certificates (with feedback)
 */
function sendBothCertificatesEmail(name, email, attachments) {
  const subject = "Congratulations on Completing the Data School Program - Your Certificates";
  
  const htmlBody = `
  <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
    <p>Dear ${name},</p>
    
    <p>Congratulations on your outstanding achievement in the May 2025 Data School Program! We are thrilled to recognize your commitment and excellence throughout the program.</p>
    
    <p>We are pleased to provide you with both your Certificate of Attendance and Certificate of Proficiency, which recognize your full participation and mastery of the program content. These certificates reflect your dedication to professional growth and the skills you've developed during the program.</p>
    
    <p>Additionally, we've included your personalized feedback document, which provides insights into your performance metrics and areas where you've excelled, as well as recommendations for future growth.</p>
    
    <p>Your certificates and feedback are attached to this email. Feel free to share your certificates on your professional profiles and with your network.</p>
    
    <p>Thank you for your active participation and remarkable performance. We hope the knowledge and skills you've gained will contribute significantly to your professional journey.</p>
    
    <p>Best regards,<br>
    The Data School Program Team</p>
  </div>`;
  
  const plainBody = `Dear ${name},

Congratulations on your outstanding achievement in the May 2025 Data School Program! We are thrilled to recognize your commitment and excellence throughout the program.

We are pleased to provide you with both your Certificate of Attendance and Certificate of Proficiency, which recognize your full participation and mastery of the program content. These certificates reflect your dedication to professional growth and the skills you've developed during the program.

Additionally, we've included your personalized feedback document, which provides insights into your performance metrics and areas where you've excelled, as well as recommendations for future growth.

Your certificates and feedback are attached to this email. Feel free to share your certificates on your professional profiles and with your network.

Thank you for your active participation and remarkable performance. We hope the knowledge and skills you've gained will contribute significantly to your professional journey.

Best regards,
The Data School Program Team`;

  sendEmailWithHtml(email, subject, htmlBody, plainBody, attachments);
}

/**
 * Send email for students receiving feedback only (no certificates)
 */
function sendFeedbackOnlyEmail(name, email, attachments) {
  const subject = "Your Data School Program Feedback";
  
  const htmlBody = `
  <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
    <p>Dear ${name},</p>
    
    <p>Thank you for your participation in the May 2025 Data School Program.</p>
    
    <p>We've prepared a personalized feedback document for you, which provides insights into your performance metrics throughout the program, as well as recommendations for future growth.</p>
    
    <p>Your feedback document is attached to this email. We hope you find the assessment helpful as you continue your professional development journey.</p>
    
    <p>We appreciate your engagement with the program and wish you success in your future endeavors.</p>
    
    <p>Best regards,<br>
    The Data School Program Team</p>
  </div>`;
  
  const plainBody = `Dear ${name},

Thank you for your participation in the May 2025 Data School Program.

We've prepared a personalized feedback document for you, which provides insights into your performance metrics throughout the program, as well as recommendations for future growth.

Your feedback document is attached to this email. We hope you find the assessment helpful as you continue your professional development journey.

We appreciate your engagement with the program and wish you success in your future endeavors.

Best regards,
The Data School Program Team`;

  sendEmailWithHtml(email, subject, htmlBody, plainBody, attachments);
}

/**
 * Send email with HTML formatting and attachments
 */
function sendEmailWithHtml(email, subject, htmlBody, plainBody, attachments) {
  try {
    GmailApp.sendEmail(
      email,
      subject,
      plainBody,
      {
        htmlBody: htmlBody,
        attachments: attachments
      }
    );
    Logger.log(`Email sent successfully to ${email}`);
  } catch (error) {
    Logger.log(`Error sending email to ${email}: ${error.toString()}`);
  }
}