/**
 * ==============================================================================
 * CREATOR CRAFT MEDIA — VIJAYAWADA WORKSHOP REGISTRATION SPREADSHEET BACKEND
 * ==============================================================================
 *
 * HOW TO SETUP IN 2 MINUTES:
 * 1. Open Google Sheets (https://sheets.google.com) and create a New Spreadsheet.
 *    Name it: "CCM Vijayawada Workshop Registrations 2026"
 *
 * 2. In Google Sheets menu, click:
 *    Extensions > Apps Script
 *
 * 3. Delete any existing code in the editor, and PASTE THIS ENTIRE FILE.
 *
 * 4. Click "Deploy" (blue button at top right) > "New deployment".
 *    - Click the Gear icon next to "Select type" and select: "Web app"
 *    - Description: "Vijayawada Workshop Registration API with PDF Passes"
 *    - Execute as: "Me" (your email)
 *    - Who has access: "Anyone"  <--- (IMPORTANT! Must select "Anyone")
 *
 * 5. Click "Deploy", approve permissions if prompted, and COPY the "Web app URL"
 *    (It looks like: https://script.google.com/macros/s/AKfycb.../exec)
 *
 * 6. Deployed Web App URL (Vijayawada 2026):
 *    https://script.google.com/macros/s/AKfycbxwkyuKv3EDjv6prDnYHcawHEgMyA9gNpyrLTmn0iYzbDH-G2aDw1ULcKz5e_ewJTibxw/exec
 *
 * THAT'S IT! All submissions, UTR numbers, payment screenshots, and WORKSHOP PDF PASSES
 * will instantly populate your Google Sheet in real-time with clickable download links!
 * ==============================================================================
 */

// ─── CONFIGURATION ─────────────────────────────────────────────────────────
var CONFIG = {
  WORKSHOP_NAME: "CCM Vijayawada Workshop 2026",
  CITY:          "Vijayawada",
  REG_PREFIX:    "CCM-VJA-",      // Registration ID prefix for Vijayawada
  DATES:         "Oct 24-25, 2026",
  HEADER_BG:     "#0F172A",
  HEADER_FG:     "#F8FAFC",
  PDF_FOLDER:    "CCM Vijayawada Workshop Passes 2026" // Dedicated Google Drive folder for PDF passes
};
// ───────────────────────────────────────────────────────────────────────────

/**
 * Helper to get or create a dedicated Google Drive folder for PDF passes
 */
function getPassesFolder() {
  try {
    var folders = DriveApp.getFoldersByName(CONFIG.PDF_FOLDER);
    if (folders.hasNext()) {
      return folders.next();
    }
    return DriveApp.createFolder(CONFIG.PDF_FOLDER);
  } catch (e) {
    return DriveApp.getRootFolder();
  }
}

/**
 * Handle POST request from workshop registration website
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    var headers = [
      "Timestamp",
      "Registration ID",
      "Full Name",
      "Mobile",
      "Email",
      "T-Shirt Size",
      "Category",
      "Knowledge Level",
      "Location / City",
      "Freelancer Interest",
      "UTR / Transaction ID",
      "Workshop Pass (PDF) — Download Link",
      "Receipt 1 — Filename",
      "Receipt 1 — Drive Link",
      "Receipt 2 — Filename",
      "Receipt 2 — Drive Link",
    ];

    // Auto-create headers if the sheet is brand new/empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);

      // Format header row (Dark luxury style)
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground(CONFIG.HEADER_BG);
      headerRange.setFontColor(CONFIG.HEADER_FG);
      headerRange.setFontWeight("bold");
      headerRange.setFontSize(11);

      // Auto-resize columns for readability
      sheet.setFrozenRows(1);
      for (var i = 1; i <= headers.length; i++) {
        sheet.setColumnWidth(i, i === 12 ? 240 : 200);
      }

      sheet.setName(CONFIG.WORKSHOP_NAME);
    } else {
      // Auto-upgrade existing sheet if "Workshop Pass" column is missing
      var firstRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 16)).getValues()[0];
      var hasPdfCol = firstRow.some(function (h) {
        return String(h).indexOf("Pass (PDF)") > -1 || String(h).indexOf("Workshop Pass") > -1;
      });

      if (!hasPdfCol && sheet.getLastColumn() >= 11) {
        sheet.insertColumnAfter(11);
        var passHeaderCell = sheet.getRange(1, 12);
        passHeaderCell.setValue("Workshop Pass (PDF) — Download Link");
        passHeaderCell.setBackground(CONFIG.HEADER_BG);
        passHeaderCell.setFontColor(CONFIG.HEADER_FG);
        passHeaderCell.setFontWeight("bold");
        passHeaderCell.setFontSize(11);
        sheet.setColumnWidth(12, 240);
      }
    }

    // ── Parse POST body ───────────────────────────────────────────────────
    var data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      data = e.parameter;
    }

    var regId = data.regId || (CONFIG.REG_PREFIX + Math.floor(10000 + Math.random() * 90000));
    var folder = getPassesFolder();

    // ── 1. Save Workshop Pass PDF to Google Drive ─────────────────────────
    var pdfDriveUrl = "Not uploaded";
    var pdfName     = data.pdfFilename || (regId + "_Workshop_Pass.pdf");
    var pdfBase64   = data.pdfBase64;

    if (pdfBase64 && pdfBase64 !== "None" && pdfBase64.length > 50) {
      try {
        var cleanPdfBase64 = pdfBase64.indexOf(",") > -1 ? pdfBase64.split(",")[1] : pdfBase64;
        var pdfBytes       = Utilities.base64Decode(cleanPdfBase64);
        var pdfBlob        = Utilities.newBlob(pdfBytes, "application/pdf", pdfName);
        var pdfFile        = folder.createFile(pdfBlob);
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        pdfDriveUrl        = pdfFile.getUrl();
      } catch (pdfErr) {
        pdfDriveUrl = "Error saving PDF: " + pdfErr.toString();
      }
    } else if (data.step === "Step 1 - Lead Captured" || data.paymentStatus === "Payment Pending") {
      pdfDriveUrl = "Pending (Step 2 Payment)";
    } else {
      // Fallback: Generate the pass PDF directly inside Google Apps Script!
      try {
        var fallbackBlob = createServerSidePassBlob(data, regId);
        var fallbackFile = folder.createFile(fallbackBlob);
        fallbackFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        pdfDriveUrl      = fallbackFile.getUrl();
      } catch (genErr) {
        pdfDriveUrl = "Not generated: " + genErr.toString();
      }
    }

    // ── 2. Save Receipt Screenshot 1 to Google Drive ───────────────────────
    var fileUrl1       = "Not uploaded";
    var receipt1Name   = data.receiptFilename || data.receipt1Filename || "None";
    var receipt1Base64 = data.receiptBase64   || data.receipt1Base64;

    if (receipt1Base64 && receipt1Name && receipt1Name !== "None") {
      try {
        var mime1  = receipt1Base64.substring(5, receipt1Base64.indexOf(";"));
        var bytes1 = Utilities.base64Decode(receipt1Base64.split(",")[1]);
        var blob1  = Utilities.newBlob(bytes1, mime1, regId + "_Receipt1_" + receipt1Name);
        var file1  = folder.createFile(blob1);
        file1.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl1   = file1.getUrl();
      } catch (driveErr1) {
        fileUrl1 = "Error: " + driveErr1.toString();
      }
    }

    // ── 3. Save Receipt Screenshot 2 to Google Drive ───────────────────────
    var fileUrl2       = "Not uploaded";
    var receipt2Name   = data.receipt2Filename || "None";
    var receipt2Base64 = data.receipt2Base64;

    if (receipt2Base64 && receipt2Name && receipt2Name !== "None") {
      try {
        var mime2  = receipt2Base64.substring(5, receipt2Base64.indexOf(";"));
        var bytes2 = Utilities.base64Decode(receipt2Base64.split(",")[1]);
        var blob2  = Utilities.newBlob(bytes2, mime2, regId + "_Receipt2_" + receipt2Name);
        var file2  = folder.createFile(blob2);
        file2.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl2   = file2.getUrl();
      } catch (driveErr2) {
        fileUrl2 = "Error: " + driveErr2.toString();
      }
    }

    // ── 4. Clickable PDF Pass Download Link in Excel/Google Sheets ────────
    var pdfCellFormula = (pdfDriveUrl && pdfDriveUrl.indexOf("http") === 0)
      ? '=HYPERLINK("' + pdfDriveUrl + '", "📄 Download PDF Pass")'
      : pdfDriveUrl;

    // ── 5. Append or Update row in sheet ──────────────────────────────────
    var row = [
      new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" }),
      regId,
      data.fullName  || "",
      data.mobile    || "",
      data.email     || "",
      data.tshirt    || "",
      data.category  || "",
      data.knowledge || "",
      data.location  || CONFIG.CITY,
      data.freelance || "",
      data.utr       || "",
      pdfCellFormula,
      receipt1Name,
      fileUrl1,
      receipt2Name,
      fileUrl2,
    ];

    // Check if registration ID already exists (e.g. from Step 1)
    var totalRows = sheet.getLastRow();
    var existingRowIndex = -1;
    if (totalRows > 1) {
      var regIds = sheet.getRange(2, 2, totalRows - 1, 1).getValues();
      for (var r = 0; r < regIds.length; r++) {
        if (regIds[r][0] && String(regIds[r][0]).trim() === String(regId).trim()) {
          existingRowIndex = r + 2;
          break;
        }
      }
    }

    var targetRow = existingRowIndex > 0 ? existingRowIndex : (totalRows + 1);
    if (existingRowIndex > 0) {
      sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
      targetRow = sheet.getLastRow();
    }

    // Alternating luxury row styling for easy scanning
    var rowRange = sheet.getRange(targetRow, 1, 1, row.length);
    rowRange.setBackground(targetRow % 2 === 0 ? "#0B0F1A" : "#111827");
    rowRange.setFontColor("#E2E8F0");

    // Make the PDF pass cell highlighted for quick access
    var pdfCellRange = sheet.getRange(targetRow, 12);
    pdfCellRange.setFontWeight("bold");
    pdfCellRange.setFontColor("#38BDF8");

    return ContentService
      .createTextOutput(JSON.stringify({
        status:  "success",
        regId:   regId,
        pdfUrl:  pdfDriveUrl,
        message: "Registration recorded successfully for " + CONFIG.WORKSHOP_NAME,
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status:  "error",
        message: error.toString(),
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

/**
 * Native Google Apps Script PDF Generator (Fallback & Direct Sheet Generation)
 * Produces the identical official Creator Craft Media Workshop Pass
 */
function createServerSidePassBlob(data, regId) {
  var isPaid = (data.utr && String(data.utr).trim().length > 3) ||
               (data.receiptFilename && data.receiptFilename !== "None") ||
               (data.receipt2Filename && data.receipt2Filename !== "None") ||
               (data.paymentStatus && data.paymentStatus.indexOf("Success") > -1);

  var statusBadge = isPaid
    ? '<span style="background-color:#064e3b;color:#34d399;border:1px solid #059669;padding:3px 10px;border-radius:12px;font-weight:bold;font-size:11px;">&#10003; Payment Successful</span>'
    : '<span style="background-color:#78350f;color:#fde68a;border:1px solid #b45309;padding:3px 10px;border-radius:12px;font-weight:bold;font-size:11px;">Payment Pending</span>';

  var topBadge = isPaid
    ? '<span style="background-color:#064e3b;color:#34d399;border:1px solid #059669;padding:2px 8px;border-radius:10px;font-size:9px;font-weight:bold;">&#9679; CONFIRMED &amp; VERIFIED</span>'
    : '<span style="background-color:#78350f;color:#fde68a;border:1px solid #b45309;padding:2px 8px;border-radius:10px;font-size:9px;font-weight:bold;">&#9679; PAYMENT PENDING</span>';

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' +
    '@page{size:A4 portrait;margin:12mm;}' +
    'body{font-family:Helvetica,Arial,sans-serif;margin:0;padding:0;color:#0f172a;background:#ffffff;line-height:1.4;}' +
    'table{border-collapse:collapse;width:100%;}' +
    '.header-table{background-color:#070a11;border-top:4px solid #00d2ff;border-radius:8px;padding:16px 20px;margin-bottom:14px;}' +
    '.card{background-color:#f8fafc;border:1px solid #cbd5e1;border-radius:8px;margin-bottom:14px;}' +
    '.meta-table td{padding:6px 10px;font-size:11px;border-bottom:1px solid #f1f5f9;}' +
    '.meta-table tr:nth-child(even){background-color:#f1f5f9;}' +
    '.curr-table td{vertical-align:top;width:50%;padding:0 6px;}' +
    '.module-card{background-color:#f8fafc;border-radius:8px;}' +
    '.day1-head{background-color:#eff6ff;border-bottom:2px solid #3b82f6;color:#1d4ed8;padding:8px 12px;font-weight:bold;font-size:11px;border-radius:8px 8px 0 0;}' +
    '.day2-head{background-color:#f5f3ff;border-bottom:2px solid #7c3aed;color:#6d28d9;padding:8px 12px;font-weight:bold;font-size:11px;border-radius:8px 8px 0 0;}' +
    '.module-body{padding:10px 14px;font-size:10.5px;color:#334155;line-height:1.6;}' +
    '.inst-box{background-color:#fefce8;border:1px solid #fef08a;border-radius:8px;padding:12px 16px;margin-bottom:14px;font-size:10px;color:#713f12;line-height:1.6;}' +
    '.footer-table{background-color:#070a11;border-top:2px solid #00d2ff;border-radius:8px;padding:14px 18px;color:#ffffff;}' +
    '</style></head><body>' +

    // Top Header with Official Logo & Badges
    '<table class="header-table">' +
      '<tr>' +
        '<td style="vertical-align:middle;width:55%;">' +
          '<img src="' + logoDataUri + '" width="210" style="display:block;margin-bottom:6px;" alt="Creator Craft Media" />' +
          '<div style="font-size:8px;letter-spacing:1.8px;color:#94a3b8;text-transform:uppercase;">CINEMATIC SHOOT &bull; AI FILM-MAKING &bull; META ADS</div>' +
        '</td>' +
        '<td style="text-align:right;vertical-align:middle;width:45%;">' +
          '<div style="display:inline-block;background-color:#1e293b;color:#f59e0b;padding:3px 10px;border-radius:4px;font-weight:bold;font-size:8.5px;letter-spacing:1px;margin-bottom:4px;">OFFICIAL VIP DELEGATE PASS</div>' +
          '<div style="font-size:18px;font-weight:bold;color:#38bdf8;letter-spacing:1px;margin:2px 0;">' + regId + '</div>' +
          '<div style="font-size:10.5px;color:#cbd5e1;margin-bottom:4px;">Oct 24-25, 2026 &bull; Vijayawada, AP</div>' +
          '<div>' + topBadge + '</div>' +
        '</td>' +
      '</tr>' +
    '</table>' +

    // Delegate Details Card
    '<div style="font-size:11px;font-weight:bold;color:#0f172a;letter-spacing:0.5px;margin-bottom:6px;">DELEGATE CREDENTIAL &amp; RESERVATION DETAILS</div>' +
    '<div class="card" style="border-left:4px solid #00d2ff;overflow:hidden;">' +
      '<table class="meta-table">' +
        '<tr>' +
          '<td style="color:#64748b;width:20%;">Participant Name:</td><td style="font-weight:bold;color:#0f172a;font-size:12px;width:30%;">' + (data.fullName || data.name || 'Participant') + '</td>' +
          '<td style="color:#64748b;width:20%;">Mobile Number:</td><td style="font-weight:bold;color:#0f172a;font-size:11.5px;width:30%;">' + (data.mobile || '-') + '</td>' +
        '</tr>' +
        '<tr>' +
          '<td style="color:#64748b;">Email Address:</td><td style="color:#334155;">' + (data.email || 'Not provided') + '</td>' +
          '<td style="color:#64748b;">Category / Role:</td><td><span style="background-color:#eef2ff;color:#4338ca;padding:2px 8px;border-radius:4px;font-weight:bold;font-size:10.5px;">' + (data.category || 'Content Creator') + '</span></td>' +
        '</tr>' +
        '<tr>' +
          '<td style="color:#64748b;">Workshop Fee:</td><td style="font-weight:bold;color:#0f172a;">Rs. 3,999 <span style="font-size:9.5px;color:#16a34a;font-weight:bold;">(Lunch &amp; Snacks Included)</span></td>' +
          '<td style="color:#64748b;">Payment Status:</td><td>' + statusBadge + '</td>' +
        '</tr>' +
        '<tr>' +
          '<td style="color:#64748b;">UTR Reference:</td><td style="font-family:monospace;font-size:11px;color:#0284c7;font-weight:bold;">' + (data.utr || 'Screenshot Verified') + '</td>' +
          '<td style="color:#64748b;">Location / City:</td><td style="font-weight:bold;color:#0f172a;">' + (data.location || CONFIG.CITY) + '</td>' +
        '</tr>' +
      '</table>' +
    '</div>' +

    // Masterclass Modules
    '<div style="font-size:11px;font-weight:bold;color:#0f172a;letter-spacing:0.5px;margin-bottom:6px;">OFFICIAL WORKSHOP PRACTICAL MODULES</div>' +
    '<table class="curr-table" style="margin-bottom:14px;">' +
      '<tr>' +
        '<td>' +
          '<div class="module-card" style="border:1px solid #bfdbfe;">' +
            '<div class="day1-head">DAY 01 &bull; OCT 24: SHOOT &amp; PRO EDITING</div>' +
            '<div class="module-body">' +
              '&bull; Mobile &amp; Camera Framing, Lighting &amp; Angles<br>' +
              '&bull; Cinematic Camera Movements &amp; Transitions<br>' +
              '&bull; Hands-on Live Shoot Session in Vijayawada<br>' +
              '&bull; <strong>CapCut Pro Video Editing Masterclass</strong><br>' +
              '&bull; Sound Design, Trending Audio &amp; Viral Pacing' +
            '</div>' +
          '</div>' +
        '</td>' +
        '<td>' +
          '<div class="module-card" style="border:1px solid #ddd6fe;">' +
            '<div class="day2-head">DAY 02 &bull; OCT 25: AI FILMS &amp; MONETIZATION</div>' +
            '<div class="module-body">' +
              '&bull; Generative AI Video, Scripting &amp; Storyboards<br>' +
              '&bull; <strong>Google Flow, Higgsfield, Seedance &amp; Kling</strong><br>' +
              '&bull; AI Voice Cloning &amp; Lip-Sync Filmmaking<br>' +
              '&bull; Meta Ads Campaign Setup for Maximum Reach<br>' +
              '&bull; Client Acquisition, Pricing &amp; Freelance Blueprint' +
            '</div>' +
          '</div>' +
        '</td>' +
      '</tr>' +
    '</table>' +

    // Important Instructions
    '<div class="inst-box">' +
      '<div style="font-weight:bold;font-size:10.5px;color:#854d0e;margin-bottom:4px;">&#128204; IMPORTANT WORKSHOP INSTRUCTIONS FOR ATTENDEES:</div>' +
      '1. <strong>Reporting Time:</strong> Arrive at venue by 8:30 AM for registration badge and kit collection.<br>' +
      '2. <strong>Equipment:</strong> Bring smartphone with 15GB free storage and power bank (Laptop optional for Day 2).<br>' +
      '3. <strong>Food &amp; Refreshments:</strong> Delicious afternoon lunch &amp; evening high-tea snacks provided on both days.<br>' +
      '4. <strong>Certification:</strong> Official Creator Craft Media Workshop Certificate awarded at closing ceremony.' +
    '</div>' +

    // Luxury Security Footer
    '<table class="footer-table">' +
      '<tr>' +
        '<td style="vertical-align:middle;width:68%;">' +
          '<div style="font-size:11px;font-weight:bold;letter-spacing:1px;color:#ffffff;margin-bottom:2px;">CREATOR CRAFT MEDIA &bull; VIJAYAWADA DESK</div>' +
          '<div style="font-size:9.5px;color:#cbd5e1;margin-bottom:2px;">Helpline / WhatsApp: +91 9121796816 &bull; Instagram: @creator_craft_media</div>' +
          '<div style="font-size:8.5px;color:#94a3b8;">UPI Merchant ID: alphaassetfinserv@sbi &bull; Please keep this pass handy for venue check-in.</div>' +
        '</td>' +
        '<td style="text-align:right;vertical-align:middle;width:32%;">' +
          '<div style="display:inline-block;border:1.5px solid #00d2ff;padding:6px 12px;border-radius:4px;text-align:center;background-color:#0f172a;">' +
            '<div style="font-size:8px;font-weight:bold;color:#00d2ff;letter-spacing:1px;">OFFICIAL VERIFICATION</div>' +
            '<div style="font-size:11px;font-weight:bold;color:#ffffff;margin:1px 0;">VERIFIED VIP PASS</div>' +
            '<div style="font-size:8px;color:#94a3b8;">' + regId + '</div>' +
          '</div>' +
        '</td>' +
      '</tr>' +
    '</table>' +

    '</body></html>';

  return Utilities.newBlob(html, "text/html").getAs("application/pdf").setName(regId + "_Workshop_Pass.pdf");
}

/**
 * ── Google Sheets Custom Menu ──────────────────────────────────────────────
 * Enables organizers to generate or re-download PDF passes directly from the sheet!
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🎟️ CCM Workshop")
    .addItem("📄 Generate / Refresh PDF for Selected Row", "generatePassForSelectedRow")
    .addItem("🔄 Generate Missing PDFs for All Rows", "generatePassesForAllMissing")
    .addToUi();
}

/**
 * Generates and attaches PDF pass for the currently selected row in Google Sheets
 */
function generatePassForSelectedRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();

  if (row <= 1) {
    SpreadsheetApp.getUi().alert("Please select a participant row (Row 2 or below).");
    return;
  }

  var regId    = sheet.getRange(row, 2).getValue();
  var fullName = sheet.getRange(row, 3).getValue();
  var mobile   = sheet.getRange(row, 4).getValue();
  var email    = sheet.getRange(row, 5).getValue();
  var category = sheet.getRange(row, 7).getValue();
  var location = sheet.getRange(row, 9).getValue();
  var utr      = sheet.getRange(row, 11).getValue();

  var data = {
    regId: regId,
    fullName: fullName,
    mobile: mobile,
    email: email,
    category: category,
    location: location,
    utr: utr
  };

  var folder = getPassesFolder();
  var blob = createServerSidePassBlob(data, regId);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var pdfUrl = file.getUrl();
  var formula = '=HYPERLINK("' + pdfUrl + '", "📄 Download PDF Pass")';
  sheet.getRange(row, 12).setValue(formula);

  SpreadsheetApp.getUi().alert("Workshop Pass PDF Generated Successfully!\n\nLink: " + pdfUrl);
}

/**
 * Scans all rows and automatically creates missing PDF passes
 */
function generatePassesForAllMissing() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No participant rows found.");
    return;
  }

  var folder = getPassesFolder();
  var count = 0;

  for (var r = 2; r <= lastRow; r++) {
    var currentPdfVal = String(sheet.getRange(r, 12).getValue());
    if (!currentPdfVal || currentPdfVal.indexOf("http") === -1 && currentPdfVal.indexOf("HYPERLINK") === -1) {
      var regId    = sheet.getRange(r, 2).getValue() || (CONFIG.REG_PREFIX + Math.floor(10000 + Math.random() * 90000));
      var fullName = sheet.getRange(r, 3).getValue();
      var mobile   = sheet.getRange(r, 4).getValue();
      var email    = sheet.getRange(r, 5).getValue();
      var category = sheet.getRange(r, 7).getValue();
      var location = sheet.getRange(r, 9).getValue();
      var utr      = sheet.getRange(r, 11).getValue();

      var data = {
        regId: regId,
        fullName: fullName,
        mobile: mobile,
        email: email,
        category: category,
        location: location,
        utr: utr
      };

      try {
        var blob = createServerSidePassBlob(data, regId);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        var pdfUrl = file.getUrl();
        sheet.getRange(r, 12).setValue('=HYPERLINK("' + pdfUrl + '", "📄 Download PDF Pass")');
        count++;
      } catch (err) {}
    }
  }

  SpreadsheetApp.getUi().alert("Completed! Generated and linked " + count + " workshop PDF pass(es).");
}

/**
 * Health check & direct pass redirect GET endpoint
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.regId) {
    var regId = e.parameter.regId;
    var files = DriveApp.getFilesByName(regId + "_Workshop_Pass.pdf");
    if (files.hasNext()) {
      return HtmlService.createHtmlOutput(
        '<script>window.location.href="' + files.next().getUrl() + '";</script>' +
        '<p>Redirecting to Workshop Pass PDF for ' + regId + '...</p>'
      );
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      status:   "active",
      workshop: CONFIG.WORKSHOP_NAME,
      city:     CONFIG.CITY,
      dates:    CONFIG.DATES,
      message:  "CCM Vijayawada Workshop Registration API with PDF Pass Sync is live!",
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
