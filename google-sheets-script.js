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
 *    - Description: "Vijayawada Workshop Registration API"
 *    - Execute as: "Me" (your email)
 *    - Who has access: "Anyone"  <--- (IMPORTANT! Must select "Anyone")
 *
 * 5. Click "Deploy", approve permissions if prompted, and COPY the "Web app URL"
 *    (It looks like: https://script.google.com/macros/s/AKfycb.../exec)
 *
 * 6. Deployed Web App URL (Vijayawada 2026):
 *    https://script.google.com/macros/s/AKfycbxSZOtraln9Pl_nNvfjPiLdaaNsqhl_ZNiHUYfJLBZvLfRFPy1lXsQp-f09LDHt_wF_Ag/exec
 *
 * THAT'S IT! All submissions, UTR numbers, and receipt details will instantly
 * populate your Google Sheet in real-time!
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
};
// ───────────────────────────────────────────────────────────────────────────

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Auto-create headers if the sheet is empty
    if (sheet.getLastRow() === 0) {
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
        "Receipt 1 — Filename",
        "Receipt 1 — Drive Link",
        "Receipt 2 — Filename",
        "Receipt 2 — Drive Link",
      ];
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
        sheet.setColumnWidth(i, 200);
      }

      // Rename sheet tab to workshop name
      sheet.setName(CONFIG.WORKSHOP_NAME);
    }

    // ── Parse POST body ───────────────────────────────────────────────────
    var data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      data = e.parameter;
    }

    // ── Save Receipt Screenshot 1 to Google Drive ─────────────────────────
    var fileUrl1       = "Not uploaded";
    var receipt1Name   = data.receiptFilename || data.receipt1Filename || "None";
    var receipt1Base64 = data.receiptBase64   || data.receipt1Base64;

    if (receipt1Base64 && receipt1Name && receipt1Name !== "None") {
      try {
        var mime1  = receipt1Base64.substring(5, receipt1Base64.indexOf(";"));
        var bytes1 = Utilities.base64Decode(receipt1Base64.split(",")[1]);
        var blob1  = Utilities.newBlob(bytes1, mime1,
                       (data.regId || "REC") + "_Receipt1_" + receipt1Name);
        var file1  = DriveApp.createFile(blob1);
        file1.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl1 = file1.getUrl();
      } catch (driveErr1) {
        fileUrl1 = "Error: " + driveErr1.toString();
      }
    }

    // ── Save Receipt Screenshot 2 to Google Drive ─────────────────────────
    var fileUrl2       = "Not uploaded";
    var receipt2Name   = data.receipt2Filename || "None";
    var receipt2Base64 = data.receipt2Base64;

    if (receipt2Base64 && receipt2Name && receipt2Name !== "None") {
      try {
        var mime2  = receipt2Base64.substring(5, receipt2Base64.indexOf(";"));
        var bytes2 = Utilities.base64Decode(receipt2Base64.split(",")[1]);
        var blob2  = Utilities.newBlob(bytes2, mime2,
                       (data.regId || "REC") + "_Receipt2_" + receipt2Name);
        var file2  = DriveApp.createFile(blob2);
        file2.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl2 = file2.getUrl();
      } catch (driveErr2) {
        fileUrl2 = "Error: " + driveErr2.toString();
      }
    }

    // ── Generate registration ID if not provided ──────────────────────────
    var regId = data.regId || (CONFIG.REG_PREFIX + Math.floor(1000 + Math.random() * 9000));

    // ── Append row to sheet ───────────────────────────────────────────────
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
      receipt1Name,
      fileUrl1,
      receipt2Name,
      fileUrl2,
    ];

    sheet.appendRow(row);

    // ── Alternating row colors for easy scanning ──────────────────────────
    var lastRow  = sheet.getLastRow();
    var rowRange = sheet.getRange(lastRow, 1, 1, row.length);
    rowRange.setBackground(lastRow % 2 === 0 ? "#0B0F1A" : "#111827");
    rowRange.setFontColor("#E2E8F0");

    return ContentService
      .createTextOutput(JSON.stringify({
        status:  "success",
        regId:   regId,
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

// ── Health-check GET endpoint ─────────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status:   "active",
      workshop: CONFIG.WORKSHOP_NAME,
      city:     CONFIG.CITY,
      dates:    CONFIG.DATES,
      message:  "CCM Vijayawada Workshop Registration API is live!",
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
