// ============================================================
//  65 BOYS — Google Apps Script Backend
//  Saves registration data to Google Sheets (Google Drive)
//
//  HOW TO DEPLOY:
//  1. Go to https://script.google.com/
//  2. Click "New Project"
//  3. Paste this entire file and save (Ctrl+S)
//  4. Click "Deploy" → "New Deployment"
//  5. Type: Web app | Execute as: Me | Who has access: Anyone
//  6. Click "Deploy" → copy the Web App URL
//  7. Paste that URL into index.html where it says:
//     const SCRIPT_URL = "YOUR_GOOGLE_APPS_SCRIPT_URL_HERE";
// ============================================================

const SHEET_NAME = "65 BOYS Registrations";

function doGet(e) {
  try {
    const name  = e.parameter.name  || "";
    const email = e.parameter.email || "";
    const phone = e.parameter.phone || "";

    if (!name && !email && !phone) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok", message: "65 BOYS backend is live!" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    saveToSheet(name, email, phone);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", message: "Registered!" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveToSheet(name, email, phone) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  // Create sheet + headers if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ["#", "Timestamp", "Full Name", "Email Address", "Phone Number"];
    const headerRow = sheet.getRange(1, 1, 1, headers.length);
    headerRow.setValues([headers]);

    // Style the header row
    headerRow.setBackground("#ff6400");
    headerRow.setFontColor("#ffffff");
    headerRow.setFontWeight("bold");
    headerRow.setFontSize(12);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 250);
    sheet.setColumnWidth(5, 180);
  }

  // Get next row number
  const lastRow = sheet.getLastRow();
  const rowNum  = lastRow; // row 1 = header, so row 2 = entry #1

  // Append the new registration
  sheet.appendRow([
    rowNum,                                               // #
    new Date().toLocaleString("en-IN", {                  // Timestamp (IST)
      timeZone: "Asia/Kolkata",
      dateStyle: "medium",
      timeStyle: "short"
    }),
    name,                                                 // Full Name
    email,                                                // Email
    phone                                                 // Phone
  ]);

  // Alternate row coloring for readability
  const newRow = sheet.getLastRow();
  if (newRow % 2 === 0) {
    sheet.getRange(newRow, 1, 1, 5).setBackground("#fff8f0");
  }
}
