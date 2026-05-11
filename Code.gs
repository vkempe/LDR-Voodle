/**
 * Voodle backend for "LDR Voodle" Google Sheet.
 *
 * Setup (one-time):
 *  1. Open the "LDR Voodle" Google Sheet.
 *  2. Extensions > Apps Script. Paste this file into Code.gs.
 *  3. Click Deploy > New deployment > Web app.
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  4. Copy the resulting Web app URL.
 *  5. Paste that URL into index.html as APPS_SCRIPT_URL.
 *
 * Each submission writes one row PER SLOT (long format) so you can
 * pivot, filter, or COUNTIF easily by poll title.
 */

const SHEET_NAME = "Responses";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet_();
    const submissionId = Utilities.getUuid();
    const serverTs = new Date();

    const rows = (data.responses || []).map(function(r) {
      return [
        serverTs,
        data.pollTitle || "",
        data.name || "",
        r.slot || "",
        r.value || "",
        data.comment || "",
        submissionId,
        data.submittedAt || ""
      ];
    });

    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
           .setValues(rows);
    }

    return ContentService
      .createTextOutput(JSON.stringify({status: "ok", rows: rows.length}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: "error", message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({status: "Voodle endpoint live"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Server Timestamp",
      "Poll Title",
      "Name",
      "Slot",
      "Response",
      "Comment",
      "Submission ID",
      "Client Timestamp"
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
