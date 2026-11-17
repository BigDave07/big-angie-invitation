/**
 * RSVP → Google Sheet
 *
 * Setup (5 minutes):
 * 1. Create a new Google Sheet (sheets.new). Name it "Angie RSVPs".
 * 2. Extensions → Apps Script. Delete what's there, paste this whole file, Save.
 * 3. Deploy → New deployment → gear icon → "Web app".
 *      Execute as: Me      Who has access: Anyone
 * 4. Authorize when asked, then copy the "Web app URL" (ends in /exec).
 * 5. Paste it into index.html:  const SHEET_URL = 'https://script.google.com/macros/s/.../exec';
 *
 * Every RSVP becomes a new row. If you edit this script later,
 * use Deploy → Manage deployments → Edit → New version (same URL).
 */
function doPost(e) {
  const p = e.parameter || {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Received', 'Name', 'Attending', 'Party size', 'Note']);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date(), clean(p.name, 120), clean(p.attending, 3), clean(p.guests, 2), clean(p.message, 500)]);
  return ContentService.createTextOutput('ok');
}

// Trim length, and stop guests from sneaking spreadsheet formulas (=, +, -, @) into cells.
function clean(v, max) {
  const s = String(v || '').slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
