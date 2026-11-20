// Receives the booking form POST from index.js and appends a row to the sheet
// this script is attached to (open the sheet -> Extensions -> Apps Script).
function doPost(e) {
  const p = e.parameter || {};
  // Prefix with ' so phone numbers like +91... stay as text.
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].appendRow([
    new Date(),
    p.name || '',
    "'" + (p.phone || ''),
    p.condition || '',
    p.time || '',
    p.message || ''
  ]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
