/**
 * DISPATCH — Google Sheets backend
 * -------------------------------------------------
 * Paste this whole file into Extensions > Apps Script
 * in the Google Sheet you want to use as the database.
 * Then deploy it as a Web App (see setup steps below).
 */

function doGet(e) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();

  var tasks = data
    .filter(function(row) { return row[0]; }) // skip blank rows
    .map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    });

  return ContentService
    .createTextOutput(JSON.stringify(tasks))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = getSheet();

  try {
    var body = JSON.parse(e.postData.contents);

    if (body.action === 'sync') {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
      }
      var rows = body.tasks.map(function(t) {
        return [t.id, t.title, t.notes, t.status, t.created];
      });
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, 5).setValues(rows);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Gets the "Tasks" sheet, creating it with headers if it
 * doesn't exist yet.
 */
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Tasks');
  if (!sheet) {
    sheet = ss.insertSheet('Tasks');
    sheet.appendRow(['id', 'title', 'notes', 'status', 'created']);
  }
  return sheet;
}
