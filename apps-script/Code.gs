// Sprachpaar sync backend — bind this script to a Google Sheet (Extensions > Apps Script)
// and deploy it as a Web App. See README.md in this folder for the exact steps.
//
// Storage model: one row per completion key in a "Completions" tab, created automatically
// on first use. Key format matches the website exactly: "YYYY-MM-DD:taskId" for solo tasks,
// "YYYY-MM-DD:taskId:person" for together tasks (e.g. "2026-09-29:tue-exercise:berkay").

var SHEET_NAME = 'Completions';

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Key', 'Done', 'UpdatedAt']);
  }
  return sheet;
}

function doGet(e) {
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var completions = {};
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0];
    var done = data[i][1];
    if (key && done) completions[key] = true;
  }
  return ContentService.createTextOutput(JSON.stringify({ completions: completions }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var key = body.key;
  var done = !!body.done;
  if (!key) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'missing key' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) {
    sheet.appendRow([key, done, new Date()]);
  } else {
    sheet.getRange(rowIndex, 2).setValue(done);
    sheet.getRange(rowIndex, 3).setValue(new Date());
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
