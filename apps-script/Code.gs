// Sprachpaar sync + auth backend — bind this script to a Google Sheet
// (Extensions > Apps Script) and deploy it as a Web App. See README.md for setup steps.
//
// Sheets used (all created automatically on first use):
//   Completions — Key | Done | UpdatedAt        (task checkboxes)
//   LoginCodes  — Email | Code | ExpiresAt       (one-time sign-in codes)
//   Sessions    — Token | Email | Person | CreatedAt
//   BoardItems  — Id | Side | Type | Content | Color | Font | X | Y | Rot | AddedBy | UpdatedAt
//
// Only these two people can sign in and write anything. Change the emails here if needed.
var ALLOWED = {
  'kemalberkayelcik@hotmail.com': 'berkay',
  'yagmurkaran6@gmail.com': 'yagmur',
};

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, headerRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headerRow);
  }
  return sheet;
}

function completionsSheet_() { return sheet_('Completions', ['Key', 'Done', 'UpdatedAt']); }
function loginCodesSheet_() { return sheet_('LoginCodes', ['Email', 'Code', 'ExpiresAt']); }
function sessionsSheet_() { return sheet_('Sessions', ['Token', 'Email', 'Person', 'CreatedAt']); }

// Seeded once, the first time this sheet is created: Yağmur's opening track
// (Dvořák, Serenade for Strings in E, Op. 22 — II. Tempo di Valse; CC BY-SA 4.0, Wikimedia Commons)
// and a shared Dutch van der Linde sticker either of you can move/rotate/resize/remove.
// Self-heals an older sheet missing the Scale column or the Dutch seed, instead of
// requiring anyone to delete real data — real rows (including their x/y/rot) are kept.
function boardSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('BoardItems');
  if (!sheet) {
    sheet = ss.insertSheet('BoardItems');
    sheet.appendRow(['Id', 'Side', 'Type', 'Content', 'Color', 'Font', 'X', 'Y', 'Rot', 'Scale', 'AddedBy', 'UpdatedAt']);
    seedBoardDefaults_(sheet);
    return sheet;
  }
  migrateBoardSheet_(sheet);
  return sheet;
}

function seedBoardDefaults_(sheet) {
  sheet.appendRow([
    Utilities.getUuid(), 'right', 'audio',
    'https://upload.wikimedia.org/wikipedia/commons/c/c2/Dvorak_String_Serenade_II_Tempo_di_Valse.ogg',
    '', '', 50, 50, 0, 1, 'yagmur', new Date(),
  ]);
  sheet.appendRow([
    Utilities.getUuid(), 'shared', 'image', 'plan.jpg',
    '', '', 15, 10, 0, 1, 'berkay', new Date(),
  ]);
}

function migrateBoardSheet_(sheet) {
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (header.indexOf('Scale') === -1) {
    var addedByCol = header.indexOf('AddedBy') + 1; // 1-based
    sheet.insertColumnBefore(addedByCol);
    sheet.getRange(1, addedByCol).setValue('Scale');
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var defaults = [];
      for (var i = 0; i < lastRow - 1; i++) defaults.push([1]);
      sheet.getRange(2, addedByCol, lastRow - 1, 1).setValues(defaults);
    }
  }
  var data = sheet.getDataRange().getValues();
  var hasShared = false;
  for (var j = 1; j < data.length; j++) {
    if (data[j][1] === 'shared') { hasShared = true; break; }
  }
  if (!hasShared) {
    sheet.appendRow([
      Utilities.getUuid(), 'shared', 'image', 'plan.jpg',
      '', '', 15, 10, 0, 1, 'berkay', new Date(),
    ]);
  }
}

function personForToken_(token) {
  if (!token) return null;
  var data = sessionsSheet_().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === token) return data[i][2];
  }
  return null;
}

function sideForPerson_(person) { return person === 'berkay' ? 'left' : 'right'; }

// 'shared' items (like the Dutch sticker) can be moved/edited by either signed-in person.
function canEditItemSide_(itemSide, person) { return itemSide === 'shared' || itemSide === sideForPerson_(person); }

// ---------- reads ----------

function doGet(e) {
  var action = (e.parameter && e.parameter.action) || 'completions';

  if (action === 'board') {
    var data = boardSheet_().getDataRange().getValues();
    var items = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      items.push({
        id: r[0], side: r[1], type: r[2], content: r[3],
        color: r[4], font: r[5], x: r[6], y: r[7], rot: r[8], scale: r[9] || 1,
      });
    }
    return jsonOut_({ items: items });
  }

  var cdata = completionsSheet_().getDataRange().getValues();
  var completions = {};
  for (var j = 1; j < cdata.length; j++) {
    if (cdata[j][0] && cdata[j][1]) completions[cdata[j][0]] = true;
  }
  return jsonOut_({ completions: completions });
}

// ---------- writes ----------

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var action = body.action;

  if (action === 'request-code') return requestCode_(body);
  if (action === 'verify-code') return verifyCode_(body);
  if (action === 'toggle') return toggleCompletion_(body);
  if (action === 'add-item') return addItem_(body);
  if (action === 'update-item') return updateItem_(body);
  if (action === 'remove-item') return removeItem_(body);

  return jsonOut_({ ok: false, error: 'unknown_action' });
}

var CODE_COOLDOWN_MS = 60 * 1000; // don't let one email trigger a new code more than once a minute

function requestCode_(body) {
  var email = String(body.email || '').toLowerCase().trim();
  if (!ALLOWED[email]) return jsonOut_({ ok: false, error: 'not_allowed' });
  var sheet = loginCodesSheet_();
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === email) {
      var requestedAt = new Date(rows[i][2]).getTime() - 10 * 60 * 1000; // ExpiresAt was set 10 min after the request
      if (Date.now() - requestedAt < CODE_COOLDOWN_MS) return jsonOut_({ ok: false, error: 'rate_limited' });
      break;
    }
  }
  var code = String(Math.floor(1000 + Math.random() * 9000));
  var expires = new Date(Date.now() + 10 * 60 * 1000);
  sheet.appendRow([email, code, expires]);
  MailApp.sendEmail(email, 'Your Sprachpaar sign-in code', 'Your code is: ' + code + '\n\nIt expires in 10 minutes.');
  return jsonOut_({ ok: true });
}

function verifyCode_(body) {
  var email = String(body.email || '').toLowerCase().trim();
  var code = String(body.code || '').trim();
  if (!ALLOWED[email]) return jsonOut_({ ok: false, error: 'not_allowed' });
  var sheet = loginCodesSheet_();
  var rows = sheet.getDataRange().getValues();
  var valid = false;
  for (var i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === email && String(rows[i][1]) === code && new Date(rows[i][2]) > new Date()) {
      valid = true;
      break;
    }
  }
  if (!valid) return jsonOut_({ ok: false, error: 'invalid_code' });
  var token = Utilities.getUuid();
  sessionsSheet_().appendRow([token, email, ALLOWED[email], new Date()]);
  return jsonOut_({ ok: true, token: token, person: ALLOWED[email] });
}

function toggleCompletion_(body) {
  var person = personForToken_(body.token);
  if (!person) return jsonOut_({ ok: false, error: 'not_authenticated' });
  var key = body.key;
  if (!key) return jsonOut_({ ok: false, error: 'missing_key' });
  var done = !!body.done;
  var sheet = completionsSheet_();
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) sheet.appendRow([key, done, new Date()]);
  else { sheet.getRange(rowIndex, 2).setValue(done); sheet.getRange(rowIndex, 3).setValue(new Date()); }
  return jsonOut_({ ok: true });
}

function addItem_(body) {
  var person = personForToken_(body.token);
  if (!person) return jsonOut_({ ok: false, error: 'not_authenticated' });
  var type = (body.type === 'text' || body.type === 'audio') ? body.type : 'image';
  var id = Utilities.getUuid();
  var side = sideForPerson_(person);
  boardSheet_().appendRow([
    id, side, type, body.content || '', body.color || '#F2DBA8', body.font || 'Public Sans',
    body.x != null ? body.x : 50, body.y != null ? body.y : 50, body.rot != null ? body.rot : 0,
    body.scale != null ? body.scale : 1, person, new Date(),
  ]);
  return jsonOut_({ ok: true, id: id, side: side });
}

function findItemRow_(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) return i + 1;
  }
  return -1;
}

function updateItem_(body) {
  var person = personForToken_(body.token);
  if (!person) return jsonOut_({ ok: false, error: 'not_authenticated' });
  var sheet = boardSheet_();
  var rowIndex = findItemRow_(sheet, body.id);
  if (rowIndex === -1) return jsonOut_({ ok: false, error: 'not_found' });
  if (!canEditItemSide_(sheet.getRange(rowIndex, 2).getValue(), person)) {
    return jsonOut_({ ok: false, error: 'not_your_side' });
  }
  if (body.content != null) sheet.getRange(rowIndex, 4).setValue(body.content);
  if (body.color != null) sheet.getRange(rowIndex, 5).setValue(body.color);
  if (body.font != null) sheet.getRange(rowIndex, 6).setValue(body.font);
  if (body.x != null) sheet.getRange(rowIndex, 7).setValue(body.x);
  if (body.y != null) sheet.getRange(rowIndex, 8).setValue(body.y);
  if (body.rot != null) sheet.getRange(rowIndex, 9).setValue(body.rot);
  if (body.scale != null) sheet.getRange(rowIndex, 10).setValue(body.scale);
  sheet.getRange(rowIndex, 12).setValue(new Date());
  return jsonOut_({ ok: true });
}

function removeItem_(body) {
  var person = personForToken_(body.token);
  if (!person) return jsonOut_({ ok: false, error: 'not_authenticated' });
  var sheet = boardSheet_();
  var rowIndex = findItemRow_(sheet, body.id);
  if (rowIndex === -1) return jsonOut_({ ok: false, error: 'not_found' });
  if (!canEditItemSide_(sheet.getRange(rowIndex, 2).getValue(), person)) {
    return jsonOut_({ ok: false, error: 'not_your_side' });
  }
  sheet.deleteRow(rowIndex);
  return jsonOut_({ ok: true });
}
