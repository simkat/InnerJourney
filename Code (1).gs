/**
 * Inner Journey — Practice Log backend  (v2)
 * Turns a Google Sheet into a tiny sync API for the tracker page.
 *
 * v2 CHANGES — you MUST redeploy after pasting this:
 *   Deploy ▸ Manage deployments ▸ Edit (pencil) ▸ Version: New version ▸ Deploy
 *   The /exec URL stays the same, so no change is needed in index.html.
 *
 * What's new: an optional "sheet" parameter on every action, so the Practice
 * Library and weekly assignments get their own tabs. Requests that omit it
 * default to "Log", so the previous version's calls keep working unchanged.
 *
 * SETUP (first time only):
 *  1. Create a Google Sheet. Extensions ▸ Apps Script.
 *  2. Delete any sample code, paste ALL of this file, Save.
 *  3. (Optional) set SECRET below and put the SAME value in index.html CONFIG.SECRET.
 *  4. Deploy ▸ New deployment ▸ Web app
 *       - Execute as: Me
 *       - Who has access: Anyone with the link
 *     Copy the /exec URL into index.html CONFIG.ENDPOINT.
 */

const SECRET = '';   // leave '' for no passphrase, or match CONFIG.SECRET

const SHEETS = {
  Log: ['id','createdAt','person','kind','date','phase',
        'lessonTitle','lessonUrl','formalDone','minutes','rating',
        'daytimeText','daytimeUrl','reflection','text','status',
        'virtueKey','virtueName',
        'practiceId','weekNumber','dayNumber','engaged','practised'],
  Practices: ['practiceId','title','core','purpose','dateIntroduced','originalWeek',
              'instructions','dailyLife'],
  PracticeAssignments: ['weekNumber','weekStart','weekEnd','practiceId',
                        'assignmentType','assignedAt']
};

/* The key column used to find an existing row, per sheet */
const KEY_COL = { Log: 'id', Practices: 'practiceId', PracticeAssignments: 'weekNumber' };

function doGet(e) {
  var which = (e && e.parameter && e.parameter.sheet) || null;
  if (which && SHEETS[which]) {
    return json_({ ok: true, rows: getRows_(which) });
  }
  // default: return everything, with `rows` kept for backwards compatibility
  return json_({
    ok: true,
    rows: getRows_('Log'),
    practices: getRows_('Practices'),
    practiceAssignments: getRows_('PracticeAssignments')
  });
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  if (SECRET && body.secret !== SECRET) {
    return json_({ ok: false, error: 'unauthorised' });
  }

  var name = body.sheet && SHEETS[body.sheet] ? body.sheet : 'Log';
  var headers = SHEETS[name];
  var keyCol = KEY_COL[name];
  var sh = sheet_(name, headers);
  var action = body.action || 'add';

  if (action === 'add') {
    var row = body.row || {};
    // upsert on the key column so re-syncing never duplicates
    var existing = findRow_(sh, headers, keyCol, row[keyCol]);
    if (existing > 0) {
      headers.forEach(function (h, i) {
        if (row[h] !== undefined) sh.getRange(existing, i + 1).setValue(serialise_(row[h]));
      });
      return json_({ ok: true, updated: true, key: row[keyCol] });
    }
    sh.appendRow(headers.map(function (h) {
      return row[h] !== undefined && row[h] !== null ? serialise_(row[h]) : '';
    }));
    return json_({ ok: true, key: row[keyCol] });
  }

  if (action === 'update') {
    var key = body.id !== undefined ? body.id : (body.key !== undefined ? body.key : null);
    var r = findRow_(sh, headers, keyCol, key);
    if (r > 0) {
      var fields = body.fields || {};
      headers.forEach(function (h, i) {
        if (fields[h] !== undefined) sh.getRange(r, i + 1).setValue(serialise_(fields[h]));
      });
    }
    return json_({ ok: true });
  }

  if (action === 'delete') {
    var dkey = body.id !== undefined ? body.id : (body.key !== undefined ? body.key : null);
    var d = findRow_(sh, headers, keyCol, dkey);
    if (d > 0) sh.deleteRow(d);
    return json_({ ok: true });
  }

  return json_({ ok: false, error: 'unknown action' });
}

/* ---------- helpers ---------- */
function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}
function getRows_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues();
  var head = values[0];
  return values.slice(1).map(function (r) {
    var o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
}
function findRow_(sh, headers, keyCol, key) {
  if (key === null || key === undefined || key === '') return -1;
  if (sh.getLastRow() < 2) return -1;
  var col = headers.indexOf(keyCol) + 1;
  if (col < 1) return -1;
  var vals = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(key)) return i + 2;
  }
  return -1;
}
/* arrays are stored as newline-joined text so the sheet stays readable */
function serialise_(v) {
  if (Object.prototype.toString.call(v) === '[object Array]') return v.join('\n');
  return v;
}
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
