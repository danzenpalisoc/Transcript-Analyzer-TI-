/**
 * ExportService.gs
 * The dashboard's "Export to Google Sheets" buttons.
 *
 *   exportAudits(view, auditRefs, title)   open to every signed-in analyst
 *
 * The page sends only a view name ('repeats', 'sales' or 'auditlog') and a list of Audit Refs. The rows are built HERE from the
 * sheets, the same way the dashboard builds them, so nobody can make the app write their own text into a spreadsheet under the
 * app owner's name. Cells that start with = + - or @ are stored as text (never as a formula). The file goes into the app owner's
 * folder "NH Analyzer Exports" and is shared, as viewer, only with the person who asked. (The old exportToSpreadsheet took the
 * rows from the browser and shared the file with the whole domain.)
 */

var EXPORT_FOLDER = 'NH Analyzer Exports';
var EXPORT_MAX_ROWS = 5000;
var EXPORT_PER_HOUR = 10;                 // per person
var EXPORT_TRIES_PER_HOUR = 60;           // per person: every try, even one that finds nothing, because each one reads the sheets
var EXPORT_CELL_MAX = 49000;              // a Google Sheets cell holds up to 50,000 characters
var EXPORT_VIEWS = { repeats: 'Repeats', sales: 'Sales', auditlog: 'Audit Log' };

// The dashboard's merged rows (index.html dbHandleData / _dbMergeRow): every Dashboard_Data column, with the Audit_Log values
// laid over the top, and an audit that is only in Dashboard_Data kept. tests/export-test.js compares it with the page.
function buildDashboardRows_(logRows, dashRows) {
  var dashMap = {};
  (dashRows || []).forEach(function(r) { var ref = String(r['Audit Ref'] || '').trim(); if (ref) dashMap[ref] = r; });
  var logFirst = ['Interaction ID', 'SAP ID', 'Team Leader', 'Operations Manager', 'Line of Business', 'Locale', 'VTID', 'Observer Name', 'Direction',
                  'Analysis Type', 'Repeat Projection %', 'Issue Resolved', 'Transfer Occurred', 'Email Status'];
  var dashOnly = ['Department', 'Call Reason', 'Call Summary', 'Overall Opportunities', 'Critical Flags', 'Call Driver', 'RCA Category', 'RCA Sub-Parameter',
                  'Product Opportunity', 'PDF Email Link'];
  var merge = function(dash, over) {
    var row = {};
    Object.keys(dash).forEach(function(k) { row[k] = dash[k]; });
    Object.keys(over).forEach(function(k) { row[k] = over[k]; });
    return row;
  };
  var out = [], haveRef = {};
  (logRows || []).forEach(function(log) {
    var ref = String(log['Audit Ref'] || '').trim();
    haveRef[ref] = true;
    var dash = dashMap[ref] || {};
    var o = { 'Audit Ref': ref };
    o['Timestamp'] = log['Submitted At'] || dash['Timestamp'] || '';
    o['Team Member'] = log['Team Member'] || ((dash['Team Member'] && dash['Team Member'].length <= 80) ? dash['Team Member'] : '') || '';
    logFirst.forEach(function(k) { o[k] = log[k] || dash[k] || ''; });
    o['Transcript Duration'] = log['Duration'] || dash['Transcript Duration'] || '';
    dashOnly.forEach(function(k) { o[k] = dash[k] || ''; });
    o['SMART Recommendation'] = dash['SMART Recommendation'] || dash['Recommendations'] || '';
    o['Top SMART Opportunity'] = dash['Top SMART Opportunity'] || dash['Top Opportunity'] || '';
    out.push(merge(dash, o));
  });
  Object.keys(dashMap).forEach(function(ref) {
    if (haveRef[ref]) return;
    var d = dashMap[ref];
    out.push(merge(d, { 'Audit Ref': ref, 'Timestamp': d['Timestamp'] || '', 'Team Member': (d['Team Member'] && d['Team Member'].length <= 80) ? d['Team Member'] : '' }));
  });
  return out;
}

// The rows of one view for the given refs: { rows, headers }.
function _exRowsFor_(view, refSet, logRows, dashRows) {
  var rows;
  if (view === 'auditlog') {
    rows = (logRows || []).filter(function(r) { return refSet[String(r['Audit Ref'] || '').trim()]; });
  } else {
    var word = view === 'repeats' ? 'Repeats' : 'Sales';
    rows = buildDashboardRows_(logRows, dashRows).filter(function(r) {
      return refSet[r['Audit Ref']] && String(r['Analysis Type'] || '').indexOf(word) !== -1;
    });
  }
  var headers = [], seen = {};
  rows.forEach(function(r) { Object.keys(r).forEach(function(k) { if (!seen[k] && k.charAt(0) !== '_') { seen[k] = true; headers.push(k); } }); });
  return { rows: rows, headers: headers };
}

// A cell value as it is written to the sheet: always text, a leading = + - @ (or a control character that starts a formula in
// some spreadsheet programs) is protected with an apostrophe, a time-like text is kept as text, and an over-long text is cut.
function _exCell_(v) {
  var s = String(v === null || v === undefined ? '' : v);
  if (s.length > EXPORT_CELL_MAX) s = s.substring(0, EXPORT_CELL_MAX);
  if (/^[=+\-@\t\r]/.test(s)) return "'" + s;
  if (typeof _DURATION_SHAPE !== 'undefined' && _DURATION_SHAPE.test(s)) return "'" + s;
  return s;
}

// Validate what the browser sent. Returns { ok, view, refs, error }.
function _exValidate_(view, auditRefs) {
  if (!EXPORT_VIEWS.hasOwnProperty(view)) return { ok: false, error: 'Unknown export.' };
  if (!Array.isArray(auditRefs) || !auditRefs.length) return { ok: false, error: 'No records to export.' };
  if (auditRefs.length > EXPORT_MAX_ROWS) return { ok: false, error: 'Too many records (' + auditRefs.length + '). Narrow the filters and try again.' };
  var seen = {}, refs = [];
  for (var i = 0; i < auditRefs.length; i++) {
    var r = typeof auditRefs[i] === 'string' ? auditRefs[i].trim() : auditRefs[i];
    if (!isValidAuditRef_(r)) return { ok: false, error: 'The selection contains something that is not an audit reference.' };
    if (!seen[r]) { seen[r] = true; refs.push(r); }
  }
  return { ok: true, view: view, refs: refs };
}

function _exFolder_() {
  return _ownedFolder_(EXPORT_FOLDER);
}

function _exRateLimitTake_(email) {
  var cache = CacheService.getScriptCache(), key = 'ex_rl_' + Utilities.base64Encode(email).substring(0, 60);
  var n = parseInt(cache.get(key) || '0', 10) || 0;
  if (n >= EXPORT_PER_HOUR) return false;
  cache.put(key, String(n + 1), 3600);
  return true;
}

function _exFriendlyError_(e) {
  var raw = String((e && e.message) || e || '');
  Logger.log('exportAudits error: ' + raw);
  var msg = /too many times|rate limit|quota|limit exceeded/i.test(raw)
    ? 'Google is busy right now. Please wait a few minutes and try again.'
    : 'The export could not be made. Please try again, and tell the QA admin team if it keeps failing.';
  var admin = false;
  try { admin = isAdminCaller_(); } catch (x) {}
  return admin ? msg + ' (Details for admins: ' + raw.substring(0, 300) + ')' : msg;
}

// Returns { success, url, title, rowCount } or { success: false, error }.
function exportAudits(view, auditRefs, title) {
  var email = '';
  try { email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) {}
  if (!email) return { success: false, error: 'Could not tell who you are, so the export cannot be shared with you. Please sign in again.' };
  var v = _exValidate_(view, auditRefs);
  if (!v.ok) return { success: false, error: v.error };

  // Counted BEFORE the lock and the reads, so a try that finds nothing cannot keep the lock and the sheets busy for everybody.
  if (!_hourlyTake_('ex_try_', email, EXPORT_TRIES_PER_HOUR)) return { success: false, error: 'Too many export requests in the last hour. Please try again later.' };
  var lock = LockService.getScriptLock(), locked = false;
  try {
    if (!lock.tryLock(2000)) return { success: false, error: 'Another export is being made. Please try again in a minute.' };
    locked = true;
    var refSet = {};
    v.refs.forEach(function(r) { refSet[r] = true; });
    var data = _exRowsFor_(v.view, refSet, readAuditLog_(), v.view === 'auditlog' ? [] : readDashboardSheet_());
    if (!data.rows.length) return { success: false, error: 'None of the selected records is in this view, so there is nothing to export.' };
    if (!_exRateLimitTake_(email)) return { success: false, error: 'You have already made ' + EXPORT_PER_HOUR + ' exports in the last hour. Please try again later.' };

    var name = String(title || '').replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim().substring(0, 100) || ('NH Audit Export - ' + EXPORT_VIEWS[v.view]);
    var ss = SpreadsheetApp.create(name + ' - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
    var sheet = ss.getActiveSheet();
    sheet.getRange(1, 1, 1, data.headers.length).setValues([data.headers]).setFontWeight('bold').setBackground('#4B286D').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    var values = data.rows.map(function(r) { return data.headers.map(function(h) { return _exCell_(r[h]); }); });
    sheet.getRange(2, 1, values.length, data.headers.length).setValues(values);
    try { sheet.autoResizeColumns(1, data.headers.length); } catch (re) {}

    var file = DriveApp.getFileById(ss.getId());
    _exFolder_().addFile(file);
    try { DriveApp.getRootFolder().removeFile(file); } catch (me) { Logger.log('export: could not remove from root: ' + me); }
    file.addViewer(email);
    Logger.log('exportAudits: ' + data.rows.length + ' rows (' + v.view + ') for ' + email + ': ' + ss.getUrl());
    return { success: true, url: ss.getUrl(), title: ss.getName(), rowCount: data.rows.length };
  } catch (e) {
    return { success: false, error: _exFriendlyError_(e) };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (x) {} }
  }
}
