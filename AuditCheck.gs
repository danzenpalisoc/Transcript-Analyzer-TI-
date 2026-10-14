/**
 * AuditCheck.gs
 *
 * verifyLatestAudit(): a read-only checklist that proves ONE audit landed everywhere
 * it should. Run it from the Apps Script editor (choose it in the function list, press
 * Run, open the Execution log) after an audit, instead of testing each place by hand.
 *
 * It is admin-only and writes nothing: no row, no cache entry. It reads the Cache
 * SHEET directly rather than the 6-hour script cache, because the script-cache copy
 * would hide a damaged sheet copy that the evaluation page shows later.
 *
 * What it cannot see is the screen: whether the email reached an inbox, or how the
 * report looks. It checks the data behind both.
 */

// All rows of a sheet whose first column (Audit Ref) is `ref`, as {header: value}.
// Reads the ref column first and then only the matching rows.
function _auditRowsForRef_(ss, sheetName, ref) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { missing: true, rows: [] };
  var last = sheet.getLastRow();
  if (last < 2) return { rows: [] };
  var cols    = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, cols).getValues()[0];
  var refs    = sheet.getRange(2, 1, last - 1, 1).getValues();
  var out = [];
  for (var i = 0; i < refs.length; i++) {
    if ((refs[i][0] || '').toString().trim() !== ref) continue;
    var vals = sheet.getRange(i + 2, 1, 1, cols).getValues()[0];
    var o = {};
    for (var k = 0; k < headers.length; k++) { if (headers[k]) o[headers[k]] = vals[k]; }
    out.push(o);
  }
  return { rows: out };
}

function _latestAuditRef_(ss) {
  var sheet = ss.getSheetByName(AUDIT_LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return '';
  var refs = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = refs.length - 1; i >= 0; i--) {
    var r = (refs[i][0] || '').toString().trim();
    if (r) return r;
  }
  return '';
}

// The newest Audit_Log row whose Observer Name is `observerName` (case-insensitive).
// '' when the name is blank or the observer has no audit. It never falls back to
// someone else's audit.
function _latestAuditRefBy_(ss, observerName) {
  var want = (observerName || '').toString().trim().toLowerCase();
  if (!want) return '';
  var sheet = ss.getSheetByName(AUDIT_LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return '';
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var oc = headers.indexOf('Observer Name') + 1;
  if (oc < 1) return '';
  var n    = sheet.getLastRow() - 1;
  var refs = sheet.getRange(2, 1, n, 1).getValues();
  var obs  = sheet.getRange(2, oc, n, 1).getValues();
  for (var i = n - 1; i >= 0; i--) {
    var ref = (refs[i][0] || '').toString().trim();
    if (ref && (obs[i][0] || '').toString().trim().toLowerCase() === want) return ref;
  }
  return '';
}

// Like verifyLatestAudit, but for the newest audit made by the person running it.
// verifyLatestAudit takes the newest audit in the whole sheet, which on a busy day is
// usually another analyst's. The observer is worked out the same way the page does it
// (resolveObserver_, from the signed-in account).
function verifyMyLatestAudit() {
  requireAdmin_();
  var ss  = getOrCreateSpreadsheet();
  var obs = {};
  try { obs = resolveObserver_() || {}; } catch (e) { Logger.log('verifyMyLatestAudit: resolveObserver_ failed: ' + e); }
  var name = (obs.name || '').toString().trim();
  Logger.log('Looking for the latest audit by observer: "' + name + '"');
  var ref = _latestAuditRefBy_(ss, name);
  if (!ref) {
    return _finishAuditCheck_([{
      level: 'fail',
      name: 'An audit by "' + name + '" exists in Audit_Log',
      detail: name ? 'none found; check the Observer Name column, or use verifyLatestAudit instead'
                   : 'the observer name could not be worked out for the account running this'
    }], '');
  }
  return verifyLatestAudit(ref);
}

// This agent's saved evaluation as the Cache SHEET holds it, or null. Mirrors how
// saveCachedResult_ writes it: the chunks of one report are consecutive rows whose
// type is the analysis type, then type_2, type_3 ... and column E is the owner.
function _cacheSheetHtml_(ss, interactionId, typeLower, sap) {
  var sheet = ss.getSheetByName(CACHE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var n = sheet.getLastRow() - 1;
  var idType = sheet.getRange(2, 1, n, 2).getValues();
  var owners = sheet.getRange(2, 5, n, 1).getValues();
  var html = '', found = false;
  for (var i = 0; i < n; i++) {
    if ((idType[i][0] || '').toString().trim() !== interactionId) continue;
    var t = (idType[i][1] || '').toString().trim().toLowerCase();
    if (!(t === typeLower || t.indexOf(typeLower + '_') === 0)) continue;
    if ((owners[i][0] || '').toString().trim() !== sap) continue;
    html += (sheet.getRange(i + 2, 4, 1, 1).getValues()[0][0] || '').toString();
    found = true;
  }
  return found ? html : null;
}

function verifyLatestAudit(auditRef) {
  requireAdmin_();
  var checks = [];
  var add = function(level, name, detail) { checks.push({ level: level, name: name, detail: detail || '' }); };
  var ss  = getOrCreateSpreadsheet();

  var ref = (auditRef || '').toString().trim() || _latestAuditRef_(ss);
  if (!ref) {
    add('fail', 'An audit exists to check', 'Audit_Log has no rows yet');
    return _finishAuditCheck_(checks, '');
  }
  add('info', 'Checking audit', ref);

  // 1. Audit_Log — the row the email link and the dashboard depend on
  var logRes = _auditRowsForRef_(ss, AUDIT_LOG_SHEET, ref);
  var logRow = logRes.rows[0] || null;
  if (logRes.rows.length === 1) add('pass', 'Audit_Log has exactly one row for this ref', '');
  else if (logRes.rows.length === 0) add('fail', 'Audit_Log has a row for this ref', 'none found: the email step would say "Audit reference not found"');
  else add('fail', 'Audit_Log lists this ref only once', logRes.rows.length + ' rows share it, so the evaluation link opens the first one for everybody');

  var interactionId = logRow ? (logRow['Interaction ID'] || '').toString().trim() : '';
  var sap           = logRow ? (logRow['SAP ID'] || '').toString().trim() : '';
  var typeLabel     = logRow ? (logRow['Analysis Type'] || '').toString().trim() : '';
  var typeLower     = /sales/i.test(typeLabel) ? 'sales' : 'repeats';

  // 2. Transcripts
  var tr = _auditRowsForRef_(ss, TRANSCRIPTS_SHEET, ref);
  if (tr.rows.length >= 1) {
    var len = (tr.rows[0]['Transcript'] || '').toString().length;
    if (len > 0) add('pass', 'Transcripts has the call text', len + ' characters');
    else add('fail', 'Transcripts has the call text', 'the row exists but the transcript is empty');
  } else add('fail', 'Transcripts has a row for this ref', 'none found');

  // 3. Analysis
  var an = _auditRowsForRef_(ss, ANALYSIS_SHEET, ref);
  if (an.rows.length >= 1) {
    var alen = (an.rows[0]['AI Analysis Result'] || '').toString().length;
    if (alen > 0) add('pass', 'Analysis has the evaluation text', alen + ' characters');
    else add('fail', 'Analysis has the evaluation text', 'the row exists but the result is empty');
  } else add('fail', 'Analysis has a row for this ref', 'none found');

  // 4. Dashboard_Data
  var db = _auditRowsForRef_(ss, DASHBOARD_DATA_SHEET, ref);
  var dRow = db.rows[0] || null;
  if (db.rows.length === 1) add('pass', 'Dashboard_Data has exactly one row for this ref', '');
  else if (db.rows.length === 0) add('fail', 'Dashboard_Data has a row for this ref', 'none found: this audit would not appear on the dashboard');
  else add('fail', 'Dashboard_Data lists this ref only once', db.rows.length + ' rows');
  if (dRow) {
    var must = ['Interaction ID', 'SAP ID', 'Team Member', 'Call Summary'];
    var optional = ['Team Leader', 'Operations Manager', 'Line of Business', 'Locale', 'VTID', 'Repeat Projection %', 'Issue Resolved', 'Call Driver', 'RCA Category'];
    var blankMust = must.filter(function(h) { return !(dRow[h] || '').toString().trim(); });
    var blankOpt  = optional.filter(function(h) { return !(dRow[h] || '').toString().trim(); });
    if (blankMust.length) add('fail', 'Dashboard_Data has its key fields', 'blank: ' + blankMust.join(', '));
    else add('pass', 'Dashboard_Data has its key fields', '');
    if (blankOpt.length) add('info', 'Dashboard_Data fields left blank (may be normal)', blankOpt.join(', '));
  }

  // 5. The same agent and interaction everywhere
  if (logRow) {
    var disagree = [];
    var same = function(label, a, b) { if (a !== undefined && (a || '').toString().trim() !== (b || '').toString().trim()) disagree.push(label); };
    if (dRow)             { same('Dashboard_Data SAP ID', dRow['SAP ID'], sap); same('Dashboard_Data Interaction ID', dRow['Interaction ID'], interactionId); }
    if (tr.rows[0])       { same('Transcripts SAP ID', tr.rows[0]['SAP ID'], sap); same('Transcripts Interaction ID', tr.rows[0]['Interaction ID'], interactionId); }
    if (an.rows[0])       { same('Analysis SAP ID', an.rows[0]['SAP ID'], sap); }
    if (disagree.length) add('fail', 'The sheets agree on SAP ID and Interaction ID', 'these differ from Audit_Log: ' + disagree.join(', '));
    else add('pass', 'The sheets agree on SAP ID and Interaction ID', '');
  }

  // 6. The saved evaluation, read from the Cache SHEET (what the evaluation page falls back to)
  if (logRow && interactionId) {
    var html = _cacheSheetHtml_(ss, interactionId, typeLower, sap);
    if (html === null) {
      add('fail', 'A saved evaluation exists for this agent', 'none in the Cache sheet: the evaluation page would be empty');
    } else if (html.indexOf('#ERROR!') !== -1) {
      add('fail', 'The saved evaluation is intact', 'contains #ERROR!: a chunk was turned into a formula error');
    } else if (!html.trim()) {
      add('fail', 'The saved evaluation is intact', 'it is empty');
    } else if (!/<\/(div|ul|table|p|section)>\s*$/i.test(html)) {
      add('fail', 'The saved evaluation is complete', 'it ends mid-way and looks cut off');
    } else {
      add('pass', 'The saved evaluation is intact and complete', html.length + ' characters');
    }
  } else if (logRow) {
    add('fail', 'A saved evaluation can be looked up', 'the audit has no Interaction ID');
  }

  // 7. Email
  if (logRow) {
    var status = (logRow['Email Status'] || '').toString().trim();
    var recips = (logRow['Recipients'] || '').toString().trim();
    if (status === 'Sent') {
      if (recips) add('pass', 'Email was sent and its recipients recorded', recips);
      else add('fail', 'Email was sent and its recipients recorded', 'status is Sent but no recipients are recorded');
    } else if (!status || status === 'Not Sent') {
      add('info', 'Email not sent yet (Email Status: ' + (status || 'blank') + ')', 'send it from the page, then run this again to confirm');
    } else {
      add('fail', 'Email was sent', 'Email Status says: ' + status);
    }
  }

  // 8. The "being processed" flag must not be left behind
  if (interactionId) {
    var flag = CacheService.getScriptCache().get(_procKey_(interactionId, sap));
    if (flag) add('fail', 'No "being processed" flag is left over', 'still set: the analyst would be told it is already being processed');
    else add('pass', 'No "being processed" flag is left over', '');
  }

  // 9. Errors logged for this interaction since the audit
  var errSheet = ss.getSheetByName(ERROR_LOG_SHEET);
  if (errSheet && errSheet.getLastRow() >= 2 && interactionId) {
    var from  = Math.max(2, errSheet.getLastRow() - 199);
    var ecols = errSheet.getLastColumn();
    var eh    = errSheet.getRange(1, 1, 1, ecols).getValues()[0];
    var erows = errSheet.getRange(from, 1, errSheet.getLastRow() - from + 1, ecols).getValues();
    var col = function(h) { return eh.indexOf(h); };
    var since = logRow && logRow['Submitted At'] ? new Date(logRow['Submitted At']).getTime() : 0;
    var hits = erows.filter(function(r) {
      var when = new Date(r[col('Timestamp')]).getTime();
      return (r[col('Interaction ID')] || '').toString().trim() === interactionId && (!since || isNaN(when) || when >= since - 1000);
    });
    if (hits.length) add('fail', 'No errors were logged for this interaction since the audit',
                         hits.length + ' in Error_Log, first: ' + (hits[0][col('Function')] || '') + ': ' + (hits[0][col('Error Message')] || '').toString().slice(0, 120));
    else add('pass', 'No errors were logged for this interaction since the audit', '');
  }

  return _finishAuditCheck_(checks, ref);
}

// ── The data BEHIND the dashboard ────────────────────────────────────────────────────
// The dashboard is built FROM Audit_Log and joined to Dashboard_Data by Audit Ref, then
// computes its KPIs from the merged rows. This reads both sheets (nothing else, nothing
// is written), checks that join and the values the charts depend on, and works out the
// KPIs the way the dashboard does, so they can be compared with what is on the screen.

// Every row of a sheet as {header: value} (raw values: Dates stay Dates). Reads the
// header row first and returns { missing, headers, rows }.
function _sheetObjects_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) return { missing: true, headers: [], rows: [] };
  var last = sheet.getLastRow(), cols = sheet.getLastColumn();
  var headers = cols ? sheet.getRange(1, 1, 1, cols).getValues()[0] : [];
  if (last < 2) return { headers: headers, rows: [] };
  var vals = sheet.getRange(2, 1, last - 1, cols).getValues();
  var rows = vals.map(function(r) {
    var o = {};
    for (var k = 0; k < headers.length; k++) { if (headers[k]) o[headers[k]] = r[k]; }
    return o;
  });
  return { headers: headers, rows: rows };
}

function verifyDashboardData() {
  requireAdmin_();
  var ss = getOrCreateSpreadsheet();
  var L = [], problems = 0, S = {};
  var line = function(level, text) { L.push({ level: level, text: text }); if (level === 'fail') problems++; };
  var str  = function(v) { return (v === null || v === undefined) ? '' : v.toString().trim(); };
  var list = function(a, n) { return a.slice(0, n || 8).join(', ') + (a.length > (n || 8) ? ' … (+' + (a.length - (n || 8)) + ' more)' : ''); };

  var log  = _sheetObjects_(ss, AUDIT_LOG_SHEET);
  var dash = _sheetObjects_(ss, DASHBOARD_DATA_SHEET);
  var count = function(name) { var s = ss.getSheetByName(name); return s ? Math.max(0, s.getLastRow() - 1) : 0; };
  S.auditLogRows = log.rows.length; S.dashboardRows = dash.rows.length;
  line('info', 'Rows: Audit_Log ' + log.rows.length + ', Dashboard_Data ' + dash.rows.length + ', Analysis ' + count(ANALYSIS_SHEET) + ', Transcripts ' + count(TRANSCRIPTS_SHEET));

  // 1. Headers: the code reads columns BY NAME, so a missing header is silently blank
  [[AUDIT_LOG_SHEET, log, AUDIT_LOG_HEADERS], [DASHBOARD_DATA_SHEET, dash, DASHBOARD_HEADERS]].forEach(function(t) {
    if (t[1].missing) { line('fail', t[0] + ' sheet is missing'); return; }
    var absent = t[2].filter(function(h) { return t[1].headers.indexOf(h) === -1; });
    if (absent.length) line('fail', t[0] + ' is missing columns the code reads: ' + absent.join(', ') + ' (they will read as blank)');
    else line('pass', t[0] + ' has every expected column');
  });

  // 2. The join the dashboard depends on
  var tally = function(rows) { var m = {}; rows.forEach(function(r) { var k = str(r['Audit Ref']); if (k) m[k] = (m[k] || 0) + 1; }); return m; };
  var blankRef = function(rows) { return rows.filter(function(r) { return !str(r['Audit Ref']); }).length; };
  var logRefs = tally(log.rows), dashRefs = tally(dash.rows);
  var dupOf = function(m) { return Object.keys(m).filter(function(k) { return m[k] > 1; }); };
  [['Audit_Log', log.rows, logRefs], ['Dashboard_Data', dash.rows, dashRefs]].forEach(function(t) {
    var b = blankRef(t[1]);
    if (b) line('fail', t[0] + ' has ' + b + ' row(s) with no Audit Ref (they cannot be joined)');
    var d = dupOf(t[2]);
    if (d.length) line('fail', t[0] + ' lists ' + d.length + ' Audit Ref(s) more than once (duplicate rows): ' + list(d));
  });
  var hidden   = Object.keys(dashRefs).filter(function(k) { return !logRefs[k]; });
  var noDash   = Object.keys(logRefs).filter(function(k) { return !dashRefs[k]; });
  S.dashboardOnly = hidden.length; S.auditOnly = noDash.length;
  if (hidden.length) line('fail', hidden.length + ' audit(s) are in Dashboard_Data but NOT in Audit_Log, so the dashboard HIDES them: ' + list(hidden));
  if (noDash.length) line('fail', noDash.length + ' audit(s) are in Audit_Log with no Dashboard_Data row; the dashboard counts them but with blank analysis: ' + list(noDash));
  if (!hidden.length && !noDash.length) line('pass', 'Audit_Log and Dashboard_Data hold the same audits');

  // The merged rows, built the way the dashboard builds them: Audit_Log first, Dashboard_Data as fallback
  var dashByRef = {};
  dash.rows.forEach(function(r) { var k = str(r['Audit Ref']); if (k) dashByRef[k] = r; });
  var merged = log.rows.map(function(l) {
    var d = dashByRef[str(l['Audit Ref'])] || {};
    var pick = function(h, dh) { return str(l[h]) || str(d[dh || h]); };
    return { ref: str(l['Audit Ref']), ts: l['Submitted At'] || d['Timestamp'] || '', sap: pick('SAP ID'), member: pick('Team Member'),
             type: pick('Analysis Type'), repeat: pick('Repeat Projection %'), resolved: pick('Issue Resolved'), email: pick('Email Status'),
             driver: str(d['Call Driver']), hasDash: !!dashByRef[str(l['Audit Ref'])] };
  });

  // 3. The same agent in both sheets
  var sapBad = [];
  log.rows.forEach(function(l) { var d = dashByRef[str(l['Audit Ref'])]; if (d && str(l['SAP ID']) && str(d['SAP ID']) && str(l['SAP ID']) !== str(d['SAP ID'])) sapBad.push(str(l['Audit Ref'])); });
  S.sapMismatch = sapBad.length;
  if (sapBad.length) line('fail', sapBad.length + ' audit(s) have a different SAP ID in Audit_Log and Dashboard_Data: ' + list(sapBad));

  // 4. Tabs: the dashboard puts an audit in Repeats or Sales by its Analysis Type
  S.repeats = merged.filter(function(m) { return m.type.indexOf('Repeats') !== -1; }).length;
  S.sales   = merged.filter(function(m) { return m.type.indexOf('Sales') !== -1; }).length;
  S.neither = merged.length - S.repeats - S.sales;
  line('info', 'Tabs: Repeats ' + S.repeats + ', Sales ' + S.sales + ' (Audit Log tab ' + merged.length + ')');
  if (S.neither) line('fail', S.neither + ' audit(s) are in NO tab (Analysis Type is neither Repeats nor Sales), so Repeats + Sales is less than the Audit Log: ' +
                       list(merged.filter(function(m) { return m.type.indexOf('Repeats') === -1 && m.type.indexOf('Sales') === -1; }).map(function(m) { return m.ref; })));

  // 5. Issue Resolved: exact answers, and what the dashboard KPI shows
  // The dashboard reads the START of the text as a verdict (same rule as _dbResolution in index.html / Dashboard.html).
  var verdictOf = function(raw) {
    var s = String(raw === null || raw === undefined ? '' : raw).trim().toLowerCase().replace(/^[^a-z]+/, '');
    if (!s) return 'unknown';
    if (/^(partial|partially|in part|mostly)\b/.test(s)) return 'partial';
    if (/^(no|not|unresolved|completely unresolved|fully unresolved|never)\b/.test(s)) return 'no';
    if (/^(yes|resolved|completely resolved|fully resolved)\b/.test(s)) return 'yes';
    // "Issue was NOT resolved." / "The issue was resolved": a verdict after a short lead-in about the issue.
    // A verdict written as a short sentence: "Issue not resolved. ...", "Issue resolved — ...", "Neither issue was resolved". A yes is not taken
    // from a sentence that goes on to qualify it (but / however / remains), so that one is left for a person to read.
    if (/^(?:the |this )?(?:primary |main |original )?(?:issue|problem|request)s?\s+(?:was |is |were |are |has been |have been )?(?:ultimately |eventually |still )?(?:not|never)\s+(?:been )?(?:fully |completely )?(?:resolved|fixed)\b/.test(s)) return 'no';
    if (/^neither (?:issue|problem|request) was (?:resolved|fixed)\b/.test(s)) return 'no';
    if (/^(?:the |this )?(?:primary |main |original )?(?:issue|problem|request)\s+(?:was |is |has been )?(?:ultimately |eventually |finally |successfully |fully |completely )*(?:resolved|fixed)\b/.test(s) && !/\b(but|however|although|except|remains?|remained|secondary|second issue|other issue)\b/.test(s.slice(0, 160))) return 'yes';
    var m = /^(?:the |this )?(?:customer'?s )?(?:issue|problem|concern|request)\b[^.]{0,25}?\b(was|is|has been|wasn'?t|isn'?t)\s+(not |never |partially |only partially )?(?:fully |completely )?(?:resolved|fixed)/.exec(s);
    if (m) return /partial/.test(m[2] || '') ? 'partial' : (/n'?t$/.test(m[1]) || /^(not|never)/.test(m[2] || '')) ? 'no' : 'yes';
    return 'unknown';
  };
  var yes = 0, partial = 0, no = 0, blank = 0, unknown = 0, unknownVals = {};
  merged.forEach(function(m) {
    var v = verdictOf(m.resolved);
    if (v === 'yes') yes++; else if (v === 'partial') partial++; else if (v === 'no') no++;
    else if (!m.resolved.trim()) blank++;
    else { unknown++; unknownVals[m.resolved.substring(0, 60)] = true; }
  });
  S.resolvedYes = yes; S.resolvedPartial = partial; S.resolvedNo = no; S.resolvedBlank = blank; S.resolvedUnknown = unknown;
  S.dashResolved = yes; S.dashPartial = partial; S.dashUnresolved = no; S.dashNotRecorded = blank + unknown;
  line('info', 'Issue Resolved as the dashboard shows it: Resolved ' + yes + ', Partially resolved ' + partial + ', Unresolved ' + no + ', Not recorded ' + (blank + unknown) + ' (blank ' + blank + ', no clear verdict ' + unknown + ')');
  if (unknown) line('warn', unknown + ' Issue Resolved value(s) do not start with a verdict, so the dashboard shows them as Not Recorded, e.g. ' + list(Object.keys(unknownVals), 3));

  // 6. Repeat risk: blank is not 0, and a real 0% belongs in the average
  var nums = [], rBlank = 0, rZero = 0, rBad = 0, rOver = 0, rText = 0, rFrac = 0;
  merged.forEach(function(m) {
    if (!m.repeat) { rBlank++; return; }
    // Same reading as _dbNum in index.html / Dashboard.html: a number, or the first "NN%" inside a sentence.
    var txt = String(m.repeat).trim(), plain = txt.replace('%', '').trim(), n = null;
    if (plain !== '' && isFinite(Number(plain))) {
      n = Number(plain);
      if (txt.indexOf('%') === -1 && n > 0 && n <= 1) { n = Math.round(n * 1000) / 10; rFrac++; }
    } else {
      var rng = /(\d+(?:\.\d+)?)\s*[-–]\s*(\d+(?:\.\d+)?)\s*%/.exec(txt), pc = /(\d+(?:\.\d+)?)\s*%/.exec(txt);
      if (rng) { n = (Number(rng[1]) + Number(rng[2])) / 2; rText++; }
      else if (pc && Number(pc[1]) <= 100) { n = Number(pc[1]); rText++; }
    }
    if (n === null) { rBad++; return; }
    nums.push(n); if (n === 0) rZero++; if (n > 100) rOver++;
  });
  var avg = function(a) { return a.length ? Math.round(a.reduce(function(s, v) { return s + v; }, 0) / a.length) : 0; };
  S.repeatFromText = rText; S.repeatFraction = rFrac;
  if (rFrac) line('info', rFrac + ' Repeat Projection % cell(s) hold a fraction (0.78 for 78%); the dashboard reads them as percentages');
  if (rText) line('warn', rText + ' Repeat Projection % cell(s) hold a sentence instead of a number; the dashboard reads the first NN% in it, but the cell should be just the number (run previewDataRepairs / applyDataRepairs)');
  S.repeatBlank = rBlank; S.repeatZero = rZero; S.repeatNumeric = nums.length; S.repeatNonNumeric = rBad; S.repeatOver100 = rOver;
  S.repeatAvgTrue = avg(nums);   // the dashboard uses the same average (real 0% counts, blank and text do not)
  line('info', 'Repeat Projection %: ' + nums.length + ' numeric (' + rZero + ' of them 0%), ' + rBlank + ' blank, ' + rBad + ' not a number, ' + rOver + ' over 100');
  if (rBad) line('fail', rBad + ' Repeat Projection % value(s) are not numbers');
  if (rOver) line('fail', rOver + ' Repeat Projection % value(s) are over 100');

  // 7. Timestamps: a date filter hides a row with no valid date
  var now = new Date().getTime(), badDates = [];
  merged.forEach(function(m) {
    var t = (m.ts instanceof Date) ? m.ts.getTime() : (m.ts ? Date.parse(m.ts.toString()) : NaN);
    if (isNaN(t) || t > now + 86400000) badDates.push(m.ref);
  });
  S.badDates = badDates.length;
  if (badDates.length) line('fail', badDates.length + ' audit(s) have a blank, invalid or future timestamp (any date filter hides them): ' + list(badDates));

  // 8. A Team Member that is a paragraph means a column shifted
  var longNames = merged.filter(function(m) { return m.member.length > 80; });
  S.longNames = longNames.length;
  if (longNames.length) line('fail', longNames.length + ' audit(s) have a Team Member over 80 characters (a column shifted): ' + list(longNames.map(function(m) { return m.ref; })));

  // 9. Email status
  var emailCounts = {};
  merged.forEach(function(m) { var k = m.email || '(blank)'; emailCounts[k] = (emailCounts[k] || 0) + 1; });
  S.emailSent = emailCounts['Sent'] || 0;
  line('info', 'Email Status: ' + Object.keys(emailCounts).map(function(k) { return k + ' ' + emailCounts[k]; }).join(', '));

  // 10. Call Drivers that differ only by capitals or spacing become separate bars
  var groups = {};
  merged.forEach(function(m) {
    if (!m.driver) return;
    var k = m.driver.toLowerCase().replace(/\s+/g, ' ').trim();
    (groups[k] = groups[k] || {})[m.driver] = 1;
  });
  var variants = Object.keys(groups).filter(function(k) { return Object.keys(groups[k]).length > 1; });
  S.driverVariantGroups = variants.length;
  variants.slice(0, 5).forEach(function(k) { line('warn', 'Call Driver written several ways (separate bars): ' + Object.keys(groups[k]).map(function(v) { return '"' + v + '"'; }).join(', ')); });

  L.forEach(function(l) { Logger.log('[' + l.level.toUpperCase() + '] ' + l.text); });
  var summary = problems ? 'RESULT: ' + problems + ' PROBLEM(S) FOUND' : 'RESULT: NO PROBLEMS FOUND';
  Logger.log(summary);
  return { ok: problems === 0, problems: problems, summary: summary, stats: S, lines: L };
}

function _finishAuditCheck_(checks, ref) {
  var bad = checks.filter(function(c) { return c.level === 'fail'; }).length;
  var real = checks.filter(function(c) { return c.level !== 'info'; }).length;
  checks.forEach(function(c) {
    Logger.log('[' + c.level.toUpperCase() + '] ' + c.name + (c.detail ? ' — ' + c.detail : ''));
  });
  var summary = bad ? 'RESULT: ' + bad + ' of ' + real + ' CHECKS FAILED' : 'RESULT: ALL ' + real + ' CHECKS PASSED';
  Logger.log(summary);
  return { ok: bad === 0, ref: ref, summary: summary, checks: checks };
}
