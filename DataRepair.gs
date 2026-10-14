/**
 * DataRepair.gs
 * Admin-only tools for the data problems verifyDashboardData() reports.
 *
 *   addMissingHeaders()   writes a header that is missing at the END of a sheet's header row
 *                         (for example "Agent Email"). It moves no data and clears nothing.
 *   diagnoseDataIssues()  read-only. Prints the exact rows behind each problem, across
 *                         Audit_Log, Dashboard_Data, Transcripts and Analysis, so a fix can be
 *                         chosen from what is really in the sheet.
 *
 * Run from the Apps Script editor.
 */

// Header row of a sheet with the empty cells at the end removed.
function _trimmedHeaders_(sheet) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return [];
  var h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(x) { return String(x === null || x === undefined ? '' : x).trim(); });
  while (h.length && !h[h.length - 1]) h.pop();
  return h;
}

// What addMissingHeaders() would do to one sheet, as a pure decision:
// { action: 'ok' | 'add' | 'skip', add: [names], from: column number, reason }
function _planMissingHeaders_(current, target) {
  var i;
  if (current.length > target.length) return { action: 'skip', add: [], from: 0, reason: 'the sheet has more header columns (' + current.length + ') than the code expects (' + target.length + ')' };
  for (i = 0; i < current.length; i++) {
    if (current[i] !== target[i]) return { action: 'skip', add: [], from: 0, reason: 'column ' + (i + 1) + ' is "' + current[i] + '" but the code expects "' + target[i] + '" (the headers are not just missing at the end, so nothing was changed)' };
  }
  if (current.length === target.length) return { action: 'ok', add: [], from: 0, reason: 'every header is there' };
  return { action: 'add', add: target.slice(current.length), from: current.length + 1, reason: '' };
}

function addMissingHeaders() {
  requireAdmin_();
  var ss = getOrCreateSpreadsheet();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var result = { added: 0, skipped: 0, sheets: {} };
  try {
    [[TRANSCRIPTS_SHEET, TRANSCRIPTS_HEADERS], [ANALYSIS_SHEET, ANALYSIS_HEADERS],
     [DASHBOARD_DATA_SHEET, DASHBOARD_HEADERS], [AUDIT_LOG_SHEET, AUDIT_LOG_HEADERS]].forEach(function(pair) {
      var name = pair[0], sheet = ss.getSheetByName(name);
      if (!sheet) { Logger.log(name + ': sheet not found, skipped'); result.skipped++; result.sheets[name] = 'missing sheet'; return; }
      var plan = _planMissingHeaders_(_trimmedHeaders_(sheet), pair[1]);
      if (plan.action === 'ok') { Logger.log(name + ': ' + plan.reason); result.sheets[name] = 'ok'; return; }
      if (plan.action === 'skip') { Logger.log(name + ': SKIPPED, ' + plan.reason); result.skipped++; result.sheets[name] = 'skipped'; return; }
      plan.add.forEach(function(h, k) { sheet.getRange(1, plan.from + k).setValue(h); });
      Logger.log(name + ': added header(s) ' + plan.add.join(', ') + ' at column ' + plan.from);
      result.added += plan.add.length;
      result.sheets[name] = 'added ' + plan.add.join(', ');
    });
    if (typeof invalidateDashboardCache_ === 'function') invalidateDashboardCache_();
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
  Logger.log('addMissingHeaders: ' + result.added + ' header(s) added, ' + result.skipped + ' sheet(s) skipped.');
  return result;
}

// One column of a sheet by header name, short strings only (never loads a transcript or an evaluation).
function _colByName_(sheet, headers, name) {
  var c = headers.indexOf(name) + 1;
  if (!sheet || c < 1 || sheet.getLastRow() < 2) return null;
  return sheet.getRange(2, c, sheet.getLastRow() - 1, 1).getValues().map(function(r) { return r[0]; });
}
function _short_(v, n) { var s = String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim(); return s.length > n ? s.substring(0, n) + '…' : s; }

// READ-ONLY. Prints the rows behind each problem verifyDashboardData() reports.
function diagnoseDataIssues() {
  requireAdmin_();
  var ss = getOrCreateSpreadsheet();
  var out = [];
  var say = function(t) { out.push(t); Logger.log(t); };
  var names = [AUDIT_LOG_SHEET, DASHBOARD_DATA_SHEET, TRANSCRIPTS_SHEET, ANALYSIS_SHEET];
  var data = {};   // sheet -> { headers, ref, cols: {name: values[]} }
  var want = ['Audit Ref', 'SAP ID', 'Interaction ID', 'Team Member', 'Timestamp', 'Submitted At', 'Analysis Type', 'Repeat Projection %', 'Issue Resolved', 'Email Status'];
  names.forEach(function(n) {
    var sh = ss.getSheetByName(n);
    var hd = _trimmedHeaders_(sh);
    var d = { headers: hd, cols: {}, rows: sh ? Math.max(0, sh.getLastRow() - 1) : 0, sheet: sh };
    want.forEach(function(w) { var v = _colByName_(sh, hd, w); if (v) d.cols[w] = v; });
    data[n] = d;
  });

  var refsIn = function(n) { return (data[n].cols['Audit Ref'] || []).map(function(r) { return String(r || '').trim(); }); };
  var show = function(ref) {
    say('  ' + ref);
    names.forEach(function(n) {
      var refs = refsIn(n), found = false;
      refs.forEach(function(r, i) {
        if (r !== ref) return;
        found = true;
        var c = data[n].cols;
        say('    ' + n + ' row ' + (i + 2) + ': SAP ' + _short_(c['SAP ID'] && c['SAP ID'][i], 12) + ' | interaction ' + _short_(c['Interaction ID'] && c['Interaction ID'][i], 30) +
            ' | member "' + _short_(c['Team Member'] && c['Team Member'][i], 40) + '" | when ' + _short_((c['Submitted At'] || c['Timestamp'] || [])[i], 25) +
            ' | type ' + _short_(c['Analysis Type'] && c['Analysis Type'][i], 30));
      });
      if (!found) say('    ' + n + ': NO ROW');
    });
  };

  say('=== 1. References used more than once ===');
  var seen = {}, dup = {};
  [AUDIT_LOG_SHEET, DASHBOARD_DATA_SHEET].forEach(function(n) {
    var cnt = {};
    refsIn(n).forEach(function(r) { if (r) cnt[r] = (cnt[r] || 0) + 1; });
    Object.keys(cnt).forEach(function(r) { if (cnt[r] > 1) dup[r] = true; });
  });
  Object.keys(dup).forEach(show);
  if (!Object.keys(dup).length) say('  none');

  say('=== 2. In Dashboard_Data but not in Audit_Log ===');
  var inLog = {}; refsIn(AUDIT_LOG_SHEET).forEach(function(r) { inLog[r] = true; });
  var hidden = {}; refsIn(DASHBOARD_DATA_SHEET).forEach(function(r) { if (r && !inLog[r]) hidden[r] = true; });
  Object.keys(hidden).forEach(show);
  if (!Object.keys(hidden).length) say('  none');

  say('=== 3. Different SAP ID in Audit_Log and Dashboard_Data ===');
  var dRefs = refsIn(DASHBOARD_DATA_SHEET), lRefs = refsIn(AUDIT_LOG_SHEET), lSap = data[AUDIT_LOG_SHEET].cols['SAP ID'] || [], dSap = data[DASHBOARD_DATA_SHEET].cols['SAP ID'] || [];
  var mism = {};
  lRefs.forEach(function(r, i) { var j = dRefs.indexOf(r); if (r && j !== -1 && String(lSap[i] || '').trim() && String(dSap[j] || '').trim() && String(lSap[i]).trim() !== String(dSap[j]).trim()) mism[r] = true; });
  Object.keys(mism).forEach(show);
  if (!Object.keys(mism).length) say('  none');

  say('=== 4. Repeat Projection % that is not a number ===');
  var bad = 0, sentences = 0, fractions = 0;
  [AUDIT_LOG_SHEET, DASHBOARD_DATA_SHEET].forEach(function(n) {
    var v = data[n].cols['Repeat Projection %'] || [], refs = refsIn(n), shownS = 0;
    v.forEach(function(x, i) {
      var s = String(x === null || x === undefined ? '' : x).replace('%', '').trim();
      if (s === '') return;
      if (isFinite(Number(s))) { if (typeof x === 'number' && x > 0 && x < 1) fractions++; return; }
      if (_extractRepeatNumber_(x) !== null) {
        sentences++;
        if (shownS++ < 5) say('  ' + n + ' row ' + (i + 2) + ' ' + refs[i] + ' (sentence, reads ' + _extractRepeatNumber_(x) + '): "' + _short_(x, 50) + '"');
      } else { bad++; say('  ' + n + ' row ' + (i + 2) + ' ' + refs[i] + ' (NO number in it): "' + _short_(x, 60) + '"'); }
    });
  });
  say('  summary: ' + sentences + ' cell(s) hold a sentence that contains a percentage, ' + bad + ' hold text with no percentage, ' + fractions + ' hold a fraction between 0 and 1 (a percent that the sheet turned into 0.xx)');

  say('=== 5. Team Member longer than 80 characters (a column shifted) ===');
  var shifted = 0;
  [AUDIT_LOG_SHEET, DASHBOARD_DATA_SHEET].forEach(function(n) {
    var m = data[n].cols['Team Member'] || [], refs = refsIn(n);
    m.forEach(function(x, i) {
      if (String(x || '').length > 80) {
        shifted++;
        var row = data[n].sheet.getRange(i + 2, 1, 1, data[n].headers.length).getValues()[0];
        say('  ' + n + ' row ' + (i + 2) + ' ' + refs[i] + ':');
        data[n].headers.forEach(function(h, k) { if (row[k] !== '' && row[k] !== null && row[k] !== undefined) say('      ' + h + ' = ' + _short_(row[k], 45)); });
      }
    });
  });
  if (!shifted) say('  none');
  say('=== end: copy everything above into the chat ===');
  return { duplicates: Object.keys(dup), dashboardOnly: Object.keys(hidden), sapMismatch: Object.keys(mism), nonNumericRepeat: bad, sentenceRepeat: sentences, fractionRepeat: fractions, shiftedRows: shifted, lines: out };
}

// ── Repairs with a preview ───────────────────────────────────────────────────
// A number from a Repeat Projection % cell, or null. A sentence gives the first "NN%" in it
// (same reading as _dbNum in index.html / Dashboard.html).
function _extractRepeatNumber_(v) {
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (s === '') return null;
  var plain = s.replace('%', '').trim();
  if (isFinite(Number(plain))) return Number(plain);
  var rng = /(\d+(?:\.\d+)?)\s*[-–]\s*(\d+(?:\.\d+)?)\s*%/.exec(s);
  if (rng) return Math.round((Number(rng[1]) + Number(rng[2])) / 2);
  var m = /(\d+(?:\.\d+)?)\s*%/.exec(s);
  if (m && Number(m[1]) <= 100) return Math.round(Number(m[1]));
  return null;
}

// Agents chosen by the admin on 2026-10-13 for two audits whose SAP ID has no clean audit to learn the name from.
// Each is used only when the name is inside the messy Team Member cell.
var _CHOSEN_AGENTS_ = {
  'NHA-20260903-0003': 'Ali Guennoun',
  'NHA-20260929-0024': 'Jonathan Cardoza'
};

// A call length that Sheets turned into a time of day (2:15:05 became a Date on the sheet's day zero, 1899-12-30). Returns it as
// "H:mm:ss" text, or '' when the value is not a time within that day (it is then left alone and reported).
function _durationFromDate_(d, tz) {
  if (Object.prototype.toString.call(d) !== '[object Date]' || isNaN(d.getTime())) return '';
  if (Utilities.formatDate(d, tz, 'yyyy-MM-dd') !== '1899-12-30') return '';
  return Utilities.formatDate(d, tz, 'H:mm:ss');
}

// A call length cell that Sheets turned into a Date was usually typed as "MM:SS" ("11:39" = 11 minutes 39 seconds, from the transcript
// header). Sheets read it as hours and minutes: 11:39 became 11 h 39 min, and "31:53" became a day and 7 h 53 min (the 1196 cells that
// are not on day zero). So a Date cannot simply be written back as H:mm:ss: for 11:39:00 that would say 11 hours. The truth is the
// "Transcript Duration:" line at the top of the saved transcript, and that text is what is written back.
var _DURATION_TEXT_ = /^\d{1,3}:[0-5]\d(?::[0-5]\d)?$/;
function _durationFromTranscript_(text) {
  var m = /Transcript Duration[:\s]+([^\n\r]+)/i.exec(String(text === null || text === undefined ? '' : text).substring(0, 3000));
  if (!m) return '';
  var s = m[1].trim();
  return _DURATION_TEXT_.test(s) ? s : '';
}
// Seconds from the sheet's day zero (1899-12-30 00:00:00) to a Date, read in the spreadsheet's own time zone.
function _clockSecondsOf_(d, tz) {
  if (Object.prototype.toString.call(d) !== '[object Date]' || isNaN(d.getTime())) return null;
  var p = Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm:ss').split(/[- :]/).map(Number);
  var days = Math.round((Date.UTC(p[0], p[1] - 1, p[2]) - Date.UTC(1899, 11, 30)) / 86400000);
  return days * 86400 + p[3] * 3600 + p[4] * 60 + p[5];
}
// The seconds that Sheets would read from a duration text typed into a cell: "11:39" is hours and minutes, "0:11:39" is hours, minutes, seconds.
function _sheetsSecondsOfText_(s) {
  var a = String(s).split(':').map(Number);
  return a.length === 2 ? a[0] * 3600 + a[1] * 60 : a[0] * 3600 + a[1] * 60 + a[2];
}
// What to write for one Date in a duration cell: { to, how } or { to: '', reason }. truth is the text from the transcript, or ''.
//  - the transcript text, when it explains the Date exactly (so a length the analyst typed over by hand is never overwritten)
//  - H:mm:ss, when the Date is on day zero and has seconds (a three-part length cannot be a mistyped MM:SS)
//  - otherwise nothing: a Date with no seconds is the "11:39" case and only the transcript can say which it was
function _durationDecision_(d, tz, truth) {
  var secs = _clockSecondsOf_(d, tz);
  if (secs === null) return { to: '', reason: 'not a valid date' };
  if (truth && _sheetsSecondsOfText_(truth) === secs) return { to: truth, how: 'transcript' };
  var asClock = _durationFromDate_(d, tz);
  if (asClock && secs % 60 !== 0) return { to: asClock, how: 'clock' };
  return { to: '', reason: truth ? 'the transcript says ' + truth + ', which does not match the cell' : (asClock ? 'it has no seconds, so it may be MM:SS typed as hours and minutes, and the saved transcript does not say' : 'a time that is not within one day and the saved transcript does not say') };
}
// Audit Ref -> the call length written in that audit's saved transcript (only the ones that have one).
function _durationTruthMap_(ss) {
  var map = {}, sh = ss.getSheetByName(TRANSCRIPTS_SHEET);
  if (!sh || sh.getLastRow() < 2) return map;
  var hd = _trimmedHeaders_(sh), refs = _colByName_(sh, hd, 'Audit Ref') || [], txt = _colByName_(sh, hd, 'Transcript') || [];
  refs.forEach(function(r, i) {
    r = String(r === null || r === undefined ? '' : r).trim();
    if (!r || map[r]) return;
    var d = _durationFromTranscript_(txt[i]);
    if (d) map[r] = d;
  });
  return map;
}

// What the repairs would change. Pure reading, no writes.
//   restores  an audit that is in Dashboard_Data but has no Audit_Log row: the row is rebuilt from Dashboard_Data
//   sapFixes  Dashboard_Data holds a different SAP ID than the three other sheets, which agree with each other
//   repeatFixes  a Repeat Projection % cell that holds a sentence: the cell becomes the number, the sentence moves to a cell note
function _planDataRepairs_(ss) {
  var plan = { restores: [], sapFixes: [], repeatFixes: [], skipped: [] };
  var L = ss.getSheetByName(AUDIT_LOG_SHEET), D = ss.getSheetByName(DASHBOARD_DATA_SHEET);
  var lh = _trimmedHeaders_(L), dh = _trimmedHeaders_(D);
  var col = function(sh, hd, n) { return _colByName_(sh, hd, n) || []; };
  var lRef = col(L, lh, 'Audit Ref').map(String), dRef = col(D, dh, 'Audit Ref').map(String);
  var trim = function(x) { return String(x === null || x === undefined ? '' : x).trim(); };
  var count = function(arr, r) { var n = 0; arr.forEach(function(x) { if (trim(x) === r) n++; }); return n; };

  // 1. audits only in Dashboard_Data
  var lSet = {}; lRef.forEach(function(r) { lSet[trim(r)] = true; });
  var seenD = {};
  dRef.forEach(function(r, i) {
    r = trim(r);
    if (!r || lSet[r] || seenD[r]) return;
    seenD[r] = true;
    if (count(dRef, r) > 1) { plan.skipped.push(r + ': in Dashboard_Data more than once, not rebuilt'); return; }
    var row = D.getRange(i + 2, 1, 1, dh.length).getValues()[0], d = {};
    dh.forEach(function(h, k) { d[h] = row[k]; });
    var member = trim(d['Team Member']);
    if (!trim(d['SAP ID']) || !trim(d['Interaction ID']) || !member || member.length > 80 || !d['Timestamp']) {
      plan.skipped.push(r + ': Dashboard_Data row lacks SAP ID, Interaction ID, Team Member or Timestamp, not rebuilt'); return;
    }
    var out = lh.map(function(h) {
      if (h === 'Submitted At') return d['Timestamp'];
      if (h === 'Duration') return d['Transcript Duration'] === undefined ? '' : d['Transcript Duration'];
      if (h === 'Recipients') return '';
      return d[h] === undefined ? '' : d[h];
    });
    plan.restores.push({ ref: r, row: out, member: member, sap: trim(d['SAP ID']) });
  });

  // 2. SAP ID: Dashboard_Data is the odd one out
  var lSap = col(L, lh, 'SAP ID'), dSap = col(D, dh, 'SAP ID');
  var other = {};
  [TRANSCRIPTS_SHEET, ANALYSIS_SHEET].forEach(function(n) {
    var sh = ss.getSheetByName(n), hd = _trimmedHeaders_(sh);
    other[n] = { refs: col(sh, hd, 'Audit Ref').map(trim), saps: col(sh, hd, 'SAP ID').map(trim) };
  });
  var doneRef = {};
  lRef.forEach(function(r, i) {
    r = trim(r);
    if (!r || doneRef[r]) return;
    doneRef[r] = true;
    var j = dRef.indexOf(r);
    if (j === -1 || count(lRef, r) !== 1 || count(dRef, r) !== 1) return;
    var a = trim(lSap[i]), b = trim(dSap[j]);
    if (!a || !b || a === b) return;
    var votes = {}; votes[a] = 1; votes[b] = (votes[b] || 0) + 1;
    var unique = true;
    Object.keys(other).forEach(function(n) {
      var idx = [], k;
      for (k = 0; k < other[n].refs.length; k++) if (other[n].refs[k] === r) idx.push(k);
      if (idx.length !== 1) { unique = false; return; }
      var s = other[n].saps[idx[0]];
      if (s) votes[s] = (votes[s] || 0) + 1;
    });
    var winner = Object.keys(votes).filter(function(s) { return votes[s] >= 3; })[0];
    if (!unique || !winner) { plan.skipped.push(r + ': SAP IDs ' + a + ' / ' + b + ' cannot be settled by the other sheets, not changed'); return; }
    if (winner !== a) { plan.skipped.push(r + ': the Audit_Log SAP ID (' + a + ') is the odd one out, not changed'); return; }
    plan.sapFixes.push({ ref: r, row: j + 2, from: b, to: a });
  });

  // 3. Repeat Projection %: a sentence in the cell
  [[L, lh, AUDIT_LOG_SHEET, lRef], [D, dh, DASHBOARD_DATA_SHEET, dRef]].forEach(function(t) {
    var vals = col(t[0], t[1], 'Repeat Projection %'), c = t[1].indexOf('Repeat Projection %') + 1;
    vals.forEach(function(x, i) {
      var s = trim(x).replace('%', '').trim();
      if (s === '' || isFinite(Number(s))) return;
      var n = _extractRepeatNumber_(x);
      if (n === null) { plan.skipped.push(t[2] + ' ' + trim(t[3][i]) + ': Repeat Projection % has no percentage in it, not changed'); return; }
      plan.repeatFixes.push({ sheet: t[2], row: i + 2, col: c, ref: trim(t[3][i]), from: String(x), to: n });
    });
  });
  // 4. A cell that holds more than it should: a list of every participant in the Team Member cell, or transcript
  // header text glued onto Team Member / Interaction ID / Direction / Duration. The agent is the clean Team Member
  // that the same SAP ID has on its other audits in Audit_Log, and only when that name is inside the messy cell.
  plan.fieldFixes = [];
  var clean = {};
  col(L, lh, 'Team Member').forEach(function(m, i) {
    m = trim(m);
    var s = trim(lSap[i]);
    if (!s || m.length < 3 || m.length > 60 || m.indexOf(',') !== -1 || /Participant/i.test(m)) return;
    clean[s] = clean[s] || {};
    clean[s][m] = (clean[s][m] || 0) + 1;
  });
  var UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;
  var durTruth = _durationTruthMap_(ss);
  var tzName = (typeof ss.getSpreadsheetTimeZone === 'function') ? ss.getSpreadsheetTimeZone() : Session.getScriptTimeZone();
  [[AUDIT_LOG_SHEET, 'Duration'], [DASHBOARD_DATA_SHEET, 'Transcript Duration'], [TRANSCRIPTS_SHEET, 'Transcript Duration'], [ANALYSIS_SHEET, '']].forEach(function(spec) {
    var sh = ss.getSheetByName(spec[0]);
    if (!sh) return;
    var hd = _trimmedHeaders_(sh), refs = col(sh, hd, 'Audit Ref'), saps = col(sh, hd, 'SAP ID');
    var push = function(field, i, from, to, text) {
      plan.fieldFixes.push({ sheet: spec[0], row: i + 2, col: hd.indexOf(field) + 1, ref: trim(refs[i]), field: field, from: String(from), to: to, text: !!text });
    };
    col(sh, hd, 'Team Member').forEach(function(m, i) {
      m = String(m === null || m === undefined ? '' : m);
      if (m.trim() === '' || (m.length <= 80 && !/Participant\(s\)/i.test(m))) return;
      var sap = trim(saps[i]), best = '', bestN = 0, chosen = _CHOSEN_AGENTS_[trim(refs[i])];
      // an agent the admin chose by hand (only used when that name is inside the messy cell)
      if (chosen && m.toLowerCase().indexOf(chosen.toLowerCase()) !== -1) best = chosen;
      if (!best) Object.keys(clean[sap] || {}).forEach(function(name) {
        if (m.toLowerCase().indexOf(name.toLowerCase()) !== -1 && clean[sap][name] > bestN) { best = name; bestN = clean[sap][name]; }
      });
      if (!best) {
        var cut = m.split(/(?:External|Internal) Participant|Direction:|Duration:/)[0].trim();
        if (cut.length >= 3 && cut.length <= 60 && cut.indexOf(',') === -1 && cut !== m.trim()) best = cut;
      }
      if (best) push('Team Member', i, m, best, false);
      else plan.skipped.push(spec[0] + ' ' + trim(refs[i]) + ': Team Member holds several names and SAP ' + sap + ' has no clean audit to say which one is the agent, not changed');
    });
    col(sh, hd, 'Interaction ID').forEach(function(v, i) {
      v = trim(v);
      if (v.length > 36 && UUID.test(v)) push('Interaction ID', i, v, v.substring(0, 36), false);
    });
    col(sh, hd, 'Direction').forEach(function(v, i) {
      // Only text glued straight onto the word ("InboundInternal Participant(s)..."). "Inbound / Outbound" is a real
      // value that more than a thousand audits hold, so a space after the word means: leave it alone.
      var s = trim(v), m = /^(Inbound|Outbound)(?=[A-Z])/.exec(s);
      if (m) push('Direction', i, s, m[1], false);
    });
    if (spec[1]) col(sh, hd, spec[1]).forEach(function(v, i) {
      if (Object.prototype.toString.call(v) === '[object Date]') {
        if (isNaN(v.getTime())) return;
        var dec = _durationDecision_(v, tzName, durTruth[trim(refs[i])] || '');
        if (dec.to) push(spec[1], i, v, dec.to, true);
        else plan.skipped.push(spec[0] + ' ' + trim(refs[i]) + ': ' + spec[1] + ' holds a time that is not a call length within one day (' + _short_(v, 30) + '): ' + dec.reason + ', not changed');
        return;
      }
      if (typeof v !== 'string') return;
      var s = trim(v), m = /^(\d{1,2}:\d{2}(?::\d{2})?)(?=\D)/.exec(s);
      if (s.length > 8 && m) push(spec[1], i, s, m[1], true);
    });
  });
  return plan;
}

function _logRepairPlan_(plan, apply) {
  Logger.log((apply ? 'APPLYING' : 'PREVIEW (nothing is changed)') + ': rebuild ' + plan.restores.length + ' Audit_Log row(s), fix ' + plan.sapFixes.length +
             ' SAP ID(s), turn ' + plan.repeatFixes.length + ' Repeat % sentence(s) into numbers, clean ' + plan.fieldFixes.length + ' messy cell(s), leave ' + plan.skipped.length + ' alone');
  plan.fieldFixes.slice(0, 40).forEach(function(f) { Logger.log('  CLEAN   ' + f.sheet + ' row ' + f.row + ' ' + f.ref + ' ' + f.field + ': "' + _short_(f.from, 45) + '" -> "' + f.to + '"'); });
  if (plan.fieldFixes.length > 40) Logger.log('  CLEAN   ... and ' + (plan.fieldFixes.length - 40) + ' more');
  plan.restores.forEach(function(r) { Logger.log('  REBUILD Audit_Log row for ' + r.ref + ' (' + r.member + ', SAP ' + r.sap + ') from Dashboard_Data'); });
  plan.sapFixes.forEach(function(f) { Logger.log('  SAP     Dashboard_Data row ' + f.row + ' ' + f.ref + ': ' + f.from + ' -> ' + f.to + ' (the other three sheets say ' + f.to + ')'); });
  plan.repeatFixes.slice(0, 8).forEach(function(f) { Logger.log('  REPEAT  ' + f.sheet + ' row ' + f.row + ' ' + f.ref + ': "' + _short_(f.from, 40) + '" -> ' + f.to); });
  if (plan.repeatFixes.length > 8) Logger.log('  REPEAT  ... and ' + (plan.repeatFixes.length - 8) + ' more');
  plan.skipped.slice(0, 20).forEach(function(s) { Logger.log('  SKIPPED ' + s); });
  if (plan.skipped.length > 20) Logger.log('  SKIPPED ... and ' + (plan.skipped.length - 20) + ' more');
}

// READ-ONLY. Shows what applyDataRepairs() would change.
function previewDataRepairs() {
  requireAdmin_();
  var plan = _planDataRepairs_(getOrCreateSpreadsheet());
  _logRepairPlan_(plan, false);
  return { restores: plan.restores.length, sapFixes: plan.sapFixes.length, repeatFixes: plan.repeatFixes.length, fieldFixes: plan.fieldFixes.length, skipped: plan.skipped };
}

// Makes the changes previewDataRepairs() shows. Safe to run again: it plans from the sheet each time, and stops
// before the script time limit, so run it again if it says it did not finish.
function applyDataRepairs() {
  requireAdmin_();
  var started = new Date().getTime(), budget = 240000;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var result = { restored: 0, sapFixed: 0, repeatFixed: 0, fieldFixed: 0, finished: true };
  try {
    var ss = getOrCreateSpreadsheet();
    var plan = _planDataRepairs_(ss);
    _logRepairPlan_(plan, true);
    var L = ss.getSheetByName(AUDIT_LOG_SHEET), D = ss.getSheetByName(DASHBOARD_DATA_SHEET);
    var dh = _trimmedHeaders_(D), sapCol = dh.indexOf('SAP ID') + 1;

    plan.restores.forEach(function(r) { appendRow(L, r.row); result.restored++; });
    plan.sapFixes.forEach(function(f) { D.getRange(f.row, sapCol).setValue(f.to); result.sapFixed++; });
    for (var i = 0; i < plan.repeatFixes.length; i++) {
      if (new Date().getTime() - started > budget) { result.finished = false; break; }
      var f = plan.repeatFixes[i], sh = (f.sheet === AUDIT_LOG_SHEET) ? L : D;
      var cell = sh.getRange(f.row, f.col);
      cell.setNote('Original text before it was turned into a number:\n' + f.from.substring(0, 4000));
      cell.setValue(f.to);
      result.repeatFixed++;
    }
    var sheetOf = function(name) { return name === AUDIT_LOG_SHEET ? L : (name === DASHBOARD_DATA_SHEET ? D : ss.getSheetByName(name)); };
    for (var k = 0; k < plan.fieldFixes.length; k++) {
      if (new Date().getTime() - started > budget) { result.finished = false; break; }
      var g = plan.fieldFixes[k], cell2 = sheetOf(g.sheet).getRange(g.row, g.col);
      cell2.setNote('Original value before it was cleaned:\n' + g.from.substring(0, 4000));
      if (g.text) cell2.setNumberFormat('@');
      cell2.setValue(g.to);
      result.fieldFixed++;
    }
    if (typeof invalidateAuditLogCache_ === 'function')  invalidateAuditLogCache_();
    if (typeof invalidateDashboardCache_ === 'function') invalidateDashboardCache_();
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
  Logger.log('applyDataRepairs: rebuilt ' + result.restored + ' row(s), fixed ' + result.sapFixed + ' SAP ID(s), fixed ' + result.repeatFixed + ' Repeat % cell(s), cleaned ' + result.fieldFixed + ' messy cell(s)' +
             (result.finished ? '. Done.' : '. NOT FINISHED (time limit), run applyDataRepairs again.'));
  return result;
}

// READ-ONLY. Three reports about Audit_Log, to decide what (if anything) should be repaired:
//   Duration       how many call lengths were turned into a time of day by Sheets (applyDataRepairs turns them back into text)
//   Email Status   where the "Not Sent" audits are: by month, by type, by observer, and whether any has recipients
//   Issue Resolved the answers the dashboard cannot read as a verdict: how they look and how they usually start
// READ-ONLY. Shows what the AI summaries are told about Issue Resolved and Repeat %, computed from the real sheets with the code that is in the
// editor right now, next to what Audit_Log alone says. When the dashboard text disagrees with the tiles, this says whether the code is old (the
// web app still runs an older version: create a new version and point the deployment at it) or the numbers themselves differ.
function diagnoseAnalyticsContext() {
  requireAdmin_();
  var ss = getOrCreateSpreadsheet(), say = function(s) { Logger.log(s); };
  var log = readAuditLog_(), dash = readDashboardSheet_();
  var byRef = {}; dash.forEach(function(r) { var k = String(r['Audit Ref'] || '').trim(); if (k) byRef[k] = r; });
  // the rows as the dashboard page builds them: the Audit_Log value, else the Dashboard_Data value
  var merged = log.map(function(r) {
    var d = byRef[String(r['Audit Ref'] || '').trim()] || {};
    return { 'Team Member': r['Team Member'], 'Issue Resolved': r['Issue Resolved'] || d['Issue Resolved'] || '', 'Repeat Projection %': r['Repeat Projection %'] || d['Repeat Projection %'] || '', 'Critical Flags': d['Critical Flags'] || '' };
  });
  var count = function(rows) { var c = { yes: 0, partial: 0, no: 0, unknown: 0 }; rows.forEach(function(r) { c[_scResolution_(r['Issue Resolved'])]++; }); return c; };
  var a = count(log), b = count(merged);
  say('=== What the AI summaries are told (code in the editor now) ===');
  say('  Audit_Log alone (' + log.length + ' audits): Resolved ' + a.yes + ', Partially ' + a.partial + ', Unresolved ' + a.no + ', Not recorded ' + a.unknown);
  say('  as the dashboard page builds the rows (' + merged.length + '): Resolved ' + b.yes + ', Partially ' + b.partial + ', Unresolved ' + b.no + ', Not recorded ' + b.unknown);
  var old = 0; merged.forEach(function(r) { if (String(r['Issue Resolved'] || '').toLowerCase().indexOf('yes') !== -1) old++; });
  say('  the OLD way (any cell that contains the letters yes): ' + old + '   <- if the dashboard summary still says this number, the web app runs old code');
  var text = buildAnalyticsContext_(merged, {}, 'repeats').split('\n').filter(function(l) { return /^(TOTAL INTERACTIONS|- Issues Resolved|- Issue Resolved, all|- Average Repeat Risk)/.test(l); });
  text.forEach(function(l) { say('  context line: ' + l); });
  say('=== end: copy everything above into the chat ===');
  return { audits: log.length, auditLog: a, merged: b, oldWay: old, contextLines: text };
}

// A copy of the data spreadsheet, made before a repair writes thousands of cells. Admin only. The copy goes in the same folder as the
// original, is named BACKUP with the date and time, and the original is not touched. It is a plain Drive copy, so it holds every sheet,
// note and format as they are now. If Drive refuses, the error is raised on purpose: do not go on to repair without a backup.
function backupSpreadsheet() {
  requireAdmin_();
  var ss = getOrCreateSpreadsheet();
  var tz = Session.getScriptTimeZone();
  var name = ss.getName() + ' - BACKUP ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH-mm');
  var file = DriveApp.getFileById(ss.getId());
  var parents = file.getParents();
  var copy = parents.hasNext() ? file.makeCopy(name, parents.next()) : file.makeCopy(name);
  Logger.log('Backup made: ' + copy.getName() + '  ' + copy.getUrl());
  return { success: true, name: copy.getName(), url: copy.getUrl(), id: copy.getId() };
}

function diagnoseDashboardGaps() {
  requireAdmin_();
  var ss = getOrCreateSpreadsheet(), out = [], say = function(s) { out.push(s); Logger.log(s); };
  var tz = (typeof ss.getSpreadsheetTimeZone === 'function') ? ss.getSpreadsheetTimeZone() : Session.getScriptTimeZone();
  var trim = function(x) { return String(x === null || x === undefined ? '' : x).trim(); };
  var L = ss.getSheetByName(AUDIT_LOG_SHEET), lh = _trimmedHeaders_(L);
  var col = function(n) { return _colByName_(L, lh, n) || []; };
  var refs = col('Audit Ref').map(trim);
  var result = {};

  say('=== A. Duration cells that Sheets turned into a time of day ===');
  var durTruth = _durationTruthMap_(ss);
  // The length each saved transcript says, by the seconds Sheets would read from it: a cell that holds a DIFFERENT audit's length was probably
  // left over from an earlier paste (the form kept the first transcript's duration when a second was pasted).
  var truthSecs = {};
  Object.keys(durTruth).forEach(function(r) { var k = _sheetsSecondsOfText_(durTruth[r]); if (!truthSecs[k]) truthSecs[k] = r; });
  var dur = { leftStale: 0, fixable: { audit: 0, dash: 0, trans: 0 }, byTranscript: 0, byClock: 0, left: { audit: 0, dash: 0, trans: 0 }, leftDays: {}, samples: [], leftSamples: [], leftNoTruth: 0, leftMismatch: 0 };
  [[AUDIT_LOG_SHEET, 'Duration', 'audit'], [DASHBOARD_DATA_SHEET, 'Transcript Duration', 'dash'], [TRANSCRIPTS_SHEET, 'Transcript Duration', 'trans']].forEach(function(s) {
    var sh = ss.getSheetByName(s[0]), hd = _trimmedHeaders_(sh), vals = _colByName_(sh, hd, s[1]) || [], rf = (_colByName_(sh, hd, 'Audit Ref') || []).map(trim);
    vals.forEach(function(v, i) {
      if (Object.prototype.toString.call(v) !== '[object Date]' || isNaN(v.getTime())) return;
      var dec = _durationDecision_(v, tz, durTruth[rf[i]] || '');
      if (dec.to) {
        dur.fixable[s[2]]++; if (dec.how === 'transcript') dur.byTranscript++; else dur.byClock++;
        if (dur.samples.length < 8) dur.samples.push(s[0] + ' ' + rf[i] + ': ' + _short_(v, 34) + ' -> ' + dec.to + ' (' + (dec.how === 'transcript' ? 'from the transcript' : 'three-part length') + ')');
      } else {
        dur.left[s[2]]++;
        var other = '';
        if (durTruth[rf[i]]) { dur.leftMismatch++; var os = truthSecs[_clockSecondsOf_(v, tz)]; if (os && os !== rf[i]) { other = os; dur.leftStale++; } } else dur.leftNoTruth++;
        var day = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        dur.leftDays[day] = (dur.leftDays[day] || 0) + 1;
        if (dur.leftSamples.length < 8) dur.leftSamples.push(s[0] + ' ' + (rf[i] || ('row ' + (i + 2))) + ': ' + Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm:ss') + (durTruth[rf[i]] ? ' (transcript says ' + durTruth[rf[i]] + (other ? '; the cell equals the length of ' + other + ')' : ')') : ' (no length in the saved transcript)'));
      }
    });
  });
  var leftTotal = dur.left.audit + dur.left.dash + dur.left.trans;
  say('  can be written back as text: Audit_Log ' + dur.fixable.audit + ', Dashboard_Data ' + dur.fixable.dash + ', Transcripts ' + dur.fixable.trans + ' (' + dur.byTranscript + ' from the length written in the saved transcript, ' + dur.byClock + ' three-part lengths such as 0:11:39)');
  dur.samples.forEach(function(x) { say('    ' + x); });
  say('  left alone (cannot be confirmed): ' + leftTotal + ' (Audit_Log ' + dur.left.audit + ', Dashboard_Data ' + dur.left.dash + ', Transcripts ' + dur.left.trans + '); the saved transcript has no length for ' + dur.leftNoTruth + ' and says something different for ' + dur.leftMismatch + ' (for ' + dur.leftStale + ' of those the cell equals the length of a DIFFERENT audit, as if left over from an earlier paste)');
  if (leftTotal) {
    say('  ... by the day the value falls on (top 6): ' + Object.keys(dur.leftDays).sort(function(x, y) { return dur.leftDays[y] - dur.leftDays[x]; }).slice(0, 6).map(function(k) { return k + ' ' + dur.leftDays[k]; }).join(', '));
    dur.leftSamples.forEach(function(x) { say('    ' + x); });
  }
  result.durationAudit = dur.fixable.audit; result.durationDash = dur.fixable.dash; result.durationTranscripts = dur.fixable.trans; result.durationNotDay = leftTotal;
  result.durationFromTranscript = dur.byTranscript; result.durationLeftNoTruth = dur.leftNoTruth; result.durationLeftMismatch = dur.leftMismatch; result.durationLeftStale = dur.leftStale;

  say('=== B. Email Status ===');
  var st = col('Email Status').map(trim), at = col('Submitted At'), rec = col('Recipients').map(trim), type = col('Analysis Type').map(trim), obs = col('Observer Name').map(trim);
  var byStatus = {}; st.forEach(function(s) { byStatus[s || '(blank)'] = (byStatus[s || '(blank)'] || 0) + 1; });
  say('  by status: ' + Object.keys(byStatus).map(function(k) { return k + ' ' + byStatus[k]; }).join(', '));
  var ns = [], byMonth = {}, byType = {}, byObs = {}, withRecipients = 0;
  st.forEach(function(s, i) {
    if (s !== 'Not Sent') return;
    ns.push(i);
    var d = at[i], m = 'unknown';
    if (Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d.getTime())) m = Utilities.formatDate(d, tz, 'yyyy-MM');
    else { var p = Date.parse(String(d)); if (!isNaN(p)) m = Utilities.formatDate(new Date(p), tz, 'yyyy-MM'); }
    byMonth[m] = (byMonth[m] || 0) + 1;
    byType[type[i] || '(blank)'] = (byType[type[i] || '(blank)'] || 0) + 1;
    byObs[obs[i] || '(blank)'] = (byObs[obs[i] || '(blank)'] || 0) + 1;
    if (rec[i]) withRecipients++;
  });
  say('  "Not Sent": ' + ns.length + ' audit(s); ' + withRecipients + ' of them have recipients written (that would be odd)');
  say('  by month: ' + Object.keys(byMonth).sort().map(function(k) { return k + ' ' + byMonth[k]; }).join(', '));
  say('  by type: ' + Object.keys(byType).map(function(k) { return k + ' ' + byType[k]; }).join(', '));
  say('  by observer (top 8): ' + Object.keys(byObs).sort(function(a, b) { return byObs[b] - byObs[a]; }).slice(0, 8).map(function(k) { return k + ' ' + byObs[k]; }).join(', '));
  result.notSent = ns.length; result.notSentWithRecipients = withRecipients; result.notSentByMonth = byMonth;

  say('=== C. Issue Resolved answers with no clear verdict ===');
  var res = col('Issue Resolved').map(function(x) { return String(x === null || x === undefined ? '' : x); });
  var cls = { blank: 0, verdict: 0, html: 0, short: 0, sentence: 0 }, ex = { html: [], short: [], sentence: [] }, lead = {};
  res.forEach(function(v, i) {
    if (!v.trim()) { cls.blank++; return; }
    if (_scResolution_(v) !== 'unknown') { cls.verdict++; return; }
    var kind = (/[<>]/.test(v) || /="/.test(v)) ? 'html' : (v.trim().length < 20 ? 'short' : 'sentence');
    cls[kind]++;
    if (ex[kind].length < 6) ex[kind].push(refs[i] + ': "' + _short_(v, 70) + '"');
    if (kind === 'sentence') { var k = v.trim().toLowerCase().split(/\s+/).slice(0, 3).join(' '); lead[k] = (lead[k] || 0) + 1; }
  });
  say('  blank ' + cls.blank + ', a clear verdict ' + cls.verdict + ', no verdict: ' + (cls.html + cls.short + cls.sentence) + ' (html-like text ' + cls.html + ', very short ' + cls.short + ', a sentence ' + cls.sentence + ')');
  ['html', 'short', 'sentence'].forEach(function(k) { ex[k].forEach(function(x) { say('    ' + k + ' ' + x); }); });
  say('  how the sentences start (top 15): ' + Object.keys(lead).sort(function(a, b) { return lead[b] - lead[a]; }).slice(0, 15).map(function(k) { return '"' + k + '" ' + lead[k]; }).join(', '));
  result.issueBlank = cls.blank; result.issueVerdict = cls.verdict; result.issueHtml = cls.html; result.issueShort = cls.short; result.issueSentence = cls.sentence;
  say('=== end: copy everything above into the chat ===');
  result.lines = out;
  return result;
}
