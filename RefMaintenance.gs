/**
 * RefMaintenance.gs
 * One-off admin tools for audit references that two audits already share.
 *
 * Before the NH and TI Analyzers got separate reference prefixes, both scripts
 * handed out NHA-YYYYMMDD-0001, -0002, ... against ONE spreadsheet, so
 * Audit_Log holds pairs of rows with the same reference. An email link carries
 * only the reference and opens the first row, so the later row's owner is
 * shown the earlier row's agent and evaluation.
 *
 * Run from the Apps Script editor, in this order:
 *   1. reportDuplicateAuditRefs()        read-only; lists what would change
 *   2. applyDuplicateAuditRefRenames()   makes the change
 *
 * The earliest row keeps its reference. Every later row gets a new unique one
 * (the old reference plus -D2, -D3, ...) in Audit_Log, Dashboard_Data,
 * Transcripts and Analysis. Each change is recorded in Ref_Rename_Log, so it
 * can be reversed by hand.
 *
 * An evaluation link that was emailed for a renamed row stops working
 * ("not found") instead of showing another agent's evaluation. The analyst can
 * resend it from the app once the row has its own reference.
 */

var REF_RENAME_LOG_SHEET   = 'Ref_Rename_Log';
var REF_RENAME_LOG_HEADERS = ['Renamed At', 'Old Ref', 'New Ref', 'SAP ID', 'Interaction ID', 'Sheets Updated', 'Note'];
// Sheets that carry an 'Audit Ref' and a 'SAP ID' column per audit. A function,
// not a top-level array, so this file never depends on Config.gs having loaded first.
function _refMaintenanceSheets_() {
  return [AUDIT_LOG_SHEET, DASHBOARD_DATA_SHEET, TRANSCRIPTS_SHEET, ANALYSIS_SHEET];
}

// Pure planning step. rows: [{ ref, sap, interactionId }] in Audit_Log order
// (oldest first). Returns one entry per row that must get a new reference:
// [{ index, ref, newRef, sap, interactionId, manual }].
// The earliest row of each reference is kept. A later row whose SAP ID equals
// another row's with the same reference cannot be told apart in every sheet
// (the Analysis sheet has no Interaction ID), so it is flagged manual and left
// unchanged.
function _planDuplicateRefRenames(rows) {
  var byRef = {};
  var used  = {};
  rows.forEach(function(r, i) {
    var ref = (r.ref || '').toString().trim();
    used[ref] = true;
    if (!ref) return;
    (byRef[ref] = byRef[ref] || []).push(i);
  });

  var plan = [];
  Object.keys(byRef).forEach(function(ref) {
    var idxs = byRef[ref];
    if (idxs.length < 2) return;
    var n = 2;
    for (var k = 1; k < idxs.length; k++) {
      var i   = idxs[k];
      var sap = (rows[i].sap || '').toString().trim();
      var clash = !sap || idxs.some(function(j) {
        return j !== i && (rows[j].sap || '').toString().trim() === sap;
      });
      var entry = { index: i, ref: ref, newRef: '', sap: sap,
                    interactionId: (rows[i].interactionId || '').toString().trim(), manual: clash };
      if (!clash) {
        var candidate;
        do { candidate = ref + '-D' + n++; } while (used[candidate]);
        used[candidate] = true;
        entry.newRef = candidate;
      }
      plan.push(entry);
    }
  });
  plan.sort(function(a, b) { return a.index - b.index; });
  return plan;
}

function _readRowsForRefMaintenance_(ss) {
  var sheet   = getOrCreateSheet(ss, AUDIT_LOG_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { sheet: sheet, rows: [], refCol: -1 };
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var refCol = headers.indexOf('Audit Ref') + 1;
  var sapCol = headers.indexOf('SAP ID') + 1;
  var idCol  = headers.indexOf('Interaction ID') + 1;
  if (refCol < 1) throw new Error('Audit_Log has no "Audit Ref" column.');
  var n    = lastRow - 1;
  var refs = sheet.getRange(2, refCol, n, 1).getValues();
  var saps = sapCol > 0 ? sheet.getRange(2, sapCol, n, 1).getValues() : null;
  var ids  = idCol  > 0 ? sheet.getRange(2, idCol,  n, 1).getValues() : null;
  var rows = [];
  for (var i = 0; i < n; i++) {
    rows.push({ ref: refs[i][0], sap: saps ? saps[i][0] : '', interactionId: ids ? ids[i][0] : '' });
  }
  return { sheet: sheet, rows: rows, refCol: refCol };
}

// STEP 1 — read-only. Lists every reference used by more than one audit and
// what applyDuplicateAuditRefRenames() would do. Changes nothing.
function reportDuplicateAuditRefs() {
  requireAdmin_();
  var ss   = getOrCreateSpreadsheet();
  var data = _readRowsForRefMaintenance_(ss);
  var plan = _planDuplicateRefRenames(data.rows);
  var dupRefs = {};
  plan.forEach(function(p) { dupRefs[p.ref] = true; });
  var manual = plan.filter(function(p) { return p.manual; });
  var report = {
    duplicateRefs: Object.keys(dupRefs).length,
    renames:       plan.length - manual.length,
    manual:        manual.length,
    plan:          plan
  };
  Logger.log('Duplicate audit references: ' + report.duplicateRefs +
             ' | rows that would be renamed: ' + report.renames +
             ' | rows needing manual review: ' + report.manual);
  plan.forEach(function(p) {
    Logger.log((p.manual ? 'MANUAL  ' : 'RENAME  ') + p.ref + (p.manual ? '' : ' -> ' + p.newRef) +
               '  (sheet row ' + (p.index + 2) + ', SAP ' + p.sap + ', interaction ' + p.interactionId + ')');
  });
  return report;
}

// Reads ONLY the Audit Ref, SAP ID and Interaction ID columns of a sheet, once.
// The Transcripts and Analysis sheets hold a whole transcript or evaluation in
// every row, so the rest of each row must never be loaded. Returns null if the
// sheet is missing or has no rows or no Audit Ref / SAP ID column.
function _loadRefIndex_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var refCol = headers.indexOf('Audit Ref') + 1;
  var sapCol = headers.indexOf('SAP ID') + 1;
  var idCol  = headers.indexOf('Interaction ID') + 1;
  if (refCol < 1 || sapCol < 1) return null;
  var n = sheet.getLastRow() - 1;
  return {
    sheet:  sheet,
    refCol: refCol,
    refs:   sheet.getRange(2, refCol, n, 1).getValues(),
    saps:   sheet.getRange(2, sapCol, n, 1).getValues(),
    ids:    idCol > 0 ? sheet.getRange(2, idCol, n, 1).getValues() : null
  };
}

// Position (0-based) in the loaded index of the one row for this audit, or -1
// if there is none or more than one. Matches reference + SAP ID, and also the
// Interaction ID where the sheet has that column and the audit has one.
function _findInRefIndex_(idx, oldRef, sap, interactionId) {
  var hit = -1, count = 0;
  for (var i = 0; i < idx.refs.length; i++) {
    if ((idx.refs[i][0] || '').toString().trim() !== oldRef) continue;
    if ((idx.saps[i][0] || '').toString().trim() !== sap) continue;
    if (idx.ids && interactionId && (idx.ids[i][0] || '').toString().trim() !== interactionId) continue;
    hit = i; count++;
  }
  return count === 1 ? hit : -1;
}

// STEP 2 — makes the change described by reportDuplicateAuditRefs().
// Safe to run more than once: after the first run no reference is shared.
function applyDuplicateAuditRefRenames() {
  requireAdmin_();
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss   = getOrCreateSpreadsheet();
    var data = _readRowsForRefMaintenance_(ss);
    var plan = _planDuplicateRefRenames(data.rows);
    var result = { renamed: 0, manual: 0 };

    var logSheet = null;
    var cache = CacheService.getScriptCache();

    // Load each other sheet's three small columns once, not once per rename.
    var indexes = {};
    _refMaintenanceSheets_().forEach(function(name) {
      if (name !== AUDIT_LOG_SHEET) indexes[name] = _loadRefIndex_(ss.getSheetByName(name));
    });

    plan.forEach(function(p) {
      if (p.manual) { result.manual++; return; }

      // Audit_Log: the row is known by position.
      data.sheet.getRange(p.index + 2, data.refCol).setValue(p.newRef);
      var touched = [AUDIT_LOG_SHEET];

      // The other sheets: this audit's row by reference + SAP ID (+ Interaction ID).
      Object.keys(indexes).forEach(function(name) {
        var idx = indexes[name];
        if (!idx) return;
        var i = _findInRefIndex_(idx, p.ref, p.sap, p.interactionId);
        if (i < 0) return;
        idx.sheet.getRange(i + 2, idx.refCol).setValue(p.newRef);
        idx.refs[i][0] = p.newRef;   // keep the in-memory copy current for later pairs
        touched.push(name);
      });

      if (!logSheet) {
        logSheet = getOrCreateSheet(ss, REF_RENAME_LOG_SHEET);
        if (logSheet.getLastRow() === 0) logSheet.appendRow(REF_RENAME_LOG_HEADERS);
      }
      logSheet.appendRow([new Date(), p.ref, p.newRef, p.sap, p.interactionId, touched.join(', '), '']);

      // The evaluation page caches by reference for up to 6 hours.
      try { cache.remove('ev2_' + Utilities.base64Encode(p.ref).substring(0, 200)); } catch (e) {}
      result.renamed++;
    });

    if (typeof invalidateAuditLogCache_ === 'function')  invalidateAuditLogCache_();
    if (typeof invalidateDashboardCache_ === 'function') invalidateDashboardCache_();
    Logger.log('applyDuplicateAuditRefRenames: renamed ' + result.renamed + ', left ' + result.manual + ' for manual review.');
    return result;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
