/**
 * SheetService.gs
 * Handles all Google Sheets read/write operations.
 */

function getOrCreateSpreadsheet() {
  return openSheetWithRetry(MAIN_SPREADSHEET_ID, 'Main Analyzer Spreadsheet');
}

// ── Open a Spreadsheet with retry on transient permission-like errors ─────────
// SpreadsheetApp.openById() can throw "You do not have permission to access
// the requested document" under heavy concurrent access even for a file's own
// owner — this deployment runs as "Execute as: Me", so a genuine per-user
// sharing gap is not the likely cause of these reports (see Error_Log /
// Action Registry #148). Retrying briefly clears the transient case; a truly
// broken share still surfaces the same error after these attempts, just
// slightly delayed, so this changes nothing for a real permissions problem.
function openSheetWithRetry(id, label) {
  var lastErr;
  for (var attempt = 1; attempt <= 3; attempt++) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      lastErr = e;
      var msg = e.toString();
      var isPermissionLike = msg.indexOf('permission') !== -1 || msg.indexOf('do not have access') !== -1;
      if (!isPermissionLike || attempt === 3) break;
      Utilities.sleep(attempt * 750); // 750ms, then 1500ms
    }
  }
  throw new Error('Could not open "' + label + '" (' + id + '): ' + lastErr);
}

function getOrCreateSheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  return sheet;
}

function ensureHeaders(sheet, headers) {
  // Fast path — if headers already exist, skip the lock entirely.
  if (sheet.getLastRow() > 0) return;

  // Double-checked locking: two concurrent executions can both see getLastRow()===0
  // and both enter the slow path. Acquiring the lock then re-checking prevents both
  // from appending a header row, which would corrupt every subsequent cache read
  // (row 2 would be a second header instead of data).
  var _ehLock = LockService.getScriptLock();
  try {
    _ehLock.waitLock(5000);
    if (sheet.getLastRow() === 0) { // re-check after acquiring lock
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#4B286D')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  } catch(le) {
    Logger.log('ensureHeaders: lock failed — ' + le);
  } finally {
    try { _ehLock.releaseLock(); } catch(le2) {}
  }
}

// Durations like "51:50" (51 min 50 sec) are written as plain strings, and Sheets
// silently reinterprets them as time values. "51:50" is not a valid time of day,
// so it becomes the duration 51 HOURS 50 minutes and displays as "51:50:00";
// "23:14" IS a valid time of day, so it silently becomes 11:14 PM. Either way the
// cell no longer means what was written, and every average or total built on the
// Duration column is meaningless. Leading the value with an apostrophe forces
// Sheets to keep it as text.
//
// Deliberately narrow: only m:ss / mm:ss / h:mm:ss shapes. Timestamps are written
// as real Date objects and are unaffected.
var _DURATION_SHAPE = /^\d{1,3}:[0-5]\d(:[0-5]\d)?$/;

function appendRow(sheet, rowData) {
  var safeRow = rowData.map(function(v) {
    if (typeof v !== 'string') return v;
    if (/^[=+\-@]/.test(v))        return "'" + v;   // formula injection
    if (_DURATION_SHAPE.test(v))   return "'" + v;   // stop time coercion
    return v;
  });
  sheet.appendRow(safeRow);
}

function getSpreadsheetUrl() {
  requireAdmin_();
  return getOrCreateSpreadsheet().getUrl();
}

// ── Backfill team-member details from prior audit history ─────────────────────
// If the client-side roster autofill didn't populate Team Leader / Ops
// Manager / LOB / Locale / VTID (e.g. the external Roster/AT-Data lookup
// failed or timed out), this SAP ID's most recent PRIOR entry in Audit_Log —
// which lives in the analyzer's own spreadsheet, not an external file — is a
// reasonable substitute for the fields still missing. Returns null if the
// SAP ID has no prior audit. Data may be stale if the team member was
// reassigned since their last audit; callers should warn when it's used
// rather than applying it silently.
function getTeamDetailsFromAuditHistory_(sapId) {
  try {
    if (!sapId) return null;
    var ss       = getOrCreateSpreadsheet();
    var sheet    = getOrCreateSheet(ss, AUDIT_LOG_SHEET);
    var lastRow  = sheet.getLastRow();
    if (lastRow < 2) return null;

    var headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var sapCol   = headers.indexOf('SAP ID');
    if (sapCol === -1) return null;
    var tlCol    = headers.indexOf('Team Leader');
    var omCol    = headers.indexOf('Operations Manager');
    var lobCol   = headers.indexOf('Line of Business');
    var locCol   = headers.indexOf('Locale');
    var vtidCol  = headers.indexOf('VTID');

    var data   = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var target = sapId.toString().trim();
    for (var i = data.length - 1; i >= 0; i--) { // scan newest-first
      var row = data[i];
      if ((row[sapCol] || '').toString().trim() === target) {
        return {
          teamLeader:     tlCol   !== -1 ? (row[tlCol]   || '').toString() : '',
          opsManager:     omCol   !== -1 ? (row[omCol]   || '').toString() : '',
          lineOfBusiness: lobCol  !== -1 ? (row[lobCol]  || '').toString() : '',
          locale:         locCol  !== -1 ? (row[locCol]  || '').toString() : '',
          vtid:           vtidCol !== -1 ? (row[vtidCol] || '').toString() : ''
        };
      }
    }
    return null;
  } catch (e) {
    Logger.log('getTeamDetailsFromAuditHistory_ error: ' + e);
    return null;
  }
}

// ── Write a submission failure to the Error_Log tab ────────────────────────────
// Apps Script execution logs require Stackdriver/GCP Cloud Logging access,
// which needs a non-default GCP project — not always available to whoever is
// troubleshooting. Writing straight to the spreadsheet the admin already has
// open gives instant visibility with no extra setup. Never throws — a logging
// failure must never mask or replace the original error being reported.
function logSubmissionError_(functionName, formData, e) {
  try {
    var ss    = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss, ERROR_LOG_SHEET);
    ensureHeaders(sheet, ERROR_LOG_HEADERS);
    appendRow(sheet, [
      new Date(),
      functionName,
      (formData && formData.observerName)   || '',
      (formData && formData.sapId)          || '',
      (formData && formData.interactionId)  || '',
      e ? e.toString() : '',
      e && e.stack ? e.stack.toString().substring(0, 2000) : ''
    ]);
  } catch (logErr) {
    Logger.log('logSubmissionError_ failed: ' + logErr);
  }
}

// ── Cache: per-agent evaluations ──────────────────────────────────────────────
// A cached evaluation belongs to ONE agent's audit. A Repeats & Transfer call
// has several agents on the same Interaction ID and each is audited on their
// own, so the cache is keyed by Interaction ID + analysis type + SAP ID.
//
// Cache sheet layout: A Interaction ID, B Analysis Type, C Timestamp,
// D HTML Result, E SAP ID. Rows written before column E existed have a blank
// owner. A blank-owner row belongs to the agent whose audit CREATED it, which
// is the EARLIEST Audit_Log row for that Interaction ID and analysis type
// (_legacyCacheOwnerSap). It is served to, and replaced by, that agent only.
var _RESULT_CS_PREFIX = 'html_result_v1_';
var _RESULT_CS_TTL    = 6 * 60 * 60; // 6 hours
var _CACHE_SAP_COL    = 5;           // column E
var _CACHE_CHUNK_SIZE = 48000;       // safe margin under Google Sheets 50k cell limit

function _normalizeAnalysisType(label) {
  return (label || '').toString().toLowerCase().indexOf('sales') !== -1 ? 'sales' : 'repeats';
}

// Audit_Log rows for one Interaction ID, oldest first: [{ ref, sap, type }].
// Finds them with TextFinder on the Interaction ID column and reads only the
// matching rows, so the cost does not grow with the size of the sheet.
function _auditRowsForInteraction(interactionId) {
  var wantId = (interactionId || '').toString().trim();
  if (!wantId) return [];
  var sheet   = getOrCreateSheet(getOrCreateSpreadsheet(), AUDIT_LOG_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var refCol  = headers.indexOf('Audit Ref');
  var idCol   = headers.indexOf('Interaction ID');
  var sapCol  = headers.indexOf('SAP ID');
  var typeCol = headers.indexOf('Analysis Type');
  if (refCol < 0 || idCol < 0) return [];
  var hits = sheet.getRange(1, idCol + 1, lastRow, 1).createTextFinder(wantId).matchEntireCell(true).findAll();
  var rowNums = hits.map(function(h) { return h.getRow(); })
                    .filter(function(n) { return n >= 2; })
                    .sort(function(a, b) { return a - b; });
  var out = [];
  for (var i = 0; i < rowNums.length; i++) {
    var v = sheet.getRange(rowNums[i], 1, 1, lastCol).getValues()[0];
    if ((v[idCol] || '').toString().trim() !== wantId) continue;
    out.push({
      ref:  (v[refCol] || '').toString().trim(),
      sap:  sapCol  >= 0 ? (v[sapCol] || '').toString().trim() : '',
      type: _normalizeAnalysisType(typeCol >= 0 ? v[typeCol] : '')
    });
  }
  return out;
}

// Ref of THIS agent's audit of this Interaction ID, or ''. analysisType may be
// 'sales'/'repeats' or the stored label. A blank sapId matches any agent.
function _findAuditRefForAgent(interactionId, analysisType, sapId) {
  try {
    var wantSap  = (sapId || '').toString().trim();
    var wantType = _normalizeAnalysisType(analysisType);
    var rows = _auditRowsForInteraction(interactionId);
    for (var i = 0; i < rows.length; i++) {
      if (wantSap && rows[i].sap !== wantSap) continue;
      if (rows[i].type !== wantType) continue;
      return rows[i].ref;
    }
    return '';
  } catch (e) {
    Logger.log('_findAuditRefForAgent error: ' + e);
    return '';
  }
}

// The Audit Ref a failed first attempt left in the Transcripts sheet for THIS agent's
// audit of this interaction, or ''. submitTranscript writes that row before the AI
// runs, so when the evaluation got cached but the audit record did not, this is the
// reference the cached report already carries; reusing it keeps the report and the
// record in agreement. Newest match wins. Only four narrow columns are read, never
// the transcript text.
function findOrphanTranscriptRef_(interactionId, sapId, analysisLabel) {
  try {
    var id = (interactionId || '').toString().trim();
    if (!id) return '';
    var wantSap = (sapId || '').toString().trim();
    var sheet = getOrCreateSpreadsheet().getSheetByName(TRANSCRIPTS_SHEET);
    if (!sheet) return '';
    var last = sheet.getLastRow();
    if (last < 2) return '';
    var colOf = function(h) { return TRANSCRIPTS_HEADERS.indexOf(h) + 1; };
    var read  = function(h) { return sheet.getRange(2, colOf(h), last - 1, 1).getValues(); };
    var refs = read('Audit Ref'), saps = read('SAP ID'), ids = read('Interaction ID'), types = read('Analysis Type');
    for (var i = refs.length - 1; i >= 0; i--) {
      if ((ids[i][0] || '').toString().trim() !== id) continue;
      if (wantSap && (saps[i][0] || '').toString().trim() !== wantSap) continue;
      if (analysisLabel && (types[i][0] || '').toString().trim() !== analysisLabel) continue;
      var ref = (refs[i][0] || '').toString().trim();
      if (ref) return ref;
    }
    return '';
  } catch (e) {
    Logger.log('findOrphanTranscriptRef_ error: ' + e);
    return '';
  }
}

// True when the sheet already has a row whose first column (Audit Ref) is `ref`.
// Used when completing the records of an audit, so a record that WAS written the
// first time is not written twice.
function _sheetHasRef_(sheet, ref) {
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var vals = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if ((vals[i][0] || '').toString().trim() === ref) return true;
  }
  return false;
}

// SAP ID of the agent whose audit created a blank-owner (pre-column-E) cache
// row for this Interaction ID and type: the earliest matching audit. '' if none.
function _legacyCacheOwnerSap(interactionId, analysisType) {
  try {
    var wantType = _normalizeAnalysisType(analysisType);
    var rows = _auditRowsForInteraction(interactionId);
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].type === wantType) return rows[i].sap;
    }
    return '';
  } catch (e) {
    Logger.log('_legacyCacheOwnerSap error: ' + e);
    return '';
  }
}

// ── Cache: look up this agent's previously processed evaluation ───────────────
// Returns the cached HTML string, or null if not found. A blank sapId matches
// any owner (used only when the audit itself has no SAP ID).
function findCachedResult_(interactionId, analysisType, sapId) {
  if (!interactionId) return null;
  try {
    var target    = interactionId.trim();
    var typeLower = (analysisType || '').trim().toLowerCase();
    var sap       = (sapId || '').toString().trim();

    // ── Layer 1: CacheService (O(1), no sheet read) ───────────────────────────
    var csKey    = _RESULT_CS_PREFIX + target + '_' + typeLower + '_' + sap;
    var csResult = CacheService.getScriptCache().get(csKey);
    if (csResult) {
      Logger.log('Cache HIT (CacheService) for: ' + target);
      return csResult;
    }

    // ── Layer 2: Cache sheet TextFinder (slower, but persistent) ─────────────
    var ss    = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss, CACHE_SHEET);
    ensureHeaders(sheet, CACHE_HEADERS);
    if (sheet.getLastRow() < 2) return null;

    // Read each matching row once; the primary row and its continuation chunks
    // are then picked from memory.
    var matches = sheet.getRange('A:A').createTextFinder(target).matchEntireCell(true).findAll();
    var rows = [];
    for (var i = 0; i < matches.length; i++) {
      var row = matches[i].getRow();
      if (row < 2) continue;
      var d = sheet.getRange(row, 1, 1, _CACHE_SAP_COL).getValues()[0];
      rows.push({
        type:  (d[1] || '').toString().trim().toLowerCase(),
        html:  (d[3] || '').toString(),
        owner: (d[_CACHE_SAP_COL - 1] || '').toString().trim()
      });
    }

    var primary = null;
    for (var p = 0; p < rows.length; p++) {
      if (rows[p].type === typeLower && rows[p].html && (!sap || rows[p].owner === sap)) { primary = rows[p]; break; }
    }
    if (!primary && sap) {
      for (var q = 0; q < rows.length; q++) {
        if (rows[q].type === typeLower && rows[q].html && rows[q].owner === '') {
          if (_legacyCacheOwnerSap(target, typeLower) === sap) primary = rows[q];
          break;
        }
      }
    }
    if (!primary) {
      Logger.log('Cache MISS for: ' + target);
      return null;
    }

    // Assemble continuation chunks (sales_2, sales_3, repeats_2, etc.) that
    // belong to the same owner as the primary row.
    var html     = primary.html;
    var chunkNum = 2;
    while (true) {
      var chunkType  = typeLower + '_' + chunkNum;
      var chunkFound = false;
      for (var j = 0; j < rows.length; j++) {
        if (rows[j].type === chunkType && rows[j].owner === primary.owner) {
          html += rows[j].html;
          chunkFound = true;
          break;
        }
      }
      if (!chunkFound) break;
      chunkNum++;
    }

    Logger.log('Cache HIT (sheet) for: ' + target + ' (assembled ' + html.length + ' chars)');
    // Promote to CacheService for next lookup. Only when the payload fits
    // CacheService's 100KB limit: truncated HTML breaks the evaluation, and the
    // multi-chunk sheet path serves large payloads correctly.
    if (html.length <= 99000) {
      try { CacheService.getScriptCache().put(csKey, html, _RESULT_CS_TTL); } catch(ce) {}
    }
    return html;
  } catch(e) {
    Logger.log('findCachedResult_ error: ' + e);
    return null;
  }
}

// ── Cache: save a result for one agent — chunks HTML > 48k chars across rows ──
// Replaces only THIS agent's rows for the Interaction ID and type, so saving or
// editing one agent's evaluation never touches another agent's.
function saveCachedResult_(interactionId, analysisType, htmlResult, sapId) {
  if (!interactionId) return;
  // Guard: empty HTML would delete all existing rows (deletion runs before append)
  // then write a blank entry — permanently destroying the previous cached result.
  if (!htmlResult) {
    Logger.log('saveCachedResult_: empty htmlResult for ' + interactionId + '/' + analysisType + ' — skipping to preserve existing cache');
    return;
  }
  try {
    // saveCachedResult_ and updateCachedResult_ are callable from any signed-in
    // user's browser, and what is stored here is later rendered for agents and
    // leaders. Sanitise on the way IN so no caller can plant markup that runs.
    var html = sanitizeEvalHtml_(htmlResult);
    if (!html) {
      Logger.log('saveCachedResult_: nothing left after sanitising for ' + interactionId + '/' + analysisType + ' — skipping to preserve existing cache');
      return;
    }
    var atype     = (analysisType || '').trim();
    var id        = interactionId.trim();
    var typeLower = atype.toLowerCase();
    var sap       = (sapId || '').toString().trim();

    // Drop the old CacheService copy FIRST, so a failed or skipped put() below
    // can never leave the previous HTML behind to be served for up to 6 hours.
    var csKey = _RESULT_CS_PREFIX + id + '_' + typeLower + '_' + sap;
    try { CacheService.getScriptCache().remove(csKey); } catch(ce) {}
    if (html.length <= 99000) {
      try { CacheService.getScriptCache().put(csKey, html, _RESULT_CS_TTL); } catch(ce) {}
    }

    // Write to Cache sheet — delete this agent's stale rows first (including
    // chunk rows like sales_2, sales_3) so a concurrent re-submission does not
    // leave orphaned rows that findCachedResult_ would concatenate with the new
    // HTML. Only columns A, B and E are read: the HTML column is never loaded.
    var ss    = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss, CACHE_SHEET);
    ensureHeaders(sheet, CACHE_HEADERS);

    // Read the rows to replace, delete them by index, then append: three steps that
    // go wrong if another save runs in between (its deletions shift the indexes, and
    // THIS save deletes another agent's row). So they run under the script lock.
    // ensureHeaders above takes the same lock, which is why it is not inside here.
    // If the lock cannot be had, save anyway: a late or unordered save is better than
    // an evaluation that is never stored.
    var lock = LockService.getScriptLock(), locked = false;
    try { lock.waitLock(10000); locked = true; }
    catch (le) { Logger.log('saveCachedResult_: could not get the script lock (' + le + ') - saving without it'); }
    try {
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var idType  = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        var owners  = sheet.getRange(2, _CACHE_SAP_COL, lastRow - 1, 1).getValues();
        var legacyMine = null; // looked up only if a blank-owner row of ours turns up
        for (var d = idType.length - 1; d >= 0; d--) {
          var rowId   = (idType[d][0] || '').toString().trim();
          var rowType = (idType[d][1] || '').toString().trim().toLowerCase();
          if (rowId !== id) continue;
          if (!(rowType === typeLower || rowType.indexOf(typeLower + '_') === 0)) continue;
          var owner = (owners[d][0] || '').toString().trim();
          var mine  = owner === sap;
          if (!mine && owner === '') {
            if (legacyMine === null) legacyMine = _legacyCacheOwnerSap(id, typeLower) === sap;
            mine = legacyMine;
          }
          if (mine) sheet.deleteRow(d + 2); // +2: 1-based index + skip header row
        }
      }

      var chunkCount = Math.ceil(html.length / _CACHE_CHUNK_SIZE) || 1;
      for (var c = 0; c < chunkCount; c++) {
        var chunk    = html.substring(c * _CACHE_CHUNK_SIZE, (c + 1) * _CACHE_CHUNK_SIZE);
        var rowType2 = c === 0 ? atype : atype + '_' + (c + 1);
        // Sheets reads what is written here as if typed. A chunk can start anywhere in
        // the HTML, so one that began with = + - @ became a formula (an error value on
        // read-back) and one that began with an apostrophe lost it, and the agent was
        // then shown a damaged evaluation. A leading apostrophe is Sheets' own "this is
        // text" marker and is not stored, so putting one on EVERY chunk is lossless.
        // (The other cells go through appendRow(), which neutralises formulas.)
        appendRow(sheet, [id, rowType2, new Date(), "'" + chunk, sap]);
      }
      Logger.log('Cached result for: ' + id + ' (' + html.length + ' chars, ' + chunkCount + ' chunk(s))');
    } finally {
      if (locked) { try { lock.releaseLock(); } catch (re) {} }
    }
  } catch(e) {
    Logger.log('saveCachedResult_ error: ' + e);
  }
}

// ── Cache: replace an existing cached result with user-edited HTML ────────────
function updateCachedResult_(interactionId, analysisType, editedHTML, sapId) {
  if (!interactionId || !editedHTML) return;
  try {
    // Deletion of this agent's stale rows is handled inside saveCachedResult_.
    saveCachedResult_(interactionId.trim(), analysisType, editedHTML, sapId);
    Logger.log('updateCachedResult_: replaced cache for ' + interactionId.trim() + ' (' + (analysisType || '').trim().toLowerCase() + ')');
  } catch(e) {
    Logger.log('updateCachedResult_ error: ' + e);
  }
}
