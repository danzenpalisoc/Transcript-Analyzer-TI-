/**
 * SheetService.gs
 * Handles all Google Sheets read/write operations.
 */

// ── Large-payload CacheService helpers ───────────────────────────────────────
// CacheService caps a single value at 100KB. Call sites used to guard with
// `if (json.length <= 99000) cache.put(...)`, which meant the biggest and most
// expensive datasets — the 12,624-row roster, the agent email map, the audit
// log — silently skipped caching entirely and were re-read from the sheet on
// EVERY execution. These split the payload across numbered keys so those
// datasets are actually cached.
//
// A partially-evicted payload must read as a miss, never as truncated JSON:
// returning half a document would break every downstream parse.
var _CACHE_CHUNK_CHARS = 90000;   // headroom under the 100KB per-key limit
var _CACHE_MAX_CHUNKS  = 120;     // ~10MB ceiling; refuse anything larger

function _cachePutLarge_(cache, key, str, ttlSeconds) {
  try {
    var n = Math.ceil(str.length / _CACHE_CHUNK_CHARS) || 1;
    if (n > _CACHE_MAX_CHUNKS) {
      Logger.log('_cachePutLarge_: ' + key + ' too large (' + str.length + ' chars) — not cached');
      return false;
    }
    var payload = {};
    for (var i = 0; i < n; i++) {
      payload[key + '_c' + i] = str.substring(i * _CACHE_CHUNK_CHARS, (i + 1) * _CACHE_CHUNK_CHARS);
    }
    payload[key + '_n'] = String(n);
    cache.putAll(payload, ttlSeconds);
    return true;
  } catch (e) {
    Logger.log('_cachePutLarge_ (' + key + '): ' + e);
    return false;
  }
}

// Must be used to invalidate anything written with _cachePutLarge_ — a plain
// cache.remove(key) would leave the numbered chunks in place and the entry
// would keep reading as a hit.
function _cacheRemoveLarge_(cache, key) {
  try {
    var n = parseInt(cache.get(key + '_n') || '0', 10);
    var keys = [key, key + '_n'];
    for (var i = 0; i < n; i++) keys.push(key + '_c' + i);
    cache.removeAll(keys);
  } catch (e) {
    Logger.log('_cacheRemoveLarge_ (' + key + '): ' + e);
  }
}

function _cacheGetLarge_(cache, key) {
  try {
    var n = parseInt(cache.get(key + '_n') || '0', 10);
    if (!n || n < 1) return null;

    var keys = [];
    for (var i = 0; i < n; i++) keys.push(key + '_c' + i);
    var got = cache.getAll(keys);

    var out = '';
    for (var j = 0; j < n; j++) {
      var part = got[key + '_c' + j];
      if (part === null || part === undefined) return null; // partial eviction -> miss
      out += part;
    }
    return out;
  } catch (e) {
    Logger.log('_cacheGetLarge_ (' + key + '): ' + e);
    return null;
  }
}


// ── Spreadsheet handle memoisation ───────────────────────────────────────────
// openById is a network round trip. Several Config constants point at the SAME
// file (MAIN_SPREADSHEET_ID / AUDIT_TRACKING_SS_ID / FCR_DASHBOARD_SS_ID are
// identical, as are TRAINEE_ROSTER_SS_ID / AT_DATA_GCP_SS_ID and
// TRAINER_LOOKUP_SS_ID / USERS_SS_ID), and getOrCreateSpreadsheet() alone runs
// five-plus times per submit. Memoised per execution.
var _ssHandleCache = {};

function openSpreadsheetCached(id) {
  if (!_ssHandleCache[id]) _ssHandleCache[id] = SpreadsheetApp.openById(id);
  return _ssHandleCache[id];
}

function getOrCreateSpreadsheet() {
  return openSpreadsheetCached(MAIN_SPREADSHEET_ID);
}

function getOrCreateSheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  return sheet;
}

// Sheets already confirmed to have headers during this execution. ensureHeaders
// runs 6+ times per submit and each call cost a getLastRow() round trip even on
// its fast path; headers cannot vanish mid-execution, so one check is enough.
var _headersChecked = {};

function ensureHeaders(sheet, headers) {
  var memoKey = sheet.getSheetId();
  if (_headersChecked[memoKey]) return;

  // Fast path — if headers already exist, skip the lock entirely.
  if (sheet.getLastRow() > 0) { _headersChecked[memoKey] = true; return; }

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
    _headersChecked[memoKey] = true;
  } catch(le) {
    Logger.log('ensureHeaders: lock failed — ' + le);
  } finally {
    try { _ehLock.releaseLock(); } catch(le2) {}
  }
}

function appendRow(sheet, rowData) {
  var safeRow = rowData.map(function(v) {
    if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
    return v;
  });
  sheet.appendRow(safeRow);
}

function getSpreadsheetUrl() {
  return getOrCreateSpreadsheet().getUrl();
}

// ── Cache: look up a previously processed Interaction ID ──────────────────────
// Uses TextFinder to avoid a full linear scan on large Cache sheets.
// Returns the cached HTML string, or null if not found.
// CacheService key for interaction HTML (avoids TextFinder on every duplicate call)
var _RESULT_CS_PREFIX = 'html_result_v1_';
var _RESULT_CS_TTL    = 6 * 60 * 60; // 6 hours

function findCachedResult(interactionId, analysisType) {
  if (!interactionId) return null;
  try {
    var target    = interactionId.trim();
    var typeLower = (analysisType || '').trim().toLowerCase();

    // ── Layer 1: CacheService (O(1), no sheet read) ───────────────────────────
    var csKey    = _RESULT_CS_PREFIX + target + '_' + typeLower;
    var csResult = _cacheGetLarge_(CacheService.getScriptCache(), csKey);
    if (csResult) {
      Logger.log('Cache HIT (CacheService) for: ' + target);
      return csResult;
    }

    // ── Layer 2: Cache sheet TextFinder (slower, but persistent) ─────────────
    var ss    = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss, CACHE_SHEET);
    ensureHeaders(sheet, CACHE_HEADERS);
    if (sheet.getLastRow() < 2) return null;

    var finder  = sheet.getRange('A:A').createTextFinder(target).matchEntireCell(true);
    var matches = finder.findAll();

    // Read each matched row exactly once, then assemble in memory. The previous
    // version issued a getValues() per match and then, for every continuation
    // chunk, re-scanned all matches with another getValues() each — O(matches^2)
    // sheet round trips for a multi-chunk report.
    var htmlByType = {};
    for (var i = 0; i < matches.length; i++) {
      var row = matches[i].getRow();
      if (row < 2) continue;
      var rowData = sheet.getRange(row, 1, 1, 4).getValues()[0];
      var rowType = (rowData[1] || '').toString().trim().toLowerCase();
      var cell    = rowData[3] ? rowData[3].toString() : '';
      // First non-empty row for a given type wins, matching the old behaviour
      // of skipping blank rows and continuing the scan.
      if (!htmlByType[rowType] && cell) htmlByType[rowType] = cell;
    }

    var html = htmlByType[typeLower];
    if (!html) {
      Logger.log('Cache MISS for: ' + target);
      return null;
    }

    // Append continuation chunks (sales_2, sales_3, repeats_2, ...)
    for (var chunkNum = 2; ; chunkNum++) {
      var part = htmlByType[typeLower + '_' + chunkNum];
      if (!part) break;
      html += part;
    }

    Logger.log('Cache HIT (sheet) for: ' + target + ' (assembled ' + html.length + ' chars)');
    // Promote to CacheService for next lookup. Chunked, so oversized
    // reports are cached too — previously anything over 100KB skipped this
    // layer entirely and paid the TextFinder scan on every repeat lookup.
    _cachePutLarge_(CacheService.getScriptCache(), csKey, html, _RESULT_CS_TTL);
    return html;
  } catch(e) {
    Logger.log('findCachedResult error: ' + e);
    return null;
  }
}

// ── Cache: save a new result — chunks HTML > 48k chars across multiple rows ───
var _CACHE_CHUNK_SIZE = 48000; // safe margin under Google Sheets 50k cell limit

function saveCachedResult(interactionId, analysisType, htmlResult) {
  if (!interactionId) return;
  // Guard: empty HTML would delete all existing rows (deletion runs before append)
  // then write a blank entry — permanently destroying the previous cached result.
  // Math.ceil(0 / CHUNK_SIZE) === 0, and 0 || 1 === 1, so a blank row IS appended.
  if (!htmlResult) {
    Logger.log('saveCachedResult: empty htmlResult for ' + interactionId + '/' + analysisType + ' — skipping to preserve existing cache');
    return;
  }
  try {
    var html      = htmlResult || '';
    var atype     = (analysisType || '').trim();
    var id        = interactionId.trim();
    var typeLower = atype.toLowerCase();

    // Write to CacheService, chunked so large reports are cached rather than
    // skipped. Never truncate here: a half-written report renders as broken
    // tables and missing tags.
    var csKey = _RESULT_CS_PREFIX + id + '_' + typeLower;
    _cachePutLarge_(CacheService.getScriptCache(), csKey, html, _RESULT_CS_TTL);

    // Write to Cache sheet — delete any stale rows first (includes chunk rows
    // like sales_2, sales_3) so a concurrent re-submission does not leave orphaned
    // rows that findCachedResult would concatenate with the new HTML.
    var ss    = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss, CACHE_SHEET);
    ensureHeaders(sheet, CACHE_HEADERS);

    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var colAB = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      var stale = [];
      for (var d = 0; d < colAB.length; d++) {
        var rowId   = (colAB[d][0] || '').toString().trim();
        var rowType = (colAB[d][1] || '').toString().trim().toLowerCase();
        if (rowId === id &&
            (rowType === typeLower || rowType.indexOf(typeLower + '_') === 0)) {
          stale.push(d + 2); // +2: 1-based index + skip header row
        }
      }
      // Delete as contiguous blocks, bottom-up so lower row numbers stay valid.
      // A report's chunk rows are appended together and so are almost always
      // adjacent, collapsing N deleteRow() calls into one deleteRows() per block.
      for (var s = stale.length - 1; s >= 0; s--) {
        var end = stale[s], start = end;
        while (s > 0 && stale[s - 1] === start - 1) { s--; start = stale[s]; }
        sheet.deleteRows(start, end - start + 1);
      }
    }

    var chunkCount = Math.ceil(html.length / _CACHE_CHUNK_SIZE) || 1;
    var stamp = new Date();
    var rows  = [];
    for (var c = 0; c < chunkCount; c++) {
      rows.push([
        id,
        c === 0 ? atype : atype + '_' + (c + 1),
        stamp,
        html.substring(c * _CACHE_CHUNK_SIZE, (c + 1) * _CACHE_CHUNK_SIZE)
      ]);
    }
    // One write instead of an appendRow() per chunk.
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    Logger.log('Cached result for: ' + id + ' (' + html.length + ' chars, ' + chunkCount + ' chunk(s))');
  } catch(e) {
    Logger.log('saveCachedResult error: ' + e);
  }
}

// ── Cache: replace an existing cached result with user-edited HTML ────────────
function updateCachedResult(interactionId, analysisType, editedHTML) {
  if (!interactionId || !editedHTML) return;
  try {
    var id        = interactionId.trim();
    var typeLower = (analysisType || '').trim().toLowerCase();

    // Deletion of stale chunk rows is handled inside saveCachedResult.
    // The deletion block that was here previously was redundant dead code —
    // saveCachedResult performs the identical bottom-up deletion before appending,
    // so this function's pre-deletion was a no-op that only added sheet API I/O.
    saveCachedResult(id, analysisType, editedHTML);
    Logger.log('updateCachedResult: replaced cache for ' + id + ' (' + typeLower + ')');
  } catch(e) {
    Logger.log('updateCachedResult error: ' + e);
  }
}

