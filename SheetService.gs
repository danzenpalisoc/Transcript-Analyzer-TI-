/**
 * SheetService.gs
 * Handles all Google Sheets read/write operations.
 */

function getOrCreateSpreadsheet() {
  return SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
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

    var finder  = sheet.getRange('A:A').createTextFinder(target).matchEntireCell(true);
    var matches = finder.findAll();

    for (var i = 0; i < matches.length; i++) {
      var row     = matches[i].getRow();
      if (row < 2) continue;
      var rowData = sheet.getRange(row, 1, 1, 4).getValues()[0];
      var rowType = (rowData[1] || '').toString().trim().toLowerCase();
      if (rowType === typeLower) {
        var html = rowData[3] ? rowData[3].toString() : null;
        if (!html) continue;

        // Assemble any continuation chunks (sales_2, sales_3, repeats_2, etc.)
        var chunkNum = 2;
        while (true) {
          var chunkType = typeLower + '_' + chunkNum;
          var chunkFound = false;
          for (var j = 0; j < matches.length; j++) {
            var cRow  = matches[j].getRow();
            if (cRow < 2) continue;
            var cData = sheet.getRange(cRow, 1, 1, 4).getValues()[0];
            if ((cData[1] || '').toString().trim().toLowerCase() === chunkType) {
              html += (cData[3] || '').toString();
              chunkFound = true;
              break;
            }
          }
          if (!chunkFound) break;
          chunkNum++;
        }

        Logger.log('Cache HIT (sheet) for: ' + target + ' (assembled ' + html.length + ' chars)');
        // Promote to CacheService for next lookup
        // Only cache if payload fits within CacheService's 100KB limit.
        // Truncating to 95000 bytes silently breaks large HTML (broken tables,
        // missing closing tags). Skipping CacheService for oversized payloads
        // is safer — the sheet-based multi-chunk path still serves them correctly.
        if (html.length <= 99000) {
          try { CacheService.getScriptCache().put(csKey, html, _RESULT_CS_TTL); } catch(ce) {}
        }
        return html;
      }
    }

    Logger.log('Cache MISS for: ' + target);
    return null;
  } catch(e) {
    Logger.log('findCachedResult error: ' + e);
    return null;
  }
}

// ── Cache: save a new result — chunks HTML > 48k chars across multiple rows ───
var _CACHE_CHUNK_SIZE = 48000; // safe margin under Google Sheets 50k cell limit

function saveCachedResult(interactionId, analysisType, htmlResult) {
  if (!interactionId) return;
  try {
    var html      = htmlResult || '';
    var atype     = (analysisType || '').trim();
    var id        = interactionId.trim();
    var typeLower = atype.toLowerCase();

    // Write to CacheService (skip if payload exceeds 100KB — truncated HTML
    // causes broken evaluations; the sheet-based path handles large payloads correctly)
    var csKey = _RESULT_CS_PREFIX + id + '_' + typeLower;
    var _csHtml = html;
    if (_csHtml.length <= 99000) {
      try { CacheService.getScriptCache().put(csKey, _csHtml, _RESULT_CS_TTL); } catch(ce) {}
    }

    // Write to Cache sheet — delete any stale rows first (includes chunk rows
    // like sales_2, sales_3) so a concurrent re-submission does not leave orphaned
    // rows that findCachedResult would concatenate with the new HTML.
    var ss    = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss, CACHE_SHEET);
    ensureHeaders(sheet, CACHE_HEADERS);

    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var colAB = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (var d = colAB.length - 1; d >= 0; d--) {
        var rowId   = (colAB[d][0] || '').toString().trim();
        var rowType = (colAB[d][1] || '').toString().trim().toLowerCase();
        if (rowId === id &&
            (rowType === typeLower || rowType.indexOf(typeLower + '_') === 0)) {
          sheet.deleteRow(d + 2); // +2: 1-based index + skip header row
        }
      }
    }

    var chunkCount = Math.ceil(html.length / _CACHE_CHUNK_SIZE) || 1;
    for (var c = 0; c < chunkCount; c++) {
      var chunk   = html.substring(c * _CACHE_CHUNK_SIZE, (c + 1) * _CACHE_CHUNK_SIZE);
      var rowType = c === 0 ? atype : atype + '_' + (c + 1);
      sheet.appendRow([id, rowType, new Date(), chunk]);
    }
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

