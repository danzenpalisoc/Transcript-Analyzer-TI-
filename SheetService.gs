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
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4B286D')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
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
        Logger.log('Cache HIT (sheet) for: ' + target);
        // Promote to CacheService for next lookup
        if (html) {
          try { CacheService.getScriptCache().put(csKey, html.substring(0, 95000), _RESULT_CS_TTL); } catch(ce) {}
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

// ── Cache: save a new result ──────────────────────────────────────────────────
function saveCachedResult(interactionId, analysisType, htmlResult) {
  if (!interactionId) return;
  try {
    // Write to CacheService immediately (next lookup will be O(1))
    var csKey = _RESULT_CS_PREFIX + interactionId.trim() + '_' + (analysisType || '').trim().toLowerCase();
    try { CacheService.getScriptCache().put(csKey, (htmlResult || '').substring(0, 95000), _RESULT_CS_TTL); } catch(ce) {}

    // Write to Cache sheet (permanent storage)
    var ss    = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss, CACHE_SHEET);
    ensureHeaders(sheet, CACHE_HEADERS);
    sheet.appendRow([
      interactionId.trim(),
      (analysisType || '').trim(),
      new Date(),
      htmlResult
    ]);
    Logger.log('Cached result for: ' + interactionId);
  } catch(e) {
    Logger.log('saveCachedResult error: ' + e);
  }
}

