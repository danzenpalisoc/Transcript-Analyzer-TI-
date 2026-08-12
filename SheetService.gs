/**
 * SheetService.gs
 * Handles all Google Sheets read/write operations.
 */

function getOrCreateSpreadsheet() {
  var files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  return SpreadsheetApp.create(SPREADSHEET_NAME);
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
  sheet.appendRow(rowData);
}

function getSpreadsheetUrl() {
  return getOrCreateSpreadsheet().getUrl();
}

// ── Cache: look up a previously processed Interaction ID ──────────────────────
// Uses TextFinder to avoid a full linear scan on large Cache sheets.
// Returns the cached HTML string, or null if not found.
function findCachedResult(interactionId, analysisType) {
  if (!interactionId) return null;
  try {
    var ss    = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss, CACHE_SHEET);
    ensureHeaders(sheet, CACHE_HEADERS);
    if (sheet.getLastRow() < 2) return null;

    var target    = interactionId.trim();
    var typeLower = (analysisType || '').trim().toLowerCase();

    // TextFinder searches the Interaction ID column (col A) without loading all rows
    var finder  = sheet.getRange('A:A').createTextFinder(target).matchEntireCell(true);
    var matches = finder.findAll();

    for (var i = 0; i < matches.length; i++) {
      var row     = matches[i].getRow();
      if (row < 2) continue;  // skip header
      var rowData = sheet.getRange(row, 1, 1, 4).getValues()[0];
      var rowType = (rowData[1] || '').toString().trim().toLowerCase();
      if (rowType === typeLower) {
        Logger.log('Cache HIT for interaction: ' + interactionId);
        return rowData[3] ? rowData[3].toString() : null;
      }
    }
    Logger.log('Cache MISS for interaction: ' + interactionId);
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

