/**
 * ObserverHelper.gs
 * User identity resolution and QA observer utilities.
 */

// ── Return all users from the Users sheet for the QA filter pill ──────────────
// Finds the sheet by GID (653814081) to avoid tab-name mismatches.
var USERS_SHEET_GID = 653814081;

function getQAUsersFromSheet() {
  try {
    var cache    = CacheService.getScriptCache();
    var cacheKey = 'qa_users_list_v4';
    var cached = _cacheGetLarge_(cache, cacheKey);
    if (cached) { try { return JSON.parse(cached); } catch(e) {} }

    var ss     = openSpreadsheetCached(USERS_SS_ID);
    // Find by GID first (most reliable), fall back to tab name
    var sheet  = null;
    var sheets = ss.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      if (sheets[s].getSheetId() === USERS_SHEET_GID) { sheet = sheets[s]; break; }
    }
    if (!sheet) sheet = ss.getSheetByName(USERS_TAB);
    if (!sheet) { Logger.log('getQAUsersFromSheet: tab not found'); return []; }

    Logger.log('getQAUsersFromSheet: reading tab "' + sheet.getName() + '"');

    var data = sheet.getDataRange().getValues();
    Logger.log('getQAUsersFromSheet: ' + data.length + ' rows, headers: ' + JSON.stringify(data[0]));
    if (data.length < 2) return [];

    var hdrs     = data[0].map(function(h){ return (h || '').toString().toLowerCase().trim(); });
    var nameCol  = hdrs.indexOf('name');
    var emailCol = hdrs.indexOf('email address');
    var roleCol  = hdrs.indexOf('role');

    if (nameCol < 0) { Logger.log('getQAUsersFromSheet: Name column not found in ' + JSON.stringify(data[0])); return []; }

    var users = [];
    for (var i = 1; i < data.length; i++) {
      var name = (data[i][nameCol] || '').toString().trim();
      if (!name) continue;
      users.push({
        name:  name,
        email: emailCol >= 0 ? (data[i][emailCol] || '').toString().trim() : '',
        role:  roleCol  >= 0 ? (data[i][roleCol]  || '').toString().trim() : ''
      });
    }

    var _qu = JSON.stringify(users);
    _cachePutLarge_(cache, cacheKey, _qu, 60 * 60);
    Logger.log('getQAUsersFromSheet: ' + users.length + ' users loaded');
    return users;
  } catch(e) {
    Logger.log('getQAUsersFromSheet error: ' + e);
    return [];
  }
}

// ── Users sheet as one shared email -> user map ───────────────────────────────
// Each caller used to cache only its OWN row, in getUserCache(). With 100+
// analysts that meant 100+ separate full-sheet reads of the same table per TTL
// window. Read once into a shared ScriptCache map instead.
var _usersByEmailInMemory = null;

function _getUsersByEmail() {
  if (_usersByEmailInMemory) return _usersByEmailInMemory;

  var cache    = CacheService.getScriptCache();
  var cacheKey = 'users_by_email_v1';
  var cached   = _cacheGetLarge_(cache, cacheKey);
  if (cached) {
    try { _usersByEmailInMemory = JSON.parse(cached); return _usersByEmailInMemory; } catch(e) {}
  }

  var map = {};
  try {
    var ss    = openSpreadsheetCached(USERS_SS_ID);
    var sheet = ss.getSheetByName(USERS_TAB);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      if (data.length >= 2) {
        var hdrs     = data[0].map(function(h){ return (h || '').toString().toLowerCase().trim(); });
        var emailCol = hdrs.indexOf('email address');
        var nameCol  = hdrs.indexOf('name');
        var roleCol  = hdrs.indexOf('role');
        if (emailCol >= 0 && nameCol >= 0) {
          for (var i = 1; i < data.length; i++) {
            var raw = (data[i][emailCol] || '').toString().trim();
            if (!raw) continue;
            map[raw.toLowerCase()] = {
              name: (data[i][nameCol] || '').toString().trim(),
              role: roleCol >= 0 ? (data[i][roleCol] || '').toString().trim() : ''
            };
          }
          _cachePutLarge_(cache, cacheKey, JSON.stringify(map), 8 * 60 * 60);
        }
      }
    }
  } catch(e) {
    Logger.log('_getUsersByEmail error: ' + e);
  }

  _usersByEmailInMemory = map;
  return map;
}

function lookupUserFromUsersSheet(email) {
  if (!email) return null;
  var rec = _getUsersByEmail()[email.toLowerCase().trim()];
  if (!rec) return null;
  // Echo back the caller's spelling of the email, as the previous version did.
  return { name: rec.name, role: rec.role, email: email };
}

// ── Backfill blank Observer Name rows in Dashboard_Data + Audit_Log ───────────
// Run ONCE from Apps Script editor when Observer Name is blank for existing records.
function backfillObserverName() {
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch(e) {}
  try { if (!email) email = Session.getEffectiveUser().getEmail() || ''; } catch(e) {}
  if (!email) { Logger.log('backfillObserverName: no email found'); return; }

  var user = lookupUserFromUsersSheet(email);
  var name = (user && user.name) ? user.name : email.split('@')[0];
  Logger.log('backfillObserverName: setting Observer Name = "' + name + '" for blank rows');

  var ss     = getOrCreateSpreadsheet();
  var total  = 0;

  [DASHBOARD_DATA_SHEET, AUDIT_LOG_SHEET].forEach(function(sheetName) {
    var sheet   = getOrCreateSheet(ss, sheetName);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    var lastCol  = sheet.getLastColumn();
    var hdrs     = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var obsCol   = hdrs.map(function(h){ return (h||'').toString().trim(); }).indexOf('Observer Name');
    if (obsCol < 0) return;
    obsCol++;
    var col     = sheet.getRange(2, obsCol, lastRow - 1, 1).getValues();
    var patched = 0;
    for (var i = 0; i < col.length; i++) {
      if (!(col[i][0] || '').toString().trim()) {
        sheet.getRange(i + 2, obsCol).setValue(name);
        patched++;
      }
    }
    Logger.log(sheetName + ': ' + patched + ' rows patched');
    total += patched;
  });

  invalidateDashboardCache();
  invalidateAuditLogCache();
  Logger.log('=== backfillObserverName done: ' + total + ' rows updated ===');
}
