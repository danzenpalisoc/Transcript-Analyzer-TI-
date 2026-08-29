/**
 * ObserverHelper.gs
 * User identity resolution and QA observer utilities.
 */

// ── Return all users from the Users sheet for the QA filter pill ──────────────
function getQAUsersFromSheet() {
  try {
    var cache    = CacheService.getScriptCache();
    var cacheKey = 'qa_users_list_v2';
    var cached   = cache.get(cacheKey);
    if (cached) { try { return JSON.parse(cached); } catch(e) {} }

    var ss    = SpreadsheetApp.openById(USERS_SS_ID);
    var sheet = ss.getSheetByName(USERS_TAB);
    if (!sheet) { Logger.log('getQAUsersFromSheet: tab not found'); return []; }

    var data    = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var hdrs    = data[0].map(function(h){ return (h || '').toString().toLowerCase().trim(); });
    var nameCol = hdrs.indexOf('name');
    if (nameCol < 0) { Logger.log('getQAUsersFromSheet: Name column not found'); return []; }

    var emailCol = hdrs.indexOf('email address');
    var roleCol  = hdrs.indexOf('role');

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

    try { cache.put(cacheKey, JSON.stringify(users), 60 * 60); } catch(e) {}
    Logger.log('getQAUsersFromSheet: ' + users.length + ' users loaded');
    return users;
  } catch(e) {
    Logger.log('getQAUsersFromSheet error: ' + e);
    return [];
  }
}

// ── Look up a user in the Users sheet by email ────────────────────────────────
function lookupUserFromUsersSheet(email) {
  if (!email) return null;
  try {
    var cache    = CacheService.getUserCache();
    var cacheKey = 'userssheet_' + email.replace(/[^a-zA-Z0-9]/g, '_');
    var cached   = cache.get(cacheKey);
    if (cached) { try { return JSON.parse(cached); } catch(e) {} }

    var ss    = SpreadsheetApp.openById(USERS_SS_ID);
    var sheet = ss.getSheetByName(USERS_TAB);
    if (!sheet) return null;

    var data    = sheet.getDataRange().getValues();
    if (data.length < 2) return null;

    var hdrs     = data[0].map(function(h){ return (h || '').toString().toLowerCase().trim(); });
    var emailCol = hdrs.indexOf('email address');
    var nameCol  = hdrs.indexOf('name');
    var roleCol  = hdrs.indexOf('role');
    if (emailCol < 0 || nameCol < 0) return null;

    var emailLower = email.toLowerCase().trim();
    for (var i = 1; i < data.length; i++) {
      if ((data[i][emailCol] || '').toString().toLowerCase().trim() !== emailLower) continue;
      var result = {
        name:  (data[i][nameCol] || '').toString().trim(),
        role:  roleCol >= 0 ? (data[i][roleCol] || '').toString().trim() : '',
        email: email
      };
      try { cache.put(cacheKey, JSON.stringify(result), 8 * 60 * 60); } catch(e) {}
      Logger.log('lookupUserFromUsersSheet: ' + result.name + ' [' + result.role + ']');
      return result;
    }
    return null;
  } catch(e) {
    Logger.log('lookupUserFromUsersSheet error: ' + e);
    return null;
  }
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
