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
    var cacheKey = 'qa_users_list_v5';   // v5: names only. v4 held every address and role for an hour.
    var cached   = cache.get(cacheKey);
    if (cached) { try { return JSON.parse(cached); } catch(e) {} }

    var ss     = openSheetWithRetry(USERS_SS_ID, 'Users / Trainer Lookup Sheet');
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
      // Names only: any signed-in user can call this function, and the page needs just the
      // name for the QA filter. The address and role of each person stay on the Users sheet.
      users.push({ name: name });
    }

    var _qu = JSON.stringify(users);
    if (_qu.length <= 99000) { try { cache.put(cacheKey, _qu, 60 * 60); } catch(e) {} }
    Logger.log('getQAUsersFromSheet: ' + users.length + ' users loaded');
    return users;
  } catch(e) {
    Logger.log('getQAUsersFromSheet error: ' + e);
    return [];
  }
}

// ── Look up a user in the Users sheet by email ────────────────────────────────
function lookupUserFromUsersSheet_(email) {
  if (!email) return null;
  try {
    var cache    = CacheService.getUserCache();
    var cacheKey = 'userssheet_' + email.replace(/[^a-zA-Z0-9]/g, '_');
    var cached   = cache.get(cacheKey);
    if (cached) { try { return JSON.parse(cached); } catch(e) {} }

    var ss    = openSheetWithRetry(USERS_SS_ID, 'Users / Trainer Lookup Sheet');
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
      Logger.log('lookupUserFromUsersSheet_: ' + result.name + ' [' + result.role + ']');
      return result;
    }
    return null;
  } catch(e) {
    Logger.log('lookupUserFromUsersSheet_ error: ' + e);
    return null;
  }
}

// ── Users sheet: lower-cased Name → Email Address ─────────────────────────────
// A name that appears on more than one row with different addresses maps to ALL of them,
// each once, in sheet order, separated by ", " (the admin chose this over leaving it blank,
// 2026-10-26; it cannot tell one person with two addresses from two people with one name).
// Throws when the Users sheet cannot be read; callers decide.
function usersEmailByName_() {
  var ss    = openSheetWithRetry(USERS_SS_ID, 'Users / Trainer Lookup Sheet');
  var sheet = ss.getSheetByName(USERS_TAB);
  var map   = {};
  if (!sheet) return map;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return map;
  var hdrs     = data[0].map(function(h){ return (h || '').toString().toLowerCase().trim(); });
  var emailCol = hdrs.indexOf('email address');
  var nameCol  = hdrs.indexOf('name');
  if (emailCol < 0 || nameCol < 0) return map;
  var lists = {};
  for (var i = 1; i < data.length; i++) {
    var n = (data[i][nameCol]  || '').toString().toLowerCase().trim();
    var e = (data[i][emailCol] || '').toString().toLowerCase().trim();
    if (!n || !e) continue;
    if (!lists[n]) lists[n] = [];
    if (lists[n].indexOf(e) === -1) lists[n].push(e);
  }
  Object.keys(lists).forEach(function(k) { map[k] = lists[k].join(', '); });
  return map;
}

// ── Email of the person who submitted an audit ────────────────────────────────
// The signed-in caller, read on the server (never the browser's observerEmail, which a
// user could change). When Google hides the caller's address, the Users sheet row for
// the observer name (every address when the name has more than one); blank when it is not
// found. Never throws: a blank
// email must not stop an audit from being saved.
function resolveAuditorEmail_(observerName) {
  var email = '';
  try { email = (Session.getActiveUser().getEmail() || '').toString().toLowerCase().trim(); } catch (e) {}
  if (email) return email;
  try { return usersEmailByName_()[(observerName || '').toString().toLowerCase().trim()] || ''; }
  catch (le) { Logger.log('resolveAuditorEmail_: Users sheet lookup failed: ' + le); return ''; }
}

// ── Backfill blank Observer Name rows in Dashboard_Data + Audit_Log ───────────
// Run ONCE from Apps Script editor when Observer Name is blank for existing records.
function backfillObserverName() {
  requireAdmin_();
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch(e) {}
  try { if (!email) email = Session.getEffectiveUser().getEmail() || ''; } catch(e) {}
  if (!email) { Logger.log('backfillObserverName: no email found'); return; }

  var user = lookupUserFromUsersSheet_(email);
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

  invalidateDashboardCache_();
  invalidateAuditLogCache_();
  Logger.log('=== backfillObserverName done: ' + total + ' rows updated ===');
}
