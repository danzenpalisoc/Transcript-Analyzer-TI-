/**
 * EmailGuard.gs
 * Who may email an audit, to whom, and how often. sendAuditEmail (Code.gs) is open to every signed-in user because the page
 * calls it, and it sends mail from the app owner's mailbox, so it must not be a way to send anything to anyone:
 *
 *   - only an admin, or the analyst who saved the audit, may email it (_callerOwnsAudit_)
 *   - the mail goes only to addresses on the company domains (_splitInternalRecipients_); an address that is not is dropped and
 *     reported, never silently sent
 *   - an analyst may send a limited number of emails per hour (_emailRateTake_)
 *   - text taken from the form is cut to one short line (_oneLine_) before it goes into a subject or a body
 */

var EMAIL_PER_HOUR = 30;          // per analyst
var EMAIL_PER_HOUR_ADMIN = 100;

// Company domains (Config.gs ADMIN_EMAIL_DOMAINS) plus any an admin lists in the Script Property EMAIL_EXTRA_DOMAINS
// (comma separated, for example "contractor.example.com").
function _emailAllowedDomains_() {
  var domains = (typeof ADMIN_EMAIL_DOMAINS !== 'undefined' ? ADMIN_EMAIL_DOMAINS : []).slice();
  try {
    var extra = PropertiesService.getScriptProperties().getProperty('EMAIL_EXTRA_DOMAINS') || '';
    extra.split(',').forEach(function(d) {
      d = d.trim().toLowerCase();
      if (/^[a-z0-9.-]+\.[a-z]{2,}$/.test(d) && domains.indexOf(d) === -1) domains.push(d);
    });
  } catch (e) {}
  return domains;
}

function _isInternalEmail_(e) {
  var s = String(e === null || e === undefined ? '' : e).trim().toLowerCase();
  var at = s.lastIndexOf('@');
  return at > 0 && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s) && _emailAllowedDomains_().indexOf(s.slice(at + 1)) !== -1;
}

// { ok: [addresses on the company domains], dropped: [addresses that are not] }
function _splitInternalRecipients_(list) {
  var ok = [], dropped = [];
  (list || []).forEach(function(e) {
    if (!e) return;
    if (_isInternalEmail_(e)) ok.push(e); else dropped.push(e);
  });
  return { ok: ok, dropped: dropped };
}

// One short line: line breaks and control characters become spaces, runs of spaces collapse, cut at max.
function _oneLine_(v, max) {
  var s = String(v === null || v === undefined ? '' : v).replace(/[\u0000-\u001F\u007F]+/g, ' ').replace(/\s+/g, ' ').trim();
  return max && s.length > max ? s.substring(0, max).trim() : s;
}

function _normName_(v) { return String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim().toLowerCase(); }

// True when the caller is the analyst whose name is on the saved audit. The page records the observer name it got from
// getObserverInfo (the Users sheet name, else the part of the email before the @), so the same three spellings are accepted.
function _callerOwnsAudit_(auditRow, observer, callerEmail) {
  var want = _normName_(auditRow && auditRow['Observer Name']);
  if (!want) return false;
  var email = String(callerEmail || '').toLowerCase().trim();
  var mine = [observer && observer.name, email, email ? email.split('@')[0] : ''].map(_normName_).filter(Boolean);
  return mine.indexOf(want) !== -1;
}

// A short, collision-free-enough key from the WHOLE address (the first 45 characters of the address used to be all that counted).

// True when this person may send another audit email this hour (and counts it).
// If the cache itself is unavailable the limit is not applied (a cache hiccup must not stop real audit emails).
function _emailRateTake_(email, isAdmin) {
  try {
    var cache = CacheService.getScriptCache(), key = 'em_rl_' + _hashKey_(String(email || 'unknown').toLowerCase());
    var n = parseInt(cache.get(key) || '0', 10) || 0;
    if (n >= (isAdmin ? EMAIL_PER_HOUR_ADMIN : EMAIL_PER_HOUR)) return false;
    cache.put(key, String(n + 1), 3600);
  } catch (e) {
    Logger.log('_emailRateTake_: cache unavailable, limit not applied: ' + e);
  }
  return true;
}
