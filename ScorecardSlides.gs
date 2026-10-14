/**
 * ScorecardSlides.gs
 * Builds the agent scorecard deck in Google Slides: one cover slide and one slide per agent.
 *
 *   createScorecardDeck(auditRefs, label, filtersText)  the dashboard button; open to every signed-in analyst
 *   previewScorecards(auditRefs)                         admin, read-only: logs the cards, creates nothing
 *   testSlidesAccess()                                   admin: makes a one-slide deck to prove the conversion works
 *
 * How the deck is made: ScorecardPptx.gs writes the slides as a PowerPoint package, this file zips it, and Drive converts it to
 * Google Slides (Drive.Files.insert, the advanced Drive service the PDF export already uses). No Google Slides API, no Cloud
 * setting to switch on, no extra permission.
 *
 * The browser sends only a list of Audit Refs. Every number is computed here from the sheets (the page is never trusted for
 * data), the deck lives in the app owner's Drive folder "NH Analyzer Scorecards", and it is shared, as editor, only with the
 * person who asked for it.
 *
 * What goes on each slide comes from ScorecardLayout.gs (the same data the approved preview was drawn from); the numbers and
 * wording come from ScorecardMetrics.gs.
 */

var SCORECARD_FOLDER = 'NH Analyzer Scorecards';
var SCORECARD_MAX_AGENTS = 40;          // one deck holds up to this many agents
var SCORECARD_MAX_REFS = 2000;
var SCORECARD_DECKS_PER_HOUR = 5;       // per person
var SCORECARD_TRIES_PER_HOUR = 30;      // per person: every try, even one that is refused or empty, because each one reads the sheets
var SCORECARD_WORKING_PER_AGENT = 3;    // newest audits per agent whose report is read for "What's Working"
var SCORECARD_PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
var SCORECARD_SLIDES_MIME = 'application/vnd.google-apps.presentation';

// ── The deck: PowerPoint package -> Drive -> Google Slides ────────────────────

// Zip the package parts (each blob is named with its path inside the zip).
function _ssPptxBlob_(parts) {
  var blobs = parts.map(function(p) { return Utilities.newBlob(p.xml, 'application/xml', p.path); });
  var zip = Utilities.zip(blobs, 'scorecards.pptx');
  zip.setContentType(SCORECARD_PPTX_MIME);
  return zip;
}

function _ssFolder_() {
  return _ownedFolder_(SCORECARD_FOLDER);
}

// Creates the deck in the scorecard folder as a Google Slides file. slides: array of layout-shape arrays. Returns { id, url }.
// Drive can convert in two ways; the one the PDF export uses is tried first, and the other if it is refused.
function _ssCreateDeck_(title, slides) {
  var blob = _ssPptxBlob_(buildPptxParts_(slides));
  var parents = [{ id: _ssFolder_().getId() }];
  var file;
  try {
    file = Drive.Files.insert({ title: title, mimeType: SCORECARD_SLIDES_MIME, parents: parents }, blob);
  } catch (e1) {
    Logger.log('scorecard: first conversion style refused, trying the other: ' + e1);
    file = Drive.Files.insert({ title: title, mimeType: SCORECARD_PPTX_MIME, parents: parents }, blob, { convert: true });
  }
  if (!file || !file.id) throw new Error('Drive did not return the new file.');
  return { id: file.id, url: 'https://docs.google.com/presentation/d/' + file.id + '/edit' };
}

// Share the deck, as editor, with the person who asked.
function _ssShare_(id, email) {
  if (email) DriveApp.getFileById(id).addEditor(email);
}

// ── Reading the data (the page sends only Audit Refs) ─────────────────────────

// Validate what the browser sent. Returns { ok, refs (unique, trimmed), error }.
function _ssValidateRefs_(auditRefs) {
  if (!Array.isArray(auditRefs)) return { ok: false, error: 'No audits were selected.' };
  if (auditRefs.length > SCORECARD_MAX_REFS) return { ok: false, error: 'Too many audits were selected (' + auditRefs.length + '). Narrow the filters and try again.' };
  var seen = {}, refs = [];
  for (var i = 0; i < auditRefs.length; i++) {
    var r = typeof auditRefs[i] === 'string' ? auditRefs[i].trim() : auditRefs[i];
    if (!isValidAuditRef_(r)) return { ok: false, error: 'The selection contains something that is not an audit reference.' };
    if (!seen[r]) { seen[r] = true; refs.push(r); }
  }
  if (!refs.length) return { ok: false, error: 'No audits were selected.' };
  return { ok: true, refs: refs };
}

// The dashboard's merge (Audit_Log first, Dashboard_Data for the analysis fields; an audit that is only in Dashboard_Data is kept),
// reduced to the fields a scorecard uses. refSet: { ref: true }. tests/scorecard-slides-test.js compares it with the page.
function _ssMergeRows_(logRows, dashRows, refSet) {
  var dashMap = {};
  (dashRows || []).forEach(function(r) { var ref = String(r['Audit Ref'] || '').trim(); if (ref) dashMap[ref] = r; });
  var out = [], haveRef = {};
  var make = function(log, dash, ref) {
    var pick = function(k) { return log[k] || dash[k] || ''; };
    var tm = log['Team Member'] || ((dash['Team Member'] && dash['Team Member'].length <= 80) ? dash['Team Member'] : '') || '';
    return { 'Audit Ref': ref, 'Timestamp': log['Submitted At'] || dash['Timestamp'] || '', 'SAP ID': pick('SAP ID'), 'Team Member': tm,
      'Team Leader': pick('Team Leader'), 'Line of Business': pick('Line of Business'), 'Locale': pick('Locale'), 'VTID': pick('VTID'),
      'Analysis Type': pick('Analysis Type'), 'Repeat Projection %': pick('Repeat Projection %'), 'Issue Resolved': pick('Issue Resolved'),
      'Transfer Occurred': pick('Transfer Occurred'), 'Email Status': pick('Email Status'),
      'Critical Flags': dash['Critical Flags'] || '', 'RCA Sub-Parameter': dash['RCA Sub-Parameter'] || '',
      'Top SMART Opportunity': dash['Top SMART Opportunity'] || dash['Top Opportunity'] || '', 'Sales Attempted': dash['Sales Attempted'] || '' };
  };
  (logRows || []).forEach(function(log) {
    var ref = String(log['Audit Ref'] || '').trim();
    haveRef[ref] = true;
    if (!refSet[ref]) return;
    out.push(make(log, dashMap[ref] || {}, ref));
  });
  Object.keys(dashMap).forEach(function(ref) {
    if (haveRef[ref] || !refSet[ref]) return;
    out.push(make({}, dashMap[ref], ref));
  });
  return out;
}

// The newest `perAgent` audit refs of each agent, newest first (the reports that are read for "What's Working").
function _ssNewestRefsPerAgent_(rows, perAgent) {
  var byAgent = {};
  rows.forEach(function(r) {
    var k = _scKey_(r['Team Member']);
    if (!k) return;
    (byAgent[k] = byAgent[k] || []).push(r);
  });
  var refs = [];
  Object.keys(byAgent).forEach(function(k) {
    byAgent[k].sort(function(a, b) { return (Date.parse(String(b['Timestamp'])) || 0) - (Date.parse(String(a['Timestamp'])) || 0); });
    byAgent[k].slice(0, perAgent).forEach(function(r) { refs.push(r['Audit Ref']); });
  });
  return refs;
}

// "What's Working" for the given refs, read from the saved report in the Analysis sheet. Reads the Audit Ref column once and then
// ONLY the report cell of each wanted audit (a report is large), and stops at the time budget. Returns { ref: [text, ...] }.
function _ssLoadWhatsWorking_(refs, budgetMs) {
  var out = {};
  if (!refs.length) return out;
  var started = new Date().getTime();
  var sheet = getOrCreateSpreadsheet().getSheetByName(ANALYSIS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return out;
  var hd = _trimmedHeaders_(sheet), refCol = hd.indexOf('Audit Ref') + 1, resCol = hd.indexOf('AI Analysis Result') + 1;
  if (refCol < 1 || resCol < 1) return out;
  var refsCol = sheet.getRange(2, refCol, sheet.getLastRow() - 1, 1).getValues();
  var rowOf = {};
  refsCol.forEach(function(v, i) { var r = String(v[0] || '').trim(); if (r && rowOf[r] === undefined) rowOf[r] = i + 2; });
  for (var i = 0; i < refs.length; i++) {
    if (new Date().getTime() - started > (budgetMs || 90000)) break;
    var row = rowOf[refs[i]];
    if (!row) continue;
    out[refs[i]] = extractWhatsWorking_(sheet.getRange(row, resCol).getValue());
  }
  return out;
}

// Rows and cards for a list of refs (no deck, no Drive).
function _ssPrepare_(refs) {
  var refSet = {};
  refs.forEach(function(r) { refSet[r] = true; });
  var rows = _ssMergeRows_(readAuditLog_(), readDashboardSheet_(), refSet);
  return { rows: rows, built: buildAgentCards_(rows, null) };
}

function _ssLegend_() {
  var TH = SCORECARD_THRESHOLDS, unit = { repeatRisk: '%', resolution: '%', flags: ' per audit', transfer: '%', salesAttempt: '%' };
  return Object.keys(TH).map(function(k) {
    return { label: TH[k].label, target: TH[k].target + unit[k] + (TH[k].direction === 'max' ? ' or lower' : ' or higher') };
  });
}

function _ssClean_(v, max) { return String(v === null || v === undefined ? '' : v).replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim().substring(0, max); }

// The slides (layout shapes) for a set of cards.
function _ssBuildSlides_(built, label, filtersText, audits) {
  var total = built.cards.length + 1;
  var slides = [scorecardCoverShapes_({ label: label, range: built.range, agents: built.cards.length, audits: audits, filters: filtersText, thresholds: _ssLegend_() }, total)];
  built.cards.forEach(function(c, i) { slides.push(scorecardSlideShapes_(c, i + 2, total)); });
  return slides;
}

// ── Who is asking, how often, one deck at a time ──────────────────────────────

function _ssCaller_() {
  try { return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) { return ''; }
}

// Returns true when this person may build another deck this hour (and counts it).
function _ssRateLimitTake_(email) {
  var cache = CacheService.getScriptCache(), key = 'sc_rl_' + Utilities.base64Encode(email).substring(0, 60);
  var n = parseInt(cache.get(key) || '0', 10) || 0;
  if (n >= SCORECARD_DECKS_PER_HOUR) return false;
  cache.put(key, String(n + 1), 3600);
  return true;
}

// A message the person can act on. The real error is logged; an admin also sees a part of it.
function _ssFriendlyError_(e) {
  var raw = String((e && e.message) || e || '');
  Logger.log('scorecard error: ' + raw);
  var msg;
  if (/too many times|rate limit|quota|limit exceeded/i.test(raw)) msg = 'Google is busy right now. Please wait a few minutes and try again.';
  else if (/permission|access denied|forbidden|insufficient/i.test(raw)) msg = 'The analyzer is not allowed to create the deck in Drive. Please tell the QA admin team.';
  else msg = 'The deck could not be built. Please try again, and tell the QA admin team if it keeps failing.';
  var admin = false;
  try { admin = isAdminCaller_(); } catch (x) {}
  return admin ? msg + ' (Details for admins: ' + raw.substring(0, 300) + ')' : msg;
}

// ── Public ────────────────────────────────────────────────────────────────────

// The dashboard button. auditRefs: the Audit Refs of the audits that pass the dashboard filters.
// Returns { success, url, agents, audits, slides, skipped } or { success: false, error }.
function createScorecardDeck(auditRefs, label, filtersText) {
  var email = _ssCaller_();
  if (!email) return { success: false, error: 'Could not tell who you are, so the deck cannot be shared with you. Please sign in again.' };
  var v = _ssValidateRefs_(auditRefs);
  if (!v.ok) return { success: false, error: v.error };

  // Counted BEFORE the lock and the reads: a try that ends empty or too big used to cost no quota, so one person could keep the
  // lock and the sheets busy for everybody by repeating it.
  if (!_hourlyTake_('sc_try_', email, SCORECARD_TRIES_PER_HOUR)) return { success: false, error: 'Too many scorecard requests in the last hour. Please try again later.' };
  var lock = LockService.getScriptLock(), locked = false;
  try {
    if (!lock.tryLock(2000)) return { success: false, error: 'Another scorecard deck is being built. Please try again in a minute.' };
    locked = true;

    var prep = _ssPrepare_(v.refs);
    if (!prep.built.cards.length) return { success: false, error: 'None of the selected audits has a usable Team Member, so there is nothing to put on a slide.' };
    if (prep.built.cards.length > SCORECARD_MAX_AGENTS) {
      return { success: false, error: 'This selection has ' + prep.built.cards.length + ' agents. A deck holds up to ' + SCORECARD_MAX_AGENTS + '. Narrow the filters (for example by Team Leader) and try again.' };
    }
    if (!_ssRateLimitTake_(email)) return { success: false, error: 'You have already built ' + SCORECARD_DECKS_PER_HOUR + ' decks in the last hour. Please try again later.' };

    var working = _ssLoadWhatsWorking_(_ssNewestRefsPerAgent_(prep.rows, SCORECARD_WORKING_PER_AGENT));
    var built = buildAgentCards_(prep.rows, { whatsWorking: working });
    var lbl = _ssClean_(label, 80), flt = _ssClean_(filtersText, 200);
    var slides = _ssBuildSlides_(built, lbl, flt, prep.rows.length);
    var deck = _ssCreateDeck_('Audit Scorecards' + (lbl ? ' - ' + lbl : '') + (built.range ? ' - ' + built.range : ''), slides);
    _ssShare_(deck.id, email);
    Logger.log('createScorecardDeck: ' + built.cards.length + ' agents, ' + prep.rows.length + ' audits, ' + slides.length + ' slides for ' + email + ': ' + deck.url);
    return { success: true, url: deck.url, agents: built.cards.length, audits: prep.rows.length, slides: slides.length, skipped: built.skipped.length };
  } catch (e) {
    return { success: false, error: _ssFriendlyError_(e) };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (x) {} }
  }
}

// READ-ONLY. Shows the card of every agent for the given refs, so the numbers can be compared with the dashboard.
function previewScorecards(auditRefs) {
  requireAdmin_();
  var v = _ssValidateRefs_(auditRefs);
  if (!v.ok) { Logger.log('previewScorecards: ' + v.error); return { success: false, error: v.error }; }
  var prep = _ssPrepare_(v.refs);
  var built = buildAgentCards_(prep.rows, null);
  Logger.log('previewScorecards: ' + built.cards.length + ' agent(s), ' + prep.rows.length + ' audit(s), ' + built.range + ', ' + built.skipped.length + ' skipped');
  built.cards.forEach(function(c) {
    Logger.log('  ' + c.name + ' (SAP ' + c.sap + ') ' + c.badge.text + ' | ' + c.tiles.map(function(t) { return t.label + ' ' + t.display + ' ' + t.statusText; }).join(' | ') +
               ' | audits ' + c.audits);
  });
  built.skipped.forEach(function(s) { Logger.log('  SKIPPED ' + s.ref + ': ' + s.reason); });
  return { success: true, agents: built.cards.length, audits: prep.rows.length, skipped: built.skipped, cards: built.cards };
}

// Admin: builds a one-agent sample deck, so the first run proves the conversion to Google Slides works.
function testSlidesAccess() {
  requireAdmin_();
  var email = _ssCaller_();
  var sample = buildAgentCards_([{ 'Audit Ref': 'SAMPLE', 'Timestamp': new Date().toISOString(), 'Team Member': 'Sample Agent', 'SAP ID': '000000', 'Analysis Type': 'Repeats & Transfer Analyzer',
    'Repeat Projection %': '20%', 'Issue Resolved': 'Yes', 'Critical Flags': '', 'Transfer Occurred': 'No', 'Email Status': 'Sent', 'Team Leader': 'Sample Leader',
    'Line of Business': 'Sample LOB', 'Locale': 'Sample Locale', 'VTID': '', 'Top SMART Opportunity': 'Sample coaching point', 'RCA Sub-Parameter': '', 'Sales Attempted': '' }], null);
  var slides = _ssBuildSlides_(sample, 'Conversion test', 'none', 1);
  var deck = _ssCreateDeck_('Scorecard conversion test', slides);
  _ssShare_(deck.id, email);
  Logger.log('testSlidesAccess: OK. A sample deck was made: ' + deck.url);
  return { success: true, url: deck.url };
}
