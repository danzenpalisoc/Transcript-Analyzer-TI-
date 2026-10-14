/**
 * Who may call what (security review, batch 2b).
 *
 * Every public function in a .gs file can be called by any signed-in user from
 * the browser console with google.script.run. Three groups, three rules:
 *
 *   BROWSER   functions the pages really call. They must stay public and must NOT
 *             be admin-gated, or a normal analyst is locked out of the app.
 *   MAINTENANCE  repair / backfill / diagnose / rename functions that nothing but
 *             an admin running them by hand ever calls. They must start with
 *             requireAdmin_().
 *   INTERNAL  helpers only server code calls (the AI key, the roster, the cache).
 *             They are renamed with a trailing underscore, which Apps Script does
 *             not expose to the browser.
 *
 * The first test is the safety net: if a page starts calling a function this file
 * does not know about, the test fails until someone decides which group it is in.
 *
 * Usage: node tests/access-control-test.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const root = path.join(__dirname, '..');

const GS = fs.readdirSync(root).filter(f => f.endsWith('.gs'));
const HTML = ['index.html', 'Dashboard.html', 'EvalView.html'];
const src = {};
GS.concat(HTML).forEach(f => { src[f] = read(f); });

const BROWSER = [
  'chatAssistant', 'detectTranscriptSpeakers', 'exportAudits', 'generateAndServePDF', 'getAIAnalytics',
  'getAuditLogData', 'getDashboardData', 'getDashboardDataFresh', 'getEvalViewData', 'getEvaluationByAuditRef',
  'getFlagDropdownOptions', 'getObserverInfo', 'getOverviewHighlightsLowlights', 'getQAUsersFromSheet', 'getRosterAll',
  'getRosterBySapId', 'getSalesAttemptedCounts', 'getSapIdForParticipant', 'getTranscriptMetadataAndRoster',
  'sendAnalyticsBriefing', 'sendAuditEmail', 'submitTranscript', 'synthesizeSpeech', 'warmLookupCaches',
  'createScorecardDeck',
];
const MAINTENANCE = [
  'diagnoseATDataColumns', 'repairNextTruncated', 'repairTruncatedSalesEvaluations', 'diagnoseRosterColumns',
  'backfillDashboardData', 'backfillLocaleInDashboard', 'enrichDashboardRCA', 'fixTruncatedLocale',
  'updateSpreadsheetSMARTFormat', 'addSalesAttemptedColumn', 'testSalesAttempted', 'backfillVTIDFromATData',
  'backfillVTIDToAuditLog', 'fixRepeatProjectionPct', 'addDashboardColumnNotes', 'showRecentErrors',
  'diagnoseFCRDashboard', 'diagnoseATDataRow', 'testLookupDirect', 'diagnoseSapIdRow', 'diagnoseATDataForAgent',
  'cleanAnalysisSheet', 'diagnoseRosterTabInMainSS', 'testAdminNotification', 'testEmailLookup', 'repairSheetHeaders',
  'createActionRegistry', 'backfillObserverName', 'reportDuplicateAuditRefs', 'applyDuplicateAuditRefRenames',
  'getSpreadsheetUrl', 'diagnoseSalesEvaluations', 'enrichDashboardData', 'generateAuditPDF', 'appendRegistryRow',
  'verifyLatestAudit', 'verifyMyLatestAudit', 'verifyDashboardData', 'addMissingHeaders', 'diagnoseDataIssues',
  'previewDataRepairs', 'applyDataRepairs', 'backupSpreadsheet', 'diagnoseAnalyticsContext', 'previewScorecards', 'testSlidesAccess', 'diagnoseDashboardGaps',
];
const INTERNAL = [
  'callFuelIX', 'getAllRosterData', 'lookupBySapId', 'lookupSapId', 'lookupVTID', 'saveCachedResult',
  'updateCachedResult', 'findCachedResult', 'getTeamDetailsFromAuditHistory', 'updateDashboardPDFLink',
  // security batch 1: these send mail, write the shared caches, spend AI quota, or return rosters, e-mail addresses and the
  // whole audit log; nothing in a browser needs them, so they stop being callable from google.script.run
  'analyzeTranscript', 'callAIForAnalytics', 'getPDFKnowledgeFromFolder', 'getLOBKnowledge', 'getEvaluationFromCache',
  'getAdminRecipients', 'getRecipientsFromRoster', 'lookupAgentEmail', 'lookupFromFCRDashboard', 'lookupUserFromUsersSheet',
  'notifyAdmins', 'logSubmissionError', 'invalidateAuditLogCache', 'invalidateDashboardCache', 'readAnalyticsCache',
  'writeAnalyticsCache', 'readAuditLog', 'readDashboardCache', 'writeDashboardCache', 'readDashboardSheet', 'resolveEmail',
  'resolveObserver', 'translateTranscriptText', 'buildAuditRecipients', 'getTraineeInfo', 'getTrainerInfo',
  'extractStructuredRCA', 'getTranscriptMetadata', 'findAuditRefForInteraction', 'doGetDashboard', 'generateAuditRef', 'buildAnalyticsContext',
];
// Public on purpose and harmless: they only turn their arguments into text or HTML, touch no sheet, mail, cache or AI service.
const PURE = [
  'adminRow', 'analysisTable', 'buildAdminEmailHTML', 'buildAgentEmailHTML', 'buildEvalFormHTML', 'buildFallbackHTML',
  'buildRepeatsHTML', 'buildRepeatsPrompt', 'buildResultHTML', 'buildSalesHTML', 'buildSalesPrompt', 'cell', 'cellHtml',
  'coachingList', 'dbUniqServer', 'emailRow', 'esc', 'escEmail', 'extractStructuredRCAFromHTML',
  'extractTextBlock', 'filterTranscriptByAgent', 'fixBadgeClasses', 'highlightBox', 'htmlToPlainText', 'infoBar',
  'parseTranscriptMetadata', 'posCard', 'renderAnalyticsHTML', 'sanitiseHTMLForPDF', 'scoreBadge', 'section', 'sharedCSS',
];
// Public because Apps Script requires it (web-app entry points) or because the arguments cannot come from a browser: a Sheet
// object is not something google.script.run can send, so these helpers cannot be driven from outside.
const ACCEPTED = {
  doGet: 'web app entry point; serves the page and nothing else',
  doPost: 'web app entry point for the registry writer; refuses a wrong key (Test 6)',
  appendRow: 'needs a Sheet object, which cannot be sent from a browser',
  ensureHeaders: 'needs a Sheet object, which cannot be sent from a browser',
  getOrCreateSheet: 'needs a Spreadsheet object, which cannot be sent from a browser',
  getOrCreateSpreadsheet: 'takes no input and returns a Spreadsheet object that cannot cross to the browser; only creates the file when it is missing',
  openSheetWithRetry: 'returns a Sheet object that cannot cross to the browser; opens nothing the caller could read',
};

function defsOf(name) {
  const out = [];
  for (const f of GS) if (new RegExp('^function\\s+' + name + '\\s*\\(', 'm').test(src[f])) out.push(f);
  return out;
}
// Only the start of a function is needed to see whether it is gated. (Matching
// braces to find the END is unreliable here: some bodies hold braces in strings.)
function body(name) {
  const f = defsOf(name)[0];
  if (!f) return null;
  const i = src[f].search(new RegExp('^function\\s+' + name + '\\s*\\(', 'm'));
  return src[f].slice(i, i + 400);
}
const startsWithGate = fn => /^function\s+\w+\s*\([^)]*\)\s*\{\s*(\/\/[^\n]*\n\s*)*requireAdmin_\(\);/.test(fn);

console.log('\nTest 1 — safety net: every function a page calls is known, public and ungated');
{
  // names the pages call: google.script.run...<name>(
  const called = new Set();
  for (const f of HTML) {
    for (const m of src[f].matchAll(/google\.script\.run((?:\s*\.\s*with\w+\((?:[^()]|\([^()]*\))*\))*)\s*\.\s*([A-Za-z_]\w*)\s*\(/g)) called.add(m[2]);
  }
  const unknown = [...called].filter(n => BROWSER.indexOf(n) === -1 && n !== 'withSuccessHandler' && n !== 'withFailureHandler');
  check('the pages call no function this test does not list', unknown.length === 0, 'unlisted: ' + unknown.join(', '));
  for (const n of BROWSER) {
    const d = defsOf(n);
    check(n + ' is defined once and public', d.length === 1 && !n.endsWith('_'), 'defined in: ' + d.join(','));
    check(n + ' is NOT admin-gated', d.length === 1 && !startsWithGate(body(n)));
  }
}

console.log('\nTest 2 — no page calls a private (underscore) function');
{
  const bad = [];
  for (const f of HTML) for (const m of src[f].matchAll(/google\.script\.run[^;]*?\.\s*([A-Za-z]\w*_)\s*\(/g)) bad.push(f + ':' + m[1]);
  check('no google.script.run call targets a name ending in _', bad.length === 0, bad.join(', '));
}

console.log('\nTest 3 — maintenance functions start with requireAdmin_()');
for (const n of MAINTENANCE) {
  const d = defsOf(n);
  check(n + ' exists exactly once', d.length === 1, 'defined in: ' + d.join(',') || 'nowhere');
  if (d.length === 1) check(n + ' is gated first thing', startsWithGate(body(n)));
}

console.log('\nTest 4 — internal helpers are private and nothing still calls the old public name');
for (const n of INTERNAL) {
  check(n + ' is no longer public', defsOf(n).length === 0, 'still defined in: ' + defsOf(n).join(','));
  check(n + '_ is defined once', defsOf(n + '_').length === 1, 'defined in: ' + defsOf(n + '_').join(','));
  const stale = [];
  for (const f of GS) {
    if (f === 'CreateActionRegistry.gs') continue;           // change-log text, not code
    src[f].split('\n').forEach((l, i) => {
      const t = l.trim();
      if (t.indexOf('//') === 0 || t.indexOf('*') === 0) return;
      let k, from = 0;
      while ((k = l.indexOf(n, from)) !== -1) {
        from = k + n.length;
        const before = l[k - 1], after = l[from];
        if (/[A-Za-z0-9_$.]/.test(before || '') || after === '_') continue;
        let j = from; while (l[j] === ' ') j++;
        if (l[j] === '(') { stale.push(f + ':' + (i + 1)); break; }
      }
    });
  }
  check('no code still calls ' + n + '(', stale.length === 0, stale.slice(0, 4).join(', '));
}

console.log('\nTest 5 — requireAdmin_ behaves');
{
  function gate(email, opts) {
    opts = opts || {};
    const sandbox = {
      Logger: { log() {} }, JSON, Array, String, Object, Error,
      PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
      Session: {
        getActiveUser: () => ({ getEmail: () => { if (opts.throws) throw new Error('no access'); return email; } }),
        // In "Execute as: Me" the effective user is the OWNER for every caller. Make it an
        // admin so a fallback to it would wrongly let a caller with a hidden email through.
        getEffectiveUser: () => ({ getEmail: () => 'danzen.palisoc@telus.com' }),
      },
    };
    const ctx = vm.createContext(sandbox);
    vm.runInContext(read('Config.gs'), ctx);
    return ctx;
  }
  const passes = (e, o) => { try { gate(e, o).requireAdmin_(); return true; } catch (x) { return false; } };
  check('an admin on telus.com passes', passes('danzen.palisoc@telus.com'));
  check('an admin on telusinternational.com passes (the owner may sign in there)', passes('danzen.palisoc@telusinternational.com'));
  check('case does not matter', passes('Danzen.Palisoc@TELUS.com'));
  check('every listed admin passes', ['lobell.bantaya', 'carlos.maldonado', 'robin.joseph'].every(u => passes(u + '@telus.com')));
  check('an ordinary analyst is refused', !passes('some.analyst@telus.com'));
  check('an admin name on a foreign domain is refused', !passes('danzen.palisoc@evil.com'));
  check('a look-alike domain is refused', !passes('danzen.palisoc@telus.com.evil.com'));
  check('a blank email is refused (never falls back to the effective user)', !passes(''));
  check('an unreadable identity is refused', !passes('x', { throws: true }));
  let msg = '';
  try { gate('some.analyst@telus.com').requireAdmin_(); } catch (e) { msg = e.message; }
  check('the refusal says it is for the admin team', /admin/i.test(msg), msg);
}

console.log('\nTest 6 — the registry secret comes from Script Properties and is not an open door');
{
  function post(key, propValue) {
    const out = { appended: 0 };
    const sandbox = {
      Logger: { log() {} }, JSON, Error,
      PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k === 'REGISTRY_SECRET' ? propValue : null) }) },
      ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ text: s, setMimeType() { return this; } }) },
      appendRegistryRow_: () => { out.appended++; return { success: true }; },
    };
    const ctx = vm.createContext(sandbox);
    const legacy = src['Code.gs'].match(/^var REGISTRY_SECRET_LEGACY[^\n]*/m);
    vm.runInContext((legacy ? legacy[0] : '') + '\n' + extractFunction(src['Code.gs'], 'getRegistrySecret_') + '\n' + extractFunction(src['Code.gs'], 'doPost'), ctx);
    const res = ctx.doPost({ postData: { contents: JSON.stringify({ key, data: { issue: 'x' } }) } });
    return { res: JSON.parse(res.text), appended: out.appended };
  }
  const a = post('wrong', 'NEW-SECRET');
  check('a wrong key is refused and nothing is written', a.res.success === false && a.appended === 0, JSON.stringify(a));
  const b = post('NEW-SECRET', 'NEW-SECRET');
  check('the Script Property value is accepted', b.res.success === true && b.appended === 1, JSON.stringify(b));
  const c = post('NH_REGISTRY_2026_DPZ', 'NEW-SECRET');
  check('once the property is set, the old committed secret stops working', c.res.success === false && c.appended === 0, JSON.stringify(c));
  const d = post(undefined, 'NEW-SECRET');
  check('a missing key is refused', d.res.success === false && d.appended === 0);
  const e = post('', undefined);
  check('with no property set an empty key is still refused', e.res.success === false && e.appended === 0);
  const f = post('NH_REGISTRY_2026_DPZ', undefined);
  check('TRANSITIONAL: with no property set the legacy secret still works', f.res.success === true && f.appended === 1, JSON.stringify(f));
  check('the old public secret constant is gone from the code path', !/key\s*!==\s*REGISTRY_SECRET\b/.test(src['Code.gs']));
}

console.log('\nTest 7 — one appendRegistryRow, and the dead mailer is gone');
{
  check('appendRegistryRow is defined exactly once across all files', GS.reduce((n, f) => n + (src[f].match(/^function\s+appendRegistryRow\s*\(/gm) || []).length, 0) === 1);
  check('appendRegistryRow_ (the real writer) is defined once', GS.reduce((n, f) => n + (src[f].match(/^function\s+appendRegistryRow_\s*\(/gm) || []).length, 0) === 1);
  check('sendSubmissionEmail (dead, public, sends mail) is removed', defsOf('sendSubmissionEmail').length === 0);
}

console.log('\nTest 8 — every public function belongs to exactly one group (nothing is public by accident)');
{
  const publicNames = [];
  for (const f of GS) for (const m of src[f].matchAll(/^function\s+([A-Za-z0-9]+)\s*\(/gm)) publicNames.push(m[1]);   // no trailing _ , so public
  const groups = { BROWSER, MAINTENANCE, PURE, ACCEPTED: Object.keys(ACCEPTED) };
  const unplaced = publicNames.filter(n => !Object.keys(groups).some(g => groups[g].indexOf(n) !== -1));
  check('every public function is listed in BROWSER, MAINTENANCE, PURE or ACCEPTED', unplaced.length === 0, 'unlisted: ' + unplaced.join(', '));
  const doubled = publicNames.filter(n => Object.keys(groups).filter(g => groups[g].indexOf(n) !== -1).length > 1);
  check('no function is in two groups', doubled.length === 0, doubled.join(', '));
  const stale = Object.keys(groups).flatMap(g => groups[g].filter(n => publicNames.indexOf(n) === -1).map(n => g + ':' + n));
  check('no group lists a function that is not public any more', stale.length === 0, stale.join(', '));
  // a PURE function must not reach for any service that reads, writes, mails or spends quota
  const SERVICES = /\b(SpreadsheetApp|MailApp|GmailApp|DriveApp|UrlFetchApp|CacheService|PropertiesService|LockService|ScriptApp|DocumentApp|SlidesApp|Drive\.)/;
  const impure = PURE.filter(n => {
    const f = defsOf(n)[0]; if (!f) return false;
    const i = src[f].search(new RegExp('^function\\s+' + n + '\\s*\\(', 'm'));
    const rest = src[f].slice(i + 10); const end = rest.search(/^function\s/m);
    return SERVICES.test(rest.slice(0, end === -1 ? rest.length : end));
  });
  check('no PURE function touches a sheet, mail, cache, Drive or the network', impure.length === 0, impure.join(', '));
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
