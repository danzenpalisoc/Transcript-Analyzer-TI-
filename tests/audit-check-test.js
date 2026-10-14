/**
 * verifyLatestAudit() (AuditCheck.gs): a read-only, admin-only checklist that proves
 * ONE audit landed everywhere it should, so an admin can run it once after an audit
 * instead of testing each place by hand.
 *
 * It must (a) say PASS for a healthy audit, (b) FAIL, naming the problem, for each
 * way an audit can be damaged, (c) never write anything, and (d) refuse non-admins.
 *
 * It reads the Cache SHEET directly, not the 6-hour script cache: the script-cache
 * copy would hide a corrupted sheet copy that the evaluation page would show later.
 *
 * Usage: node tests/audit-check-test.js
 */
const vm = require('vm');
const { read, extractFunction, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const codeSrc = read('Code.gs');

function makeEnv() {
  const store = {}, logs = [], writes = { cache: 0, sheet: 0 };
  const who = { email: 'danzen.palisoc@telus.com', observerName: 'QA One' };
  let refSeq = 0;
  const sandbox = {
    Logger: { log: m => logs.push(String(m)) },
    CacheService: { getScriptCache: () => ({ get: k => (k in store ? store[k] : null), put: (k, v) => { writes.cache++; store[k] = v; }, remove: k => { delete store[k]; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Utilities: { base64Encode: s => Buffer.from(String(s)).toString('base64'), sleep() {} },
    ScriptApp: { getScriptId: () => 'x' },
    Session: { getActiveUser: () => ({ getEmail: () => who.email }) },
    resolveObserver_: () => ({ name: who.observerName }),     // who the roster says the caller is
    parseTranscriptMetadata: txt => { const m = /INT:(\S+)/.exec(txt || ''); return { interactionId: m ? m[1] : '' }; },
    filterTranscriptByAgent: t => ({ filtered: t, stats: { applied: false, targetFound: false } }),
    translateTranscriptText_: t => t,
    analyzeTranscript_: () => '<div class="report-wrap"><div class="ai-chip">ok</div></div>',
    lookupBySapId_: () => null,
    generateAuditRef_: () => 'NHA-TEST-' + String(++refSeq).padStart(4, '0'),
    // a believable extraction so the Dashboard row is filled in
    extractTextBlock: (h, label) => (label === 'Call Summary' ? 'A short summary.' : ''),
    extractStructuredRCAFromHTML: () => null,
    invalidateAuditLogCache_: () => {}, invalidateDashboardCache_: () => {},
    logSubmissionError_: () => {},
    sharedCSS: () => '', fixBadgeClasses: h => h,
  };
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('Config.gs'), ctx);
  vm.runInContext(read('HtmlSanitizer.gs'), ctx);
  vm.runInContext(read('SheetService.gs'), ctx);
  ['_procKey_', 'htmlToPlainText', 'findAuditRefForInteraction_', 'produceEvaluation_', 'writeAuditRecords_', '_cellSafeTranscript_', 'submitTranscript']
    .forEach(n => vm.runInContext(extractFunction(codeSrc, n), ctx));
  vm.runInContext(read('AuditCheck.gs'), ctx);

  const sheets = {};
  [['TRANSCRIPTS_SHEET', 'TRANSCRIPTS_HEADERS'], ['ANALYSIS_SHEET', 'ANALYSIS_HEADERS'], ['DASHBOARD_DATA_SHEET', 'DASHBOARD_HEADERS'],
   ['AUDIT_LOG_SHEET', 'AUDIT_LOG_HEADERS'], ['CACHE_SHEET', 'CACHE_HEADERS'], ['ERROR_LOG_SHEET', 'ERROR_LOG_HEADERS']]
    .forEach(([s, h]) => { sheets[ctx[s]] = new FakeSheet(ctx[h], []); });
  const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new FakeSheet([], [])), getUrl: () => 'https://sheet' };
  ctx.getOrCreateSpreadsheet = () => ss;
  return { ctx, sheets, store, logs, writes, who };
}
const FORM = () => ({
  sapId: '111', transcript: 'INT:abc agent: hello', interactionId: 'abc', participant: 'Agent A',
  teamLeader: 'TL', opsManager: 'OM', lineOfBusiness: 'LOB', locale: 'Site', vtid: 'V1',
  analysisType: 'repeats', observerName: 'QA One', agentEmail: 'a@telus.com',
});
function healthy() {
  const e = makeEnv();
  const r = e.ctx.submitTranscript(FORM());
  if (!r.success) throw new Error('fixture submit failed: ' + JSON.stringify(r));
  return { e, ref: r.auditRef };
}
const failed = r => r.checks.filter(c => c.level === 'fail');
const names = r => failed(r).map(c => c.name).join(' | ');
const dropRows = (sheet, ref) => { sheet.grid = [sheet.grid[0]].concat(sheet.grid.slice(1).filter(r => String(r[0]).trim() !== ref)); };

console.log('\nTest 1 — a healthy audit passes every check');
{
  const { e, ref } = healthy();
  const r = e.ctx.verifyLatestAudit();
  check('ok === true', r.ok === true, names(r));
  check('it picked the latest audit', r.ref === ref, r.ref);
  check('there is a real checklist (not an empty pass)', r.checks.length >= 10, 'checks: ' + r.checks.length);
  check('the log shows PASS lines and a final RESULT line', e.logs.some(l => /\[PASS\]/.test(l)) && e.logs.some(l => /RESULT:.*PASSED/.test(l)), e.logs.slice(-3).join(' | '));
  check('"Not Sent" is information, not a failure', r.checks.some(c => c.level === 'info' && /Not Sent/i.test(c.detail + c.name)));
  const named = e.ctx.verifyLatestAudit(ref);
  check('an explicit ref works too', named.ok === true && named.ref === ref);
}

console.log('\nTest 2 — each way an audit can be damaged is FAILED and named');
{
  let x = healthy();
  dropRows(x.e.sheets[x.e.ctx.DASHBOARD_DATA_SHEET], x.ref);
  let r = x.e.ctx.verifyLatestAudit();
  check('missing Dashboard_Data row', r.ok === false && /Dashboard_Data/i.test(names(r)), names(r));

  x = healthy(); dropRows(x.e.sheets[x.e.ctx.ANALYSIS_SHEET], x.ref);
  r = x.e.ctx.verifyLatestAudit();
  check('missing Analysis row', r.ok === false && /Analysis/i.test(names(r)), names(r));

  x = healthy(); dropRows(x.e.sheets[x.e.ctx.TRANSCRIPTS_SHEET], x.ref);
  r = x.e.ctx.verifyLatestAudit();
  check('missing Transcripts row', r.ok === false && /Transcripts/i.test(names(r)), names(r));

  x = healthy(); const log = x.e.sheets[x.e.ctx.AUDIT_LOG_SHEET]; log.grid.push(log.grid[1].slice());
  r = x.e.ctx.verifyLatestAudit(x.ref);
  check('the same ref twice in Audit_Log (the email link would open the wrong one)', r.ok === false && /more than once|share/i.test(failed(r).map(c => c.name + c.detail).join(' ')), names(r));

  x = healthy(); dropRows(x.e.sheets[x.e.ctx.AUDIT_LOG_SHEET], x.ref);
  r = x.e.ctx.verifyLatestAudit(x.ref);
  check('an audit that is not in Audit_Log at all', r.ok === false && /Audit_Log/i.test(names(r)));

  x = healthy(); const dash = x.e.sheets[x.e.ctx.DASHBOARD_DATA_SHEET];
  dash.grid[1][x.e.ctx.DASHBOARD_HEADERS.indexOf('SAP ID')] = '999';
  r = x.e.ctx.verifyLatestAudit();
  check('SAP ID differs between sheets', r.ok === false && /match|consistent|differ/i.test(failed(r).map(c => c.name + c.detail).join(' ')), names(r));

  x = healthy(); x.e.store['proc_abc_111'] = String(Date.now());
  r = x.e.ctx.verifyLatestAudit();
  check('the "being processed" flag is stuck', r.ok === false && /processing|flag/i.test(names(r)), names(r));

  x = healthy(); const err = x.e.sheets[x.e.ctx.ERROR_LOG_SHEET];
  err.appendRow(x.e.ctx.ERROR_LOG_HEADERS.map(h => ({ 'Timestamp': new Date(Date.now() + 1000), 'Function': 'sendAuditEmail', 'SAP ID': '111', 'Interaction ID': 'abc', 'Error Message': 'boom' }[h] || '')));
  r = x.e.ctx.verifyLatestAudit();
  check('an error was logged for this interaction after the audit', r.ok === false && /error/i.test(names(r)) && /boom/.test(failed(r).map(c => c.detail).join(' ')), names(r));
}

console.log('\nTest 3 — the saved evaluation (Cache sheet) is checked, including a damaged copy the script cache would hide');
{
  let x = healthy();
  const cache = x.e.sheets[x.e.ctx.CACHE_SHEET];
  check('fixture: the evaluation really is in the Cache sheet', cache.grid.length > 1);
  cache.grid[1][3] = '#ERROR!';
  let r = x.e.ctx.verifyLatestAudit();
  check('a chunk turned into an error value is FAILED even though the script cache still holds a good copy', r.ok === false && /evaluation|cache/i.test(names(r)) && /ERROR/.test(failed(r).map(c => c.detail + c.name).join(' ')), names(r));

  x = healthy(); x.e.sheets[x.e.ctx.CACHE_SHEET].grid = [x.e.sheets[x.e.ctx.CACHE_SHEET].grid[0]];
  r = x.e.ctx.verifyLatestAudit();
  check('no saved evaluation at all', r.ok === false && /evaluation|cache/i.test(names(r)), names(r));

  x = healthy(); x.e.sheets[x.e.ctx.CACHE_SHEET].grid[1][3] = '<div class="report-wrap"><div>cut off in the midd';
  r = x.e.ctx.verifyLatestAudit();
  check('a report that looks cut off is FAILED', r.ok === false && /cut off|complete|truncat/i.test(failed(r).map(c => c.name + c.detail).join(' ')), names(r));

  x = healthy(); const c2 = x.e.sheets[x.e.ctx.CACHE_SHEET];
  c2.grid[1][4] = 'someone-else';
  r = x.e.ctx.verifyLatestAudit();
  check('an evaluation stored under another agent is not counted as this agent\'s', r.ok === false && /evaluation|cache/i.test(names(r)), names(r));
}

console.log('\nTest 4 — email status');
{
  let x = healthy();
  const log = x.e.sheets[x.e.ctx.AUDIT_LOG_SHEET], H = x.e.ctx.AUDIT_LOG_HEADERS;
  log.grid[1][H.indexOf('Email Status')] = 'Sent'; log.grid[1][H.indexOf('Recipients')] = 'a@telus.com, tl@telus.com';
  let r = x.e.ctx.verifyLatestAudit();
  check('Sent with recipients passes', r.ok === true, names(r));
  log.grid[1][H.indexOf('Recipients')] = '';
  r = x.e.ctx.verifyLatestAudit();
  check('Sent but no recipients recorded is FAILED', r.ok === false && /recipient/i.test(failed(r).map(c => c.name + c.detail).join(' ')), names(r));
  log.grid[1][H.indexOf('Email Status')] = 'Failed: quota exceeded';
  r = x.e.ctx.verifyLatestAudit();
  check('a failed email is FAILED and shows the reason', r.ok === false && /quota exceeded/.test(failed(r).map(c => c.detail).join(' ')), names(r));
}

console.log('\nTest 5 — it never writes, and only an admin may run it');
{
  const { e } = healthy();
  const before = { cache: e.writes.cache, rows: Object.keys(e.sheets).map(k => e.sheets[k].grid.length).join('/') };
  let del = 0, app = 0;
  Object.keys(e.sheets).forEach(k => { const s = e.sheets[k]; const a = s.appendRow.bind(s), d = s.deleteRow.bind(s); s.appendRow = x => { app++; return a(x); }; s.deleteRow = n => { del++; return d(n); }; });
  e.ctx.verifyLatestAudit();
  check('no rows appended or deleted', app === 0 && del === 0, 'append ' + app + ', delete ' + del);
  check('no sheet changed size', Object.keys(e.sheets).map(k => e.sheets[k].grid.length).join('/') === before.rows);
  check('nothing was put in the script cache', e.writes.cache === before.cache, 'puts: ' + (e.writes.cache - before.cache));

  e.who.email = 'some.analyst@telus.com';
  let threw = false, msg = '';
  try { e.ctx.verifyLatestAudit(); } catch (x) { threw = true; msg = x.message; }
  check('a non-admin is refused', threw && /admin/i.test(msg), msg);
  e.who.email = '';
  threw = false;
  try { e.ctx.verifyLatestAudit(); } catch (x) { threw = true; }
  check('a hidden identity is refused', threw);
}

console.log('\nTest 6 — an empty workbook is a clear failure, not a crash');
{
  const e = makeEnv();
  const r = e.ctx.verifyLatestAudit();
  check('ok === false with a message', r.ok === false && failed(r).length >= 1, names(r));
}

console.log('\nTest 7 — verifyMyLatestAudit checks the CALLER\'s newest audit, not whoever audited last');
{
  const e = makeEnv();
  const mine = e.ctx.submitTranscript(FORM());                       // observer "QA One"
  const other = FORM(); other.observerName = 'Someone Else'; other.interactionId = 'zzz'; other.transcript = 'INT:zzz hi'; other.sapId = '222';
  const theirs = e.ctx.submitTranscript(other);                      // a LATER audit by another analyst
  check('fixture: the newest audit in the sheet is the other analyst\'s', theirs.success && theirs.auditRef !== mine.auditRef);
  check('verifyLatestAudit picks the newest overall (the other analyst\'s)', e.ctx.verifyLatestAudit().ref === theirs.auditRef);
  const r = e.ctx.verifyMyLatestAudit();
  check('verifyMyLatestAudit picks MY audit', r.ref === mine.auditRef, r.ref + ' vs mine ' + mine.auditRef);
  check('...and it passes', r.ok === true, names(r));
  check('...and says whose audit it is looking for', e.logs.some(l => /QA One/.test(l)), e.logs.slice(-14).join(' | '));

  // my newest of several is chosen
  const again = FORM(); again.interactionId = 'def'; again.transcript = 'INT:def hi'; again.sapId = '333';
  const mine2 = e.ctx.submitTranscript(again);
  check('with two audits of mine, the newer one is checked', e.ctx.verifyMyLatestAudit().ref === mine2.auditRef);

  e.who.observerName = 'Nobody Known';
  const none = e.ctx.verifyMyLatestAudit();
  check('no audit by this observer is a clear failure, not a crash or someone else\'s audit', none.ok === false && /Nobody Known/.test(none.checks.map(c => c.name + c.detail).join(' ')), names(none));
  check('...and it does not fall back to another analyst\'s audit', !none.ref);

  e.who.observerName = '';
  const blank = e.ctx.verifyMyLatestAudit();
  check('an unknown observer name is a clear failure', blank.ok === false && !blank.ref, names(blank));

  e.who.observerName = 'qa one';
  check('the name match ignores case', e.ctx.verifyMyLatestAudit().ref === mine2.auditRef);

  e.who.observerName = 'QA One'; e.who.email = 'some.analyst@telus.com';
  let threw = false;
  try { e.ctx.verifyMyLatestAudit(); } catch (x) { threw = /admin/i.test(x.message); }
  check('a non-admin is refused', threw);

  // the refusal must not depend on finding an audit (verifyLatestAudit's own gate would
  // otherwise hide a missing gate here, and the failure path would run for anybody)
  e.who.observerName = 'Nobody Known';
  threw = false;
  try { e.ctx.verifyMyLatestAudit(); } catch (x) { threw = /admin/i.test(x.message); }
  check('a non-admin is refused even when no audit is found', threw);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
