/**
 * Smoke test: a whole audit through the real submitTranscript / writeAuditRecords_ puts the
 * submitter's email in the LAST cell of the Dashboard_Data row, and nothing else moves.
 *
 * Only Google (sheets, session, cache, lock) and the AI are fake. The Users sheet is a fake with
 * a name that has two addresses, one with one, and one that is missing.
 *
 * Usage: node tests/auditor-email-submit-test.js
 */
const vm = require('vm');
const { read, extractFunction, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const codeSrc = read('Code.gs');

function makeEnv(opts) {
  opts = opts || {};
  const store = {}, logs = [];
  const who = { email: opts.email === undefined ? 'qa.one@telus.com' : opts.email };
  let refSeq = 0;
  const sandbox = {
    Logger: { log: m => logs.push(String(m)) },
    CacheService: { getScriptCache: () => ({ get: k => (k in store ? store[k] : null), put: (k, v) => { store[k] = v; }, remove: k => { delete store[k]; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Utilities: { base64Encode: s => Buffer.from(String(s)).toString('base64'), sleep() {} },
    ScriptApp: { getScriptId: () => 'x' },
    Session: { getActiveUser: () => ({ getEmail: () => who.email }) },
    parseTranscriptMetadata: txt => { const m = /INT:(\S+)/.exec(txt || ''); return { interactionId: m ? m[1] : '' }; },
    filterTranscriptByAgent: t => ({ filtered: t, stats: { applied: false, targetFound: false } }),
    translateTranscriptText_: t => t,
    analyzeTranscript_: () => '<div class="report-wrap"><div class="ai-chip">ok</div></div>',
    lookupBySapId_: () => null,
    generateAuditRef_: () => 'NHA-TEST-' + String(++refSeq).padStart(4, '0'),
    extractTextBlock: () => '', extractStructuredRCAFromHTML: () => null,
    invalidateAuditLogCache_: () => {}, invalidateDashboardCache_: () => {},
    logSubmissionError_: () => {},
    sharedCSS: () => '', fixBadgeClasses: h => h,
  };
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('Config.gs'), ctx);
  vm.runInContext(read('HtmlSanitizer.gs'), ctx);
  vm.runInContext(read('SheetService.gs'), ctx);
  vm.runInContext(read('ObserverHelper.gs'), ctx);
  ['_procKey_', 'htmlToPlainText', 'findAuditRefForInteraction_', 'produceEvaluation_', 'writeAuditRecords_', '_cellSafeTranscript_', 'submitTranscript']
    .forEach(n => vm.runInContext(extractFunction(codeSrc, n), ctx));

  const sheets = {};
  [['TRANSCRIPTS_SHEET', 'TRANSCRIPTS_HEADERS'], ['ANALYSIS_SHEET', 'ANALYSIS_HEADERS'], ['DASHBOARD_DATA_SHEET', 'DASHBOARD_HEADERS'],
   ['AUDIT_LOG_SHEET', 'AUDIT_LOG_HEADERS'], ['CACHE_SHEET', 'CACHE_HEADERS']]
    .forEach(([s, h]) => { sheets[ctx[s]] = new FakeSheet(ctx[h], []); });
  const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new FakeSheet([], [])), getUrl: () => 'https://sheet' };
  ctx.getOrCreateSpreadsheet = () => ss;

  const users = new FakeSheet(['Name', 'Email Address', 'Role'], [
    ['QA One', 'qa.one@telus.com', 'QA'],
    ['Twin QA', 'twin.qa@telus.com', 'QA'], ['Twin QA', 'twin.qa@telusinternational.com', 'QA'],
  ]);
  users.getDataRange = () => ({ getValues: () => users.grid.map(r => r.slice()) });
  const usersSs = { getSheetByName: n => (n === ctx.USERS_TAB ? users : null) };
  ctx.openSheetWithRetry = id => {
    if (opts.usersThrow && id === ctx.USERS_SS_ID) throw new Error('permission denied');
    return id === ctx.USERS_SS_ID ? usersSs : ss;
  };
  const dash = () => sheets[ctx.DASHBOARD_DATA_SHEET];
  const cell = (sheet, header, rowNo) => { const g = sheet.grid, c = g[0].indexOf(header); return g[rowNo || g.length - 1][c]; };
  return { ctx, sheets, dash, cell, who, logs };
}

const FORM = (observer, int) => ({
  sapId: '111', transcript: 'INT:' + (int || 'abc') + ' agent: hello', interactionId: int || 'abc', participant: 'Agent A',
  teamLeader: 'TL', opsManager: 'OM', lineOfBusiness: 'LOB', locale: 'Site', vtid: 'V1',
  analysisType: 'repeats', observerName: observer, agentEmail: 'a@telus.com',
});

console.log('\nTest 1 — the signed-in submitter is recorded, in the last column, and no other cell moves');
{
  const e = makeEnv({ email: 'QA.One@Telus.com' });
  const r = e.ctx.submitTranscript(FORM('QA One'));
  const D = e.ctx.DASHBOARD_HEADERS, row = e.dash().grid[1];
  check('success', r.success === true, JSON.stringify(r).slice(0, 160));
  check('Dashboard_Data row has one cell per header', row.length === D.length, row.length + ' vs ' + D.length);
  check('Auditor Email is the last cell, lower-cased', row[D.length - 1] === 'qa.one@telus.com', String(row[D.length - 1]));
  check('Agent Email is still in its own column (before it)', e.cell(e.dash(), 'Agent Email') === 'a@telus.com');
  check('Observer Name is still where it was', e.cell(e.dash(), 'Observer Name') === 'QA One');
  check('Audit_Log row has the same width as its headers (no new column there)', e.sheets[e.ctx.AUDIT_LOG_SHEET].grid[1].length === e.ctx.AUDIT_LOG_HEADERS.length);
}

console.log('\nTest 2 — Google hides the submitter: the Users sheet is used');
{
  let e = makeEnv({ email: '' });
  e.ctx.submitTranscript(FORM('QA One'));
  check('one address for a name with one address', e.cell(e.dash(), 'Auditor Email') === 'qa.one@telus.com', e.cell(e.dash(), 'Auditor Email'));

  e = makeEnv({ email: '' });
  e.ctx.submitTranscript(FORM('Twin QA'));
  check('both addresses for a name with two', e.cell(e.dash(), 'Auditor Email') === 'twin.qa@telus.com, twin.qa@telusinternational.com', e.cell(e.dash(), 'Auditor Email'));

  e = makeEnv({ email: '' });
  const r = e.ctx.submitTranscript(FORM('Nobody'));
  check('a name that is not there: audit saved, email blank', r.success === true && e.cell(e.dash(), 'Auditor Email') === '', JSON.stringify(r).slice(0, 120));
}

console.log('\nTest 3 — a failing Users sheet never loses the audit');
{
  const e = makeEnv({ email: '', usersThrow: true });
  const r = e.ctx.submitTranscript(FORM('QA One'));
  check('success and no warning about it', r.success === true && !(r.warnings || []).length, JSON.stringify(r.warnings));
  check('all four sheets still got their row', [e.ctx.TRANSCRIPTS_SHEET, e.ctx.ANALYSIS_SHEET, e.ctx.DASHBOARD_DATA_SHEET, e.ctx.AUDIT_LOG_SHEET].every(n => e.sheets[n].grid.length === 2));
  check('email blank', e.cell(e.dash(), 'Auditor Email') === '');
}

console.log('\nTest 4 — the form cannot choose the email');
{
  const e = makeEnv({ email: 'qa.one@telus.com' });
  const f = FORM('QA One'); f.observerEmail = 'someone.else@telus.com'; f.auditorEmail = 'someone.else@telus.com';
  e.ctx.submitTranscript(f);
  check('the session email is stored, not the one in the form', e.cell(e.dash(), 'Auditor Email') === 'qa.one@telus.com', e.cell(e.dash(), 'Auditor Email'));
}

console.log('\nTest 5 — completing an audit from the cache records the email too');
{
  const e = makeEnv({ email: 'qa.one@telus.com' });
  e.ctx.saveCachedResult_('abc', 'repeats', '<div class="report-wrap">cached</div>', '111');
  const r = e.ctx.submitTranscript(FORM('QA One'));
  check('success, AI not needed', r.success === true, JSON.stringify(r).slice(0, 120));
  check('email recorded on the recovered row', e.cell(e.dash(), 'Auditor Email') === 'qa.one@telus.com', e.cell(e.dash(), 'Auditor Email'));
}

console.log('\nTest 6 — two audits keep each their own submitter');
{
  const e = makeEnv({ email: 'qa.one@telus.com' });
  e.ctx.submitTranscript(FORM('QA One', 'i1'));
  e.who.email = 'other.qa@telus.com';
  e.ctx.submitTranscript(FORM('Other QA', 'i2'));
  check('row 1 and row 2 have different submitters', e.cell(e.dash(), 'Auditor Email', 1) === 'qa.one@telus.com' && e.cell(e.dash(), 'Auditor Email', 2) === 'other.qa@telus.com', e.dash().grid.map(r => r[r.length - 1]).join(' | '));
}

console.log('\n' + (check.state.failures === 0 ? 'ALL PASSED' : check.state.failures + ' FAILED'));
process.exit(check.state.failures === 0 ? 0 : 1);
