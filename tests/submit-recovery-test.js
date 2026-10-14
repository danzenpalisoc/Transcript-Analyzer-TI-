/**
 * submitTranscript must not lose an audit, and must not lock an analyst out.
 *
 * Before: the report was cached (step 5), then the Analysis sheet was appended with
 * nothing around it. If THAT threw, the Dashboard and Audit_Log rows were never
 * written, and the retry hit the cache, answered "success" with an EMPTY audit
 * reference, and wrote nothing, so the analyst saw a report that was in no sheet and
 * the email failed with "Audit reference not found".
 *
 * Also: the "being processed" flag was only released on two of its four exits (a cache
 * hit and "Access denied" left it for 10 minutes), and with a blank Interaction ID on
 * the form (it is then read from the transcript) there was no guard at all.
 *
 * Usage: node tests/submit-recovery-test.js
 */
const vm = require('vm');
const { read, extractFunction, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const codeSrc = read('Code.gs');

function makeEnv() {
  const store = {};
  const logs = [], errors = [];
  const ai = { calls: 0, impl: () => '<div class="report-wrap"><div class="ai-chip">ok</div></div>' };
  let refSeq = 0;
  const sandbox = {
    Logger: { log: m => logs.push(String(m)) },
    CacheService: { getScriptCache: () => ({ get: k => (k in store ? store[k] : null), put: (k, v) => { store[k] = v; }, remove: k => { delete store[k]; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Utilities: { base64Encode: s => Buffer.from(String(s)).toString('base64'), sleep() {} },
    ScriptApp: { getScriptId: () => 'x' },
    // collaborators of submitTranscript that are not under test
    parseTranscriptMetadata: txt => { const m = /INT:(\S+)/.exec(txt || ''); return { interactionId: m ? m[1] : '' }; },
    filterTranscriptByAgent: t => ({ filtered: t, stats: { applied: false, targetFound: false } }),
    translateTranscriptText_: t => t,
    analyzeTranscript_: (...a) => { ai.calls++; return ai.impl(...a); },
    lookupBySapId_: () => null,
    generateAuditRef_: () => 'NHA-TEST-' + String(++refSeq).padStart(4, '0'),
    extractTextBlock: () => '', extractStructuredRCAFromHTML: () => null,
    invalidateAuditLogCache_: () => {}, invalidateDashboardCache_: () => {},
    logSubmissionError_: (fn, fd, e) => errors.push(fn + ': ' + (e && e.message)),
    sharedCSS: () => '', fixBadgeClasses: h => h,
  };
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('Config.gs'), ctx);
  vm.runInContext(read('HtmlSanitizer.gs'), ctx);
  vm.runInContext(read('SheetService.gs'), ctx);
  ['_procKey_', 'htmlToPlainText', 'findAuditRefForInteraction_', 'produceEvaluation_', 'writeAuditRecords_', '_cellSafeTranscript_', 'submitTranscript']
    .forEach(n => vm.runInContext(extractFunction(codeSrc, n), ctx));

  const sheets = {};
  [['TRANSCRIPTS_SHEET', 'TRANSCRIPTS_HEADERS'], ['ANALYSIS_SHEET', 'ANALYSIS_HEADERS'], ['DASHBOARD_DATA_SHEET', 'DASHBOARD_HEADERS'],
   ['AUDIT_LOG_SHEET', 'AUDIT_LOG_HEADERS'], ['CACHE_SHEET', 'CACHE_HEADERS']]
    .forEach(([s, h]) => { sheets[ctx[s]] = new FakeSheet(ctx[h], []); });
  const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new FakeSheet([], [])), getUrl: () => 'https://sheet' };
  ctx.getOrCreateSpreadsheet = () => ss;
  const rows = name => sheets[ctx[name]].grid.length - 1;
  return { ctx, sheets, store, logs, errors, ai, rows };
}

const SAP = '111', INT = 'abc';
const FORM = () => ({
  sapId: SAP, transcript: 'INT:' + INT + ' agent: hello', interactionId: INT, participant: 'Agent A',
  teamLeader: 'TL', opsManager: 'OM', lineOfBusiness: 'LOB', locale: 'Site', vtid: 'V1',
  analysisType: 'repeats', observerName: 'QA One', agentEmail: 'a@telus.com',
});
const flag = e => e.store['proc_' + INT + '_' + SAP];

console.log('\nTest 1 — the happy path writes every record and releases the flag');
{
  const e = makeEnv();
  const r = e.ctx.submitTranscript(FORM());
  check('success', r.success === true, JSON.stringify(r).slice(0, 200));
  check('an audit ref is returned', /^NHA-TEST-/.test(r.auditRef || ''), r.auditRef);
  check('Transcripts, Analysis, Dashboard_Data and Audit_Log each got one row',
        e.rows('TRANSCRIPTS_SHEET') === 1 && e.rows('ANALYSIS_SHEET') === 1 && e.rows('DASHBOARD_DATA_SHEET') === 1 && e.rows('AUDIT_LOG_SHEET') === 1,
        [e.rows('TRANSCRIPTS_SHEET'), e.rows('ANALYSIS_SHEET'), e.rows('DASHBOARD_DATA_SHEET'), e.rows('AUDIT_LOG_SHEET')].join('/'));
  check('the evaluation was cached', e.ctx.findCachedResult_(INT, 'repeats', SAP) !== null);
  check('the processing flag is released', !flag(e));
}

console.log('\nTest 2 — THE BUG: the Analysis sheet fails; the audit is still recorded');
{
  const e = makeEnv();
  e.sheets[e.ctx.ANALYSIS_SHEET].appendRow = () => { throw new Error('Analysis sheet unavailable'); };
  const r = e.ctx.submitTranscript(FORM());
  check('still reported as success', r.success === true, JSON.stringify(r).slice(0, 200));
  check('Dashboard_Data row was written', e.rows('DASHBOARD_DATA_SHEET') === 1);
  check('Audit_Log row was written (so the email can find the audit)', e.rows('AUDIT_LOG_SHEET') === 1);
  check('the analyst is warned about the missing Analysis row, with the ref', (r.warnings || []).some(w => /Analysis/i.test(w) && w.indexOf(r.auditRef) !== -1), JSON.stringify(r.warnings));
  check('the processing flag is released', !flag(e));
}

console.log('\nTest 3 — THE BUG: cached report but NO audit record -> the retry completes the records instead of lying');
{
  const e = makeEnv();
  // what an interrupted first attempt leaves behind: the Transcripts row and the cache, nothing else
  e.sheets[e.ctx.TRANSCRIPTS_SHEET].grid.push(e.ctx.TRANSCRIPTS_HEADERS.map(h => ({ 'Audit Ref': 'NHA-OLD-0001', 'SAP ID': SAP, 'Interaction ID': INT, 'Analysis Type': 'Repeats & Transfer Analyzer' }[h] || '')));
  e.ctx.saveCachedResult_(INT, 'repeats', '<div class="report-wrap">cached evaluation NHA-OLD-0001</div>', SAP);
  const r = e.ctx.submitTranscript(FORM());
  check('success', r.success === true, JSON.stringify(r).slice(0, 200));
  check('the ORIGINAL audit ref is reused, so the report and the ref agree', r.auditRef === 'NHA-OLD-0001', r.auditRef);
  check('the AI was NOT called again', e.ai.calls === 0, 'AI calls: ' + e.ai.calls);
  check('the cached evaluation is what is returned', /cached evaluation/.test(r.html || ''));
  check('Audit_Log now has the row', e.rows('AUDIT_LOG_SHEET') === 1);
  check('Dashboard_Data and Analysis have their rows', e.rows('DASHBOARD_DATA_SHEET') === 1 && e.rows('ANALYSIS_SHEET') === 1);
  check('no second Transcripts row was added', e.rows('TRANSCRIPTS_SHEET') === 1);
  check('the audit can now be found by the email step', e.ctx.findAuditRefForInteraction_(INT, 'repeats', SAP) === 'NHA-OLD-0001');
  check('the analyst is told it was recovered', (r.warnings || []).some(w => /recover|restor|already/i.test(w)), JSON.stringify(r.warnings));
  check('the processing flag is released', !flag(e));
}

console.log('\nTest 4 — same, but no Transcripts row survived: a new ref is made, everything is still recorded');
{
  const e = makeEnv();
  e.ctx.saveCachedResult_(INT, 'repeats', '<div class="report-wrap">cached</div>', SAP);
  const r = e.ctx.submitTranscript(FORM());
  check('success with a ref', r.success === true && /^NHA-TEST-/.test(r.auditRef || ''), JSON.stringify(r).slice(0, 160));
  check('AI not called', e.ai.calls === 0);
  check('all four sheets have their row', e.rows('TRANSCRIPTS_SHEET') === 1 && e.rows('ANALYSIS_SHEET') === 1 && e.rows('DASHBOARD_DATA_SHEET') === 1 && e.rows('AUDIT_LOG_SHEET') === 1);
}

console.log('\nTest 5 — a cache hit that DOES have an audit record behaves as before and releases the flag');
{
  const e = makeEnv();
  const first = e.ctx.submitTranscript(FORM());
  const aiAfterFirst = e.ai.calls;
  const before = [e.rows('TRANSCRIPTS_SHEET'), e.rows('ANALYSIS_SHEET'), e.rows('DASHBOARD_DATA_SHEET'), e.rows('AUDIT_LOG_SHEET')].join('/');
  const second = e.ctx.submitTranscript(FORM());
  check('success from cache', second.success === true && second.fromCache === true, JSON.stringify(second).slice(0, 160));
  check('same audit ref as the first time', second.auditRef === first.auditRef);
  check('no AI call and no new rows', e.ai.calls === aiAfterFirst && [e.rows('TRANSCRIPTS_SHEET'), e.rows('ANALYSIS_SHEET'), e.rows('DASHBOARD_DATA_SHEET'), e.rows('AUDIT_LOG_SHEET')].join('/') === before);
  check('THE BUG: the flag is released after a cache hit (it used to stay for 10 minutes)', !flag(e), 'flag=' + flag(e));
  const third = e.ctx.submitTranscript(FORM());
  check('an immediate resubmit is NOT told "already being processed"', third.success === true, JSON.stringify(third).slice(0, 160));
}

console.log('\nTest 6 — "Access denied" releases the flag');
{
  const e = makeEnv();
  e.sheets[e.ctx.TRANSCRIPTS_SHEET].appendRow = () => { throw new Error('You do not have permission to access the requested document'); };
  const r = e.ctx.submitTranscript(FORM());
  check('reports access denied', r.success === false && /Access denied/.test(r.error || ''), JSON.stringify(r));
  check('THE BUG: the flag is released (it used to stay for 10 minutes)', !flag(e), 'flag=' + flag(e));
}

console.log('\nTest 7 — a blank Interaction ID on the form is read from the transcript, and still guarded');
{
  const e = makeEnv();
  e.store['proc_' + INT + '_' + SAP] = String(Date.now());           // the first submit is still running
  const f = FORM(); f.interactionId = '';
  const r = e.ctx.submitTranscript(f);
  check('the second submit is refused as already being processed', r.success === false && /already being processed/.test(r.error || ''), JSON.stringify(r).slice(0, 200));
  check('and the AI was not called', e.ai.calls === 0);
  check('the other submit\'s flag is left alone', !!flag(e));
}

console.log('\nTest 8 — an AI failure writes no records, releases the flag, and can be retried');
{
  const e = makeEnv();
  e.ai.impl = () => { throw new Error('AI_SERVICE_ERROR: 503'); };
  const r = e.ctx.submitTranscript(FORM());
  check('friendly failure', r.success === false && /AI analysis service/.test(r.error || ''), JSON.stringify(r).slice(0, 200));
  check('no Analysis / Dashboard / Audit_Log rows', e.rows('ANALYSIS_SHEET') === 0 && e.rows('DASHBOARD_DATA_SHEET') === 0 && e.rows('AUDIT_LOG_SHEET') === 0);
  check('the flag is released', !flag(e));
  e.ai.impl = () => '<div class="report-wrap">second try</div>';
  const r2 = e.ctx.submitTranscript(FORM());
  check('the retry succeeds and records the audit', r2.success === true && e.rows('AUDIT_LOG_SHEET') === 1, JSON.stringify(r2).slice(0, 160));
}

console.log('\nTest 9 — end to end: the Audit_Log write fails once; the retry repairs it with the SAME ref and no new AI call');
{
  const e = makeEnv();
  const log = e.sheets[e.ctx.AUDIT_LOG_SHEET];
  const realAppend = log.appendRow.bind(log);
  log.appendRow = () => { throw new Error('Audit_Log unavailable'); };
  const first = e.ctx.submitTranscript(FORM());
  check('first attempt: success with a warning', first.success === true && (first.warnings || []).some(w => /Audit log/i.test(w)), JSON.stringify(first.warnings));
  check('first attempt: no Audit_Log row yet', e.rows('AUDIT_LOG_SHEET') === 0);
  const aiCalls = e.ai.calls;
  log.appendRow = realAppend;
  const second = e.ctx.submitTranscript(FORM());
  check('retry: success', second.success === true, JSON.stringify(second).slice(0, 160));
  check('retry: the SAME audit ref as the first attempt', second.auditRef === first.auditRef, first.auditRef + ' vs ' + second.auditRef);
  check('retry: no new AI call', e.ai.calls === aiCalls);
  check('retry: the Audit_Log row exists now', e.rows('AUDIT_LOG_SHEET') === 1);
  check('retry: no duplicate Transcripts / Analysis / Dashboard rows from the first attempt',
        e.rows('TRANSCRIPTS_SHEET') === 1 && e.rows('ANALYSIS_SHEET') === 1 && e.rows('DASHBOARD_DATA_SHEET') === 1,
        [e.rows('TRANSCRIPTS_SHEET'), e.rows('ANALYSIS_SHEET'), e.rows('DASHBOARD_DATA_SHEET')].join('/'));
}

console.log('\nTest 10 — a cut-off AI answer is refused: nothing is saved, the analyst is told why, and can submit again');
{
  const e = makeEnv();
  e.ai.impl = () => { throw new Error('AI_SERVICE_ERROR[truncated]: the AI answer was cut off by the token limit (finish_reason=length).'); };
  const r = e.ctx.submitTranscript(FORM());
  check('the submit fails with a message that says the report was cut off and to submit again', r.success === false && /cut off/.test(r.error) && /submit again/i.test(r.error), JSON.stringify(r));
  check('the message is not the generic "did not respond" one', !/did not respond/.test(r.error));
  check('no row was written anywhere', ['TRANSCRIPTS_SHEET', 'ANALYSIS_SHEET', 'DASHBOARD_DATA_SHEET', 'AUDIT_LOG_SHEET'].every(n => e.rows(n) === 0) || e.rows('DASHBOARD_DATA_SHEET') === 0 && e.rows('AUDIT_LOG_SHEET') === 0);
  check('nothing was cached as if it were a finished report', e.ctx.findCachedResult_(INT, 'repeats', SAP) === null);
  check('the processing flag is released, so the analyst can try again at once', !flag(e));
  check('the failure is logged with its reason', e.logs.some(l => /submitTranscript error: .*AI_SERVICE_ERROR\[truncated\]/.test(l)), e.logs.slice(-2).join(' | ').slice(0, 200));
  e.ai.impl = () => '<div class="report-wrap"><div class="ai-chip">ok</div></div>';
  check('the next attempt works normally', e.ctx.submitTranscript(FORM()).success === true);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
