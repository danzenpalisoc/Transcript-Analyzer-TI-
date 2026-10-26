/**
 * Saving an audit holds the script lock, the same lock the sheet repairs take.
 *
 * repairSheetHeaders reads a whole sheet, clears it and writes it back from that snapshot, so a row
 * appended while it runs was wiped (the audit then had no Audit_Log / Dashboard_Data row while the
 * report and the email had already gone out). The comment on the repair said the audit save took the
 * same lock; it did not. Now the Transcripts row and the three audit rows are written under it.
 *
 * The fake lock is NOT re-entrant (like the real one): taking it while it is held fails, so a nested
 * take (ensureHeaders inside the save) would show up here.
 *
 * Usage: node tests/audit-save-lock-test.js
 */
const vm = require('vm');
const { read, extractFunction, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const codeSrc = read('Code.gs');

function makeEnv() {
  const store = {}, logs = [];
  const ai = { calls: 0 };
  const lock = { held: false, busy: false, takes: 0, refused: 0, nested: 0 };
  let refSeq = 0;
  const sandbox = {
    Logger: { log: m => logs.push(String(m)) },
    CacheService: { getScriptCache: () => ({ get: k => (k in store ? store[k] : null), put: (k, v) => { store[k] = v; }, remove: k => { delete store[k]; } }) },
    LockService: { getScriptLock: () => ({
      waitLock() {
        if (lock.busy) { lock.refused++; throw new Error('Lock timeout'); }
        if (lock.held) { lock.nested++; throw new Error('Lock already held'); }
        lock.held = true; lock.takes++;
      },
      releaseLock() { lock.held = false; },
    }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Utilities: { base64Encode: s => Buffer.from(String(s)).toString('base64'), sleep() {} },
    ScriptApp: { getScriptId: () => 'x' },
    Session: { getActiveUser: () => ({ getEmail: () => 'qa.one@telus.com' }) },
    parseTranscriptMetadata: txt => { const m = /INT:(\S+)/.exec(txt || ''); return { interactionId: m ? m[1] : '' }; },
    filterTranscriptByAgent: t => ({ filtered: t, stats: { applied: false, targetFound: false } }),
    translateTranscriptText_: t => t,
    analyzeTranscript_: () => { ai.calls++; return '<div class="report-wrap"><div class="ai-chip">ok</div></div>'; },
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

  const sheets = {}, heldAtAppend = {};
  [['TRANSCRIPTS_SHEET', 'TRANSCRIPTS_HEADERS'], ['ANALYSIS_SHEET', 'ANALYSIS_HEADERS'], ['DASHBOARD_DATA_SHEET', 'DASHBOARD_HEADERS'],
   ['AUDIT_LOG_SHEET', 'AUDIT_LOG_HEADERS'], ['CACHE_SHEET', 'CACHE_HEADERS']]
    .forEach(([s, h]) => {
      const sh = new FakeSheet(ctx[h], []);
      const orig = sh.appendRow.bind(sh);
      heldAtAppend[ctx[s]] = [];
      sh.appendRow = arr => { heldAtAppend[ctx[s]].push(lock.held); orig(arr); };
      sheets[ctx[s]] = sh;
    });
  const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new FakeSheet([], [])), getUrl: () => 'https://sheet' };
  ctx.getOrCreateSpreadsheet = () => ss;
  const rows = name => sheets[ctx[name]].grid.length - 1;
  const all4 = () => [rows('TRANSCRIPTS_SHEET'), rows('ANALYSIS_SHEET'), rows('DASHBOARD_DATA_SHEET'), rows('AUDIT_LOG_SHEET')].join('/');
  return { ctx, sheets, store, logs, ai, lock, rows, all4, heldAtAppend };
}

const SAP = '111', INT = 'abc';
const FORM = () => ({
  sapId: SAP, transcript: 'INT:' + INT + ' agent: hello', interactionId: INT, participant: 'Agent A',
  teamLeader: 'TL', opsManager: 'OM', lineOfBusiness: 'LOB', locale: 'Site', vtid: 'V1',
  analysisType: 'repeats', observerName: 'QA One', agentEmail: 'a@telus.com',
});
const flag = e => e.store['proc_' + INT + '_' + SAP];

console.log('\nTest 1 — every row of an audit is written while the lock is held, and the lock is released');
{
  const e = makeEnv();
  const r = e.ctx.submitTranscript(FORM());
  check('success', r.success === true, JSON.stringify(r).slice(0, 160));
  check('one row in each of the four sheets', e.all4() === '1/1/1/1', e.all4());
  const h = e.heldAtAppend;
  check('Transcripts row written under the lock', h[e.ctx.TRANSCRIPTS_SHEET].length === 1 && h[e.ctx.TRANSCRIPTS_SHEET].every(Boolean), JSON.stringify(h[e.ctx.TRANSCRIPTS_SHEET]));
  check('Analysis row written under the lock', h[e.ctx.ANALYSIS_SHEET].every(Boolean) && h[e.ctx.ANALYSIS_SHEET].length === 1);
  check('Dashboard_Data row written under the lock', h[e.ctx.DASHBOARD_DATA_SHEET].every(Boolean) && h[e.ctx.DASHBOARD_DATA_SHEET].length === 1);
  check('Audit_Log row written under the lock', h[e.ctx.AUDIT_LOG_SHEET].every(Boolean) && h[e.ctx.AUDIT_LOG_SHEET].length === 1);
  check('the lock is released at the end', e.lock.held === false);
  check('nothing tried to take the lock while it was held (no nested take)', e.lock.nested === 0, 'nested: ' + e.lock.nested);
  check('no warnings for the analyst', !(r.warnings || []).length, JSON.stringify(r.warnings));
  check('the processing flag is released', !flag(e));
}

console.log('\nTest 2 — a repair holds the lock for the whole call: nothing is saved, the AI is not called, the analyst is told');
{
  const e = makeEnv();
  e.lock.busy = true;
  const r = e.ctx.submitTranscript(FORM());
  check('not a success', r.success === false, JSON.stringify(r).slice(0, 160));
  check('the message says it is being repaired and to submit again', /repair/i.test(r.error || '') && /again/i.test(r.error || ''), r.error);
  check('nothing was written anywhere', e.all4() === '0/0/0/0', e.all4());
  check('the AI was not called', e.ai.calls === 0, 'AI calls: ' + e.ai.calls);
  check('the processing flag is released', !flag(e));
  e.lock.busy = false;
  const again = e.ctx.submitTranscript(FORM());
  check('once the repair is over the same call saves normally', again.success === true && e.all4() === '1/1/1/1', e.all4());
}

console.log('\nTest 3 — a repair starts AFTER the AI ran: no audit rows, report kept, a second submit completes them once');
{
  const e = makeEnv();
  const tr = e.sheets[e.ctx.TRANSCRIPTS_SHEET], orig = tr.appendRow.bind(tr);
  tr.appendRow = arr => { orig(arr); e.lock.busy = true; };   // the repair takes the lock right after the Transcripts row
  const first = e.ctx.submitTranscript(FORM());
  check('first attempt: not a success', first.success === false, JSON.stringify(first).slice(0, 160));
  check('the message says the report is kept', /report is kept/i.test(first.error || ''), first.error);
  check('only the Transcripts row exists; no audit rows were written', e.all4() === '1/0/0/0', e.all4());
  check('the AI ran once and the report is cached', e.ai.calls === 1 && e.ctx.findCachedResult_(INT, 'repeats', SAP) !== null, 'AI calls: ' + e.ai.calls);
  check('the processing flag is released', !flag(e));

  e.lock.busy = false;
  const second = e.ctx.submitTranscript(FORM());
  check('second attempt: success', second.success === true, JSON.stringify(second).slice(0, 160));
  check('the AI was NOT called again', e.ai.calls === 1, 'AI calls: ' + e.ai.calls);
  check('every sheet has exactly one row (no duplicate Transcripts row)', e.all4() === '1/1/1/1', e.all4());
  check('the audit ref of the first attempt is reused', second.auditRef === e.sheets[e.ctx.TRANSCRIPTS_SHEET].grid[1][0], second.auditRef);
  check('the lock is released', e.lock.held === false);
}

console.log('\nTest 4 — the lock is released even when a write throws');
{
  const e = makeEnv();
  e.sheets[e.ctx.DASHBOARD_DATA_SHEET].appendRow = () => { throw new Error('Dashboard sheet unavailable'); };
  const r = e.ctx.submitTranscript(FORM());
  check('still a success, with a warning naming the ref', r.success === true && (r.warnings || []).some(w => /Dashboard/i.test(w) && w.indexOf(r.auditRef) !== -1), JSON.stringify(r.warnings));
  check('the lock is released', e.lock.held === false);
  const next = e.ctx.submitTranscript(Object.assign(FORM(), { interactionId: 'other', transcript: 'INT:other agent: hi' }));
  check('the next audit is not blocked by a lock left behind', next.success === true && e.lock.nested === 0, JSON.stringify(next).slice(0, 120));
}

console.log('\nTest 5 — the helper itself');
{
  const e = makeEnv();
  let ran = false;
  const ok = e.ctx._withSheetLock_(1000, () => { ran = true; return 42; });
  check('runs fn under the lock and returns its value', ok.locked === true && ok.value === 42 && ran && e.lock.held === false);
  e.lock.busy = true; ran = false;
  const no = e.ctx._withSheetLock_(1000, () => { ran = true; });
  check('lock not got: fn is not run', no.locked === false && ran === false);
  e.lock.busy = false;
  let threw = false;
  try { e.ctx._withSheetLock_(1000, () => { throw new Error('boom'); }); } catch (x) { threw = true; }
  check('fn throws: the error passes through and the lock is released', threw && e.lock.held === false);
}

console.log('\n' + (check.state.failures === 0 ? 'ALL PASSED' : check.state.failures + ' FAILED'));
process.exit(check.state.failures === 0 ? 0 : 1);
