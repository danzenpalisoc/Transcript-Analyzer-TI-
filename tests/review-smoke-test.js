/**
 * Smoke test of the three fixes of the app-wide review (registry 188), run together on the REAL code:
 *
 *   1. an audit saved while repairSheetHeaders is rewriting a sheet is not wiped (the real
 *      repairSheetHeaders and the real submitTranscript run against each other, with a fake lock that,
 *      like the real one, cannot be taken twice);
 *   2. the Users directory the page receives holds names only, and the real index.html takes it;
 *   3. the registry accepts nothing without the Script Property, and not the old public secret.
 *
 * Only Google (sheets, session, cache, lock) and the AI are fake.
 *
 * Usage: node tests/review-smoke-test.js
 */
const vm = require('vm');
const { read, extractFunction, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const { load } = require('./lib/page-harness');
const check = makeCheck();
const codeSrc = read('Code.gs');

function makeEnv() {
  const store = {}, logs = [];
  const ai = { calls: 0 };
  const lock = { held: false, refused: 0, nested: 0 };
  const props = {};
  let refSeq = 0;
  const sandbox = {
    Logger: { log: m => logs.push(String(m)) },
    CacheService: { getScriptCache: () => ({ get: k => (k in store ? store[k] : null), put: (k, v) => { store[k] = v; }, remove: k => { delete store[k]; } }) },
    // Like the real lock: while another execution holds it, taking it waits and then times out, and an
    // execution can only release a lock it holds itself (the submit's duplicate guard calls releaseLock
    // after a failed wait, which must not free the repair's lock).
    LockService: { getScriptLock: () => {
      let mine = false;
      return {
        waitLock() { if (lock.held) { lock.refused++; throw new Error('Lock timeout'); } lock.held = true; mine = true; },
        releaseLock() { if (mine) { lock.held = false; mine = false; } },
      };
    } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null) }) },
    Utilities: { base64Encode: s => Buffer.from(String(s)).toString('base64'), sleep() {} },
    ScriptApp: { getScriptId: () => 'x' },
    Session: { getActiveUser: () => ({ getEmail: () => 'danzen.palisoc@telus.com' }) },   // an admin, who may run the repair
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ text: s, setMimeType() { return this; } }) },
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
  ['_procKey_', 'htmlToPlainText', 'findAuditRefForInteraction_', 'produceEvaluation_', 'writeAuditRecords_', '_cellSafeTranscript_', 'submitTranscript',
   'repairSheetHeaders', '_repairSheet', '_writeHeaderRow', 'getRegistrySecret_', 'doPost']
    .forEach(n => vm.runInContext(extractFunction(codeSrc, n), ctx));

  const sheets = {};
  [['TRANSCRIPTS_SHEET', 'TRANSCRIPTS_HEADERS'], ['ANALYSIS_SHEET', 'ANALYSIS_HEADERS'], ['DASHBOARD_DATA_SHEET', 'DASHBOARD_HEADERS'],
   ['AUDIT_LOG_SHEET', 'AUDIT_LOG_HEADERS'], ['CACHE_SHEET', 'CACHE_HEADERS']]
    .forEach(([s, h]) => { sheets[ctx[s]] = new FakeSheet(ctx[h], []); });
  const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new FakeSheet([], [])), getUrl: () => 'https://sheet' };
  ctx.getOrCreateSpreadsheet = () => ss;

  const users = new FakeSheet(['Name', 'Email Address', 'Role'], [
    ['Ana Cruz', 'ana.cruz@telus.com', 'QA'], ['Ben Reyes', 'ben.reyes@telus.com', 'Admin'],
  ]);
  users.getDataRange = () => ({ getValues: () => users.grid.map(r => r.slice()) });
  users.getName = () => 'Users'; users.getSheetId = () => ctx.USERS_SHEET_GID;
  ctx.openSheetWithRetry = id => (id === ctx.USERS_SS_ID ? { getSheets: () => [users], getSheetByName: () => users } : ss);

  const rows = n => sheets[ctx[n]].grid.length - 1;
  const all4 = () => ['TRANSCRIPTS_SHEET', 'ANALYSIS_SHEET', 'DASHBOARD_DATA_SHEET', 'AUDIT_LOG_SHEET'].map(rows).join('/');
  return { ctx, sheets, store, logs, ai, lock, props, rows, all4 };
}

const form = (ref, sap) => ({
  sapId: sap, transcript: 'INT:' + ref + ' agent: hello', interactionId: ref, participant: 'Agent ' + sap,
  teamLeader: 'TL', opsManager: 'OM', lineOfBusiness: 'LOB', locale: 'Site', vtid: 'V1',
  analysisType: 'repeats', observerName: 'QA One', agentEmail: 'a@telus.com',
});

console.log('\nSmoke 1 — an audit is saved while repairSheetHeaders is rewriting Dashboard_Data');
{
  const e = makeEnv();
  const ctx = e.ctx;
  // two audits already in the sheets
  check('audit A saved before the repair', ctx.submitTranscript(form('callA', '111')).success === true);
  check('audit B saved before the repair', ctx.submitTranscript(form('callB', '222')).success === true);
  const dashBefore = e.sheets[ctx.DASHBOARD_DATA_SHEET].grid.slice(1).map(r => r[0]).join(',');
  check('two rows in each sheet', e.all4() === '2/2/2/2', e.all4());

  // make the Dashboard_Data headers wrong (two columns swapped) so the real repair has work to do
  const dd = e.sheets[ctx.DASHBOARD_DATA_SHEET];
  const swapped = dd.grid[0].slice(); [swapped[3], swapped[4]] = [swapped[4], swapped[3]];
  dd.grid.forEach((row, i) => { if (i > 0) { const t = row[3]; row[3] = row[4]; row[4] = t; } });
  dd.grid[0] = swapped;

  // a third analyst saves while the repair has the sheet cleared (clearContents is the dangerous moment)
  const race = { result: null, rowsDuring: null };
  const orig = dd.clearContents.bind(dd);
  dd.clearContents = () => {
    orig();
    race.result = ctx.submitTranscript(form('callC', '333'));
    // Dashboard_Data has been cleared by the repair, so its grid is empty (no header, no rows): count its cells, not its rows
    race.rowsDuring = [e.rows('TRANSCRIPTS_SHEET'), e.rows('ANALYSIS_SHEET'), dd.grid.length, e.rows('AUDIT_LOG_SHEET')].join('/');
  };
  ctx.repairSheetHeaders();
  dd.clearContents = orig;

  check('the repair ran (Dashboard_Data headers are the code\'s again)', JSON.stringify(dd.grid[0].slice(0, ctx.DASHBOARD_HEADERS.length)) === JSON.stringify(ctx.DASHBOARD_HEADERS));
  check('the repair rewrote both earlier audits', dd.grid.slice(1).map(r => r[0]).join(',') === dashBefore, dd.grid.slice(1).map(r => r[0]).join(','));
  check('the save that arrived during the repair was refused, not lost silently', race.result && race.result.success === false && /repair/i.test(race.result.error || ''), JSON.stringify(race.result).slice(0, 200));
  check('nothing was written during the repair (Transcripts 2, Analysis 2, Dashboard_Data still empty, Audit_Log 2)', race.rowsDuring === '2/2/0/2', 'during the repair: ' + race.rowsDuring);
  check('the AI was not run for the refused audit', e.ai.calls === 2, 'AI calls: ' + e.ai.calls);
  check('the lock is free again', e.lock.held === false);

  const retry = ctx.submitTranscript(form('callC', '333'));
  check('the same call, submitted again after the repair, saves', retry.success === true, JSON.stringify(retry).slice(0, 160));
  check('every sheet has audits A, B and C once each', e.all4() === '3/3/3/3', e.all4());
  const refs = dd.grid.slice(1).map(r => r[0]);
  check('Dashboard_Data has three different refs', new Set(refs).size === 3 && refs.length === 3, refs.join(','));
  check('Auditor Email is filled on the new audit', dd.grid[3][ctx.DASHBOARD_HEADERS.indexOf('Auditor Email')] === 'danzen.palisoc@telus.com', String(dd.grid[3][ctx.DASHBOARD_HEADERS.indexOf('Auditor Email')]));
}

console.log('\nSmoke 2 — the Users directory reaches the real page as names only');
{
  const e = makeEnv();
  const list = e.ctx.getQAUsersFromSheet();
  check('the server returns names only', list.length === 2 && list.every(u => Object.keys(u).join() === 'name'), JSON.stringify(list));
  const page = load();
  let threw = '';
  try {
    page.ctx._loadQAUsers();
    const call = page.take('getQAUsersFromSheet');
    check('the page asks the server for the list', !!call);
    call.success(JSON.parse(JSON.stringify(list)));
  } catch (x) { threw = x.message; }
  check('the page takes the list without an error', threw === '', threw);
  const got = (page.ctx._dbQAUsers || []).map(u => u.name);
  check('the page holds both names', got.join('|') === 'Ana Cruz|Ben Reyes', got.join('|'));
  check('the page holds no address or role', !/@|Admin/.test(JSON.stringify(page.ctx._dbQAUsers || [])), JSON.stringify(page.ctx._dbQAUsers));
}

console.log('\nSmoke 3 — the registry takes nothing without its Script Property');
{
  const post = (e, key) => JSON.parse(e.ctx.doPost({ postData: { contents: JSON.stringify({ key, data: { issue: 'x' } }) } }).text);
  const e = makeEnv();
  let appended = 0; e.ctx.appendRegistryRow_ = () => { appended++; return { success: true }; };
  check('no property: the old public secret is refused', post(e, 'NH_REGISTRY_2026_DPZ').success === false && appended === 0);
  check('no property: an empty key is refused', post(e, '').success === false && appended === 0);
  e.props.REGISTRY_SECRET = 'a-new-secret';
  check('property set: the old public secret is still refused', post(e, 'NH_REGISTRY_2026_DPZ').success === false && appended === 0);
  check('property set: the new secret is accepted', post(e, 'a-new-secret').success === true && appended === 1);
}

console.log('\n' + (check.state.failures === 0 ? 'ALL PASSED' : check.state.failures + ' FAILED'));
process.exit(check.state.failures === 0 ? 0 : 1);
