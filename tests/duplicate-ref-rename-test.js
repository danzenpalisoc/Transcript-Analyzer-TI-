/**
 * Test for the admin tools that clean up audit references that two audits
 * already share (RefMaintenance.gs).
 *
 * Before the NH and TI scripts got separate prefixes, both handed out
 * NHA-YYYYMMDD-0001, -0002, ... against one spreadsheet, so Audit_Log holds
 * pairs of rows with the same reference. An email link carries only the
 * reference and opens the first row, so the later row's owner sees the
 * earlier row's agent.
 *
 * reportDuplicateAuditRefs() only reads. applyDuplicateAuditRefRenames()
 * keeps the earliest row's reference and gives every later row a new unique
 * one in Audit_Log, Dashboard_Data, Transcripts and Analysis, and records
 * every change in Ref_Rename_Log so it can be reversed.
 *
 * Usage: node tests/duplicate-ref-rename-test.js
 * Exits non-zero if any check fails.
 */
const vm = require('vm');
const { read, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

const who = { email: 'danzen.palisoc@telus.com' };   // an admin; a test can switch it

function makeEnv() {
  const removed = [];
  const sandbox = {
    Logger: { log() {} },
    CacheService: { getScriptCache: () => ({ get: () => null, put() {}, remove: k => removed.push(k) }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Utilities: { base64Encode: s => Buffer.from(String(s)).toString('base64'), sleep() {} },
    ScriptApp: { getScriptId: () => 'x' },
    // reportDuplicateAuditRefs / applyDuplicateAuditRefRenames are admin-only
    Session: { getActiveUser: () => ({ getEmail: () => (who.email) }) },
  };
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('Config.gs'), ctx);
  vm.runInContext(read('HtmlSanitizer.gs'), ctx);
  vm.runInContext(read('SheetService.gs'), ctx);
  vm.runInContext(read('RefMaintenance.gs'), ctx);
  const sheets = {};
  const mk = (name, headers) => (sheets[name] = new FakeSheet(headers, []));
  mk(ctx.AUDIT_LOG_SHEET, ctx.AUDIT_LOG_HEADERS);
  mk(ctx.DASHBOARD_DATA_SHEET, ctx.DASHBOARD_HEADERS);
  mk(ctx.TRANSCRIPTS_SHEET, ctx.TRANSCRIPTS_HEADERS);
  mk(ctx.ANALYSIS_SHEET, ['Audit Ref', 'Timestamp', 'SAP ID', 'Team Member', 'Observer Name', 'Analysis Type', 'AI Analysis Result']);
  const ss = {
    getSheetByName: n => sheets[n] || null,
    insertSheet: n => { const s = new FakeSheet([], []); s.grid = []; return (sheets[n] = s); },
  };
  ctx.getOrCreateSpreadsheet = () => ss;
  return { ctx, sheets, removed };
}
function put(env, sheetName, o) {
  const h = env.sheets[sheetName].grid[0];
  env.sheets[sheetName].appendRow(h.map(c => (o[c] !== undefined ? o[c] : '')));
}
function seed(env, ref, sap, intId, who) {
  const row = { 'Audit Ref': ref, 'SAP ID': sap, 'Interaction ID': intId, 'Team Member': who };
  [env.ctx.AUDIT_LOG_SHEET, env.ctx.DASHBOARD_DATA_SHEET, env.ctx.TRANSCRIPTS_SHEET, env.ctx.ANALYSIS_SHEET].forEach(n => put(env, n, row));
}
const refsIn = (env, name) => {
  const h = env.sheets[name].grid[0], c = h.indexOf('Audit Ref');
  return env.sheets[name].grid.slice(1).map(r => r[c]);
};

console.log('NH Analyzer — duplicate audit reference clean-up test');

console.log('\nTest 1 — the plan keeps the earliest row and renames only the later twins');
{
  const { ctx } = makeEnv();
  const rows = [
    { ref: 'NHA-1', sap: 'A', interactionId: 'i1' },
    { ref: 'NHA-1', sap: 'B', interactionId: 'i2' },
    { ref: 'NHA-2', sap: 'C', interactionId: 'i3' },
    { ref: 'NHA-1', sap: 'D', interactionId: 'i4' },
  ];
  const plan = ctx._planDuplicateRefRenames(rows);
  const byIndex = {}; plan.forEach(p => { byIndex[p.index] = p; });
  check('exactly the two later rows are planned', plan.length === 2 && byIndex[1] && byIndex[3], JSON.stringify(plan));
  check('the earliest row (index 0) is not renamed', !byIndex[0]);
  check('a reference used once is not renamed', !byIndex[2]);
  check('each new reference is unique', byIndex[1].newRef !== byIndex[3].newRef && byIndex[1].newRef !== 'NHA-1', JSON.stringify(plan));
  check('the new reference keeps the original as its start', /^NHA-1-/.test(byIndex[1].newRef), byIndex[1].newRef);
}

console.log('\nTest 2 — a new reference never collides with one that already exists');
{
  const { ctx } = makeEnv();
  const first = ctx._planDuplicateRefRenames([{ ref: 'NHA-1', sap: 'A' }, { ref: 'NHA-1', sap: 'B' }])[0].newRef;
  const second = ctx._planDuplicateRefRenames([
    { ref: 'NHA-1', sap: 'A' }, { ref: 'NHA-1', sap: 'B' }, { ref: first, sap: 'Z' },
  ])[0].newRef;
  check('an already-used new reference is skipped', second !== first, first + ' vs ' + second);
}

console.log('\nTest 3 — twins that cannot be told apart are reported, not renamed');
{
  const { ctx } = makeEnv();
  const plan = ctx._planDuplicateRefRenames([
    { ref: 'NHA-1', sap: 'A', interactionId: 'i1' }, { ref: 'NHA-1', sap: 'A', interactionId: 'i1' },
  ]);
  check('the later row is flagged for manual review', plan.length === 1 && plan[0].manual === true, JSON.stringify(plan));
}

console.log('\nTest 4 — report only reads');
{
  const env = makeEnv();
  seed(env, 'NHA-1', 'A', 'i1', 'Fatima'); seed(env, 'NHA-1', 'B', 'i2', 'Rose');
  const before = JSON.stringify(Object.keys(env.sheets).map(k => env.sheets[k].grid));
  const rep = env.ctx.reportDuplicateAuditRefs();
  check('it reports the one pair', rep && rep.duplicateRefs === 1 && rep.renames === 1, JSON.stringify(rep));
  check('no sheet was changed', JSON.stringify(Object.keys(env.sheets).map(k => env.sheets[k].grid)) === before);
  check('no log sheet was created', !env.sheets[env.ctx.REF_RENAME_LOG_SHEET]);
}

console.log('\nTest 5 — apply renames the later twin in every sheet and logs it');
{
  const env = makeEnv();
  seed(env, 'NHA-1', 'A', 'i1', 'Fatima'); seed(env, 'NHA-1', 'B', 'i2', 'Rose'); seed(env, 'NHA-2', 'C', 'i3', 'Other');
  const counts = Object.keys(env.sheets).map(k => env.sheets[k].grid.length);
  const res = env.ctx.applyDuplicateAuditRefRenames();
  check('one row was renamed', res && res.renamed === 1, JSON.stringify(res));
  [env.ctx.AUDIT_LOG_SHEET, env.ctx.DASHBOARD_DATA_SHEET, env.ctx.TRANSCRIPTS_SHEET, env.ctx.ANALYSIS_SHEET].forEach(n => {
    const refs = refsIn(env, n);
    check(n + ': the earliest keeps NHA-1', refs[0] === 'NHA-1', JSON.stringify(refs));
    check(n + ': the later twin has a new reference', refs[1] !== 'NHA-1' && /^NHA-1-/.test(refs[1]), JSON.stringify(refs));
    check(n + ': an unrelated reference is untouched', refs[2] === 'NHA-2');
  });
  const log = env.sheets[env.ctx.REF_RENAME_LOG_SHEET];
  check('the change is recorded in the log sheet', log && log.grid.length === 2 && log.grid[1].indexOf('NHA-1') !== -1, JSON.stringify(log && log.grid));
  check('the log holds the old and new reference and the SAP ID',
        log && log.grid[1].indexOf('B') !== -1 && log.grid[1].some(v => /^NHA-1-/.test(v)));
  check('the stale evaluation-page cache entry for the old reference is cleared',
        env.removed.indexOf('ev2_' + Buffer.from('NHA-1').toString('base64')) !== -1, JSON.stringify(env.removed));
  check('no rows were added or removed', JSON.stringify(Object.keys(env.sheets).filter(k => k !== env.ctx.REF_RENAME_LOG_SHEET).map(k => env.sheets[k].grid.length)) === JSON.stringify(counts));
}

console.log('\nTest 5b — apply never reads the big text columns and reads each sheet a bounded number of times');
{
  const env = makeEnv();
  for (let k = 0; k < 30; k++) { seed(env, 'NHA-' + k, 'A' + k, 'i' + k, 'First ' + k); seed(env, 'NHA-' + k, 'B' + k, 'j' + k, 'Second ' + k); }
  const tr = env.sheets[env.ctx.TRANSCRIPTS_SHEET], an = env.sheets[env.ctx.ANALYSIS_SHEET];
  const transcriptCol = env.ctx.TRANSCRIPTS_HEADERS.indexOf('Transcript') + 1;
  Object.keys(env.sheets).forEach(k => env.sheets[k].resetReads());
  const res = env.ctx.applyDuplicateAuditRefRenames();
  check('all 30 pairs were renamed', res.renamed === 30, JSON.stringify(res));
  check('the Transcripts text column was never read', tr.cellsReadInColumn(transcriptCol, 2) === 0, 'read ' + tr.cellsReadInColumn(transcriptCol, 2) + ' cells');
  check('the Analysis result column was never read', an.cellsReadInColumn(7, 2) === 0, 'read ' + an.cellsReadInColumn(7, 2) + ' cells');
  check('Transcripts is read a fixed number of times, not once per rename', tr.reads.length <= 6, 'reads: ' + tr.reads.length);
  check('Analysis is read a fixed number of times, not once per rename', an.reads.length <= 6, 'reads: ' + an.reads.length);
}

console.log('\nTest 6 — running apply again changes nothing');
{
  const env = makeEnv();
  seed(env, 'NHA-1', 'A', 'i1', 'Fatima'); seed(env, 'NHA-1', 'B', 'i2', 'Rose');
  env.ctx.applyDuplicateAuditRefRenames();
  const snapshot = JSON.stringify(Object.keys(env.sheets).map(k => env.sheets[k].grid));
  const again = env.ctx.applyDuplicateAuditRefRenames();
  check('nothing is renamed the second time', again.renamed === 0, JSON.stringify(again));
  check('no sheet changed the second time', JSON.stringify(Object.keys(env.sheets).map(k => env.sheets[k].grid)) === snapshot);
}

console.log('\nTest — an ordinary analyst cannot rename audit references or list them');
{
  const env = makeEnv();
  put(env, env.ctx.AUDIT_LOG_SHEET, { 'Audit Ref': 'NHA-1', 'SAP ID': '111' });
  put(env, env.ctx.AUDIT_LOG_SHEET, { 'Audit Ref': 'NHA-1', 'SAP ID': '222' });
  const before = JSON.stringify(Object.keys(env.sheets).map(k => env.sheets[k].grid));
  who.email = 'some.analyst@telus.com';
  let applyThrew = false, reportThrew = false;
  try { env.ctx.applyDuplicateAuditRefRenames(); } catch (e) { applyThrew = /admin/i.test(e.message); }
  try { env.ctx.reportDuplicateAuditRefs(); } catch (e) { reportThrew = /admin/i.test(e.message); }
  who.email = 'danzen.palisoc@telus.com';
  check('applyDuplicateAuditRefRenames is refused for a non-admin', applyThrew);
  check('reportDuplicateAuditRefs is refused for a non-admin', reportThrew);
  check('and nothing was written', JSON.stringify(Object.keys(env.sheets).map(k => env.sheets[k].grid)) === before);
}

console.log(check.state.failures === 0 ? '\nALL TESTS PASSED' : '\n' + check.state.failures + ' CHECK(S) FAILED');
process.exit(check.state.failures === 0 ? 0 : 1);
