/**
 * Admin repair tools must not destroy data (code review batch 3).
 *
 *  1. repairSheetHeaders cleared each sheet and rewrote only the columns the code knows, so an older column (for example
 *     Recommendations) was deleted for good, and it ran without a lock, so an audit saved in between was wiped.
 *  2. repairNextTruncated deleted every Sales evaluation of an Interaction ID, including the other agents' who share it.
 *
 * Usage: node tests/admin-tools-test.js
 */
const vm = require('vm');
const { read, extractFunction, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const codeSrc = read('Code.gs');

console.log('\nTest 1 — repairSheetHeaders keeps every column');
{
  const logs = [], lockLog = { waited: 0, released: 0, failWait: false };
  const sandbox = { Logger: { log: m => logs.push(String(m)) }, JSON, Math, String, Array, Object, Error,
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    LockService: { getScriptLock: () => ({ waitLock() { lockLog.waited++; if (lockLog.failWait) throw new Error('Lock timeout'); }, releaseLock() { lockLog.released++; } }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'danzen.palisoc@telus.com' }) } };
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('Config.gs'), ctx);
  ['repairSheetHeaders', '_repairSheet', '_writeHeaderRow'].forEach(n => vm.runInContext(extractFunction(codeSrc, n), ctx));
  const H = ctx.AUDIT_LOG_HEADERS;
  const mk = () => {
    // an old layout: two known columns swapped, an old column (Recommendations), a column with no header that holds data, and an empty one
    const hdr = [H[1], H[0], 'Recommendations', '', ''].concat([]);
    const sheet = new FakeSheet(hdr, [['2026-10-13', 'NHA-1', 'Call back in 2 days', 'note', ''], ['2026-10-12', 'NHA-2', '', '', '']]);
    return sheet;
  };
  const ss = sheets => ({ getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new FakeSheet([], [])) });
  const sh = mk(); const sheets = {}; sheets[ctx.AUDIT_LOG_SHEET] = sh;
  ctx._repairSheet(ss(sheets), ctx.AUDIT_LOG_SHEET, H);
  const hdr = sh.grid[0];
  check('the standard headers come first, in the right order', H.every((h, i) => hdr[i] === h), JSON.stringify(hdr.slice(0, 3)));
  check('the old Recommendations column is KEPT after them, with its header', hdr[H.length] === 'Recommendations', JSON.stringify(hdr.slice(H.length)));
  const col = hdr.indexOf('Recommendations');
  check('and its data is still there, on the right rows', sh.grid[1][col] === 'Call back in 2 days' && sh.grid[2][col] === '', JSON.stringify([sh.grid[1][col], sh.grid[2][col]]));
  check('a column with no header that holds data is kept', hdr.length === H.length + 2 && sh.grid[1][hdr.length - 1] === 'note', JSON.stringify(hdr.slice(H.length)));
  check('an empty column with no header is dropped', hdr.length === H.length + 2);
  check('the known columns moved with their data (Audit Ref NHA-1 is under Audit Ref)', sh.grid[1][hdr.indexOf('Audit Ref')] === 'NHA-1' && sh.grid[1][hdr.indexOf('Submitted At')] === '2026-10-13');
  check('the repair says what it kept', logs.some(l => /kept 2 extra column/.test(l)), logs.join(' | '));
  const before = JSON.stringify(sh.grid);
  ctx._repairSheet(ss(sheets), ctx.AUDIT_LOG_SHEET, H);
  check('running it again changes nothing (headers already correct, extra columns after them)', JSON.stringify(sh.grid) === before);

  ctx.repairSheetHeaders = ctx.repairSheetHeaders;
  ctx.getOrCreateSpreadsheet = () => ss({});
  ctx.invalidateDashboardCache_ = () => {};
  ctx.repairSheetHeaders();
  check('repairSheetHeaders takes the script lock and releases it', lockLog.waited === 1 && lockLog.released === 1, JSON.stringify(lockLog));
  lockLog.failWait = true; lockLog.waited = 0; lockLog.released = 0;
  const target = ss({}); let touched = false;
  const real = target.insertSheet; target.insertSheet = n => { touched = true; return real(n); };
  ctx.getOrCreateSpreadsheet = () => target;
  let err = '';
  try { ctx.repairSheetHeaders(); } catch (e) { err = e.message; }
  check('when the lock cannot be had it stops and touches nothing', /Lock timeout/.test(err) && !touched, err);
}

console.log('\nTest 2 — repairNextTruncated only touches the one agent\'s evaluation');
{
  const GOOD = '<div class="call-summary">Call Summary ' + 'x'.repeat(600) + '</div>';
  function env(cacheRows) {
    const logs = [], saved = [], analyzed = [];
    const cache = new FakeSheet(['Interaction ID', 'Type', 'Created', 'Result', 'SAP ID'], cacheRows);
    const trans = new FakeSheet(['Interaction ID', 'Type', 'Created', 'Transcript', 'SAP ID'], [['INT-1', 't', '', 'Agent: hi', ''], ['INT-2', 't', '', 'Agent: hello', '']]);
    const ssObj = { sheets: {}, getSheetByName(n) { return this.sheets[n] || null; } };
    const sandbox = { Logger: { log: m => logs.push(String(m)) }, JSON, Math, String, Array, Object, Error, RegExp,
      PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
      Session: { getActiveUser: () => ({ getEmail: () => 'danzen.palisoc@telus.com' }) },
      SpreadsheetApp: { flush() {} } };
    const ctx = vm.createContext(sandbox);
    vm.runInContext(read('Config.gs'), ctx);
    ['diagnoseSalesEvaluations', 'repairNextTruncated'].forEach(n => vm.runInContext(extractFunction(codeSrc, n), ctx));
    ssObj.sheets[ctx.CACHE_SHEET] = cache; ssObj.sheets[ctx.TRANSCRIPTS_SHEET] = trans;
    ctx.getOrCreateSpreadsheet = () => ssObj;
    ctx.getOrCreateSheet = (s, n) => s.sheets[n];
    ctx.analyzeTranscript_ = (t, type) => { analyzed.push(t); return GOOD; };
    ctx.fixBadgeClasses = h => h;
    ctx.saveCachedResult_ = (id, type, html, sap) => { saved.push({ id, type, sap }); };
    ctx._clearEv2CacheForInteraction_ = () => {};
    return { ctx, cache, logs, saved, analyzed };
  }
  const cut = '<div>Call Summary half'; // truncated: too short, not ending cleanly
  // one agent only
  let e = env([['INT-1', 'sales', '', cut, '111'], ['INT-1', 'repeats', '', 'keep me', '111'], ['INT-1', 'sales', '', 'x', '222']].slice(0, 2));
  e.ctx.repairNextTruncated();
  check('a single-agent interaction is repaired and saved for that agent (SAP 111)', e.saved.length === 1 && e.saved[0].sap === '111' && e.saved[0].id === 'INT-1', JSON.stringify(e.saved));
  check('its old sales row is gone but the Repeats row of the same interaction is kept', e.cache.grid.every(r => r[1] !== 'sales') && e.cache.grid.some(r => r[3] === 'keep me'), JSON.stringify(e.cache.grid));

  // two agents share the interaction
  e = env([['INT-1', 'sales', '', cut, '111'], ['INT-1', 'sales', '', GOOD, '222']]);
  const before = JSON.stringify(e.cache.grid);
  e.ctx.repairNextTruncated();
  check('with two agents on the interaction nothing is deleted, nothing is asked of the AI, nothing is saved', JSON.stringify(e.cache.grid) === before && e.analyzed.length === 0 && e.saved.length === 0, JSON.stringify({ a: e.analyzed.length, s: e.saved.length }));
  check('the log says to repair it by hand', e.logs.some(l => /several agents share this interaction/.test(l)), e.logs.join(' | '));

  // an orphan chunk of ANOTHER agent (its main row is gone) carries the same Interaction ID: only this agent's rows are deleted
  e = env([['INT-1', 'sales', '', cut, '111'], ['INT-1', 'sales_2', '', 'orphan chunk of agent 222', '222']]);
  e.ctx.repairNextTruncated();
  check('the other agent\'s orphan chunk row is NOT deleted (rows are matched on the SAP ID too)', e.cache.grid.some(r => r[3] === 'orphan chunk of agent 222') && !e.cache.grid.some(r => r[3] === cut), JSON.stringify(e.cache.grid.map(r => r[3])));

  // the first truncated one is shared, the second is not: the second is repaired
  e = env([['INT-1', 'sales', '', cut, '111'], ['INT-1', 'sales', '', cut, '222'], ['INT-2', 'sales', '', cut, '333']]);
  e.ctx.repairNextTruncated();
  check('a shared one is skipped and the next safe one is repaired', e.saved.length === 1 && e.saved[0].id === 'INT-2' && e.saved[0].sap === '333', JSON.stringify(e.saved));
  check('the shared interaction\'s two rows are untouched', e.cache.grid.filter(r => r[0] === 'INT-1').length === 2);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
