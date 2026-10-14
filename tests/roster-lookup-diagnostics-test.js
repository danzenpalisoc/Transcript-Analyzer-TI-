/**
 * getRosterBySapId() must tell the page WHY it has nothing, not just return null.
 *
 * Before: every failure path (roster sheet unreadable, permission blip, timeout)
 * was swallowed and returned null, and the page treats null as "this agent is not
 * on the roster" and leaves the form blank without a word. An analyst with a
 * perfectly valid SAP ID saw exactly the same thing as one with a typo, and the
 * cause could not be told apart even in the Executions log.
 *
 * Contract now:
 *   row found                       -> the row, unchanged
 *   nothing found, nothing failed   -> null (genuinely not on the roster)
 *   nothing found, a source failed  -> { _error: 'source_unavailable', sources: [...] }
 *   lookupBySapId_ threw             -> { _error: 'lookup_failed' }
 *
 * Usage: node tests/roster-lookup-diagnostics-test.js
 */
const vm = require('vm');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');

const code = read('Code.gs');
const ai   = read('AIService.gs');
const check = makeCheck();

const varLine = (ai.match(/^var _lookupSourceErrors\s*=\s*\[\];/m) || [''])[0];
const helperSrc = [
  varLine,
  extractFunction(ai, '_noteSourceError_'),
  extractFunction(ai, '_resetSourceErrors_'),
  extractFunction(code, 'getRosterBySapId'),
].join('\n');

function sandbox(extra) {
  const logs = [];
  const ctx = vm.createContext(Object.assign({
    Logger: { log: m => logs.push(String(m)) }, JSON, Error, Array, String,
  }, extra));
  vm.runInContext(helperSrc, ctx);
  ctx.__logs = logs;
  return ctx;
}

console.log('\nTest 1 — a found row is returned unchanged');
{
  const row = { participant: 'A Agent', teamLeader: 'TL' };
  const ctx = sandbox({ lookupBySapId_: () => row });
  check('same row', ctx.getRosterBySapId('243141') === row);
}

console.log('\nTest 2 — genuinely not on the roster stays null');
{
  const ctx = sandbox({ lookupBySapId_: () => null });
  check('null', ctx.getRosterBySapId('999999') === null);
}

console.log('\nTest 3 — not found because a source could not be read is reported');
{
  const ctx = sandbox({});
  ctx.lookupBySapId_ = () => { ctx._noteSourceError_('Roster', new Error('permission')); return null; };
  const r = ctx.getRosterBySapId('243141');
  check('is an error object', !!r && r._error === 'source_unavailable', JSON.stringify(r));
  check('names the unreadable source', !!r && r.sources && r.sources.join() === 'Roster', JSON.stringify(r));
}

console.log('\nTest 4 — a hit wins even if a lower tier could not be read');
{
  const ctx = sandbox({});
  const row = { participant: 'A Agent' };
  ctx.lookupBySapId_ = () => { ctx._noteSourceError_('Global Roster', new Error('x')); return row; };
  check('row still returned', ctx.getRosterBySapId('243141') === row);
}

console.log('\nTest 5 — an exception is reported, not swallowed to null');
{
  const ctx = sandbox({ lookupBySapId_: () => { throw new Error('boom'); } });
  const r = ctx.getRosterBySapId('243141');
  check('lookup_failed', !!r && r._error === 'lookup_failed', JSON.stringify(r));
  check('logged', ctx.__logs.some(l => /boom/.test(l)));
}

console.log('\nTest 6 — failures from one call do not leak into the next');
{
  const ctx = sandbox({});
  ctx.lookupBySapId_ = () => { ctx._noteSourceError_('Roster', 'x'); return null; };
  ctx.getRosterBySapId('1');
  ctx.lookupBySapId_ = () => null;
  check('second call is a plain null', ctx.getRosterBySapId('2') === null);
}

console.log('\nTest 7 — the sheet loaders record their failure instead of hiding it');
for (const [fn, label] of [['_getRosterSheetData', 'Roster'], ['_getGlobalRosterData', 'Global Roster'],
                           ['_getTraineeRosterData', 'Trainee Roster'], ['_getATDataGCPSheetData', 'AT Data']]) {
  const ctx = sandbox({
    CacheService: { getScriptCache: () => ({ get: () => null, put() {} }) },
    openSheetWithRetry: () => { throw new Error('You do not have permission'); },
    ROSTER_SHEET_ID: 'a', GLOBAL_ROSTER_SS_ID: 'b', TRAINEE_ROSTER_SS_ID: 'c', AT_DATA_GCP_SS_ID: 'd',
  });
  vm.runInContext('var _rosterDataInMemory=null,_globalRosterDataInMemory=null,_traineeRosterInMemory=null,_atDataInMemory=null;' +
                  extractFunction(ai, fn), ctx);
  let threw = false;
  try { ctx[fn](); } catch (e) { threw = true; }
  check(fn + ' does not throw', !threw);
  check(fn + ' notes "' + label + '"', ctx._lookupSourceErrors.indexOf(label) !== -1, JSON.stringify(ctx._lookupSourceErrors));
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
