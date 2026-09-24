/**
 * Regression test for openSheetWithRetry() (SheetService.gs).
 *
 * Reported symptom: analysts intermittently got "A lookup source could not
 * be opened... the file is shared with a different account" and the audit
 * was not saved — even though this deployment runs as "Execute as: Me" and
 * the owner has full access to every backend file. SpreadsheetApp.openById()
 * is known to intermittently throw a "You do not have permission..." error
 * under heavy concurrent access, even for a file's own owner. Since a
 * genuine per-user sharing gap doesn't fit this deployment's config, the fix
 * is to retry briefly on that specific error shape before giving up.
 *
 * Usage: node tests/sheet-retry-test.js [path-to-SheetService.gs]
 * Exits non-zero if any check fails.
 */
const fs = require('fs');
const vm = require('vm');
const path = require('path');

const gsPath = process.argv[2] || path.join(__dirname, '..', 'SheetService.gs');
const src = fs.readFileSync(gsPath, 'utf8');

function extractFunction(source, name) {
  const startMatch = source.match(new RegExp('function\\s+' + name + '\\s*\\('));
  if (!startMatch) throw new Error('function ' + name + ' not found in ' + gsPath);
  let i = startMatch.index;
  const braceStart = source.indexOf('{', i);
  let depth = 0, end = -1;
  for (let j = braceStart; j < source.length; j++) {
    if (source[j] === '{') depth++;
    else if (source[j] === '}') { depth--; if (depth === 0) { end = j + 1; break; } }
  }
  if (end === -1) throw new Error('could not find end of function ' + name);
  return source.slice(i, end);
}

const fnSource = extractFunction(src, 'openSheetWithRetry');

function run(SpreadsheetApp, extraGlobals) {
  const sleeps = [];
  const sandbox = Object.assign({
    SpreadsheetApp,
    Utilities: { sleep(ms) { sleeps.push(ms); } },
  }, extraGlobals);
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(fnSource, ctx);
  return { openSheetWithRetry: ctx.openSheetWithRetry, sleeps };
}

let failures = 0;
function check(name, cond) {
  console.log('  ' + (cond ? 'PASS' : 'FAIL') + '  ' + name);
  if (!cond) failures++;
}

console.log('NH Analyzer — openSheetWithRetry regression test');
console.log('file: ' + gsPath);

console.log('\nTest 1 — succeeds immediately with no retry when openById works');
{
  let calls = 0;
  const SpreadsheetApp = { openById() { calls++; return { id: 'ok' }; } };
  const { openSheetWithRetry, sleeps } = run(SpreadsheetApp);
  const result = openSheetWithRetry('sheet-id', 'Roster Sheet');
  check('returned the spreadsheet', result && result.id === 'ok');
  check('called openById exactly once', calls === 1);
  check('never slept', sleeps.length === 0);
}

console.log('\nTest 2 — retries a permission-like error and succeeds once it clears');
{
  let calls = 0;
  const SpreadsheetApp = {
    openById() {
      calls++;
      if (calls < 3) throw new Error('Exception: You do not have permission to access the requested document.');
      return { id: 'ok-after-retry' };
    }
  };
  const { openSheetWithRetry, sleeps } = run(SpreadsheetApp);
  const result = openSheetWithRetry('sheet-id', 'Roster Sheet');
  check('returned the spreadsheet on the 3rd attempt', result && result.id === 'ok-after-retry');
  check('called openById exactly 3 times', calls === 3);
  check('slept between retries (2 sleeps)', sleeps.length === 2);
}

console.log('\nTest 3 — does not retry a non-permission error (fails fast)');
{
  let calls = 0;
  const SpreadsheetApp = {
    openById() { calls++; throw new Error('Exception: Service Spreadsheets timed out while accessing document.'); }
  };
  const { openSheetWithRetry, sleeps } = run(SpreadsheetApp);
  let threw = null;
  try { openSheetWithRetry('sheet-id', 'Roster Sheet'); } catch (e) { threw = e; }
  check('threw', !!threw);
  check('called openById exactly once (no retry for unrelated errors)', calls === 1);
  check('never slept', sleeps.length === 0);
}

console.log('\nTest 4 — gives up after exhausting retries and names the file in the error');
{
  let calls = 0;
  const SpreadsheetApp = {
    openById() { calls++; throw new Error('Exception: You do not have permission to access the requested document.'); }
  };
  const { openSheetWithRetry, sleeps } = run(SpreadsheetApp);
  let threw = null;
  try { openSheetWithRetry('the-sheet-id', 'Roster Sheet'); } catch (e) { threw = e; }
  check('threw after exhausting retries', !!threw);
  check('called openById exactly 3 times', calls === 3);
  check('slept between attempts (2 sleeps)', sleeps.length === 2);
  check('error names the label', threw && threw.message.indexOf('Roster Sheet') !== -1);
  check('error names the id', threw && threw.message.indexOf('the-sheet-id') !== -1);
}

console.log(failures === 0 ? '\nALL TESTS PASSED' : '\n' + failures + ' CHECK(S) FAILED');
process.exit(failures === 0 ? 0 : 1);
