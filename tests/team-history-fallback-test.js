/**
 * Regression test for getTeamDetailsFromAuditHistory() (SheetService.gs).
 *
 * Idea (from the admin, in response to the "lookup source" investigation):
 * if the live Roster/AT-Data/Global-Roster lookup fails or leaves a field
 * blank, fall back to this SAP ID's most recent PRIOR entry in Audit_Log —
 * which lives in the analyzer's own spreadsheet, not an external file — so
 * the submission can still succeed instead of failing outright. Must never
 * override a field the live lookup DID resolve, must pick the MOST RECENT
 * prior entry (not the oldest), and must return null (not throw) when there
 * is no prior history for that SAP ID.
 *
 * Usage: node tests/team-history-fallback-test.js [path-to-SheetService.gs]
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

const fnSource = extractFunction(src, 'getTeamDetailsFromAuditHistory');

const AUDIT_LOG_HEADERS = [
  'Audit Ref', 'Submitted At', 'Interaction ID', 'SAP ID', 'Team Member',
  'Team Leader', 'Operations Manager', 'Line of Business', 'Locale', 'VTID',
  'Observer Name', 'Analysis Type', 'Direction', 'Duration',
  'Repeat Projection %', 'Issue Resolved', 'Transfer Occurred',
  'Email Status', 'Recipients', 'Agent Email'
];

// Builds a fake Audit_Log sheet + spreadsheet pair from an array of row
// objects (only the fields under test need to be set; the rest default to '').
function makeAuditLogSheet(rows) {
  const grid = rows.map(r => AUDIT_LOG_HEADERS.map(h => r[h] !== undefined ? r[h] : ''));
  const sheet = {
    getLastRow: () => grid.length + 1, // +1 for the header row
    getLastColumn: () => AUDIT_LOG_HEADERS.length,
    getRange(row, col, numRows, numCols) {
      if (row === 1) return { getValues: () => [AUDIT_LOG_HEADERS.slice()] };
      // row 2 == grid[0]
      return { getValues: () => grid.slice(row - 2, row - 2 + numRows) };
    },
  };
  return {
    AUDIT_LOG_SHEET: 'Audit_Log',
    getOrCreateSpreadsheet: () => ({}),
    getOrCreateSheet: () => sheet,
  };
}

function run(rows, extraGlobals) {
  const logs = [];
  const sandbox = Object.assign(
    { Logger: { log(msg) { logs.push(String(msg)); } } },
    makeAuditLogSheet(rows),
    extraGlobals
  );
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(fnSource, ctx);
  return { getTeamDetailsFromAuditHistory: ctx.getTeamDetailsFromAuditHistory, logs };
}

let failures = 0;
function check(name, cond) {
  console.log('  ' + (cond ? 'PASS' : 'FAIL') + '  ' + name);
  if (!cond) failures++;
}

console.log('NH Analyzer — getTeamDetailsFromAuditHistory regression test');
console.log('file: ' + gsPath);

console.log('\nTest 1 — no history for this SAP ID returns null, does not throw');
{
  const { getTeamDetailsFromAuditHistory } = run([
    { 'SAP ID': '1111111', 'Team Leader': 'Someone Else' },
  ]);
  const result = getTeamDetailsFromAuditHistory('2222222');
  check('returns null', result === null);
}

console.log('\nTest 2 — empty Audit_Log (header only) returns null, does not throw');
{
  const { getTeamDetailsFromAuditHistory } = run([]);
  const result = getTeamDetailsFromAuditHistory('2222222');
  check('returns null', result === null);
}

console.log('\nTest 3 — one prior entry is returned with all fields');
{
  const { getTeamDetailsFromAuditHistory } = run([
    {
      'SAP ID': '2007888', 'Team Leader': 'TL-A', 'Operations Manager': 'OM-A',
      'Line of Business': 'MOB CLS', 'Locale': 'PH', 'VTID': 'V-100',
    },
  ]);
  const result = getTeamDetailsFromAuditHistory('2007888');
  check('found the entry', !!result);
  check('Team Leader',         result && result.teamLeader === 'TL-A');
  check('Operations Manager',  result && result.opsManager === 'OM-A');
  check('Line of Business',    result && result.lineOfBusiness === 'MOB CLS');
  check('Locale',               result && result.locale === 'PH');
  check('VTID',                 result && result.vtid === 'V-100');
}

console.log('\nTest 4 — multiple prior entries: the MOST RECENT (last row) wins, not the oldest');
{
  const { getTeamDetailsFromAuditHistory } = run([
    { 'SAP ID': '2007888', 'Team Leader': 'TL-OLD', 'Line of Business': 'MOB CLS' },   // oldest
    { 'SAP ID': '9999999', 'Team Leader': 'Unrelated Agent' },                          // different SAP ID in between
    { 'SAP ID': '2007888', 'Team Leader': 'TL-NEW', 'Line of Business': 'FFH CxSS' },   // most recent
  ]);
  const result = getTeamDetailsFromAuditHistory('2007888');
  check('used the most recent entry, not the oldest', result && result.teamLeader === 'TL-NEW');
  check('used the most recent Line of Business too',  result && result.lineOfBusiness === 'FFH CxSS');
}

console.log('\nTest 5 — SAP ID matching trims whitespace and compares as text');
{
  const { getTeamDetailsFromAuditHistory } = run([
    { 'SAP ID': ' 2007888 ', 'Team Leader': 'TL-A' },
  ]);
  const result = getTeamDetailsFromAuditHistory('2007888');
  check('matched despite stray whitespace in the sheet', result && result.teamLeader === 'TL-A');
}

console.log('\nTest 6 — a broken sheet read is swallowed, returns null instead of throwing');
{
  const sandbox = {
    Logger: { log() {} },
    AUDIT_LOG_SHEET: 'Audit_Log',
    getOrCreateSpreadsheet: () => ({}),
    getOrCreateSheet: () => { throw new Error('boom'); },
  };
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(fnSource, ctx);
  let threw = false;
  let result;
  try { result = ctx.getTeamDetailsFromAuditHistory('2007888'); } catch (e) { threw = true; }
  check('did not throw', !threw);
  check('returned null', result === null);
}

console.log(failures === 0 ? '\nALL TESTS PASSED' : '\n' + failures + ' CHECK(S) FAILED');
process.exit(failures === 0 ? 0 : 1);
