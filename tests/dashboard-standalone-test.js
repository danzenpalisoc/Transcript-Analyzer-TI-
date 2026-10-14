/**
 * The standalone Dashboard.html page (code review batch 3): its date filter must agree with index.html, and its CSV must not carry formulas.
 *
 * Usage: node tests/dashboard-standalone-test.js
 */
process.env.TZ = 'Asia/Manila';          // east of UTC: the case where "midnight UTC" is 8 AM local
const path = require('path');
const { load } = require('./lib/page-harness');
const { makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

const h = load(path.join(__dirname, '..', 'Dashboard.html'));
const ix = load();
const tzOk = new Date('2026-10-13').getHours() === 8;
check('the test runs in a time zone east of UTC (Asia/Manila), where the old rule failed', tzOk, String(new Date('2026-10-13').getHours()));

console.log('\nTest 1 — the date filter uses local midnight, like index.html');
{
  const rows = [
    { 'Audit Ref': 'A', 'Timestamp': '2026-10-13T00:30:00', 'Analysis Type': 'Repeats & Transfer Analyzer', 'Team Member': 'x' },   // first hour of the day, local
    { 'Audit Ref': 'B', 'Timestamp': '2026-10-12T23:30:00', 'Analysis Type': 'Repeats & Transfer Analyzer', 'Team Member': 'x' },   // the day before
    { 'Audit Ref': 'C', 'Timestamp': '2026-10-13T23:50:00', 'Analysis Type': 'Repeats & Transfer Analyzer', 'Team Member': 'x' },   // last minutes of the day
    { 'Audit Ref': 'D', 'Timestamp': '', 'Analysis Type': 'Repeats & Transfer Analyzer', 'Team Member': 'x' },                       // no usable date
    { 'Audit Ref': 'E', 'Timestamp': '2026-10-14T00:10:00', 'Analysis Type': 'Repeats & Transfer Analyzer', 'Team Member': 'x' },   // the day after
  ];
  h.ctx._allData = rows;
  h.ctx.getMultiSelect = () => [];
  h.ctx.renderStats = () => {}; h.ctx.renderCharts = () => {}; h.ctx.renderTable = () => {};
  const run = (from, to) => { h.els['r-date-from'] = Object.assign(h.document.getElementById('r-date-from'), { value: from }); h.els['r-date-to'] = Object.assign(h.document.getElementById('r-date-to'), { value: to }); h.ctx.applyFilters('r'); return h.ctx._rFiltered.map(r => r['Audit Ref']).join(''); };
  check('one day (13 Oct to 13 Oct) gives A and C: the first hour of the day is NOT cut off', run('2026-10-13', '2026-10-13') === 'AC', run('2026-10-13', '2026-10-13'));
  check('from 13 Oct only: A, C, E (D has no date so it is left out)', run('2026-10-13', '') === 'ACE', run('2026-10-13', ''));
  check('to 12 Oct only: B', run('', '2026-10-12') === 'B', run('', '2026-10-12'));
  check('with no date filter every row is shown, the one with no date too', run('', '') === 'ABCDE', run('', ''));
}

console.log('\nTest 2 — the CSV carries no formulas');
{
  const c = h.ctx.csvCell;
  check('a formula is turned into text with an apostrophe', c('=HYPERLINK("http://evil.example","x")') === '"\'=HYPERLINK(""http://evil.example"",""x"")"', c('=HYPERLINK("http://evil.example","x")'));
  check('+ - @ and a leading tab are neutralised too', c('+1+1') === "'+1+1" && c('-2+3') === "'-2+3" && c('@SUM(A1)') === "'@SUM(A1)" && c('\tcmd').charAt(0) === "'");
  check('ordinary text, numbers and a dash in the middle are unchanged', c('Maria Santos') === 'Maria Santos' && c('40%') === '40%' && c('2026-10-13') === '2026-10-13' && c('a-b') === 'a-b');
  check('commas, quotes and line breaks are still quoted', c('a,b') === '"a,b"' && c('say "hi"') === '"say ""hi"""' && c('l1\nl2') === '"l1\nl2"');
  check('an empty cell stays empty', c('') === '');
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
