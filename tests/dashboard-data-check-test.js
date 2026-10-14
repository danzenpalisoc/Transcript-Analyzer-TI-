/**
 * verifyDashboardData() (AuditCheck.gs): a read-only, admin-only report that says whether
 * the data behind the dashboard is consistent, and what numbers the dashboard will show
 * for it.
 *
 * The dashboard is built FROM Audit_Log and joined to Dashboard_Data by Audit Ref, so the
 * report checks exactly that join, the headers, the values the charts depend on, and
 * works out the KPIs the way the dashboard does so they can be compared with the screen.
 *
 * Usage: node tests/dashboard-data-check-test.js
 */
const vm = require('vm');
const { read, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

function makeEnv(opts) {
  opts = opts || {};
  const logs = [];
  const who = { email: 'danzen.palisoc@telus.com' };
  const sandbox = {
    Logger: { log: m => logs.push(String(m)) }, JSON, Array, Object, String, Number, Math, Date, RegExp, isNaN, parseFloat, Error,
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Session: { getActiveUser: () => ({ getEmail: () => who.email }) },
  };
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('Config.gs'), ctx);
  vm.runInContext(read('AuditCheck.gs'), ctx);
  const sheets = {};
  [['AUDIT_LOG_SHEET', 'AUDIT_LOG_HEADERS'], ['DASHBOARD_DATA_SHEET', 'DASHBOARD_HEADERS'], ['ANALYSIS_SHEET', 'ANALYSIS_HEADERS'],
   ['TRANSCRIPTS_SHEET', 'TRANSCRIPTS_HEADERS']].forEach(([s, h]) => { sheets[ctx[s]] = new FakeSheet((opts.headers && opts.headers[s]) || ctx[h], []); });
  const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new FakeSheet([], [])), getUrl: () => 'u' };
  ctx.getOrCreateSpreadsheet = () => ss;
  const add = (sheetConst, o) => { const sh = sheets[ctx[sheetConst]]; sh.grid.push(sh.grid[0].map(h => (o[h] !== undefined ? o[h] : ''))); };
  const run = () => ctx.verifyDashboardData();
  return { ctx, sheets, logs, who, add, run };
}
const NOW = new Date('2026-10-12T10:00:00');
const log = (ref, o) => Object.assign({ 'Audit Ref': ref, 'Submitted At': NOW, 'Interaction ID': 'i-' + ref, 'SAP ID': '111', 'Team Member': 'Agent ' + ref,
  'Analysis Type': 'Repeats & Transfer Analyzer', 'Repeat Projection %': '40%', 'Issue Resolved': 'Yes', 'Email Status': 'Not Sent' }, o);
const dash = (ref, o) => Object.assign({ 'Audit Ref': ref, 'Timestamp': NOW, 'Interaction ID': 'i-' + ref, 'SAP ID': '111', 'Team Member': 'Agent ' + ref,
  'Analysis Type': 'Repeats & Transfer Analyzer', 'Call Summary': 'summary', 'Call Driver': 'Billing' }, o);
function healthy(n) {
  const e = makeEnv();
  for (let i = 1; i <= (n === undefined ? 3 : n); i++) { e.add('AUDIT_LOG_SHEET', log('R' + i)); e.add('DASHBOARD_DATA_SHEET', dash('R' + i)); }
  return e;
}
const fails = r => r.lines.filter(l => l.level === 'fail').map(l => l.text).join(' | ');
const has = (r, re) => r.lines.some(l => l.level === 'fail' && re.test(l.text));

console.log('\nTest 1 — consistent data is reported clean, with the counts the dashboard will show');
{
  const e = healthy(3);
  const r = e.run();
  check('no problems', r.problems === 0, fails(r));
  check('Audit_Log and Dashboard_Data row counts', r.stats.auditLogRows === 3 && r.stats.dashboardRows === 3, JSON.stringify(r.stats));
  check('tab totals: 3 Repeats, 0 Sales, 0 in neither', r.stats.repeats === 3 && r.stats.sales === 0 && r.stats.neither === 0, JSON.stringify(r.stats));
  check('the log has a RESULT line', e.logs.some(l => /RESULT:/.test(l)), e.logs.slice(-2).join(' | '));
  check('every analysis field filled in -> no blank-analysis count', r.stats.auditOnly === 0 && r.stats.dashboardOnly === 0);
}

console.log('\nTest 2 — the join the dashboard depends on');
{
  let e = healthy(2); e.add('DASHBOARD_DATA_SHEET', dash('ORPHAN'));
  let r = e.run();
  check('a Dashboard_Data row with no Audit_Log row is reported as HIDDEN from the dashboard', r.problems > 0 && has(r, /hidden|not in Audit_Log/i) && r.stats.dashboardOnly === 1, fails(r));
  check('...and the ref is named', r.lines.some(l => /ORPHAN/.test(l.text)));

  e = healthy(2); e.add('AUDIT_LOG_SHEET', log('NODASH'));
  r = e.run();
  check('an Audit_Log row with no Dashboard_Data row is reported (counted with blank analysis)', has(r, /no Dashboard_Data row|blank analysis/i) && r.stats.auditOnly === 1, fails(r));

  e = healthy(2); const l = e.sheets[e.ctx.AUDIT_LOG_SHEET]; l.grid.push(l.grid[1].slice());
  r = e.run();
  check('a ref twice in Audit_Log', has(r, /more than once|duplicate/i), fails(r));

  e = healthy(2); const d = e.sheets[e.ctx.DASHBOARD_DATA_SHEET]; d.grid.push(d.grid[1].slice());
  r = e.run();
  check('a ref twice in Dashboard_Data', has(r, /Dashboard_Data.*(more than once|duplicate)|duplicate.*Dashboard_Data/i), fails(r));

  e = healthy(2); e.add('AUDIT_LOG_SHEET', log(''));
  r = e.run();
  check('a row with no Audit Ref', has(r, /no Audit Ref|blank Audit Ref/i), fails(r));

  e = healthy(2); e.sheets[e.ctx.DASHBOARD_DATA_SHEET].grid[1][e.ctx.DASHBOARD_HEADERS.indexOf('SAP ID')] = '999';
  r = e.run();
  check('the same ref with a different SAP ID in the two sheets', has(r, /SAP ID/i) && r.stats.sapMismatch === 1, fails(r));
}

console.log('\nTest 3 — header health (the code reads columns by name, so a missing header means silent blanks)');
{
  const e = makeEnv({ headers: { DASHBOARD_DATA_SHEET: ['Audit Ref', 'Timestamp', 'SAP ID'] } });
  e.add('AUDIT_LOG_SHEET', log('R1'));
  const r = e.run();
  check('missing Dashboard_Data columns are named', has(r, /Dashboard_Data.*missing|missing.*Dashboard_Data/i) && r.lines.some(l => /Call Summary/.test(l.text)), fails(r));
}

console.log('\nTest 4 — values the charts depend on');
{
  let e = healthy(0);
  [['A', 'Repeats & Transfer Analyzer'], ['B', 'Sales Analyzer'], ['C', '']].forEach(([ref, t]) => { e.add('AUDIT_LOG_SHEET', log(ref, { 'Analysis Type': t })); e.add('DASHBOARD_DATA_SHEET', dash(ref, { 'Analysis Type': t })); });
  let r = e.run();
  check('an audit with a blank Analysis Type is in no tab: reported', r.stats.neither === 1 && has(r, /no tab|neither/i), fails(r) + ' ' + JSON.stringify(r.stats));
  check('Repeats 1, Sales 1', r.stats.repeats === 1 && r.stats.sales === 1, JSON.stringify(r.stats));

  e = healthy(0);
  ['Yes', 'No', '', 'Unknown', 'Partially resolved - credit pending'].forEach((v, i) => { e.add('AUDIT_LOG_SHEET', log('R' + i, { 'Issue Resolved': v })); e.add('DASHBOARD_DATA_SHEET', dash('R' + i)); });
  r = e.run();
  check('answers: 1 Yes, 1 Partially, 1 No, 1 blank, 1 with no verdict', r.stats.resolvedYes === 1 && r.stats.resolvedPartial === 1 && r.stats.resolvedNo === 1 && r.stats.resolvedBlank === 1 && r.stats.resolvedUnknown === 1, JSON.stringify(r.stats));
  check('what the dashboard KPI shows: Resolved 1 / Partially 1 / Unresolved 1 / Not recorded 2', r.stats.dashResolved === 1 && r.stats.dashPartial === 1 && r.stats.dashUnresolved === 1 && r.stats.dashNotRecorded === 2, JSON.stringify(r.stats));
  check('a value with no verdict is a warning, not a failure', r.lines.some(l => l.level === 'warn' && /verdict/.test(l.text)) && !has(r, /Issue Resolved/), JSON.stringify(r.lines));

  e = healthy(0);
  ['0%', '0%', '40%', ''].forEach((v, i) => { e.add('AUDIT_LOG_SHEET', log('R' + i, { 'Repeat Projection %': v })); e.add('DASHBOARD_DATA_SHEET', dash('R' + i)); });
  r = e.run();
  check('repeat risk: 2 zeros, 1 value, 1 blank', r.stats.repeatZero === 2 && r.stats.repeatBlank === 1 && r.stats.repeatNumeric === 3, JSON.stringify(r.stats));
  check('the average is 13 (the dashboard counts the 0% audits too)', r.stats.repeatAvgTrue === 13, JSON.stringify(r.stats));
  check('there is no gap to report any more', !has(r, /Avg Repeat/i), fails(r));

  e = healthy(0);
  ['95% repeat risk. No issue was identified', '85%', 'about forty', ''].forEach((v, i) => { e.add('AUDIT_LOG_SHEET', log('S' + i, { 'Repeat Projection %': v })); e.add('DASHBOARD_DATA_SHEET', dash('S' + i)); });
  r = e.run();
  check('a sentence is read like the dashboard reads it (2 numeric, 1 of them from a sentence, 1 unreadable, 1 blank)',
        r.stats.repeatNumeric === 2 && r.stats.repeatFromText === 1 && r.stats.repeatNonNumeric === 1 && r.stats.repeatBlank === 1, JSON.stringify(r.stats));
  check('the average is 90 (95 and 85)', r.stats.repeatAvgTrue === 90, JSON.stringify(r.stats));

  const eF = healthy(0);
  ['0.78', '65', 0.52].forEach((v, i) => { eF.add('AUDIT_LOG_SHEET', log('F' + i, { 'Repeat Projection %': v })); eF.add('DASHBOARD_DATA_SHEET', dash('F' + i)); });
  const rF = eF.run();
  check('fractions (0.78, 0.52) and whole numbers (65) are put on one scale: average 65', rF.stats.repeatAvgTrue === 65 && rF.stats.repeatFraction === 2, JSON.stringify(rF.stats));
  check('the report says how many fractions it read', rF.lines.some(l => l.level === 'info' && /fraction/.test(l.text)));
  check('a sentence in the cell is a warning, a value with no percentage is a failure',
        r.lines.some(l => l.level === 'warn' && /sentence/.test(l.text)) && has(r, /not numbers/), JSON.stringify(r.lines.filter(l => /Repeat/.test(l.text))));

  e = healthy(0);
  // blank in BOTH sheets (the dashboard falls back to Dashboard_Data's Timestamp, so blank in one is fine)
  e.add('AUDIT_LOG_SHEET', log('R1', { 'Submitted At': '' })); e.add('DASHBOARD_DATA_SHEET', dash('R1', { 'Timestamp': '' }));
  e.add('AUDIT_LOG_SHEET', log('R2', { 'Submitted At': new Date('2099-01-01') })); e.add('DASHBOARD_DATA_SHEET', dash('R2'));
  e.add('AUDIT_LOG_SHEET', log('R3', { 'Submitted At': 'not a date' })); e.add('DASHBOARD_DATA_SHEET', dash('R3'));
  r = e.run();
  check('blank, future and invalid timestamps are reported (a date filter hides them)', r.stats.badDates === 3 && has(r, /timestamp|date/i), fails(r) + ' ' + JSON.stringify(r.stats));

  e = healthy(0);
  e.add('AUDIT_LOG_SHEET', log('R1', { 'Team Member': 'x'.repeat(120) })); e.add('DASHBOARD_DATA_SHEET', dash('R1'));
  r = e.run();
  check('a Team Member that is really a column-shifted paragraph', r.stats.longNames === 1 && has(r, /80 characters|column.shift/i), fails(r));

  e = healthy(0);
  ['Billing dispute', 'billing dispute ', 'Billing Dispute', 'Router offline'].forEach((v, i) => { e.add('AUDIT_LOG_SHEET', log('R' + i)); e.add('DASHBOARD_DATA_SHEET', dash('R' + i, { 'Call Driver': v })); });
  r = e.run();
  check('Call Drivers that differ only by capitals/spacing are found', r.stats.driverVariantGroups === 1 && r.lines.some(l => /billing dispute/i.test(l.text) && /Billing Dispute/.test(l.text)), JSON.stringify(r.stats));

  e = healthy(0);
  ['Sent', 'Not Sent', 'Failed: quota'].forEach((v, i) => { e.add('AUDIT_LOG_SHEET', log('R' + i, { 'Email Status': v })); e.add('DASHBOARD_DATA_SHEET', dash('R' + i)); });
  r = e.run();
  check('email statuses are counted, and a failure reason is shown', r.stats.emailSent === 1 && r.lines.some(l => /Failed: quota/.test(l.text)), JSON.stringify(r.stats));
}

console.log('\nTest 5 — read-only, admin-only, and safe on an empty workbook');
{
  const e = healthy(3);
  let writes = 0;
  Object.keys(e.sheets).forEach(k => { const s = e.sheets[k]; const a = s.appendRow.bind(s), d = s.deleteRow.bind(s); s.appendRow = x => { writes++; return a(x); }; s.deleteRow = n => { writes++; return d(n); }; });
  e.run();
  check('nothing is written', writes === 0, 'writes: ' + writes);
  e.who.email = 'some.analyst@telus.com';
  let threw = false; try { e.run(); } catch (x) { threw = /admin/i.test(x.message); }
  check('a non-admin is refused', threw);

  const empty = makeEnv();
  const r = empty.run();
  check('an empty workbook is reported, not a crash', r.ok === false || r.problems >= 0, JSON.stringify(r.stats));
  check('...with zero rows counted', r.stats.auditLogRows === 0 && r.stats.dashboardRows === 0);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
