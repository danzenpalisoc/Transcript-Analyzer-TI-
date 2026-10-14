/**
 * Dashboard features that were silently broken (code review batch 2).
 *
 *  1. Switching a tab threw ReferenceError: _dbCurrentFilters was never declared, so the chart redraw, the analytics button and
 *     _lastAnalyticsType (all after that line) never ran. The tab itself showed first, so nobody noticed.
 *  2. The Flag Records button called dbDownloadCSV, which exists nowhere, so the button did nothing but throw.
 *
 * Usage: node tests/dashboard-broken-features-test.js
 */
const { load } = require('./lib/page-harness');
const { makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const h = load();
const alerts = [];
h.ctx.alert = m => alerts.push(String(m));

const btn = { classList: { add() {}, remove() {}, toggle() {} }, attrs: {}, setAttribute(k, v) { this.attrs[k] = v; } };
const rows = [
  { 'Audit Ref': 'NHA-1', 'Analysis Type': 'Repeats & Transfer Analyzer', 'Critical Flags': 'Wrong LOB | No probing', 'Team Member': 'A' },
  { 'Audit Ref': 'NHA-2', 'Analysis Type': 'Repeats & Transfer Analyzer', 'Critical Flags': '', 'Team Member': 'B' },
  { 'Audit Ref': 'NHA-3', 'Analysis Type': 'Repeats & Transfer Analyzer', 'Critical Flags': '  ', 'Team Member': 'C' },
  { 'Audit Ref': 'NHA-1', 'Analysis Type': 'Repeats & Transfer Analyzer', 'Critical Flags': 'Wrong LOB', 'Team Member': 'A' },
  { 'Audit Ref': 'NHA-4', 'Analysis Type': 'Repeats & Transfer Analyzer', 'Critical Flags': 'Hold too long', 'Team Member': 'D' },
];

console.log('\nTest 1 — switching tabs runs to the end');
{
  h.ctx._dbLoaded = true;
  h.ctx._dbAllData = rows; h.ctx._dbRFiltered = rows; h.ctx._dbSFiltered = [];
  h.ctx._dbAuditLog = [{ 'Audit Ref': 'x' }];
  const results = {};
  ['overview', 'repeats', 'sales', 'auditlog'].forEach(tab => {
    let err = '';
    try { h.ctx.dbSwitchTab(tab, btn); } catch (e) { err = e.name + ': ' + e.message; }
    results[tab] = err;
    check('dbSwitchTab(\'' + tab + '\') does not throw', err === '', err);
  });
  h.ctx.dbSwitchTab('sales', btn);
  check('the last tab is remembered for the analytics button (sales -> s)', h.ctx._lastAnalyticsType === 's', String(h.ctx._lastAnalyticsType));
  h.ctx.dbSwitchTab('repeats', btn);
  check('and repeats -> r', h.ctx._lastAnalyticsType === 'r', String(h.ctx._lastAnalyticsType));
  h.ctx.dbSwitchTab('overview', btn);
  check('the analytics button is hidden on Overview, which has no per-type analytics', h.els['analytics-fab'] && h.els['analytics-fab'].style.display === 'none', h.els['analytics-fab'] && h.els['analytics-fab'].style.display);
}

console.log('\nTest 2 — Flag Records exports only the flagged audits');
{
  h.reset(); alerts.length = 0;
  h.ctx._dbRFiltered = rows; h.ctx._dbSFiltered = [];
  let err = '';
  try { h.ctx.dbFlagOnly('r'); } catch (e) { err = e.name + ': ' + e.message; }
  check('the button does not throw (dbDownloadCSV does not exist)', err === '', err);
  const call = h.take('exportAudits');
  check('it asks the server for an export', !!call, JSON.stringify(h.pending.map(p => p.method)));
  if (call) {
    check('view is repeats', call.args[0] === 'repeats', JSON.stringify(call.args[0]));
    check('only the audits that have a flag are sent, each once, in order (NHA-1, NHA-4)', JSON.stringify(call.args[1]) === JSON.stringify(['NHA-1', 'NHA-4']), JSON.stringify(call.args[1]));
    check('the title says it is the flagged records', /Flagged/.test(call.args[2]), String(call.args[2]));
  }

  h.reset(); alerts.length = 0;
  h.ctx._dbRFiltered = [rows[1], rows[2]]; h.ctx._dbSFiltered = [{ 'Audit Ref': 'NHA-9', 'Analysis Type': 'Sales Analyzer', 'Critical Flags': 'No offer' }];
  h.ctx.dbFlagOnly('r');
  check('with no flagged audit in the view the person is told and nothing is exported', alerts.some(m => /No flagged records/.test(m)) && !h.take('exportAudits'), JSON.stringify(alerts));

  h.reset(); alerts.length = 0;
  h.ctx.dbFlagOnly('s');
  const sc = h.take('exportAudits');
  check('the Sales tab exports from the Sales view', sc && sc.args[0] === 'sales' && JSON.stringify(sc.args[1]) === JSON.stringify(['NHA-9']), sc && JSON.stringify(sc.args));
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
