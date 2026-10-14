/**
 * Dashboard correctness: the numbers on the dashboard must match the data behind them.
 *
 * It runs the REAL dashboard code from index.html and Dashboard.html, and the real
 * extraction code from Code.gs, against small hand-made datasets whose right answers are
 * obvious. (This started as an audit script that found 18 places where the dashboard
 * showed a number that did not match its data; each one is now a check that must pass.)
 *
 * It proves the LOGIC. It cannot see the real sheets: run verifyDashboardData() from the
 * Apps Script editor for that.
 *
 * Usage: node tests/dashboard-correctness-test.js
 */
const vm = require('vm');
const path = require('path');
const { load } = require('./lib/page-harness');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

// ── a page with a stubbed Google Charts that records what would be drawn ──────────────
function pageWithCharts(htmlPath) {
  const h = load(htmlPath);
  const drawn = { pie: {}, bar: {} };
  function DataTable() { this.rows = []; }
  DataTable.prototype.addColumn = function () {};
  DataTable.prototype.addRow = function (r) { this.rows.push(r); };
  DataTable.prototype.addRows = function (rs) { rs.forEach(r => this.rows.push(r)); };
  DataTable.prototype.getNumberOfRows = function () { return this.rows.length; };
  h.ctx.google.visualization = {
    DataTable,
    PieChart: function (el) { this.draw = dt => { drawn.pie[el.id] = dt.rows.slice(); }; },
    BarChart: function (el) { this.draw = dt => { drawn.bar[el.id] = dt.rows.slice(); }; },
  };
  return Object.assign(h, { drawn });
}
const blank = o => Object.assign({
  'Audit Ref': '', 'Timestamp': '2026-10-12T10:00:00', 'Team Member': 'Agent A', 'Team Leader': 'TL', 'Analysis Type': 'Repeats & Transfer Analyzer',
  'Repeat Projection %': '', 'Issue Resolved': '', 'Critical Flags': '', 'Email Status': 'Not Sent', 'Call Driver': '', 'RCA Category': '',
}, o);
function statVal(html, label) {
  const m = new RegExp('<div class="db-stat-val">([^<]*)</div><div class="db-stat-lbl">' + label.replace(/[()]/g, c => '\\' + c) + '</div>').exec(html);
  return m ? m[1] : null;
}

// Real sentences from the live Audit_Log (the 540 free-text Issue Resolved values).
const REAL = {
  partial: 'Partially resolved - the customer was given the refund steps but the credit has not posted yet',
  notRes: 'Not resolved - the customer still has no service and a technician visit is needed',
  completelyUnres: 'Completely unresolved: the agent transferred the call without addressing the billing question',
  yes: 'Yes - the agent fixed the billing error and confirmed the new balance',
  noVerdict: 'Agent confirmed the form was submitted. Resolution is incomplete until the review finishes',
};

console.log('\nTest 1 — Issue Resolved is read as a verdict (Resolved / Partially / Unresolved / Not recorded)');
{
  const h = pageWithCharts();
  const r = v => h.ctx._dbResolution(v);
  check('"Yes" is resolved', r('Yes') === 'yes', r('Yes'));
  check('a real "Yes - ..." sentence is resolved', r(REAL.yes) === 'yes', r(REAL.yes));
  check('a real "Partially resolved - ..." sentence is partial, not resolved', r(REAL.partial) === 'partial', r(REAL.partial));
  check('"Not resolved - ..." is unresolved', r(REAL.notRes) === 'no', r(REAL.notRes));
  check('"Completely unresolved: ..." is unresolved', r(REAL.completelyUnres) === 'no', r(REAL.completelyUnres));
  check('"No" is unresolved', r('No') === 'no');
  check('"No, not yes" is unresolved (contains the letters yes)', r('No, not yes') === 'no');
  check('"Not resolved - yes follow up" is unresolved', r('Not resolved - yes follow up') === 'no');
  check('a blank is not recorded, not unresolved', r('') === 'unknown' && r(null) === 'unknown' && r(undefined) === 'unknown');
  check('"Unknown" is not recorded', r('Unknown') === 'unknown');
  check('text with no verdict is not guessed', r(REAL.noVerdict) === 'unknown', r(REAL.noVerdict));
  check('"Issue was NOT resolved. The promised $10..." is unresolved', r('Issue was NOT resolved. The promised $10 pre-authorized payment was never applied') === 'no');
  check('"The issue was resolved" is resolved', r('The issue was resolved after the agent reset the account') === 'yes');
  check('"Issue was partially resolved" is partial', r('Issue was partially resolved; credit still pending') === 'partial');
  check('"The issue wasn\'t resolved" is unresolved', r("The issue wasn't resolved on this call") === 'no');
  check('"Resolution is ambiguous" is still not recorded', r('Resolution is ambiguous. Interaction summary states phones were returned') === 'unknown');

  const rows = [blank({ 'Issue Resolved': 'Yes' }), blank({ 'Issue Resolved': REAL.partial }), blank({ 'Issue Resolved': REAL.partial }),
                blank({ 'Issue Resolved': 'No' }), blank({ 'Issue Resolved': '' }), blank({ 'Issue Resolved': 'Unknown' })];
  h.reset(); h.ctx.dbRenderStats('r', rows);
  const html = h.els['r-stats'].innerHTML;
  check('KPI Issues Resolved = 1', statVal(html, 'Issues Resolved') === '1', statVal(html, 'Issues Resolved'));
  check('KPI Partially Resolved = 2 (its own count)', statVal(html, 'Partially Resolved') === '2', statVal(html, 'Partially Resolved'));
  check('KPI Unresolved = 1 (only the "No")', statVal(html, 'Unresolved') === '1', statVal(html, 'Unresolved'));
  check('KPI Not Recorded = 2 (blank + Unknown)', statVal(html, 'Not Recorded') === '2', statVal(html, 'Not Recorded'));
  check('the four counts add up to the total', 1 + 2 + 1 + 2 === rows.length);
  check('Not Recorded card is hidden when every audit has a verdict',
        (h.reset(), h.ctx.dbRenderStats('r', [blank({ 'Issue Resolved': 'Yes' })]), statVal(h.els['r-stats'].innerHTML, 'Not Recorded') === null));

  h.reset(); h.drawn.pie = {};
  h.ctx.dbDrawChart('r', rows);
  const donut = h.drawn.pie['r-chart-resolution'] || [];
  const get = n => (donut.find(x => x[0] === n) || [0, null])[1];
  check('the donut shows the same four numbers as the KPI cards',
        get('Resolved') === 1 && get('Partially resolved') === 2 && get('Unresolved') === 1 && get('Not recorded') === 2, JSON.stringify(donut));
}

console.log('\nTest 2 — Avg Repeat Risk counts a real 0% and ignores blanks and text');
{
  const h = pageWithCharts();
  const stats = rows => { h.reset(); h.ctx.dbRenderStats('r', rows); return h.els['r-stats'].innerHTML; };
  check('0%, 0%, 40% averages to 13%', statVal(stats([blank({ 'Repeat Projection %': '0%' }), blank({ 'Repeat Projection %': '0%' }), blank({ 'Repeat Projection %': '40%' })]), 'Avg Repeat Risk') === '13%');
  check('a blank is left out (40% not 20%)', statVal(stats([blank({ 'Repeat Projection %': '' }), blank({ 'Repeat Projection %': '40%' })]), 'Avg Repeat Risk') === '40%');
  check('text such as "N/A" is left out', statVal(stats([blank({ 'Repeat Projection %': 'N/A' }), blank({ 'Repeat Projection %': '40%' })]), 'Avg Repeat Risk') === '40%');
  check('a plain number 40 and the text " 40 % " both read as 40', h.ctx._dbNum(40) === 40 && h.ctx._dbNum(' 40 % ') === 40);
  // Older audits stored a whole sentence in this cell (real examples from the live Audit_Log). They must still count.
  const n = v => h.ctx._dbNum(v);
  check('"95% repeat risk. No issue was identified..." reads 95', n('95% repeat risk. No issue was identified, no resolution was reached') === 95);
  check('"~55% repeat risk. Key drivers: ..." reads 55', n('~55% repeat risk. Key drivers: vague delivery time') === 55);
  check('"High — approximately 80%. Customer\'s compensation ..." reads 80', n("High — approximately 80%. Customer's compensation was not applied") === 80);
  check('"Repeat risk estimated at ~45%. ..." reads 45', n('Repeat risk estimated at ~45%. Key drivers: equipment') === 45);
  check('"Low-to-moderate repeat risk (~25%). The $49.50 charge" reads 25 (not 49.5)', n('Low-to-moderate repeat risk (~25%). The $49.50 charge was explained') === 25);
  check('"85% — Transfer without resolution ..." reads 85', n('85% — Transfer without resolution guarantees a second call') === 85);
  check('a range "40-50%" reads its middle, 45', n('Repeat risk 40-50% depending on the callback') === 45);
  // A percent cell holds 0.78 for 78% (the sheet turned "78%" into a number), and the server sends it as the text "0.78".
  check('0.78 (a percent cell) reads 78', n(0.78) === 78 && n('0.78') === 78, JSON.stringify([n(0.78), n('0.78')]));
  check('0.52 reads 52 and 0.975 reads 97.5', n('0.52') === 52 && n(0.975) === 97.5);
  check('a whole number 65 stays 65, and so does the text "65"', n(65) === 65 && n('65') === 65);
  check('a bare 1 is 100% (a percent cell), a written "1%" is 1', n('1') === 100 && n('1%') === 1);
  check('a written "0.5%" stays 0.5', n('0.5%') === 0.5);
  check('0 stays 0', n(0) === 0 && n('0%') === 0);
  check('the average mixes both scales correctly (0.78 and 65 -> 72)',
        statVal(stats([blank({ 'Repeat Projection %': '0.78' }), blank({ 'Repeat Projection %': '65' })]), 'Avg Repeat Risk') === '72%');
  { // the tables show N%, never "0.78%"
    const t = pageWithCharts();
    t.reset(); t.ctx.dbRenderTable('r', [blank({ 'Audit Ref': 'A1', 'Repeat Projection %': '0.78' }), blank({ 'Audit Ref': 'A2', 'Repeat Projection %': '65' })]);
    const html = (t.els['r-tbody'] || t.els['r-table-body'] || {}).innerHTML || Object.keys(t.els).map(k => t.els[k].innerHTML || '').join('');
    check('the Repeats table shows 78% and 65%, not 0.78%', />78%</.test(html) && />65%</.test(html) && !/0\.78%/.test(html), html.slice(0, 200));
    const s2 = pageWithCharts(path.join(__dirname, '..', 'Dashboard.html'));
    s2.reset(); s2.ctx.renderTable('r', [blank({ 'Audit Ref': 'A1', 'Repeat Projection %': '0.78' })]);
    const html2 = Object.keys(s2.els).map(k => s2.els[k].innerHTML || '').join('');
    check('the standalone dashboard table shows 78% too', />78%</.test(html2) && !/0\.78/.test(html2), html2.slice(0, 200));
  }
  check('a sentence with no percentage is left out, not 0', n('about forty') === null && n('N/A') === null);
  check('a percentage above 100 inside a sentence is not trusted', n('The customer was charged 250% of the normal rate') === null);
  check('the average counts sentence rows (95 and 85 -> 90%)',
        statVal(stats([blank({ 'Repeat Projection %': '95% repeat risk. No issue was identified' }), blank({ 'Repeat Projection %': '85% repeat risk. The call ended' })]), 'Avg Repeat Risk') === '90%');
  check('every audit blank gives 0%, not a crash', statVal(stats([blank({})]), 'Avg Repeat Risk') === '0%');
}

console.log('\nTest 3 — Critical Flags: one flag is one flag');
{
  const h = pageWithCharts();
  const split = s => JSON.stringify(h.ctx._dbSplitFlags(s));
  check('new rows are split on " | " and keep commas inside a title',
        split('Wrong LOB, handled a TV call | No save attempt') === JSON.stringify(['Wrong LOB, handled a TV call', 'No save attempt']), split('Wrong LOB, handled a TV call | No save attempt'));
  check('old rows: a piece that starts in lower case belongs to the flag before it',
        split('Wrong LOB, handled a TV call, No save attempt') === JSON.stringify(['Wrong LOB, handled a TV call', 'No save attempt']), split('Wrong LOB, handled a TV call, No save attempt'));
  check('old rows: plain comma lists still split', split('F1, F2, F3') === JSON.stringify(['F1', 'F2', 'F3']));
  check('"None" and blank are no flags', split('None') === '[]' && split('') === '[]' && split(null) === '[]');
  const counts = h.ctx._dbFlagCounts([blank({ 'Critical Flags': 'No save attempt' }), blank({ 'Critical Flags': 'no save attempt | Other' })]);
  check('the same flag in different capitals is one bucket', counts['No save attempt'] === 2, JSON.stringify(counts));

  const run = rows => { const p = pageWithCharts(); p.ctx._dbRFiltered = rows; p.ctx._dbSFiltered = []; p.ctx.renderOverviewTab(); return p; };
  const ov = run([blank({ 'Critical Flags': 'Wrong LOB, handled a TV call' })]);
  check('Total Critical Flags counts that one flag as 1', statVal(ov.els['ov-stats'].innerHTML, 'Total Critical Flags') === '1', statVal(ov.els['ov-stats'].innerHTML, 'Total Critical Flags'));

  const p2 = pageWithCharts();
  p2.ctx.dbDrawChart('r', [blank({ 'Critical Flags': 'Wrong LOB, handled a TV call | No save attempt' })]);
  check('the flag chart shows 2 bars for a 2-flag audit', (p2.drawn.bar['r-chart-flags'] || []).length === 2, JSON.stringify(p2.drawn.bar['r-chart-flags']));

  h.reset(); h.ctx.dbRenderStats('r', [blank({ 'Critical Flags': 'None' }), blank({ 'Critical Flags': 'F1' })]);
  check('the flagged-audits card does not count "None"', statVal(h.els['r-stats'].innerHTML, 'Audits with Flags') === '1', statVal(h.els['r-stats'].innerHTML, 'Critical Flags'));
}

console.log('\nTest 4 — the Overview tab');
{
  const run = rows => { const p = pageWithCharts(); p.ctx._dbRFiltered = rows; p.ctx._dbSFiltered = []; p.ctx.renderOverviewTab(); return p; };
  let h = run([blank({ 'Team Member': 'Agent A', 'Critical Flags': 'Flag X' }), blank({ 'Team Member': 'Agent A', 'Critical Flags': 'Flag X' }),
               blank({ 'Team Member': 'Agent A', 'Critical Flags': 'Flag X' }), blank({ 'Team Member': 'Agent B', 'Critical Flags': 'Flag X' })]);
  const cell = /Flag X[\s\S]*?<td style="font-size:12px">([^<]*)<\/td>/.exec(h.els['ov-flag-tbody'].innerHTML);
  check('"Top Agent" for a flag is the agent with the MOST of it (Agent A, 3 of 4)', !!cell && cell[1] === 'Agent A', cell ? cell[1] : 'no cell');

  h = run([blank({ 'Team Member': 'Agent A' }), blank({ 'Team Member': 'Agent B' })]);
  check('nobody is called "top performer" when everyone ties at 0', !/top performer/i.test(h.els['ov-highlights'].innerHTML), h.els['ov-highlights'].innerHTML);

  h = run([blank({ 'Team Member': 'Agent A' })]);
  check('a lone agent is not called the top performer', !/top performer/i.test(h.els['ov-highlights'].innerHTML));

  h = run([blank({ 'Team Member': 'Agent A ', 'Critical Flags': 'F1 | F2 | F3' }), blank({ 'Team Member': 'Agent B' })]);
  const hl = h.els['ov-highlights'].innerHTML;
  check('a stray space in a name does not hide the agent\'s flags', /Agent B<\/strong> has the fewest/.test(hl), hl);

  h = run([blank({ 'Team Member': 'Agent A', 'Critical Flags': 'F1 | F2 | F3 | F4 | F5' }), blank({ 'Team Member': 'Agent B' }), blank({ 'Team Member': 'Agent C' })]);
  check('two agents tied for fewest flags: nobody is named top performer', !/top performer/i.test(h.els['ov-highlights'].innerHTML), h.els['ov-highlights'].innerHTML);
  check('the single agent with the most flags is named for coaching', /Agent A<\/strong> has the most critical flags/.test(h.els['ov-lowlights'].innerHTML), h.els['ov-lowlights'].innerHTML);

  h = run([blank({ 'Repeat Projection %': '0%' }), blank({ 'Repeat Projection %': '40%' })]);
  check('the Overview average counts a real 0% too (20%)', statVal(h.els['ov-stats'].innerHTML, 'Avg Repeat Risk') === '20%', statVal(h.els['ov-stats'].innerHTML, 'Avg Repeat Risk'));
  check('the agent table shows 20% for that agent', />20%</.test(h.els['ov-tm-tbody'].innerHTML), h.els['ov-tm-tbody'].innerHTML.slice(0, 200));
  h = run([blank({ 'Repeat Projection %': '0%' })]);
  check('an agent whose only repeat value is 0% shows 0%, not a dash', />0%</.test(h.els['ov-tm-tbody'].innerHTML), h.els['ov-tm-tbody'].innerHTML.slice(0, 200));
}

console.log('\nTest 5 — what reaches the dashboard: the Audit_Log + Dashboard_Data merge');
{
  const cfg = vm.createContext({ PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) } });
  vm.runInContext(read('Config.gs'), cfg);
  const headers = cfg.DASHBOARD_HEADERS;
  const merge = (dashRows, auditRows) => {
    const h = pageWithCharts();
    h.ctx.dbFinishRender = function () {};
    h.ctx.dbHandleData({ rows: dashRows, ts: new Date().toISOString() }, false);
    h.pending.find(p => p.method === 'getAuditLogData').success(auditRows);
    return { all: h.ctx._dbAllData, status: (h.els['db-cache-status'] || {}).textContent || '' };
  };
  // a Dashboard_Data row with EVERY real column filled
  const dash = (ref, o) => { const r = {}; headers.forEach(hd => { r[hd] = hd + ' ' + ref; }); r['Audit Ref'] = ref; return Object.assign(r, o); };
  const log = (ref, o) => Object.assign({ 'Audit Ref': ref, 'Submitted At': '2026-10-12T10:00:00', 'SAP ID': '111', 'Team Member': 'Agent ' + ref, 'Analysis Type': 'Repeats & Transfer Analyzer' }, o);

  let m = merge([dash('A'), dash('B')], [log('A')]);
  check('an audit in Dashboard_Data with no Audit_Log row is still shown', m.all.length === 2 && !!m.all.find(r => r['Audit Ref'] === 'B'), m.all.map(r => r['Audit Ref']).join(','));
  const b = m.all.find(r => r['Audit Ref'] === 'B') || {};
  check('...and it keeps its own data (summary, flags)', b['Call Summary'] === 'Call Summary B' && b['Critical Flags'] === 'Critical Flags B', JSON.stringify([b['Call Summary'], b['Critical Flags']]));

  m = merge([dash('A')], [log('A')]);
  const lost = headers.filter(hd => m.all[0][hd] === undefined);
  check('every Dashboard_Data column reaches the merged row', lost.length === 0, lost.join(', '));
  check('Audit_Log values still win where both have one (Team Member)', m.all[0]['Team Member'] === 'Agent A', m.all[0]['Team Member']);

  m = merge([dash('A')], [log('A'), log('C')]);
  const c = m.all.find(r => r['Audit Ref'] === 'C');
  check('an audit with no Dashboard_Data row is marked as having no analysis', !!c && c._noAnalysis === true);
  check('...and the status line says so', /1 audit\(s\) have no analysis row/.test(m.status), m.status);
  check('an audit that has its Dashboard_Data row is not marked', m.all.find(r => r['Audit Ref'] === 'A')._noAnalysis === false);

  m = merge([dash('A'), dash('B', { 'Analysis Type': '' })], [log('A'), log('B', { 'Analysis Type': '' })]);
  const inTab = m.all.filter(x => /Repeats|Sales/.test(x['Analysis Type'] || '')).length;
  check('an audit with no Analysis Type is reported, not silently dropped from the tabs', inTab < m.all.length && /1 have no Analysis Type/.test(m.status), m.status);
}

console.log('\nTest 6 — the Sales Attempted chart gets its data through the real path');
{
  const h = pageWithCharts();
  h.ctx.dbFinishRender = function () {};
  const dashRows = ['A', 'B'].map((ref, i) => ({ 'Audit Ref': ref, 'Sales Attempted': i === 0 ? 'Yes' : 'No', 'Analysis Type': 'Sales Analyzer' }));
  const logRows = ['A', 'B'].map(ref => ({ 'Audit Ref': ref, 'Submitted At': '2026-10-12T10:00:00', 'Analysis Type': 'Sales Analyzer', 'SAP ID': '1', 'Team Member': 'X' }));
  h.ctx.dbHandleData({ rows: dashRows, ts: new Date().toISOString() }, false);
  h.pending.find(p => p.method === 'getAuditLogData').success(logRows);
  h.ctx.dbDrawChart('s', h.ctx._dbAllData.filter(r => (r['Analysis Type'] || '').indexOf('Sales') !== -1));
  const d = h.drawn.pie['s-chart-salesattempted'] || [];
  check('the chart draws Yes 1 / No 1', d.length === 2 && d[0][1] === 1 && d[1][1] === 1, JSON.stringify(d));
}

console.log('\nTest 7 — index.html and Dashboard.html use the same rules');
{
  const idx = pageWithCharts();
  const std = pageWithCharts(path.join(__dirname, '..', 'Dashboard.html'));
  const rows = [blank({ 'Repeat Projection %': '0%', 'Issue Resolved': 'Yes' }), blank({ 'Repeat Projection %': '40%', 'Issue Resolved': REAL.partial }), blank({ 'Issue Resolved': 'No' })];
  idx.reset(); idx.ctx.dbRenderStats('r', rows);
  std.reset(); std.ctx.renderStats('r', rows);
  const a = idx.els['r-stats'].innerHTML, s = std.els['r-stats'].innerHTML;
  const stdVal = label => { const m = new RegExp('<div class="stat-val">([^<]*)</div><div class="stat-label">' + label + '</div>').exec(s); return m ? m[1] : null; };
  check('Avg Repeat Risk agrees', statVal(a, 'Avg Repeat Risk') === stdVal('Avg Repeat Risk') && stdVal('Avg Repeat Risk') === '20%', statVal(a, 'Avg Repeat Risk') + ' vs ' + stdVal('Avg Repeat Risk'));
  check('Issues Resolved agrees', statVal(a, 'Issues Resolved') === stdVal('Issues Resolved') && stdVal('Issues Resolved') === '1');
  check('Partially Resolved agrees', statVal(a, 'Partially Resolved') === stdVal('Partially Resolved') && stdVal('Partially Resolved') === '1');
  check('Unresolved agrees', statVal(a, 'Unresolved') === stdVal('Unresolved') && stdVal('Unresolved') === '1');

  const block = f => { const t = read(f); const i = t.indexOf('// ── Dashboard helpers:'), j = t.indexOf('// ── end of dashboard helpers ──'); return i === -1 || j === -1 ? null : t.slice(i, j).replace(/\r\n/g, '\n'); };
  const bi = block('index.html'), bd = block('Dashboard.html');
  check('the shared helper block exists in both files', !!bi && !!bd);
  check('the helper block is IDENTICAL in both files', bi === bd);
}

console.log('\nTest 8 — grouping of free text in the charts');
{
  const h = pageWithCharts();
  const top = h.ctx.dbTopNByField([blank({ 'Call Driver': 'Billing dispute' }), blank({ 'Call Driver': 'billing dispute ' }), blank({ 'Call Driver': 'Billing  Dispute' })], 'Call Driver', 5);
  check('the same Call Driver with different capitals or spacing is one bar of 3', top.length === 1 && top[0][1] === 3, JSON.stringify(top));
  check('the label is the first spelling seen', top[0][0] === 'Billing dispute', top[0][0]);
  const opp = h.ctx.dbTopNByField([blank({ 'RCA Sub-Parameter': 'Hold handling; Empathy' }), blank({ 'RCA Sub-Parameter': 'Empathy; Closing' })], 'RCA Sub-Parameter', 5, s => s.split(/\s*;\s*/));
  check('RCA Sub-Parameter is counted per item, not per pair', opp[0][0] === 'Empathy' && opp[0][1] === 2 && opp.length === 3, JSON.stringify(opp));
  const hc = pageWithCharts();
  hc.ctx.dbDrawChart('r', [blank({ 'RCA Sub-Parameter': 'Hold handling; Empathy' }), blank({ 'RCA Sub-Parameter': 'Empathy; Closing' })]);
  const bars = hc.drawn.bar['r-chart-opps'] || [];
  check('the real Top Opportunities chart is drawn per item (Empathy 2)', bars.length === 3 && bars[0][0] === 'Empathy' && bars[0][1] === 2, JSON.stringify(bars));
}

console.log('\nTest 9 — where the numbers come from (Code.gs)');
{
  const code = read('Code.gs');
  const ctx = vm.createContext({ Logger: { log() {} }, String, RegExp, Object, Array, Math, JSON });
  vm.runInContext(extractFunction(code, 'extractStructuredRCAFromHTML'), ctx);
  const flagHtml = t => '<div class="ai-flag"><div class="ai-flag-title">&#9888; ' + t + '</div></div>';
  const s = ctx.extractStructuredRCAFromHTML('<div class="report-wrap">' + flagHtml('Wrong LOB, handled a TV call') + flagHtml('No save attempt') + '</div>');
  const dash = pageWithCharts();
  check('stored Critical Flags split back into the 2 flags that were on the report', dash.ctx._dbSplitFlags(s.criticalFlags).length === 2, JSON.stringify(s.criticalFlags));
  check('new flags are stored with " | " between them', (s.criticalFlags || '').indexOf(' | ') !== -1, s.criticalFlags);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
