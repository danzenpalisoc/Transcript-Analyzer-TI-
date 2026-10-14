/**
 * One filter, the same audits and the same numbers on every tab (app-wide review of the dashboard).
 *
 * A generated set of audits goes through the REAL path (dbHandleData: Audit_Log + Dashboard_Data merged, as on the live page), then many filter
 * combinations are applied with the real dbApplyAll. For each one the audits that every tab shows, the tiles, the tables, the charts, the exports,
 * the scorecard selection and the text the AI is given are compared with an answer worked out here from the raw data with simple, separate code.
 *
 * It proves the LOGIC against a dataset with the awkward cases (blank cells in Audit_Log that Dashboard_Data fills, names with stray spaces, flags written
 * twice, "None", the old comma format, dates on the edge of a day, audits with no date or no type). It cannot see the real sheets: verifyDashboardData()
 * does that.
 *
 * Usage: node tests/filter-consistency-test.js
 */
const vm = require('vm');
const { load } = require('./lib/page-harness');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

// ── the page, with Google Charts replaced by a recorder ────────────────────────────────
const h = load();
h.ctx.alert = () => {};
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

// The harness has no real checkboxes: the pill stores are small fakes that behave like the page's (see dashboard-refresh-test.js).
const FIELD = { member: 'Team Member', leader: 'Team Leader', ops: 'Operations Manager', lob: 'Line of Business', locale: 'Locale', vtid: 'VTID', atype: 'Analysis Type', ref: 'Audit Ref', qa: 'Observer Name' };
const stores = {};
h.ctx.dbPopulateFilters = function () {
  Object.keys(FIELD).forEach(key => {
    const store = h.document.getElementById('sh-ms-' + key);
    const values = [...new Set([].concat(h.ctx._dbAllData || [], h.ctx._dbAuditLog || []).map(r => r[FIELD[key]]).filter(Boolean))];
    store.boxes = values.map(v => ({ value: v, checked: true }));
    store.querySelectorAll = sel => (/:checked/.test(sel) ? store.boxes.filter(b => b.checked) : store.boxes);
    stores[key] = store;
  });
};

// ── a generated set of audits ──────────────────────────────────────────────────────────
let seed = 20261014;
const rnd = () => { seed = (seed * 1664525 + 1013904223) % 4294967296; return seed / 4294967296; };
const pick = a => a[Math.floor(rnd() * a.length)];
const AGENTS = ['Maria Santos', 'Aisha Khan', 'Carlos Mendoza', 'Li Wei', 'Nora Dela Cruz', 'Omar Haddad', 'Priya Nair', 'Sam Reyes'];
const TLS = ['TL One', 'TL Two', 'TL Three'], OPS = ['Ops A', 'Ops B'], LOBS = ['MOB', 'HOME', 'ONE CLS'], LOCALES = ['PH', 'MA', 'CO'], VTIDS = ['V1', 'V2', 'V3'], QAS = ['QA One', 'QA Two', 'QA Three'];
const FLAGS = ['No Proactive Expectation-Setting Before Transfer', 'Incomplete authentication', 'No retention attempt', 'Callback Number Not Confirmed at Start of Call', 'Zero product presentation'];
const RESOLVED = ['Yes', 'No', 'Partially resolved - credit has not posted yet', '', 'Yes - fixed the billing error'];
const TIMES = ['2026-10-01T10:00:00', '2026-10-02T00:00:00', '2026-10-02T12:30:00', '2026-10-02T23:59:59', '2026-10-03T00:00:01', '2026-10-04T09:15:00', '2026-10-05T18:45:00', '2026-10-06T23:59:59', '2026-10-07T07:00:00'];

function flagCell() {
  const n = Math.floor(rnd() * 5);
  if (n === 0) return pick(['', 'None']);
  const list = []; for (let i = 0; i < n; i++) list.push(pick(FLAGS));
  if (rnd() < 0.2) list.push(list[0]);                       // the same flag written twice in one audit
  if (rnd() < 0.15) return list.slice(0, 2).join(', ');      // the old comma format (the flag titles used here have no comma, so this is a plain list)
  return list.join(' | ');
}
const dashRows = [], logRows = [];
const N = 150;
for (let i = 1; i <= N; i++) {
  const ref = 'NHA-' + String(i).padStart(4, '0');
  const sales = rnd() < 0.3;
  const type = sales ? 'Sales Analyzer' : 'Repeats & Transfer Analyzer';
  const agent = pick(AGENTS), tl = pick(TLS), ops = pick(OPS), lob = pick(LOBS), locale = pick(LOCALES);
  const ts = i % 40 === 0 ? '' : pick(TIMES);                 // an audit with no date at all
  const dash = { 'Audit Ref': ref, 'Timestamp': ts, 'Team Member': agent, 'Team Leader': tl, 'Operations Manager': ops, 'Line of Business': lob, 'Locale': locale, 'VTID': pick(VTIDS),
    'Analysis Type': type, 'Critical Flags': flagCell(), 'Repeat Projection %': pick(['40%', '80', '', '0%', '65%']), 'Issue Resolved': pick(RESOLVED), 'Call Driver': pick(['Billing', 'Tech', 'Plan change']),
    'Observer Name': pick(QAS), 'Email Status': pick(['Sent', 'Not Sent']), 'SAP ID': String(1000 + i) };
  if (sales) dash['Sales Attempted'] = pick(['Yes', 'No']);
  if (i % 17 !== 0) dashRows.push(dash);                       // some audits have no Dashboard_Data row
  const log = { 'Audit Ref': ref, 'Submitted At': ts, 'SAP ID': String(1000 + i), 'Interaction ID': 'INT-' + i, 'Team Member': agent, 'Team Leader': tl, 'Operations Manager': ops,
    'Line of Business': lob, 'Locale': locale, 'VTID': dash['VTID'], 'Observer Name': dash['Observer Name'], 'Analysis Type': type, 'Email Status': dash['Email Status'],
    'Repeat Projection %': dash['Repeat Projection %'], 'Issue Resolved': dash['Issue Resolved'] };
  if (i % 7 === 0) log['Team Member'] = '';                    // blank in Audit_Log, filled from Dashboard_Data
  if (i % 11 === 0) log['Line of Business'] = '';
  if (i % 13 === 0) log['Submitted At'] = '';
  if (i % 19 === 0) log['Team Member'] = agent + '  ';         // stray spaces
  if (i % 23 === 0) { log['Interaction ID'] = ''; }
  if (i % 29 === 0) log['Observer Name'] = '';
  if (i === 5) { log['Team Member'] = ''; dash['Team Member'] = ''; }   // blank in both
  logRows.push(log);
}
// ── what every audit is, worked out here without the page's code ───────────────────────
const truth = logRows.map(log => {
  const d = dashRows.find(x => x['Audit Ref'] === log['Audit Ref']) || {};
  const pickv = (a, b) => (log[a] !== undefined && String(log[a]).trim() !== '' ? log[a] : (d[b || a] || ''));
  return { ref: log['Audit Ref'], hasDash: !!d['Audit Ref'], 'Timestamp': log['Submitted At'] || d['Timestamp'] || '', 'Team Member': log['Team Member'] || (d['Team Member'] || ''), 'Team Leader': pickv('Team Leader'),
    'Operations Manager': pickv('Operations Manager'), 'Line of Business': log['Line of Business'] || d['Line of Business'] || '', 'Locale': pickv('Locale'), 'VTID': pickv('VTID'),
    'Observer Name': log['Observer Name'] || d['Observer Name'] || '', 'Analysis Type': log['Analysis Type'] || d['Analysis Type'] || '', 'SAP ID': log['SAP ID'] || d['SAP ID'] || '',
    'Email Status': log['Email Status'] || d['Email Status'] || '', flags: d['Critical Flags'] || '', repeat: log['Repeat Projection %'] || d['Repeat Projection %'] || '',
    resolved: log['Issue Resolved'] || d['Issue Resolved'] || '', 'Sales Attempted': d['Sales Attempted'] || '', driver: d['Call Driver'] || '' };
});
// Audits that exist only in Dashboard_Data are none here (every audit has an Audit_Log row), so the truth is Audit_Log plus the fill.

const splitFlags = s => { const raw = String(s || '').trim(); if (!raw || /^none$/i.test(raw)) return []; const parts = raw.indexOf(' | ') !== -1 ? raw.split(' | ') : raw.split(','); const seen = {}; return parts.map(p => p.trim()).filter(p => p && !/^none$/i.test(p)).filter(p => { const k = p.toLowerCase(); if (seen[k]) return false; seen[k] = 1; return true; }); };
const verdict = v => { const s = String(v || '').trim().toLowerCase(); if (!s) return 'unknown'; if (/^partial/.test(s)) return 'partial'; if (/^no\b/.test(s)) return 'no'; if (/^yes\b/.test(s)) return 'yes'; return 'unknown'; };
const num = v => { const s = String(v || '').replace('%', '').trim(); if (s === '') return null; const n = Number(s); return isNaN(n) ? null : n; };
const avg = rows => { const v = rows.map(r => num(r.repeat)).filter(x => x !== null); return v.length ? Math.round(v.reduce((a, b) => a + b, 0) / v.length) : 0; };
const norm = s => String(s || '').replace(/\s+/g, ' ').trim();

// the filters, applied by simple code of the same meaning as the page's
function oracleFilter(rows, f) {
  return rows.filter(r => {
    if (f.dateFrom || f.dateTo) {
      const d = r['Timestamp'] ? new Date(r['Timestamp']) : null;
      if (!d || isNaN(d.getTime())) return false;
      if (f.dateFrom && d < new Date(f.dateFrom + 'T00:00:00')) return false;
      if (f.dateTo && d > new Date(f.dateTo + 'T23:59:59')) return false;
    }
    if (f.sapQ && String(r['SAP ID']).toLowerCase().indexOf(f.sapQ) === -1) return false;
    const inList = (k, field) => !(f[k] && f[k].length) || f[k].indexOf(r[field]) !== -1;
    return inList('members', 'Team Member') && inList('leaders', 'Team Leader') && inList('ops', 'Operations Manager') && inList('lobs', 'Line of Business') &&
      inList('locales', 'Locale') && inList('vtids', 'VTID') && inList('refs', 'ref') && inList('qas', 'Observer Name') &&
      (!(f.atypes && f.atypes.length) || f.atypes.some(a => r['Analysis Type'].toLowerCase().indexOf(a.toLowerCase()) !== -1));
  });
}

// ── load the data through the real path ────────────────────────────────────────────────
h.ctx.dbHandleData({ rows: dashRows, ts: new Date().toISOString() }, false);
h.pending.find(p => p.method === 'getAuditLogData').success(logRows);
const roster = h.pending.find(p => p.method === 'getRosterAll'); if (roster) roster.success([]);

const KEYS = { members: 'member', leaders: 'leader', ops: 'ops', lobs: 'lob', locales: 'locale', vtids: 'vtid', atypes: 'atype', refs: 'ref', qas: 'qa' };
function apply(f) {
  drawn.pie = {}; drawn.bar = {};   // a chart that is not drawn (no data) must not keep the previous filter's drawing
  Object.keys(KEYS).forEach(k => {
    const vals = f[k] || [];
    stores[KEYS[k]].boxes.forEach(b => { b.checked = !vals.length || vals.indexOf(b.value) !== -1; });
  });
  h.document.getElementById('sh-date-from').value = f.dateFrom || '';
  h.document.getElementById('sh-date-to').value = f.dateTo || '';
  h.document.getElementById('sh-sap-search').value = f.sapQ || '';
  h.ctx.dbApplyAll();
  h.ctx.renderOverviewTab();
}
const statVal = (html, label) => { const m = new RegExp('<div class="db-stat-val">([^<]*)</div><div class="db-stat-lbl">' + label + '</div>').exec(html); return m ? m[1] : null; };
const rowsIn = id => { const html = h.els[id].innerHTML; return /db-empty|No audit records/.test(html) ? 0 : (html.match(/<tr>/g) || []).length; };
const refsOf = rows => rows.map(r => r['Audit Ref']).sort().join(',');
const sum = rows => rows.reduce((a, r) => a + r[1], 0);

// server side: the text the AI is given
const code = read('Code.gs');
const sctx = vm.createContext({ PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) } });
vm.runInContext(read('ScorecardMetrics.gs'), sctx);
['dbUniqServer', 'buildAnalyticsContext_'].forEach(n => vm.runInContext(extractFunction(code, n), sctx));

const SPACED = (logRows.find(l => /s$/.test(l['Team Member'])) || {})['Team Member'];
const FILTERS = [
  ['no filter', {}],
  ['one agent', { members: ['Maria Santos'] }],
  ['two agents', { members: ['Aisha Khan', 'Li Wei'] }],
  ['an agent written with stray spaces', { members: [SPACED] }],
  ['a team leader', { leaders: ['TL Two'] }],
  ['an ops manager', { ops: ['Ops B'] }],
  ['a line of business (blank in Audit_Log for some)', { lobs: ['HOME'] }],
  ['another line of business', { lobs: ['MOB', 'ONE CLS'] }],
  ['a locale', { locales: ['MA'] }],
  ['a VTID', { vtids: ['V2'] }],
  ['a QA observer', { qas: ['QA Two'] }],
  ['Sales only', { atypes: ['Sales Analyzer'] }],
  ['Repeats only', { atypes: ['Repeats & Transfer Analyzer'] }],
  ['one audit ref', { refs: ['NHA-0014'] }],
  ['a SAP ID search', { sapQ: '101' }],
  ['dates from', { dateFrom: '2026-10-02' }],
  ['dates to', { dateTo: '2026-10-03' }],
  ['one day (rows at 00:00:00 and 23:59:59)', { dateFrom: '2026-10-02', dateTo: '2026-10-02' }],
  ['a range', { dateFrom: '2026-10-03', dateTo: '2026-10-06' }],
  ['agent + LOB', { members: ['Maria Santos', 'Sam Reyes'], lobs: ['MOB'] }],
  ['date + type + leader', { dateFrom: '2026-10-02', dateTo: '2026-10-05', atypes: ['Sales Analyzer'], leaders: ['TL One', 'TL Three'] }],
  ['everything narrowed', { members: ['Priya Nair'], leaders: ['TL One'], ops: ['Ops A'], lobs: ['MOB'], locales: ['PH'], vtids: ['V1'], qas: ['QA One'], dateFrom: '2026-10-01', dateTo: '2026-10-07' }],
  ['nothing matches (a date after every audit)', { members: ['Maria Santos'], dateFrom: '2027-01-01' }],
];

FILTERS.forEach(([name, f]) => {
  console.log('\nFilter — ' + name);
  const full = Object.assign({ dateFrom: '', dateTo: '', sapQ: '' }, f);
  apply(full);
  const all = oracleFilter(truth, full);
  const R = all.filter(r => r['Analysis Type'].indexOf('Repeats') !== -1), S = all.filter(r => r['Analysis Type'].indexOf('Sales') !== -1);
  const AL = all;   // the Audit Log has every audit that passes the filter, with or without a type

  // 1. the same audits
  check('Repeats shows the audits the filter selects', refsOf(h.ctx._dbRFiltered) === refsOf(R.map(r => ({ 'Audit Ref': r.ref }))), h.ctx._dbRFiltered.length + ' vs ' + R.length);
  check('Sales shows the audits the filter selects', refsOf(h.ctx._dbSFiltered) === refsOf(S.map(r => ({ 'Audit Ref': r.ref }))), h.ctx._dbSFiltered.length + ' vs ' + S.length);
  check('the Audit Log shows the audits the filter selects (the same ones as Repeats and Sales, plus any with no type)', refsOf(h.ctx._dbAlFiltered) === refsOf(AL.map(r => ({ 'Audit Ref': r.ref }))), h.ctx._dbAlFiltered.length + ' vs ' + AL.length);

  // 2. the tiles of Repeats and Sales
  [['r', R], ['s', S]].forEach(([p, rows]) => {
    const html = h.els[p + '-stats'].innerHTML;
    const c = { yes: 0, partial: 0, no: 0, unknown: 0 }; rows.forEach(r => c[verdict(r.resolved)]++);
    const flagged = rows.filter(r => splitFlags(r.flags).length > 0).length;
    check(p + ': Total Interactions', statVal(html, 'Total Interactions') === String(rows.length), statVal(html, 'Total Interactions') + ' vs ' + rows.length);
    check(p + ': Issues Resolved', statVal(html, 'Issues Resolved') === String(c.yes), statVal(html, 'Issues Resolved') + ' vs ' + c.yes);
    check(p + ': Partially Resolved', statVal(html, 'Partially Resolved') === String(c.partial));
    check(p + ': Unresolved', statVal(html, 'Unresolved') === String(c.no));
    check(p + ': Not Recorded (shown only when there are some)', (c.unknown ? statVal(html, 'Not Recorded') === String(c.unknown) : statVal(html, 'Not Recorded') === null), String(statVal(html, 'Not Recorded')) + ' vs ' + c.unknown);
    check(p + ': Avg Repeat Risk', statVal(html, 'Avg Repeat Risk') === avg(rows) + '%', statVal(html, 'Avg Repeat Risk') + ' vs ' + avg(rows) + '%');
    check(p + ': Audits with Flags counts audits with at least one real flag ("None" is none, a repeated flag is one)', statVal(html, 'Audits with Flags') === String(flagged), statVal(html, 'Audits with Flags') + ' vs ' + flagged);
    check(p + ': Audits Emailed', statVal(html, 'Audits Emailed') === String(rows.filter(r => r['Email Status'] === 'Sent').length));
    // 3. the table and the count under the filters
    check(p + ': the table has one row per audit', rowsIn(p + '-tbody') === rows.length, rowsIn(p + '-tbody') + ' vs ' + rows.length);
    check(p + ': the record count text', h.els[p + '-row-count'].textContent === rows.length + ' record' + (rows.length !== 1 ? 's' : ''), h.els[p + '-row-count'].textContent);
    // 4. the charts
    if (rows.length) {
      const per = drawn.pie[p + '-chart-flagdist'] || [];
      check(p + ': the Critical Flags per Audit chart adds up to the audits', sum(per.filter(x => x[0] !== 'No audits')) === rows.length, JSON.stringify(per));
      const buckets = [0, 0, 0, 0, 0]; rows.forEach(r => buckets[Math.min(splitFlags(r.flags).length, 4)]++);
      check(p + ': ...and each slice is right', JSON.stringify(per.map(x => x[1])) === JSON.stringify(buckets.filter(x => x > 0)), JSON.stringify(per) + ' vs ' + JSON.stringify(buckets));
      const bar = drawn.bar[p + '-chart-flags'] || [];
      const counts = {}; rows.forEach(r => splitFlags(r.flags).forEach(fl => { counts[fl] = (counts[fl] || 0) + 1; }));
      const top = Object.keys(counts).sort((a, b) => counts[b] - counts[a]);
      check(p + ': the Top Critical Flags chart leads with the most common flag and its count (no bar when no flags)', top.length ? (bar.length && bar[0][1] === counts[top[0]]) : bar.length === 0, JSON.stringify(bar[0]) + ' vs ' + top[0] + ' ' + counts[top[0]]);
      check(p + ': ...and no bar is higher than the real count of its flag', bar.every(b => counts[b[0]] === b[1]), JSON.stringify(bar));
      if (p === 'r') {
        const res = drawn.pie['r-chart-resolution'] || [];
        const cc = { yes: 0, partial: 0, no: 0, unknown: 0 }; rows.forEach(r => cc[verdict(r.resolved)]++);
        check('r: the Resolved vs Unresolved chart equals the tiles', JSON.stringify(res.map(x => x[1])) === JSON.stringify([cc.yes, cc.partial, cc.no, cc.unknown]), JSON.stringify(res));
      } else {
        const sa = drawn.pie['s-chart-salesattempted'] || [];
        const yes = rows.filter(r => /^yes$/i.test(r['Sales Attempted'])).length, no = rows.filter(r => /^no$/i.test(r['Sales Attempted'])).length;
        check('s: the Sales Attempted chart counts only the audits that pass the filter', yes + no === 0 || (sa.length === 2 && sa[0][1] === yes && sa[1][1] === no), JSON.stringify(sa) + ' vs ' + yes + '/' + no);
      }
    }
  });

  // the Call Driver bar charts
  const driverCounts = rows => { const c = {}; rows.forEach(r => { if (r.driver) c[r.driver] = (c[r.driver] || 0) + 1; }); return c; };
  const driversOk = (bar, rows) => { const c = driverCounts(rows), max = Math.max.apply(null, [0].concat(Object.keys(c).map(k => c[k]))); return bar.every(b => c[b[0]] === b[1]) && (max === 0 ? bar.length === 0 : (bar.length > 0 && bar[0][1] === max)); };
  check('r: the Call Drivers chart counts only the audits of the view', driversOk(drawn.bar['r-chart-drivers'] || [], R), JSON.stringify(drawn.bar['r-chart-drivers']));
  check('s: the Call Drivers chart counts only the audits of the view', driversOk(drawn.bar['s-chart-drivers'] || [], S), JSON.stringify(drawn.bar['s-chart-drivers']));
  if (R.concat(S).length) check('Overview: the Call Drivers chart counts the audits of Repeats and Sales together', driversOk(drawn.bar['ov-chart-drivers'] || [], R.concat(S)), JSON.stringify(drawn.bar['ov-chart-drivers']));

  // Home counts everything (it is not filtered); with no filter it must say what the Overview says
  if (name === 'no filter') {
    h.ctx.homeUpdate();
    const typed = truth.filter(r => /Repeats|Sales/.test(r['Analysis Type']));
    check('Home: evaluations = the Overview total with no filter', h.els['home-stat-evals'].textContent === statVal(h.els['ov-stats'].innerHTML, 'Total Evaluations') && h.els['home-stat-evals'].textContent === String(typed.length), h.els['home-stat-evals'].textContent);
    check('Home: agents = the Overview unique agents', h.els['home-stat-agents'].textContent === statVal(h.els['ov-stats'].innerHTML, 'Unique Agents'), h.els['home-stat-agents'].textContent);
    check('Home: the Repeats / Sales split', h.els['home-stat-split'].textContent === R.length + ' / ' + S.length, h.els['home-stat-split'].textContent);
  }

  // 5. the Overview is Repeats and Sales together
  const ov = h.els['ov-stats'].innerHTML, both = R.concat(S);
  if (both.length) {
    const agents = new Set(both.map(r => norm(r['Team Member'])).filter(Boolean));
    const flags = both.reduce((a, r) => a + splitFlags(r.flags).length, 0);
    check('Overview: Total Evaluations = Repeats + Sales', statVal(ov, 'Total Evaluations') === String(both.length), statVal(ov, 'Total Evaluations') + ' vs ' + both.length);
    check('Overview: the split is right', statVal(ov, 'Repeats & Transfer') === String(R.length) && statVal(ov, 'Sales Analyzed') === String(S.length));
    check('Overview: Avg Repeat Risk counts the audits of both', statVal(ov, 'Avg Repeat Risk') === avg(both) + '%', statVal(ov, 'Avg Repeat Risk') + ' vs ' + avg(both) + '%');
    check('Overview: Unique Agents (same name with stray spaces is one agent)', statVal(ov, 'Unique Agents') === String(agents.size), statVal(ov, 'Unique Agents') + ' vs ' + agents.size);
    check('Overview: Total Critical Flags = the flags of Repeats and Sales (a repeated flag counts once per audit)', statVal(ov, 'Total Critical Flags') === String(flags), statVal(ov, 'Total Critical Flags') + ' vs ' + flags);
    const tm = h.els['ov-tm-tbody'].innerHTML;
    const tmFlags = [...tm.matchAll(/font-weight:700">(\d+)<\/td><td style="max-width:200px/g)].reduce((a, m) => a + Number(m[1]), 0);
    const named = both.filter(r => { const m = norm(r['Team Member']); return m; }).reduce((a, r) => a + splitFlags(r.flags).length, 0);
    check('Overview: the agent table adds up to the flags of the agents it lists', tmFlags === named, tmFlags + ' vs ' + named);
    const ty = drawn.pie['ov-chart-type'] || [];
    check('Overview: the Repeats vs Sales donut', ty.length === 2 && ty[0][1] === R.length && ty[1][1] === S.length, JSON.stringify(ty));
  } else {
    check('Overview: no audits says so instead of showing numbers', /No records match/.test(h.els['ov-highlights'].innerHTML), h.els['ov-highlights'].innerHTML.slice(0, 100));
  }

  // 6. the Audit Log tiles against the other tabs
  const al = h.els['al-stats'].innerHTML;
  check('Audit Log: Total Audits = the rows of the Audit Log', statVal(al, 'Total Audits') === String(AL.length) && rowsIn('al-tbody') === AL.length, statVal(al, 'Total Audits') + ' / ' + rowsIn('al-tbody') + ' vs ' + AL.length);
  check('Audit Log: Total Audits = Repeats + Sales + audits with no type', AL.length === R.length + S.length + AL.filter(r => !/Repeats|Sales/.test(r['Analysis Type'])).length);
  const sentAL = AL.filter(r => r['Email Status'] === 'Sent').length;
  check('Audit Log: Emailed', statVal(al, 'Emailed') === String(sentAL), statVal(al, 'Emailed') + ' vs ' + sentAL);
  check('Audit Log: Emailed = Audits Emailed of Repeats + Sales (every audit here has a type)', sentAL === R.filter(r => r['Email Status'] === 'Sent').length + S.filter(r => r['Email Status'] === 'Sent').length);
  check('Audit Log: Unique Agents agrees with the Overview', both.length === 0 || statVal(al, 'Unique Agents') === statVal(ov, 'Unique Agents'), statVal(al, 'Unique Agents') + ' vs ' + statVal(ov, 'Unique Agents'));

  // 7. exports and the scorecard take the audits of the view
  const exp = (fn, pfx) => { h.pending.length = 0; h.ctx[fn](pfx); const c = h.pending.find(p => p.method === 'exportAudits'); h.pending.length = 0; return c ? { view: c.args[0], refs: c.args[1].slice().sort() } : null; };
  [['r', R], ['s', S]].forEach(([p, rows]) => {
    const e = exp('dbExportFiltered', p);
    check(p + ': "Export Filtered" sends exactly the audits of the view', rows.length === 0 ? e === null : (!!e && e.refs.join(',') === rows.map(r => r.ref).sort().join(',')), e ? e.refs.length + ' vs ' + rows.length : 'nothing sent');
    const fl = rows.filter(r => splitFlags(r.flags).length > 0);
    const e2 = exp('dbFlagOnly', p);
    check(p + ': "Flagged only" sends the audits with a real flag ("None" is not one)', fl.length === 0 ? e2 === null : (!!e2 && e2.refs.join(',') === fl.map(r => r.ref).sort().join(',')), e2 ? e2.refs.length + ' vs ' + fl.length : 'nothing sent');
  });
  const eal = exp('dbExportFiltered', 'al');
  check('al: "Export Filtered" sends exactly the audits of the Audit Log view', AL.length === 0 ? eal === null : (!!eal && eal.refs.join(',') === AL.map(r => r.ref).sort().join(',')));
  const sel = h.ctx.dbScorecardSelection();
  check('the scorecard takes the audits of Repeats and Sales', sel.refs.slice().sort().join(',') === both.map(r => r.ref).sort().join(','), sel.refs.length + ' vs ' + both.length);

  // 8. what the AI is given
  const ai = h.ctx._dbAiFilters('sales');
  check('the AI is told the filters that are applied and the tab', ai.analysisType === 'sales' && ai.dateFrom === (full.dateFrom || '') && ai.dateTo === (full.dateTo || '') && JSON.stringify((ai.members || []).slice().sort()) === JSON.stringify((full.members || []).slice().sort()), JSON.stringify(ai).slice(0, 160));
  check('the Overview tells the AI no tab', h.ctx._dbAiFilters('').analysisType === undefined);
  [['sales', S], ['repeats', R]].forEach(([kind, rows]) => {
    if (!rows.length) return;
    const dashRowsOfTab = (kind === 'sales' ? h.ctx._dbSFiltered : h.ctx._dbRFiltered);
    const txt = sctx.buildAnalyticsContext_(dashRowsOfTab, h.ctx._dbAiFilters(kind), kind);
    const c = { yes: 0, partial: 0, no: 0, unknown: 0 }; rows.forEach(r => c[verdict(r.resolved)]++);
    check(kind + ': the AI text names the right analyzer', new RegExp('ANALYSIS TYPE: ' + (kind === 'sales' ? 'Sales Analyzer' : 'Repeats & Transfer Analyzer')).test(txt));
    check(kind + ': the AI text has the audits of the tab and the tiles\' resolved count', txt.indexOf('TOTAL INTERACTIONS: ' + rows.length + '\n') !== -1 && txt.indexOf('Issues Resolved: ' + c.yes + ' of ' + rows.length) !== -1, txt.split('\n').slice(0, 12).join(' / ').slice(0, 220));
    check(kind + ': the AI text has the tiles\' repeat risk', txt.indexOf('Average Repeat Risk: ' + avg(rows) + '%') !== -1);
    const counts = {}; rows.forEach(r => splitFlags(r.flags).forEach(fl => { counts[fl] = (counts[fl] || 0) + 1; }));
    const top = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0];
    check(kind + ': the AI text counts the most common flag as the chart does', !top || txt.indexOf('- ' + top + ' (' + counts[top] + ')') !== -1, top + ' ' + counts[top]);
    const dated = !!(full.dateFrom || full.dateTo);
    check(kind + ': the AI text states the date range that is applied (not "All time" under a date filter)', dated ? txt.indexOf('DATE RANGE: ' + (full.dateFrom || 'All time')) !== -1 : txt.indexOf('DATE RANGE: All time') !== -1);
  });
  if (both.length) {
    const txt = sctx.buildAnalyticsContext_(h.ctx._dbRFiltered.concat(h.ctx._dbSFiltered), h.ctx._dbAiFilters(''), 'overview');
    check('overview: the AI text says it covers both analyzers, with all the audits', /ANALYSIS TYPE: All analyzers/.test(txt) && txt.indexOf('TOTAL INTERACTIONS: ' + both.length + '\n') !== -1, txt.split('\n')[0]);
  }
});

console.log('\nThe AI Analytics panel uses the same day boundaries as the dashboard');
{
  // rows right on the edge of a day: 00:00:00 and 23:59:59 local time belong to that day
  const rows = [{ 'Timestamp': '2026-10-02T00:00:00', 'Audit Ref': 'a' }, { 'Timestamp': '2026-10-02T23:59:59', 'Audit Ref': 'b' }, { 'Timestamp': '2026-10-01T23:59:59', 'Audit Ref': 'c' }, { 'Timestamp': '2026-10-03T00:00:00', 'Audit Ref': 'd' }, { 'Timestamp': '', 'Audit Ref': 'e' }];
  h.ctx._lastAnalyticsType = 'r'; h.ctx._dbRFiltered = rows; h.ctx.anApplyFilters = () => {};
  h.ctx._anFilters = { dateFrom: '2026-10-02', dateTo: '2026-10-02', member: [], leader: [], ops: [], lob: [], locale: [], vtid: [], atype: [], ref: [], sapQ: '' };
  const got = h.ctx.anGetFilteredRows().map(r => r['Audit Ref']).join(',');
  check('one day: both edges are in, the neighbours and the audit with no date are out', got === 'a,b', got);
}

console.log('\nThe filters work the same when an Audit_Log cell is blank and Dashboard_Data has the value');
{
  apply({ dateFrom: '', dateTo: '', sapQ: '', lobs: ['HOME'] });
  const blanksFilled = h.ctx._dbAlFiltered.filter(r => !logRows.find(l => l['Audit Ref'] === r['Audit Ref'])['Line of Business']).length;
  const expected = truth.filter(r => r['Line of Business'] === 'HOME' && !logRows.find(l => l['Audit Ref'] === r.ref)['Line of Business']).length;
  check('this data has audits whose LOB is blank in Audit_Log (so the check is real)', expected > 0, String(expected));
  check('the Audit Log finds them under the LOB filter, as Repeats and Sales do', blanksFilled === expected, blanksFilled + ' vs ' + expected);
  const shown = h.els['al-tbody'].innerHTML;
  check('and the table shows the filled value', (shown.match(/<td>HOME<\/td>/g) || []).length === h.ctx._dbAlFiltered.length, String((shown.match(/<td>HOME<\/td>/g) || []).length));
}

console.log('\nThe AI Highlights and Lowlights call tells the AI which tab it is (through the real server function)');
{
  let prompt = '';
  Object.assign(sctx, {
    Utilities: { computeDigest: () => [1, 2, 3], DigestAlgorithm: { MD5: 1 } },
    readAnalyticsCache_: () => null, writeAnalyticsCache_: () => {},
    FUELIX_CONFIG: { model: 'm', apiKey: 'k', baseUrl: 'https://x' },
    UrlFetchApp: { fetch: (u, o) => { prompt = JSON.parse(o.payload).messages[0].content; return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ choices: [{ message: { content: '{"highlights":[],"lowlights":[]}' } }] }) }; } },
  });
  vm.runInContext(extractFunction(code, 'getOverviewHighlightsLowlights'), sctx);
  const rows = [{ 'Audit Ref': 'A', 'Team Member': 'X', 'Critical Flags': '' }];
  [['sales', 'Sales Analyzer'], ['repeats', 'Repeats & Transfer Analyzer'], ['', 'All analyzers']].forEach(([kind, label]) => {
    prompt = '';
    sctx.getOverviewHighlightsLowlights(rows, h.ctx._dbAiFilters(kind));
    check('the "' + (kind || 'overview') + '" call reaches the AI as: ' + label, prompt.indexOf('ANALYSIS TYPE: ' + label) !== -1, prompt.split('\n').filter(l => /ANALYSIS TYPE/.test(l))[0] || 'no prompt');
  });
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
