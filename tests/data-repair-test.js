/**
 * DataRepair.gs: addMissingHeaders() and diagnoseDataIssues().
 *
 * addMissingHeaders may only ever add a header at the END of a header row, may never move
 * or clear data, and must do nothing when the headers are out of order. diagnoseDataIssues
 * is read-only and must find each problem verifyDashboardData() reports.
 *
 * Usage: node tests/data-repair-test.js
 */
const vm = require('vm');
const { read, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

function makeEnv(headerOverrides) {
  const logs = [];
  const who = { email: 'danzen.palisoc@telus.com' };
  const sandbox = {
    Logger: { log: m => logs.push(String(m)) }, JSON, Array, Object, String, Number, Math, Date, RegExp, isFinite, isNaN, Error,
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Session: { getActiveUser: () => ({ getEmail: () => who.email }), getScriptTimeZone: () => 'UTC' },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    // formats in UTC, so the tests do not depend on the time zone of the computer that runs them
    Utilities: { formatDate: (d, tz, f) => {
      const p = n => String(n).padStart(2, '0');
      const m = { yyyy: d.getUTCFullYear(), MM: p(d.getUTCMonth() + 1), dd: p(d.getUTCDate()), HH: p(d.getUTCHours()), H: d.getUTCHours(), mm: p(d.getUTCMinutes()), ss: p(d.getUTCSeconds()) };
      return f.replace(/yyyy|MM|dd|HH|H|mm|ss/g, k => String(m[k]));
    } },
  };
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('Config.gs'), ctx);
  vm.runInContext(read('ScorecardMetrics.gs'), ctx);
  vm.runInContext(read('DataRepair.gs'), ctx);
  const sheets = {};
  [['AUDIT_LOG_SHEET', 'AUDIT_LOG_HEADERS'], ['DASHBOARD_DATA_SHEET', 'DASHBOARD_HEADERS'], ['ANALYSIS_SHEET', 'ANALYSIS_HEADERS'],
   ['TRANSCRIPTS_SHEET', 'TRANSCRIPTS_HEADERS']].forEach(([s, h]) => {
    const hd = (headerOverrides && headerOverrides[s]) || ctx[h];
    sheets[ctx[s]] = new FakeSheet(hd, []);
  });
  const ss = { getSheetByName: n => sheets[n] || null };
  ctx.getOrCreateSpreadsheet = () => ss;
  ctx.appendRow = (sh, row) => sh.appendRow(row);
  let invalidated = 0;
  ctx.invalidateDashboardCache_ = () => { invalidated++; };
  const add = (sheetConst, o) => { const sh = sheets[ctx[sheetConst]]; sh.grid.push(sh.grid[0].map(h => (o[h] !== undefined ? o[h] : ''))); };
  return { ctx, sheets, logs, who, add, invalidated: () => invalidated };
}
const withoutLast = (arr, n) => arr.slice(0, arr.length - (n || 1));

console.log('\nTest 1 — a missing header is added at the end and no data moves');
{
  const e = makeEnv();
  const AL = e.ctx.AUDIT_LOG_HEADERS, DD = e.ctx.DASHBOARD_HEADERS;
  const shortAL = withoutLast(AL), shortDD = withoutLast(DD);
  const e2 = makeEnv({ AUDIT_LOG_SHEET: shortAL, DASHBOARD_DATA_SHEET: shortDD });
  e2.sheets[e2.ctx.AUDIT_LOG_SHEET].grid.push(shortAL.map((h, i) => 'v' + i));
  e2.sheets[e2.ctx.DASHBOARD_DATA_SHEET].grid.push(shortDD.map((h, i) => 'd' + i));
  const beforeAL = JSON.stringify(e2.sheets[e2.ctx.AUDIT_LOG_SHEET].grid.slice(1));
  const beforeDD = JSON.stringify(e2.sheets[e2.ctx.DASHBOARD_DATA_SHEET].grid.slice(1));
  const r = e2.ctx.addMissingHeaders();
  const hAL = e2.sheets[e2.ctx.AUDIT_LOG_SHEET].grid[0], hDD = e2.sheets[e2.ctx.DASHBOARD_DATA_SHEET].grid[0];
  check('Audit_Log now ends with Agent Email', hAL[hAL.length - 1] === 'Agent Email' && JSON.stringify(hAL) === JSON.stringify(AL), JSON.stringify(hAL.slice(-3)));
  check('Dashboard_Data now ends with Agent Email', hDD[hDD.length - 1] === 'Agent Email' && JSON.stringify(hDD) === JSON.stringify(DD));
  check('two headers were added', r.added === 2 && r.skipped === 0, JSON.stringify(r));
  check('no data row changed in Audit_Log', JSON.stringify(e2.sheets[e2.ctx.AUDIT_LOG_SHEET].grid.slice(1)) === beforeAL);
  check('no data row changed in Dashboard_Data', JSON.stringify(e2.sheets[e2.ctx.DASHBOARD_DATA_SHEET].grid.slice(1)) === beforeDD);
  check('the dashboard cache was cleared', e2.invalidated() === 1);
  const again = e2.ctx.addMissingHeaders();
  check('running it again changes nothing', again.added === 0 && again.skipped === 0, JSON.stringify(again));
}

console.log('\nTest 2 — it refuses to touch headers that are out of order, and never clears anything');
{
  const AL = makeEnv().ctx.AUDIT_LOG_HEADERS.slice();
  const swapped = AL.slice(); const t = swapped[3]; swapped[3] = swapped[4]; swapped[4] = t;
  const e = makeEnv({ AUDIT_LOG_SHEET: withoutLast(swapped) });
  const sh = e.sheets[e.ctx.AUDIT_LOG_SHEET];
  sh.grid.push(withoutLast(swapped).map((h, i) => 'v' + i));
  const before = JSON.stringify(sh.grid);
  const r = e.ctx.addMissingHeaders();
  check('swapped columns: skipped, nothing written', r.skipped === 1 && JSON.stringify(sh.grid) === before, JSON.stringify(r));
  check('the log says why', e.logs.some(l => /SKIPPED/.test(l) && /Audit_Log/.test(l)), e.logs.join(' | '));

  const e3 = makeEnv({ AUDIT_LOG_SHEET: AL.concat(['Extra']) });
  const r3 = e3.ctx.addMissingHeaders();
  check('a sheet with MORE headers than the code expects is skipped', r3.sheets[e3.ctx.AUDIT_LOG_SHEET] === 'skipped', JSON.stringify(r3));

  const e4 = makeEnv();
  const r4 = e4.ctx.addMissingHeaders();
  check('complete headers: nothing added, nothing skipped', r4.added === 0 && r4.skipped === 0, JSON.stringify(r4));
}

console.log('\nTest 3 — trailing blank header cells do not confuse it');
{
  const AL = makeEnv().ctx.AUDIT_LOG_HEADERS;
  const e = makeEnv({ AUDIT_LOG_SHEET: withoutLast(AL).concat(['', '']) });
  const r = e.ctx.addMissingHeaders();
  const h = e.sheets[e.ctx.AUDIT_LOG_SHEET].grid[0];
  check('Agent Email goes in the first empty column', h[AL.length - 1] === 'Agent Email' && r.added === 1, JSON.stringify(h.slice(-3)) + JSON.stringify(r));
}

console.log('\nTest 4 — only admins');
{
  const e = makeEnv();
  e.who.email = 'some.analyst@telus.com';
  let a = '', b = '';
  try { e.ctx.addMissingHeaders(); } catch (x) { a = x.message; }
  try { e.ctx.diagnoseDataIssues(); } catch (x) { b = x.message; }
  check('addMissingHeaders refuses an analyst', /admin/i.test(a), a);
  check('diagnoseDataIssues refuses an analyst', /admin/i.test(b), b);
}

console.log('\nTest 5 — diagnoseDataIssues finds each problem and changes nothing');
{
  const e = makeEnv();
  const T = new Date('2026-10-12T10:00:00');
  const log = (ref, o) => Object.assign({ 'Audit Ref': ref, 'Submitted At': T, 'Interaction ID': 'i-' + ref, 'SAP ID': '111', 'Team Member': 'Agent ' + ref, 'Analysis Type': 'Repeats & Transfer Analyzer', 'Repeat Projection %': '40%' }, o);
  const dash = (ref, o) => Object.assign({ 'Audit Ref': ref, 'Timestamp': T, 'Interaction ID': 'i-' + ref, 'SAP ID': '111', 'Team Member': 'Agent ' + ref, 'Analysis Type': 'Repeats & Transfer Analyzer' }, o);
  e.add('AUDIT_LOG_SHEET', log('OK')); e.add('DASHBOARD_DATA_SHEET', dash('OK'));
  e.add('AUDIT_LOG_SHEET', log('DUP')); e.add('AUDIT_LOG_SHEET', log('DUP', { 'SAP ID': '222' }));
  e.add('DASHBOARD_DATA_SHEET', dash('DUP')); e.add('DASHBOARD_DATA_SHEET', dash('DUP', { 'SAP ID': '222' }));
  e.add('DASHBOARD_DATA_SHEET', dash('HIDDEN'));
  e.add('AUDIT_LOG_SHEET', log('SAPX')); e.add('DASHBOARD_DATA_SHEET', dash('SAPX', { 'SAP ID': '999' }));
  e.add('AUDIT_LOG_SHEET', log('BADPCT', { 'Repeat Projection %': 'about forty' })); e.add('DASHBOARD_DATA_SHEET', dash('BADPCT'));
  const paragraph = 'This is a long paragraph of call text that landed in the Team Member column because the columns shifted by one position to the right';
  e.add('AUDIT_LOG_SHEET', log('SHIFT', { 'Team Member': paragraph })); e.add('DASHBOARD_DATA_SHEET', dash('SHIFT'));
  const snap = JSON.stringify(Object.keys(e.sheets).map(k => e.sheets[k].grid));
  const r = e.ctx.diagnoseDataIssues();
  check('the duplicate ref is found', r.duplicates.length === 1 && r.duplicates[0] === 'DUP', JSON.stringify(r.duplicates));
  check('the audit missing from Audit_Log is found', r.dashboardOnly.length === 1 && r.dashboardOnly[0] === 'HIDDEN', JSON.stringify(r.dashboardOnly));
  check('the SAP mismatch is found', r.sapMismatch.indexOf('SAPX') !== -1, JSON.stringify(r.sapMismatch));
  check('the non-numeric repeat value is found', r.nonNumericRepeat === 1, String(r.nonNumericRepeat));
  check('the shifted row is found', r.shiftedRows === 1, String(r.shiftedRows));
  check('the report shows the row for a duplicate in each sheet', r.lines.some(l => /Audit_Log row \d+: SAP 222/.test(l)) && r.lines.some(l => /Transcripts: NO ROW/.test(l)), r.lines.slice(0, 12).join('\n'));
  check('a long cell is cut short in the report', r.lines.every(l => l.length < 400), String(Math.max.apply(null, r.lines.map(l => l.length))));
  check('nothing in any sheet changed', JSON.stringify(Object.keys(e.sheets).map(k => e.sheets[k].grid)) === snap);
  check('clean data is reported as "none"', (() => { const c = makeEnv(); c.add('AUDIT_LOG_SHEET', log('A')); c.add('DASHBOARD_DATA_SHEET', dash('A')); const x = c.ctx.diagnoseDataIssues(); return !x.duplicates.length && !x.dashboardOnly.length && !x.sapMismatch.length && !x.nonNumericRepeat && !x.shiftedRows; })());
}

// ── repairs with a preview ──────────────────────────────────────────────────
const T0 = new Date('2026-09-16T13:00:45');
const L = (ref, o) => Object.assign({ 'Audit Ref': ref, 'Submitted At': T0, 'Interaction ID': 'i-' + ref, 'SAP ID': '111', 'Team Member': 'Agent ' + ref, 'Analysis Type': 'Repeats & Transfer Analyzer', 'Repeat Projection %': '40%', 'Email Status': 'Sent' }, o);
const D = (ref, o) => Object.assign({ 'Audit Ref': ref, 'Timestamp': T0, 'Interaction ID': 'i-' + ref, 'SAP ID': '111', 'Team Member': 'Agent ' + ref, 'Analysis Type': 'Repeats & Transfer Analyzer', 'Transcript Duration': '12:30', 'Repeat Projection %': '40%' }, o);
const TR = (ref, o) => Object.assign({ 'Audit Ref': ref, 'SAP ID': '111', 'Team Member': 'Agent ' + ref }, o);
const snapshot = e => JSON.stringify(Object.keys(e.sheets).map(k => e.sheets[k].grid));
function seeded() {
  const e = makeEnv();
  ['OK', 'SAPX', 'SAPL', 'SENT', 'NOPCT'].forEach(r => { e.add('AUDIT_LOG_SHEET', L(r)); e.add('DASHBOARD_DATA_SHEET', D(r)); e.add('TRANSCRIPTS_SHEET', TR(r)); e.add('ANALYSIS_SHEET', TR(r)); });
  return e;
}

console.log('\nTest 6 — an audit missing from Audit_Log is rebuilt from Dashboard_Data');
{
  const e = seeded();
  e.add('DASHBOARD_DATA_SHEET', D('HID', { 'SAP ID': '2004754', 'Team Member': 'Kanika', 'Issue Resolved': 'Yes', 'Email Status': 'Sent', 'Observer Name': 'QA One' }));
  const before = e.sheets[e.ctx.AUDIT_LOG_SHEET].grid.length;
  const p = e.ctx.previewDataRepairs();
  check('the preview finds one row to rebuild', p.restores === 1, JSON.stringify(p));
  check('the preview changes nothing', e.sheets[e.ctx.AUDIT_LOG_SHEET].grid.length === before);
  const r = e.ctx.applyDataRepairs();
  const al = e.sheets[e.ctx.AUDIT_LOG_SHEET], h = al.grid[0], row = al.grid[al.grid.length - 1];
  const at = n => row[h.indexOf(n)];
  check('one Audit_Log row was added', r.restored === 1 && al.grid.length === before + 1, JSON.stringify(r));
  check('it carries the audit ref, SAP ID, member and interaction', at('Audit Ref') === 'HID' && at('SAP ID') === '2004754' && at('Team Member') === 'Kanika' && at('Interaction ID') === 'i-HID');
  check('Submitted At comes from Timestamp and Duration from Transcript Duration', at('Submitted At') === T0 && at('Duration') === '12:30', JSON.stringify([at('Submitted At'), at('Duration')]));
  check('Observer, Issue Resolved and Email Status are carried over; Recipients is left blank', at('Observer Name') === 'QA One' && at('Issue Resolved') === 'Yes' && at('Email Status') === 'Sent' && at('Recipients') === '');
  check('the row has one value per Audit_Log column', row.length === h.length);
  check('running it again adds nothing', e.ctx.applyDataRepairs().restored === 0 && al.grid.length === before + 1);

  const e2 = seeded();
  e2.add('DASHBOARD_DATA_SHEET', D('NOSAP', { 'SAP ID': '' }));
  e2.add('DASHBOARD_DATA_SHEET', D('LONG', { 'Team Member': 'x'.repeat(120) }));
  e2.add('DASHBOARD_DATA_SHEET', D('TWICE')); e2.add('DASHBOARD_DATA_SHEET', D('TWICE'));
  const p2 = e2.ctx.previewDataRepairs();
  check('rows with missing or shifted data, or listed twice, are NOT rebuilt', p2.restores === 0 && p2.skipped.filter(s => /not rebuilt/.test(s)).length === 3, JSON.stringify(p2));
}

console.log('\nTest 7 — a SAP ID that only Dashboard_Data disagrees with is corrected');
{
  const e = seeded();
  const set = (sheetConst, ref, col, val) => { const sh = e.sheets[e.ctx[sheetConst]]; const i = sh.grid[0].indexOf('Audit Ref'); const r = sh.grid.find((x, k) => k > 0 && x[i] === ref); r[sh.grid[0].indexOf(col)] = val; };
  set('DASHBOARD_DATA_SHEET', 'SAPX', 'SAP ID', '999');                       // Dashboard_Data is the odd one out
  set('AUDIT_LOG_SHEET', 'SAPL', 'SAP ID', '888');                           // Audit_Log is the odd one out
  const p = e.ctx.previewDataRepairs();
  check('one SAP fix is planned and the Audit_Log-odd-one-out case is skipped', p.sapFixes === 1 && p.skipped.some(s => /SAPL/.test(s)), JSON.stringify(p));
  e.ctx.applyDataRepairs();
  const dd = e.sheets[e.ctx.DASHBOARD_DATA_SHEET], di = dd.grid[0].indexOf('Audit Ref'), ds = dd.grid[0].indexOf('SAP ID');
  const sap = ref => dd.grid.find((x, k) => k > 0 && x[di] === ref)[ds];
  check('Dashboard_Data now has the SAP ID the other three sheets agree on', sap('SAPX') === '111', sap('SAPX'));
  check('Audit_Log was not touched', e.sheets[e.ctx.AUDIT_LOG_SHEET].grid.find(x => x[0] === 'SAPX')[3] === '111');
  check('the case where Audit_Log is the odd one is left alone', sap('SAPL') === '111');

  const e3 = seeded();
  const put3 = (sc, ref, col, val) => { const sh = e3.sheets[e3.ctx[sc]]; const r = sh.grid.find((x, k) => k > 0 && x[0] === ref); r[sh.grid[0].indexOf(col)] = val; };
  put3('DASHBOARD_DATA_SHEET', 'SAPX', 'SAP ID', '999'); put3('ANALYSIS_SHEET', 'SAPX', 'SAP ID', '222');   // 111, 999, 111, 222: no 3-of-4
  const p3 = e3.ctx.previewDataRepairs();
  check('without three sheets agreeing nothing is changed, and it is reported', p3.sapFixes === 0 && p3.skipped.some(s => /SAPX/.test(s)), JSON.stringify(p3));

  const e2 = seeded();
  const dsh = e2.sheets[e2.ctx.DASHBOARD_DATA_SHEET]; dsh.grid.find((x, k) => k > 0 && x[0] === 'OK')[dsh.grid[0].indexOf('SAP ID')] = '999';
  e2.add('AUDIT_LOG_SHEET', L('OK')); e2.add('DASHBOARD_DATA_SHEET', D('OK'));      // the ref is now used twice
  check('a ref used twice is never touched', e2.ctx.previewDataRepairs().sapFixes === 0);
}

console.log('\nTest 8 — a sentence in Repeat Projection % becomes the number; the sentence is kept as a note');
{
  const e = seeded();
  const put = (sheetConst, ref, val) => { const sh = e.sheets[e.ctx[sheetConst]]; const r = sh.grid.find((x, k) => k > 0 && x[0] === ref); r[sh.grid[0].indexOf('Repeat Projection %')] = val; };
  const sentence = '95% repeat risk. No issue was identified, no resolution was reached';
  put('AUDIT_LOG_SHEET', 'OK', sentence); put('DASHBOARD_DATA_SHEET', 'OK', sentence);
  put('AUDIT_LOG_SHEET', 'SENT', 'High — approximately 80%. Customer was not helped');
  put('AUDIT_LOG_SHEET', 'NOPCT', 'about forty');
  const snap = snapshot(e);
  const p = e.ctx.previewDataRepairs();
  check('three sentences are planned (two sheets for OK, one for SENT)', p.repeatFixes === 3, JSON.stringify(p));
  check('the preview changes nothing', snapshot(e) === snap);
  const r = e.ctx.applyDataRepairs();
  const al = e.sheets[e.ctx.AUDIT_LOG_SHEET], rc = al.grid[0].indexOf('Repeat Projection %');
  const cell = ref => al.grid.find((x, k) => k > 0 && x[0] === ref)[rc];
  check('Audit_Log: the sentence is now 95 and the other reads 80', cell('OK') === 95 && cell('SENT') === 80, JSON.stringify([cell('OK'), cell('SENT')]));
  check('a clean value ("40%") is not touched', cell('SAPX') === '40%');
  check('text with no percentage is left exactly as it was', cell('NOPCT') === 'about forty');
  const rowOf = ref => al.grid.findIndex((x, k) => k > 0 && x[0] === ref) + 1;
  check('the original sentence is kept in a cell note', ((al.notes || {})[rowOf('OK') + ',' + (rc + 1)] || '').indexOf('No issue was identified') !== -1, JSON.stringify(al.notes));
  check('both sheets were fixed', r.repeatFixed === 3 && r.finished === true, JSON.stringify(r));
  check('running it again changes nothing', e.ctx.applyDataRepairs().repeatFixed === 0);
  check('the unreadable one is reported as left alone', e.ctx.previewDataRepairs().skipped.some(s => /NOPCT/.test(s)));
}

console.log('\nTest 10 — a messy cell is cleaned: participant lists and glued-on transcript header text');
{
  const e = seeded();
  const UUID = '2250f9a7-ff7c-442e-9d04-d4221d812634';
  // SAP 2006193: two clean audits say the agent is Nouamane Badrane; two messy rows have the transcript header glued on
  e.add('AUDIT_LOG_SHEET', L('C1', { 'SAP ID': '2006193', 'Team Member': 'Nouamane Badrane' }));
  e.add('AUDIT_LOG_SHEET', L('C2', { 'SAP ID': '2006193', 'Team Member': 'Nouamane Badrane' }));
  const glued = { 'SAP ID': '2006193', 'Team Member': 'Nouamane BadraneExternal Participant(s): HARVEY', 'Interaction ID': UUID + 'Transcript Start Time',
                  'Direction': 'InboundInternal Participant(s): Nouamane Badrane' };
  e.add('AUDIT_LOG_SHEET', L('G1', Object.assign({}, glued, { 'Duration': '10:10Direction: InboundInternal Participant(s)' })));
  e.add('DASHBOARD_DATA_SHEET', D('G1', Object.assign({}, glued, { 'Transcript Duration': '10:10Direction: InboundInternal Participant(s)' })));
  e.add('TRANSCRIPTS_SHEET', TR('G1', Object.assign({}, glued, { 'Transcript Duration': '10:10Direction: Inbound' })));
  // SAP 2008221: a participant list; a clean audit names the agent
  e.add('AUDIT_LOG_SHEET', L('C3', { 'SAP ID': '2008221', 'Team Member': 'Mark Enumerables' }));
  e.add('AUDIT_LOG_SHEET', L('LST', { 'SAP ID': '2008221', 'Team Member': 'Beatriz Adriana Yoc Pirir, Mark Enumerables, Omar baba, Johnnie Lou Ines, Someone Else Here' }));
  e.add('DASHBOARD_DATA_SHEET', D('LST', { 'SAP ID': '2008221', 'Team Member': 'Beatriz Adriana Yoc Pirir, Mark Enumerables, Omar baba, Johnnie Lou Ines, Someone Else Here' }));
  // SAP 5555555: a participant list and NO clean audit to say who the agent is
  e.add('AUDIT_LOG_SHEET', L('NOCLEAN', { 'SAP ID': '5555555', 'Team Member': 'Ana Maria Lopez Garcia, Pedro Pablo Ruiz Diaz, Carla Sofia Vega Mora, Juan Jose Perez Cruz' }));
  // SAP 6666666: clean audits say Zed Zebra, but he is NOT in this list, so the list must not be "fixed" to him
  e.add('AUDIT_LOG_SHEET', L('C4', { 'SAP ID': '6666666', 'Team Member': 'Zed Zebra' }));
  e.add('AUDIT_LOG_SHEET', L('WRONG', { 'SAP ID': '6666666', 'Team Member': 'Ana Maria Lopez Garcia, Pedro Pablo Ruiz Diaz, Carla Sofia Vega Mora, Juan Jose Perez Cruz' }));
  // two audits whose agent the admin chose by hand: no clean audit exists for their SAP IDs
  const list = 'Ali Guennoun, Luis Fernando Martinez Guevara, Someone Else Here, Another Person Name';
  e.add('AUDIT_LOG_SHEET', L('NHA-20260903-0003', { 'SAP ID': '2006749', 'Team Member': list }));
  e.add('DASHBOARD_DATA_SHEET', D('NHA-20260903-0003', { 'SAP ID': '2006749', 'Team Member': list }));
  e.add('AUDIT_LOG_SHEET', L('NHA-20260929-0024', { 'SAP ID': '4573347', 'Team Member': 'Oscar Rodrigo Hernandez Flores, Jonathan Cardoza, Krishna Singh Person, Another Person Name' }));
  e.add('DASHBOARD_DATA_SHEET', D('NHA-20260929-0024', { 'SAP ID': '4573347', 'Team Member': 'Oscar Rodrigo Hernandez Flores, Krishna Singh Person, Another Person Name, Yet Another Person' }));  // chosen name NOT in this cell
  e.add('AUDIT_LOG_SHEET', L('NHA-20260929-9999', { 'SAP ID': '4573348', 'Team Member': 'Oscar Rodrigo Hernandez Flores, Jonathan Cardoza, Krishna Singh Person, Another Person Name' }));  // same list, ref not chosen
  // a clean row that must not be touched
  e.add('AUDIT_LOG_SHEET', L('CLEANDUR', { 'Direction': 'Inbound', 'Duration': '1:09:13' }));
  // "Inbound / Outbound" is a real value many audits hold: it must never be shortened
  e.add('AUDIT_LOG_SHEET', L('BOTH', { 'Direction': 'Inbound / Outbound' }));
  e.add('DASHBOARD_DATA_SHEET', D('BOTH', { 'Direction': 'Inbound / Outbound' }));

  const snap = snapshot(e);
  const p = e.ctx.previewDataRepairs();
  check('the preview plans the cleaning', p.fieldFixes > 0, JSON.stringify(p));
  check('the preview changes nothing', snapshot(e) === snap);
  check('the participant list with no clean audit is skipped and listed', p.skipped.some(s => /NOCLEAN/.test(s) && /several names/.test(s)), JSON.stringify(p.skipped));
  const r = e.ctx.applyDataRepairs();
  const cellOf = (sc, ref, col) => { const sh = e.sheets[e.ctx[sc]], h = sh.grid[0]; const row = sh.grid.find((x, k) => k > 0 && x[h.indexOf('Audit Ref')] === ref); return row[h.indexOf(col)]; };
  check('Audit_Log: Team Member, Interaction ID and Direction are clean', cellOf('AUDIT_LOG_SHEET', 'G1', 'Team Member') === 'Nouamane Badrane' && cellOf('AUDIT_LOG_SHEET', 'G1', 'Interaction ID') === UUID && cellOf('AUDIT_LOG_SHEET', 'G1', 'Direction') === 'Inbound',
        JSON.stringify([cellOf('AUDIT_LOG_SHEET', 'G1', 'Team Member'), cellOf('AUDIT_LOG_SHEET', 'G1', 'Interaction ID'), cellOf('AUDIT_LOG_SHEET', 'G1', 'Direction')]));
  check('Audit_Log: the duration is 10:10', cellOf('AUDIT_LOG_SHEET', 'G1', 'Duration') === '10:10', cellOf('AUDIT_LOG_SHEET', 'G1', 'Duration'));
  check('Dashboard_Data and Transcripts are cleaned the same way', cellOf('DASHBOARD_DATA_SHEET', 'G1', 'Team Member') === 'Nouamane Badrane' && cellOf('DASHBOARD_DATA_SHEET', 'G1', 'Transcript Duration') === '10:10' && cellOf('TRANSCRIPTS_SHEET', 'G1', 'Interaction ID') === UUID);
  const alh = e.sheets[e.ctx.AUDIT_LOG_SHEET];
  const rowG1 = alh.grid.findIndex((x, k) => k > 0 && x[0] === 'G1') + 1;
  check('the duration cell is stored as text so Sheets does not turn 10:10 into a time', ((alh.formats || {})[rowG1 + ',' + (alh.grid[0].indexOf('Duration') + 1)]) === '@', JSON.stringify(alh.formats));
  check('the original messy text is kept in a cell note', ((alh.notes || {})[rowG1 + ',' + (alh.grid[0].indexOf('Team Member') + 1)] || '').indexOf('External Participant(s)') !== -1);
  check('a participant list becomes the agent the same SAP ID has on clean audits', cellOf('AUDIT_LOG_SHEET', 'LST', 'Team Member') === 'Mark Enumerables' && cellOf('DASHBOARD_DATA_SHEET', 'LST', 'Team Member') === 'Mark Enumerables');
  check('the list with no clean audit is left exactly as it was', cellOf('AUDIT_LOG_SHEET', 'NOCLEAN', 'Team Member').indexOf('Ana Maria Lopez Garcia, Pedro') === 0);
  check('a clean name that is not inside the list is never used', cellOf('AUDIT_LOG_SHEET', 'WRONG', 'Team Member').indexOf('Ana Maria Lopez Garcia') === 0, cellOf('AUDIT_LOG_SHEET', 'WRONG', 'Team Member'));
  check('a clean Direction and Duration are not touched', cellOf('AUDIT_LOG_SHEET', 'CLEANDUR', 'Direction') === 'Inbound' && cellOf('AUDIT_LOG_SHEET', 'CLEANDUR', 'Duration') === '1:09:13');
  check('the two agents the admin chose are applied (Ali Guennoun, Jonathan Cardoza)', cellOf('AUDIT_LOG_SHEET', 'NHA-20260903-0003', 'Team Member') === 'Ali Guennoun' && cellOf('DASHBOARD_DATA_SHEET', 'NHA-20260903-0003', 'Team Member') === 'Ali Guennoun' && cellOf('AUDIT_LOG_SHEET', 'NHA-20260929-0024', 'Team Member') === 'Jonathan Cardoza');
  check('a chosen name that is NOT inside that cell is never written', cellOf('DASHBOARD_DATA_SHEET', 'NHA-20260929-0024', 'Team Member').indexOf('Oscar Rodrigo') === 0, cellOf('DASHBOARD_DATA_SHEET', 'NHA-20260929-0024', 'Team Member'));
  check('the same list on an audit that was not chosen is still left alone', cellOf('AUDIT_LOG_SHEET', 'NHA-20260929-9999', 'Team Member').indexOf('Oscar Rodrigo') === 0);
  check('"Inbound / Outbound" is left exactly as it is, in both sheets', cellOf('AUDIT_LOG_SHEET', 'BOTH', 'Direction') === 'Inbound / Outbound' && cellOf('DASHBOARD_DATA_SHEET', 'BOTH', 'Direction') === 'Inbound / Outbound');
  check('the preview does not list "Inbound / Outbound" as a cell to clean', !p.skipped.concat([]).some(s => /BOTH/.test(s)));
  check('the result counts the cleaned cells', r.fieldFixed > 0 && r.finished === true, JSON.stringify(r));
  check('running it again changes nothing', e.ctx.applyDataRepairs().fieldFixed === 0);
}

console.log('\nTest 11 — a call length that Sheets turned into a time of day goes back to text');
{
  const dayZero = (h, m, s) => new Date(Date.UTC(1899, 11, 30, h, m, s));
  const e = makeEnv();
  check('2:15:05 on the sheet\'s day zero reads "2:15:05"; 0:48:15 and 23:59:59 too', e.ctx._durationFromDate_(dayZero(2, 15, 5), 'UTC') === '2:15:05' && e.ctx._durationFromDate_(dayZero(0, 48, 15), 'UTC') === '0:48:15' && e.ctx._durationFromDate_(dayZero(23, 59, 59), 'UTC') === '23:59:59');
  check('anything that is not a time on day zero gives "" (a 2026 date, a day later, not a date, an invalid date)', e.ctx._durationFromDate_(new Date(Date.UTC(2026, 9, 1, 2, 15, 5)), 'UTC') === '' && e.ctx._durationFromDate_(new Date(Date.UTC(1900, 0, 1, 0, 15, 0)), 'UTC') === '' && e.ctx._durationFromDate_('2:15:05', 'UTC') === '' && e.ctx._durationFromDate_(new Date('x'), 'UTC') === '');

  e.add('AUDIT_LOG_SHEET', L('D1', { 'Duration': dayZero(2, 15, 5) }));
  e.add('DASHBOARD_DATA_SHEET', D('D1', { 'Transcript Duration': dayZero(2, 15, 5) }));
  e.add('TRANSCRIPTS_SHEET', TR('D1', { 'Transcript Duration': dayZero(2, 15, 5) }));
  e.add('AUDIT_LOG_SHEET', L('D2', { 'Duration': '1:09:13' }));
  e.add('AUDIT_LOG_SHEET', L('D3', { 'Duration': new Date(Date.UTC(1900, 0, 1, 0, 15, 0)) }));
  const snap = snapshot(e);
  const p = e.ctx.previewDataRepairs();
  check('the preview plans the three cells of D1 and skips D3 (not within a day), listing it', p.fieldFixes === 3 && p.skipped.some(s => /D3/.test(s) && /not a call length within one day/.test(s)), JSON.stringify(p));
  check('the preview changes nothing', snapshot(e) === snap);
  const r = e.ctx.applyDataRepairs();
  const cellOf = (sc, ref, col) => { const sh = e.sheets[e.ctx[sc]], h = sh.grid[0]; const row = sh.grid.find((x, k) => k > 0 && x[h.indexOf('Audit Ref')] === ref); return row[h.indexOf(col)]; };
  check('all three become the text 2:15:05', cellOf('AUDIT_LOG_SHEET', 'D1', 'Duration') === '2:15:05' && cellOf('DASHBOARD_DATA_SHEET', 'D1', 'Transcript Duration') === '2:15:05' && cellOf('TRANSCRIPTS_SHEET', 'D1', 'Transcript Duration') === '2:15:05' && r.fieldFixed === 3, JSON.stringify(r));
  const al = e.sheets[e.ctx.AUDIT_LOG_SHEET], rowD1 = al.grid.findIndex((x, k) => k > 0 && x[0] === 'D1') + 1, c = al.grid[0].indexOf('Duration') + 1;
  check('the cell is stored as text so Sheets does not turn it into a time again, and the original is kept in a note', (al.formats || {})[rowD1 + ',' + c] === '@' && /1899/.test((al.notes || {})[rowD1 + ',' + c] || ''));
  check('a normal duration and the not-a-day value are left exactly as they were', cellOf('AUDIT_LOG_SHEET', 'D2', 'Duration') === '1:09:13' && Object.prototype.toString.call(cellOf('AUDIT_LOG_SHEET', 'D3', 'Duration')) === '[object Date]');
  check('running it again changes nothing', e.ctx.applyDataRepairs().fieldFixed === 0);
}

console.log('\nTest 11b — a "11:39" that Sheets read as hours and minutes comes back as the text the transcript says');
{
  const e = makeEnv();
  const c = e.ctx;
  const at = (day, h, m, s) => new Date(Date.UTC(1899, 11, 30 + day, h, m, s));
  check('the call length is read from the transcript header (MM:SS or H:MM:SS), and from nothing else', c._durationFromTranscript_('Interaction ID: X\nTranscript Duration: 11:39\nAgent: hi') === '11:39' && c._durationFromTranscript_('Transcript Duration: 0:11:39\r\n') === '0:11:39' && c._durationFromTranscript_('Transcript Duration: about ten minutes') === '' && c._durationFromTranscript_('no header') === '' && c._durationFromTranscript_(null) === '');
  check('11:39 typed as MM:SS and read as 11 h 39 min is written back as 11:39, NOT as 11:39:00 (which would say 11 hours)', c._durationDecision_(at(0, 11, 39, 0), 'UTC', '11:39').to === '11:39' && c._durationDecision_(at(0, 11, 39, 0), 'UTC', '11:39').how === 'transcript');
  check('31:53 became a day and 7 h 53 min (1899-12-31): it is written back as 31:53', c._durationDecision_(at(1, 7, 53, 0), 'UTC', '31:53').to === '31:53');
  check('49:18 became 1900-01-01 01:18: written back as 49:18', c._durationDecision_(at(2, 1, 18, 0), 'UTC', '49:18').to === '49:18');
  check('a three-part length in the transcript is written as it is', c._durationDecision_(at(0, 0, 11, 39), 'UTC', '0:11:39').to === '0:11:39');
  check('a Date with no seconds and no length in the transcript is NOT guessed', c._durationDecision_(at(0, 11, 39, 0), 'UTC', '').to === '' && c._durationDecision_(at(1, 7, 53, 0), 'UTC', '').to === '');
  check('a length the analyst typed over by hand (the transcript says something else) is not overwritten', c._durationDecision_(at(0, 11, 39, 0), 'UTC', '12:00').to === '' && /does not match/.test(c._durationDecision_(at(0, 11, 39, 0), 'UTC', '12:00').reason));
  check('a Date WITH seconds on day zero cannot be a mistyped MM:SS, so it is written as H:mm:ss even without a transcript', c._durationDecision_(at(0, 2, 15, 5), 'UTC', '').to === '2:15:05' && c._durationDecision_(at(0, 2, 15, 5), 'UTC', '').how === 'clock');
  check('an invalid date gives nothing', c._durationDecision_(new Date('x'), 'UTC', '1:00').to === '');

  // end to end: three sheets, the transcript row says what the length was
  const tx = len => (len ? 'Interaction ID: I1\nTranscript Duration: ' + len + '\nAgent: hello' : 'Agent: hello (no header)');
  e.add('AUDIT_LOG_SHEET', L('M1', { 'Duration': at(0, 11, 39, 0) }));
  e.add('DASHBOARD_DATA_SHEET', D('M1', { 'Transcript Duration': at(0, 11, 39, 0) }));
  e.add('TRANSCRIPTS_SHEET', TR('M1', { 'Transcript Duration': at(0, 11, 39, 0), 'Transcript': tx('11:39') }));
  e.add('AUDIT_LOG_SHEET', L('M2', { 'Duration': at(1, 7, 53, 0) }));
  e.add('TRANSCRIPTS_SHEET', TR('M2', { 'Transcript Duration': at(1, 7, 53, 0), 'Transcript': tx('31:53') }));
  e.add('AUDIT_LOG_SHEET', L('M3', { 'Duration': at(0, 4, 57, 0) }));                         // no transcript length: left alone
  e.add('TRANSCRIPTS_SHEET', TR('M3', { 'Transcript Duration': at(0, 4, 57, 0), 'Transcript': tx('') }));
  e.add('AUDIT_LOG_SHEET', L('M4', { 'Duration': at(0, 11, 39, 0) }));                        // typed over by the analyst
  e.add('TRANSCRIPTS_SHEET', TR('M4', { 'Transcript Duration': at(0, 11, 39, 0), 'Transcript': tx('20:00') }));
  const p = e.ctx.previewDataRepairs();
  check('the preview plans M1 (3 cells) and M2 (2 cells) and skips M3 and M4 with the reason', p.fieldFixes === 5 && p.skipped.some(s => /M3/.test(s) && /no seconds/.test(s)) && p.skipped.some(s => /M4/.test(s) && /does not match/.test(s)), JSON.stringify(p));
  const dg = e.ctx.diagnoseDashboardGaps();
  check('the diagnosis agrees with the preview: 5 cells can be written back, all from the transcript; 4 are left (2 with no length in the transcript, 2 that the transcript contradicts)', dg.durationFromTranscript === 5 && dg.durationNotDay === 4 && dg.durationLeftNoTruth === 2 && dg.durationLeftMismatch === 2 && dg.durationLeftStale === 2 && dg.lines.some(l => /5 from the length written in the saved transcript/.test(l)), JSON.stringify([dg.durationFromTranscript, dg.durationNotDay, dg.durationLeftNoTruth, dg.durationLeftMismatch]));
  check('a cell that holds the length of a DIFFERENT audit (M4 holds 11:39, which is M1) is counted as left over from an earlier paste, and the line says whose length it is', dg.lines.some(l => /Audit_Log M4: .*(transcript says 20:00; the cell equals the length of M1)/.test(l)) && dg.lines.some(l => /for 2 of those the cell equals the length of a DIFFERENT audit/.test(l)), dg.lines.filter(l => /M4|DIFFERENT/.test(l)).join(' | '));
  e.ctx.applyDataRepairs();
  const cellOf = (sc, ref, col) => { const sh = e.sheets[e.ctx[sc]], h = sh.grid[0]; const row = sh.grid.find((x, k) => k > 0 && x[h.indexOf('Audit Ref')] === ref); return row[h.indexOf(col)]; };
  check('M1 is 11:39 in all three sheets (the original Date stays in a note)', cellOf('AUDIT_LOG_SHEET', 'M1', 'Duration') === '11:39' && cellOf('DASHBOARD_DATA_SHEET', 'M1', 'Transcript Duration') === '11:39' && cellOf('TRANSCRIPTS_SHEET', 'M1', 'Transcript Duration') === '11:39');
  check('M2 is 31:53', cellOf('AUDIT_LOG_SHEET', 'M2', 'Duration') === '31:53' && cellOf('TRANSCRIPTS_SHEET', 'M2', 'Transcript Duration') === '31:53');
  check('M3 and M4 are still the Date they were', Object.prototype.toString.call(cellOf('AUDIT_LOG_SHEET', 'M3', 'Duration')) === '[object Date]' && Object.prototype.toString.call(cellOf('AUDIT_LOG_SHEET', 'M4', 'Duration')) === '[object Date]');
}

console.log('\nTest 12 — diagnoseDashboardGaps: Duration, Not Sent and Issue Resolved, read-only');
{
  const dayZero = (h, m, s) => new Date(Date.UTC(1899, 11, 30, h, m, s));
  const e = makeEnv();
  const at = (y, mo, d) => new Date(Date.UTC(y, mo - 1, d, 10, 0, 0));
  e.add('AUDIT_LOG_SHEET', L('G1', { 'Submitted At': at(2026, 9, 5), 'Email Status': 'Not Sent', 'Issue Resolved': 'Yes', 'Duration': dayZero(1, 0, 5), 'Observer Name': 'QA One' }));
  e.add('AUDIT_LOG_SHEET', L('G2', { 'Submitted At': at(2026, 10, 2), 'Email Status': 'Not Sent', 'Issue Resolved': '', 'Observer Name': 'QA One' }));
  e.add('AUDIT_LOG_SHEET', L('G3', { 'Submitted At': at(2026, 10, 3), 'Email Status': 'Not Sent', 'Recipients': 'a@telus.com', 'Issue Resolved': '<div style="font-size:13px;">Manager coaching tip', 'Observer Name': 'QA Two' }));
  e.add('AUDIT_LOG_SHEET', L('G4', { 'Email Status': 'Sent', 'Issue Resolved': 'N/A' }));
  e.add('AUDIT_LOG_SHEET', L('G5', { 'Email Status': 'Sent', 'Issue Resolved': 'Agent confirmed form was submitted to trigger a physical inventory' }));
  e.add('AUDIT_LOG_SHEET', L('G6', { 'Email Status': 'Sent', 'Issue Resolved': 'Agent confirmed form was submitted. Resolution is incomplete' }));
  e.add('AUDIT_LOG_SHEET', L('G7', { 'Email Status': 'Sent', 'Issue Resolved': 'Resolution is ambiguous. Interaction summary states phones were returned' }));
  // a Duration cell holding a Date that is not on day zero (a short time read the wrong way by the time zone, or a real date)
  e.add('AUDIT_LOG_SHEET', L('G8', { 'Email Status': 'Sent', 'Issue Resolved': 'Yes', 'Duration': new Date(Date.UTC(1899, 11, 29, 23, 58, 0)) }));
  let a = '';
  e.who.email = 'some.analyst@telus.com';
  try { e.ctx.diagnoseDashboardGaps(); } catch (x) { a = x.message; }
  check('an analyst is refused', /admin/i.test(a), a);
  e.who.email = 'danzen.palisoc@telus.com';
  const snap = snapshot(e);
  const r = e.ctx.diagnoseDashboardGaps();
  check('it changes nothing', snapshot(e) === snap);
  check('Duration: the three-part cell in Audit_Log is found, with its conversion shown', r.durationAudit === 1 && r.lines.some(l => /Audit_Log G1: .* -> 1:00:05 \(three-part length\)/.test(l)), r.lines.slice(0, 4).join(' | '));
  check('Duration: a Date that is not on day zero is counted apart, with its sheet, its day and a sample so the cause can be seen', r.durationNotDay === 1 && r.lines.some(l => /left alone .*: 1 \(Audit_Log 1, Dashboard_Data 0, Transcripts 0\)/.test(l)) && r.lines.some(l => /by the day the value falls on/.test(l) && /1899-12-29 1/.test(l)) && r.lines.some(l => /Audit_Log G8: 1899-12-29/.test(l)), r.lines.filter(l => /not within|by sheet|falls on|G8/.test(l)).join(' | '));
  check('Not Sent: 3 audits, 1 of them with recipients written, by month', r.notSent === 3 && r.notSentWithRecipients === 1 && r.notSentByMonth['2026-09'] === 1 && r.notSentByMonth['2026-10'] === 2, JSON.stringify([r.notSent, r.notSentWithRecipients, r.notSentByMonth]));
  check('Not Sent is also split by observer', r.lines.some(l => /by observer/.test(l) && /QA One 2/.test(l) && /QA Two 1/.test(l)));
  check('Issue Resolved: 1 blank, 1 clear verdict, 1 html-like, 1 very short, 3 sentences', r.issueBlank === 1 + 0 + 0 && r.issueVerdict === 2 && r.issueHtml === 1 && r.issueShort === 1 && r.issueSentence === 3, JSON.stringify([r.issueBlank, r.issueVerdict, r.issueHtml, r.issueShort, r.issueSentence]));
  check('it shows how the sentences usually start (the pattern to teach the dashboard)', r.lines.some(l => /how the sentences start/.test(l) && /"agent confirmed form" 2/.test(l)), r.lines.filter(l => /start/.test(l)).join(' | '));
  check('every report line is short enough to read in the log', r.lines.every(l => l.length < 700));
}

console.log('\nTest 9 — only admins');
{
  const e = seeded(); e.who.email = 'some.analyst@telus.com';
  let a = '', b = '';
  try { e.ctx.previewDataRepairs(); } catch (x) { a = x.message; }
  try { e.ctx.applyDataRepairs(); } catch (x) { b = x.message; }
  check('previewDataRepairs refuses an analyst', /admin/i.test(a), a);
  check('applyDataRepairs refuses an analyst', /admin/i.test(b), b);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
