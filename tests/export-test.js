/**
 * The dashboard export (ExportService.gs + the Export buttons in index.html).
 *
 * The old exportToSpreadsheet took the ROWS from the browser, wrote them to a spreadsheet in the app owner's Drive and shared it
 * with the whole domain. Anyone could make the app publish any text under the owner's name, a cell starting with = became a
 * formula, and there was no limit. The new exportAudits takes only a view name and Audit Refs and builds everything itself.
 *
 *  1. What the page may send.  2. The rows equal the dashboard's rows.  3. The right rows for each view.
 *  4. Cells are text, never formulas.  5. The flow: folder, viewer-only sharing, limits, failures.  6. The page buttons.
 *
 * Usage: node tests/export-test.js
 */
const vm = require('vm');
const fs = require('fs');
const path = require('path');
const { load } = require('./lib/page-harness');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

function makeEnv() {
  const logs = [], created = [], drive = { folders: [], viewers: [], editors: [], sharing: [], removed: [], createdFolders: 0 };
  const who = { email: 'analyst.one@telus.com' };
  const cacheStore = {};
  const lockState = { busy: false, released: 0 };
  const api = { fail: null };
  const makeSheet = () => {
    const s = { writes: [], frozen: 0 };
    s.getRange = (r, c, n, m) => {
      const chain = { setValues(v) { s.writes.push({ r, c, n, m, values: v }); return chain; }, setFontWeight() { return chain; }, setBackground() { return chain; }, setFontColor() { return chain; } };
      return chain;
    };
    s.setFrozenRows = n => { s.frozen = n; };
    s.autoResizeColumns = () => {};
    return s;
  };
  const sandbox = {
    Logger: { log: m => logs.push(String(m)) }, JSON, Array, Object, String, Number, Math, Date, RegExp, isFinite, isNaN, Error, parseInt,
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Session: { getActiveUser: () => ({ getEmail: () => who.email }), getScriptTimeZone: () => 'America/New_York' },
    LockService: { getScriptLock: () => ({ tryLock: () => !lockState.busy, releaseLock() { lockState.released++; } }) },
    CacheService: { getScriptCache: () => ({ get: k => (k in cacheStore ? cacheStore[k] : null), put: (k, v) => { cacheStore[k] = v; } }) },
    Utilities: { base64Encode: s => Buffer.from(String(s)).toString('base64'), formatDate: () => '2026-10-13 12:00' },
    SpreadsheetApp: { create: name => {
      if (api.fail) throw new Error(api.fail);
      const sheet = makeSheet(), id = 'SS' + (created.length + 1);
      const ss = { id, name, sheet, getId: () => id, getUrl: () => 'https://docs.google.com/spreadsheets/d/' + id + '/edit', getName: () => name, getActiveSheet: () => sheet };
      created.push(ss); return ss;
    } },
    DriveApp: {
      getFoldersByName: n => ({ hasNext: () => drive.folders.length > 0 && drive.folders[0].name === n, next: () => drive.folders[0] }),
      createFolder: n => { drive.createdFolders++; const f = { name: n, addFile: file => (f.files = (f.files || []).concat(file.id)) }; drive.folders.push(f); return f; },
      getFileById: id => ({ id, addViewer: e => drive.viewers.push({ id, email: e }), addEditor: e => drive.editors.push({ id, email: e }), setSharing: (a, p) => drive.sharing.push([a, p]) }),
      getRootFolder: () => ({ removeFile: f => drive.removed.push(f.id) }),
    },
  };
  const ctx = vm.createContext(sandbox);
  ['Config.gs', 'ExportService.gs'].forEach(f => vm.runInContext(read(f), ctx, { filename: f }));
  vm.runInContext(read('SheetService.gs').match(/^var _DURATION_SHAPE[^\n]*/m)[0], ctx);
  vm.runInContext(extractFunction(read('Code.gs'), 'isValidAuditRef_'), ctx);
  const data = { log: [], dash: [] };
  ctx.readAuditLog_ = () => data.log;
  ctx.readDashboardSheet_ = () => data.dash;
  return { ctx, logs, created, drive, who, cacheStore, lockState, api, data };
}

let n = 0;
function add(env, o, dash) {
  const ref = (o && o.ref) || 'NHA-2026101' + (++n);
  env.data.log.push(Object.assign({ 'Audit Ref': ref, 'Submitted At': 'Mon Oct 12 2026 10:00:00 GMT-0400', 'Interaction ID': 'i-' + ref, 'SAP ID': '111', 'Team Member': 'Agent A', 'Team Leader': 'TL One',
    'Analysis Type': 'Repeats & Transfer Analyzer', 'Duration': '12:30', 'Repeat Projection %': '40%', 'Issue Resolved': 'Yes', 'Email Status': 'Sent' }, o || {}));
  env.data.dash.push(Object.assign({ 'Audit Ref': ref, 'Call Summary': 'Summary of ' + ref, 'Critical Flags': 'F1 | F2', 'Call Driver': 'Billing', 'Sales Attempted': '', 'SMART S (Specific)': 'S' }, dash || {}));
  return ref;
}
delete add.ref;
const refsOf = env => env.data.log.map(r => r['Audit Ref']);

console.log('\nTest 1 — the page may send only a view name and Audit Refs');
{
  const e = makeEnv();
  const V = (view, refs) => e.ctx._exValidate_(view, refs);
  check('only the three views are known', V('repeats', ['A1']).ok && V('sales', ['A1']).ok && V('auditlog', ['A1']).ok && !V('everything', ['A1']).ok && !V('', ['A1']).ok && !V(undefined, ['A1']).ok);
  check('names that exist on every JavaScript object are not views', !V('__proto__', ['A1']).ok && !V('constructor', ['A1']).ok && !V('toString', ['A1']).ok && !V('hasOwnProperty', ['A1']).ok);
  check('not an array, an empty list, or the old rows-from-the-browser shape is refused', !V('repeats', 'A1').ok && !V('repeats', []).ok && !V('repeats', [{ 'Audit Ref': 'A1', 'Call Summary': 'text' }]).ok);
  check('a ref with a quote, space or slash is refused', !V('repeats', ['A1"']).ok && !V('repeats', ['A 1']).ok && !V('repeats', ['../x']).ok);
  const ok = V('repeats', ['NHA-1', ' NHA-1 ', 'NHA-2']);
  check('good refs are trimmed and de-duplicated', ok.ok && ok.refs.join() === 'NHA-1,NHA-2');
  const many = []; for (let i = 0; i < 5001; i++) many.push('NHA-' + i);
  check('more than 5,000 refs are refused, 5,000 are fine', !V('repeats', many).ok && V('repeats', many.slice(0, 5000)).ok);
  const r = e.ctx.exportAudits('repeats', [{ x: 1 }], 'x');
  check('exportAudits refuses bad input and creates nothing', r.success === false && e.created.length === 0, JSON.stringify(r));
}

console.log('\nTest 2 — the rows equal the dashboard\'s rows (parity with the page)');
{
  const e = makeEnv();
  add(e, { ref: 'BOTH' }, { 'Issue Resolved': 'No', 'Team Leader': 'Other TL', 'Recommendations': 'old name', 'Agent Email': 'a@telus.com', 'Department': 'Dept' });
  add(e, { ref: 'NOLOGVALUES', 'Team Member': '', 'Team Leader': '' }, { 'Team Member': 'From Dashboard', 'Team Leader': 'Dash TL', 'Top Opportunity': 'old top' });
  e.data.log.push({ 'Audit Ref': 'LOGONLY', 'Submitted At': 'Tue Oct 13 2026 09:00:00 GMT-0400', 'Team Member': 'Agent Z', 'SAP ID': '9', 'Analysis Type': 'Sales Analyzer', 'Duration': '5:00' });
  e.data.dash.push({ 'Audit Ref': 'DASHONLY', 'Timestamp': 'Mon Oct 05 2026 10:00:00 GMT-0400', 'Team Member': 'Agent C', 'SAP ID': '3', 'Analysis Type': 'Repeats & Transfer Analyzer', 'Critical Flags': 'F9' });
  e.data.dash.push({ 'Audit Ref': 'LONGNAME', 'Team Member': 'x'.repeat(120), 'Critical Flags': 'F8' });
  e.data.log.push({ 'Audit Ref': 'LONGNAME', 'Submitted At': 'Mon Oct 05 2026 11:00:00 GMT-0400', 'Team Member': '', 'SAP ID': '4', 'Analysis Type': 'Repeats & Transfer Analyzer' });
  const mine = e.ctx.buildDashboardRows_(e.data.log, e.data.dash);

  const h = load();
  h.ctx.dbFinishRender = function () {};
  h.ctx.dbHandleData({ rows: e.data.dash, ts: new Date().toISOString() }, false);
  h.pending.find(p => p.method === 'getAuditLogData').success(e.data.log);
  const page = h.ctx._dbAllData;
  const diffs = [];
  page.forEach(p => {
    const m = mine.find(x => x['Audit Ref'] === p['Audit Ref']);
    if (!m) { diffs.push('missing ' + p['Audit Ref']); return; }
    const keys = {}; Object.keys(p).concat(Object.keys(m)).forEach(k => { if (k.charAt(0) !== '_') keys[k] = true; });
    Object.keys(keys).forEach(k => { if (String(m[k] === undefined ? '' : m[k]) !== String(p[k] === undefined ? '' : p[k])) diffs.push(p['Audit Ref'] + '.' + k + ': "' + m[k] + '" vs "' + p[k] + '"'); });
  });
  check('every field of every merged row equals the page (' + page.length + ' audits, all columns)', diffs.length === 0 && mine.length === page.length, diffs.slice(0, 4).join(' | ') + ' counts ' + mine.length + '/' + page.length);
  check('an audit only in Dashboard_Data is kept and a paragraph is not used as a name', mine.some(r => r['Audit Ref'] === 'DASHONLY') && mine.find(r => r['Audit Ref'] === 'LONGNAME')['Team Member'] === '');
}

console.log('\nTest 3 — the right rows for each view');
{
  const e = makeEnv();
  const r1 = add(e, {}, {}); const r2 = add(e, {}, {}); const s1 = add(e, { 'Analysis Type': 'Sales Analyzer' }, { 'Sales Attempted': 'Yes' }); add(e, {}, {});
  const refSet = x => { const o = {}; x.forEach(r => { o[r] = true; }); return o; };
  const rep = e.ctx._exRowsFor_('repeats', refSet([r1, s1]), e.data.log, e.data.dash);
  check('Repeats: only Repeats audits among the selected refs (not the Sales one, not the unselected ones)', rep.rows.length === 1 && rep.rows[0]['Audit Ref'] === r1);
  const sal = e.ctx._exRowsFor_('sales', refSet([r1, s1]), e.data.log, e.data.dash);
  check('Sales: only the Sales audit', sal.rows.length === 1 && sal.rows[0]['Audit Ref'] === s1 && sal.rows[0]['Sales Attempted'] === 'Yes');
  const al = e.ctx._exRowsFor_('auditlog', refSet([r1, r2, s1]), e.data.log, []);
  check('Audit Log: the selected Audit_Log rows exactly as they are (the log columns, no dashboard columns)', al.rows.length === 3 && al.headers.indexOf('Submitted At') !== -1 && al.headers.indexOf('Call Summary') === -1);
  check('the headers cover every column of every row, in first-seen order', rep.headers.indexOf('Audit Ref') === 0 && rep.headers.indexOf('Call Summary') !== -1 && rep.headers.indexOf('SMART S (Specific)') !== -1);
  check('a refs list with nothing in the view gives no rows', e.ctx._exRowsFor_('sales', refSet([r1]), e.data.log, e.data.dash).rows.length === 0);
}

console.log('\nTest 4 — cells are text, never formulas');
{
  const e = makeEnv();
  const c = v => e.ctx._exCell_(v);
  check('= + - @ at the start are protected with an apostrophe', c('=HYPERLINK("http://evil","click")') === "'=HYPERLINK(\"http://evil\",\"click\")" && c('+1') === "'+1" && c('-2') === "'-2" && c('@SUM(A1)') === "'@SUM(A1)");
  check('a tab or carriage return at the start is protected too', c('\tcmd') === "'\tcmd" && c('\rcmd') === "'\rcmd");
  check('a time-like text stays text (10:10, 1:09:13)', c('10:10') === "'10:10" && c('1:09:13') === "'1:09:13");
  check('ordinary text, numbers and blanks are unchanged', c('Warm greeting') === 'Warm greeting' && c(65) === '65' && c(null) === '' && c(undefined) === '' && c('A=B') === 'A=B');
  check('a text with a formula inside (not at the start) is left alone', c('see =A1') === 'see =A1');
  const long = c('x'.repeat(60000));
  check('an over-long text is cut to fit a Sheets cell', long.length === 49000);
  check('a formula-looking over-long text is cut and still protected', c('=' + 'x'.repeat(60000)).length === 49001 && c('=' + 'x'.repeat(60000)).charAt(0) === "'");
}

console.log('\nTest 5 — the flow: folder, viewer-only sharing, limits, failures');
{
  const e = makeEnv();
  add(e, {}, { 'Call Summary': '=1+1' }); add(e, {}, {});
  const refs = refsOf(e);
  const r = e.ctx.exportAudits('repeats', refs, '  My export\nwith break ');
  check('an analyst gets a spreadsheet', r.success === true && r.rowCount === 2 && /docs\.google\.com\/spreadsheets\/d\/SS1/.test(r.url), JSON.stringify(r));
  const ss = e.created[0];
  check('the title is cleaned and has the time', /^My export with break - 2026-10-13 12:00$/.test(ss.name), ss.name);
  const w = ss.sheet.writes;
  check('the header row, then the data rows, are written', w.length === 2 && w[0].values[0].indexOf('Audit Ref') === 0 && w[1].values.length === 2 && w[1].r === 2 && ss.sheet.frozen === 1);
  check('a formula in the data is written as text', w[1].values.some(row => row.indexOf("'=1+1") !== -1));
  check('the file is in the "NH Analyzer Exports" folder and out of the root', e.drive.folders[0].name === 'NH Analyzer Exports' && e.drive.folders[0].files[0] === 'SS1' && e.drive.removed[0] === 'SS1');
  check('it is shared, as VIEWER, with the person who asked and nobody else', e.drive.viewers.length === 1 && e.drive.viewers[0].email === 'analyst.one@telus.com' && e.drive.editors.length === 0);
  check('it is NEVER shared with the domain or the public', e.drive.sharing.length === 0);
  check('the lock is released', e.lockState.released === 1);
  e.ctx.exportAudits('repeats', refs, '');
  check('the folder is reused', e.drive.createdFolders === 1 && e.created[1].name.indexOf('NH Audit Export - Repeats') === 0);

  // rate limit: 10 per person per hour (2 so far)
  for (let i = 0; i < 8; i++) e.ctx.exportAudits('repeats', refs, '');
  const before = e.created.length;
  const eleventh = e.ctx.exportAudits('repeats', refs, '');
  check('the 11th export in an hour is refused with a clear message and builds nothing', eleventh.success === false && /10 exports/.test(eleventh.error) && e.created.length === before, JSON.stringify(eleventh));
  e.who.email = 'someone.else@telus.com';
  check('the limit is per person', e.ctx.exportAudits('repeats', refs, '').success === true);

  const b = makeEnv(); add(b, {}, {}); b.lockState.busy = true;
  const busy = b.ctx.exportAudits('repeats', refsOf(b), '');
  check('while another export is being made: try again in a minute', busy.success === false && /another export/i.test(busy.error) && b.created.length === 0);
  const nid = makeEnv(); add(nid, {}, {}); nid.who.email = '';
  check('without an identity nothing is made (it could not be shared)', nid.ctx.exportAudits('repeats', refsOf(nid), '').success === false && nid.created.length === 0);
  const none = makeEnv(); add(none, { 'Analysis Type': 'Sales Analyzer' }, {});
  const noRows = none.ctx.exportAudits('repeats', refsOf(none), '');
  check('refs that are not in the chosen view: a clear message, nothing made, not counted against the hourly export limit (only against the separate try counter)', noRows.success === false && none.created.length === 0 && Object.keys(none.cacheStore).filter(k => k.indexOf('ex_rl_') === 0).length === 0 && Object.keys(none.cacheStore).filter(k => k.indexOf('ex_try_') === 0).length === 1, JSON.stringify(noRows));
  const f = makeEnv(); add(f, {}, {}); f.api.fail = 'Service invoked too many times for one day: spreadsheet. secret-internal-detail';
  const fail = f.ctx.exportAudits('repeats', refsOf(f), '');
  check('a Google failure is plain words, without the raw error, for an analyst; the lock is released', fail.success === false && /Google is busy/.test(fail.error) && fail.error.indexOf('secret-internal-detail') === -1 && f.lockState.released === 1, fail.error);
  f.who.email = 'danzen.palisoc@telus.com'; f.cacheStore = {};
  check('an admin also sees the details', /secret-internal-detail/.test(f.ctx.exportAudits('repeats', refsOf(f), '').error));
}

console.log('\nTest 6 — the page buttons');
{
  const h = load();
  const src = h.html;
  check('the old exportToSpreadsheet is gone from the page', src.indexOf('exportToSpreadsheet') === -1 && src.indexOf('.exportAudits(view, refs, title)') !== -1);
  const opened = [], alerts = [];
  h.ctx.window.open = u => opened.push(u);
  h.ctx.alert = m => alerts.push(m);
  h.reset();
  h.ctx._dbRFiltered = [{ 'Audit Ref': 'R1', 'Call Summary': 'SECRET TEXT', 'Team Member': 'A' }, { 'Audit Ref': 'R2' }, { 'Audit Ref': 'R1' }, { 'Audit Ref': '' }];
  h.ctx.dbExportFiltered('r');
  check('Export Filtered sends the view and the unique Audit Refs only', h.pending.length === 1 && h.pending[0].method === 'exportAudits' && h.pending[0].args[0] === 'repeats' && JSON.stringify(h.pending[0].args[1]) === JSON.stringify(['R1', 'R2']), JSON.stringify(h.pending.map(p => p.args)));
  check('none of the row data goes to the server', JSON.stringify(h.pending[0].args).indexOf('SECRET TEXT') === -1);
  check('the title says Filtered View', /Repeats — Filtered View/.test(h.pending[0].args[2]));
  check('the spinner shows', h.document.getElementById('db-export-overlay').style.display === 'flex');
  h.pending[0].success({ success: true, url: 'https://docs.google.com/spreadsheets/d/X/edit' });
  check('a Google Sheets link is opened and the spinner stops', opened[0] === 'https://docs.google.com/spreadsheets/d/X/edit' && h.document.getElementById('db-export-overlay').style.display === 'none');
  h.reset(); h.ctx._dbSFiltered = [{ 'Audit Ref': 'S1' }]; h.ctx.dbExportFiltered('s');
  check('the Sales button uses the sales view', h.pending[0].args[0] === 'sales');
  h.pending[0].success({ success: true, url: 'javascript:alert(1)' });
  check('a link that is not Google Sheets is never opened', opened.length === 1 && /not a Google Sheets link/.test(alerts[alerts.length - 1] || ''));
  h.reset(); h.ctx._dbAuditLog = [{ 'Audit Ref': 'L1' }, { 'Audit Ref': 'L2' }]; h.ctx.dbExportAll('al');
  check('Export All on the Audit Log sends the audit log view and its refs', h.pending[0].args[0] === 'auditlog' && JSON.stringify(h.pending[0].args[1]) === JSON.stringify(['L1', 'L2']) && /All Records/.test(h.pending[0].args[2]));
  h.pending[0].success({ success: false, error: 'You have already made 10 exports in the last hour.' });
  check('an error from the server is shown', /10 exports/.test(alerts[alerts.length - 1]) && h.document.getElementById('db-export-overlay').style.display === 'none');
  h.reset(); h.ctx.dbExportAll('al'); h.ctx._dbAuditLog = [];
  h.reset(); h.ctx._dbRFiltered = []; h.ctx.dbExportFiltered('r');
  check('with no records it says so and calls nothing', h.pending.length === 0 && /No records to export/.test(alerts[alerts.length - 1]));
  h.reset(); h.ctx._dbRFiltered = [{ 'Audit Ref': 'R9' }]; h.ctx.dbExportFiltered('r'); h.pending[0].failure({ message: 'boom' });
  check('a failed request is shown', /Export error: boom/.test(alerts[alerts.length - 1]));
}

console.log('\nTest — every try counts, so repeated empty tries cannot keep the lock and the sheets busy');
{
  const e = makeEnv(); add(e, { 'Analysis Type': 'Sales Analyzer' }, {});     // the audit is not in the Repeats view, so every try finds nothing
  let last = null, refused = 0;
  for (let i = 0; i < 60; i++) { last = e.ctx.exportAudits('repeats', refsOf(e), ''); if (/Too many export requests/.test(last.error || '')) refused++; }
  check('the first 60 tries are answered normally (nothing to export)', refused === 0 && /nothing to export/.test(last.error), JSON.stringify(last));
  const lockBefore = e.lockState.released;
  const r61 = e.ctx.exportAudits('repeats', refsOf(e), '');
  check('the 61st try in the hour is refused', r61.success === false && /Too many export requests/.test(r61.error), JSON.stringify(r61));
  check('and it is refused BEFORE the lock and the sheet reads (the lock was not taken)', e.lockState.released === lockBefore);
  const other = makeEnv(); add(other, { 'Analysis Type': 'Sales Analyzer' }, {}); other.who.email = 'someone.else@telus.com';
  Object.keys(e.cacheStore).forEach(k => { other.cacheStore[k] = e.cacheStore[k]; });
  check('the count is per person: somebody else is not affected', !/Too many export/.test(other.ctx.exportAudits('repeats', refsOf(other), '').error || ''));
}

console.log('\nTest — the folder is the one the app owner owns, not one somebody shared with the owner');
{
  const mk = (name, owner) => ({ name, getOwner: () => ({ getEmail: () => owner }), getId: () => name + '@' + owner });
  function run(folders, effective) {
    const created = [];
    const sb = { Logger: { log() {} }, JSON, Math, Date, parseInt, String, Error, PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
      Session: { getEffectiveUser: () => ({ getEmail: () => effective }) },
      DriveApp: { getFoldersByName: n => { const l = folders.filter(f => f.name === n); let i = 0; return { hasNext: () => i < l.length, next: () => l[i++] }; },
                  createFolder: n => { const f = mk(n, effective); created.push(f); return f; } } };
    const c = vm.createContext(sb); vm.runInContext(read('Config.gs'), c);
    return { picked: c._ownedFolder_('NH Analyzer Exports'), created };
  }
  const planted = mk('NH Analyzer Exports', 'attacker@telus.com'), mine = mk('NH Analyzer Exports', 'owner@telus.com');
  const a = run([planted, mine], 'owner@telus.com');
  check('with a planted folder listed first, the owner\'s own folder is used', a.picked === mine && a.created.length === 0, a.picked.getId());
  const b = run([planted], 'owner@telus.com');
  check('with only somebody else\'s folder, a new folder is made for the owner (the planted one is never used)', b.picked !== planted && b.created.length === 1);
  const c = run([], 'owner@telus.com');
  check('with none, one is created', c.created.length === 1);
  const d = run([mk('NH Analyzer Exports', 'OWNER@telus.com')], 'owner@telus.com');
  check('the owner comparison ignores case', d.created.length === 0);
  const f = run([planted], '');
  check('when the owner cannot be read the first match is used, as before', f.picked === planted && f.created.length === 0);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
