/**
 * Dashboard_Data "Auditor Email" column.
 *
 * The column is the LAST one (the sheet is written by position, so nothing in the middle may
 * move), it holds the signed-in caller's address as the server reads it, and falls back to the
 * Users sheet by name. A name with several different addresses gets all of them, once each,
 * separated by ", ". backfillAuditorEmail() fills only blank cells and writes only that column.
 *
 * Usage: node tests/auditor-email-test.js
 */
const vm = require('vm');
const { read, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

function makeEnv(opts) {
  opts = opts || {};
  const logs = [];
  const who = { email: opts.email === undefined ? 'danzen.palisoc@telus.com' : opts.email };
  const sandbox = {
    Logger: { log: m => logs.push(String(m)) }, JSON, Array, Object, String, Number, Math, Date, RegExp, isFinite, isNaN, Error,
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Session: { getActiveUser: () => ({ getEmail: () => who.email }), getScriptTimeZone: () => 'UTC' },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    CacheService: { getUserCache: () => ({ get: () => null, put() {} }), getScriptCache: () => ({ get: () => null, put() {} }) },
  };
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('Config.gs'), ctx);
  vm.runInContext(read('ObserverHelper.gs'), ctx);
  vm.runInContext(read('DataRepair.gs'), ctx);

  const users = new FakeSheet(['Name', 'Email Address', 'Role'], opts.users || [
    ['Danzen Palisoc', 'Danzen.Palisoc@telus.com', 'QA'],
    ['Ana Cruz', 'ana.cruz@telus.com', 'QA'],
    ['Twin Name', 'twin.one@telus.com', 'QA'],
    ['Twin Name', 'twin.two@telus.com', 'QA'],
    ['No Email', '', 'QA'],
  ]);
  users.getDataRange = () => ({ getValues: () => users.grid.map(r => r.slice()) });
  const dash = new FakeSheet(opts.headers || ctx.DASHBOARD_HEADERS, []);
  const mainSs = { getSheetByName: n => (n === ctx.DASHBOARD_DATA_SHEET ? dash : null) };
  const usersSs = { getSheetByName: n => (n === ctx.USERS_TAB && !opts.noUsersTab ? users : null) };
  ctx.openSheetWithRetry = id => {
    if (opts.usersThrow && id === ctx.USERS_SS_ID) throw new Error('permission denied');
    return id === ctx.USERS_SS_ID ? usersSs : mainSs;
  };
  ctx.getOrCreateSpreadsheet = () => mainSs;
  let invalidated = 0;
  ctx.invalidateDashboardCache_ = () => { invalidated++; };
  const addRow = o => dash.grid.push(dash.grid[0].map(h => (o[h] !== undefined ? o[h] : '')));
  return { ctx, dash, users, logs, who, addRow, invalidated: () => invalidated };
}

console.log('\nTest 1 — the column is the last header, after Agent Email');
{
  const H = makeEnv().ctx.DASHBOARD_HEADERS;
  check('last header is Auditor Email', H[H.length - 1] === 'Auditor Email', H.slice(-3).join(' | '));
  check('Agent Email is still just before it (no column moved)', H[H.length - 2] === 'Agent Email');
  check('Observer Name is still column 11', H.indexOf('Observer Name') === 10);
  check('no duplicate header', new Set(H).size === H.length);
}

console.log('\nTest 2 — resolveAuditorEmail_');
{
  let e = makeEnv({ email: 'Some.Analyst@Telus.com ' });
  check('uses the signed-in caller, lower-cased and trimmed', e.ctx.resolveAuditorEmail_('Danzen Palisoc') === 'some.analyst@telus.com');

  e = makeEnv({ email: '' });
  check('caller hidden: falls back to the Users sheet by name (any case)', e.ctx.resolveAuditorEmail_('  danzen PALISOC ') === 'danzen.palisoc@telus.com');
  check('caller hidden, name on two rows with two addresses: both, in sheet order', e.ctx.resolveAuditorEmail_('Twin Name') === 'twin.one@telus.com, twin.two@telus.com', e.ctx.resolveAuditorEmail_('Twin Name'));
  check('three rows, one address twice: each address once', makeEnv({ email: '', users: [['Dup', 'a@x.com', ''], ['Dup', 'A@x.com', ''], ['Dup', 'b@y.com', '']] }).ctx.resolveAuditorEmail_('dup') === 'a@x.com, b@y.com');
  check('caller hidden, name not on the sheet: blank', e.ctx.resolveAuditorEmail_('Nobody Here') === '');
  check('caller hidden, row without an email: blank', e.ctx.resolveAuditorEmail_('No Email') === '');
  check('caller hidden, no name: blank', e.ctx.resolveAuditorEmail_('') === '');

  e = makeEnv({ email: '', usersThrow: true });
  let out = 'unset', threw = false;
  try { out = e.ctx.resolveAuditorEmail_('Danzen Palisoc'); } catch (x) { threw = true; }
  check('an unreadable Users sheet never stops the audit: blank, no throw', !threw && out === '');
  check('and the reason is logged', e.logs.some(l => /resolveAuditorEmail_/.test(l)));

  e = makeEnv({ email: '', noUsersTab: true });
  check('no Users tab: blank', e.ctx.resolveAuditorEmail_('Danzen Palisoc') === '');
}

console.log('\nTest 3 — backfillAuditorEmail fills only blank cells in that one column');
{
  const e = makeEnv();
  e.addRow({ 'Audit Ref': 'A1', 'Observer Name': 'Danzen Palisoc', 'Team Leader': 'TL' });
  e.addRow({ 'Audit Ref': 'A2', 'Observer Name': 'ana cruz' });
  e.addRow({ 'Audit Ref': 'A3', 'Observer Name': 'Ana Cruz', 'Auditor Email': 'kept@telus.com' });
  e.addRow({ 'Audit Ref': 'A4', 'Observer Name': 'Twin Name' });
  e.addRow({ 'Audit Ref': 'A5', 'Observer Name': 'Nobody Here' });
  e.addRow({ 'Audit Ref': 'A6', 'Observer Name': '' });
  const H = e.dash.grid[0], ec = H.indexOf('Auditor Email');
  const others = () => JSON.stringify(e.dash.grid.map(r => r.filter((v, i) => i !== ec)));
  const before = others();
  const r = e.ctx.backfillAuditorEmail();
  const col = e.dash.grid.slice(1).map(row => row[ec]);
  check('Danzen filled (Users sheet address, lower-cased)', col[0] === 'danzen.palisoc@telus.com', JSON.stringify(col));
  check('name matched ignoring case', col[1] === 'ana.cruz@telus.com');
  check('an email already there is never overwritten', col[2] === 'kept@telus.com');
  check('a name on two rows gets both addresses', col[3] === 'twin.one@telus.com, twin.two@telus.com', JSON.stringify(col));
  check('an unknown name is left blank', col[4] === '');
  check('a blank name is left blank', col[5] === '');
  check('no other column changed', others() === before);
  check('counts are right', r.filled === 3 && r.alreadySet === 1, JSON.stringify(r));
  check('the report names who was not matched and who got several addresses', r.noMatch.join() === 'Nobody Here' && r.multiple.join() === 'Twin Name', JSON.stringify(r));
  check('the dashboard cache was cleared', e.invalidated() === 1);
  const again = e.ctx.backfillAuditorEmail();
  check('running it again changes nothing', again.filled === 0 && again.alreadySet === 4, JSON.stringify(again));
}

console.log('\nTest 4 — backfill refuses safely');
{
  const short = makeEnv().ctx.DASHBOARD_HEADERS.slice(0, -1);
  const e = makeEnv({ headers: short });
  e.dash.grid.push(short.map((h, i) => (h === 'Observer Name' ? 'Danzen Palisoc' : '')));
  const before = JSON.stringify(e.dash.grid);
  const r = e.ctx.backfillAuditorEmail();
  check('header not added yet: nothing changed', JSON.stringify(e.dash.grid) === before && r.filled === 0);
  check('the log says to run addMissingHeaders first', e.logs.some(l => /addMissingHeaders/.test(l)), e.logs.join(' | '));

  const a = makeEnv();
  a.who.email = 'some.analyst@telus.com';
  let msg = '';
  try { a.ctx.backfillAuditorEmail(); } catch (x) { msg = x.message; }
  check('only admins can run it', /admin/i.test(msg), msg);

  const empty = makeEnv();
  const r2 = empty.ctx.backfillAuditorEmail();
  check('no data rows: nothing to do', r2.filled === 0 && empty.dash.grid.length === 1);
}

console.log('\nTest 5 — the submit writes the email as the last cell of the Dashboard_Data row');
{
  const src = read('Code.gs');
  const m = src.match(/appendRow\(dSheet, \[[\s\S]*?\]\);/);
  check('the Dashboard_Data row ends with resolveAuditorEmail_(observerName)', !!m && /formData\.agentEmail \|\| '',\s*resolveAuditorEmail_\(observerName\)\s*\]\)/.test(m[0]));
}

console.log('\n' + (check.state.failures === 0 ? 'ALL PASSED' : check.state.failures + ' FAILED'));
process.exit(check.state.failures === 0 ? 0 : 1);
