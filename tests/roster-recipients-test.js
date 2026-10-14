/**
 * getRecipientsFromRoster_() — who gets an audit email in TEST mode (Admin/Dev) and
 * for the QA / trainer lists.
 *
 * Reported: sending an audit in TEST mode failed with "Could not resolve recipient
 * email addresses". The log said only `getRecipientsFromRoster_("Admin/Dev"): 0 found`
 * although the sheet has Admin/Dev rows. Nothing said WHICH tab was read, whether the
 * rows came from a 2-hour cache, or whether the columns were found, so a stale cache
 * and a same-named tab with other columns looked identical.
 *
 * Checks: the right rows are found from the real layout; an old v2 cache entry can no
 * longer hide them; and the log names the tab, the columns, the source and the roles
 * present.
 *
 * Usage: node tests/roster-recipients-test.js
 */
const vm = require('vm');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const code = read('Code.gs');

const HEADERS = ['Email Address', 'Name', 'Role', 'Supervisor', 'Supervisor Email', 'Status', 'ccList'];
const ROWS = [
  ['Lobell.Bantaya@telus.com', 'Lobell Bantaya', 'Admin/Dev', 'Lobell Bantaya', 'Lobell.Bantaya@telus.com', 'Approved', ''],
  ['Carlos.Maldonado@telus.com', 'Carlos Maldonado', 'Admin', 'Lobell Bantaya', 'Lobell.Bantaya@telus.com', 'Approved', ''],
  ['danzen.palisoc@telus.com', 'Danzen Palisoc', 'Admin/Dev', 'Reymark Valor', 'reymark.valor@telus.com', 'Approved', ''],
  ['lauralea.upshall@telus.com', 'Laura Lea Upshall', 'Admin/No Test', '', '', 'Approved', ''],
  ['DANZEN.PALISOC@telus.com', 'Danzen Palisoc', 'Admin/Dev', '', '', 'Approved', ''],   // same person again
  ['not-an-email', 'Broken Row', 'Admin/Dev', '', '', 'Approved', ''],
];

function env(opts) {
  opts = opts || {};
  const logs = [];
  const store = {};
  const tabs = opts.tabs || { Roster: { name: 'Roster', data: [HEADERS].concat(ROWS) } };
  const sandbox = {
    Logger: { log: m => logs.push(String(m)) }, JSON, Array, String, Object, Error, RegExp,
    AUDIT_TRACKING_SS_ID: 'ssid-123456',
    CacheService: { getScriptCache: () => ({ get: k => (k in store ? store[k] : null), put: (k, v) => { store[k] = v; }, remove: k => { delete store[k]; } }) },
    openSheetWithRetry: () => ({
      getSheetByName: n => (tabs[n] ? { getName: () => tabs[n].name, getDataRange: () => ({ getValues: () => tabs[n].data }) } : null),
    }),
  };
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext('var _rosterRecipientsInMemory = null;\n' + extractFunction(code, 'getRecipientsFromRoster_'), ctx);
  return { ctx, logs, store };
}
const last = e => e.logs[e.logs.length - 1];

console.log('\nTest 1 — Admin/Dev rows are found from the real layout, de-duplicated, invalid emails skipped');
{
  const e = env();
  const r = e.ctx.getRecipientsFromRoster_('Admin/Dev');
  const emails = r.map(x => x.email.toLowerCase()).sort();
  check('finds the two distinct Admin/Dev people', emails.join() === 'danzen.palisoc@telus.com,lobell.bantaya@telus.com', emails.join());
  check('"Admin" does not match "Admin/Dev" and vice versa', e.ctx.getRecipientsFromRoster_('Admin').length === 1);
  check('role matching ignores case', env().ctx.getRecipientsFromRoster_('admin/dev').length === 2);
}

console.log('\nTest 2 — a stale v2 cache entry can no longer hide the rows');
{
  const e = env();
  // what an old, wrong read would have left behind: rows that match nobody
  e.store['audit_roster_rows_v2'] = JSON.stringify([{ name: 'x', role: '', email: '', supervisor: '', supervisorEmail: '' }]);
  const r = e.ctx.getRecipientsFromRoster_('Admin/Dev');
  check('the stale v2 entry is ignored', r.length === 2, 'found ' + r.length);
  check('the fresh read is cached under a new key', !!e.store['audit_roster_rows_v3'] && /Lobell/.test(e.store['audit_roster_rows_v3']));
}

console.log('\nTest 3 — the log says which tab, which columns, where the rows came from, and which roles exist');
{
  const e = env();
  e.ctx.getRecipientsFromRoster_('Admin/Dev');
  const read1 = e.logs.find(l => /read tab/.test(l)) || '';
  check('names the tab', /read tab "Roster"/.test(read1), read1);
  check('reports the detected columns', /name col=1, role col=2, email col=0/.test(read1), read1);
  check('final line reports source=sheet', /source=sheet/.test(last(e)), last(e));
  check('final line lists the roles present', /"Admin\/Dev":4/.test(last(e)) && /"Admin":1/.test(last(e)), last(e));
  e.ctx.getRecipientsFromRoster_('Admin/Dev');   // second call in the same execution
  check('a repeat call in one execution reports source=memory', /source=memory/.test(last(e)), last(e));

  const fresh = env();
  fresh.store['audit_roster_rows_v3'] = e.store['audit_roster_rows_v3'];
  fresh.ctx.getRecipientsFromRoster_('Admin/Dev');
  check('a call served from the script cache reports source=cache', /source=cache/.test(last(fresh)), last(fresh));
}

console.log('\nTest 4 — a same-named tab with other columns is called out, not silent');
{
  const e = env({ tabs: { Roster: { name: 'Roster', data: [['sap_ID', 'Agent_Name', 'Domain'], ['1', 'A', 'B']] } } });
  const r = e.ctx.getRecipientsFromRoster_('Admin/Dev');
  check('returns nobody', r.length === 0);
  check('logs a WARNING that the Role/Email columns are missing', e.logs.some(l => /WARNING/.test(l) && /Role/.test(l)), e.logs.join(' | '));
}

console.log('\nTest 5 — a missing tab is reported and returns nobody');
{
  const e = env({ tabs: {} });
  check('no tab -> empty list', e.ctx.getRecipientsFromRoster_('Admin/Dev').length === 0);
  check('the log says the tab was not found', e.logs.some(l => /not found/i.test(l)), e.logs.join(' | '));
}

console.log('\nTest 6 — sendAuditEmail explains WHICH list came up empty');
{
  const send = extractFunction(code, 'sendAuditEmail');
  check('TEST-mode message names the Admin/Dev list', /TEST mode: no Admin\/Dev address was found/.test(send));
  check('LIVE-mode message names agent, leader, ops manager and QA', /LIVE mode: none of the agent, team leader, ops manager or QA/.test(send));
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
