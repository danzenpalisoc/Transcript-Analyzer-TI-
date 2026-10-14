/**
 * An agent's evaluation must be emailed to the person the SAP ID identifies,
 * never to someone who merely shares their name.
 *
 * buildAuditRecipients_ resolved the tenured agent's address by NAME first
 * (lookupAgentEmail_ -> resolveEmail_) and only fell back to the address the SAP
 * lookup had shown the analyst. _getAgentEmailMap built that name map with "last
 * row wins", so two roster agents with the same name meant one of them silently
 * received the other's evaluation.
 *
 * Contract now:
 *   1. the roster row for the SAP ID (exact),
 *   2. the address the SAP lookup showed the analyst,
 *   3. by name -- but only when exactly one person has that name.
 * A name shared by people with different addresses resolves to nobody. Sending no
 * email is recoverable; sending it to the wrong person is not.
 *
 * Usage: node tests/audit-recipient-test.js
 */
const vm = require('vm');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');

const code = read('Code.gs');
const check = makeCheck();

const SRC = [
  extractFunction(code, 'buildAuditRecipients_'),
  extractFunction(code, '_agentEmailForSap_'),
  extractFunction(code, 'lookupAgentEmail_'),
  extractFunction(code, 'resolveEmail_'),
  extractFunction(code, '_getAgentEmailMap'),
].join('\n');

// `extra` is visible to the code under test from the start. `override` replaces
// functions that SRC itself declares (lookupAgentEmail_, resolveEmail_); it must be
// applied AFTER SRC loads, or the real declaration silently wins over the stub.
function sandbox(extra, override) {
  const logs = [];
  const ctx = vm.createContext(Object.assign({
    Logger: { log: m => logs.push(String(m)) }, JSON, Error, Array, String, Object,
    ROSTER_COL_SAP_ID: 0, QA_ROLE: 'QA',
    getTraineeInfo_: () => null, getTrainerInfo_: () => null, lookupSapId_: () => '',
    getRecipientsFromRoster_: () => [],
  }, extra));
  vm.runInContext('var _agentEmailMapInMemory = null;' + SRC, ctx);
  Object.assign(ctx, override || {});
  ctx.__logs = logs;
  return ctx;
}

// A roster sheet backed by rows: [agent name, email].
function rosterEnv(rows) {
  const data = [['Agent_Name', 'Team Member Email']].concat(rows);
  let opens = 0;
  return {
    opens: () => opens,
    env: {
      CacheService: { getScriptCache: () => ({ get: () => null, put() {}, remove() {} }) },
      openSheetWithRetry: () => {
        opens++;
        return { getSheetByName: n => (n === 'roster' || n === 'Roster' ? { getDataRange: () => ({ getValues: () => data }) } : null) };
      },
      ROSTER_SHEET_ID: 'x', _getGlobalRosterData: () => ({ header: [], rows: [] }),
    },
  };
}

console.log('\nTest 1 — the roster row for the SAP ID wins over a name match');
{
  const ctx = sandbox({ _findRosterRow_: () => ({ row: [], email: 'sap-person@telus.com' }) },
    { lookupAgentEmail_: n => (n === 'Maria Cruz' ? 'namesake@telus.com' : ''),
      resolveEmail_:     n => (n === 'Maria Cruz' ? 'namesake@telus.com' : '') });
  const r = ctx.buildAuditRecipients_({ participant: 'Maria Cruz', sapId: '1234567', agentEmail: '' });
  check('SAP-keyed address is used', r.indexOf('sap-person@telus.com') !== -1, JSON.stringify(r));
  check('namesake is NOT emailed', r.indexOf('namesake@telus.com') === -1, JSON.stringify(r));
}

console.log('\nTest 2 — no address on the SAP row: use the one the analyst was shown');
{
  const ctx = sandbox({ _findRosterRow_: () => ({ row: [], email: '' }) },
    { lookupAgentEmail_: () => '', resolveEmail_: () => '' });
  const r = ctx.buildAuditRecipients_({ participant: 'Maria Cruz', sapId: '1234567', agentEmail: 'shown@telus.com' });
  check('form address used', r.indexOf('shown@telus.com') !== -1, JSON.stringify(r));
}

console.log('\nTest 3 — shared name, no SAP ID, no address: nobody is emailed');
{
  const ctx = sandbox({ _findRosterRow_: () => null }, { lookupAgentEmail_: () => '', resolveEmail_: () => '' });
  const r = ctx.buildAuditRecipients_({ participant: 'Maria Cruz', sapId: '', agentEmail: '' });
  check('no agent address', r.length === 0, JSON.stringify(r));
}

console.log('\nTest 4 — a unique name still resolves when there is no SAP ID');
{
  const ctx = sandbox({ _findRosterRow_: () => null }, { lookupAgentEmail_: () => 'jose@telus.com', resolveEmail_: () => '' });
  const r = ctx.buildAuditRecipients_({ participant: 'Jose Reyes', sapId: '', agentEmail: '' });
  check('name address used', r.indexOf('jose@telus.com') !== -1, JSON.stringify(r));
}

console.log('\nTest 5 — the name map marks a shared name as ambiguous instead of last-row-wins');
{
  const { env, opens } = rosterEnv([
    ['Maria Cruz', 'maria.one@telus.com'],
    ['Maria Cruz', 'maria.two@telus.com'],
    ['Jose Reyes', 'jose@telus.com'],
    ['Ana Lim', 'ana@telus.com'],
    ['Ana Lim', 'ANA@telus.com'],   // same person listed twice, not ambiguous
  ]);
  const ctx = sandbox(env);
  check('shared name resolves to nobody (lookupAgentEmail_)', ctx.lookupAgentEmail_('Maria Cruz') === '', ctx.lookupAgentEmail_('Maria Cruz'));
  check('unique name resolves', ctx.lookupAgentEmail_('Jose Reyes') === 'jose@telus.com', ctx.lookupAgentEmail_('Jose Reyes'));
  check('same address listed twice is not ambiguous', ctx.lookupAgentEmail_('Ana Lim').toLowerCase() === 'ana@telus.com', ctx.lookupAgentEmail_('Ana Lim'));
  const before = opens();
  check('resolveEmail_ also refuses a shared name', ctx.resolveEmail_('Maria Cruz') === '', ctx.resolveEmail_('Maria Cruz'));
  check('...without falling through to a first-match sheet scan', opens() === before, 'extra opens: ' + (opens() - before));
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
