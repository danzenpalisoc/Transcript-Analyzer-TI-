/**
 * Nothing that belongs to one agent may survive into another agent's audit.
 *
 * agentEmail, Team Leader, Ops Manager, LOB, Locale and VTID are filled from a
 * roster row. _applyRosterRow only wrote the fields that were truthy and never
 * cleared the rest, so when agent B's row had a blank Team Leader or VTID,
 * agent A's value stayed in the form and was submitted with B -- and A's email
 * address (a HIDDEN input) went on receiving B's evaluation.
 *
 * Related holes closed here:
 *  - editing Team Member after a SAP lookup replaced the SAP ID but left A's
 *    roster fields and email in place;
 *  - Clear All did not invalidate an in-flight transcript autofill, so it could
 *    write the previous agent's identity and Interaction ID into the blank form;
 *  - the name lookup and the SAP lookup shared one timer, so typing in one field
 *    cancelled a pending lookup from the other.
 *
 * Usage: node tests/identity-reset-test.js
 */
const { load, makeChecker } = require('./lib/page-harness');
const h = load(process.argv[2]);
const check = makeChecker();
const { ctx, document, pending, sleep, reset, field, setField, typeSap, take } = h;

const A = { participant: 'Agent A', teamLeader: 'TL-A', opsManager: 'OM-A', lineOfBusiness: 'LOB-A', locale: 'Site-A', vtid: 'V-A', agentEmail: 'a@telus.com' };
const B_SPARSE = { participant: 'Agent B', teamLeader: '', opsManager: '', lineOfBusiness: '', locale: '', vtid: '', agentEmail: '' };
const ROSTER_FIELDS = ['teamLeader', 'opsManager', 'lineOfBusiness', 'locale', 'vtid', 'agentEmail'];

async function resolveA() {
  await typeSap('1111111');
  take('getRosterBySapId').success(A);
}

async function main() {
  console.log('\nTest 1 — a sparse row for agent B must not inherit agent A\'s fields');
  reset(); await resolveA();
  check('precondition: A is filled', field('teamLeader') === 'TL-A' && field('agentEmail') === 'a@telus.com');
  await typeSap('2222222');
  take('getRosterBySapId').success(B_SPARSE);
  check('Team Member is B', field('participant') === 'Agent B', field('participant'));
  for (const id of ROSTER_FIELDS) check(id + ' is blank, not A\'s', field(id) === '', field(id));

  console.log('\nTest 2 — while B is still loading, A\'s fields are already gone');
  reset(); await resolveA();
  await typeSap('2222222');
  for (const id of ROSTER_FIELDS) check(id + ' cleared at lookup start', field(id) === '', field(id));
  check('A\'s resolved name cleared too', field('participant') === '', field('participant'));

  console.log('\nTest 3 — B not found: nothing of A is left behind');
  reset(); await resolveA();
  await typeSap('2222222');
  take('getRosterBySapId').success(null);
  for (const id of ['participant'].concat(ROSTER_FIELDS)) check(id + ' blank', field(id) === '', field(id));

  console.log('\nTest 4 — a name the analyst typed themselves is not wiped by a lookup');
  reset();
  setField('participant', 'Typed Name');
  await typeSap('2222222');
  check('typed name kept while loading', field('participant') === 'Typed Name', field('participant'));
  take('getRosterBySapId').success({ participant: '', teamLeader: 'TL-X', opsManager: '', lineOfBusiness: '', locale: '', vtid: '', agentEmail: '' });
  check('typed name kept after a reply with no name', field('participant') === 'Typed Name', field('participant'));
  check('roster data still applied', field('teamLeader') === 'TL-X', field('teamLeader'));

  console.log('\nTest 5 — editing Team Member re-resolves the roster, not just the SAP ID');
  reset(); await resolveA();
  setField('participant', 'Agent B');
  ctx.debounceLookupSap();
  await sleep(700);
  const nameReq = take('getSapIdForParticipant');
  check('a name lookup was sent', !!nameReq);
  if (nameReq) {
    nameReq.success('2222222');
    check('SAP ID replaced', field('sapId') === '2222222', field('sapId'));
    for (const id of ROSTER_FIELDS) check(id + ' no longer A\'s', field(id) !== A[id] || field(id) === '', field(id));
    const rosterReq = pending.find(p => p.method === 'getRosterBySapId' && p.args[0] === '2222222');
    check('the full roster lookup was requested for the new SAP ID', !!rosterReq);
    if (rosterReq) {
      pending.splice(pending.indexOf(rosterReq), 1);
      rosterReq.success({ participant: 'Agent B', teamLeader: 'TL-B', opsManager: 'OM-B', lineOfBusiness: 'LOB-B', locale: 'Site-B', vtid: 'V-B', agentEmail: 'b@telus.com' });
      check('B\'s email applied', field('agentEmail') === 'b@telus.com', field('agentEmail'));
      check('B\'s team leader applied', field('teamLeader') === 'TL-B', field('teamLeader'));
    }
  }

  console.log('\nTest 6 — the name lookup and SAP lookup no longer cancel each other');
  reset();
  setField('sapId', '2222222'); ctx.debounceSapLookup();
  setField('participant', 'Agent B'); ctx.debounceLookupSap();
  await sleep(800);
  check('SAP lookup still fired', !!take('getRosterBySapId'));
  check('name lookup fired too', !!take('getSapIdForParticipant'));

  console.log('\nTest 7 — Clear All cancels a lookup that has not fired yet');
  reset();
  setField('sapId', '2222222'); ctx.debounceSapLookup();
  ctx.clearAllFields();
  await sleep(800);
  check('no lookup was sent after Clear All', !take('getRosterBySapId'));

  console.log('\nTest 8 — an in-flight transcript autofill cannot refill a cleared form');
  reset();
  setField('transcript', 'Agent: hello');
  ctx.scheduleAutofill();
  await sleep(1000);
  const af = take('getTranscriptMetadataAndRoster');
  check('autofill was sent', !!af);
  if (af) {
    ctx.clearAllFields();
    af.success({ meta: { participant: 'Agent A', interactionId: 'INT-OLD', startTime: '09:00', duration: '10:00' },
                 sapId: '1111111', roster: A });
    check('Interaction ID stays blank', field('interactionId') === '', field('interactionId'));
    check('SAP ID stays blank', field('sapId') === '', field('sapId'));
    check('Team Member stays blank', field('participant') === '', field('participant'));
    check('agentEmail stays blank', field('agentEmail') === '', field('agentEmail'));
  }

  console.log('\nTest 9 — autofill still works when nothing was cleared');
  reset();
  setField('transcript', 'Agent: hello');
  ctx.scheduleAutofill();
  await sleep(1000);
  const af2 = take('getTranscriptMetadataAndRoster');
  if (af2) {
    af2.success({ meta: { participant: 'Agent A', interactionId: 'INT-1' }, sapId: '1111111', roster: A });
    check('SAP ID filled', field('sapId') === '1111111', field('sapId'));
    check('email filled', field('agentEmail') === 'a@telus.com', field('agentEmail'));
  } else { check('autofill was sent', false); }

  console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
  process.exit(check.state.failures ? 1 : 0);
}
main();
