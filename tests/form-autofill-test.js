/**
 * The call details read from a transcript (code review batch 3).
 *
 * Pasting a SECOND transcript used to leave the first one's Interaction ID, start time and duration in the form (setIfEmpty only
 * fills an empty field), so the next audit was saved under the wrong Interaction ID. A field still holding what an earlier
 * transcript put there is now replaced; a field the person typed or changed is never touched.
 *
 * Also: sendAuditEmail looks the recipients up for the agent on the SAVED audit, even when the form's SAP ID has been blanked or edited.
 *
 * Usage: node tests/form-autofill-test.js
 */
const { load } = require('./lib/page-harness');
const { makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const h = load();

const T = (id, start, dur) => 'Interaction ID: ' + id + '\nTranscript Start Time: ' + start + '\nTranscript Duration: ' + dur + '\nAgent: hello\nCustomer: hi';
const val = id => h.document.getElementById(id).value;
const paste = async txt => { h.document.getElementById('transcript').value = txt; h.ctx.scheduleAutofill(); await h.sleep(900); };

(async () => {
  console.log('\nTest 1 — a second transcript replaces the first one\'s call details');
  {
    h.ctx._formEpoch = h.ctx._formEpoch || 0;
    await paste(T('INT-AAA', '2026-10-13 09:00', '0:05:00'));
    check('the first transcript fills the fields', val('interactionId') === 'INT-AAA' && val('startTime') === '2026-10-13 09:00' && val('duration') === '0:05:00', JSON.stringify([val('interactionId'), val('startTime'), val('duration')]));
    await paste(T('INT-BBB', '2026-10-13 10:30', '0:12:10'));
    check('a second transcript REPLACES the Interaction ID (it kept INT-AAA before)', val('interactionId') === 'INT-BBB', val('interactionId'));
    check('and the start time and the duration', val('startTime') === '2026-10-13 10:30' && val('duration') === '0:12:10', JSON.stringify([val('startTime'), val('duration')]));
    // the person types over a field: it is theirs now
    h.document.getElementById('interactionId').value = 'MY-OWN-ID';
    await paste(T('INT-CCC', '2026-10-14 08:00', '0:01:00'));
    check('a value the person typed is NOT overwritten by the next paste', val('interactionId') === 'MY-OWN-ID', val('interactionId'));
    check('while the fields they did not touch still follow the transcript', val('startTime') === '2026-10-14 08:00' && val('duration') === '0:01:00');
    // a transcript that has no Interaction ID line must not leave the old one behind
    await paste(T('INT-DDD', '2026-10-15 08:00', '0:02:00'));
    h.document.getElementById('interactionId').value = '';
    await paste(T('INT-EEE', '2026-10-15 09:00', '0:02:30'));
    check('an empty field is filled again', val('interactionId') === 'INT-EEE', val('interactionId'));
    await paste('Agent: hello\nCustomer: hi (no header lines at all)');
    check('a transcript with no Interaction ID line clears the one an earlier transcript set, rather than keeping it', val('interactionId') === '', val('interactionId'));
  }

  console.log('\nTest 2 — helper behaviour');
  {
    const set = (id, v) => { h.document.getElementById(id).value = v; };
    set('interactionId', ''); h.ctx._afFromTranscript = {};
    h.ctx.setFromTranscript('interactionId', 'A1');
    check('empty field: filled', val('interactionId') === 'A1');
    h.ctx.setFromTranscript('interactionId', 'A2');
    check('still what we put there: replaced', val('interactionId') === 'A2');
    set('interactionId', 'typed');
    h.ctx.setFromTranscript('interactionId', 'A3');
    check('changed by the person: left alone', val('interactionId') === 'typed');
    h.ctx.setFromTranscript('interactionId', '');
    check('and not cleared by a transcript without it', val('interactionId') === 'typed');
    h.ctx.setFromTranscript('nonexistentField', 'x');
    check('a missing element is ignored', true);
  }

  console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
  process.exit(check.state.failures ? 1 : 0);
})();
