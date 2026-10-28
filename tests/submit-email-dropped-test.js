/**
 * Submit & Send Email must tell the analyst when an address was NOT emailed.
 *
 * sendAuditEmail drops any address that is not on a company domain (EmailGuard) and returns it as droppedRecipients. The page
 * used to ignore that field, so when the team member's own address was dropped (for example name@gapps.telus.com) the analyst saw
 * "Audit emailed successfully" and nobody knew that person had not received the audit.
 *
 * Runs the real handleSubmitEmail from index.html with a fake page and a fake server reply.
 *
 * Usage: node tests/submit-email-dropped-test.js
 */
const vm = require('vm');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const src = extractFunction(read('index.html'), 'handleSubmitEmail');

function run(reply) {
  const shown = { success: '', warning: '', error: '', locked: null };
  const els = {
    teamLeader: { value: 'Jude Ngah' }, opsManager: { value: 'Kim Robert Osian' },
    'btn-submit': { disabled: false }, 'email-overlay': { classList: { add() {}, remove() {} } }
  };
  const sandbox = {
    document: { getElementById: id => els[id] || { value: '', classList: { add() {}, remove() {} } } },
    _lastHTML: '<div></div>', _emailMode: 'live', _currentAuditRef: 'NHA-1',
    getFormData: () => ({ auditRef: 'NHA-1' }), getEffectiveHTML: () => '<div></div>',
    showError: m => { shown.error = m; }, showWarning: m => { shown.warning = m; }, showSuccess: m => { shown.success = m; },
    lockFormAfterSubmit: ref => { shown.locked = ref; }, console,
    google: { script: { run: {
      withSuccessHandler(ok) { this._ok = ok; return this; },
      withFailureHandler() { return this; },
      sendAuditEmail() { this._ok(reply); }
    } } }
  };
  vm.createContext(sandbox);
  vm.runInContext(src + '\nhandleSubmitEmail();', sandbox);
  return shown;
}

console.log('Test 1 - a dropped address is shown to the analyst');
const a = run({ success: true, recipients: ['tl@telus.com', 'qa@telus.com'], auditRef: 'NHA-1', droppedRecipients: ['jeimi.perez@gapps.telus.com'] });
check('the success line still names who was emailed', /tl@telus.com, qa@telus.com/.test(a.success));
check('a warning names the address that was NOT emailed', /NOT emailed/.test(a.warning) && a.warning.indexOf('jeimi.perez@gapps.telus.com') !== -1);
check('the warning says the person did not receive the audit', /did not receive this audit/.test(a.warning));
check('the form still locks', a.locked === 'NHA-1');

console.log('\nTest 2 - nothing dropped: no warning');
const b = run({ success: true, recipients: ['a@telus.com'], auditRef: 'NHA-1' });
check('no warning', b.warning === '');
check('the success line is shown', /a@telus.com/.test(b.success));

console.log('\nTest 3 - a quota warning and a dropped address are both shown');
const c = run({ success: true, emailWarning: 'Audit saved. Email could not be sent.', recipients: [], auditRef: 'NHA-1', droppedRecipients: ['x@gmail.com'] });
check('the one warning carries both messages', /Email could not be sent/.test(c.warning) && c.warning.indexOf('x@gmail.com') !== -1);

console.log('\nTest 4 - a failed send is still an error');
const d = run({ success: false, error: 'boom' });
check('error shown, no lock', /boom/.test(d.error) && d.locked === null);

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
