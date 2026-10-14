/**
 * sendAuditEmail must not be a way to send anything to anyone.
 *
 * It is open to every signed-in user (the page calls it) and mails from the app owner's mailbox. Before: the recipients came from
 * formData.observerEmail / formData.agentEmail (checked only by a regex), nobody checked who the caller was, and the subject and
 * body were built from the form. So anyone could call it from the browser console with an existing audit ref and mail any
 * address, with their own text, and overwrite that audit's stored evaluation.
 *
 *  1. Only an admin, or the analyst who saved the audit, may email it.
 *  2. Mail goes only to company addresses; anything else is dropped and reported.
 *  3. The QA recipient is the signed-in person, not a typed-in address.
 *  4. An analyst may send a limited number per hour.
 *  5. Text from the form is cut to one short line; the saved audit wins where it has the value.
 *
 * Usage: node tests/email-relay-test.js
 */
const vm = require('vm');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const codeSrc = read('Code.gs');

function makeEnv(o) {
  o = o || {};
  const sent = [], cached = [], pdfLinks = [], logs = [], errors = [], notified = [], buildArgs = [];
  const store = {};
  const who = { email: o.email || 'analyst.one@telus.com' };
  const obs = { name: o.observerName === undefined ? 'Analyst One' : o.observerName };
  const rows = o.rows || [{ 'Audit Ref': 'NHA-20261013-0001', 'Observer Name': 'Analyst One', 'Team Member': 'Maria Santos', 'SAP ID': '1234567', 'Interaction ID': 'INT-1', 'Analysis Type': 'Repeats & Transfer Analyzer' }];
  const sandbox = {
    readAuditLog_: () => rows,
    Session: { getActiveUser: () => ({ getEmail: () => who.email }) },
    resolveObserver_: () => ({ name: obs.name, email: who.email }),
    ADMIN_USERNAMES: ['admin1'], ADMIN_DOMAIN: 'telus.com', ADMIN_EMAIL_DOMAINS: ['telus.com', 'telusinternational.com'],
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k === 'EMAIL_EXTRA_DOMAINS' ? (o.extraDomains || null) : null) }) },
    CacheService: { getScriptCache: () => ({
      get: k => { if (o.cacheBroken) throw new Error('cache down'); return k in store ? store[k] : null; },
      put: (k, v) => { if (o.cacheBroken) throw new Error('cache down'); store[k] = v; }, remove: k => { delete store[k]; } }) },
    getRecipientsFromRoster_: () => [{ email: 'admin@telus.com' }, { email: 'out@elsewhere.com' }],
    buildAuditRecipients_: fd => { buildArgs.push(JSON.parse(JSON.stringify(fd))); return o.recipients || ['agent@telus.com', 'tl@telus.com']; },
    buildAgentEmailHTML: (...a) => { sandbox.htmlArgs = a; return '<html></html>'; },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://example.com/exec' }) },
    sendEmailInBatches_: (list, subject, body, html, sender) => { sent.push({ list, subject, body }); },
    extractTextBlock: () => 'Yes',
    updateDashboardPDFLink_: (...a) => { pdfLinks.push(a); },
    notifyAdmins_: (fd, ref) => { notified.push(ref); },
    updateCachedResult_: (...a) => { cached.push(a); },
    Utilities: { base64Encode: s => Buffer.from(String(s)).toString('base64') },
    logSubmissionError_: (fn, fd, e) => errors.push(fn + ': ' + (e && e.message)),
    Logger: { log: m => logs.push(String(m)) },
  };
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('HtmlSanitizer.gs'), ctx);
  vm.runInContext(read('EmailGuard.gs'), ctx);
  vm.runInContext(extractFunction(read('Config.gs'), 'isAdminCaller_'), ctx);
  vm.runInContext(extractFunction(read('Config.gs'), '_hashKey_'), ctx);
  vm.runInContext(extractFunction(codeSrc, 'sendAuditEmail'), ctx);
  return { ctx, sent, cached, pdfLinks, logs, errors, notified, who, obs, store, buildArgs, sandbox };
}
const FORM = o => Object.assign({ auditRef: 'NHA-20261013-0001', participant: 'Maria Santos', sapId: '1234567', interactionId: 'INT-1', analysisType: 'repeats', emailMode: 'live',
  teamLeader: 'TL', opsManager: 'OM', observerName: 'Analyst One' }, o || {});
const send = (e, f) => e.ctx.sendAuditEmail(f || FORM(), '<div class="report-wrap">report</div>');

console.log('\nTest 1 — only an admin, or the analyst who saved the audit, may email it');
{
  const e = makeEnv({ email: 'someone.else@telus.com', observerName: 'Someone Else' });
  const r = send(e);
  check('a different analyst is refused with a plain message', r.success === false && /Only the analyst who saved this audit, or an admin/.test(r.error), JSON.stringify(r));
  check('nothing was mailed, cached, recorded or announced', e.sent.length === 0 && e.cached.length === 0 && e.pdfLinks.length === 0 && e.notified.length === 0);
  check('the attempt is written to the error log', e.errors.some(x => /notYourAudit/.test(x)));

  const ok1 = makeEnv({});
  check('the analyst whose NAME is on the audit may send it', send(ok1).success === true && ok1.sent.length === 1);
  const ok2 = makeEnv({ observerName: '', email: 'analyst.one@telus.com', rows: [{ 'Audit Ref': 'NHA-20261013-0001', 'Observer Name': 'analyst.one' }] });
  check('the page also records the part of the email before the @, and that matches too', send(ok2).success === true);
  const ok3 = makeEnv({ observerName: '  analyst   ONE ' });
  check('capitals and spacing do not matter', send(ok3).success === true);
  const ok4 = makeEnv({ email: 'admin1@telus.com', observerName: 'Admin Person' });
  check('an admin may email any audit', send(ok4).success === true);
  const off = makeEnv({ email: 'admin1@evil.com', observerName: 'Admin Person' });
  check('an admin name on a foreign domain is just another user', send(off).success === false);
  const blank = makeEnv({ rows: [{ 'Audit Ref': 'NHA-20261013-0001', 'Observer Name': '' }] });
  check('an audit with no observer on it can only be emailed by an admin', send(blank).success === false);
  const noId = makeEnv({ email: '', observerName: '' });
  check('a caller whose identity cannot be read is refused', send(noId).success === false && noId.sent.length === 0);
}

console.log('\nTest 2 — THE ATTACK: any ref, any recipient, any text');
{
  const e = makeEnv({ email: 'attacker@telus.com', observerName: 'Attacker', recipients: ['victim@gmail.com', 'v2@evil.com'] });
  const r = e.ctx.sendAuditEmail({ auditRef: 'NHA-20261013-0001', emailMode: 'live', participant: 'Your account is locked', observerEmail: 'victim@gmail.com', agentEmail: 'v2@evil.com', customerBAN: 'click http://evil' }, '<div>anything</div>');
  check('refused: the attacker is not the observer of that audit', r.success === false && e.sent.length === 0);
  check('...and the stored evaluation of that audit is not overwritten', e.cached.length === 0);

  const own = makeEnv({ recipients: ['agent@telus.com', 'victim@gmail.com', 'v2@evil.com'] });
  const r2 = send(own, FORM({ observerEmail: 'victim@gmail.com', agentEmail: 'v2@evil.com' }));
  check('even the real observer cannot mail outside addresses: only company addresses are sent to', r2.success === true && own.sent[0].list.join() === 'agent@telus.com', JSON.stringify(own.sent));
  check('the dropped addresses are reported back and logged', r2.droppedRecipients && r2.droppedRecipients.join() === 'victim@gmail.com,v2@evil.com' && own.errors.some(x => /externalRecipient/.test(x)));
  const allOut = makeEnv({ recipients: ['victim@gmail.com'] });
  const r3 = send(allOut);
  check('if every recipient is outside the company nothing is sent and the message says why', r3.success === false && /None of the recipients is a company address/.test(r3.error) && allOut.sent.length === 0, JSON.stringify(r3));
  const tricky = makeEnv({ recipients: ['a@telus.com.evil.com', 'b@evil-telus.com', 'c@@telus.com', 'd@TELUS.COM', 'e@telusinternational.com'] });
  const r4 = send(tricky);
  check('look-alike domains are refused, real ones accepted in any case', tricky.sent[0].list.join() === 'd@TELUS.COM,e@telusinternational.com', JSON.stringify(tricky.sent));
  const extra = makeEnv({ extraDomains: 'contractor.example.com, bad domain', recipients: ['x@contractor.example.com', 'y@other.com'] });
  send(extra);
  check('an admin can allow one more domain through the Script Property EMAIL_EXTRA_DOMAINS', extra.sent[0].list.join() === 'x@contractor.example.com', JSON.stringify(extra.sent));
  const adminTest = makeEnv({ email: 'admin1@telus.com', observerName: 'Admin', recipients: [] });
  send(adminTest, FORM({ emailMode: 'test' }));
  check('TEST mode (the Admin/Dev roster list) is filtered too', adminTest.sent.length === 1 && adminTest.sent[0].list.join() === 'admin@telus.com', JSON.stringify(adminTest.sent));
}

console.log('\nTest 3 — the QA recipient is the signed-in person');
{
  const e = makeEnv({});
  send(e, FORM({ observerEmail: 'someone.else@gmail.com' }));
  check('an analyst\'s typed-in observerEmail is replaced by their own address', e.buildArgs[0].observerEmail === 'analyst.one@telus.com', JSON.stringify(e.buildArgs[0].observerEmail));
  const a = makeEnv({ email: 'admin1@telus.com', observerName: 'Admin' });
  send(a, FORM({ observerEmail: 'x@gmail.com' }));
  check('for an admin the typed-in address is cleared and the roster resolves the observer by the name on the audit', a.buildArgs[0].observerEmail === '' && a.buildArgs[0].observerName === 'Analyst One', JSON.stringify(a.buildArgs[0]));
}

console.log('\nTest 4 — a limit per person per hour');
{
  const e = makeEnv({});
  let ok = 0; for (let i = 0; i < 30; i++) if (send(e).success) ok++;
  check('an analyst may send 30 in an hour', ok === 30, String(ok));
  const r = send(e);
  check('the 31st is refused with a clear message and sends nothing', r.success === false && /a lot of audit emails/.test(r.error) && e.sent.length === 30, JSON.stringify(r));
  e.who.email = 'other.analyst@telus.com'; e.obs.name = 'Analyst One';
  check('the limit is per person', send(e).success === true);
  const a = makeEnv({ email: 'admin1@telus.com', observerName: 'Admin' });
  let okA = 0; for (let i = 0; i < 100; i++) if (send(a).success) okA++;
  check('an admin may send 100', okA === 100 && send(a).success === false);
  const b = makeEnv({ cacheBroken: true });
  check('if the cache is down the limit is not applied, so real audit emails are not stopped', send(b).success === true && b.logs.some(l => /limit not applied/.test(l)));
  const long = makeEnv({}); long.who.email = ('a'.repeat(60) + '1@telus.com'); long.obs.name = 'Analyst One';
  const long2 = ('a'.repeat(60) + '2@telus.com');
  send(long); long.who.email = long2; send(long);
  check('two long addresses with the same first characters have separate counters', Object.keys(long.store).length === 2, JSON.stringify(Object.keys(long.store)));
}

console.log('\nTest 5 — text from the form is one short line, and the saved audit wins');
{
  const e = makeEnv({});
  send(e, FORM({ participant: 'x\r\nBcc: evil@x.com', observerName: 'Analyst\nOne', customerBAN: 'B'.repeat(500), direction: 'D\r\nX', startTime: 'S'.repeat(300), locale: 'L\nM', auditType: 'T\r\nU', teamLeader: 'TL' }));
  const m = e.sent[0];
  check('no line break can get into the subject', !/[\r\n]/.test(m.subject), JSON.stringify(m.subject));
  check('the agent name on the SAVED audit is used, not what the form says', m.subject.indexOf('Maria Santos') === 0 && m.subject.indexOf('evil') === -1, m.subject);
  check('a very long BAN and date are cut in the plain-text body', !/B{50}/.test(m.body) && !/S{50}/.test(m.body) && /Customer BAN:\s+B{1,40}\n/.test(m.body));
  check('in the HTML body too', e.sandbox.htmlArgs.every(x => String(x).length <= 130), e.sandbox.htmlArgs.map(x => String(x).length).join());
  check('the SAP ID and Interaction ID come from the saved audit', /Interaction ID:   INT-1\n/.test(m.body) && e.sandbox.htmlArgs[3] === '1234567', e.sandbox.htmlArgs[3]);
  check('the audit ref in the mail is the one that was checked', e.sandbox.htmlArgs[4] === 'NHA-20261013-0001');
  const blank = makeEnv({});
  send(blank, FORM({ sapId: '' }));
  check('with the SAP ID blanked on the form, the recipients are still looked up for the agent of the SAVED audit', blank.buildArgs.length === 1 && blank.buildArgs[0].sapId === '1234567', JSON.stringify(blank.buildArgs.map(x => x.sapId)));
}

console.log('\nTest 6 — the helpers');
{
  const e = makeEnv({});
  const c = e.ctx;
  check('_isInternalEmail_: company addresses yes, look-alikes and junk no', c._isInternalEmail_('a@telus.com') && c._isInternalEmail_('A@TELUSINTERNATIONAL.COM') && !c._isInternalEmail_('a@telus.com.evil.com') && !c._isInternalEmail_('a@gmail.com') && !c._isInternalEmail_('') && !c._isInternalEmail_(null) && !c._isInternalEmail_('telus.com') && !c._isInternalEmail_('a b@telus.com'));
  check('_callerOwnsAudit_ is false for a missing row or a blank observer', !c._callerOwnsAudit_(null, { name: 'A' }, 'a@telus.com') && !c._callerOwnsAudit_({ 'Observer Name': '' }, { name: '' }, ''));
  check('_oneLine_ handles null, numbers and the empty string', c._oneLine_(null, 5) === '' && c._oneLine_(12345, 3) === '123' && c._oneLine_('  a   b  ', 10) === 'a b');
  check('_oneLine_ removes control characters, not only line breaks (a header cannot be smuggled in)', c._oneLine_('a\u0000b\u0007c\u001Bd\r\nBcc: x@evil.com', 60) === 'a b c d Bcc: x@evil.com', JSON.stringify(c._oneLine_('a\u0000b\u0007c\u001Bd\r\nBcc: x@evil.com', 60)));
  check('_hashKey_ differs for different addresses and is stable', c._hashKey_('a@telus.com') !== c._hashKey_('b@telus.com') && c._hashKey_('a@telus.com') === c._hashKey_('a@telus.com'));
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
