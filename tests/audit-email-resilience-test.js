/**
 * Regression test for sendAuditEmail() (Code.gs).
 *
 * Reported symptom: analysts saw "Email failed: You do not have permission to
 * access the requested document." even though the email itself was actually
 * delivered. Two separate gaps caused this:
 *   1. updateDashboardPDFLink_() ran unguarded AFTER the email had already
 *      been sent — a transient failure there made the whole request report
 *      as failed even though the email went out fine.
 *   2. The outer catch returned the raw Google error string instead of an
 *      actionable message (submitTranscript already had this treatment;
 *      sendAuditEmail did not).
 *
 * This loads the REAL sendAuditEmail() out of Code.gs into a sandbox with
 * every dependency stubbed, so the exact production logic is exercised.
 *
 * Usage: node tests/audit-email-resilience-test.js [path-to-Code.gs]
 * Exits non-zero if any check fails.
 */
const fs = require('fs');
const vm = require('vm');
const path = require('path');

const gsPath = process.argv[2] || path.join(__dirname, '..', 'Code.gs');
const src = fs.readFileSync(gsPath, 'utf8');

function extractFunction(source, name) {
  const startMatch = source.match(new RegExp('function\\s+' + name + '\\s*\\('));
  if (!startMatch) throw new Error('function ' + name + ' not found in ' + gsPath);
  let i = startMatch.index;
  const braceStart = source.indexOf('{', i);
  let depth = 0, end = -1;
  for (let j = braceStart; j < source.length; j++) {
    if (source[j] === '{') depth++;
    else if (source[j] === '}') { depth--; if (depth === 0) { end = j + 1; break; } }
  }
  if (end === -1) throw new Error('could not find end of function ' + name);
  return source.slice(i, end);
}

const fnSource = extractFunction(src, 'sendAuditEmail');

function run(overrides) {
  const logs = [];
  const errorLogCalls = [];
  const defaults = {
    readAuditLog_: () => [{ 'Audit Ref': 'NHA-20260929-0001', 'Observer Name': 'Analyst One' }],
    resolveObserver_: () => ({ name: 'Analyst One', email: 'analyst@telus.com' }),
    Session: { getActiveUser: () => ({ getEmail: () => 'analyst@telus.com' }) },
    ADMIN_USERNAMES: ['admin1'],
    ADMIN_DOMAIN: 'telus.com',
    ADMIN_EMAIL_DOMAINS: ['telus.com', 'telusinternational.com'],
    getRecipientsFromRoster_: () => [{ email: 'admin@telus.com' }],
    buildAuditRecipients_: () => ['tl@telus.com', 'om@telus.com'],
    buildAgentEmailHTML: () => '<html></html>',
    ScriptApp: { getService: () => ({ getUrl: () => 'https://example.com/exec' }) },
    sendEmailInBatches_: () => {},      // succeeds by default
    extractTextBlock: () => 'Yes',
    updateDashboardPDFLink_: () => {},   // succeeds by default
    notifyAdmins_: () => {},
    updateCachedResult_: () => {},
    CacheService: { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) },
    Utilities: { base64Encode: s => Buffer.from(String(s)).toString('base64') },
    logSubmissionError_: (fnName, formData, e) => { errorLogCalls.push({ fnName, err: e && e.toString() }); },
    Logger: { log: (m) => logs.push(m) },
  };
  const sandbox = Object.assign({}, defaults, overrides);
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'HtmlSanitizer.gs'), 'utf8'), ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'EmailGuard.gs'), 'utf8'), ctx);
  // the real admin rule, so sendAuditEmail is tested against it and not a stub
  vm.runInContext(extractFunction(fs.readFileSync(path.join(__dirname, '..', 'Config.gs'), 'utf8'), 'isAdminCaller_'), ctx);
  vm.runInContext(fnSource, ctx);
  return { sendAuditEmail: ctx.sendAuditEmail, logs, errorLogCalls };
}

const BASE_FORM = {
  auditRef: 'NHA-20260929-0001', participant: 'Juan Dela Cruz',
  teamLeader: 'TL Person', opsManager: 'OM Person', interactionId: 'INT-1',
  analysisType: 'repeats', sapId: '1234567', startTime: '2026-09-29',
};

let failures = 0;
function check(name, cond, detail) {
  console.log('  ' + (cond ? 'PASS' : 'FAIL') + '  ' + name);
  if (!cond) { failures++; if (detail) console.log('          ' + detail); }
}

console.log('NH Analyzer — sendAuditEmail resilience test');
console.log('file: ' + gsPath);

console.log('\nTest 1 — happy path still returns success');
{
  const { sendAuditEmail } = run({});
  const result = sendAuditEmail(BASE_FORM, '<html>report</html>');
  check('success === true', result.success === true);
  check('recipients returned', Array.isArray(result.recipients) && result.recipients.length > 0);
}

console.log('\nTest 2 — THE BUG: dashboard-status write fails AFTER a successful send');
{
  const { sendAuditEmail, logs, errorLogCalls } = run({
    updateDashboardPDFLink_: () => {
      throw new Error('Exception: You do not have permission to access the requested document.');
    }
  });
  const result = sendAuditEmail(BASE_FORM, '<html>report</html>');
  check('email is still reported as sent (success === true)', result.success === true,
        'got ' + JSON.stringify(result));
  check('recipients still returned', Array.isArray(result.recipients) && result.recipients.length > 0);
  check('failure was logged to Logger', logs.some(l => l.indexOf('updateDashboardPDFLink in sendAuditEmail failed') !== -1));
  check('failure was logged to Error_Log', errorLogCalls.some(c => c.fnName === 'sendAuditEmail:updateDashboardPDFLink'));
}

console.log('\nTest 3 — a genuine pre-send failure (permission-shaped) gets a friendly message');
{
  const { sendAuditEmail, errorLogCalls } = run({
    readAuditLog_: () => { throw new Error('Exception: You do not have permission to access the requested document.'); }
  });
  const result = sendAuditEmail(BASE_FORM, '<html>report</html>');
  check('success === false', result.success === false);
  check('error message is friendly, not raw', result.error.indexOf('temporary system error') !== -1,
        'got ' + JSON.stringify(result.error));
  check('raw Google text not shown to analyst', result.error.indexOf('do not have permission') === -1);
  check('logged to Error_Log', errorLogCalls.some(c => c.fnName === 'sendAuditEmail'));
}

console.log('\nTest 4 — a non-permission failure still returns its own message unchanged');
{
  const { sendAuditEmail } = run({
    readAuditLog_: () => { throw new Error('Some totally different failure'); }
  });
  const result = sendAuditEmail(BASE_FORM, '<html>report</html>');
  check('success === false', result.success === false);
  check('raw message preserved for non-permission errors', result.error.indexOf('Some totally different failure') !== -1,
        'got ' + JSON.stringify(result.error));
}

console.log('\nTest 5 — quota-exceeded during send is still a warning, not a hard failure (regression)');
{
  const { sendAuditEmail } = run({
    sendEmailInBatches_: () => { throw new Error('QUOTA_EXCEEDED: sendEmailInBatches_: 1 of 1 batch(es) failed.'); }
  });
  const result = sendAuditEmail(BASE_FORM, '<html>report</html>');
  check('success === true (audit still saved)', result.success === true);
  check('emailWarning is set', !!result.emailWarning);
}

console.log('\nTest 6 — TEST mode is honoured for an admin, on either TELUS domain, and never silently becomes LIVE');
{
  const asUser = email => ({ Session: { getActiveUser: () => ({ getEmail: () => email }) } });
  const TEST_FORM = Object.assign({}, BASE_FORM, { emailMode: 'test' });
  const LIVE_FORM = Object.assign({}, BASE_FORM, { emailMode: 'live' });
  const who = r => (r.recipients || []).join();

  const a = run(asUser('admin1@telus.com')).sendAuditEmail(TEST_FORM, '<p>x</p>');
  check('admin on telus.com + TEST -> only the Admin/Dev list', a.success === true && who(a) === 'admin@telus.com', JSON.stringify(a));

  const b = run(asUser('admin1@telusinternational.com')).sendAuditEmail(TEST_FORM, '<p>x</p>');
  check('admin on telusinternational.com + TEST -> still TEST, not LIVE', b.success === true && who(b) === 'admin@telus.com', JSON.stringify(b));

  let live = 0;
  const c = run(Object.assign(asUser('analyst@telus.com'), { buildAuditRecipients_: () => { live++; return ['tl@telus.com']; } })).sendAuditEmail(TEST_FORM, '<p>x</p>');
  check('a non-admin asking for TEST is REFUSED', c.success === false && /only for the admin team/i.test(c.error || ''), JSON.stringify(c));
  check('...and nobody real was looked up or emailed', live === 0, 'buildAuditRecipients_ ran ' + live + ' time(s)');

  const d = run(Object.assign(asUser(''), { buildAuditRecipients_: () => { live++; return ['tl@telus.com']; } })).sendAuditEmail(TEST_FORM, '<p>x</p>');
  check('a hidden/blank email asking for TEST is REFUSED, not treated as the owner', d.success === false && live === 0, JSON.stringify(d));

  const e = run(asUser('admin1@evil.com')).sendAuditEmail(TEST_FORM, '<p>x</p>');
  check('an admin name on a foreign domain asking for TEST is REFUSED', e.success === false);

  const f = run(asUser('analyst@telus.com')).sendAuditEmail(LIVE_FORM, '<p>x</p>');
  check('an ordinary analyst on LIVE still works and reaches the real recipients', f.success === true && who(f) === 'tl@telus.com,om@telus.com', JSON.stringify(f));

  const g = run(asUser('admin1@telus.com')).sendAuditEmail(LIVE_FORM, '<p>x</p>');
  check('an admin who chose LIVE gets LIVE', g.success === true && who(g) === 'tl@telus.com,om@telus.com', JSON.stringify(g));

  const h = run(asUser('analyst@telus.com')).sendAuditEmail(BASE_FORM, '<p>x</p>');
  check('no mode given defaults to LIVE', h.success === true && who(h) === 'tl@telus.com,om@telus.com', JSON.stringify(h));
}

console.log('\nTest 7 — the screen and the server use the same admin rule');
{
  const codeSrc = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
  const obs = extractFunction(codeSrc, 'getObserverInfo');
  check('getObserverInfo decides isAdmin with isAdminCaller_()', /obs\.isAdmin\s*=\s*isAdminCaller_\(\)/.test(obs));
  check('...and no longer compares only the local part of a possibly-owner email', !/ADMIN_USERNAMES\.indexOf\(userLocal\)/.test(obs));
  check('sendAuditEmail uses isAdminCaller_() too', /callerIsAdmin\s*=\s*isAdminCaller_\(\)/.test(fnSource));
  check('...and no longer requires exactly ADMIN_DOMAIN', !/callerDomain\s*===\s*ADMIN_DOMAIN/.test(fnSource));
}

console.log(failures === 0 ? '\nALL TESTS PASSED' : '\n' + failures + ' CHECK(S) FAILED');
process.exit(failures === 0 ? 0 : 1);
