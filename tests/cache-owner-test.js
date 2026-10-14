/**
 * Regression test for who OWNS a cached evaluation.
 *
 * Cache rows written before the SAP ID column existed have a blank owner.
 * They belong to the agent whose audit CREATED them, which is the EARLIEST
 * Audit_Log row for that Interaction ID and analysis type. Treating a blank
 * row as "mine" for any agent who merely has an audit let a second agent's
 * save delete the first agent's only evaluation, and served it to the second.
 *
 * Also covers: the cache layer must not read other rows' HTML just to compare
 * IDs, a failed CacheService write must not leave the old HTML to be served,
 * and sendAuditEmail must use the SAVED audit's analysis type and refuse a
 * reference that two audits share.
 *
 * Usage: node tests/cache-owner-test.js
 * Exits non-zero if any check fails.
 */
const vm = require('vm');
const { read, extractFunction, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const codeSrc = read('Code.gs');

function makeEnv(opts) {
  opts = opts || {};
  const store = {};
  const cs = {
    get: k => (k in store ? store[k] : null),
    put: (k, v) => { if (opts.failPuts && opts.failPuts()) throw new Error('CacheService put failed'); store[k] = v; },
    remove: k => { delete store[k]; },
  };
  const sandbox = {
    Logger: { log() {} },
    CacheService: { getScriptCache: () => cs },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Utilities: { base64Encode: s => Buffer.from(String(s)).toString('base64'), sleep() {} },
    ScriptApp: { getScriptId: () => 'x', getService: () => ({ getUrl: () => 'https://example.com/exec' }) },
    sharedCSS: () => '', fixBadgeClasses: h => h, invalidateDashboardCache_: () => {},
  };
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('Config.gs'), ctx);
  vm.runInContext(read('HtmlSanitizer.gs'), ctx);
  vm.runInContext(read('EmailGuard.gs'), ctx);
  vm.runInContext(read('SheetService.gs'), ctx);
  vm.runInContext(extractFunction(codeSrc, 'sendAuditEmail'), ctx);
  const sheets = {};
  sheets[ctx.AUDIT_LOG_SHEET] = new FakeSheet(ctx.AUDIT_LOG_HEADERS, []);
  sheets[ctx.CACHE_SHEET] = new FakeSheet(ctx.CACHE_HEADERS, []);
  const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new FakeSheet([], [])), getUrl: () => 'u' };
  ctx.getOrCreateSpreadsheet = () => ss;
  return { ctx, sheets, store, audit: sheets[ctx.AUDIT_LOG_SHEET], cache: sheets[ctx.CACHE_SHEET] };
}
function addAudit(env, ref, sap, intId, type) {
  env.audit.appendRow(env.ctx.AUDIT_LOG_HEADERS.map(h => ({
    'Audit Ref': ref, 'Interaction ID': intId, 'SAP ID': sap,
    'Analysis Type': type || 'Repeats & Transfer Analyzer',
  }[h] || '')));
}

const INT = 'interaction-1', SAP_A = '373106', SAP_B = '2008630';
const HTML_A = '<div>A</div>', HTML_B = '<div>B</div>';
const legacyRow = [INT, 'repeats', new Date(), HTML_A];           // 4 columns: written before column E existed

console.log('NH Analyzer — cache ownership test');

console.log('\nTest 1 — a second agent saving or editing never deletes the first agent\'s legacy evaluation');
{
  const env = makeEnv();
  env.cache.appendRow(legacyRow);
  addAudit(env, 'NHA-1', SAP_A, INT);     // A's audit created the legacy row
  addAudit(env, 'NHA-2', SAP_B, INT);     // B was audited later
  env.ctx.updateCachedResult_(INT, 'repeats', '<div>B edited</div>', SAP_B);
  env.store && Object.keys(env.store).forEach(k => delete env.store[k]);
  check('agent A still gets A\'s evaluation after B saved', env.ctx.findCachedResult_(INT, 'repeats', SAP_A) === HTML_A,
        'got ' + JSON.stringify(env.ctx.findCachedResult_(INT, 'repeats', SAP_A)));
  check('agent B gets B\'s own saved evaluation', env.ctx.findCachedResult_(INT, 'repeats', SAP_B) === '<div>B edited</div>');
}

console.log('\nTest 2 — a legacy row is served only to the agent whose audit created it');
{
  const env = makeEnv();
  env.cache.appendRow(legacyRow);
  addAudit(env, 'NHA-1', SAP_A, INT);
  addAudit(env, 'NHA-2', SAP_B, INT);
  check('the earliest audit\'s agent gets the legacy row', env.ctx.findCachedResult_(INT, 'repeats', SAP_A) === HTML_A);
  check('a LATER audit\'s agent does NOT, even though they have an audit',
        env.ctx.findCachedResult_(INT, 'repeats', SAP_B) === null,
        'got ' + JSON.stringify(env.ctx.findCachedResult_(INT, 'repeats', SAP_B)));
}

console.log('\nTest 3 — the legacy owner replacing their own evaluation leaves no stale blank-owner row');
{
  const env = makeEnv();
  env.cache.appendRow(legacyRow);
  addAudit(env, 'NHA-1', SAP_A, INT);
  env.ctx.updateCachedResult_(INT, 'repeats', '<div>A edited</div>', SAP_A);
  const blank = env.cache.grid.slice(1).filter(r => r[0] === INT && !(r[4] || '').toString().trim());
  check('no blank-owner row is left behind', blank.length === 0, JSON.stringify(blank.map(r => r[3])));
  check('the owner now gets the edited evaluation', env.ctx.findCachedResult_(INT, 'repeats', SAP_A) === '<div>A edited</div>');
}

console.log('\nTest 4 — the cache layer does not read other rows\' HTML');
{
  const env = makeEnv();
  for (let i = 0; i < 40; i++) env.cache.appendRow(['other-' + i, 'repeats', new Date(), '<div>' + 'x'.repeat(500) + '</div>', '111' + i]);
  env.cache.resetReads();
  env.ctx.saveCachedResult_(INT, 'repeats', HTML_A, SAP_A);
  check('saving reads none of the existing HTML cells (column D)', env.cache.cellsReadInColumn(4) === 0,
        'read ' + env.cache.cellsReadInColumn(4) + ' column-D cells');
  Object.keys(env.store).forEach(k => delete env.store[k]);
  env.cache.resetReads();
  check('lookup still finds the saved evaluation', env.ctx.findCachedResult_(INT, 'repeats', SAP_A) === HTML_A);
  check('lookup reads only the matching row\'s HTML, not the other 40', env.cache.cellsReadInColumn(4) <= 1,
        'read ' + env.cache.cellsReadInColumn(4) + ' column-D cells');
  check('a normal owner-stamped lookup never touches Audit_Log', env.audit.reads.length === 0,
        'Audit_Log reads: ' + env.audit.reads.length);
}

console.log('\nTest 5 — a failed CacheService write must not leave the OLD html to be served');
{
  let failing = false;
  const env = makeEnv({ failPuts: () => failing });
  env.ctx.saveCachedResult_(INT, 'repeats', '<div>old</div>', SAP_A);
  failing = true;
  env.ctx.saveCachedResult_(INT, 'repeats', '<div>new</div>', SAP_A);
  failing = false;
  check('the new evaluation is served, not the old one', env.ctx.findCachedResult_(INT, 'repeats', SAP_A) === '<div>new</div>',
        'got ' + JSON.stringify(env.ctx.findCachedResult_(INT, 'repeats', SAP_A)));
}

console.log('\nTest 6 — sendAuditEmail uses the saved audit\'s analysis type and refuses a shared reference');
{
  function runSend(form, rows) {
    const env = makeEnv();
    let cachedWith = null, sent = 0;
    Object.assign(env.ctx, {
      readAuditLog_: () => rows.map(r => Object.assign({ 'Observer Name': 'QA One' }, r)),
      Session: { getActiveUser: () => ({ getEmail: () => 'analyst@telus.com' }) },
      resolveObserver_: () => ({ name: 'QA One', email: 'analyst@telus.com' }),
      ADMIN_USERNAMES: [], ADMIN_DOMAIN: 'telus.com',
      getRecipientsFromRoster_: () => [], buildAuditRecipients_: () => ['agent@telus.com'],
      buildAgentEmailHTML: () => '<html></html>',
      sendEmailInBatches_: () => { sent++; },
      extractTextBlock: () => 'Yes',
      updateDashboardPDFLink_: () => {}, notifyAdmins_: () => {},
      updateCachedResult_: (id, type, html, sap) => { cachedWith = { id, type, sap }; },
      logSubmissionError_: () => {},
    });
    return { result: env.ctx.sendAuditEmail(form, '<html>r</html>'), cachedWith: () => cachedWith, sent: () => sent };
  }
  const rows = [{ 'Audit Ref': 'NHA-1', 'SAP ID': SAP_A, 'Interaction ID': INT, 'Analysis Type': 'Sales Analyzer' }];
  const form = { auditRef: 'NHA-1', participant: 'A', sapId: SAP_A, interactionId: INT, analysisType: 'repeats', startTime: 'x' };

  let r = runSend(form, rows);
  check('the email is sent', r.result.success === true && r.sent() === 1, JSON.stringify(r.result));
  check('the cache is updated under the SAVED type (sales), not the form radio (repeats)',
        r.cachedWith() && r.cachedWith().type === 'sales', JSON.stringify(r.cachedWith()));
  check('the cache is updated for the saved audit\'s agent', r.cachedWith() && r.cachedWith().sap === SAP_A);

  const shared = [
    { 'Audit Ref': 'NHA-1', 'SAP ID': SAP_A, 'Interaction ID': INT, 'Analysis Type': 'Repeats & Transfer Analyzer' },
    { 'Audit Ref': 'NHA-1', 'SAP ID': SAP_B, 'Interaction ID': 'other', 'Analysis Type': 'Repeats & Transfer Analyzer' },
  ];
  r = runSend(Object.assign({}, form, { sapId: SAP_B, interactionId: 'other' }), shared);
  check('a reference shared by two audits is refused, nothing is emailed', r.result.success === false && r.sent() === 0, JSON.stringify(r.result));
  check('the message tells the analyst to contact the admin team', /admin/i.test(r.result.error || ''), JSON.stringify(r.result.error));
  r = runSend(form, shared);
  check('the first audit\'s owner is refused too, because the link would be ambiguous', r.result.success === false && r.sent() === 0,
        JSON.stringify(r.result));
}

console.log(check.state.failures === 0 ? '\nALL TESTS PASSED' : '\n' + check.state.failures + ' CHECK(S) FAILED');
process.exit(check.state.failures === 0 ? 0 : 1);
