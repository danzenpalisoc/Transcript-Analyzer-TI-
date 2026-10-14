/**
 * diagnoseAnalyticsContext(): read-only, admin only. Shows what the AI summaries are told, from the real sheets, so a dashboard text that
 * disagrees with the tiles can be traced to old code or to different numbers.
 *
 * Usage: node tests/diagnose-context-test.js
 */
const vm = require('vm');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

function makeEnv(logRows, dashRows) {
  const who = { email: 'danzen.palisoc@telus.com' }, logs = [];
  const sandbox = { Logger: { log: m => logs.push(String(m)) }, JSON, Math, String, Array, Object, Number, parseFloat, isNaN, Error,
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Session: { getActiveUser: () => ({ getEmail: () => who.email }), getScriptTimeZone: () => 'UTC' } };
  const ctx = vm.createContext(sandbox);
  ['Config.gs', 'ScorecardMetrics.gs', 'DataRepair.gs'].forEach(f => vm.runInContext(read(f), ctx));
  ['dbUniqServer', 'buildAnalyticsContext_'].forEach(n => vm.runInContext(extractFunction(read('Code.gs'), n), ctx));
  let writes = 0;
  ctx.getOrCreateSpreadsheet = () => ({ setName() { writes++; } });
  ctx.readAuditLog_ = () => logRows; ctx.readDashboardSheet_ = () => dashRows;
  return { ctx, who, logs, writes: () => writes };
}
const L = (ref, ir) => ({ 'Audit Ref': ref, 'Team Member': 'A', 'Issue Resolved': ir, 'Repeat Projection %': '40%' });
const D = (ref, ir) => ({ 'Audit Ref': ref, 'Issue Resolved': ir, 'Critical Flags': '' });

console.log('\nTest 1 — it reports the counts three ways and the context lines');
{
  const log = [L('R1', 'Yes'), L('R2', 'Partially resolved - yes the credit posted'), L('R3', 'Issue resolved — fixed'), L('R4', 'No'), L('R5', ''), L('R6', '')];
  const dash = [D('R5', 'The issue was resolved after the reset'), D('R6', '')];
  const e = makeEnv(log, dash);
  const r = e.ctx.diagnoseAnalyticsContext();
  check('Audit_Log alone: Resolved 2, Partially 1, Unresolved 1, Not recorded 2', r.auditLog.yes === 2 && r.auditLog.partial === 1 && r.auditLog.no === 1 && r.auditLog.unknown === 2, JSON.stringify(r.auditLog));
  check('as the dashboard page builds the rows (the Dashboard_Data answer fills a blank): Resolved 3, Not recorded 1', r.merged.yes === 3 && r.merged.unknown === 1, JSON.stringify(r.merged));
  check('the OLD way (the letters yes anywhere) is shown for comparison: 2 (Yes, and the Partially one that contains yes)', r.oldWay === 2, String(r.oldWay));
  check('the context lines the AI is given are shown', r.contextLines.some(l => /Issues Resolved: 3 of 6/.test(l)) && r.contextLines.some(l => /Resolved 3, Partially resolved 1, Unresolved 1, Not recorded 1/.test(l)) && r.contextLines.some(l => /Average Repeat Risk: 40%/.test(l)), JSON.stringify(r.contextLines));
  check('the log ends with the line that says to copy it', e.logs[e.logs.length - 1] === '=== end: copy everything above into the chat ===');
  check('it writes nothing', e.writes() === 0);
}

console.log('\nTest 2 — only an admin');
{
  const e = makeEnv([L('R1', 'Yes')], []); e.who.email = 'some.analyst@telus.com';
  let err = ''; try { e.ctx.diagnoseAnalyticsContext(); } catch (x) { err = x.message; }
  check('an analyst is refused', /admin/i.test(err), err);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
