/**
 * The text the AI is given for the Highlights and Lowlights and the analytics (buildAnalyticsContext_ in Code.gs).
 *
 * The AI quotes whatever numbers it is given. They must be the numbers the dashboard tiles show, so the page does not say 23 percent in a
 * summary and 25 percent on a tile. Resolved used to be "the cell contains the letters yes anywhere" and Repeat % used parseFloat.
 *
 * Usage: node tests/analytics-context-test.js
 */
const vm = require('vm');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');
const { load } = require('./lib/page-harness');
const check = makeCheck();

const sandbox = { Logger: { log() {} }, JSON, Math, String, Array, Object, Number, parseFloat, isNaN,
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) } };
const ctx = vm.createContext(sandbox);
vm.runInContext(read('Config.gs'), ctx);
vm.runInContext(read('ScorecardMetrics.gs'), ctx);
const code = read('Code.gs');
['dbUniqServer', 'buildAnalyticsContext_'].forEach(n => vm.runInContext(extractFunction(code, n), ctx));
const page = load();

const row = o => Object.assign({ 'Team Member': 'A', 'Critical Flags': '', 'Issue Resolved': '', 'Repeat Projection %': '', 'RCA Category': '', 'Call Driver': '', 'RCA Sub-Parameter': '' }, o);
const line = (txt, label) => { const m = new RegExp('^' + label + '.*$', 'm').exec(txt); return m ? m[0] : ''; };

console.log('\nTest 1 — the counts of Issue Resolved are the dashboard\'s');
{
  const answers = ['Yes', 'Issue resolved — the agent fixed the billing error', 'The issue was resolved after the reset', 'Primary issue resolved after the second reset',            // 4 resolved
    'Partially resolved - yes the credit posted but not the refund', 'Issue was partially resolved; credit pending',                   // 2 partial (the first contains "yes")
    'No', 'Issue not resolved. Suggestions were limited to', 'Not resolved - yes, a follow-up call was promised', 'Neither issue was resolved on this call', // 4 unresolved
    '', 'Resolution is ambiguous. Interaction summary', 'Technically resolved — but the customer will call again'];                       // 3 not recorded
  const rows = answers.map(a => row({ 'Issue Resolved': a }));
  const txt = ctx.buildAnalyticsContext_(rows, {}, 'repeats');
  check('Issues Resolved: 4 of 13 (31%), counting a sentence that starts with its verdict and NOT a cell that merely contains "yes"', /- Issues Resolved: 4 of 13 \(31%\)/.test(txt), line(txt, '- Issues Resolved'));
  check('all the answers are given: Resolved 4, Partially resolved 2, Unresolved 4, Not recorded 3', /Resolved 4, Partially resolved 2, Unresolved 4, Not recorded 3/.test(txt), line(txt, '- Issue Resolved, all'));
  const dash = answers.reduce((n, a) => n + (page.ctx._dbResolution(a) === 'yes' ? 1 : 0), 0);
  check('the dashboard (index.html) counts the same 4 resolved', dash === 4, String(dash));
}

console.log('\nTest 2 — the Repeat % average is the dashboard\'s');
{
  const vals = ['40%', '0.6', '95% repeat risk. No issue was identified', '', null, 'about forty', '0%'];
  const txt = ctx.buildAnalyticsContext_(vals.map(v => row({ 'Repeat Projection %': v })), {}, 'repeats');
  const dashVals = vals.map(v => page.ctx._dbNum(v)).filter(x => x !== null);
  const want = Math.round(dashVals.reduce((s, x) => s + x, 0) / dashVals.length);
  check('the average of 40, 60, 95 and a real 0 is ' + want + ' percent, as the tiles read them (a fraction, a sentence, a zero)', new RegExp('- Average Repeat Risk: ' + want + '%').test(txt) && want === 49, line(txt, '- Average Repeat Risk') + ' / dashboard ' + want);
  check('a cell with no number does not count', ctx.buildAnalyticsContext_([row({ 'Repeat Projection %': '' })], {}, 'repeats').indexOf('- Average Repeat Risk: 0%') !== -1);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
