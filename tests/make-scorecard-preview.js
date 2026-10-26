/**
 * Writes an HTML preview of sample scorecard slides (made-up agents) to the path given, for looking at in a browser.
 * Not a test; the checks that keep the look safe are in tests/scorecard-metrics-test.js.
 *
 * Usage: node tests/make-scorecard-preview.js <output.html>
 */
const fs = require('fs');
const vm = require('vm');
const { read } = require('./lib/gas-fakes');
const { renderDeckHtml } = require('./lib/scorecard-preview');

const ctx = vm.createContext({ Logger: { log() {} }, JSON, Array, Object, String, Number, Math, Date, RegExp, isFinite, isNaN, Error,
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) } });
['Config.gs', 'ScorecardMetrics.gs', 'ScorecardLayout.gs'].forEach(f => vm.runInContext(read(f), ctx, { filename: f }));

let n = 0;
const row = o => Object.assign({ 'Audit Ref': 'S' + (++n), 'Timestamp': '2026-10-0' + (1 + n % 9) + 'T10:00:00Z', 'Team Member': 'Sample Agent', 'SAP ID': '100001',
  'Analysis Type': 'Repeats & Transfer Analyzer', 'Repeat Projection %': '20%', 'Issue Resolved': 'Yes', 'Critical Flags': '', 'Transfer Occurred': 'No',
  'Sales Attempted': '', 'Email Status': 'Sent', 'Team Leader': 'Taylor Brooks', 'Line of Business': 'MOB CLS', 'Locale': 'TI Philippines', 'VTID': 'VTI2026000101',
  'Top SMART Opportunity': '', 'RCA Sub-Parameter': '' }, o);

const rows = [];
// 1. everything on target
for (let i = 0; i < 6; i++) rows.push(row({ 'Team Member': 'Maria Santos', 'SAP ID': '100001', 'Repeat Projection %': ['10%', '0.15', '25', '20%', '12%', '28%'][i], 'Issue Resolved': i === 5 ? 'Partially resolved - the credit has not posted yet' : 'Yes' }));
// 2. one tile below (needs attention)
for (let i = 0; i < 5; i++) rows.push(row({ 'Team Member': 'Aisha Khan', 'SAP ID': '100002', 'Repeat Projection %': ['15%', '22', '0.18', '25%', '30%'][i], 'Issue Resolved': 'Yes',
  'Critical Flags': i === 2 ? 'No expectation set for the callback' : '', 'Top SMART Opportunity': i === 2 ? 'Confirm the callback time and channel before closing the call' : '' }));
// 3. priority support, with the longest realistic text
const long = 'Agent did not confirm the customer\'s understanding before closing and moved to the next step without a summary of what was agreed';
for (let i = 0; i < 6; i++) rows.push(row({ 'Team Member': 'Jonathan Reyes-Villanueva', 'SAP ID': '100003', 'Repeat Projection %': ['85% repeat risk. The call ended without a confirmed fix', '0.72', '65', '90%', '55%', '78'][i],
  'Issue Resolved': ['No', 'Partially resolved - callback promised', 'No', 'No', 'Yes', 'Issue was NOT resolved. The credit was never applied'][i],
  'Critical Flags': ['No needs discovery before offer | Generic pricing explanation | Transfer without warm handoff', 'No needs discovery before offer | Missed callback policy', 'Transfer without warm handoff', 'No needs discovery before offer', 'Generic pricing explanation', 'No expectation set after the appointment'][i],
  'Transfer Occurred': ['Yes', 'Yes', 'No', 'Yes', 'No', 'Yes'][i],
  'Top SMART Opportunity': [long, 'Use customer cues to personalize the offer before presenting any price', long, 'Set a clear expectation for the next step and confirm the customer agrees', long, 'Ask one open-ended discovery question at the start of every call'][i] }));
// 4. a Sales agent
for (let i = 0; i < 5; i++) rows.push(row({ 'Team Member': 'Carlos Mendoza', 'SAP ID': '100004', 'Analysis Type': 'Sales Analyzer', 'Repeat Projection %': '', 'Issue Resolved': 'Yes',
  'Sales Attempted': ['Yes', 'No', 'Yes', 'Yes', 'N/A'][i], 'Critical Flags': i % 2 ? 'No retention save attempt after the customer asked to cancel' : 'Offer made before needs discovery', 'Line of Business': 'MOB CxSS',
  'Top SMART Opportunity': 'Ask a needs-based question before presenting any offer or price' }));

const extras = { whatsWorking: {} };
rows.forEach(r => {
  if (r['Team Member'] === 'Maria Santos') extras.whatsWorking[r['Audit Ref']] = ['Warm greeting and clear ownership of the customer\'s issue', 'Accurate explanation of the new bill'];
  if (r['Team Member'] === 'Aisha Khan') extras.whatsWorking[r['Audit Ref']] = ['Verified the account before making any change'];
  if (r['Team Member'] === 'Jonathan Reyes-Villanueva') extras.whatsWorking[r['Audit Ref']] = ['Stayed calm and polite with a frustrated customer'];
  if (r['Team Member'] === 'Carlos Mendoza') extras.whatsWorking[r['Audit Ref']] = ['Clear explanation of the plan benefits'];
});
const built = ctx.buildAgentCards_(rows, extras);
// the whole deck, exactly as the real builder assembles it (cover, group slides, one slide per agent, next steps)
const slides = ctx.scorecardDeckShapes_(built, 'MOB CLS Sample Deck', 'Team Leader: Taylor Brooks; Locale: TI Philippines', built.total, []);
const titles = slides.map((shapes, i) => {
  const t = shapes.find(s => s.id === 'title') || shapes.find(s => s.id === 'name');
  return (i + 1) + '. ' + (t ? t.paras[0].runs.map(r => r.t).join('') : 'Slide');
});

module.exports = { slides, titles, ctx };   // other scripts (for example one that takes a picture of each slide) can reuse the sample deck

if (require.main === module) {
  const out = process.argv[2];
  if (!out) { console.error('Usage: node tests/make-scorecard-preview.js <output.html>'); process.exit(1); }
  fs.writeFileSync(out, renderDeckHtml(slides, ctx.SCORECARD_THEME, ctx.SCORECARD_PAGE, 1.6, titles));
  console.log('wrote ' + out + ' (' + slides.length + ' slides)');
}
