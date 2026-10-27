/**
 * The LOB scorecards in Generate Slides (registry 190).
 *
 *  1. The numbers: every FCR (and the MOB CxSS Transfers) threshold, target and stretch of every LOB and quarter, written down again here
 *     from the admin's scorecards, so a changed number in ScorecardTargets.gs fails this test.
 *  2. The 28 spellings of Line of Business in the data and the scorecard each one uses.
 *  3. The quarter: from the newest audit, and the closest quarter when the scorecard has none.
 *  4. The four levels (below, threshold, target, stretch), higher-is-better and lower-is-better.
 *  5. The agent cards: which tiles use the LOB scorecard, which keep the app's target and say so, the badge.
 *  6. The sections: one for each LOB scorecard, "Other LOB" last; without the targets file nothing changes.
 *  7. The group numbers of a section.
 *  8. The deck: a section for each LOB (group slides, its agents, its next steps), the order, the page numbers, the tag, the legends.
 *  9. The look of every slide: inside the page, unique ids, no font under 11 pt, contrast 4.5:1, nothing overflowing.
 *
 * Usage: node tests/scorecard-lob-test.js
 */
const vm = require('vm');
const { read, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

function load(files) {
  const ctx = vm.createContext({ Logger: { log() {} }, JSON, Array, Object, String, Number, Math, Date, RegExp, isFinite, isNaN, Error,
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) } });
  files.forEach(f => vm.runInContext(read(f), ctx, { filename: f }));
  return ctx;
}
const ctx = load(['Config.gs', 'ScorecardMetrics.gs', 'ScorecardTargets.gs', 'ScorecardLayout.gs', 'ScorecardPptx.gs']);
const T = ctx.SCORECARD_THEME, P = ctx.SCORECARD_PAGE;

console.log('\nTest 1 — the numbers of the scorecards');
{
  // LOB quarter thresh target stretch, as read from the screenshots
  const FCR = `PF CxSS 1 74.5 78 81|PF CxSS 2 74.5 78 81|PF CxSS 3 74 78 80.5|PF CxSS 4 77 79.5 82
PF CLS 1 74 78 83|PF CLS 2 74 78 83|PF CLS 3 74 79.5 82|PF CLS 4 76.5 79.5 82
PF TS 1 70 76 79|PF TS 2 70 76 79|PF TS 3 73.5 78 80.5|PF TS 4 76 78.5 80.5
SHS CxSS 1 75 79 84|SHS CxSS 2 75 79 84|SHS CxSS 3 76.5 80.5 83.5|SHS CxSS 4 78.5 81 83.5
SHS CLS 1 75 80 85|SHS CLS 2 75 80 85|SHS CLS 3 77.5 81.5 84|SHS CLS 4 79 81.5 84
SHS TS 1 74.5 78.5 82.5|SHS TS 2 74.5 80 87|SHS TS 3 76.5 81.5 84|SHS TS 4 79 81.5 84
MOB CLS 1 72 77 82|MOB CLS 2 72.5 77.5 82.5|MOB CLS 3 77.5 81 84|MOB CLS 4 78.5 81 84
MOB CxSS 1 77 80 83|MOB CxSS 2 77 80 83|MOB CxSS 3 77.5 81 83.5|MOB CxSS 4 80 82.5 85
One CAM 2 77.5 81 86|One CAM 3 81.5 86 88.5|One CAM 4 81.5 86 88.5`.split(/\||\n/).map(s => s.trim()).filter(Boolean);
  const got = [];
  Object.keys(ctx.SCORECARD_LOB_FCR).forEach(lob => Object.keys(ctx.SCORECARD_LOB_FCR[lob]).forEach(q => got.push(lob + ' ' + q + ' ' + ctx.SCORECARD_LOB_FCR[lob][q].join(' '))));
  check('all ' + FCR.length + ' FCR rows (thresh, target, stretch) are as on the scorecards', JSON.stringify(got.sort()) === JSON.stringify(FCR.slice().sort()), got.filter(x => FCR.indexOf(x) === -1).join(' | '));
  check('9 scorecards (8 LOBs of the popup and One CAM); One CAM has no Q1', Object.keys(ctx.SCORECARD_LOB_FCR).length === 9 && !ctx.SCORECARD_LOB_FCR['One CAM'][1] && Object.keys(ctx.SCORECARD_LOB_FCR['One CAM']).join() === '2,3,4');
  check('FCR: threshold < target < stretch everywhere (higher is better)', got.every(r => { const a = r.split(' ').slice(-3).map(Number); return a[0] < a[1] && a[1] < a[2]; }));
  const tr = ctx.SCORECARD_LOB_TRANSFERS;
  check('Transfers: only MOB CxSS, Q1 to Q3, threshold 34, target 28, stretch 20', Object.keys(tr).join() === 'MOB CxSS' && Object.keys(tr['MOB CxSS']).join() === '1,2,3' && [1, 2, 3].every(q => tr['MOB CxSS'][q].join() === '34,28,20'));
  check('Transfers: threshold > target > stretch (lower is better)', tr['MOB CxSS'][1][0] > tr['MOB CxSS'][1][1] && tr['MOB CxSS'][1][1] > tr['MOB CxSS'][1][2]);
}

console.log('\nTest 2 — the 28 spellings of Line of Business in the data');
{
  const MAP = { 'PF CxSS': 'PF CxSS', 'PF CLS': 'PF CLS', 'PF TS': 'PF TS', 'SHS CLS': 'SHS CLS', 'SHS TS': 'SHS TS', 'MOB CLS': 'MOB CLS', 'MOB CxSS': 'MOB CxSS',
    'MOB TS': 'MOB CxSS', 'Koodo': 'MOB CxSS', 'Koodo TS': 'MOB CxSS', 'KDO CxSS': 'MOB CxSS',
    'MOB ONE CLS': 'MOB CLS', 'MOB ONR CLS': 'MOB CLS', 'CLS2 Access + Mob': 'MOB CLS',
    'ONE CLS': 'PF CLS', 'One CLS': 'PF CLS', 'one cls': 'PF CLS',
    'ONE TS': 'PF TS', 'One TS': 'PF TS', 'OHE TS': 'PF TS', 'FFH_FBR_TS_TIIN': 'PF TS', 'FFH_FBR_TS_TIP': 'PF TS', 'Copper / PF TS': 'PF TS',
    'SHS / Vivint TS': 'SHS TS', 'CxSS': 'PF CxSS', 'OHE CxSS': 'PF CxSS', 'FFH Sales': 'PF CxSS', 'One CAM': 'One CAM' };
  const bad = Object.keys(MAP).filter(k => ctx.scorecardLobKey_(k) !== MAP[k]);
  check('all ' + Object.keys(MAP).length + ' spellings found in the live data go to the scorecard the admin chose', bad.length === 0, bad.map(k => k + ' -> ' + ctx.scorecardLobKey_(k)).join('; '));
  check('SHS CxSS (in the popup, no audit yet) has its own scorecard', ctx.scorecardLobKey_('SHS CxSS') === 'SHS CxSS');
  check('capitals and extra spaces do not matter', ctx.scorecardLobKey_('  pf   CXSS ') === 'PF CxSS' && ctx.scorecardLobKey_('mob  cls') === 'MOB CLS');
  check('an unknown LOB, a blank and null have no scorecard', ctx.scorecardLobKey_('Brand New LOB') === '' && ctx.scorecardLobKey_('') === '' && ctx.scorecardLobKey_(null) === '' && ctx.scorecardLobKey_(undefined) === '');
  check('every scorecard an alias points to exists', Object.keys(ctx.SCORECARD_LOB_ALIASES).every(a => !!ctx.SCORECARD_LOB_FCR[ctx.SCORECARD_LOB_ALIASES[a]]));
  check('the 11 LOBs of the popup are all covered', ['PF CxSS', 'PF CLS', 'PF TS', 'SHS CxSS', 'SHS CLS', 'SHS TS', 'MOB CxSS', 'MOB CLS', 'MOB TS', 'Koodo', 'Koodo TS'].every(l => !!ctx.scorecardLobKey_(l)));
}

console.log('\nTest 3 — the quarter');
{
  const at = iso => Date.parse(iso);
  const q = (key, iso) => ctx.scorecardLobTargets_(key, iso === null ? null : at(iso));
  check('an audit in Aug 2026 uses Q3', q('PF CxSS', '2026-08-20T10:00:00Z').quarter === 3 && q('PF CxSS', '2026-08-20T10:00:00Z').label === 'Q3 2026' && q('PF CxSS', '2026-08-20T10:00:00Z').exact);
  check('Oct 1 and Dec 31 are Q4, Jan 1 2026 is Q1, Mar 31 is Q1, Apr 1 is Q2', [q('PF CxSS', '2026-10-01T12:00:00Z').quarter, q('PF CxSS', '2026-12-31T12:00:00Z').quarter, q('PF CxSS', '2026-01-01T12:00:00Z').quarter, q('PF CxSS', '2026-03-31T12:00:00Z').quarter, q('PF CxSS', '2026-04-01T12:00:00Z').quarter].join() === '4,4,1,1,2');
  const t4 = q('PF CxSS', '2026-10-27T10:00:00Z');
  check('the targets of the quarter are returned (PF CxSS Q4: 77.0 / 79.5 / 82.0)', t4.fcr.thresh === 77 && t4.fcr.target === 79.5 && t4.fcr.stretch === 82);
  const later = q('PF CxSS', '2027-02-10T10:00:00Z');
  check('after 2026 the closest quarter is used (Q4 2026) and it is said it is not exact', later.quarter === 4 && later.label === 'Q4 2026' && later.exact === false);
  const cam1 = q('One CAM', '2026-02-10T10:00:00Z');
  check('One CAM has no Q1: Jan to Mar use Q2 (closest), not exact', cam1.quarter === 2 && cam1.exact === false && cam1.fcr.target === 81);
  check('before 2026 the earliest quarter is used', q('PF CxSS', '2025-06-01T10:00:00Z').quarter === 1 && q('One CAM', '2025-06-01T10:00:00Z').quarter === 2);
  check('with no date the last quarter is used', q('PF CxSS', null).quarter === 4);
  check('MOB CxSS has Transfers in Q1 to Q3 and not in Q4', !!q('MOB CxSS', '2026-05-01T10:00:00Z').transfers && !!q('MOB CxSS', '2026-08-01T10:00:00Z').transfers && q('MOB CxSS', '2026-11-01T10:00:00Z').transfers === null);
  check('other LOBs never have Transfers; an unknown scorecard gives null', q('PF CLS', '2026-05-01T10:00:00Z').transfers === null && ctx.scorecardLobTargets_('Nope', 1) === null && ctx.scorecardLobTargets_('', 1) === null);
}

console.log('\nTest 4 — the four levels');
{
  const hi = { thresh: 77, target: 79.5, stretch: 82 }, lo = { thresh: 34, target: 28, stretch: 20 };
  const L = (v, t, h) => ctx.scorecardLevel_(v, t, h);
  check('higher is better: 70 below, 77 threshold, 79 threshold, 79.5 target, 81 target, 82 stretch, 95 stretch', [70, 77, 79, 79.5, 81, 82, 95].map(v => L(v, hi, true)).join() === 'below,threshold,threshold,target,target,stretch,stretch');
  check('just under the threshold is below (76.9)', L(76.9, hi, true) === 'below');
  check('lower is better: 40 below, 34 threshold, 30 threshold, 28 target, 25 target, 20 stretch, 0 stretch', [40, 34, 30, 28, 25, 20, 0].map(v => L(v, lo, false)).join() === 'below,threshold,threshold,target,target,stretch,stretch');
  check('no number (null, undefined, NaN) is "none"', L(null, hi, true) === 'none' && L(undefined, hi, true) === 'none' && L(NaN, hi, true) === 'none');
}

let n = 0;
const row = o => Object.assign({ 'Audit Ref': 'L' + (++n), 'Timestamp': '2026-10-05T10:00:00Z', 'Team Member': 'Agent A', 'SAP ID': '100', 'Analysis Type': 'Repeats & Transfer Analyzer',
  'Repeat Projection %': '10%', 'Issue Resolved': 'Yes', 'Critical Flags': '', 'Transfer Occurred': 'No', 'Sales Attempted': '', 'Email Status': 'Sent', 'Team Leader': 'TL One',
  'Line of Business': 'PF CxSS', 'Locale': 'TI Philippines', 'VTID': 'V1', 'Top SMART Opportunity': '', 'RCA Sub-Parameter': '' }, o);
// an agent with `yes` resolved and `no` unresolved audits (resolution = yes / (yes + no))
const agent = (name, lob, yes, no, o) => { const r = []; for (let i = 0; i < yes + no; i++) r.push(row(Object.assign({ 'Team Member': name, 'Line of Business': lob, 'Issue Resolved': i < yes ? 'Yes' : 'No' }, o || {}))); return r; };
const build = (rows, c) => (c || ctx).buildAgentCards_(rows, null);
const card = (built, name) => built.cards.find(c => c.name === name);
const tile = (c, key) => c.tiles.find(t => t.key === key);

console.log('\nTest 5 — the agent cards');
{
  const rows = [].concat(agent('Below', 'PF CxSS', 7, 3), agent('Thr', 'PF CxSS', 7, 2), agent('Tgt', 'PF CxSS', 4, 1), agent('Str', 'PF CxSS', 9, 1));
  const b = build(rows);
  const res = nm => tile(card(b, nm), 'resolution');
  check('PF CxSS Q4 (77.0 / 79.5 / 82.0): 70% below, 78% at threshold, 80% on target, 90% stretch', [res('Below').status, res('Thr').status, res('Tgt').status, res('Str').status].join() === 'below,threshold,target,stretch', [res('Below').display, res('Thr').display, res('Tgt').display, res('Str').display].join());
  check('the words say it, never colour alone', [res('Below').statusText, res('Thr').statusText, res('Tgt').statusText, res('Str').statusText].join() === 'Below threshold,At threshold,On target,Stretch');
  check('the tile is labelled as the audit and names the LOB target', res('Tgt').label === 'Resolution (audit)' && res('Tgt').target === 'Target >= 79.5% (LOB)' && res('Tgt').lobKpi === 'FCR', res('Tgt').target);
  const c = card(b, 'Tgt');
  check('the card knows its scorecard and quarter', c.scorecard.key === 'PF CxSS' && c.scorecard.label === 'Q4 2026' && c.scorecard.exact === true && c.section === 'PF CxSS');
  check('Repeat Risk and Transfer keep the app target and say "(audit)"; Critical flags keep their words', tile(c, 'repeatRisk').target === 'Target <= 30% (audit)' && tile(c, 'transfer').target === 'Target <= 20% (audit)' && tile(c, 'flags').target === 'Target <= 0 per audit');
  check('their status is still decided by the app target (two states)', ['target', 'below'].indexOf(tile(c, 'repeatRisk').status) !== -1 && ['target', 'below'].indexOf(tile(c, 'transfer').status) !== -1);
  check('the badge counts only "below threshold": an agent at threshold or on stretch is not penalised', card(b, 'Thr').belowCount === 0 || tile(card(b, 'Thr'), 'resolution').status !== 'below');
  const sales = build(agent('Sam', 'PF CxSS', 3, 1, { 'Analysis Type': 'Sales Analyzer', 'Sales Attempted': 'Yes' }));
  check('a Sales agent: Resolution uses the LOB scorecard, Sales attempted keeps 80% and says "(audit)"', tile(sales.cards[0], 'resolution').lobKpi === 'FCR' && tile(sales.cards[0], 'salesAttempt').target === 'Target >= 80% (audit)');

  // MOB CxSS Transfers (lower is better) in Q2, none in Q4
  const mk = (name, yes, no, iso) => [].concat(agent(name, 'MOB CxSS', 4, 0).slice(0, 0), (() => { const r = []; for (let i = 0; i < yes + no; i++) r.push(row({ 'Team Member': name, 'Line of Business': 'MOB CxSS', 'Timestamp': iso, 'Transfer Occurred': i < yes ? 'Yes' : 'No' })); return r; })());
  const q2 = build([].concat(mk('T0', 0, 4, '2026-05-10T10:00:00Z'), mk('T25', 1, 3, '2026-05-10T10:00:00Z'), mk('T33', 1, 2, '2026-05-10T10:00:00Z'), mk('T50', 1, 1, '2026-05-10T10:00:00Z')));
  const tr = nm => tile(card(q2, nm), 'transfer');
  check('MOB CxSS Q2 Transfers (34 / 28 / 20, lower is better): 0% stretch, 25% target, 33% threshold, 50% below', [tr('T0').status, tr('T25').status, tr('T33').status, tr('T50').status].join() === 'stretch,target,threshold,below', [tr('T0').display, tr('T25').display, tr('T33').display, tr('T50').display].join());
  check('the Transfer tile says it is the audit and the LOB target', tr('T25').label === 'Transfer rate (audit)' && tr('T25').target === 'Target <= 28% (LOB)' && tr('T25').lobKpi === 'Transfers');
  const q4 = build(mk('Late', 1, 3, '2026-11-10T10:00:00Z'));
  check('MOB CxSS Q4 has no Transfers KPI: the tile keeps the app target and says "(audit)"', tile(q4.cards[0], 'transfer').target === 'Target <= 20% (audit)' && !tile(q4.cards[0], 'transfer').lobKpi && tile(q4.cards[0], 'resolution').lobKpi === 'FCR');

  // an agent audited in two LOBs goes where most of their audits are
  const two = build([].concat(agent('Both', 'PF CLS', 1, 0), agent('Both', 'MOB CLS', 2, 1)));
  check('an agent audited in two LOBs is judged by the LOB with most of their audits', two.cards[0].section === 'MOB CLS' && two.cards[0].scorecard.key === 'MOB CLS');
  const noLob = build(agent('Blank', '', 2, 0));
  check('a blank LOB goes to "Other LOB" with the app targets and no scorecard', noLob.cards[0].section === 'Other LOB' && noLob.cards[0].scorecard === null && tile(noLob.cards[0], 'resolution').target === 'Target >= 80%' && tile(noLob.cards[0], 'resolution').label === 'Resolution rate');
}

console.log('\nTest 6 — the sections');
{
  const rows = [].concat(agent('P1', 'PF CxSS', 4, 1), agent('P2', 'pf cxss', 4, 1), agent('M1', 'MOB CLS', 4, 1), agent('C1', 'one cls', 4, 1), agent('W1', 'Brand New LOB', 4, 1), agent('K1', 'Koodo', 4, 1));
  const b = build(rows);
  check('sections: the LOB scorecards in alphabetical order, "Other LOB" last', b.sections.map(s => s.name).join() === 'MOB CLS,MOB CxSS,PF CLS,PF CxSS,Other LOB', b.sections.map(s => s.name).join());
  check('"one cls" is in PF CLS and Koodo is in MOB CxSS', b.sections.find(s => s.name === 'PF CLS').cards[0].name === 'C1' && b.sections.find(s => s.name === 'MOB CxSS').cards[0].name === 'K1');
  check('two spellings of the same LOB are one section', b.sections.find(s => s.name === 'PF CxSS').cards.map(c => c.name).join() === 'P1,P2');
  check('every agent is in exactly one section and the cards keep their alphabetical order', b.sections.reduce((a, s) => a + s.cards.length, 0) === b.cards.length && b.cards.map(c => c.name).join() === 'C1,K1,M1,P1,P2,W1');
  check('Other LOB has no scorecard; the others have one', b.sections.every(s => (s.name === 'Other LOB') === (s.scorecard === null)));

  // the quarter of a section is the newest audit of the whole section, for every agent in it
  const mixed = build([].concat(agent('Old', 'PF CxSS', 4, 1, { 'Timestamp': '2026-08-10T10:00:00Z' }), agent('New', 'PF CxSS', 4, 1, { 'Timestamp': '2026-10-10T10:00:00Z' })));
  check('one section, one quarter: an agent audited in Aug is judged by Q4 when a colleague was audited in Oct', mixed.sections[0].scorecard.label === 'Q4 2026' && card(mixed, 'Old').scorecard.label === 'Q4 2026');
  const sep = build([].concat(agent('Old', 'PF CxSS', 4, 1, { 'Timestamp': '2026-08-10T10:00:00Z' }), agent('New', 'PF CLS', 4, 1, { 'Timestamp': '2026-10-10T10:00:00Z' })));
  check('two sections have their own quarter', sep.sections.find(s => s.name === 'PF CxSS').scorecard.label === 'Q3 2026' && sep.sections.find(s => s.name === 'PF CLS').scorecard.label === 'Q4 2026');

  const bare = load(['Config.gs', 'ScorecardMetrics.gs', 'ScorecardLayout.gs']);
  const old = bare.buildAgentCards_(agent('Old', 'PF CxSS', 4, 1), null);
  check('without ScorecardTargets.gs nothing changes: one section without a name and the app targets', old.sections.length === 1 && old.sections[0].name === '' && old.cards[0].scorecard === null && tile(old.cards[0], 'resolution').target === 'Target >= 80%');
}

console.log('\nTest 7 — the group numbers of a section');
{
  const rows = [].concat(agent('Below', 'PF CxSS', 7, 3), agent('Thr', 'PF CxSS', 7, 2), agent('Tgt', 'PF CxSS', 4, 1), agent('Str', 'PF CxSS', 9, 1));
  const b = build(rows), sec = b.sections[0], g = ctx.buildGroupSummary_(b, 'B', '', 40, sec);
  const k = g.kpis.find(x => x.key === 'resolution');
  check('the section is described (name, quarter, the FCR targets)', g.section.name === 'PF CxSS' && g.section.hasLob && g.section.quarterLabel === 'Q4 2026' && g.section.fcr.target === 79.5);
  check('resolution: 1 below, 1 at threshold, 1 on target, 1 stretch; "on target" counts target and stretch (2 of 4)', k.below === 1 && k.threshold === 1 && k.stretch === 1 && k.onTarget === 2 && k.measured === 4 && k.applicable === 4, JSON.stringify([k.below, k.threshold, k.stretch, k.onTarget]));
  check('the average (70, 78, 80, 90 = 79.5 -> 80) is judged on the same four levels: on target', k.display === '80%' && k.status === 'target' && k.statusText === 'On target' && k.lobLevels === true);
  check('the label, target and chart target are the LOB ones', k.label === 'Resolution (audit)' && k.targetShort === '>= 79.5%' && k.targetValue === 79.5 && k.threshValue === 77);
  const rr = g.kpis.find(x => x.key === 'repeatRisk');
  check('Repeat Risk keeps the app target and says "(audit)"', rr.lobLevels === false && rr.target === 'Target <= 30% (audit)' && rr.targetValue === 30);
  const step = g.nextSteps.find(s => s.title === 'Resolution (audit)');
  check('the next step says "threshold" for a LOB KPI and gives the threshold (77%)', step && /^1 of 4 agents? (is|are) below the 77% threshold\.$/.test(step.body), step && step.body);
  const flat = ctx.buildGroupSummary_(b, 'B', '', 40);
  check('without a section all the cards are one group with no section', flat.section === null && flat.agents === 4);
  const other = build(agent('X', 'Brand New LOB', 4, 1)), go = ctx.buildGroupSummary_(other, 'B', '', 5, other.sections[0]);
  check('Other LOB: no scorecard, the app targets', go.section.hasLob === false && go.kpis.find(x => x.key === 'resolution').lobLevels === false && go.kpis.find(x => x.key === 'resolution').target === 'Target >= 80%');
}

console.log('\nTest 8 — the deck: a section for each LOB');
const mixedRows = [].concat(agent('Ana', 'PF CxSS', 7, 3), agent('Ben', 'PF CxSS', 9, 1), agent('Cy', 'MOB CLS', 4, 1), agent('Di', 'one cls', 4, 1), agent('Eve', 'Brand New LOB', 4, 1));
const mixedBuilt = build(mixedRows);
const deck = ctx.scorecardDeckShapes_(mixedBuilt, 'Batch 15', 'Team Leader: TL One', mixedRows.length, []);
const title = s => { const t = s.find(x => x.id === 'title') || s.find(x => x.id === 'name'); return t ? t.paras[0].runs.map(r => r.t).join('') : ''; };
const text = shapes => shapes.map(s => s.paras.map(p => p.runs.map(r => r.t).join('')).join(' ')).join(' | ');
{
  // 1 cover + each section: 6 group slides (summary, results, landed, heatmap, findings, readiness) + its agents + next steps
  const want = 1 + (6 + 1 + 1) + (6 + 1 + 1) + (6 + 1 + 2) + (6 + 1 + 1);
  check('the deck has ' + want + ' slides: the cover and four sections (MOB CLS 1, PF CLS 1, PF CxSS 2, Other LOB 1 agents)', deck.length === want, 'got ' + deck.length);
  const titles = deck.map(title);
  check('the cover first, then a section: summary, results, landed, heatmap, findings, readiness, the agents, next steps',
        titles[0] === 'Audit Scorecards' && titles.slice(1, 9).join('|') === ['Executive summary', 'Scorecard results', 'Where the agents landed on each KPI', 'Scorecard heatmap by agent', 'Audit findings: strengths and gaps', 'Readiness', 'Cy', 'Next steps'].join('|'), titles.slice(0, 10).join(' | '));
  const tags = deck.map(s => { const t = s.find(x => x.id === 'lob-tag'); return t ? text([t]) : ''; });
  check('every group slide has a purple tag with the LOB; the cover and the agent slides do not', tags[1] === 'MOB CLS' && tags[2] === 'MOB CLS' && tags[7] === '' && tags[8] === 'MOB CLS' && tags[0] === '');
  check('the sections are in order: MOB CLS, PF CLS, PF CxSS, Other LOB', [tags[1], tags[9], tags[17], tags[27]].join() === 'MOB CLS,PF CLS,PF CxSS,Other LOB', [tags[1], tags[9], tags[17], tags[27]].join());
  check('the page numbers run 1 to ' + want + ' in order, each "n / ' + want + '"', deck.every((sl, i) => JSON.stringify(sl.find(s => s.id === 'page')).indexOf((i + 1) + ' / ' + want) !== -1));
  check('the cover lists the LOBs and how many agents each has', /LOBs: MOB CLS \(1\), PF CLS \(1\), PF CxSS \(2\), Other LOB \(1\)/.test(text(deck[0])), text(deck[0]).slice(-120));
  check('a group slide footer names the quarter of the targets', /Q4 2026 targets/.test(text([deck[1].find(s => s.id === 'footer')])) && !/targets/.test(text([deck[27].find(s => s.id === 'footer')])));

  const results = deck[17 + 1];   // results of PF CxSS
  check('the results slide of PF CxSS says which FCR scorecard it uses, with the three numbers', /FCR scorecard of PF CxSS Q4 2026: Threshold 77%, Target 79.5%, Stretch 82%/.test(text(results)), text(results).slice(-220));
  const landed = deck[17 + 2];
  check('the landed slide has the five levels in its legend for a LOB section, three for Other LOB', /Below threshold/.test(text(landed)) && /At threshold/.test(text(landed)) && /Stretch/.test(text(landed)) && !/At threshold/.test(text(deck[27 + 2])), '');
  const heat = deck[17 + 3];
  check('the heatmap legend has the four levels, its note names the scorecard; Other LOB has the old note', /At threshold/.test(text(heat)) && /FCR scorecard of PF CxSS Q4 2026/.test(text(heat)) && !/FCR scorecard/.test(text(deck[27 + 3])));
  const cells = heat.filter(s => /^l-c-\d+-1$/.test(s.id)).map(s => s.fill);
  check('the heatmap Resolve cells use the level colours (Ana 70% red, Ben 90% purple)', cells[0] === T.red && cells[1] === T.purple, cells.join());
  const agentSlide = deck[17 + 7];
  check('the slide of an agent says the LOB scorecard and quarter, and the Resolution tile says "(audit)"', /PF CxSS Q4 2026 targets/.test(text(agentSlide)) && /Resolution \(audit\)/.test(text(agentSlide)) && /Target >= 79.5% \(LOB\)/.test(text(agentSlide)), text(agentSlide).slice(0, 220));
}

console.log('\nTest 9 — the look of every slide of the mixed deck and of a 40-agent, 6-LOB deck');
{
  const bgAt = (shapes, i) => {
    const s = shapes[i];
    if (s.fill) return s.fill;
    for (let j = i - 1; j >= 0; j--) { const b = shapes[j]; if (b.fill && b.x <= s.x + 0.5 && b.y <= s.y + 0.5 && b.x + b.w >= s.x + s.w - 0.5 && b.y + b.h >= s.y + s.h - 0.5) return b.fill; }
    return T.white;
  };
  const LOBS = ['PF CxSS', 'PF CLS', 'PF TS', 'MOB CLS', 'MOB CxSS', 'one cls', 'Brand New LOB'];
  const many = []; for (let i = 0; i < 40; i++) many.push(...agent(i % 5 === 0 ? 'Maria Guadalupe Hernandez-Villanueva de la Cruz ' + i : 'Agent Number ' + i, LOBS[i % LOBS.length], 3 + (i % 4), 1 + (i % 3), { 'Repeat Projection %': ['10%', '45%', '80%'][i % 3], 'Critical Flags': i % 2 ? 'Customer was transferred without a warm handoff and the agent did not set any expectation for the next step on the call' : '', 'Transfer Occurred': i % 4 ? 'No' : 'Yes' }));
  const decks = { mixed: deck, forty: ctx.scorecardDeckShapes_(build(many), 'MOB CLS Batch 15 ' + 'x'.repeat(40), 'Team Leader: ' + 'y'.repeat(80), many.length, []) };
  const pairs = [[T.white, T.purple], [T.onDark, T.purple], [T.purple, T.lilac], [T.purple, T.white], [T.ink, T.lilac], [T.ink, T.white], [T.grey, T.white], [T.grey, T.lilac], [T.green, T.lilac], [T.green, T.white],
                 [T.red, T.lilac], [T.red, T.white], [T.amber, T.lilac], [T.amber, T.white], [T.white, T.green], [T.white, T.red], [T.white, T.grey], [T.white, T.amber]];
  const low = pairs.filter(([a, b]) => ctx._slContrast_(a, b) < 4.5);
  check('all the colours used for text, including amber and purple for threshold and stretch, have contrast >= 4.5:1', low.length === 0, JSON.stringify(low.map(([a, b]) => a + ' on ' + b + ' = ' + ctx._slContrast_(a, b).toFixed(2))));
  Object.keys(decks).forEach(name => {
    const slides = decks[name], pr = { outside: [], dup: [], small: [], contrast: [], overflow: [] };
    slides.forEach((shapes, si) => {
      const tag = name + ' slide ' + (si + 1) + ' (' + title(shapes).slice(0, 24) + ')', ids = shapes.map(s => s.id);
      if (new Set(ids).size !== ids.length) pr.dup.push(tag);
      shapes.forEach((s, i) => {
        if (s.x < -0.01 || s.y < -0.01 || s.x + s.w > P.w + 0.01 || s.y + s.h > P.h + 0.01) pr.outside.push(tag + ' ' + s.id);
        const bg = bgAt(shapes, i);
        s.paras.forEach(p => p.runs.forEach(r => {
          const minFont = (s.id === 'footer' || s.id === 'page') ? ctx.SCORECARD_FOOTER_FONT : ctx.SCORECARD_MIN_FONT;
          if (r.size < minFont) pr.small.push(tag + ' ' + s.id + ' ' + r.size);
          if (ctx._slContrast_(r.color, bg) < 4.5) pr.contrast.push(tag + ' ' + s.id + ' ' + r.color + ' on ' + bg);
        }));
        if (s.paras.length) {
          let need = 2 * s.pad; const inner = s.w - 2 * s.pad;
          s.paras.forEach(p => { const size = Math.max.apply(null, p.runs.map(r => r.size)); need += ctx._slHeight_(ctx._slLines_(p.runs.map(r => r.t).join(''), inner - (p.bullet ? 14 : 0), size), size, p.spacing); });
          if (need > s.h + 0.5) pr.overflow.push(tag + ' ' + s.id + ' needs ' + need.toFixed(0) + ' has ' + s.h);
        }
      });
    });
    check(name + ' (' + slides.length + ' slides): every shape is inside the page', pr.outside.length === 0, pr.outside.slice(0, 4).join('; '));
    check(name + ': every shape id is unique on its slide', pr.dup.length === 0, pr.dup.slice(0, 3).join('; '));
    check(name + ': no font below 11 pt (footer 9 pt)', pr.small.length === 0, pr.small.slice(0, 4).join('; '));
    check(name + ': every run has contrast >= 4.5:1 against what is behind it', pr.contrast.length === 0, pr.contrast.slice(0, 4).join('; '));
    check(name + ': no text overflows its box', pr.overflow.length === 0, pr.overflow.slice(0, 6).join('; '));
  });
  const pageNums = decks.forty.map(s => JSON.stringify(s.find(x => x.id === 'page'))), tot = decks.forty.length;
  check('40 agents in 7 sections: the page numbers are 1 to ' + tot, pageNums.every((p, i) => p.indexOf((i + 1) + ' / ' + tot) !== -1));
  const parts = ctx.buildPptxParts_(decks.forty);
  check('the PowerPoint package has one slide file per slide', parts.filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p.path)).length === tot);
}

console.log('\n' + (check.state.failures === 0 ? 'ALL PASSED' : check.state.failures + ' FAILED'));
process.exit(check.state.failures === 0 ? 0 : 1);
