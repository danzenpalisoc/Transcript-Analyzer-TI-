/**
 * Agent scorecard: the numbers (ScorecardMetrics.gs) and the look (ScorecardLayout.gs).
 *
 *  1. The server rules give the same answer as the dashboard rules (parity), so a scorecard never disagrees with the dashboard.
 *  2. The numbers: averages, rates, thresholds, badge, tile choice, wording.
 *  3. What's Working is read from a saved report.
 *  4. The look: contrast, minimum font size, nothing spills out of its box or the page, even for the worst-case agent.
 *
 * Usage: node tests/scorecard-metrics-test.js
 */
const vm = require('vm');
const { load } = require('./lib/page-harness');
const { read, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

const sandbox = { Logger: { log() {} }, JSON, Array, Object, String, Number, Math, Date, RegExp, isFinite, isNaN, Error,
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) } };
const ctx = vm.createContext(sandbox);
['Config.gs', 'ScorecardMetrics.gs', 'ScorecardLayout.gs'].forEach(f => vm.runInContext(read(f), ctx, { filename: f }));
const page = load();     // index.html, with the dashboard helper block

const row = o => Object.assign({ 'Audit Ref': 'R' + Math.random().toString(36).slice(2, 8), 'Timestamp': '2026-10-05T10:00:00Z', 'Team Member': 'Agent A', 'SAP ID': '111',
  'Analysis Type': 'Repeats & Transfer Analyzer', 'Repeat Projection %': '40%', 'Issue Resolved': 'Yes', 'Critical Flags': '', 'Transfer Occurred': 'No',
  'Sales Attempted': '', 'Email Status': 'Sent', 'Team Leader': 'TL One', 'Line of Business': 'MOB CLS', 'Locale': 'TI Philippines', 'VTID': 'VTI1',
  'Top SMART Opportunity': '', 'RCA Sub-Parameter': '' }, o);
const cards = (rows, extras) => ctx.buildAgentCards_(rows, extras);
const one = (rows, extras) => cards(rows, extras).cards[0];
const tile = (c, key) => c.tiles.find(t => t.key === key);

console.log('\nTest 1 — parity: the scorecard reads the data exactly as the dashboard does');
{
  const nums = [0, 0.78, '0.78', 0.975, 1, '1', '1%', '0.5%', 65, '65', ' 40 % ', '40%', '', null, undefined, 'N/A', 'about forty', '95% repeat risk. No issue was identified',
    '~55% repeat risk. Key drivers', "High — approximately 80%. Customer's compensation", 'Repeat risk estimated at ~45%.', 'Low-to-moderate repeat risk (~25%). The $49.50 charge',
    'Repeat risk 40-50% depending', 'charged 250% of the normal rate', '85% — Transfer without resolution'];
  check('numbers: ' + nums.length + ' examples agree', nums.every(v => ctx._scNum_(v) === page.ctx._dbNum(v)),
        JSON.stringify(nums.filter(v => ctx._scNum_(v) !== page.ctx._dbNum(v))));
  const res = ['Yes', 'No', '', null, 'Unknown', 'No, not yes', 'Not resolved - yes follow up', 'Partially resolved - the credit has not posted', 'Completely unresolved: the agent transferred',
    'Issue was NOT resolved. The promised $10 payment', 'The issue was resolved after the reset', 'Issue was partially resolved; credit pending', "The issue wasn't resolved on this call",
    'Resolution is ambiguous. Interaction summary', 'Agent confirmed the form was submitted',
    // sentences that begin with a verdict (the real ones that the diagnoseDashboardGaps run found, shortened)
    'Issue not resolved. Suggestions were limited to: check cable connections', 'Issue NOT resolved. Customer wanted internet relocation', 'Issue not resolved',
    'Neither issue was resolved on this call', 'Issue resolved — the agent fixed the billing error', 'Primary issue resolved after the second reset',
    'The primary issue was resolved but a secondary issue remains', 'Issue was ultimately not resolved after two calls', 'Issue was ultimately resolved after a technician visit',
    'Issue resolved, however the customer is still unhappy', 'Technically resolved — but the customer will call again', 'Service was restored but no ticket was created',
    'Issue about previously resolved bills was discussed', 'Issue resolvedX'];
  check('Issue Resolved: ' + res.length + ' examples agree', res.every(v => ctx._scResolution_(v) === page.ctx._dbResolution(v)),
        JSON.stringify(res.filter(v => ctx._scResolution_(v) !== page.ctx._dbResolution(v))));
  // the copy inside AuditCheck.gs (verifyDashboardData) must agree too
  const acSrc = require('fs').readFileSync(require('path').join(__dirname, '..', 'AuditCheck.gs'), 'utf8').replace(/\r\n/g, '\n');
  const i0 = acSrc.indexOf('var verdictOf = function(raw) {'), i1 = acSrc.indexOf('\n  };', i0);
  const verdictOf = vm.runInNewContext('(' + acSrc.slice(i0 + 'var verdictOf = '.length, i1 + 5).replace(/;$/, '') + ')', { String });
  check('the copy in AuditCheck.gs agrees on the same ' + res.length + ' examples', res.every(v => verdictOf(v) === ctx._scResolution_(v)), JSON.stringify(res.filter(v => verdictOf(v) !== ctx._scResolution_(v))));
  const want = { 'Issue not resolved. Suggestions were limited to: check cable connections': 'no', 'Issue NOT resolved. Customer wanted internet relocation': 'no', 'Issue not resolved': 'no',
    'Neither issue was resolved on this call': 'no', 'Issue resolved — the agent fixed the billing error': 'yes', 'Primary issue resolved after the second reset': 'yes',
    'The primary issue was resolved but a secondary issue remains': 'unknown', 'Issue was ultimately not resolved after two calls': 'no', 'Issue was ultimately resolved after a technician visit': 'yes',
    'Issue resolved, however the customer is still unhappy': 'unknown', 'Technically resolved — but the customer will call again': 'unknown', 'Service was restored but no ticket was created': 'unknown',
    'Issue about previously resolved bills was discussed': 'unknown', 'Issue resolvedX': 'unknown', 'The issue was resolved after the reset': 'yes' };
  const wrong = Object.keys(want).filter(k => ctx._scResolution_(k) !== want[k]).map(k => k.slice(0, 40) + ' => ' + ctx._scResolution_(k));
  check('sentences that start with a clear verdict are read (not resolved = no, resolved = yes) and the qualified or unclear ones are left for a person', wrong.length === 0, JSON.stringify(wrong));
  const flags = ['', 'None', null, 'F1, F2, F3', 'Wrong LOB, handled a TV call | No save attempt', 'Wrong LOB, handled a TV call, No save attempt', 'No save attempt | None'];
  check('flags: ' + flags.length + ' examples agree', flags.every(v => JSON.stringify(ctx._scSplitFlags_(v)) === JSON.stringify(page.ctx._dbSplitFlags(v))));
  check('names and keys agree', ['  Agent   A ', 'Agent A', '', null].every(v => ctx._scName_(v) === page.ctx._dbName(v) && ctx._scKey_(v) === page.ctx._dbKey(v)));
}

console.log('\nTest 2 — the numbers on a card');
{
  const rows = [
    row({ 'Repeat Projection %': '0.78', 'Issue Resolved': 'Yes', 'Critical Flags': 'No save attempt | Generic offer', 'Transfer Occurred': 'Yes' }),
    row({ 'Repeat Projection %': '65', 'Issue Resolved': 'Partially resolved - credit pending', 'Critical Flags': 'No save attempt', 'Transfer Occurred': 'No' }),
    row({ 'Repeat Projection %': '0%', 'Issue Resolved': 'No', 'Critical Flags': '', 'Transfer Occurred': 'The call was transferred five times' }),
    row({ 'Repeat Projection %': '95% repeat risk. No issue was identified', 'Issue Resolved': '', 'Critical Flags': 'None', 'Transfer Occurred': 'Yes' }),
  ];
  const c = one(rows);
  check('one agent, four audits', c.name === 'Agent A' && c.audits === 4);
  check('Avg Repeat Risk = (78 + 65 + 0 + 95) / 4 = 60: fractions, text and a real 0% all count', tile(c, 'repeatRisk').value === 60, JSON.stringify(tile(c, 'repeatRisk')));
  check('Resolution = Resolved of audits with a verdict: 1 of 3 = 33%', tile(c, 'resolution').value === 33 && /Partial 1 \| Unresolved 1/.test(tile(c, 'resolution').sub), JSON.stringify(tile(c, 'resolution')));
  check('Critical flags = 3 flags in 4 audits = 0.8 per audit ("None" is no flag)', tile(c, 'flags').value === 0.8 && /3 flags in 4 audits/.test(tile(c, 'flags').sub), JSON.stringify(tile(c, 'flags')));
  check('Transfer rate counts Yes and No only: 2 of 3 = 67% (a sentence is not counted)', tile(c, 'transfer').value === 67, JSON.stringify(tile(c, 'transfer')));
  check('four tiles in this order', c.tiles.map(t => t.key).join() === 'repeatRisk,resolution,flags,transfer');
  check('supporting: audits 4, repeats 4, sales 0, emailed 100%', c.supporting.map(s => s.display).join() === '4,4,0,100%', JSON.stringify(c.supporting));
  check('Not recorded is not in the denominator (4 audits, 3 with a verdict)', true);

  const blankCard = one([row({ 'Repeat Projection %': '', 'Issue Resolved': '', 'Transfer Occurred': '' })]);
  check('no repeat value / no verdict / no transfer answer: tiles say No data, not 0', tile(blankCard, 'repeatRisk').status === 'none' && tile(blankCard, 'resolution').status === 'none' && tile(blankCard, 'transfer').status === 'none' && tile(blankCard, 'repeatRisk').display === 'N/A');
  check('a real 0% repeat risk is a number, not "no data"', one([row({ 'Repeat Projection %': '0%' })]).tiles[0].value === 0);
}

console.log('\nTest 3 — thresholds are compared as shown, and the badge counts the tiles below');
{
  const rc = pct => one([row({ 'Repeat Projection %': pct + '%' })]).tiles[0].status;
  check('Repeat Risk 30% is on target, 31% is below (lower is better)', rc(30) === 'target' && rc(31) === 'below');
  const mk = (yes, total) => { const r = []; for (let i = 0; i < total; i++) r.push(row({ 'Issue Resolved': i < yes ? 'Yes' : 'No' })); return one(r); };
  check('Resolution 80% is on target, 79% is below (higher is better)', tile(mk(4, 5), 'resolution').status === 'target' && tile(mk(79, 100), 'resolution').status === 'below');
  check('Critical flags: 0 flags is on target, one flag in 4 audits (0.3) is below', one([row({}), row({}), row({}), row({})]).tiles[2].status === 'target' && one([row({ 'Critical Flags': 'F1' }), row({}), row({}), row({})]).tiles[2].status === 'below');
  const trs = (yes, total) => { const r = []; for (let i = 0; i < total; i++) r.push(row({ 'Transfer Occurred': i < yes ? 'Yes' : 'No' })); return tile(one(r), 'transfer').status; };
  check('Transfer 20% is on target, 25% is below', trs(1, 5) === 'target' && trs(1, 4) === 'below');

  const good = one([row({ 'Repeat Projection %': '10%', 'Issue Resolved': 'Yes', 'Critical Flags': '', 'Transfer Occurred': 'No' })]);
  check('no tile below: On track', good.badge.level === 'ontrack' && good.badge.text === 'On track' && good.belowCount === 0, JSON.stringify(good.badge));
  const one1 = one([row({ 'Repeat Projection %': '10%', 'Issue Resolved': 'Yes', 'Critical Flags': 'F1', 'Transfer Occurred': 'No' })]);
  check('one tile below: Needs attention', one1.badge.level === 'attention' && one1.belowCount === 1, JSON.stringify(one1.badge));
  const bad = one([row({ 'Repeat Projection %': '90%', 'Issue Resolved': 'No', 'Critical Flags': 'F1', 'Transfer Occurred': 'Yes' })]);
  check('two or more below: Priority support', bad.badge.level === 'priority' && bad.badge.text === 'Priority support' && bad.belowCount === 4, JSON.stringify(bad.badge));
  const none = one([row({ 'Repeat Projection %': '', 'Issue Resolved': '', 'Critical Flags': '', 'Transfer Occurred': '' })]);
  check('flags tile has data even with no flags, so the badge is never empty', none.badge.level === 'ontrack' || none.badge.level === 'nodata');
}

console.log('\nTest 4 — Sales agents, mixed agents, grouping');
{
  const sales = [row({ 'Analysis Type': 'Sales Analyzer', 'Sales Attempted': 'Yes' }), row({ 'Analysis Type': 'Sales Analyzer', 'Sales Attempted': 'No' }),
                 row({ 'Analysis Type': 'Sales Analyzer', 'Sales Attempted': 'N/A' }), row({ 'Analysis Type': 'Sales Analyzer', 'Sales Attempted': 'Yes' })];
  const c = one(sales);
  check('an agent with only Sales audits gets the Sales attempted tile (2 of 3 = 67%, N/A not counted)', tile(c, 'salesAttempt') && !tile(c, 'transfer') && tile(c, 'salesAttempt').value === 67, JSON.stringify(c.tiles.map(t => t.key)));
  const mixed = one(sales.concat([row({ 'Transfer Occurred': 'Yes' })]));
  check('an agent with both gets the Transfer tile, and the sales rate goes into Also noted', tile(mixed, 'transfer') && /Sales attempted on 67%/.test(mixed.alsoNoted), mixed.alsoNoted);

  const g = cards([row({ 'Team Member': 'Agent A' }), row({ 'Team Member': ' agent   a ' }), row({ 'Team Member': 'Agent B' }), row({ 'Team Member': '' }),
                   row({ 'Team Member': 'x'.repeat(120), 'Audit Ref': 'LONG' }), row({ 'Team Member': 'Name\nwith newline', 'Audit Ref': 'NL' })]);
  check('names with different spacing or capitals are one agent', g.cards.length === 2 && g.cards.find(c => c.audits === 2));
  check('a blank name, a paragraph and a name with a newline are skipped and listed', g.skipped.length === 3 && g.skipped.some(s => s.ref === 'LONG') && g.skipped.some(s => s.ref === 'NL'), JSON.stringify(g.skipped));
  check('cards are sorted by name', g.cards.map(c => c.name).join() === 'Agent A,Agent B');
  check('the period covers the audits', /Oct 5, 2026/.test(g.range), g.range);
}

console.log('\nTest 5 — the wording');
{
  const rows = [
    row({ 'Critical Flags': 'No save attempt | Generic offer | Missed callback', 'Top SMART Opportunity': 'Ask one open-ended discovery question at the start of every call', 'Audit Ref': 'W1' }),
    row({ 'Critical Flags': 'no save attempt | Generic offer', 'Top SMART Opportunity': 'Ask one open-ended discovery question at the start of every call', 'Audit Ref': 'W2' }),
    row({ 'Critical Flags': 'No save attempt', 'Top SMART Opportunity': 'Confirm the callback time before closing', 'Audit Ref': 'W3' }),
  ];
  const c = one(rows, { whatsWorking: { W1: ['Warm greeting and clear ownership', 'Accurate billing explanation'], W2: ['Warm greeting and clear ownership'], W3: [] } });
  check('Gaps: the most common flags first, capitals ignored: No save attempt (3), Generic offer (2), Missed callback (1)', c.gaps.join('|') === 'No save attempt|Generic offer|Missed callback', c.gaps.join('|'));
  check('Strengths: the most common "What\'s Working" first', c.strengths[0] === 'Warm greeting and clear ownership' && c.strengths.length === 2, JSON.stringify(c.strengths));
  check('Coaching: the most common Top SMART Opportunity first', c.coaching[0].indexOf('Ask one open-ended') === 0 && c.coaching.length === 2, JSON.stringify(c.coaching));
  check('Also noted names the areas below target with the target', /Avg Repeat Risk: 40% \(target <= 30%\)\. Critical flags: 2\.0 per audit \(target <= 0 per audit\)\./.test(c.alsoNoted), c.alsoNoted);

  const none = one([row({ 'Critical Flags': '' })]);
  check('with no report text, Strengths falls back to the areas that meet their target', none.strengths.length > 0 && /Meets the target/.test(none.strengths[0]), JSON.stringify(none.strengths));
  const rcaCard = one([row({ 'Critical Flags': 'F1', 'RCA Sub-Parameter': 'Hold handling; Empathy' })]);
  check('with fewer than 3 flags, Gaps adds RCA sub-parameters (each item, not the pair)', rcaCard.gaps.join('|') === 'F1|Hold handling|Empathy', rcaCard.gaps.join('|'));
  check('nothing is invented when there is nothing: empty lists, not made-up text', one([row({})]).gaps.length === 0 && one([row({})]).coaching.length === 0);
}

console.log('\nTest 6b — the report as the Analysis sheet really holds it: PLAIN TEXT made by htmlToPlainText');
{
  // The real saved text: the report HTML (the template of AIService.gs) run through the real htmlToPlainText of Code.gs.
  const { extractFunction } = require('./lib/gas-fakes');
  vm.runInContext(extractFunction(read('Code.gs'), 'htmlToPlainText'), ctx);
  const tpl = '<div class="report-3col">\n  <div class="report-col report-col-working">\n    <div class="report-col-head">&#9989; What\'s Working</div>\n    <ul class="report-col-list">\n' +
    '      <li contenteditable="true">Warm greeting &amp; clear ownership: "I can help with that"</li>\n      <li contenteditable="true">Accurate <b>billing</b> explanation</li>\n    </ul>\n  </div>\n' +
    '  <div class="report-col report-col-change">\n    <div class="report-col-head">&#128205; What Needs to Change</div>\n    <ul class="report-col-list">\n      <li contenteditable="true">Should not appear</li>\n    </ul>\n  </div>\n</div>';
  const plain = ctx.htmlToPlainText(tpl);
  check('the saved text really is plain (no report-col-working class left), so the HTML reader alone cannot work', plain.indexOf('report-col-working') === -1 && /What.s Working/.test(plain), JSON.stringify(plain));
  const got = ctx.extractWhatsWorking_(plain);
  check('the two strengths are read from the plain text, bullets and tags gone', JSON.stringify(got) === JSON.stringify(['Warm greeting & clear ownership: "I can help with that"', 'Accurate billing explanation']), JSON.stringify(got));
  check('the next column heading and its bullets are not read', got.every(x => x.indexOf('Should not') === -1 && x.indexOf('Needs to Change') === -1));
  check('plain text with no such heading gives [] and never throws', ctx.extractWhatsWorking_('Some report\n• a bullet').length === 0 && ctx.extractWhatsWorking_('').length === 0);
  check('a heading with no bullets gives []', ctx.extractWhatsWorking_("What's Working\n\nWhat Needs to Change\n• x").length === 0);
  check('at most 3 items are taken', ctx.extractWhatsWorking_("What's Working\n• a\n• b\n• c\n• d").length === 3);
  const c = one([row({ 'Audit Ref': 'P1' })], { whatsWorking: { P1: got } });
  check('and they reach the card as Strengths, instead of "Meets the target"', c.strengths.length === 2 && !/Meets the target/.test(c.strengths.join(' ')), JSON.stringify(c.strengths));
}

console.log('\nTest 6 — "What\'s Working" is read from a saved report');
{
  const html = '<div class="report-wrap"><div class="report-cols">' +
    '<div class="report-col report-col-working"><div class="report-col-head">&#9989; What\'s Working</div><ul>' +
    '<li contenteditable="true">Warm greeting &amp; clear <strong>ownership</strong> of the issue</li><li>Accurate billing explanation</li></ul></div>' +
    '<div class="report-col report-col-needs"><div class="report-col-head">What Needs to Change</div><ul><li>Should not appear</li></ul></div></div></div>';
  check('the two items are read, tags removed and entities decoded', JSON.stringify(ctx.extractWhatsWorking_(html)) === JSON.stringify(['Warm greeting & clear ownership of the issue', 'Accurate billing explanation']), JSON.stringify(ctx.extractWhatsWorking_(html)));
  check('the next column is not read', ctx.extractWhatsWorking_(html).every(x => x.indexOf('Should not') === -1));
  check('a report without the column gives [] and never throws', ctx.extractWhatsWorking_('<div>nothing</div>').length === 0 && ctx.extractWhatsWorking_(null).length === 0 && ctx.extractWhatsWorking_(undefined).length === 0);
  check('a column cut off in the middle still gives what is there', ctx.extractWhatsWorking_('<div class="report-col report-col-working"><ul><li>One</li><li>Tw').join() === 'One');
}

console.log('\nTest 7 — text is cut at a word end and never overflows');
{
  check('a short text is untouched', ctx._scTruncate_('Short text', 50) === 'Short text');
  const t = ctx._scTruncate_('Ask one open-ended discovery question at the start of every single call', 40);
  check('a long text ends at a word with an ellipsis, within the limit', t.length <= 40 && /…$/.test(t) && !/\s…$/.test(t), t);
  const f = ctx._slFitText_('word '.repeat(80), 100, 11, 2);
  check('_slFitText_ keeps the text inside the line budget', ctx._slLines_(f, 100, 11) <= 2 && /…$/.test(f), f);
  check('a single word longer than a line is still cut', ctx._slLines_(ctx._slFitText_('x'.repeat(300), 100, 11, 2), 100, 11) <= 2);
}

console.log('\nTest 8 — the look: contrast, font size, and nothing spills');
{
  const T = ctx.SCORECARD_THEME;
  const pairs = [[T.purple, T.white], [T.purple, T.lilac], [T.white, T.purple], [T.brightGreen, T.purple], [T.green, T.white], [T.green, T.lilac], [T.red, T.white], [T.red, T.lilac],
                 [T.amber, T.white], [T.white, T.green], [T.white, T.amber], [T.white, T.red], [T.white, T.grey], [T.grey, T.white], [T.grey, T.lilac], [T.ink, T.white], [T.ink, T.lilac]];
  const low = pairs.filter(([a, b]) => ctx._slContrast_(a, b) < 4.5);
  check('every text/background colour pair has contrast >= 4.5:1', low.length === 0, JSON.stringify(low.map(([a, b]) => a + ' on ' + b + ' = ' + ctx._slContrast_(a, b).toFixed(2))));

  const LONG = 'Customer was transferred without a warm handoff and the agent did not set any expectation for the next step on the call '.repeat(3);
  const worst = Object.assign(one([row({ 'Team Member': 'Maria Guadalupe Hernandez-Villanueva de la Cruz Jr' })]), {});
  worst.name = 'Maria Guadalupe Hernandez-Villanueva de la Cruz Jr Third';
  worst.strengths = [LONG, LONG]; worst.gaps = [LONG, LONG, LONG]; worst.alsoNoted = LONG; worst.coaching = [LONG, LONG, LONG];
  worst.teamLeader = 'A very long team leader name that goes on'; worst.lob = 'MOB CxSS and some more text'; worst.locale = 'TI El Salvador (Merliot) and more'; worst.vtid = 'VTI2026000103';
  worst.tiles.forEach(t => { t.display = '100%'; t.target = 'Target <= 30%'; });
  const typical = one([row({ 'Critical Flags': 'No save attempt | Generic offer', 'Top SMART Opportunity': 'Ask one discovery question' })]);
  const noData = one([row({ 'Repeat Projection %': '', 'Issue Resolved': '', 'Transfer Occurred': '' })]);
  const salesC = one([row({ 'Analysis Type': 'Sales Analyzer', 'Sales Attempted': 'No' })]);
  const decks = { worst, typical, noData, salesC };

  const bgAt = (shapes, i) => {
    const s = shapes[i];
    if (s.fill) return s.fill;      // text inside a filled shape sits on that fill
    for (let j = i - 1; j >= 0; j--) {
      const b = shapes[j];
      if (b.fill && b.x <= s.x + 0.5 && b.y <= s.y + 0.5 && b.x + b.w >= s.x + s.w - 0.5 && b.y + b.h >= s.y + s.h - 0.5) return b.fill;
    }
    return T.white;
  };
  Object.keys(decks).forEach(name => {
    const shapes = ctx.scorecardSlideShapes_(decks[name], 2, 5);
    const ids = shapes.map(s => s.id);
    check(name + ': every shape id is unique', new Set(ids).size === ids.length);
    const outside = shapes.filter(s => s.x < 0 || s.y < 0 || s.x + s.w > 720.01 || s.y + s.h > 405.01);
    check(name + ': every shape is inside the 720 x 405 page', outside.length === 0, outside.map(s => s.id).join());
    const small = [];
    const overflow = [];
    const lowContrast = [];
    shapes.forEach((s, i) => {
      const bg = bgAt(shapes, i);
      s.paras.forEach(p => {
        p.runs.forEach(r => {
          const minFont = (s.id === 'footer' || s.id === 'page') ? ctx.SCORECARD_FOOTER_FONT : ctx.SCORECARD_MIN_FONT;
          if (r.size < minFont) small.push(s.id + ' ' + r.size);
          if (ctx._slContrast_(r.color, bg) < 4.5) lowContrast.push(s.id + ' ' + r.color + ' on ' + bg);
        });
      });
      if (s.paras.length) {
        const inner = s.w - 2 * s.pad;
        let need = 2 * s.pad;
        s.paras.forEach(p => {
          const size = Math.max.apply(null, p.runs.map(r => r.size));
          const w = inner - (p.bullet ? 14 : 0);
          need += ctx._slHeight_(ctx._slLines_(p.runs.map(r => r.t).join(''), w, size), size, p.spacing);
        });
        if (need > s.h + 0.5) overflow.push(s.id + ' needs ' + need.toFixed(0) + ' has ' + s.h);
      }
    });
    check(name + ': no font below 11 pt (footer 9 pt)', small.length === 0, small.join());
    check(name + ': every run has contrast >= 4.5:1 against what is behind it', lowContrast.length === 0, lowContrast.join());
    check(name + ': no text overflows its box', overflow.length === 0, overflow.join('; '));
    const boxes = shapes.filter(s => /^(tile|sup)-/.test(s.id));
    let overlap = '';
    for (let a = 0; a < boxes.length; a++) for (let b = a + 1; b < boxes.length; b++) {
      const A = boxes[a], B = boxes[b];
      if (A.x < B.x + B.w && B.x < A.x + A.w && A.y < B.y + B.h && B.y < A.y + A.h) overlap += A.id + '/' + B.id + ' ';
    }
    check(name + ': tiles and supporting boxes do not overlap', !overlap, overlap);
    const cardOf = { 'tile': 'card-score', 'sup': 'card-score', 'feedback': 'card-feedback', 'coaching': 'card-coaching' };
    const bad = shapes.filter(s => { const k = s.id.replace(/-\d+$/, ''); const c = cardOf[k] && shapes.find(x => x.id === cardOf[k]); return c && (s.x < c.x || s.y < c.y || s.x + s.w > c.x + c.w || s.y + s.h > c.y + c.h); });
    check(name + ': content stays inside its card', bad.length === 0, bad.map(s => s.id).join());
  });

  const cover = ctx.scorecardCoverShapes_({ label: 'MOB CLS Batch 15', range: 'Oct 1 - Oct 13, 2026', agents: 12, audits: 87, filters: 'Team Leader: TL One; Locale: TI Philippines',
    thresholds: Object.keys(ctx.SCORECARD_THRESHOLDS).map(k => ({ label: ctx.SCORECARD_THRESHOLDS[k].label, target: 'Target' })) }, 13);
  const coverOver = [];
  cover.forEach(s => { if (s.paras.length) { let need = 2 * s.pad; s.paras.forEach(p => { const size = Math.max.apply(null, p.runs.map(r => r.size)); need += ctx._slHeight_(ctx._slLines_(p.runs.map(r => r.t).join(''), s.w - 2 * s.pad, size), size, p.spacing); }); if (need > s.h + 0.5) coverOver.push(s.id + ' ' + need.toFixed(0) + '>' + s.h); } });
  check('cover slide: nothing overflows', coverOver.length === 0, coverOver.join('; '));
  check('cover slide: inside the page', cover.every(s => s.x >= 0 && s.y >= 0 && s.x + s.w <= 720.01 && s.y + s.h <= 405.01));

  const slide = ctx.scorecardSlideShapes_(bad0(), 2, 5);
  function bad0() { return one([row({ 'Repeat Projection %': '90%', 'Issue Resolved': 'No', 'Critical Flags': 'F1', 'Transfer Occurred': 'Yes' })]); }
  const words = JSON.stringify(slide);
  check('status is said in words on a failing slide: "Below threshold" and the badge text', words.indexOf('Below threshold') !== -1 && words.indexOf('Priority support') !== -1);
  const fine = JSON.stringify(ctx.scorecardSlideShapes_(one([row({ 'Repeat Projection %': '10%', 'Critical Flags': '' })]), 2, 5));
  check('status is said in words on a passing slide: "On target" and "On track"', fine.indexOf('On target') !== -1 && fine.indexOf('On track') !== -1);
  check('no emoji or symbol outside Arial is used', !/[✓✕✔✖▲▼]/.test(words + fine));
}

console.log('\nTest 7 — a small number of flags is not rounded away');
{
  const many = n => Array.from({ length: n }, (_, i) => row({ 'Audit Ref': 'F' + i, 'Critical Flags': i === 0 ? 'One flag' : '' }));
  const c21 = one(many(21));
  check('1 flag in 21 audits is Below threshold (the target is 0), not "On target" because 0.048 rounds to 0.0', tile(c21, 'flags').status === 'below', JSON.stringify(tile(c21, 'flags')));
  check('and it is shown as <0.1, never as 0.0 next to "Below threshold"', tile(c21, 'flags').display === '<0.1', tile(c21, 'flags').display);
  const c0 = one(many(21).map(r => Object.assign(r, { 'Critical Flags': '' })));
  check('no flags at all is On target and shown 0.0', tile(c0, 'flags').status === 'target' && tile(c0, 'flags').display === '0.0', JSON.stringify(tile(c0, 'flags')));
  const c2 = one([row({ 'Critical Flags': 'A | B' }), row({ 'Critical Flags': 'C' })]);
  check('1.5 flags per audit is shown as 1.5 and is below', tile(c2, 'flags').display === '1.5' && tile(c2, 'flags').status === 'below');
  check('the sub-text still counts the real flags', /1 flag in 21 audits/.test(tile(c21, 'flags').sub), tile(c21, 'flags').sub);
}

console.log('\nTest 8 — dates are in the script time zone, not UTC');
{
  // 8 PM New York on Oct 13 is 00:00 UTC on Oct 14. With UTC the audit moved to the next day.
  const sb = Object.assign({}, sandbox, {
    Session: { getScriptTimeZone: () => 'America/New_York' },
    Utilities: { formatDate: (d, tz, fmt) => {
      const p = new Intl.DateTimeFormat('en-US', { timeZone: tz, year: 'numeric', month: 'numeric', day: 'numeric' }).formatToParts(d);
      const g = t => p.find(x => x.type === t).value;
      return g('year') + '-' + g('month') + '-' + g('day');
    } },
  });
  const c2 = vm.createContext(sb);
  ['Config.gs', 'ScorecardMetrics.gs'].forEach(f => vm.runInContext(read(f), c2, { filename: f }));
  const late = Date.parse('2026-10-14T00:30:00Z');       // 8:30 PM on Oct 13 in New York
  check('an audit at 8:30 PM ET on Oct 13 is Oct 13', c2._scDate_(late) === 'Oct 13', c2._scDate_(late));
  check('a range ending at that audit says Oct 13', c2._scDateRange_(Date.parse('2026-10-01T15:00:00Z'), late) === 'Oct 1 - Oct 13, 2026', c2._scDateRange_(Date.parse('2026-10-01T15:00:00Z'), late));
  const nye = Date.parse('2027-01-01T03:00:00Z');        // Dec 31, 10 PM ET
  check('the year follows the same zone (Dec 31, 2026, not Jan 1, 2027)', c2._scDateRange_(nye, nye) === 'Dec 31, 2026', c2._scDateRange_(nye, nye));
  check('without Apps Script (these tests) it falls back to UTC and still works', ctx._scDate_(late) === 'Oct 14', ctx._scDate_(late));
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
