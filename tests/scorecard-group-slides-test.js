/**
 * The slides about the whole group (registry 189): cover, executive summary, scorecard results, where the agents landed,
 * heatmap, audit findings, readiness, next steps.
 *
 *  1. The numbers (averages, counts at each level, readiness, gaps, next steps) on an example worked out by hand.
 *  2. The deck: the order of the slides, the page numbers, one heatmap slide per 14 agents, one agent slide each.
 *  3. The look, on every kind of deck (typical, one agent, 40 agents with very long names, no data, a Sales agent):
 *     inside the page, unique ids, no font under 11 pt (footer 9), contrast 4.5:1, no text outside its box.
 *  4. What each slide says: heatmap colours follow the status, no agent is lost, "+ N more" is right, bars are to scale.
 *  5. The package: the new circle shape is a real ellipse in the PowerPoint file.
 *
 * Usage: node tests/scorecard-group-slides-test.js
 */
const vm = require('vm');
const { read, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

const ctx = vm.createContext({ Logger: { log() {} }, JSON, Array, Object, String, Number, Math, Date, RegExp, isFinite, isNaN, Error,
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) } });
['Config.gs', 'ScorecardMetrics.gs', 'ScorecardLayout.gs', 'ScorecardPptx.gs'].forEach(f => vm.runInContext(read(f), ctx, { filename: f }));
const T = ctx.SCORECARD_THEME, P = ctx.SCORECARD_PAGE;

let n = 0;
const row = o => Object.assign({ 'Audit Ref': 'G' + (++n), 'Timestamp': '2026-10-05T10:00:00Z', 'Team Member': 'Agent A', 'SAP ID': '100', 'Analysis Type': 'Repeats & Transfer Analyzer',
  'Repeat Projection %': '10%', 'Issue Resolved': 'Yes', 'Critical Flags': '', 'Transfer Occurred': 'No', 'Sales Attempted': '', 'Email Status': 'Sent', 'Team Leader': 'TL One',
  'Line of Business': 'MOB CLS', 'Locale': 'TI Philippines', 'VTID': 'V1', 'Top SMART Opportunity': '', 'RCA Sub-Parameter': '' }, o);
const build = (rows, extras) => ctx.buildAgentCards_(rows, extras || null);
const tile = (card, key) => card.tiles.find(t => t.key === key);
const text = shapes => shapes.map(s => s.paras.map(p => p.runs.map(r => r.t).join('')).join(' ')).join(' | ');
const titleOf = shapes => { const t = shapes.find(s => s.id === 'title') || shapes.find(s => s.id === 'name'); return t ? t.paras[0].runs.map(r => r.t).join('') : ''; };

console.log('\nTest 1 — the group numbers, on an example worked out by hand');
{
  // three agents, one audit each: repeat risk 10 / 20 / 60, resolved yes / yes / no, flags 0 / 0 / 2, transfer no / no / yes
  const rows = [
    row({ 'Team Member': 'Ana', 'Repeat Projection %': '10%', 'Issue Resolved': 'Yes', 'Critical Flags': '', 'Transfer Occurred': 'No', 'Critical Flags ': '' }),
    row({ 'Team Member': 'Ben', 'Repeat Projection %': '20%', 'Issue Resolved': 'Yes', 'Critical Flags': '', 'Transfer Occurred': 'No' }),
    row({ 'Team Member': 'Cy', 'Repeat Projection %': '60%', 'Issue Resolved': 'No', 'Critical Flags': 'Dead air | Interrupting', 'Transfer Occurred': 'Yes' }),
  ];
  const g = ctx.buildGroupSummary_(build(rows), 'Batch 1', 'Team Leader: TL One', 3);
  const k = key => g.kpis.find(x => x.key === key);
  check('three KPIs have numbers (no Sales agent, so no Sales attempted)', g.kpis.map(x => x.key).join() === 'repeatRisk,resolution,flags,transfer', g.kpis.map(x => x.key).join());
  check('repeat risk: the mean of 10, 20, 60 is 30%, which meets the 30% limit', k('repeatRisk').display === '30%' && k('repeatRisk').status === 'target', JSON.stringify(k('repeatRisk')));
  check('repeat risk: 2 of 3 agents on target, 1 below', k('repeatRisk').onTarget === 2 && k('repeatRisk').below === 1 && k('repeatRisk').measured === 3);
  check('resolution: the mean of 100, 100, 0 is 67%, below the 80% target', k('resolution').display === '67%' && k('resolution').status === 'below');
  check('flags: the mean of 0, 0, 2 is 0.7 per audit, below target (the target is 0)', k('flags').display === '0.7' && k('flags').status === 'below');
  check('transfer: the mean of 0, 0, 100 is 33%, below the 20% limit', k('transfer').display === '33%' && k('transfer').status === 'below');
  check('the target is said in words and in short', k('resolution').target === 'Target >= 80%' && k('resolution').targetShort === '>= 80%' && k('flags').targetShort === '<= 0/audit');
  check('readiness: Ana and Ben are On track, Cy needs Priority support', g.counts.ontrack === 2 && g.counts.priority === 1 && g.counts.attention === 0 && g.levels.priority[0].name === 'Cy', JSON.stringify(g.counts));
  check('the reason names the KPIs that are below, in short', /^Below: /.test(g.levels.priority[0].reason) && /Resolve/.test(g.levels.priority[0].reason) && g.levels.ontrack[0].reason === 'All measured KPIs on target', g.levels.priority[0].reason);
  check('gaps: each flag is counted once per agent', g.gaps.length === 2 && g.gaps.every(x => x.count === 1), JSON.stringify(g.gaps));
  // each of the four KPIs has one agent below its target (Cy), so four KPI cards, then the priority agents
  check('next steps: a card for each KPI that has an agent below target, then the priority agents', g.nextSteps.length === 5 && g.nextSteps[4].title === 'Priority support' && /Cy/.test(g.nextSteps[4].body) &&
        ['Avg Repeat Risk', 'Resolution rate', 'Critical flags', 'Transfer rate'].every(t => g.nextSteps.some(s => s.title === t)), JSON.stringify(g.nextSteps.map(s => s.title)));
  check('every next-step body says how many of how many agents', g.nextSteps.slice(0, 3).every(s => /^\d+ of 3 agents are (above|below) the /.test(s.body)), JSON.stringify(g.nextSteps.map(s => s.body)));

  const two = [row({ 'Team Member': 'Ana', 'Critical Flags': 'Dead air | Interrupting' }), row({ 'Team Member': 'Ben', 'Critical Flags': 'Dead air' }), row({ 'Team Member': 'Ben', 'Critical Flags': 'dead  AIR' })];
  const g2 = ctx.buildGroupSummary_(build(two), '', '', 3);
  check('a gap two agents share counts 2, even with other capitals and spaces; an agent with it twice counts once', g2.gaps[0].count === 2 && /dead air/i.test(g2.gaps[0].label), JSON.stringify(g2.gaps));

  const sales = [row({ 'Team Member': 'Sam', 'Analysis Type': 'Sales Analyzer', 'Sales Attempted': 'Yes' }), row({ 'Team Member': 'Ann' })];
  const g3 = ctx.buildGroupSummary_(build(sales), '', '', 2);
  const sk = g3.kpis.find(x => x.key === 'salesAttempt'), tk = g3.kpis.find(x => x.key === 'transfer');
  check('a Sales agent is counted only in Sales attempted, a Repeats agent only in Transfer', sk.applicable === 1 && tk.applicable === 1, JSON.stringify([sk.applicable, tk.applicable]));

  const none = ctx.buildGroupSummary_(build([row({ 'Repeat Projection %': '', 'Issue Resolved': '', 'Transfer Occurred': '', 'Critical Flags': '' })].map(r => Object.assign(r, { 'Critical Flags': '' }))), '', '', 1);
  check('a group with nothing measurable still gives a summary (flags are 0 per audit, so only that KPI)', Array.isArray(none.kpis) && none.nextSteps.length >= 1);
}

console.log('\nTest 2 — the order of the slides and the page numbers');
{
  const mk = count => build(Array.from({ length: count }, (_, i) => row({ 'Team Member': 'Agent ' + String.fromCharCode(65 + (i % 26)) + (i >= 26 ? i : ''), 'SAP ID': String(100 + i) })));
  [[1, 1], [14, 1], [15, 2], [28, 2], [29, 3], [40, 3]].forEach(([agents, heat]) => {
    const built = mk(agents), slides = ctx.scorecardDeckShapes_(built, 'Batch', 'filters', agents, []);
    const want = 4 + heat + 2 + agents + 1;
    check(agents + ' agent(s): ' + want + ' slides (' + heat + ' heatmap slide' + (heat === 1 ? '' : 's') + ')', slides.length === want, 'got ' + slides.length);
    const titles = slides.map(titleOf);
    check(agents + ' agent(s): cover, summary, results, landed, heatmap(s), findings, readiness, the agents, next steps',
          titles[0] === 'Audit Scorecards' && titles[1] === 'Executive summary' && titles[2] === 'Scorecard results' && titles[3] === 'Where the agents landed on each KPI' &&
          titles.slice(4, 4 + heat).every(t => /^Scorecard heatmap by agent/.test(t)) && titles[4 + heat] === 'Audit findings: strengths and gaps' && titles[5 + heat] === 'Readiness' &&
          titles[titles.length - 1] === 'Next steps' && titles.slice(6 + heat, titles.length - 1).length === agents, titles.slice(0, 8).join(' | '));
    check(agents + ' agent(s): the page numbers run 1 to ' + want + ', each "n / ' + want + '"', slides.every((sl, i) => JSON.stringify(sl.find(s => s.id === 'page')).indexOf((i + 1) + ' / ' + want) !== -1));
  });
  const three = ctx.scorecardDeckShapes_(mk(29), 'B', '', 29, []);
  check('with several heatmap slides each title says which one ("1 of 3")', /\(1 of 3\)/.test(titleOf(three[4])) && /\(3 of 3\)/.test(titleOf(three[6])));
}

console.log('\nTest 3 — the look: inside the page, ids, font size, contrast, nothing spills');
{
  const LONGNAME = 'Maria Guadalupe Hernandez-Villanueva de la Cruz Jr Third';
  const LONG = 'Customer was transferred without a warm handoff and the agent did not set any expectation for the next step on the call';
  const many = Array.from({ length: 40 }, (_, i) => row({ 'Team Member': i % 3 === 0 ? LONGNAME + ' ' + i : 'Agent Number ' + i, 'SAP ID': String(200 + i),
    'Repeat Projection %': ['10%', '45%', '80%'][i % 3], 'Issue Resolved': ['Yes', 'No', 'Yes'][i % 3], 'Critical Flags': i % 2 ? LONG + ' | Another long flag that goes on and on for a while' : '',
    'Transfer Occurred': ['No', 'Yes', 'No'][i % 3], 'Top SMART Opportunity': LONG }));
  const extras = { whatsWorking: {} }; many.forEach(r => { extras.whatsWorking[r['Audit Ref']] = [LONG + ' and more words to make it long']; });
  const decks = {
    typical: build([row({ 'Team Member': 'Ana' }), row({ 'Team Member': 'Ben', 'Repeat Projection %': '70%', 'Issue Resolved': 'No', 'Critical Flags': 'Dead air' }),
      row({ 'Team Member': 'Cy', 'Analysis Type': 'Sales Analyzer', 'Sales Attempted': 'No' }), row({ 'Team Member': 'Di', 'Repeat Projection %': '', 'Issue Resolved': '', 'Transfer Occurred': '' })]),
    one: build([row({})]),
    forty: build(many, extras),
    nodata: build([row({ 'Repeat Projection %': '', 'Issue Resolved': '', 'Transfer Occurred': '', 'Analysis Type': 'Other' })]),
  };
  const pairs = [[T.white, T.purple], [T.onDark, T.purple], [T.purple, T.lilac], [T.purple, T.white], [T.ink, T.lilac], [T.ink, T.white], [T.grey, T.white], [T.grey, T.lilac],
                 [T.green, T.lilac], [T.red, T.lilac], [T.amber, T.lilac], [T.grey, T.lilac], [T.white, T.green], [T.white, T.red], [T.white, T.grey], [T.white, T.amber]];
  const low = pairs.filter(([a, b]) => ctx._slContrast_(a, b) < 4.5);
  check('the colours used for text have contrast >= 4.5:1 (including the new cover colours)', low.length === 0, JSON.stringify(low.map(([a, b]) => a + ' on ' + b + ' = ' + ctx._slContrast_(a, b).toFixed(2))));

  const bgAt = (shapes, i) => {
    const s = shapes[i];
    if (s.fill) return s.fill;
    for (let j = i - 1; j >= 0; j--) {
      const b = shapes[j];
      if (b.fill && b.x <= s.x + 0.5 && b.y <= s.y + 0.5 && b.x + b.w >= s.x + s.w - 0.5 && b.y + b.h >= s.y + s.h - 0.5) return b.fill;
    }
    return T.white;
  };
  Object.keys(decks).forEach(name => {
    const built = decks[name], slides = ctx.scorecardDeckShapes_(built, 'MOB CLS Sample Deck ' + 'x'.repeat(40), 'Team Leader: ' + LONG.slice(0, 80), built.total, []);
    const problems = { outside: [], dup: [], small: [], contrast: [], overflow: [], bg: [] };
    slides.forEach((shapes, si) => {
      const tag = name + ' slide ' + (si + 1) + ' (' + titleOf(shapes).slice(0, 24) + ')';
      const ids = shapes.map(s => s.id);
      if (new Set(ids).size !== ids.length) problems.dup.push(tag);
      if (!shapes.length || shapes[0].id !== 'bg') problems.bg.push(tag);
      shapes.forEach((s, i) => {
        if (s.x < -0.01 || s.y < -0.01 || s.x + s.w > P.w + 0.01 || s.y + s.h > P.h + 0.01) problems.outside.push(tag + ' ' + s.id);
        const bg = bgAt(shapes, i);
        s.paras.forEach(p => p.runs.forEach(r => {
          const minFont = (s.id === 'footer' || s.id === 'page') ? ctx.SCORECARD_FOOTER_FONT : ctx.SCORECARD_MIN_FONT;
          if (r.size < minFont) problems.small.push(tag + ' ' + s.id + ' ' + r.size);
          if (ctx._slContrast_(r.color, bg) < 4.5) problems.contrast.push(tag + ' ' + s.id + ' ' + r.color + ' on ' + bg);
        }));
        if (s.paras.length) {
          const inner = s.w - 2 * s.pad;
          let need = 2 * s.pad;
          s.paras.forEach(p => {
            const size = Math.max.apply(null, p.runs.map(r => r.size));
            need += ctx._slHeight_(ctx._slLines_(p.runs.map(r => r.t).join(''), inner - (p.bullet ? 14 : 0), size), size, p.spacing);
          });
          if (need > s.h + 0.5) problems.overflow.push(tag + ' ' + s.id + ' needs ' + need.toFixed(0) + ' has ' + s.h);
        }
      });
    });
    check(name + ' (' + slides.length + ' slides): every shape is inside the 720 x 405 page', problems.outside.length === 0, problems.outside.slice(0, 4).join('; '));
    check(name + ': every shape id is unique on its slide, and the first shape is the background', problems.dup.length === 0 && problems.bg.length === 0, problems.dup.concat(problems.bg).slice(0, 3).join('; '));
    check(name + ': no font below 11 pt (footer 9 pt)', problems.small.length === 0, problems.small.slice(0, 4).join('; '));
    check(name + ': every run has contrast >= 4.5:1 against what is behind it', problems.contrast.length === 0, problems.contrast.slice(0, 4).join('; '));
    check(name + ': no text overflows its box', problems.overflow.length === 0, problems.overflow.slice(0, 4).join('; '));
  });
}

console.log('\nTest 4 — what the slides say');
{
  const rows = [
    row({ 'Team Member': 'Ana', 'Repeat Projection %': '10%' }), row({ 'Team Member': 'Ben', 'Repeat Projection %': '70%', 'Issue Resolved': 'No', 'Critical Flags': 'Dead air' }),
    row({ 'Team Member': 'Cy', 'Analysis Type': 'Sales Analyzer', 'Sales Attempted': 'No' }), row({ 'Team Member': 'Di', 'Repeat Projection %': '', 'Issue Resolved': '', 'Transfer Occurred': '' }),
  ];
  const built = build(rows), g = ctx.buildGroupSummary_(built, 'B', '', 4);
  const heat = ctx.scorecardHeatmapShapes_(built.cards, 0, 1, 5, 12, g);
  const fillOf = id => (heat.find(s => s.id === id) || {}).fill;
  check('the heatmap has a row for each of the 4 agents, in order', ['l-name-0', 'l-name-1', 'l-name-2', 'l-name-3'].map(id => text([heat.find(s => s.id === id)])).join() === 'Ana,Ben,Cy,Di', ['l-name-0', 'l-name-1', 'l-name-2', 'l-name-3'].map(id => text([heat.find(s => s.id === id)])).join());
  check('a cell is green on target, red below, grey with no data (Ana repeat 10% is green, Ben repeat 70% is red, Di blank is grey)', fillOf('l-c-0-0') === T.green && fillOf('l-c-1-0') === T.red && fillOf('l-c-3-0') === T.grey,
        [fillOf('l-c-0-0'), fillOf('l-c-1-0'), fillOf('l-c-3-0')].join());
  check('every cell of the heatmap says its number in text (never colour alone)', heat.filter(s => /^l-c-/.test(s.id)).every(s => /\S/.test(text([s]))));
  check('the heatmap says what Xfer / Sales means and what the colours are', /Transfer rate for Repeats agents/.test(text(heat)) && /On target/.test(text(heat)) && /Below threshold/.test(text(heat)) && /No data/.test(text(heat)));

  const sev = Array.from({ length: 10 }, (_, i) => row({ 'Team Member': 'Agent ' + String.fromCharCode(65 + i), 'Repeat Projection %': '90%', 'Issue Resolved': 'No', 'Critical Flags': 'F', 'Transfer Occurred': 'Yes' }));
  const gs = ctx.buildGroupSummary_(build(sev), '', '', 10), ready = ctx.scorecardReadinessShapes_(gs, 7, 12);
  const list = ready.find(s => s.id === 'col-priority-list');
  check('10 agents in one column: 6 are listed, then "+ 4 more"', list.paras.filter(p => p.runs[0].r === undefined && p.runs[0].bold && p.runs[0].size === 12).length === 6 && /\+ 4 more/.test(text([list])), text([list]).slice(-60));
  check('the column head counts all 10', /Priority support \(10\)/.test(text(ready)));

  const res = ctx.scorecardResultsShapes_(g, 3, 12), kr = g.kpis.find(k => k.key === 'resolution');
  const avgBar = res.find(s => s.id === 'bar-1-avg'), tgtBar = res.find(s => s.id === 'bar-1-target');
  check('the chart bars are to scale (1.7 pt per percent): average and target', Math.abs(avgBar.h - kr.avg * 1.7) < 0.01 && Math.abs(tgtBar.h - 80 * 1.7) < 0.01, avgBar.h + ' / ' + tgtBar.h);
  check('both bars stand on the same baseline', Math.abs((avgBar.y + avgBar.h) - (tgtBar.y + tgtBar.h)) < 0.01);
  check('the table lists each KPI with its target, average and "n of m"', g.kpis.every((k, r) => text([res.find(s => s.id === 'td-' + r + '-3')]) === k.onTarget + ' of ' + k.measured && text([res.find(s => s.id === 'td-' + r + '-2')]) === k.display));
  check('flags are not in the chart (a count per audit, not a percent), and the slide says so', !res.some(s => /^cat-/.test(s.id) && /Flags/.test(text([s]))) && /not in the chart/.test(text(res)));

  const landed = ctx.scorecardLandedShapes_(g, 4, 12), kRep = g.kpis.find(k => k.key === 'repeatRisk');
  const seg = (i, s) => landed.find(x => x.id === 'lb-' + i + '-' + s);
  check('a stacked bar is as wide as its KPI and each part is in proportion', Math.abs(['below', 'on', 'none'].reduce((w, s) => w + (seg(0, s) ? seg(0, s).w : 0), 0) - 296) < 0.01 &&
        Math.abs(seg(0, 'below').w - 296 * kRep.below / kRep.applicable) < 0.01, JSON.stringify(['below', 'on', 'none'].map(s => seg(0, s) && seg(0, s).w)));

  const cover = ctx.scorecardCoverShapes_({ label: 'B', range: 'Oct 1 - Oct 9, 2026', agents: 4, audits: 22, filters: 'x' }, 12);
  check('the cover is dark purple with two circles and no legend any more', cover[0].fill === T.purple && cover.filter(s => s.type === 'ELLIPSE').length === 2 && !/How to read/.test(text(cover)));
  check('the cover says the title, the period, the counts and the filters', /Audit Scorecards/.test(text(cover)) && /Oct 1 - Oct 9, 2026/.test(text(cover)) && /4 agents/.test(text(cover)) && /Filters: x/.test(text(cover)));
  check('the targets are said on the results slide (so "Below threshold" is still explained)', /Target/.test(text(res)) && /<= 30%/.test(text(res)) && /Badge: no KPI below threshold/.test(text(res)));
}

console.log('\nTest 5 — the PowerPoint package draws the circles as ellipses');
{
  const built = build([row({ 'Team Member': 'Ana' }), row({ 'Team Member': 'Ben', 'Repeat Projection %': '70%' })]);
  const slides = ctx.scorecardDeckShapes_(built, 'B', 'f', 2, []);
  const parts = ctx.buildPptxParts_(slides), by = {}; parts.forEach(p => { by[p.path] = p.xml; });
  const wantEllipses = slides.reduce((c, sl) => c + sl.filter(s => s.type === 'ELLIPSE').length, 0), got = parts.reduce((c, p) => c + (p.xml.match(/prst="ellipse"/g) || []).length, 0);
  check('there is an ellipse in the file for every ELLIPSE in the layout (' + wantEllipses + ')', wantEllipses > 0 && got === wantEllipses, got + ' vs ' + wantEllipses);
  check('an ellipse is not a text box, and it keeps its fill', slides.every((sl, si) => sl.filter(s => s.type === 'ELLIPSE').every(s => {
    const x = by['ppt/slides/slide' + (si + 1) + '.xml'], i = x.indexOf('name="' + s.id + '"'), end = x.indexOf('</p:sp>', i), sp = x.slice(i, end);
    return sp.indexOf('txBox') === -1 && sp.indexOf('prst="ellipse"') !== -1 && sp.indexOf('srgbClr val="' + s.fill.replace('#', '').toUpperCase() + '"') !== -1;
  })));
  check('the package has one slide file per slide', slides.every((_, i) => !!by['ppt/slides/slide' + (i + 1) + '.xml']) && !by['ppt/slides/slide' + (slides.length + 1) + '.xml']);
  const want = []; slides.forEach(sl => sl.filter(s => s.id !== 'bg').forEach(s => s.paras.forEach(p => p.runs.forEach(r => { if (r.t.length) want.push(r.t); }))));
  const gotText = parts.filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p.path)).reduce((c, p) => c + (p.xml.match(/<a:t>/g) || []).length, 0);
  check('every piece of text of the whole deck is in the package (' + want.length + ')', gotText === want.length, gotText + ' vs ' + want.length);
}

console.log('\n' + (check.state.failures === 0 ? 'ALL PASSED' : check.state.failures + ' FAILED'));
process.exit(check.state.failures === 0 ? 0 : 1);
