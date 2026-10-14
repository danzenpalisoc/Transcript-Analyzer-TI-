/**
 * The PowerPoint package the scorecard deck is made from (ScorecardPptx.gs). Drive converts it to Google Slides, so it has to be a
 * clean, complete package, and it has to carry exactly what the approved layout says.
 *
 *  1. Every part is well-formed XML (a small strict checker below, since a stray character would make Drive refuse the file).
 *  2. The package is complete: content types, relationships and slide list agree with the parts that exist.
 *  3. Positions, sizes, fills, corners, margins, fonts, colours, spacing and bullets are written as the layout says.
 *  4. Text is escaped, and characters XML cannot hold are dropped instead of breaking the file.
 *
 * Usage: node tests/scorecard-pptx-test.js
 */
const vm = require('vm');
const { read, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

const ctx = vm.createContext({ Logger: { log() {} }, JSON, Array, Object, String, Number, Math, Date, RegExp, isFinite, isNaN, Error, parseInt,
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) } });
['Config.gs', 'ScorecardMetrics.gs', 'ScorecardLayout.gs', 'ScorecardPptx.gs'].forEach(f => vm.runInContext(read(f), ctx, { filename: f }));

// A strict well-formedness check: one root, tags balanced, attributes quoted, only the five entities and numeric references.
function xmlError(xml) {
  if (!/^<\?xml version="1\.0" encoding="UTF-8" standalone="yes"\?>\n/.test(xml)) return 'missing or wrong XML declaration';
  const body = xml.replace(/^<\?xml[^>]*\?>\n/, '');
  const re = /<(\/?)([A-Za-z_][\w:.-]*)((?:\s+[A-Za-z_][\w:.-]*="[^"<]*")*)\s*(\/?)>|([^<]+)/g;
  const validEnt = s => !/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(s);
  const stack = []; let m, pos = 0, roots = 0;
  while ((m = re.exec(body)) !== null) {
    if (m.index !== pos) return 'unparsable text near character ' + pos;
    pos = re.lastIndex;
    if (m[5] !== undefined) { if (!validEnt(m[5])) return 'bad entity in text: ' + m[5].slice(0, 40); if (stack.length === 0 && m[5].trim()) return 'text outside the root'; continue; }
    if (m[3] && !validEnt(m[3])) return 'bad entity in an attribute of <' + m[2] + '>';
    if (m[1]) { const top = stack.pop(); if (top !== m[2]) return 'closing </' + m[2] + '> but <' + top + '> is open'; }
    else if (!m[4]) { if (stack.length === 0) roots++; stack.push(m[2]); }
    else if (stack.length === 0) roots++;
  }
  if (pos !== body.length) return 'trailing characters';
  if (stack.length) return 'unclosed <' + stack[stack.length - 1] + '>';
  if (roots !== 1) return roots + ' root elements';
  return '';
}
const unesc = s => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (a, n) => String.fromCharCode(Number(n))).replace(/&amp;/g, '&');

// real slides from the real layout
let n = 0;
const row = o => Object.assign({ 'Audit Ref': 'P' + (++n), 'Timestamp': '2026-10-05T10:00:00Z', 'Team Member': 'Agent A', 'SAP ID': '111', 'Analysis Type': 'Repeats & Transfer Analyzer',
  'Repeat Projection %': '40%', 'Issue Resolved': 'Yes', 'Critical Flags': 'No save attempt | Generic offer', 'Transfer Occurred': 'No', 'Email Status': 'Sent', 'Team Leader': 'TL One',
  'Line of Business': 'MOB CLS', 'Locale': 'TI Philippines', 'VTID': 'V1', 'Top SMART Opportunity': 'Ask one discovery question & confirm "next steps"', 'RCA Sub-Parameter': '' }, o);
const built = ctx.buildAgentCards_([row({}), row({ 'Team Member': 'Agent B <Sales> & Co', 'Analysis Type': 'Sales Analyzer', 'Sales Attempted': 'No' })], { whatsWorking: { P1: ['Warm greeting & clear ownership'] } });
const slides = [ctx.scorecardCoverShapes_({ label: 'Batch "15"', range: 'Oct 1 - Oct 5, 2026', agents: 2, audits: 2, filters: 'Team Leader: TL One', thresholds: [{ label: 'Avg Repeat Risk', target: '30% or lower' }] }, 3)]
  .concat(built.cards.map((c, i) => ctx.scorecardSlideShapes_(c, i + 2, 3)));
const parts = ctx.buildPptxParts_(slides);
const byPath = {}; parts.forEach(p => { byPath[p.path] = p.xml; });

console.log('\nTest 1 — every part is well-formed XML');
{
  const bad = parts.map(p => [p.path, xmlError(p.xml)]).filter(x => x[1]);
  check('all ' + parts.length + ' parts parse (declaration, one root, balanced tags, quoted attributes, legal entities)', bad.length === 0, JSON.stringify(bad));
  check('the checker itself catches a broken file (so the check above means something)', xmlError('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a><b></a>') !== '' && xmlError('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a>x & y</a>') !== '' && xmlError('<a/>') !== '');
  check('parts are named like a real package: [Content_Types].xml first, then _rels/.rels', parts[0].path === '[Content_Types].xml' && parts[1].path === '_rels/.rels');
  check('nothing external is referenced (no links, no scripts, no embedded objects)', !parts.some(p => /http:\/\/|https:\/\//.test(p.xml.replace(/xmlns(:\w+)?="[^"]*"/g, '').replace(/Type="[^"]*"/g, '').replace(/ContentType="[^"]*"/g, ''))));
}

console.log('\nTest 2 — the package is complete and consistent');
{
  const paths = parts.map(p => p.path);
  check('paths are unique', new Set(paths).size === paths.length);
  const ct = byPath['[Content_Types].xml'];
  const overrides = [...ct.matchAll(/PartName="([^"]+)"/g)].map(m => m[1].replace(/^\//, ''));
  check('every part that is not a .rels file has a content type, and every content type points at a part', paths.filter(p => !/\.rels$/.test(p) && p !== '[Content_Types].xml').every(p => overrides.indexOf(p) !== -1) && overrides.every(p => paths.indexOf(p) !== -1),
        JSON.stringify(overrides.filter(p => paths.indexOf(p) === -1)));
  const relTargets = [];
  parts.filter(p => /\.rels$/.test(p.path)).forEach(p => {
    const dir = p.path.replace(/_rels\/[^/]*$/, '');
    [...p.xml.matchAll(/Target="([^"]+)"/g)].forEach(m => {
      const segs = (dir + m[1]).split('/'), out = [];
      segs.forEach(s => { if (s === '..') out.pop(); else if (s !== '.' && s !== '') out.push(s); });
      relTargets.push([p.path, out.join('/')]);
    });
  });
  const missing = relTargets.filter(([, t]) => paths.indexOf(t) === -1);
  check('every relationship target exists in the package (' + relTargets.length + ' checked)', missing.length === 0, JSON.stringify(missing));
  check('the root relationship points at the presentation', /Target="ppt\/presentation\.xml"/.test(byPath['_rels/.rels']));
  const pres = byPath['ppt/presentation.xml'];
  const sldIds = [...pres.matchAll(/<p:sldId id="(\d+)" r:id="(rId\d+)"/g)];
  check('the slide list has one entry per slide, with unique ids from 256', sldIds.length === slides.length && new Set(sldIds.map(m => m[1])).size === sldIds.length && Number(sldIds[0][1]) === 256);
  const presRels = byPath['ppt/_rels/presentation.xml.rels'];
  check('each slide id points, through the presentation relationships, at its own slide file', sldIds.every((m, i) => new RegExp('Id="' + m[2] + '"[^>]*Target="slides/slide' + (i + 1) + '\\.xml"').test(presRels)));
  check('the page is 16:9, 9,144,000 x 5,143,500 EMU', /<p:sldSz cx="9144000" cy="5143500"\/>/.test(pres));
  check('every slide uses the blank layout and the layout uses the master', slides.every((s, i) => /slideLayout1\.xml/.test(byPath['ppt/slides/_rels/slide' + (i + 1) + '.xml.rels'])) && /slideMaster1\.xml/.test(byPath['ppt/slideLayouts/_rels/slideLayout1.xml.rels']));
  check('the theme has the 3 fill, 3 line, 3 effect and 3 background styles the format requires', [['fillStyleLst', 'a:solidFill'], ['lnStyleLst', 'a:ln '], ['effectStyleLst', 'a:effectStyle>'], ['bgFillStyleLst', 'a:solidFill']].every(([tag, item]) => {
    const m = new RegExp('<a:' + tag + '>([\\s\\S]*?)</a:' + tag + '>').exec(byPath['ppt/theme/theme1.xml']);
    return m && (m[1].split('<' + item).length - 1) >= 3;
  }));
  check('the theme names the 12 colours and the Arial font', ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].every(c => byPath['ppt/theme/theme1.xml'].indexOf('<a:' + c + '>') !== -1) && /typeface="Arial"/.test(byPath['ppt/theme/theme1.xml']));
}

console.log('\nTest 3 — the slides carry exactly what the layout says');
{
  const emu = pt => Math.round(pt * 12700);
  slides.forEach((shapes, si) => {
    const xml = byPath['ppt/slides/slide' + (si + 1) + '.xml'];
    const label = 'slide ' + (si + 1) + ': ';
    const sps = xml.split('<p:sp>').slice(1);
    const drawn = shapes.filter(s => s.id !== 'bg');
    check(label + 'one drawn shape per layout shape (the background is the slide background, not a shape)', sps.length === drawn.length, sps.length + ' vs ' + drawn.length);
    const bad = [];
    drawn.forEach((s, i) => {
      const x = sps[i];
      if (x.indexOf('name="' + s.id + '"') === -1) bad.push(s.id + ' name');
      const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(x);
      if (!off || +off[1] !== emu(s.x) || +off[2] !== emu(s.y) || +off[3] !== emu(s.w) || +off[4] !== emu(s.h)) bad.push(s.id + ' geometry');
      if (s.fill && x.indexOf('<a:solidFill><a:srgbClr val="' + s.fill.replace('#', '').toUpperCase() + '"/></a:solidFill><a:ln') === -1) bad.push(s.id + ' fill');
      if (!s.fill && x.indexOf('</a:prstGeom><a:noFill/>') === -1) bad.push(s.id + ' noFill');
      if (s.line && x.indexOf('<a:ln w="9525"><a:solidFill><a:srgbClr val="' + s.line.replace('#', '').toUpperCase() + '"/>') === -1) bad.push(s.id + ' line');
      if (!s.line && x.indexOf('<a:ln><a:noFill/></a:ln>') === -1) bad.push(s.id + ' noLine');
      if (s.type === 'TEXT_BOX' && x.indexOf('<p:cNvSpPr txBox="1"/>') === -1) bad.push(s.id + ' txBox');
      if (s.type !== 'TEXT_BOX' && x.indexOf('txBox') !== -1) bad.push(s.id + ' not a text box');
      if (s.paras.length) {
        const pad = emu(s.pad);
        if (x.indexOf('<a:bodyPr wrap="square" lIns="' + pad + '" tIns="' + pad + '" rIns="' + pad + '" bIns="' + pad + '" rtlCol="0" anchor="' + (s.valign === 'MIDDLE' ? 'ctr' : 't') + '"><a:noAutofit/></a:bodyPr>') === -1) bad.push(s.id + ' bodyPr');
      } else if (x.indexOf('<p:txBody>') !== -1) bad.push(s.id + ' has text but should not');
    });
    check(label + 'name, position, size, fill, outline, margin, anchor, text box flag and no-autofit all match', bad.length === 0, bad.slice(0, 6).join(', '));
    const want = []; drawn.forEach(s => s.paras.forEach(p => p.runs.forEach(r => { if (r.t.length) want.push(r.t); })));
    const got = [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(m => unesc(m[1]));
    check(label + 'all ' + want.length + ' pieces of text are there, in order, unchanged', JSON.stringify(got) === JSON.stringify(want), got.length + ' vs ' + want.length);
    check(label + 'the slide background is set from the layout', xml.indexOf('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="' + shapes.find(s => s.id === 'bg').fill.replace('#', '').toUpperCase() + '"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>') !== -1);
  });
  const agent = byPath['ppt/slides/slide2.xml'], sh = slides[1];
  const find = id => { const i = agent.indexOf('name="' + id + '"'); return agent.slice(i, agent.indexOf('</p:sp>', i)); };
  check('the badge is a pill (corner 50,000), a card has a 10 pt corner, a tile an 8 pt corner, a text box has none',
        find('badge').indexOf('prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/>') !== -1 &&
        new RegExp('fmla="val ' + Math.round(10 / 268 * 100000) + '"').test(find('card-score')) && new RegExp('fmla="val ' + Math.round(8 / 88 * 100000) + '"').test(find('tile-0')) &&
        /prst="rect"><a:avLst\/>/.test(find('name')));
  check('sizes are in hundredths of a point, bold is 1 or 0, colours are 6-digit upper-case hex',
        /sz="3000" b="1"/.test(find('name')) && /sz="1100" b="0"/.test(find('tile-0')) && !/srgbClr val="#/.test(agent) && [...agent.matchAll(/srgbClr val="([^"]+)"/g)].every(m => /^[0-9A-F]{6}$/.test(m[1])));
  check('the font is Arial everywhere', [...agent.matchAll(/<a:latin typeface="([^"]+)"/g)].every(m => m[1] === 'Arial'));
  check('line spacing is a percentage (1.1 -> 110,000)', /<a:lnSpc><a:spcPct val="110000"\/>/.test(find('feedback')) && /<a:lnSpc><a:spcPct val="100000"\/>/.test(find('name')));
  check('alignment: centre for the badge, left for the name', /<a:pPr algn="ctr"/.test(find('badge')) && /<a:pPr algn="l"/.test(find('name')));
  check('the page number is right-aligned', /<a:pPr algn="r"/.test(find('page')));
  check('coaching items are real bullets with a hanging indent of 14 pt in the text colour', /<a:pPr algn="l" marL="177800" indent="-177800">/.test(find('coaching')) && /<a:buClr><a:srgbClr val="FFFFFF"\/><\/a:buClr><a:buFont typeface="Arial"\/><a:buChar char="&#8226;"\/>/.test(find('coaching')));
  check('other paragraphs have no bullet', /<a:buNone\/>/.test(find('name')) && !/buChar/.test(find('feedback')));
  check('no empty text pieces and every paragraph ends with its end-of-paragraph size', !/<a:t><\/a:t>/.test(agent) && agent.split('<a:p>').length === agent.split('<a:endParaRPr').length);
  check('the shape ids on a slide are unique and start at 2', (() => { const ids = [...agent.matchAll(/<p:cNvPr id="(\d+)" name="[^"]+"/g)].map(m => +m[1]); return ids[0] === 2 && new Set(ids).size === ids.length; })());
}

console.log('\nTest 4 — text is escaped, and characters XML cannot hold are dropped');
{
  const bad = String.fromCharCode(1) + String.fromCharCode(0xD800) + '|' + String.fromCharCode(0xDC00) + String.fromCharCode(0xFFFE);   // a lone high surrogate, then a lone low one (not adjacent, so not a pair)
  const smile = String.fromCharCode(0xD83D, 0xDE00);
  const text = 'A & B <tag> "quoted" \'single\' ' + bad + 'ok ' + smile + ' end\tof\nline';
  const shape = { id: 'x', type: 'TEXT_BOX', x: 0, y: 0, w: 100, h: 20, fill: null, line: null, radius: 0, valign: 'TOP', pad: 0,
    paras: [{ runs: [{ t: text, size: 12, bold: false, color: '#2B2B2B' }], align: 'START', bullet: false, spacing: 1 }] };
  const xml = ctx._pxSlide_([{ id: 'bg', type: 'RECTANGLE', x: 0, y: 0, w: 720, h: 405, fill: '#FFFFFF', line: null, radius: 0, paras: [], valign: 'TOP', pad: 0 }, shape]);
  check('the slide is still well-formed XML', xmlError(xml) === '', xmlError(xml));
  const t = unesc(/<a:t>([\s\S]*?)<\/a:t>/.exec(xml)[1]);
  check('special characters survive a round trip', t.indexOf('A & B <tag> "quoted" \'single\' ') === 0, t);
  check('control characters and lone surrogates are dropped', t.indexOf(String.fromCharCode(1)) === -1 && t.indexOf(String.fromCharCode(0xD800)) === -1 && t.indexOf(String.fromCharCode(0xDC00)) === -1 && t.indexOf(String.fromCharCode(0xFFFE)) === -1 && /ok /.test(t));
  check('a real emoji (a surrogate pair) is kept', t.indexOf(smile) !== -1);
  check('tabs and line breaks become spaces', t.indexOf('end of line') !== -1);
  check('the raw characters never appear unescaped in the file', !/<a:t>[^<]*[&<>"][^<]*<\/a:t>/.test(xml.replace(/&(amp|lt|gt|quot|apos|#\d+);/g, '')));
  const injected = ctx._pxShape_(Object.assign({}, shape, { id: 'a"b<c>' }), 2);
  check('a hostile shape name cannot break out of its attribute', xmlError(ctx._pxSlide_([Object.assign({}, shape, { id: 'a"b<c>&' })])) === '' && injected.indexOf('name="a&quot;b&lt;c&gt;"') !== -1);
  const deckHtml = JSON.stringify(parts.map(p => p.xml));
  check('agent names with markup characters are escaped in the real slides', byPath['ppt/slides/slide3.xml'].indexOf('Agent B &lt;Sales&gt; &amp; Co') !== -1 && deckHtml.indexOf('<Sales>') === -1);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
