/**
 * Smoke test of Generate Slides, end to end, on the REAL code (createScorecardDeck -> the layout -> the PowerPoint package -> the
 * zipped file handed to Drive). Only Google is fake: the sheets, Drive, the session, the cache and the lock.
 *
 * It looks at the PowerPoint file itself, the thing Drive has to convert to Google Slides:
 *   - every part is well-formed XML, the content types, relationships and slide list agree;
 *   - the file has the right number of slides, in the right order, and every agent is on the heatmap and on a slide of their own;
 *   - names and texts with & < > ' " are escaped, so one odd character cannot make Drive refuse the whole deck;
 *   - the biggest deck (40 agents) is built quickly and the file stays small.
 * What it cannot see is how Google Drive converts the file (no Google here): that needs a real deck.
 *
 * Usage: node tests/scorecard-deck-smoke-test.js
 */
const vm = require('vm');
const { read, extractFunction, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

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

function makeEnv() {
  const inserts = [], drive = { folders: [], editors: [] };
  const cacheStore = {};
  const sandbox = {
    Logger: { log() {} }, JSON, Array, Object, String, Number, Math, Date, RegExp, isFinite, isNaN, Error, parseInt,
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'analyst.one@telus.com' }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock() {}, releaseLock() {} }) },
    CacheService: { getScriptCache: () => ({ get: k => (k in cacheStore ? cacheStore[k] : null), put: (k, v) => { cacheStore[k] = v; } }) },
    Utilities: {
      base64Encode: s => Buffer.from(String(s)).toString('base64'),
      newBlob: (data, type, name) => ({ data, type, name, setContentType(t) { this.type = t; return this; } }),
      zip: (blobs, name) => ({ zipOf: blobs, name, type: 'application/zip', setContentType(t) { this.type = t; return this; } }),
    },
    Drive: { Files: { insert: (resource, blob, args) => { inserts.push({ resource, blob, args }); return { id: 'PRES1' }; } } },
    DriveApp: {
      getFoldersByName: nm => ({ hasNext: () => drive.folders.length > 0 && drive.folders[0].name === nm, next: () => drive.folders[0] }),
      createFolder: nm => { const f = { name: nm, getId: () => 'FOLDER1' }; drive.folders.push(f); return f; },
      getFileById: id => ({ id, addEditor: e => drive.editors.push({ id, email: e }) }),
    },
  };
  const ctx = vm.createContext(sandbox);
  ['Config.gs', 'DataRepair.gs', 'ScorecardMetrics.gs', 'ScorecardTargets.gs', 'ScorecardLayout.gs', 'ScorecardPptx.gs', 'ScorecardSlides.gs'].forEach(f => vm.runInContext(read(f), ctx, { filename: f }));
  vm.runInContext(extractFunction(read('Code.gs'), 'isValidAuditRef_'), ctx);
  const data = { log: [], dash: [] };
  ctx.readAuditLog_ = () => data.log;
  ctx.readDashboardSheet_ = () => data.dash;
  const analysis = new FakeSheet(ctx.ANALYSIS_HEADERS, []);
  const ss = { getSheetByName: nm => (nm === ctx.ANALYSIS_SHEET ? analysis : null) };
  ctx.getOrCreateSpreadsheet = () => ss;
  return { ctx, inserts, drive, data, analysis };
}

let counter = 0;
const REPORT = items => '<div class="report-col report-col-working"><div class="report-col-head">What\'s Working</div><ul>' + items.map(i => '<li>' + i + '</li>').join('') + '</ul></div><div class="report-col report-col-needs"><ul><li>needs</li></ul></div>';
function addAudit(env, agent, o) {
  const ref = 'NHA-2026101' + String(++counter).padStart(4, '0');
  env.data.log.push({ 'Audit Ref': ref, 'Submitted At': 'Thu Oct 08 2026 10:00:00 GMT-0400 (Eastern Daylight Time)', 'SAP ID': o.sap, 'Team Member': agent, 'Team Leader': 'TL One', 'Line of Business': o.lob || 'MOB CLS',
    'Locale': 'TI Philippines', 'VTID': 'V1', 'Analysis Type': o.type || 'Repeats & Transfer Analyzer', 'Repeat Projection %': o.repeat, 'Issue Resolved': o.resolved || 'Yes',
    'Transfer Occurred': o.transfer || 'No', 'Email Status': 'Sent' });
  env.data.dash.push({ 'Audit Ref': ref, 'Critical Flags': o.flags || '', 'RCA Sub-Parameter': '', 'Top SMART Opportunity': o.coach || 'Ask one discovery question', 'Sales Attempted': o.sales || '' });
  env.analysis.grid.push(env.analysis.grid[0].map(h => h === 'Audit Ref' ? ref : (h === 'AI Analysis Result' ? (o.report || '') : '')));
  return ref;
}
const refsOf = env => env.data.log.map(r => r['Audit Ref']);
const LONG = 'Customer was transferred without a warm handoff and the agent did not set any expectation for the next step on the call, then closed the call without a recap';

function fill(env, count, lobs) {
  for (let i = 0; i < count; i++) {
    const name = i === 0 ? 'Zoë O\'Brien & <Sons> "Jr"' : (i % 4 === 0 ? 'Maria Guadalupe Hernandez-Villanueva de la Cruz ' + i : 'Agent ' + String.fromCharCode(65 + (i % 26)) + i);
    for (let a = 0; a < 3; a++) addAudit(env, name, { lob: lobs ? lobs[i % lobs.length] : undefined, sap: String(3000 + i), repeat: [10, 30, 55, 80][(i + a) % 4] + '%', resolved: (i + a) % 5 === 0 ? 'No' : 'Yes',
      flags: (i + a) % 3 === 0 ? 'Dead air | ' + LONG : '', transfer: i % 7 === 0 ? 'Yes' : 'No', type: i % 9 === 8 ? 'Sales Analyzer' : undefined, sales: i % 9 === 8 ? (a ? 'Yes' : 'No') : '',
      coach: LONG, report: REPORT(['Probing & investigation <skills> are "solid"', LONG]) });
  }
}
const namesOf = env => { const s = {}; env.data.log.forEach(r => { s[r['Team Member']] = true; }); return Object.keys(s).sort((a, b) => a.toLowerCase() < b.toLowerCase() ? -1 : 1); };
const partsOf = env => env.inserts[0].blob.zipOf;
const slideXml = (parts, i) => (parts.find(p => p.name === 'ppt/slides/slide' + i + '.xml') || {}).data || '';
const slideText = (parts, i) => [...slideXml(parts, i).matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(m => unesc(m[1])).join(' ');

[[1, 9], [14, 4 + 1 + 2 + 14 + 1], [40, 4 + 3 + 2 + 40 + 1]].forEach(([agents, slides]) => {
  console.log('\nSmoke ' + agents + ' agent' + (agents === 1 ? '' : 's') + ' — createScorecardDeck, the file Drive receives');
  const e = makeEnv(); fill(e, agents);
  const t0 = Date.now();
  const r = e.ctx.createScorecardDeck(refsOf(e), 'MOB CLS Batch 15', 'Team Leader: TL One');
  const ms = Date.now() - t0;
  check('the deck is created and reported', r.success === true && r.agents === agents && r.slides === slides && /docs\.google\.com\/presentation\/d\/PRES1/.test(r.url), JSON.stringify(r).slice(0, 200));
  check('one Drive call, into the Scorecards folder, shared with the person who asked', e.inserts.length === 1 && e.inserts[0].resource.parents[0].id === 'FOLDER1' && e.drive.editors.length === 1 && e.drive.editors[0].email === 'analyst.one@telus.com');
  const parts = partsOf(e);
  const bad = parts.map(p => [p.name, xmlError(p.data)]).filter(x => x[1]);
  check('every one of the ' + parts.length + ' parts is well-formed XML', bad.length === 0, JSON.stringify(bad.slice(0, 3)));
  const slideParts = parts.filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p.name));
  check('the file has exactly ' + slides + ' slides', slideParts.length === slides, 'slides in file: ' + slideParts.length);
  const pres = parts.find(p => p.name === 'ppt/presentation.xml').data, ids = [...pres.matchAll(/<p:sldId id="(\d+)" r:id="(rId\d+)"/g)];
  check('the slide list names every slide once, with unique ids', ids.length === slides && new Set(ids.map(m => m[1])).size === slides);
  const ct = parts.find(p => p.name === '[Content_Types].xml').data;
  check('every slide has a content type', slideParts.every(p => ct.indexOf('PartName="/' + p.name + '"') !== -1));
  const first = slideText(parts, 1), last = slideText(parts, slides);
  check('the first slide is the cover and the last is Next steps', /Audit Scorecards/.test(first) && /MOB CLS Batch 15/.test(first) && /Next steps/.test(last), first.slice(0, 60) + ' ... ' + last.slice(0, 40));
  const order = [2, 3, 4].map(i => slideText(parts, i).slice(0, 40));
  check('then the summary, the results and where the agents landed', /^Executive summary/.test(order[0]) && /^Scorecard results/.test(order[1]) && /^Where the agents landed/.test(order[2]), order.join(' | '));

  const names = namesOf(e), heatPages = Math.ceil(agents / 14);
  let onHeat = '';
  for (let p = 0; p < heatPages; p++) onHeat += ' ' + slideText(parts, 5 + p);
  check('every agent is on the heatmap (as far as the name fits, a long name is cut with an ellipsis)', names.every(n => onHeat.indexOf(n.split(' ')[0]) !== -1));
  const agentSlides = []; for (let i = 6 + heatPages + 1; i < slides; i++) agentSlides.push(slideText(parts, i));
  // a name too long for the title is cut at a word end with an ellipsis, so compare the beginning of the name
  check('every agent has a slide of their own, in alphabetical order', agentSlides.length === agents && agentSlides.every((t, i) => t.indexOf(names[i].slice(0, 20)) === 0),
        agentSlides.length + ' agent slides; first mismatch: ' + agentSlides.map((t, i) => [i, t.slice(0, 30), names[i].slice(0, 30)]).filter(x => x[1].indexOf(x[2].slice(0, 20)) !== 0).slice(0, 1).join(' vs '));
  check('the page numbers are "n / ' + slides + '" on every slide', slideParts.every((p, i) => slideText(parts, i + 1).indexOf((i + 1) + ' / ' + slides) !== -1));
  const all = parts.map(p => p.data).join('');
  check('the odd characters of a name (& < > \' ") reach the file escaped and read back unchanged', agents === 0 || slideText(parts, 6 + heatPages + 1).indexOf('Zoë O\'Brien & <Sons> "Jr"') !== -1 || agentSlides.some(t => t.indexOf('Zoë O\'Brien & <Sons> "Jr"') !== -1));
  check('no raw "<Sons>" and no unescaped ampersand is in the XML text', !/<a:t>[^<]*<Sons>/.test(all) && !/&(?!(amp|lt|gt|quot|apos|#\d+);)/.test(all));
  const size = all.length;
  check('the package is small (' + Math.round(size / 1024) + ' KB of XML, far below the Drive limit)', size < 3 * 1024 * 1024, size + ' chars');
  check('it was built in under 3 seconds here (' + ms + ' ms)', ms < 3000, ms + ' ms');
});

console.log('\nSmoke mixed LOBs — 20 agents in 5 LOBs, a section for each');
{
  const e = makeEnv(); fill(e, 20, ['PF CxSS', 'MOB CLS', 'one cls', 'Koodo', 'Brand New LOB']);
  const r = e.ctx.createScorecardDeck(refsOf(e), 'Mixed', 'x');
  const want = 1 + 5 * (6 + 1 + 4);
  check('the deck is created with the cover and five sections of 4 agents (' + want + ' slides)', r.success === true && r.agents === 20 && r.slides === want, JSON.stringify(r).slice(0, 200));
  const parts = partsOf(e), bad = parts.map(p => [p.name, xmlError(p.data)]).filter(x => x[1]);
  check('every part of the file is well-formed XML', bad.length === 0, JSON.stringify(bad.slice(0, 3)));
  const count = parts.filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p.name)).length;
  check('the file has ' + want + ' slides', count === want, 'slides in file: ' + count);
  check('the cover lists the five LOBs', /LOBs:\s+MOB CLS \(4\), MOB CxSS \(4\), PF CLS \(4\), PF CxSS \(4\), Other LOB \(4\)/.test(slideText(parts, 1)), slideText(parts, 1).slice(-160));
  check('the first section is MOB CLS and its slides name the quarter of the targets', /^Executive summary/.test(slideText(parts, 2)) && /MOB CLS/.test(slideText(parts, 2)) && /Q\d 2026 targets/.test(slideText(parts, 2)), slideText(parts, 2).slice(-90));
  const lastSection = 1 + 4 * 11 + 1;
  check('the last section is Other LOB, with no targets named', /Other LOB/.test(slideText(parts, lastSection)) && !/targets/.test(slideText(parts, lastSection)), slideText(parts, lastSection).slice(-70));
  check('the page numbers are "n / ' + want + '"', parts.filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p.name)).every((p, i) => slideText(parts, i + 1).indexOf((i + 1) + ' / ' + want) !== -1));
}

console.log('\nSmoke 41 agents — refused, nothing is created');
{
  const e = makeEnv(); fill(e, 41);
  const r = e.ctx.createScorecardDeck(refsOf(e), 'x', '');
  check('more than 40 agents is refused with a plain message and nothing is sent to Drive', r.success === false && /41 agents/.test(r.error || '') && e.inserts.length === 0 && e.drive.editors.length === 0, JSON.stringify(r).slice(0, 160));
}

console.log('\nSmoke — an admin\'s sample deck (testSlidesAccess) is the full deck for one agent');
{
  const e = makeEnv();
  e.ctx.requireAdmin_ = () => {};
  const t = e.ctx.testSlidesAccess();
  const parts = e.inserts.length ? partsOf(e) : [];
  check('the sample deck is made, with all 9 slides, and every part is well-formed', t.success === true && parts.filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p.name)).length === 9 && parts.every(p => !xmlError(p.data)), JSON.stringify(t).slice(0, 120));
}

console.log('\n' + (check.state.failures === 0 ? 'ALL PASSED' : check.state.failures + ' FAILED'));
process.exit(check.state.failures === 0 ? 0 : 1);
