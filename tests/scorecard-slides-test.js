/**
 * Scorecard deck (ScorecardSlides.gs): the guards and the whole flow, with fake Google services.
 *
 *  1. What the page may send: only a list of Audit Refs, checked and capped.
 *  2. The data the deck is built from is the dashboard's data (parity with the page's merge).
 *  3. The deck is a PowerPoint package that Drive converts to Google Slides (no Slides API, no extra permission).
 *  4. How it is handed to Drive: zipped parts, the folder, the two conversion styles, errors.
 *  5. Any analyst can build a deck, only for themselves, with limits (rate, size, one at a time).
 *  6. Guards: size, one at a time, identity, failures reported in plain words.
 *  7. Only what is needed is read from the Analysis sheet.
 *  8. The admin tools are admin only and the preview creates nothing.
 *
 * Usage: node tests/scorecard-slides-test.js
 */
const vm = require('vm');
const fs = require('fs');
const path = require('path');
const { load } = require('./lib/page-harness');
const { read, extractFunction, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

function makeEnv(opts) {
  opts = opts || {};
  const logs = [], inserts = [], drive = { folders: [], editors: [], created: 0 };
  const who = { email: 'analyst.one@telus.com' };
  const cacheStore = {};
  const lockState = { busy: false, released: 0 };
  const api = { failFirst: null, failAll: null };
  const sandbox = {
    Logger: { log: m => logs.push(String(m)) }, JSON, Array, Object, String, Number, Math, Date, RegExp, isFinite, isNaN, Error, parseInt,
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Session: { getActiveUser: () => ({ getEmail: () => who.email }) },
    LockService: { getScriptLock: () => ({ tryLock: () => !lockState.busy, waitLock() {}, releaseLock() { lockState.released++; } }) },
    CacheService: { getScriptCache: () => ({ get: k => (k in cacheStore ? cacheStore[k] : null), put: (k, v) => { cacheStore[k] = v; } }) },
    Utilities: {
      base64Encode: s => Buffer.from(String(s)).toString('base64'),
      newBlob: (data, type, name) => ({ data, type, name, setContentType(t) { this.type = t; return this; } }),
      zip: (blobs, name) => ({ zipOf: blobs, name, type: 'application/zip', setContentType(t) { this.type = t; return this; } }),
    },
    Drive: { Files: { insert: (resource, blob, args) => {
      inserts.push({ resource, blob, args });
      if (api.failAll) throw new Error(api.failAll);
      if (api.failFirst && inserts.length === 1) throw new Error(api.failFirst);
      return { id: 'PRES1' };
    } } },
    DriveApp: {
      getFoldersByName: n => ({ hasNext: () => drive.folders.length > 0 && drive.folders[0].name === n, next: () => drive.folders[0] }),
      createFolder: n => { drive.created++; const f = { name: n, getId: () => 'FOLDER1' }; drive.folders.push(f); return f; },
      getFileById: id => ({ id, addEditor: e => drive.editors.push({ id, email: e }) }),
    },
  };
  const ctx = vm.createContext(sandbox);
  ['Config.gs', 'DataRepair.gs', 'ScorecardMetrics.gs', 'ScorecardLayout.gs', 'ScorecardPptx.gs', 'ScorecardSlides.gs'].forEach(f => vm.runInContext(read(f), ctx, { filename: f }));
  vm.runInContext(extractFunction(read('Code.gs'), 'isValidAuditRef_'), ctx);

  const data = { log: [], dash: [] };
  ctx.readAuditLog_ = () => data.log;
  ctx.readDashboardSheet_ = () => data.dash;
  const analysis = new FakeSheet(ctx.ANALYSIS_HEADERS, []);
  const ss = { getSheetByName: n => (n === ctx.ANALYSIS_SHEET && !opts.noAnalysis ? analysis : null) };
  ctx.getOrCreateSpreadsheet = () => ss;
  return { ctx, logs, inserts, drive, who, cacheStore, lockState, api, data, analysis };
}

let counter = 0;
function addAudit(env, agent, o, html) {
  o = o || {};
  const ref = o.ref || 'NHA-2026100' + (++counter);
  env.data.log.push({ 'Audit Ref': ref, 'Submitted At': o.at || 'Thu Oct 0' + (1 + counter % 8) + ' 2026 10:00:00 GMT-0400 (Eastern Daylight Time)', 'SAP ID': o.sap || '1' + agent.length, 'Team Member': agent,
    'Team Leader': 'TL One', 'Line of Business': 'MOB CLS', 'Locale': 'TI Philippines', 'VTID': 'V1', 'Analysis Type': o.type || 'Repeats & Transfer Analyzer',
    'Repeat Projection %': o.repeat || '40%', 'Issue Resolved': o.resolved || 'Yes', 'Transfer Occurred': o.transfer || 'No', 'Email Status': 'Sent' });
  env.data.dash.push({ 'Audit Ref': ref, 'Critical Flags': o.flags || 'No save attempt | Generic offer', 'RCA Sub-Parameter': '', 'Top SMART Opportunity': o.coach || 'Ask one discovery question', 'Sales Attempted': o.sales || '' });
  const row = env.analysis.grid[0].map(h => h === 'Audit Ref' ? ref : (h === 'AI Analysis Result' ? (html || '') : ''));
  env.analysis.grid.push(row);
  return ref;
}
const REPORT = items => '<div class="report-col report-col-working"><div class="report-col-head">What\'s Working</div><ul>' + items.map(i => '<li>' + i + '</li>').join('') + '</ul></div><div class="report-col report-col-needs"><ul><li>needs</li></ul></div>';
const refsOf = env => env.data.log.map(r => r['Audit Ref']);
const textOfDeck = env => JSON.stringify((env.inserts[0] ? env.inserts[0].blob.zipOf : []).map(b => b.data));

console.log('\nTest 1 — the page may send only a list of Audit Refs');
{
  const e = makeEnv();
  const V = x => e.ctx._ssValidateRefs_(x);
  check('not an array is refused', !V('NHA-1').ok && !V(null).ok && !V({}).ok && !V(undefined).ok);
  check('an empty list is refused', !V([]).ok);
  check('a ref with a quote, a space or a slash is refused (nothing is sent on)', !V(['NHA-1"']).ok && !V(['NHA 1']).ok && !V(['../x']).ok && !V(['NHA-1', '<script>']).ok);
  check('a number or an object in the list is refused', !V([5]).ok && !V([{}]).ok && !V([null]).ok);
  const ok = V(['NHA-20261001-0001', ' NHA-20261001-0001-D2 ', 'NHA-20261001-0001']);
  check('good refs are trimmed and de-duplicated', ok.ok && ok.refs.join() === 'NHA-20261001-0001,NHA-20261001-0001-D2', JSON.stringify(ok));
  const many = []; for (let i = 0; i < 2001; i++) many.push('NHA-' + i);
  check('more than 2,000 refs are refused', !V(many).ok && V(many.slice(0, 2000)).ok);
  const r = e.ctx.createScorecardDeck(['bad ref!'], 'x', '');
  check('createScorecardDeck refuses bad input and creates nothing', r.success === false && e.inserts.length === 0 && e.drive.editors.length === 0, JSON.stringify(r));
}

console.log('\nTest 2 — the deck uses the dashboard\'s data (parity with the page merge)');
{
  const e = makeEnv();
  addAudit(e, 'Agent A', { repeat: '0.78', flags: 'F1, F2' });
  addAudit(e, 'Agent B', { type: 'Sales Analyzer', sales: 'Yes' });
  // both sheets hold a value and they differ: Audit_Log wins, exactly as on the dashboard
  e.data.dash[0]['Issue Resolved'] = 'No'; e.data.dash[0]['Team Leader'] = 'Other TL';
  // an audit only in Dashboard_Data, and a Dashboard row whose Team Member is a paragraph
  e.data.dash.push({ 'Audit Ref': 'ONLYDASH', 'Timestamp': 'Mon Oct 05 2026 10:00:00 GMT-0400', 'Team Member': 'Agent C', 'SAP ID': '3', 'Analysis Type': 'Repeats & Transfer Analyzer', 'Critical Flags': 'F9', 'Issue Resolved': 'No' });
  e.data.log.push({ 'Audit Ref': 'LONGDASH', 'Submitted At': 'Mon Oct 05 2026 11:00:00 GMT-0400', 'Team Member': '', 'SAP ID': '4', 'Analysis Type': 'Repeats & Transfer Analyzer' });
  e.data.dash.push({ 'Audit Ref': 'LONGDASH', 'Team Member': 'x'.repeat(120), 'Critical Flags': 'F8' });
  const all = {}; e.data.log.concat(e.data.dash).forEach(r => { all[r['Audit Ref']] = true; });
  const mine = e.ctx._ssMergeRows_(e.data.log, e.data.dash, all);

  const h = load();
  h.ctx.dbFinishRender = function () {};
  h.ctx.dbHandleData({ rows: e.data.dash, ts: new Date().toISOString() }, false);
  h.pending.find(p => p.method === 'getAuditLogData').success(e.data.log);
  const page = h.ctx._dbAllData;
  const keys = ['Audit Ref', 'Timestamp', 'SAP ID', 'Team Member', 'Team Leader', 'Line of Business', 'Locale', 'VTID', 'Analysis Type', 'Repeat Projection %', 'Issue Resolved',
                'Transfer Occurred', 'Email Status', 'Critical Flags', 'RCA Sub-Parameter', 'Top SMART Opportunity', 'Sales Attempted'];
  const diffs = [];
  page.forEach(p => {
    const m = mine.find(x => x['Audit Ref'] === p['Audit Ref']);
    if (!m) { diffs.push('missing ' + p['Audit Ref']); return; }
    keys.forEach(k => { if (String(m[k]) !== String(p[k] === undefined ? '' : p[k])) diffs.push(p['Audit Ref'] + '.' + k + ': ' + m[k] + ' vs ' + p[k]); });
  });
  check('every field a scorecard uses equals the page\'s merged row (' + page.length + ' audits)', diffs.length === 0 && mine.length === page.length, diffs.slice(0, 4).join(' | ') + ' counts ' + mine.length + '/' + page.length);
  check('an audit only in Dashboard_Data is kept', mine.some(r => r['Audit Ref'] === 'ONLYDASH' && r['Team Member'] === 'Agent C'));
  check('a Dashboard_Data paragraph is not used as a name', mine.find(r => r['Audit Ref'] === 'LONGDASH')['Team Member'] === '');
  const sub = e.ctx._ssMergeRows_(e.data.log, e.data.dash, { 'ONLYDASH': true });
  check('only the selected refs are returned', sub.length === 1 && sub[0]['Audit Ref'] === 'ONLYDASH');
}

console.log('\nTest 3 — the deck is a PowerPoint package converted by Drive (no Slides API)');
{
  const src = ['ScorecardSlides.gs', 'ScorecardPptx.gs', 'ScorecardLayout.gs', 'ScorecardMetrics.gs'].map(f => fs.readFileSync(path.join(__dirname, '..', f), 'utf8')).join('\n');
  check('no Google Slides API, no UrlFetchApp and no SlidesApp anywhere in the scorecard code (nothing to switch on in Google Cloud)', !/UrlFetchApp|slides\.googleapis|SlidesApp|presentations\/v1/.test(src));
  const e = makeEnv();
  addAudit(e, 'Agent A', { flags: 'F1 | F2 | F3', coach: 'Ask one open-ended discovery question at the start of every call' }, REPORT(['Warm greeting', 'Clear billing explanation']));
  addAudit(e, 'Agent A', { flags: 'F1', resolved: 'No' }, REPORT(['Warm greeting']));
  const refs = refsOf(e), refSet = {}; refs.forEach(r => { refSet[r] = true; });
  const rows = e.ctx._ssMergeRows_(e.data.log, e.data.dash, refSet);
  const built = e.ctx.buildAgentCards_(rows, { whatsWorking: { [refs[0]]: ['Warm greeting', 'Clear billing explanation'], [refs[1]]: ['Warm greeting'] } });
  const slides = e.ctx._ssBuildSlides_(built, 'Batch 15', 'Team Leader: TL One', rows.length);
  // cover, summary, results, landed, heatmap (1), findings, readiness, the agent, next steps. Even for one agent all of them are built.
  check('a cover, the group slides, one slide per agent and next steps (9 for one agent)', slides.length === 9, 'slides: ' + slides.length);
  const title = i => { const t = slides[i].find(s => s.id === 'title') || slides[i].find(s => s.id === 'name'); return t ? t.paras[0].runs.map(r => r.t).join('') : ''; };
  check('in the order: cover, summary, results, landed, heatmap, findings, readiness, agent, next steps',
        [0, 1, 2, 3, 4, 5, 6, 7, 8].map(title).join(' | ') === 'Audit Scorecards | Executive summary | Scorecard results | Where the agents landed on each KPI | Scorecard heatmap by agent | Audit findings: strengths and gaps | Readiness | Agent A | Next steps',
        [0, 1, 2, 3, 4, 5, 6, 7, 8].map(title).join(' | '));
  check('the page numbers run 1 to 9 in order, "n / 9"', slides.every((sl, i) => JSON.stringify(sl.find(s => s.id === 'page')).indexOf((i + 1) + ' / 9') !== -1));
  const parts = e.ctx.buildPptxParts_(slides);
  const xml = JSON.stringify(parts.map(p => p.xml));
  check('the package has all 9 slides and no 10th', parts.some(p => p.path === 'ppt/slides/slide1.xml') && parts.some(p => p.path === 'ppt/slides/slide9.xml') && !parts.some(p => p.path === 'ppt/slides/slide10.xml'));
  check('the text of the slides is in the package (name, flags, strengths, label, filters)', xml.indexOf('Agent A') !== -1 && xml.indexOf('Warm greeting') !== -1 && xml.indexOf('Batch 15') !== -1 && xml.indexOf('Team Leader: TL One') !== -1);
  const blob = e.ctx._ssPptxBlob_(parts);
  check('the parts are zipped under their own paths, [Content_Types].xml first, and the file is typed as PowerPoint', blob.zipOf.length === parts.length && blob.zipOf[0].name === '[Content_Types].xml' &&
        blob.zipOf.every((b, i) => b.name === parts[i].path && b.data === parts[i].xml) && blob.type === 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
}

console.log('\nTest 4 — how the deck is handed to Drive');
{
  const shapes = [{ id: 'bg', type: 'RECTANGLE', x: 0, y: 0, w: 720, h: 405, fill: '#FFFFFF', line: null, radius: 0, paras: [], valign: 'TOP', pad: 0 }];
  const e = makeEnv();
  const deck = e.ctx._ssCreateDeck_('My title', [shapes, shapes]);
  check('one Drive call creates the deck as a Google Slides file, with the title', e.inserts.length === 1 && e.inserts[0].resource.title === 'My title' && e.inserts[0].resource.mimeType === 'application/vnd.google-apps.presentation');
  check('it is created straight into the "NH Analyzer Scorecards" folder (created when missing)', e.inserts[0].resource.parents[0].id === 'FOLDER1' && e.drive.folders[0].name === 'NH Analyzer Scorecards' && e.drive.created === 1);
  check('the content sent is the zipped PowerPoint package', e.inserts[0].blob.type === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' && e.inserts[0].blob.zipOf[0].name === '[Content_Types].xml');
  check('the link opens the new presentation', deck.url === 'https://docs.google.com/presentation/d/PRES1/edit' && deck.id === 'PRES1');
  e.ctx._ssCreateDeck_('Again', [shapes]);
  check('the folder is reused, not created again', e.drive.created === 1);

  const f = makeEnv();
  f.api.failFirst = 'Invalid conversion request';
  const second = f.ctx._ssCreateDeck_('T', [shapes]);
  check('if Drive refuses the first conversion style, the other one is tried (convert: true, PowerPoint type)', f.inserts.length === 2 && f.inserts[1].args && f.inserts[1].args.convert === true &&
        f.inserts[1].resource.mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' && second.id === 'PRES1');
  check('and the refusal is logged', f.logs.some(l => /first conversion style refused/.test(l)));
  const g = makeEnv();
  g.api.failAll = 'Drive said no';
  let msg = '';
  try { g.ctx._ssCreateDeck_('T', [shapes]); } catch (x) { msg = x.message; }
  check('if both styles fail the error is passed on', /Drive said no/.test(msg) && g.inserts.length === 2, msg);
  const h2 = makeEnv();
  h2.ctx.Drive.Files.insert = () => ({});
  let msg2 = '';
  try { h2.ctx._ssCreateDeck_('T', [shapes]); } catch (x) { msg2 = x.message; }
  check('a reply with no file id is an error, not a broken link', /did not return/.test(msg2), msg2);
}

console.log('\nTest 5 — any analyst can build a deck; limits and sharing');
{
  const e = makeEnv();
  addAudit(e, 'Agent A', {}, REPORT(['Warm greeting']));
  addAudit(e, 'Agent A', {}, REPORT(['Warm greeting']));
  addAudit(e, 'Agent B', { type: 'Sales Analyzer', sales: 'Yes' }, REPORT(['Clear plan explanation']));
  const refs = refsOf(e);
  const r = e.ctx.createScorecardDeck(refs, '  MOB CLS\nBatch 15 ', 'Team Leader: TL One');
  check('an ordinary analyst (not an admin) gets a deck', r.success === true && r.agents === 2 && r.audits === 3 && r.slides === 10 && /docs\.google\.com\/presentation\/d\/PRES1/.test(r.url), JSON.stringify(r));
  check('the deck goes into the "NH Analyzer Scorecards" folder', e.inserts.length === 1 && e.inserts[0].resource.parents[0].id === 'FOLDER1');
  check('it is shared, as editor, with the person who asked and nobody else', e.drive.editors.length === 1 && e.drive.editors[0].email === 'analyst.one@telus.com' && e.drive.editors[0].id === 'PRES1');
  check('the title carries the label (cleaned of line breaks)', e.inserts[0].resource.title.indexOf('MOB CLS Batch 15') !== -1, e.inserts[0].resource.title);
  check('the lock is released', e.lockState.released === 1);
  const text = textOfDeck(e);
  check('the numbers and the "What\'s Working" text from the report are on the slides', text.indexOf('Warm greeting') !== -1 && text.indexOf('Clear plan explanation') !== -1 && text.indexOf('Agent A') !== -1);
  check('the label and the filters reach the cover slide', text.indexOf('MOB CLS Batch 15') !== -1 && text.indexOf('Team Leader: TL One') !== -1);

  // rate limit: 5 per person per hour (1 built so far)
  e.ctx.createScorecardDeck(refs, '', ''); e.ctx.createScorecardDeck(refs, '', ''); e.ctx.createScorecardDeck(refs, '', ''); e.ctx.createScorecardDeck(refs, '', '');
  const before = e.inserts.length;
  const sixth = e.ctx.createScorecardDeck(refs, '', '');
  check('the sixth deck in an hour is refused with a clear message and builds nothing', sixth.success === false && /5 decks/.test(sixth.error) && e.inserts.length === before, JSON.stringify(sixth));
  e.who.email = 'someone.else@telus.com';
  check('the limit is per person', e.ctx.createScorecardDeck(refs, '', '').success === true);
}

console.log('\nTest 6 — guards: size, one at a time, identity, failures');
{
  const e = makeEnv();
  for (let i = 0; i < 41; i++) addAudit(e, 'Agent Number ' + i, {});
  const refs = refsOf(e);
  const big = e.ctx.createScorecardDeck(refs, '', '');
  check('more than 40 agents is refused with advice to narrow the filters, and nothing is built', big.success === false && /41 agents/.test(big.error) && /Narrow the filters/.test(big.error) && e.inserts.length === 0, JSON.stringify(big));
  check('...and it does not count against the hourly deck limit (only against the separate try counter)', Object.keys(e.cacheStore).filter(k => k.indexOf('sc_rl_') === 0).length === 0 && Object.keys(e.cacheStore).filter(k => k.indexOf('sc_try_') === 0).length === 1);
  check('the lock is released after a refusal', e.lockState.released === 1);
  check('exactly 40 agents is allowed', e.ctx.createScorecardDeck(refs.slice(0, 40), '', '').success === true);

  const b = makeEnv(); addAudit(b, 'Agent A', {});
  b.lockState.busy = true;
  const busy = b.ctx.createScorecardDeck(refsOf(b), '', '');
  check('while another deck is being built the answer is "try again in a minute"', busy.success === false && /another scorecard deck/i.test(busy.error) && b.inserts.length === 0);

  const n = makeEnv(); addAudit(n, 'Agent A', {}); n.who.email = '';
  const noId = n.ctx.createScorecardDeck(refsOf(n), '', '');
  check('if the caller cannot be identified nothing is built (it could not be shared)', noId.success === false && /who you are/.test(noId.error) && n.inserts.length === 0);

  const none = makeEnv(); addAudit(none, '', {});
  check('audits with no usable Team Member give a clear message', none.ctx.createScorecardDeck(refsOf(none), '', '').success === false && none.inserts.length === 0);

  const f = makeEnv(); addAudit(f, 'Agent A', {});
  f.api.failAll = 'Service invoked too many times for one day: drive. secret-internal-detail';
  const fail = f.ctx.createScorecardDeck(refsOf(f), '', '');
  check('a Drive failure is reported in plain words, without the raw error, to an analyst', fail.success === false && /Google is busy/.test(fail.error) && fail.error.indexOf('secret-internal-detail') === -1, fail.error);
  check('the lock is released after a failure, and nothing was shared', f.lockState.released === 1 && f.drive.editors.length === 0);
  f.who.email = 'danzen.palisoc@telus.com';
  f.cacheStore = {};
  const adminFail = f.ctx.createScorecardDeck(refsOf(f), '', '');
  check('an admin also sees the details', /Details for admins/.test(adminFail.error) && /secret-internal-detail/.test(adminFail.error), adminFail.error);
  check('the real error is logged', f.logs.some(l => /scorecard error/.test(l)));
  const p = makeEnv(); addAudit(p, 'Agent A', {}); p.api.failAll = 'Access denied: forbidden';
  check('a permission problem says so', /not allowed to create the deck/.test(p.ctx.createScorecardDeck(refsOf(p), '', '').error));
  const u = makeEnv(); addAudit(u, 'Agent A', {}); u.api.failAll = 'something odd';
  check('anything else gets a calm, general message', /could not be built/.test(u.ctx.createScorecardDeck(refsOf(u), '', '').error));
}

console.log('\nTest 7 — only what is needed is read from the Analysis sheet');
{
  const e = makeEnv();
  const stamps = ['Mon Oct 12 2026 10:00:00 GMT-0400', 'Sun Oct 11 2026 10:00:00 GMT-0400', 'Sat Oct 10 2026 10:00:00 GMT-0400', 'Fri Oct 09 2026 10:00:00 GMT-0400', 'Thu Oct 08 2026 10:00:00 GMT-0400', 'Wed Oct 07 2026 10:00:00 GMT-0400'];
  const aRefs = stamps.map((s, i) => addAudit(e, 'Agent A', { at: s }, REPORT(['Strength number ' + i])));
  addAudit(e, 'Agent B', { at: 'Mon Oct 12 2026 12:00:00 GMT-0400' }, REPORT(['B strength']));
  e.analysis.resetReads();
  const prep = e.ctx._ssPrepare_(refsOf(e));
  const wanted = e.ctx._ssNewestRefsPerAgent_(prep.rows, 3);
  check('the 3 newest audits of each agent are chosen (3 + 1)', wanted.length === 4 && aRefs.slice(0, 3).every(r => wanted.indexOf(r) !== -1) && aRefs.slice(3).every(r => wanted.indexOf(r) === -1), wanted.join());
  const ww = e.ctx._ssLoadWhatsWorking_(wanted);
  const resCol = e.analysis.grid[0].indexOf('AI Analysis Result') + 1;
  check('only those reports are read: 4 cells of the report column, never the whole column', e.analysis.cellsReadInColumn(resCol, 2) === 4, String(e.analysis.cellsReadInColumn(resCol, 2)));
  check('the items come back per audit', ww[aRefs[0]][0] === 'Strength number 0' && ww[wanted.find(r => ww[r] && ww[r][0] === 'B strength')][0] === 'B strength');
  const budget = e.ctx._ssLoadWhatsWorking_(wanted, -1);
  check('a time budget that has run out stops the reading instead of running on', Object.keys(budget).length === 0);

  const n = makeEnv({ noAnalysis: true });
  addAudit(n, 'Agent A', {});
  check('with no Analysis sheet the deck is still built (Strengths fall back to the data)', n.ctx.createScorecardDeck(refsOf(n), '', '').success === true);
  const broken = makeEnv(); addAudit(broken, 'Agent A', {}, '#ERROR!');
  check('a damaged report cell does not break the deck', broken.ctx.createScorecardDeck(refsOf(broken), '', '').success === true);
}

console.log('\nTest 8 — the admin tools');
{
  const e = makeEnv(); addAudit(e, 'Agent A', {}); addAudit(e, 'Agent B', { type: 'Sales Analyzer', sales: 'No' });
  let a = '', b = '';
  try { e.ctx.previewScorecards(refsOf(e)); } catch (x) { a = x.message; }
  try { e.ctx.testSlidesAccess(); } catch (x) { b = x.message; }
  check('previewScorecards refuses an analyst', /admin/i.test(a), a);
  check('testSlidesAccess refuses an analyst', /admin/i.test(b), b);
  check('nothing was created for the analyst', e.inserts.length === 0 && e.drive.created === 0);
  e.who.email = 'danzen.palisoc@telus.com';
  const p = e.ctx.previewScorecards(refsOf(e));
  check('the preview lists each agent with the numbers', p.success && p.agents === 2 && e.logs.some(l => /Agent A/.test(l) && /Avg Repeat Risk/.test(l)), e.logs.join(' | '));
  check('the preview creates no file and shares nothing', e.inserts.length === 0 && e.drive.created === 0 && e.drive.editors.length === 0);
  check('a bad list gets an error, not a crash', e.ctx.previewScorecards(['bad ref']).success === false);
  const t = e.ctx.testSlidesAccess();
  check('testSlidesAccess builds a small deck, files it and shares it with the admin', t.success && /PRES1/.test(t.url) && e.drive.editors[0].email === 'danzen.palisoc@telus.com' && e.inserts.length === 1, JSON.stringify(t));
  check('the test deck has all the slides of a one-agent deck (9)', e.inserts[0].blob.zipOf.filter(z => /^ppt\/slides\/slide\d+\.xml$/.test(z.name)).length === 9);
}

console.log('\nTest — every try counts, before the lock and the reads');
{
  const e = makeEnv({});
  const refs = ['NHA-20261013-9999'];                          // not in the sheets: every try ends empty and used to cost no quota
  let refused = 0, last = null;
  for (let i = 0; i < 30; i++) { last = e.ctx.createScorecardDeck(refs, '', ''); if (/Too many scorecard requests/.test(last.error || '')) refused++; }
  check('the first 30 tries are answered normally', refused === 0, JSON.stringify(last));
  const before = e.lockState.released;
  const r31 = e.ctx.createScorecardDeck(refs, '', '');
  check('the 31st try in the hour is refused', r31.success === false && /Too many scorecard requests/.test(r31.error), JSON.stringify(r31));
  check('and the lock was not taken for it', e.lockState.released === before);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
