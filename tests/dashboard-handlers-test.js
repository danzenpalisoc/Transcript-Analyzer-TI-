/**
 * A name or reference from the sheet must never become code in the dashboard.
 *
 * Before: the filter pills and the Audit Ref links were built as  onclick="dbPillOnly('member','" + name + "')"  with the name
 * escaped by escHtml/escAttr. The browser decodes an attribute (so &#39; is a quote again) BEFORE the script runs, so an agent
 * named  x');alert(document.domain);//  closed the string and ran code in whoever opened the filter, an admin included. And a
 * name such as O'Brien simply broke its own checkbox with a syntax error.
 *
 * The test does what the browser does: it takes the HTML the page builds, decodes each handler attribute, and RUNS it with the
 * element as `this`, watching what the page functions receive and whether anything else executes.
 *
 * Usage: node tests/dashboard-handlers-test.js
 */
const vm = require('vm');
const { load } = require('./lib/page-harness');
const { makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

const unescapeHtml = s => s.replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
const HOSTILE = "x');alert(document.domain);//";
const NAMES = [HOSTILE, "O'Brien", 'Normal Name', 'Quote "name"', '<img src=x onerror=alert(1)>', 'back\\slash'];

// every inline handler in a piece of HTML: [{ attr, text }]
const handlersOf = html => [...String(html).matchAll(/\s(onclick|onchange|oninput)="([^"]*)"/g)].map(m => ({ attr: m[1], text: m[2] }));
// the attribute with the given name on the element that carries the given handler
function attrsOfElementWith(html, handlerText, attrName) {
  const out = [];
  [...String(html).matchAll(/<(input|span|a)\b[^>]*>/g)].forEach(m => {
    if (m[0].indexOf('"' + handlerText + '"') === -1) return;
    const a = new RegExp('\\s' + attrName + '="([^"]*)"').exec(m[0]);
    out.push(a ? a[1] : null);
  });
  return out;
}
// run a handler the way the browser does: decode the attribute, call it with `this` = the element
function runHandler(text, el) {
  const calls = [], flags = { pwned: false };
  const spy = name => (...a) => { calls.push({ name, args: a }); };
  const sandbox = { dbPillMemberChange: spy('dbPillMemberChange'), dbPillCrossSelect: spy('dbPillCrossSelect'), dbPillOnly: spy('dbPillOnly'), dbMemberChange: spy('dbMemberChange'),
    dbOpenEvaluation: spy('dbOpenEvaluation'), dbApplyAll: spy('dbApplyAll'), dbPillSelectAll: spy('dbPillSelectAll'), dbPillSearch: spy('dbPillSearch'),
    alert: () => { flags.pwned = true; }, document: { domain: 'x' }, event: { stopPropagation() {}, preventDefault() {} } };
  vm.createContext(sandbox);
  let error = '';
  try { vm.runInContext('(function(){ ' + unescapeHtml(text) + ' }).call(__el)', Object.assign(sandbox, { __el: el })); } catch (e) { error = e.message; }
  return { calls, pwned: flags.pwned, error };
}

const h = load();

console.log('\nTest 1 — the filter pill panel: no name ever reaches a handler as code');
{
  h.ctx._dbAllData = NAMES.map(n => ({ 'Team Member': n, 'Audit Ref': 'R-' + n.length }));
  h.ctx._dbAuditLog = []; h.ctx._dbQAUsers = [];
  const panelHtml = h.ctx.dbBuildPanelHTML('member');
  const hs = handlersOf(panelHtml);
  check('the panel has handlers to check (' + hs.length + ')', hs.length >= NAMES.length * 2);
  const texts = new Set(hs.map(x => x.text));
  const staticOnly = hs.every(x => !NAMES.some(n => unescapeHtml(x.text).indexOf(n) !== -1));
  check('no handler contains any of the names, however it is decoded', staticOnly, JSON.stringify([...texts].filter(t => NAMES.some(n => unescapeHtml(t).indexOf(n) !== -1)).slice(0, 2)));
  check('so the handlers are the SAME few fixed texts for every name (' + texts.size + ' different texts for ' + NAMES.length + ' names)', texts.size <= 6, [...texts].join(' || '));

  // run each checkbox handler with each name as the element value
  let allOk = true, detail = '';
  NAMES.forEach(n => {
    const change = hs.find(x => x.attr === 'onchange' && /dbPillMemberChange/.test(x.text));
    const r = runHandler(change.text, { value: n, checked: true });
    if (r.pwned || r.error || r.calls.length !== 1 || r.calls[0].args[0] !== n) { allOk = false; detail = n + ' -> ' + JSON.stringify(r); }
  });
  check('the member checkbox hands the page the EXACT name as a string, runs nothing else (all ' + NAMES.length + ' names, incl. O\'Brien and the hostile one)', allOk, detail);

  // the "only" link: the name is in data-v and read back with getAttribute
  const only = hs.find(x => x.attr === 'onclick' && /dbPillOnly/.test(x.text));
  const dvs = attrsOfElementWith(panelHtml, only.text, 'data-v').map(v => v && unescapeHtml(v));
  check('every "only" link carries its name in data-v, decoded back to the exact name', NAMES.every(n => dvs.indexOf(n) !== -1), JSON.stringify(dvs));
  let onlyOk = true, d2 = '';
  NAMES.forEach(n => {
    const r = runHandler(only.text, { getAttribute: k => (k === 'data-v' ? n : null) });
    if (r.pwned || r.error || r.calls.length !== 1 || r.calls[0].args[0] !== 'member' || r.calls[0].args[1] !== n) { onlyOk = false; d2 = n + ' -> ' + JSON.stringify(r); }
  });
  check('the "only" link passes the exact name and the key, and nothing else runs', onlyOk, d2);
  check('no attribute value in the panel contains a raw double quote or a raw < from a name', !/="[^"]*(<img|<script)[^"]*"/.test(panelHtml));

  const other = h.ctx.dbBuildPanelHTML('leader');
  h.ctx._dbAllData = NAMES.map(n => ({ 'Team Leader': n }));
  const cross = handlersOf(h.ctx.dbBuildPanelHTML('leader')).find(x => /dbPillCrossSelect/.test(x.text));
  const rc = runHandler(cross.text, { value: HOSTILE, checked: true });
  check('the other pills (leader, ops, ...) are safe in the same way', !rc.pwned && !rc.error && rc.calls[0].args[0] === 'leader' && rc.calls[0].args[1] === HOSTILE && rc.calls[0].args[2] === true, JSON.stringify(rc));
}

console.log('\nTest 2 — the member list that feeds the filters');
{
  const items = [];
  h.document.createElement = () => { const it = { className: '', innerHTML: '' }; items.push(it); return it; };
  h.ctx.dbBuildMs('sh-ms-member', NAMES, 'sh', 'member');
  const hs = items.flatMap(it => handlersOf(it.innerHTML));
  check('one checkbox per name, each with a handler', items.length === NAMES.length && hs.length === NAMES.length);
  check('no handler contains a name', hs.every(x => !NAMES.some(n => unescapeHtml(x.text).indexOf(n) !== -1)));
  const r = runHandler(hs[0].text, { value: HOSTILE });
  check('running it hands over the exact hostile name and nothing else', !r.pwned && !r.error && r.calls.length === 1 && r.calls[0].args[0] === 'sh' && r.calls[0].args[1] === HOSTILE, JSON.stringify(r));
}

console.log('\nTest 3 — the Audit Ref link in the Repeats and Sales tables');
{
  h.reset();
  const rows = [{ 'Audit Ref': HOSTILE, 'Analysis Type': 'Repeats & Transfer Analyzer', 'Team Member': 'A', 'Repeat Projection %': '40%', 'Issue Resolved': 'Yes' },
                { 'Audit Ref': "NHA'1", 'Analysis Type': 'Repeats & Transfer Analyzer', 'Team Member': 'B', 'Repeat Projection %': '40%', 'Issue Resolved': 'Yes' }];
  h.ctx.dbRenderTable('r', rows);
  const html = Object.keys(h.els).map(k => h.els[k].innerHTML || '').join('');
  const hs = handlersOf(html).filter(x => /dbOpenEvaluation/.test(x.text));
  check('each row has an Audit Ref link', hs.length === 2, String(hs.length));
  check('the handler is the same fixed text for both references and contains neither', new Set(hs.map(x => x.text)).size === 1 && hs.every(x => x.text.indexOf('NHA') === -1 && x.text.indexOf('alert') === -1), JSON.stringify(hs.map(x => x.text)));
  const refs = attrsOfElementWith(html, hs[0].text, 'data-ref').map(v => v && unescapeHtml(v));
  check('the reference travels in data-ref and decodes back exactly', refs.indexOf(HOSTILE) !== -1 && refs.indexOf("NHA'1") !== -1, JSON.stringify(refs));
  const r = runHandler(hs[0].text, { getAttribute: k => (k === 'data-ref' ? HOSTILE : null) });
  check('clicking it gives dbOpenEvaluation the exact string and runs nothing else', !r.pwned && !r.error && r.calls.length === 1 && r.calls[0].args[0] === HOSTILE, JSON.stringify(r));
}

console.log('\nTest 4 — the report window');
{
  h.reset();
  const writes = [], alerts = [];
  h.ctx.alert = m => alerts.push(m);
  const fakeWin = { document: { write: s => writes.push(s), close() {} } };
  h.ctx.window.open = () => fakeWin;
  const REF = '</title><script>alert(1)</script>';
  h.ctx.dbOpenEvaluation(REF);
  h.pending[0].success({ html: '<div class="report-wrap">report</div>', css: '' });
  const page = writes.join('');
  check('the reference is escaped in the title and the heading (no raw </title><script>)', page.indexOf(REF) === -1 && page.indexOf('&lt;/title&gt;&lt;script&gt;alert(1)&lt;/script&gt;') !== -1, page.slice(0, 160));
  check('the only <script> in the window is the page\'s own Play handler', (page.match(/<script>/g) || []).length === 1 && /speechSynthesis/.test(page.slice(page.indexOf('<script>'))));

  h.reset(); writes.length = 0;
  h.ctx.window.open = () => null;
  h.ctx.dbOpenEvaluation('NHA-1');
  let threw = '';
  try { h.pending[0].success({ html: '<div>report</div>', css: '' }); } catch (e) { threw = e.message; }
  check('a blocked pop-up is explained to the person and does not throw', threw === '' && alerts.some(m => /blocked the pop-up/.test(m)), threw + ' | ' + alerts.join(' | '));
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
