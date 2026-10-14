/**
 * Server-side input hardening (security review, batch 2).
 *
 * Every public function in a .gs file can be called by any signed-in user from
 * the browser console, so what a function ACCEPTS matters as much as what the UI
 * sends it. This covers the entry points that took client-supplied text and put
 * it into a page, a cache, or an email:
 *
 *   - the ?ref= value printed into EvalView (reflected XSS) -> isValidAuditRef_
 *   - sendAnalyticsBriefing: was an open mail relay        -> recipient allow-list
 *   - renderAnalyticsHTML: filter values were not escaped  -> stored HTML injection
 *   - saveCachedResult_ / updateCachedResult_ / sendAuditEmail / generateAndServePDF
 *     stored or mailed caller-supplied HTML                -> sanitised on the way in
 *
 * Usage: node tests/server-input-hardening-test.js
 */
const vm = require('vm');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const code = read('Code.gs');

const EVIL = '<img src=x onerror=alert(1)>';

function load(extra) {
  const mails = [];
  const sandbox = Object.assign({
    Logger: { log() {} }, JSON, Array, Number, String, Object, Error, RegExp,
    _emailRateTake_: () => true, isAdminCaller_: () => false,   // sendAnalyticsBriefing counts against the hourly limit
    MailApp: { sendEmail: m => mails.push(m) },
    Utilities: { formatDate: () => 'January 1, 2026' },
    Session: { getScriptTimeZone: () => 'UTC' },
  }, extra);
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('HtmlSanitizer.gs'), ctx);
  const domains = code.match(/^var BRIEFING_EMAIL_DOMAINS[^\n]*/m);
  vm.runInContext((domains ? domains[0] : '') + '\n' +
    ['isValidAuditRef_', 'isAllowedBriefingRecipient_', 'sendAnalyticsBriefing', 'renderAnalyticsHTML', 'esc']
      .map(n => extractFunction(code, n)).join('\n'), ctx);
  ctx.__mails = mails;
  return ctx;
}

console.log('\nTest 1 — only a real audit ref is accepted for the evaluation link');
{
  const ctx = load();
  for (const good of ['NHA-20261012-0001', 'TIA-20261012-0042', 'NHA-20261012-0001-D2'])
    check('accepts ' + good, ctx.isValidAuditRef_(good) === true);
  for (const bad of ['', null, undefined, 42, 'a b', "a'b", 'a"b', '</script><script>alert(1)</script>', 'x<y', 'a;b', 'a\nb',
                     'NHA-1/../x', 'x'.repeat(65)])
    check('rejects ' + JSON.stringify(bad), ctx.isValidAuditRef_(bad) === false);
}

console.log('\nTest 2 — the briefing can only go to ONE address on a TELUS domain');
{
  const ctx = load();
  for (const good of ['a@telus.com', 'First.Last@telusinternational.com', "o'brien@telus.com", 'x+tag@TELUS.COM'])
    check('allows ' + good, ctx.isAllowedBriefingRecipient_(good) === true);
  for (const bad of ['a@evil.com', 'a@telus.com.evil.com', 'a@sub.telus.com', 'a@xtelus.com', 'a@telus.com,b@evil.com',
                     'a@telus.com;b@evil.com', 'a@telus.com b@evil.com', 'a@telus.com\nBcc: x@evil.com', '<a@telus.com>',
                     'Name <a@telus.com>', '', null, undefined, '@telus.com', 'a@', 'a@telus.com ', ' a@telus.com',
                     'a'.repeat(250) + '@telus.com'])
    check('refuses ' + JSON.stringify(bad), ctx.isAllowedBriefingRecipient_(bad) === false);
}

console.log('\nTest 3 — sendAnalyticsBriefing does not mail a refused recipient, and sanitises what it does send');
{
  const ctx = load();
  const r1 = ctx.sendAnalyticsBriefing('victim@evil.com', '<p>x</p>', {}, 'repeats', 5);
  check('refused recipient -> success:false', r1 && r1.success === false, JSON.stringify(r1));
  check('refused recipient -> nothing mailed', ctx.__mails.length === 0, 'mails: ' + ctx.__mails.length);
  check('the error says who is allowed', /telus/i.test((r1 && r1.error) || ''), r1 && r1.error);

  const r2 = ctx.sendAnalyticsBriefing('a@telus.com,b@evil.com', '<p>x</p>', {}, 'repeats', 5);
  check('a list of recipients is refused', r2.success === false && ctx.__mails.length === 0);

  const r3 = ctx.sendAnalyticsBriefing('  analyst@telus.com ', '<p>ok</p>' + EVIL + '<script>alert(2)</script>', {}, 'sales', '12<b>');
  check('an allowed recipient is mailed', r3.success === true && ctx.__mails.length === 1, JSON.stringify(r3));
  const m = ctx.__mails[0] || {};
  check('to is exactly the trimmed address', m.to === 'analyst@telus.com', m.to);
  check('the HTML body carries no handler or script', !/onerror|<script/i.test(m.htmlBody || ''), (m.htmlBody || '').slice(-200));
  check('the legitimate content is still there', /<p>ok<\/p>/.test(m.htmlBody || ''));
  check('rowCount cannot inject into the subject', !/[<>]/.test(m.subject || ''), m.subject);
}

console.log('\nTest 4 — analytics filter values are escaped in the saved/emailed briefing');
{
  const ctx = load();
  const filters = { members: [EVIL], leaders: ['<script>alert(1)</script>'], lobs: ['a" onmouseover="x'], locales: ["'><svg/onload=1>"],
                    vtids: ['<b>'], dateFrom: '<i>', dateTo: EVIL };
  const html = ctx.renderAnalyticsHTML({}, filters, 'repeats', 3);
  check('no raw <img from a member filter', html.indexOf('<img') === -1, html.slice(html.indexOf('Filters:'), html.indexOf('Filters:') + 200));
  check('no raw <script from a leader filter', html.indexOf('<script') === -1);
  check('no raw <svg from a locale filter', html.indexOf('<svg') === -1);
  check('no raw <b> / <i> from vtid or date filter', !/<b>|<i>/.test(html));
  check('the values are still shown, escaped', /&lt;img src=x onerror=alert\(1\)&gt;/.test(html), html.slice(html.indexOf('Filters:'), html.indexOf('Filters:') + 200));
  check('a filter that is not an array does not throw or inject', (function () {
    try { return ctx.renderAnalyticsHTML({}, { members: EVIL, leaders: { length: 1 } }, 'repeats', 1).indexOf('<img') === -1; } catch (e) { return false; }
  })());
  check('missing filters object does not throw', (function () { try { ctx.renderAnalyticsHTML({}, null, 'repeats', 1); return true; } catch (e) { return false; } })());
  check('no filters still reads "All records"', /All records/.test(ctx.renderAnalyticsHTML({}, {}, 'repeats', 1)));
}

console.log('\nTest 5 — the cache and the email/PDF paths sanitise caller-supplied HTML');
{
  // saveCachedResult_ is exercised against real fakes in cache-owner-test.js; here
  // we check, in source, that each entry point sanitises BEFORE it uses the HTML,
  // so a refactor that drops the call is caught.
  const sheetSrc = read('SheetService.gs');
  const save = extractFunction(sheetSrc, 'saveCachedResult_');
  check('saveCachedResult_ sanitises its input', /sanitizeEvalHtml_\(htmlResult\)/.test(save));
  check('...and does so before the first write', save.indexOf('sanitizeEvalHtml_') < save.indexOf('CacheService'));
  const send = extractFunction(code, 'sendAuditEmail');
  check('sendAuditEmail sanitises htmlResult', /htmlResult\s*=\s*sanitizeEvalHtml_\(htmlResult\)/.test(send));
  check('...before it reads or sends anything', send.indexOf('sanitizeEvalHtml_') < send.indexOf('readAuditLog_'));
  const pdf = extractFunction(code, 'generateAndServePDF');
  check('generateAndServePDF sanitises htmlResult and puts the SERVER stylesheet back',
        /htmlResult\s*=\s*sharedCSS\(\)\s*\+\s*sanitizeEvalHtml_\(htmlResult\)/.test(pdf));
  check('...before building the document', pdf.indexOf('sanitizeEvalHtml_') < pdf.indexOf('buildEvalFormHTML'));
  // the model output is produced (and sanitised, and cached) in produceEvaluation_
  const submit = extractFunction(code, 'produceEvaluation_');
  check('submitTranscript sanitises the model output', /html\s*=\s*sanitizeEvalHtml_\(html\)/.test(submit));
  check('...before the cache is written', submit.indexOf('sanitizeEvalHtml_(html)') < submit.indexOf('saveCachedResult_('));
}

console.log('\nTest 6 — doGet refuses a hostile ref and escapes what it prints');
{
  const doGet = extractFunction(code, 'doGet');
  check('doGet validates the ref', /isValidAuditRef_\(ref\)/.test(doGet));
  check('preloaded JSON has every "<" escaped, not only "</script"', /replace\(\/<\/g,\s*'\\\\u003c'\)/.test(doGet));
  const view = read('EvalView.html');
  check('EvalView escapes "<" in the ref it prints', /JSON\.stringify\(auditRef \|\| ''\)\.replace\(\/<\/g/.test(view));
}

console.log('\nTest 7 — the pages still parse: no script block is cut short, and each one is valid JS');
{
  // An HTML parser ends a <script> block at the first closing script tag, even
  // inside a JS comment or string. Read the blocks the way a browser does.
  for (const file of ['index.html', 'EvalView.html']) {
    const html = read(file);
    const blocks = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi)].map(m => m[1]);
    let bad = null;
    blocks.forEach((b, i) => {
      // EvalView is an Apps Script template: scriptlets are replaced by a value for the syntax check
      const js = file === 'EvalView.html' ? b.replace(/<\?!?=[\s\S]*?\?>/g, 'null') : b;
      try { new Function(js); } catch (e) { if (!bad) bad = 'block ' + i + ': ' + e.message; }
    });
    check(file + ': all ' + blocks.length + ' inline script blocks are valid JavaScript', bad === null, bad);
    const holder = blocks.find(b => b.indexOf('function sanitizeEvalHtml_') !== -1);
    check(file + ': the sanitiser is inside ONE intact block (not cut by an early closing tag)',
          !!holder && holder.indexOf("return out.join('')") !== -1);
  }
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
