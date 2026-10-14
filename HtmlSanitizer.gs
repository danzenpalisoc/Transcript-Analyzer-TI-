/**
 * HtmlSanitizer.gs
 *
 * One allow-list HTML sanitiser for everything that came from the AI model or
 * from an analyst's edit of it, before it is stored, emailed, turned into a PDF
 * or put on a page.
 *
 * Why a function that REBUILDS the markup instead of stripping bad parts out:
 * the old sanitiser was a handful of regexes that removed `on*="..."` attributes.
 * `<img src=x onerror=alert(1)>` (unquoted) and `<svg/onload=...>` walked
 * straight through it, and the model's output includes whatever a customer or
 * agent said on the call, so a transcript can ask the model to emit them. This
 * function reads the input as a token stream and writes out ONLY tags and
 * attributes on a short list, with every value re-escaped. Anything it does not
 * understand is dropped or turned into plain text, so it fails closed.
 *
 * It is a pure function of a string (no DOM), so the same text runs in Apps
 * Script, in index.html and in EvalView.html, and under node in the tests. A
 * test asserts the three copies are identical -- change all of them together.
 *
 * The name ends in an underscore so Apps Script does not expose it to
 * google.script.run.
 */
function sanitizeEvalHtml_(html) {
  if (html === null || html === undefined) return '';
  html = String(html);

  var ALLOWED = {
    div:1, span:1, p:1, br:1, hr:1, ul:1, ol:1, li:1, b:1, strong:1, i:1, em:1, u:1, s:1,
    small:1, sub:1, sup:1, mark:1, abbr:1, cite:1, q:1, del:1, ins:1, code:1, pre:1,
    blockquote:1, h1:1, h2:1, h3:1, h4:1, h5:1, h6:1, table:1, thead:1, tbody:1, tfoot:1,
    tr:1, th:1, td:1, caption:1, colgroup:1, col:1, a:1, section:1, article:1, header:1,
    footer:1, aside:1, dl:1, dt:1, dd:1, details:1, summary:1, figure:1, figcaption:1,
    // The Play button on sample statements is a <button> holding a tiny <svg>
    // triangle. Allowed because they are inert here: no handlers survive, a button
    // always comes out type="button", and SVG is limited to plain shapes (no
    // script, style, links, images, animation or foreignObject).
    button:1, svg:1, g:1, path:1, polygon:1, polyline:1, circle:1, rect:1, line:1
  };
  var VOID = { br:1, hr:1, col:1 };
  // Elements whose CONTENT is dropped with them (script text, CSS, embedded
  // documents, form controls). Any other unlisted tag is dropped but its text kept.
  var DROP_BODY = {
    script:1, style:1, iframe:1, object:1, embed:1, applet:1, noscript:1, template:1,
    math:1, textarea:1, select:1, option:1, title:1, head:1, xmp:1, plaintext:1,
    noembed:1, noframes:1, form:1, frameset:1, frame:1
  };
  var ATTR_OK = { 'class':1, title:1, colspan:1, rowspan:1, align:1, role:1, scope:1,
                  width:1, height:1, viewbox:1, fill:1, stroke:1, 'stroke-width':1, points:1, d:1,
                  cx:1, cy:1, r:1, x:1, y:1, x1:1, y1:1, x2:1, y2:1 };
  var ENT = { amp:'&', lt:'<', gt:'>', quot:'"', apos:"'", colon:':', tab:'\t',
              newline:'\n', sol:'/', lpar:'(', rpar:')' };

  function fromCode(cp) {
    if (!(cp > 0) || cp > 0x10FFFF) return '';
    if (cp > 0xFFFF) { cp -= 0x10000; return String.fromCharCode(0xD800 + (cp >> 10), 0xDC00 + (cp & 0x3FF)); }
    return String.fromCharCode(cp);
  }
  // Decode the way a browser would before it interprets an attribute value, so
  // `&#106;avascript:` is seen as what it is. Numeric entities decode without a
  // trailing semicolon too, as in browsers.
  function decode(v) {
    return v.replace(/&(?:#(\d+)|#[xX]([0-9a-fA-F]+)|([a-zA-Z][a-zA-Z0-9]*));?/g, function (m, d, x, n) {
      if (d || x) return fromCode(d ? parseInt(d, 10) : parseInt(x, 16));
      var k = n.toLowerCase();
      return Object.prototype.hasOwnProperty.call(ENT, k) ? ENT[k] : m;
    });
  }
  function escAttr(v) {
    return v.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function safeHref(v) {
    var d = decode(v), t = '', c;
    for (var k = 0; k < d.length; k++) {
      c = d.charCodeAt(k);
      // controls, spaces and invisible separators are ignored by URL parsers
      if (c <= 32 || (c >= 127 && c <= 160) || c === 5760 || c === 6158 || (c >= 8192 && c <= 8207) ||
          c === 8232 || c === 8233 || c === 8239 || c === 8287 || c === 12288 || c === 65279) continue;
      t += d.charAt(k);
    }
    return /^(https?:|mailto:|tel:|#)/i.test(t) ? t : null;
  }
  function safeStyle(v) {
    var d = decode(v);
    if (/[\\<>]|\/\*|url\s*\(|expression|@import|javascript:|behavio|binding|position\s*:\s*(fixed|absolute|sticky)/i.test(d)) return null;
    return d;
  }
  function parseAttrs(str, tag) {
    var res = '', seen = {}, hasHref = false, a;
    var re = /([^\s"'<>\/=]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g;
    while ((a = re.exec(str))) {
      var n = a[1].toLowerCase();
      if (seen[n] || n.indexOf('on') === 0) continue;
      var v = a[2] !== undefined ? a[2] : a[3] !== undefined ? a[3] : a[4] !== undefined ? a[4] : '';
      var ok;
      if (n === 'href') {
        if (tag !== 'a') continue;
        ok = safeHref(v);
        if (ok === null) continue;
        hasHref = true;
      } else if (n === 'style') {
        ok = safeStyle(v);
        if (ok === null) continue;
      } else if (ATTR_OK[n] || /^data-[a-z0-9-]+$/.test(n) || /^aria-[a-z-]+$/.test(n)) {
        ok = decode(v);
        if ((n === 'fill' || n === 'stroke') && /url\s*\(|javascript:/i.test(ok)) continue;
      } else {
        continue;
      }
      seen[n] = 1;
      res += ' ' + n + '="' + escAttr(ok) + '"';
    }
    if (hasHref) res += ' rel="noopener noreferrer"';
    if (tag === 'button') res += ' type="button"';   // never a submit button
    return res;
  }

  var out = [], stack = [], last = 0, m;
  // Comments, doctypes and processing instructions are dropped; a tag is `<`,
  // a name, then attributes in which quotes must balance. Anything that merely
  // looks like the start of a tag but never closes (`<img src=x onerror=...`)
  // does not match, so its `<` is escaped below and it becomes inert text.
  var re = /<!--[\s\S]*?-->|<(\/?)([a-zA-Z][a-zA-Z0-9:-]*)((?:[^>"']|"[^"]*"|'[^']*')*)>|<[!?][^>]*>/g;
  function text(t) { out.push(t.replace(/</g, '&lt;')); }

  while ((m = re.exec(html))) {
    text(html.slice(last, m.index));
    last = re.lastIndex;
    if (m[2] === undefined) continue;
    var closing = m[1] === '/', name = m[2].toLowerCase();

    if (!closing && DROP_BODY[name]) {
      var end = new RegExp('</' + name + '\\s*>', 'i').exec(html.slice(last));
      if (!end) { last = html.length; break; }          // never closed: drop the rest
      last += end.index + end[0].length;
      re.lastIndex = last;
      continue;
    }
    if (!ALLOWED[name]) continue;
    if (closing) {
      var at = stack.lastIndexOf(name);
      if (at === -1) continue;                           // closer with no opener
      while (stack.length > at) out.push('</' + stack.pop() + '>');
      continue;
    }
    out.push('<' + name + parseAttrs(m[3], name) + '>');
    if (!VOID[name]) stack.push(name);
  }
  text(html.slice(last));
  while (stack.length) out.push('</' + stack.pop() + '>');   // never leave a tag open
  return out.join('');
}
