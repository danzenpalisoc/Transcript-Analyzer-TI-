/**
 * Tests the Issue Resolution override.
 *
 * The AI decides Issue Resolution from the transcript alone and gets it wrong in
 * a way that matters — a call with a technician visit still pending has been
 * marked "Resolved: Yes", which inflates FCR. The analyst can now click the
 * badge to overrule it.
 *
 * Two halves, both covered here:
 *   client — _injectResolutionToggle flips the badge and binds only once
 *   server — extractTextBlock reads the analyst's value back out of the edited
 *            report, which is how the sheet gets corrected on send
 *
 * Usage:  node tests/resolution-toggle-test.js
 * Exits non-zero if any check fails.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
let failures = 0;
function check(name, actual, expected) {
  const ok = actual === expected;
  if (!ok) failures++;
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}`);
  if (!ok) console.log(`          expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
}

// ── A DOM just rich enough for the toggle ───────────────────────────────────
function el(cls, text) {
  const classes = new Set((cls || '').split(' ').filter(Boolean));
  const node = {
    textContent: text || '',
    style: {},
    attrs: {},
    children: [],
    handlers: {},
    classList: {
      contains: c => classes.has(c),
      add: (...cs) => cs.forEach(c => classes.add(c)),
      remove: (...cs) => cs.forEach(c => classes.delete(c)),
    },
    setAttribute(k, v) { this.attrs[k] = v; },
    removeAttribute(k) { delete this.attrs[k]; },
    addEventListener(ev, fn) { this.handlers[ev] = fn; },
    click() { if (this.handlers.click) this.handlers.click({ stopPropagation() {} }); },
    matches(sel) { return sel.split(',').map(s => s.trim().replace(/^\./, '')).some(c => classes.has(c)); },
    querySelectorAll(sel) {
      const out = [];
      (function walk(n) {
        n.children.forEach(c => { if (c.matches(sel)) out.push(c); walk(c); });
      })(this);
      return out;
    },
    querySelector(sel) { return this.querySelectorAll(sel)[0] || null; },
    contains() { return true; },
  };
  return node;
}
function chipWith(labelText, badgeClass, badgeText) {
  const chip = el('ai-chip');
  const label = el('ai-chip-label', labelText);
  const val = el('ai-chip-val');
  const badge = el(badgeClass, badgeText);
  val.children.push(badge);
  chip.children.push(label, val);
  return { chip, badge };
}

// Lift _injectResolutionToggle out of index.html.
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const start = html.indexOf('function _injectResolutionToggle');
if (start === -1) { console.error('_injectResolutionToggle not found'); process.exit(2); }
let depth = 0, end = -1;
for (let i = html.indexOf('{', start); i < html.length; i++) {
  if (html[i] === '{') depth++;
  else if (html[i] === '}') { depth--; if (depth === 0) { end = i + 1; break; } }
}
const ctx = vm.createContext({
  document: { getElementById: () => null },
  _lastHTML: '',
});
vm.runInContext(html.slice(start, end) + '\nthis.__fn = _injectResolutionToggle;', ctx);
const injectToggle = ctx.__fn;

console.log('Issue Resolution override');

console.log('\nClient — the badge flips both ways');
{
  const container = el('root');
  const { chip, badge } = chipWith('Issue Resolution', 'report-badge-no', 'No');
  container.children.push(chip);

  injectToggle(container);
  check('starts as No', badge.textContent, 'No');
  check('is marked clickable', badge.style.cursor, 'pointer');
  check('explains itself on hover', typeof badge.attrs.title, 'string');

  badge.click();
  check('click gives Yes', badge.textContent, 'Yes');
  check('class follows the text', badge.classList.contains('report-badge-yes'), true);
  check('old class removed', badge.classList.contains('report-badge-no'), false);

  badge.click();
  check('clicking again gives No', badge.textContent, 'No');
  check('class follows back', badge.classList.contains('report-badge-no'), true);
}

console.log('\nClient — re-running the injector must not double-bind');
{
  const container = el('root');
  const { chip, badge } = chipWith('Issue Resolution', 'report-badge-yes', 'Yes');
  container.children.push(chip);
  injectToggle(container);
  injectToggle(container);   // the applied view re-injects on every render
  badge.click();
  check('one click is one flip', badge.textContent, 'No');
}

console.log('\nClient — other chips must stay untouched');
{
  const container = el('root');
  const a = chipWith('Issue Resolution', 'report-badge-yes', 'Yes');
  const b = chipWith('Pitched a Sale?', 'report-badge-no', 'No');
  container.children.push(a.chip, b.chip);
  injectToggle(container);
  check('Pitched a Sale is not made clickable', b.badge.style.cursor, undefined);
  b.badge.click();
  check('Pitched a Sale does not flip', b.badge.textContent, 'No');
}

// ── Server half ─────────────────────────────────────────────────────────────
const code = fs.readFileSync(path.join(ROOT, 'Code.gs'), 'utf8');
const s2 = code.indexOf('function extractTextBlock');
let d2 = 0, e2 = -1;
for (let i = code.indexOf('{', s2); i < code.length; i++) {
  if (code[i] === '{') d2++;
  else if (code[i] === '}') { d2--; if (d2 === 0) { e2 = i + 1; break; } }
}
const sctx = vm.createContext({ Logger: { log() {} } });
vm.runInContext(code.slice(s2, e2) + '\nthis.__fn = extractTextBlock;', sctx);
const extractTextBlock = sctx.__fn;

function reportWith(badgeClass, badgeText) {
  return `<div class="report-wrap"><div class="ai-info">
    <div class="ai-chip"><span class="ai-chip-label">Issue Resolution</span>` +
    `<span class="ai-chip-val"><span class="${badgeClass}">${badgeText}</span></span></div>
  </div></div>`;
}

console.log('\nServer — the analyst\'s value is what gets read back');
check('an overridden No reads as No', extractTextBlock(reportWith('report-badge-no', 'No'), 'Issue Resolution'), 'No');
check('an overridden Yes reads as Yes', extractTextBlock(reportWith('report-badge-yes', 'Yes'), 'Issue Resolution'), 'Yes');

console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} CHECK(S) FAILED`);
process.exit(failures === 0 ? 0 : 1);
