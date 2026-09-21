/**
 * Tests for extractTextBlock() in Code.gs.
 *
 * extractTextBlock is a pure function, so even though it lives in a .gs file it
 * can be lifted out and exercised in Node. The fixture below is the report shape
 * the evaluation prompt actually asks for, using real values from audit
 * NHA-20260920-0001.
 *
 * Usage:  node tests/extract-test.js [path-to-Code.gs]
 *         (defaults to ../Code.gs)
 * Exits non-zero if any check fails.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const codePath = process.argv[2] || path.join(__dirname, '..', 'Code.gs');
const source = fs.readFileSync(codePath, 'utf8');

// Lift just the function under test, brace-matched from its declaration.
const start = source.indexOf('function extractTextBlock');
if (start === -1) { console.error('extractTextBlock not found in ' + codePath); process.exit(2); }
let depth = 0, end = -1;
for (let i = source.indexOf('{', start); i < source.length; i++) {
  if (source[i] === '{') depth++;
  else if (source[i] === '}') { depth--; if (depth === 0) { end = i + 1; break; } }
}
const fnSrc = source.slice(start, end);

const ctx = vm.createContext({ Logger: { log() {} } });
vm.runInContext(fnSrc + '\nthis.__fn = extractTextBlock;', ctx);
const extractTextBlock = ctx.__fn;

// ── Fixture: the report the prompt specifies ────────────────────────────────
const SUMMARY = 'Karen (daughter of account holder Jorge) was transferred to Roan in the ' +
  'Loyalty/Smart Home Tech department following a failed doorbell camera installation.';

const REPORT = `<div class="report-wrap">
<div class="report-header">
  <div class="report-header-title">&#128204; TELUS NH Analyzer</div>
  <div class="report-header-sub">Repeats &amp; Transfer Audit Report</div>
</div>
<div class="ai-info">
  <div class="ai-chip"><span class="ai-chip-label">Agent Name</span><span class="ai-chip-val">Roan Valdez</span></div>
  <div class="ai-chip"><span class="ai-chip-label">Issue Resolution</span><span class="ai-chip-val"><span class="report-badge-yes">Yes</span></span></div>
  <div class="ai-chip"><span class="ai-chip-label">Repeat Risk</span><span class="ai-chip-val">52%</span></div>
  <div class="ai-chip"><span class="ai-chip-label">Audit Reference</span><span class="ai-chip-val">NHA-20260920-0001</span></div>
</div>
<div class="report-summary-wrap">
  <div class="report-summary-label">&#128222; Call Summary &amp; Key Interaction Details</div>
  <div class="ai-summary" contenteditable="true">${SUMMARY}</div>
  <div class="report-summary-meta">Duration: 51 min 50 sec &nbsp;|&nbsp; Direction: Inbound &nbsp;|&nbsp; Transfer: Yes</div>
</div>
<div class="report-flags">
  <div class="report-flags-head">&#128681; AI Spotted Flags</div>
  <div class="report-flags-body">
    <div class="ai-flag">
      <div class="ai-flag-title">&#9888; ETF disclosed only in the closing recap</div>
      <div class="ai-flag-detail" contenteditable="true">Roan disclosed the $15/month ETF only in the end-of-call summary.</div>
    </div>
  </div>
</div>
</div>`;

let failures = 0;
function check(name, actual, predicate, expectation) {
  const ok = predicate(actual);
  if (!ok) failures++;
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}`);
  if (!ok) console.log(`          ${expectation}\n          got: ${JSON.stringify(actual)}`);
}
const noMarkup = v => !/[<>]|\/div|class=|margin-|font-weight|">/.test(v);

console.log('extractTextBlock — Code.gs');
console.log('file: ' + codePath);

// The defect: "Transfer" matched the TITLE ("Repeats & Transfer Audit Report")
// and sliced raw HTML mid-tag, writing '/div> Karen (daugh...' into the sheet.
console.log('\nTransfer — must read the summary meta line, not the report title');
const transfer = extractTextBlock(REPORT, 'Transfer');
check('is exactly "Yes"', transfer, v => v === 'Yes', 'expected "Yes"');
check('contains no markup', transfer, noMarkup, 'must not contain tag fragments');
check('is not the call summary', transfer, v => v.indexOf('Karen') === -1,
      'must not bleed into the summary text');

// These two already worked (they are chips) — guard against regression.
console.log('\nChip-backed fields must keep working');
check('Repeat Risk is "52%"', extractTextBlock(REPORT, 'Repeat Risk'), v => v === '52%', 'expected "52%"');
check('Issue Resolution is "Yes"', extractTextBlock(REPORT, 'Issue Resolution'), v => v === 'Yes', 'expected "Yes"');

console.log('\nOther meta fields resolve from the same line');
check('Direction is "Inbound"', extractTextBlock(REPORT, 'Direction'), v => v === 'Inbound', 'expected "Inbound"');

// The last-resort branch must never emit markup, whatever it matches.
console.log('\nLast-resort branch must never emit markup');
['Call Reason', 'Opportunities', 'Recommendation', 'Critical Flag'].forEach(function (kw) {
  check(kw + ' is clean or empty', extractTextBlock(REPORT, kw), noMarkup,
        'must be plain text or empty, never a tag fragment');
});

console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} CHECK(S) FAILED`);
process.exit(failures === 0 ? 0 : 1);
