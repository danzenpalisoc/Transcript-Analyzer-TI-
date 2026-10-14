/**
 * The three report columns are editable, and the Play button speaks what the analyst
 * wrote.
 *
 * What broke: the model's report carries contenteditable="true" on the bullets and the
 * sample statements, but the HTML sanitiser removes that attribute (on purpose: the
 * evaluation page that agents open must stay read-only), and the page only put it back
 * AFTER "Apply Suggestions & Close". So in the modal the columns were read-only. And a
 * Play button spoke its data-tts, which is fixed when the button is made, so an edited
 * statement still played the model's original words.
 *
 * The real browser behaviour (typing, Select-All, the Play button coming back) was
 * tried in Chrome against the real helper code and CSS; node has no DOM, so this test
 * covers the pure helper and, as source checks, the wiring that must not be lost.
 *
 * Usage: node tests/report-editing-test.js
 */
const vm = require('vm');
const { read, extractFunction, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const html = read('index.html');
const fn = name => extractFunction(html, name);

console.log('\nTest 1 — _cleanStatement: what the voice is given');
{
  const ctx = vm.createContext({ String, RegExp, Object });
  vm.runInContext(fn('_cleanStatement'), ctx);
  const c = ctx._cleanStatement;
  check('the model\'s wrapping quotes are removed', c('"Hello there."') === 'Hello there.', c('"Hello there."'));
  check('a quote left in the middle (typed after the closing quote) is removed too', c('"Hello." and more') === 'Hello. and more', c('"Hello." and more'));
  check('curly quotes are removed', c(String.fromCharCode(8220) + 'Hi' + String.fromCharCode(8221)) === 'Hi');
  check('apostrophes are kept (they are part of words)', c("I'll call you") === "I'll call you");
  check('runs of spaces and newlines become single spaces', c('  a   b\n c ') === 'a b c', JSON.stringify(c('  a   b\n c ')));
  check('empty and null give an empty string', c('') === '' && c(null) === '' && c(undefined) === '');
}

console.log('\nTest 2 — what is editable');
{
  const sel = html.match(/var _EDITABLE_SELECTORS = \[([\s\S]*?)\];/);
  check('the editable list exists', !!sel);
  const list = sel ? sel[1] : '';
  check('the call summary is editable', /'\.ai-summary'/.test(list));
  check('the bullets of What\'s Working and What Needs to Change are editable', /'\.report-col-list:not\(\.report-col-roleplays\) li'/.test(list));
  check('the sample statements are editable', /'\.ai-flag-stmt'/.test(list));
  check('the flag text is editable', /'\.ai-flag-detail'/.test(list));
  check('the How-to-Change <li> itself is NOT editable (Select-All + type would wipe the statement and its Play button)',
        !/'\.report-col-list li'/.test(list));
}

console.log('\nTest 3 — the wiring');
{
  const open = fn('openModal');
  check('openModal makes the report editable once the modal is built',
        /initPlayButtons\(\);[\s\S]*_makeReportEditable\(document\.getElementById\('modal-ai-result'\)\)/.test(open));
  const apply = fn('applyAndClose');
  check('applyAndClose copies the analyst\'s words into data-tts BEFORE it serialises the modal',
        apply.indexOf('_syncTtsText(modalAIResult)') !== -1 && apply.indexOf('_syncTtsText(modalAIResult)') < apply.indexOf('modalAIResult.innerHTML'));
  check('applyAndClose makes the applied view editable with the same list', /_makeReportEditable\(out\)/.test(apply));
  const eff = fn('getEffectiveHTML');
  check('the HTML that is emailed / saved carries the edited statements', /_syncTtsText\(clone\)/.test(eff));
  const play = fn('playStatement');
  check('Play speaks the statement as it reads NOW', /_ttsTextOf\(btn\)/.test(play) && !/var text = btn\.getAttribute\('data-tts'\)/.test(play), play.slice(0, 160));
  check('an emptied statement plays nothing', /if \(!text\) return;/.test(play));
  check('the typed text is kept in step as it is typed', /addEventListener\('input'[\s\S]*_statementText\(host\)/.test(html));
  check('a Play button removed by backspacing is put back on blur', /addEventListener\('focusout'[\s\S]*_ensurePlayButton\(host\)/.test(html));
}

console.log('\nTest 4 — the evaluation page stays read-only: the sanitiser still strips contenteditable');
{
  const ctx = vm.createContext({ String, Object, parseInt, RegExp });
  vm.runInContext(extractFunction(read('HtmlSanitizer.gs'), 'sanitizeEvalHtml_'), ctx);
  const out = ctx.sanitizeEvalHtml_('<ul><li contenteditable="true">x</li></ul><div class="ai-flag-stmt" contenteditable="true">y</div>');
  check('no contenteditable survives, so agents cannot edit a saved evaluation', out.indexOf('contenteditable') === -1, out);
  check('the text does', /x/.test(out) && /y/.test(out));
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
