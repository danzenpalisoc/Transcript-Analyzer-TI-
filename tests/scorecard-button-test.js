/**
 * The "Scorecard Slides" button on the dashboard (index.html).
 *
 * The button must exist, open a dialog that says what will be built, and send ONLY the Audit Refs of the audits that pass
 * the current dashboard filters (plus an optional title and the filters in words). It must refuse an empty or oversized
 * selection before calling the server, ignore double clicks, show the result as a link only when it is a Google Slides link,
 * and show a clear message when something fails.
 *
 * Usage: node tests/scorecard-button-test.js
 */
const { load } = require('./lib/page-harness');
const { makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

const h = load();
const el = id => h.document.getElementById(id);
const row = (ref, name) => ({ 'Audit Ref': ref, 'Team Member': name });
function open(rowsR, rowsS) { h.reset(); h.ctx._dbRFiltered = rowsR; h.ctx._dbSFiltered = rowsS || []; h.ctx._scBusy = false; h.ctx.dbOpenScorecardDialog(); }

console.log('\nTest 1 — the button and the dialog are in the page');
{
  const html = h.html;
  const top = html.indexOf('id="db-topbar"'), tabs = html.indexOf('class="db-tab-bar"'), btn = html.indexOf('id="db-scorecard-btn"');
  const sec = html.indexOf('<!-- Row 4: Team Member Progression'), table = html.indexOf('class="ov-tm-table"');
  check('the button is in the header of "Team Member Trends & Repeat Offenders" (Overview tab), above the table', sec !== -1 && btn > sec && btn < table && html.slice(sec, btn).indexOf('Team Member Trends') !== -1);
  check('it is NOT in the top bar any more', top !== -1 && html.slice(top, tabs).indexOf('db-scorecard-btn') === -1);
  check('it sits beside the High / Moderate / Low legend, in the same header row', html.slice(btn, table).indexOf('class="ov-legend"') !== -1 && html.slice(sec, btn).indexOf('justify-content:space-between') !== -1);
  check('it opens the dialog', /id="db-scorecard-btn"[^>]*onclick="dbOpenScorecardDialog\(\)"/.test(html));
  check('it is a real button (type="button") called "Generate Slides"', /<button id="db-scorecard-btn" class="sc-trigger" type="button"/.test(html) && /Generate Slides\s*<\/button>/.test(html));
  check('its icon is a decorative SVG (not an emoji) hidden from screen readers', /<button id="db-scorecard-btn"[\s\S]*?<svg[^>]*aria-hidden="true"/.test(html));
  check('it is a filled purple button at least 44 px tall with a visible focus ring', /\.sc-trigger \{[^}]*min-height:44px[^}]*background:#4B286D/.test(html) && /\.sc-trigger:focus-visible/.test(html));
  check('the dialog and its buttons use the same name ("Generate Slides", "Generate slides")', /<h2 id="sc-title">Generate Slides<\/h2>/.test(html) && /dbScorecardStart\(\)">Generate slides<\/button>/.test(html) && html.indexOf('Create slides') === -1);
  check('the dialog is a labelled modal dialog', /id="db-scorecard-overlay"[^>]*role="dialog"[^>]*aria-modal="true"[^>]*aria-labelledby="sc-title"/.test(html));
  check('the status message is announced to screen readers', /id="sc-message" role="status" aria-live="polite"/.test(html));
  check('the title field has a visible label', /<label for="sc-label">/.test(html) && /<input id="sc-label"/.test(html));
  check('the dialog buttons are at least 44 px tall and have a visible focus ring', /\.sc-btn \{[^}]*min-height:44px/.test(html) && /\.sc-btn:focus-visible/.test(html));
  check('the spinner stops for people who ask for reduced motion', /prefers-reduced-motion: reduce\) \{ \.sc-spin/.test(html));
  const code = html.slice(html.indexOf('// ── Scorecard Slides'), html.indexOf('function dbUniq(arr){'));
  check('the dialog code never puts text into the page as HTML (no innerHTML)', code.length > 500 && code.indexOf('innerHTML') === -1);
  check('the dialog code calls the one server function the access test knows', /\.createScorecardDeck\(/.test(code));
}

console.log('\nTest 2 — what the dialog says and sends');
{
  open([row('R1', 'Agent A'), row('R2', ' agent   a '), row('R3', 'Agent B'), row('R1', 'Agent A')], [row('S1', 'Agent C'), row('S2', '')]);
  check('the dialog opens', el('db-scorecard-overlay').style.display === 'flex');
  check('it says how many agents and audits will be used (3 agents, 5 audits; a blank name is not an agent; a repeated ref counts once)', /3 agents/.test(el('sc-summary').textContent) && /5 audits/.test(el('sc-summary').textContent), el('sc-summary').textContent);
  check('the create button is shown', el('sc-create').style.display === 'inline-flex');
  h.setField('sc-label', '  MOB CLS Batch 15  ');
  h.setField('sh-date-from', '2026-10-01');
  h.ctx.dbScorecardStart();
  check('one call to createScorecardDeck is made', h.pending.length === 1 && h.pending[0].method === 'createScorecardDeck', JSON.stringify(h.pending.map(p => p.method)));
  const a = h.pending[0].args;
  check('it sends only the Audit Refs (each once, nothing else from the rows)', JSON.stringify(a[0]) === JSON.stringify(['R1', 'R2', 'R3', 'S1', 'S2']), JSON.stringify(a[0]));
  check('the title is sent trimmed', a[1] === 'MOB CLS Batch 15', a[1]);
  check('the active filters are sent in words', /Dates: 2026-10-01 to \.\.\./.test(a[2]), a[2]);
  check('while it works: the spinner shows, the form is hidden, both buttons are disabled', el('sc-working').style.display === 'block' && el('sc-form').style.display === 'none' && el('sc-create').disabled && el('sc-cancel').disabled);
  check('a message says it can take up to a minute', /up to a minute/.test(el('sc-message').textContent));
  h.ctx.dbScorecardStart(); h.ctx.dbOpenScorecardDialog(); h.ctx.dbScorecardClose();
  check('a double click, another open, or closing while it works does nothing', h.pending.length === 1 && el('db-scorecard-overlay').style.display === 'flex');
}

console.log('\nTest 3 — the result');
{
  open([row('R1', 'Agent A')], []);
  h.ctx.dbScorecardStart();
  h.pending[0].success({ success: true, url: 'https://docs.google.com/presentation/d/ABC/edit', slides: 2, agents: 1, audits: 1, skipped: 0 });
  check('a Google Slides link is shown as "Open in Google Slides"', el('sc-open').href === 'https://docs.google.com/presentation/d/ABC/edit' && el('sc-open').style.display === 'inline-flex');
  check('the summary says the deck is ready, how many slides, and that it is shared', /ready: 2 slides for 1 agent\./.test(el('sc-summary').textContent) && /shared with you/.test(el('sc-summary').textContent), el('sc-summary').textContent);
  check('the spinner stops, the create button is gone, Close is available', el('sc-working').style.display === 'none' && el('sc-create').style.display === 'none' && !el('sc-cancel').disabled && el('sc-cancel').textContent === 'Close');
  check('it can be used again afterwards', h.ctx._scBusy === false);

  open([row('R1', 'Agent A')], []);
  h.ctx.dbScorecardStart();
  h.pending[0].success({ success: true, url: 'https://docs.google.com/presentation/d/ABC/edit', slides: 3, agents: 2, skipped: 2 });
  check('left-out audits are mentioned', /2 audits were left out/.test(el('sc-message').textContent), el('sc-message').textContent);

  open([row('R1', 'Agent A')], []);
  h.ctx.dbScorecardStart();
  h.pending[0].success({ success: true, url: 'javascript:alert(1)', slides: 2, agents: 1 });
  check('a link that is not a Google Slides link is never used', el('sc-open').style.display !== 'inline-flex' && el('sc-open').href !== 'javascript:alert(1)' && el('sc-message').className === 'sc-error');
}

console.log('\nTest 4 — when something goes wrong');
{
  open([row('R1', 'Agent A')], []);
  h.ctx.dbScorecardStart();
  h.pending[0].success({ success: false, error: 'You have already built 5 decks in the last hour. Please try again later.' });
  check('the server\'s message is shown, in the error style', el('sc-message').textContent.indexOf('5 decks') !== -1 && el('sc-message').className === 'sc-error');
  check('the form comes back and the button says "Try again"', el('sc-form').style.display === 'block' && el('sc-create').textContent === 'Try again' && !el('sc-create').disabled && h.ctx._scBusy === false);

  open([row('R1', 'Agent A')], []);
  h.ctx.dbScorecardStart();
  h.pending[0].failure({ message: 'Script function not found' });
  check('a failed request shows a plain message', /The request failed: Script function not found/.test(el('sc-message').textContent) && el('sc-message').className === 'sc-error');

  open([row('R1', 'Agent A')], []);
  h.ctx.dbScorecardStart();
  h.pending[0].success(null);
  check('an empty reply does not crash and says so', /could not be built/.test(el('sc-message').textContent));
}

console.log('\nTest 5 — refused before calling the server');
{
  open([], []);
  check('no audits match: a clear message, no create button, no call', /No audits match the current filters/.test(el('sc-summary').textContent) && el('sc-create').style.display === 'none' && h.pending.length === 0);
  h.ctx.dbScorecardStart();
  check('...and starting anyway does nothing', h.pending.length === 0);

  const many = []; for (let i = 0; i < 41; i++) many.push(row('R' + i, 'Agent ' + i));
  open(many, []);
  check('41 agents: told to narrow the filters, no create button', /41 agents/.test(el('sc-summary').textContent) && /Narrow the filters/.test(el('sc-summary').textContent) && el('sc-create').style.display === 'none');
  h.ctx.dbScorecardStart();
  check('...and starting anyway does nothing', h.pending.length === 0);
  open(many.slice(0, 40), []);
  check('exactly 40 agents is allowed', el('sc-create').style.display === 'inline-flex');
  open([row('R1', 'x'.repeat(100)), row('R2', 'Agent A')], []);
  check('a Team Member that is a paragraph is not counted as an agent', /1 agent /.test(el('sc-summary').textContent), el('sc-summary').textContent);
}

console.log('\nTest 6 — the filters in words');
{
  h.reset();
  check('no filters: an empty text', h.ctx.dbScorecardFilterText() === '');
  h.setField('sh-date-from', '2026-10-01'); h.setField('sh-date-to', '2026-10-13'); h.setField('sh-sap-search', ' 2008221 ');
  const t = h.ctx.dbScorecardFilterText();
  check('dates and SAP ID are in the text', /Dates: 2026-10-01 to 2026-10-13/.test(t) && /SAP ID: 2008221/.test(t), t);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
