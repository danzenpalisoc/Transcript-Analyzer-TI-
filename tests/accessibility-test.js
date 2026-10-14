/**
 * Accessibility of the pages (code review batch 4).
 *
 *  Static rules on index.html and EvalView.html (they cannot slip back): labels tied to controls, dialogs with a role and a name,
 *  messages announced, no text under 12 px, no text colour that is too pale on white.
 *  Behaviour of the keyboard layer: Enter / Space on role=button, arrows on tabs, Escape for pills and dialogs, Tab kept inside a dialog.
 *  The Yes/No badge and the EvalView stars.
 *
 * Usage: node tests/accessibility-test.js
 */
const fs = require('fs');
const path = require('path');
const { load } = require('./lib/page-harness');
const { makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const read = f => fs.readFileSync(path.join(__dirname, '..', f), 'utf8');
const IX = read('index.html'), EV = read('EvalView.html');

console.log('\nTest 1 — labels, dialogs and messages (static)');
{
  const ids = new Set([...IX.matchAll(/\bid="([^"]+)"/g)].map(m => m[1]));
  const forTargets = [...IX.matchAll(/<label[^>]*\bfor="([^"]+)"/g)].map(m => m[1]);
  check('every <label for=...> points at an element that exists (the speaker select is built at run time)', forTargets.every(f => ids.has(f) || f === 'spk-agent-select'), forTargets.filter(f => !ids.has(f)).join(','));
  const bare = [...IX.matchAll(/<label>[^<]*(?:<span[^>]*>[^<]*<\/span>[^<]*)?<\/label>/g)].length;
  check('no plain <label> without for is left in the page markup (it was 12 form labels)', bare === 0, String(bare));
  ['sapId', 'participant', 'teamLeader', 'opsManager', 'lineOfBusiness', 'locale', 'vtid', 'startTime', 'customerBAN', 'interactionId', 'direction', 'duration']
    .forEach(id => check('the label of #' + id + ' is tied to it', new RegExp('<label for="' + id + '"').test(IX)));
  [['analysis-modal', 'analysis-modal-title'], ['audit-type-overlay', 'audit-type-title'], ['lob-overlay', 'lob-title'], ['spk-confirm-overlay', 'spk-title']].forEach(([id, title]) => {
    const tag = new RegExp('<div id="' + id + '"[^>]*>').exec(IX);
    check('#' + id + ' is a modal dialog named by #' + title, tag && /role="dialog"/.test(tag[0]) && /aria-modal="true"/.test(tag[0]) && tag[0].indexOf('aria-labelledby="' + title + '"') !== -1 && ids.has(title), tag && tag[0].slice(0, 120));
  });
  check('the analytics panel is a dialog named by its title', /<div id="analytics-panel" role="dialog" aria-labelledby="an-panel-title">/.test(IX) && ids.has('an-panel-title'));
  check('the progress overlays are announced (role=status, aria-live)', ['email-overlay', 'processing-overlay', 'db-export-overlay'].every(id => new RegExp('<div id="' + id + '"[^>]*role="status"[^>]*aria-live="polite"').test(IX)));
  check('an error is announced as an alert, a warning as a status', /id="alert-error" role="alert"/.test(IX) && /id="alert-warning" role="status"/.test(IX));
  check('the page language is set on both pages', /<html lang="en">/.test(IX) && /<html lang="en">/.test(EV));
  check('icon-only buttons have a name (close, chat, send)', /class="modal-close"[^>]*aria-label="Close"/.test(IX) && /id="chat-send"[^>]*aria-label=/.test(IX) && /id="chat-fab"[^>]*aria-label=/.test(IX) && /class="an-close"[^>]*aria-label=/.test(IX));
  check('the transcript box, the date boxes and the search boxes have a name', ['id="transcript" aria-label', 'id="sh-date-from" title="From date" aria-label', 'id="sh-date-to"   title="To date" aria-label', 'id="sh-sap-search" aria-label', 'id="chat-input" rows="1" aria-label'].every(s => IX.indexOf(s) !== -1));
  check('the pill search box built at run time has a name', /class="db-pill-panel-search" aria-label="Search '/.test(IX));
  check('the "only" link of a pill is a focusable button with a name', /data-v="[^"]*" role="button" tabindex="0" aria-label="Show only /.test(IX));
}

console.log('\nTest 2 — text size and colour (static)');
{
  const sizes = s => [...s.matchAll(/font-size:\s*([0-9.]+)px/g)].map(m => parseFloat(m[1]));
  check('index.html declares no text under 12 px', sizes(IX).every(n => n >= 12), 'under 12: ' + [...new Set(sizes(IX).filter(n => n < 12))].join(','));
  check('EvalView.html declares no text under 12 px', sizes(EV).every(n => n >= 12), 'under 12: ' + [...new Set(sizes(EV).filter(n => n < 12))].join(','));
  const pale = s => (s.match(/color:\s*#(888|999|aaa|bbb|AAA|BBB)\b/g) || []);
  check('no text colour #888 / #999 / #aaa / #bbb on white (all under 4.5 to 1)', pale(IX).length === 0 && pale(EV).length === 0, JSON.stringify(pale(IX).concat(pale(EV))));
  check('a visible keyboard focus ring is defined for buttons, inputs, tabs and role=button', /:focus-visible/.test(IX) && /outline: 3px solid #4B286D !important/.test(IX));
}

console.log('\nTest 3 — the keyboard layer');
{
  const h = load();
  const ctx = h.ctx;
  const mkEvent = (key, target, extra) => Object.assign({ key, target, defaultPrevented: false, preventDefault() { this.defaultPrevented = true; } }, extra || {});
  // a span with role=button: Enter and Space click it
  let clicks = 0;
  const span = { tagName: 'SPAN', getAttribute: k => (k === 'role' ? 'button' : null), click() { clicks++; } };
  ctx.a11yKeydown(mkEvent('Enter', span)); ctx.a11yKeydown(mkEvent(' ', span));
  check('Enter and Space activate a span that is a button by role', clicks === 2, String(clicks));
  const nat = { tagName: 'BUTTON', getAttribute: k => (k === 'role' ? 'button' : null), click() { clicks += 10; } };
  ctx.a11yKeydown(mkEvent('Enter', nat));
  check('a real <button> is left to the browser (no double click)', clicks === 2);
  ctx.a11yKeydown(mkEvent('a', span));
  check('other keys do nothing', clicks === 2);

  // tabs: arrows move and activate; Home / End jump; wraps around
  const tabs = [0, 1, 2, 3].map(i => ({ i, tagName: 'BUTTON', focused: 0, clicked: 0, focus() { this.focused++; }, click() { this.clicked++; }, getAttribute: k => (k === 'role' ? 'tab' : null) }));
  const bar = { querySelectorAll: () => tabs }; tabs.forEach(t => { t.parentNode = bar; });
  ctx.a11yKeydown(mkEvent('ArrowRight', tabs[0]));
  check('ArrowRight moves to the next tab and opens it', tabs[1].focused === 1 && tabs[1].clicked === 1);
  ctx.a11yKeydown(mkEvent('ArrowLeft', tabs[0]));
  check('ArrowLeft on the first tab wraps to the last', tabs[3].focused === 1 && tabs[3].clicked === 1);
  ctx.a11yKeydown(mkEvent('End', tabs[1])); ctx.a11yKeydown(mkEvent('Home', tabs[2]));
  check('End and Home jump to the last and the first tab', tabs[3].focused === 2 && tabs[0].focused === 1);

  // Escape closes an open pill first, then the top dialog that may be closed
  let closedPill = 0; let focusedPill = 0;
  ctx.dbClosePill = () => { closedPill++; ctx._openPill = null; };
  const pillEl = { focus() { focusedPill++; } };
  const dlg = {};
  const mkDialog = (id, visible) => { const el = h.document.getElementById(id); el.classList = { contains: c => visible && c === 'visible' }; el.style = { display: '' }; el.querySelectorAll = () => []; return el; };
  mkDialog('analysis-modal', false); mkDialog('spk-confirm-overlay', false); mkDialog('audit-type-overlay', false); mkDialog('lob-overlay', false); mkDialog('analytics-panel', false);
  h.els['pill-member'] = pillEl;
  ctx._openPill = 'member';
  const e1 = mkEvent('Escape', {});
  ctx.a11yKeydown(e1);
  check('Escape closes an open filter pill and puts the focus back on it', closedPill === 1 && focusedPill === 1 && e1.defaultPrevented);

  let modalClosed = 0, analyticsClosed = 0, spkClosed = 0;
  ctx.closeModal = () => { modalClosed++; }; ctx.closeAnalyticsPanel = () => { analyticsClosed++; }; ctx.spkConfirmCancel = () => { spkClosed++; };
  mkDialog('analysis-modal', true);
  ctx.a11yKeydown(mkEvent('Escape', {}));
  check('Escape closes the Process Call Script dialog', modalClosed === 1);
  mkDialog('analytics-panel', true);
  ctx.a11yKeydown(mkEvent('Escape', {}));
  check('with two open, Escape closes only the top one (the analytics panel is above)', analyticsClosed === 1 && modalClosed === 1);
  mkDialog('analysis-modal', false); mkDialog('analytics-panel', false);
  mkDialog('audit-type-overlay', true);
  ctx.a11yKeydown(mkEvent('Escape', {}));
  check('a dialog that asks for a required choice (audit type) is NOT closed by Escape', modalClosed === 1 && analyticsClosed === 1 && spkClosed === 0);
  mkDialog('audit-type-overlay', false);
  const spk = mkDialog('spk-confirm-overlay', false); spk.classList = { contains: () => false }; spk.style = { display: 'flex' };
  ctx.a11yKeydown(mkEvent('Escape', {}));
  check('the speaker confirmation (shown with style.display) closes with Escape', spkClosed === 1);
  spk.style = { display: 'none' };

  // Tab stays inside the open dialog
  const f1 = { offsetParent: {}, focused: 0, focus() { this.focused++; ctx.document.activeElement = this; } }, f2 = { offsetParent: {}, focused: 0, focus() { this.focused++; ctx.document.activeElement = this; } };
  const md = mkDialog('analysis-modal', true); md.querySelectorAll = () => [f1, f2]; md.contains = x => x === f1 || x === f2;
  h.document.activeElement = f2;
  const tabFwd = mkEvent('Tab', f2, { shiftKey: false }); ctx.a11yKeydown(tabFwd);
  check('Tab on the last control goes to the first (the focus cannot leave the dialog)', tabFwd.defaultPrevented && f1.focused === 1, String(f1.focused));
  h.document.activeElement = f1;
  const tabBack = mkEvent('Tab', f1, { shiftKey: true }); ctx.a11yKeydown(tabBack);
  check('Shift+Tab on the first control goes to the last', tabBack.defaultPrevented && f2.focused === 1);
  h.document.activeElement = { not: 'inside' };
  const tabOut = mkEvent('Tab', {}, { shiftKey: false }); ctx.a11yKeydown(tabOut);
  check('if the focus is somewhere outside the open dialog, Tab pulls it back in', tabOut.defaultPrevented && f1.focused === 2);
  mkDialog('analysis-modal', false);
  const free = mkEvent('Tab', {}, { shiftKey: false }); ctx.a11yKeydown(free);
  check('with no dialog open Tab is left alone', !free.defaultPrevented);
}

console.log('\nTest 4 — the Yes/No badge works from the keyboard');
{
  const h = load();
  const mkBadge = () => {
    const handlers = {}, attrs = {}, cls = new Set(['report-badge-yes']);
    return { attrs, handlers, textContent: 'Yes', style: {}, _resToggleBound: false,
      classList: { contains: c => cls.has(c), add: (...a) => a.forEach(x => cls.add(x)), remove: (...a) => a.forEach(x => cls.delete(x)) },
      setAttribute: (k, v) => { attrs[k] = v; }, addEventListener: (ev, fn) => { handlers[ev] = fn; },
      click() { handlers.click({ stopPropagation() {} }); } };
  };
  h.els['applied-output'] = Object.assign(h.document.getElementById('applied-output'), { contains: () => false });
  const badge = mkBadge();
  const chip = { querySelector: s => (/ai-chip-label/.test(s) ? { textContent: 'Issue Resolution' } : badge) };
  h.ctx._injectResolutionToggle({ querySelectorAll: () => [chip] });
  check('the badge is a focusable button with a name that says its state', badge.attrs.role === 'button' && badge.attrs.tabindex === '0' && /Issue resolved: Yes/.test(badge.attrs['aria-label']), JSON.stringify(badge.attrs));
  const k = { key: 'Enter', preventDefault() { this.p = true; } };
  badge.handlers.keydown(k);
  check('Enter flips it from Yes to No', badge.textContent === 'No' && k.p === true, badge.textContent);
  check('and the name follows (read out as No)', /Issue resolved: No/.test(badge.attrs['aria-label']), badge.attrs['aria-label']);
  badge.handlers.keydown({ key: ' ', preventDefault() {} });
  check('Space flips it back', badge.textContent === 'Yes');
  badge.handlers.keydown({ key: 'a', preventDefault() {} });
  check('other keys do nothing', badge.textContent === 'Yes');
}

console.log('\nTest 5 — the EvalView stars (static)');
{
  check('the stars are a radio group with a name and five radios', /id="ev-stars" role="radiogroup" aria-label="[^"]+"/.test(EV) && (EV.match(/role="radio"/g) || []).length === 5);
  check('each star has its own name ("3 stars") and only one is in the tab order to start', (EV.match(/aria-label="[1-5] stars?"/g) || []).length === 5 && (EV.match(/role="radio"[^>]*tabindex="0"/g) || []).length === 1);
  check('the arrow keys, Home, End, Enter and Space choose a star', ['ArrowRight', 'ArrowLeft', 'Home', 'End', "'Enter'"].every(k => EV.indexOf(k) !== -1) && /function _setStars/.test(EV));
  check('the chosen star is told to screen readers (aria-checked)', /setAttribute\('aria-checked'/.test(EV));
}

console.log('\nTest 6 — every pair of colours in the design tokens reads well (WCAG)');
{
  const root = /:root \{([\s\S]*?)\n\}/.exec(IX)[1];
  const tok = {}; [...root.matchAll(/(--[a-z0-9-]+):\s*(#[0-9A-Fa-f]{6})/g)].forEach(m => { tok[m[1]] = m[2]; });
  const lum = hex => { const c = [1, 3, 5].map(i => parseInt(hex.substr(i, 2), 16) / 255).map(v => (v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4))); return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]; };
  const ratio = (a, b) => { const x = lum(a), y = lum(b); return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05); };
  const W = '#FFFFFF', V = n => tok[n] || n;
  const pairs = [
    ['white text on the navigation bar', W, '--c-nav', 4.5], ['white text on the open tab', W, '--c-nav-active', 4.5], ['green accent on the navigation bar', '--c-neon', '--c-nav', 4.5], ['green accent on the hero start', '--c-neon', '--c-hero-start', 4.5],
    ['green accent on deep purple', '--c-neon', '--c-deep', 4.5], ['white on deep purple (buttons, labels)', W, '--c-deep', 4.5],
    ['white on the hero middle', W, '--c-hero-mid', 4.5], ['white on the hero start', W, '--c-hero-start', 4.5],
    ['green text on white', '--telus-green', W, 4.5], ['white text on the green fill', W, '--telus-green', 4.5], ['dark green text on white', '--c-green-text', W, 4.5],
    ['dark green on the mint chip', '--c-green-dark', '--c-tier-good', 4.5], ['dark green on the mint tile', '--c-green-dark', '--c-tier-good-bg', 4.5],
    ['rose text on the rose chip', '--c-tier-low-text', '--c-tier-low', 4.5], ['rose text on the rose tile', '--c-tier-low-text', '--c-tier-low-bg', 4.5],
    ['amber text on the yellow chip', '--c-tier-mid-text', '--c-tier-mid', 4.5], ['amber text on the yellow tile', '--c-tier-mid-text', '--c-tier-mid-bg', 4.5],
    ['body text on white', '--c-body', W, 4.5], ['body text on the page', '--c-body', '--telus-grey-bg', 4.5], ['title text on the page', '--telus-text', '--telus-grey-bg', 4.5],
    ['muted text on white', '--c-muted', W, 4.5], ['muted text on the page', '--c-muted', '--telus-grey-bg', 4.5], ['sub text on white', '--telus-subtext', W, 4.5],
    ['purple on the purple chip tint', '--telus-purple', '--c-tint-purple', 4.5], ['purple text on white', '--telus-purple', W, 4.5], ['red text on white', '--telus-red', W, 4.5],
  ];
  const bad = pairs.filter(p => ratio(V(p[1]), V(p[2])) < p[3]).map(p => p[0] + ' ' + ratio(V(p[1]), V(p[2])).toFixed(2));
  check('all ' + pairs.length + ' pairs are at least 4.5 to 1', bad.length === 0, bad.join(' | '));
  check('every colour the pairs name exists as a token (a typo would hide a failure)', pairs.every(p => [p[1], p[2]].every(x => x === W || tok[x])), JSON.stringify(pairs.filter(p => [p[1], p[2]].some(x => x !== W && !tok[x])).map(p => p[0])));
}

console.log('\nTest 7 — the brand colours are the TELUS palette');
{
  // The palette table of the brand: purples DFD5E5 C1AFCD B9A0CA 946DAF 6E3D8F 4B286D; greens DBF9BD C7EF9F AAE074 84D533 66CC00 5AB400 248700; greys 54595F 6A6E74 9DA5AA.
  const PALETTE = ['#DFD5E5', '#C1AFCD', '#B9A0CA', '#946DAF', '#6E3D8F', '#4B286D', '#DBF9BD', '#C7EF9F', '#AAE074', '#84D533', '#66CC00', '#5AB400', '#248700', '#54595F', '#6A6E74', '#9DA5AA'];
  const root = /:root \{([\s\S]*?)\n\}/.exec(IX)[1];
  const tok = {}; [...root.matchAll(/(--[a-z0-9-]+):\s*(#[0-9A-Fa-f]{6})/g)].forEach(m => { tok[m[1]] = m[2].toUpperCase(); });
  const brand = ['--telus-purple', '--telus-green', '--c-nav', '--c-nav-active', '--c-deep', '--c-hero-mid', '--c-glow', '--c-neon', '--c-green-text', '--c-tint-purple', '--c-tier-good', '--c-body', '--c-muted'];
  check('every brand token is a colour of the TELUS table', brand.every(n => PALETTE.indexOf(tok[n]) !== -1), brand.filter(n => PALETTE.indexOf(tok[n]) === -1).map(n => n + '=' + tok[n]).join(', '));
  check('the purple is 4B286D and the green is 66CC00 (the two colours of the brand)', tok['--telus-purple'] === '#4B286D' && tok['--c-deep'] === '#4B286D' && tok['--c-nav'] === '#4B286D' && tok['--c-neon'] === '#66CC00' && tok['--c-glow'] === '#66CC00');
  check('the lighter purple (open tab, hero) is 6E3D8F', tok['--c-nav-active'] === '#6E3D8F' && tok['--c-hero-mid'] === '#6E3D8F');
  // the values that were taken from screenshots of a wide-gamut screen, which records the TELUS colours more saturated, must not come back
  const SAMPLED = ['#6F008E', '#57006E', '#811F9C', '#6D008B', '#440058', '#00EB00', '#47C50F', '#00802F', '#007332', '#F8E0FF', '#BAFFD2', '#A146BB', '#25EA26', '#00DA00', '#8D0093', '#650070'];
  const found = SAMPLED.filter(c => IX.toUpperCase().indexOf(c) !== -1);
  check('no colour sampled from a screenshot is left in the page (they are the TELUS colours seen on a wide-gamut screen)', found.length === 0, found.join(', '));
  check('the open tab reads white on the lighter purple (the green accent there is only 3.7 to 1) and keeps the green as its underline', IX.indexOf('.nav-tab.nav-active { background: var(--c-nav-active); color: #fff; border-bottom-color: var(--c-neon); }') !== -1);
  check('the Highlights titles use the dark green on the mint box (248700 on F0FFF4 is only 4.47 to 1)', IX.indexOf('.ov-hl-high .ov-hl-title { color: var(--c-green-dark); }') !== -1 && IX.indexOf('.an-hl-high .an-hl-title { color: var(--c-green-dark); }') !== -1);
  check('no chip puts the green 248700 text on the mint EDF7E6 (4.19 to 1; the chips use 1A5E00)', !/background:\s*#EDF7E6;\s*color:\s*#248700/i.test(IX));
  check('the report documents keep the TELUS purple too (the print window and the report style)', (IX.match(/#4B286D/gi) || []).length >= 10);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
