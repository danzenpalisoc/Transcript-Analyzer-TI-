/**
 * The new navigation: Home | Overview | Repeats | Sales | Audit Log | Analyzer Form.
 *
 * One view at a time (the form page and the dashboard are hidden, never removed), the router reuses openDashboard / closeDashboard /
 * dbSwitchTab, the navigation bar is a proper tab list, and the dashboard shell sits under the bar.
 *
 * Usage: node tests/navigation-test.js
 */
const fs = require('fs');
const path = require('path');
const { load } = require('./lib/page-harness');
const { makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const IX = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

console.log('\nTest 1 — the navigation bar (static)');
{
  const tabs = [...IX.matchAll(/<button type="button" class="nav-tab[^"]*" role="tab" id="tab-(\w+)" data-view="(\w+)" aria-selected="(true|false)" aria-controls="([\w-]+)" tabindex="(-?\d)" onclick="showView\('(\w+)'\)">[\s\S]*?<span>([^<]+)<\/span>/g)];
  check('six tabs, in the order Home, Overview, Repeats, Sales, Audit Log, Analyzer Form', tabs.map(m => m[2]).join() === 'home,overview,repeats,sales,auditlog,form' && tabs.map(m => m[7]).join('|') === 'Home|Overview|Repeats &amp; Transfer|Sales|Audit Log|Analyzer Form', tabs.map(m => m[7]).join('|'));
  check('each tab opens its own view (onclick, data-view and id agree)', tabs.every(m => m[1] === m[2] && m[2] === m[6]));
  const ids = new Set([...IX.matchAll(/\bid="([^"]+)"/g)].map(m => m[1]));
  check('each tab controls an element that exists', tabs.every(m => ids.has(m[4])), tabs.filter(m => !ids.has(m[4])).map(m => m[4]).join());
  check('only the first tab is selected and in the tab order to start with (roving tabindex)', tabs.filter(m => m[3] === 'true').length === 1 && tabs[0][3] === 'true' && tabs.filter(m => m[5] === '0').length === 1 && tabs[0][5] === '0');
  check('the bar is a tab list with a name', /<nav class="nav-tabs" role="tablist" aria-label="Main views" id="app-tabs">/.test(IX));
  check('the tab icons are SVG and hidden from screen readers (no emoji as icons)', (IX.match(/class="nav-tab[\s\S]{0,400}?<svg[^>]*aria-hidden="true"/g) || []).length === 6);
  check('the page starts on Home (data-view="home" in the HTML, so the first paint is Home)', /<body data-view="home">/.test(IX));
}

console.log('\nTest 2 — the shell (static)');
{
  check('only the Analyzer Form view shows the form page; Home shows only for Home', /body\[data-view\] \.page \{ display: none; \}/.test(IX) && /body\[data-view="form"\] \.page \{ display: block; \}/.test(IX) && /body\[data-view="home"\] #home-view \{ display: block; \}/.test(IX));
  check('whichever view is showing is the main landmark of the page (Home is <main>; the form page and the dashboard are role=main), so a screen reader can always jump to the content', /<main id="home-view"/.test(IX) && /<div class="page" id="form-page" role="main">/.test(IX) && /<div id="dashboard-panel" role="main" aria-label="Dashboard">/.test(IX));
  check('the form page has an id and holds the Process Script button (it only exists on that view now)', /<div class="page" id="form-page" role="main">/.test(IX) && /<div class="page" id="form-page" role="main">[\s\S]*?id="btn-process"/.test(IX));
  // position: fixed at 56 px from the window top put the title row behind the bar when the page sits lower than the window top (the Apps Script frame, by 16 px)
  const panelCss = (/#dashboard-panel \{([^}]*)\}/.exec(IX) || [])[1] || '';
  const topbarCss = (/#db-topbar \{([^}]*)\}/.exec(IX) || [])[1] || '';
  check('the dashboard panel is in the page flow under the bar (not fixed), so it can never start behind the bar', /position:\s*relative/.test(panelCss) && !/position:\s*(fixed|absolute)/.test(panelCss) && !/inset:/.test(panelCss), panelCss.replace(/\s+/g, ' '));
  check('the title row sticks right under the bar while the page scrolls', /position:\s*sticky/.test(topbarCss) && /top:\s*var\(--nav-h\)/.test(topbarCss), topbarCss.replace(/\s+/g, ' '));
  const z = re => Number((re.exec(IX) || [])[1]);
  check('the bar is above the dashboard panel (z-index)', z(/\.app-nav \{[^}]*z-index: (\d+)/) > z(/#dashboard-panel \{[^}]*z-index: (\d+)/), z(/\.app-nav \{[^}]*z-index: (\d+)/) + ' vs ' + z(/#dashboard-panel \{[^}]*z-index: (\d+)/));
  check('the old dashboard tab bar stays in the page but hidden (dbSwitchTab and the tests use it)', /\.db-tab-bar \{[^}]*display: none;/.test(IX) && (IX.match(/class="db-tab-btn[^"]*" data-tab="(overview|repeats|sales|auditlog)"/g) || []).length === 4);
  check('the dashboard toolbar kept the Refresh button and the data status, and lost the back button', /id="db-refresh-btn"/.test(IX) && /id="db-cache-status"/.test(IX) && !/class="db-back"/.test(IX));
  check('the page opens on Home and still loads the dashboard data right away', /showView\('home'\);\n\s*if \(!_dbLoaded\) loadDashboardData\(\);/.test(IX));
}

console.log('\nTest 3 — the router');
{
  const h = load();
  const ctx = h.ctx, attrs = {}, calls = [];
  h.document.body.setAttribute = (k, v) => { attrs[k] = v; };
  const tab = v => { const b = { dataset: { view: v }, attrs: {}, cls: new Set(), getAttribute(k) { return k === 'data-view' ? v : this.attrs[k]; }, setAttribute(k, x) { this.attrs[k] = x; }, classList: { toggle: (c, on) => (on ? b.cls.add(c) : b.cls.delete(c)) } }; return b; };
  const navTabs = ['home', 'overview', 'repeats', 'sales', 'auditlog', 'form'].map(tab);
  const dbBtns = ['overview', 'repeats', 'sales', 'auditlog'].map(n => ({ tab: n }));
  h.document.querySelectorAll = sel => (sel === '.nav-tab' ? navTabs : []);
  h.document.querySelector = sel => { const m = /data-tab="(\w+)"/.exec(sel); return m ? dbBtns.find(b => b.tab === m[1]) : null; };
  ctx.openDashboard = () => calls.push('open');
  ctx.closeDashboard = () => calls.push('close');
  ctx.dbSwitchTab = (name, btn) => calls.push('tab:' + name + ':' + (btn && btn.tab));
  const title = h.document.getElementById('db-toolbar-title');
  const reset = () => { calls.length = 0; };

  ctx.showView('overview');
  check('Overview opens the dashboard on its Overview tab, once', calls.join() === 'open,tab:overview:overview' && attrs['data-view'] === 'overview', calls.join());
  check('and the toolbar says Overview', title.textContent === 'Overview');
  reset(); ctx.showView('sales');
  check('Sales stays in the dashboard and switches to the Sales tab (no close and open)', calls.join() === 'open,tab:sales:sales' && title.textContent === 'Sales Analyzer — Agent', calls.join());
  reset(); ctx._tabRenderHash = { overview: 'x' }; ctx.showView('auditlog');
  check('moving between dashboard tabs keeps the render memo (no needless redraw)', ctx._tabRenderHash.overview === 'x');
  reset(); ctx.showView('form');
  check('leaving the dashboard for the form closes the dashboard', calls.join() === 'close' && attrs['data-view'] === 'form', calls.join());
  reset(); ctx.showView('home');
  check('going from the form to Home does not touch the dashboard', calls.length === 0 && attrs['data-view'] === 'home', calls.join());
  reset(); ctx._tabRenderHash = { overview: 'x' }; ctx.showView('repeats');
  check('coming to the dashboard from Home clears the render memo (charts drawn while it was hidden have no size, so they are drawn again)', Object.keys(ctx._tabRenderHash).length === 0 && calls.join() === 'open,tab:repeats:repeats', JSON.stringify(ctx._tabRenderHash) + ' ' + calls.join());
  reset(); ctx.showView('nonsense');
  check('an unknown name goes to Home', attrs['data-view'] === 'home' && calls.join() === 'close', attrs['data-view'] + ' ' + calls.join());

  ctx.showView('sales');
  const state = () => navTabs.map(b => b.attrs['aria-selected'] + '/' + b.attrs.tabindex + '/' + (b.cls.has('nav-active') ? 'A' : '-')).join(' ');
  check('the bar marks only the open tab: selected, in the tab order, highlighted', navTabs.filter(b => b.attrs['aria-selected'] === 'true').length === 1 && navTabs[3].attrs['aria-selected'] === 'true' && navTabs[3].attrs.tabindex === '0' && navTabs.filter(b => b.attrs.tabindex === '0').length === 1 && navTabs[3].cls.has('nav-active') && navTabs.filter(b => b.cls.has('nav-active')).length === 1, state());
}

console.log('\nTest 4 — the Home page (static)');
{
  const home = /<main id="home-view"[\s\S]*?<\/main>/.exec(IX)[0];
  const plain = home.replace(/<[^>]+>/g, ' ').replace(/&amp;/g, '&').replace(/&middot;/g, '·').replace(/&mdash;/g, '—').replace(/&rarr;/g, '→').replace(/\s+/g, ' ');
  check('one h1 (the hero title) and h2 section titles in order', (home.match(/<h1\b/g) || []).length === 1 && /<h1 id="home-title">/.test(home) && (home.match(/<h2 /g) || []).length === 2 && /aria-labelledby="home-title"/.test(home));
  const cards = [...home.matchAll(/<article class="home-card" onclick="showView\('(\w+)'\)">[\s\S]*?<h3 class="home-card-title">([^<]+)<\/h3>[\s\S]*?<button type="button" class="home-open">Open ([^<]+) <span/g)];
  check('five cards open Overview, Repeats, Sales, Audit Log and Analyzer Form (not Home itself), and each button names its destination', cards.map(m => m[1]).join() === 'overview,repeats,sales,auditlog,form' && cards.every(m => m[2] === m[3]), cards.map(m => m[1] + ':' + m[2] + '/' + m[3]).join(' '));
  check('every card has a real button inside it (the keyboard path), and the icons are SVG hidden from screen readers', (home.match(/<button type="button" class="home-open">/g) || []).length === 5 && (home.match(/class="home-icon"><svg[^>]*aria-hidden="true"/g) || []).length === 5);
  check('the four "how it works" cards are there', (home.match(/<article class="home-how">/g) || []).length === 4);
  ['home-stat-evals', 'home-stat-agents', 'home-stat-split', 'home-info-latest', 'home-info-asof'].forEach(id => check('the number #' + id + ' has a place', home.indexOf('id="' + id + '"') !== -1));
  check('the "latest data" box is announced politely when it changes', /class="home-info" role="status" aria-live="polite"/.test(home));

  // what the page says must be what the code does
  const cfg = fs.readFileSync(path.join(__dirname, '..', 'Config.gs'), 'utf8');
  const th = k => { const m = new RegExp(k + ':\\s*\\{[^}]*target:\\s*(\\d+)').exec(cfg); return m ? m[1] : null; };
  check('the targets on Home are the ones in SCORECARD_THRESHOLDS (repeat risk, resolution, flags, transfer, sales attempted)', plain.indexOf('Avg repeat risk — ' + th('repeatRisk') + '% or lower') !== -1 && plain.indexOf('Resolution rate — ' + th('resolution') + '% or higher') !== -1 && plain.indexOf('Critical flags — ' + th('flags') + ' per audit') !== -1 && plain.indexOf('Transfer rate — ' + th('transfer') + '% or lower') !== -1 && plain.indexOf('Sales attempted — ' + th('salesAttempt') + '% or higher') !== -1, [th('repeatRisk'), th('resolution'), th('flags'), th('transfer'), th('salesAttempt')].join());
  const ms = name => { const m = new RegExp(name + '\\s*=\\s*([0-9 *]+);').exec(IX); return m ? m[1].split('*').reduce((p, x) => p * Number(x.trim()), 1) : null; };
  check('the browser copies named on Home match the code: dashboard rows 30 minutes, audit log 15 minutes, roster 2 hours', ms('LS_CACHE_MAX_AGE') === 30 * 60000 && ms('LS_AUDITLOG_AGE') === 15 * 60000 && ms('LS_ROSTER_AGE') === 2 * 3600000 && /dashboard rows 30 minutes, audit log 15 minutes, roster 2 hours/.test(plain), [ms('LS_CACHE_MAX_AGE'), ms('LS_AUDITLOG_AGE'), ms('LS_ROSTER_AGE')].join());
  check('the deck limit named on Home is the real one (SC_MAX_AGENTS)', Number((/var SC_MAX_AGENTS = (\d+)/.exec(IX) || [])[1]) === 40 && /up to 40 agents/.test(plain));
  check('the example Audit Ref has the real format (NHA-yyyymmdd-nnnn)', /NHA-\d{8}-\d{4}/.test(plain) && /^NHA-\d{8}-\d{4}$/.test((/\(for example (NHA-[\d-]+)\)/.exec(plain) || [])[1] || ''));
  const bad = ['TELUS', 'NHA'].filter(() => false);
  check('the green glow of the hero gives way to a plain purple gradient on narrower screens, where the text would sit on green (white on the glow is about 2.4 to 1)', /@media \(max-width: 1100px\) \{\s*\.home-hero \{ background: linear-gradient\(105deg, var\(--c-hero-start\) 0%, var\(--c-deep\) 100%\)/.test(IX) && /\.home-lead \{ max-width: 680px;/.test(IX));
  check('the page does not promise anything the code does not do (no "real-time", no "automatically every")', !/real-time|automatically every/i.test(plain), bad.join());
}

console.log('\nTest 5 — the numbers on Home are counted like the Overview tiles');
{
  const h = load();
  const el = id => h.document.getElementById(id);
  const row = (n, type, ts, agent) => ({ 'Audit Ref': 'R' + n, 'Analysis Type': type, 'Team Member': agent, 'Timestamp': ts });
  h.ctx.homeUpdate();
  check('with no data the numbers are left as they were (a dash, not 0)', el('home-stat-evals').textContent === '' || el('home-stat-evals').textContent === '—' || el('home-stat-evals').textContent === undefined);
  h.ctx._dbAllData = [
    row(1, 'Repeats & Transfer Analyzer', '2026-10-01T10:00:00', 'Maria Santos'), row(2, 'Repeats & Transfer Analyzer', '2026-10-03T10:00:00', '  maria   santos '), row(3, 'Sales Analyzer', '2026-10-05T10:00:00', 'Aisha Khan'),
    row(4, 'Repeats & Transfer Analyzer', '2026-10-02T10:00:00', 'Name\nSecond'), row(5, 'Repeats & Transfer Analyzer', '2026-10-02T10:00:00', 'x'.repeat(90)), row(6, '', '2026-12-31T10:00:00', 'No Type'),
  ];
  h.ctx._dbDataTs = Date.parse('2026-10-14T03:39:00');
  h.ctx.homeUpdate();
  check('evaluations are the Repeats and Sales audits (an audit with no analysis type is not counted), as the Overview tile says', el('home-stat-evals').textContent === '5', el('home-stat-evals').textContent);
  check('agents are counted by the Overview rule: a name with a line break or over 80 characters is not an agent; spacing is cleaned', el('home-stat-agents').textContent === '3', el('home-stat-agents').textContent);
  check('the split is Repeats / Sales', el('home-stat-split').textContent === '4 / 1', el('home-stat-split').textContent);
  check('the latest audit is the newest of the counted ones (Oct 5, not the untyped Dec 31)', /Oct/.test(el('home-info-latest').textContent) && /5/.test(el('home-info-latest').textContent) && !/Dec/.test(el('home-info-latest').textContent), el('home-info-latest').textContent);
  check('"data loaded" is the time the data was produced', /Oct/.test(el('home-info-asof').textContent) && /14/.test(el('home-info-asof').textContent), el('home-info-asof').textContent);
  const h2 = load(); h2.ctx.dbHandleData({ rows: [], ts: 1760000000000 }, false);
  check('the time of the data is remembered when the dashboard data arrives (dbHandleData)', h2.ctx._dbDataTs === 1760000000000, String(h2.ctx._dbDataTs));
  const dash = h.ctx.dbUniq(['a', 'a', 'b', '']).length;
  check('it uses the same helper as the dashboard (dbUniq), not a second counting method', dash === 2 && /var agents = dbUniq\(all\.map/.test(IX.slice(IX.indexOf('function homeUpdate'), IX.indexOf('function openDashboard'))));
  check('it runs whenever the dashboard data has been rendered (dbFinishRender) so Refresh updates it', /function dbFinishRender\(\) \{[\s\S]{0,300}homeUpdate\(\);/.test(IX));
}

console.log('\nTest 6 — the loading screen');
{
  const vm = require('vm');
  check('it is the first thing in the body, so it paints before any script runs', /<body data-view="home">\s*<!--[^>]*-->\s*<div id="app-splash"/.test(IX));
  check('it is announced politely and named (role=status, aria-label), and its decoration is hidden from screen readers', /<div id="app-splash" role="status" aria-live="polite" aria-label="Loading Call Analyzer">/.test(IX) && /class="splash-bubbles" aria-hidden="true"/.test(IX) && /class="splash-dots" aria-hidden="true"/.test(IX) && /class="splash-curve" aria-hidden="true"/.test(IX));
  check('it has the green curve, the bubbles and the three dots of the reference', (IX.match(/class="splash-b"/g) || []).length >= 10 && /splash-curve/.test(IX) && (/class="splash-dots"[^>]*>(<span><\/span>){3}/.test(IX)));
  check('it respects reduced motion (no floating, no pulsing)', /@media \(prefers-reduced-motion: reduce\) \{ \.splash-b, \.splash-dots span \{ animation: none; \}/.test(IX));
  check('it is above everything else (z-index above the dialogs) and ignores clicks once it is hiding', Number((/#app-splash \{[^}]*z-index: (\d+)/.exec(IX) || [])[1]) > 10000 && /#app-splash\.splash-hide \{ opacity: 0; pointer-events: none; \}/.test(IX));
  check('with scripts off it does not block the page', /<noscript><style>#app-splash \{ display: none; \}<\/style><\/noscript>/.test(IX));
  check('the page init takes it away once Home is ready', /showView\('home'\);\n\s*if \(!_dbLoaded\) loadDashboardData\(\);\n\s*if \(typeof hideSplash === 'function'\) hideSplash\(\);/.test(IX));

  // the small inline script, run on its own with a fake clock
  const src = /<script>\s*(window\.__splashStart[\s\S]*?)<\/script>/.exec(IX)[1];
  function run(elapsedMs) {
    const timers = [], cls = new Set(), attrs = {}, style = {}; let now = 1000000;
    const el = { classList: { add: c => cls.add(c) }, getAttribute: k => attrs[k], setAttribute: (k, v) => { attrs[k] = v; }, style };
    const win = {}; const sb = { window: win, document: { getElementById: id => (id === 'app-splash' ? el : null) }, Date: { now: () => now }, setTimeout: (fn, ms) => { timers.push({ fn, ms }); return timers.length; } };
    vm.runInNewContext(src, sb);
    now += elapsedMs;
    return { win, timers, cls, attrs, style, el, advance: () => { const todo = timers.splice(0); todo.forEach(x => x.fn()); } };
  }
  const r = run(100);
  check('the 8 second failsafe is set up when the script runs', r.timers.length === 1 && r.timers[0].ms === 8000, JSON.stringify(r.timers.map(x => x.ms)));
  r.timers.length = 0; r.win.hideSplash();
  check('100 ms after the start, it waits the other 500 ms (never a flash)', r.timers.length === 1 && r.timers[0].ms === 500, JSON.stringify(r.timers.map(x => x.ms)));
  r.advance();
  check('then it fades (class) and is marked gone', r.cls.has('splash-hide') && r.attrs['data-gone'] === '1');
  check('and it is taken out of the layout 350 ms later', r.timers.length === 1 && r.timers[0].ms === 350 && (r.advance(), r.style.display === 'none'), JSON.stringify(r.style));
  const slow = run(5000); slow.timers.length = 0; slow.win.hideSplash();
  check('when the page took 5 seconds to load it goes at once (no extra wait)', slow.timers.length === 1 && slow.timers[0].ms === 0, JSON.stringify(slow.timers.map(x => x.ms)));
  const twice = run(5000); twice.timers.length = 0; twice.win.hideSplash(); twice.advance(); twice.timers.length = 0; twice.win.hideSplash();
  check('calling it again does nothing (the 8 second failsafe after a normal hide is harmless)', twice.timers.length === 0);
  check('it never throws when the element is missing', (() => { const sb = { window: {}, document: { getElementById: () => null }, Date: { now: () => 1 }, setTimeout: () => 1 }; vm.runInNewContext(src, sb); try { sb.window.hideSplash(); return true; } catch (e) { return false; } })());
  check('the dashboard loading line uses the same dots, and the old spinner is gone from it', /class="db-dots" aria-hidden="true"><span><\/span><span><\/span><span><\/span><\/span>Loading dashboard data/.test(IX) && !/<div class="db-spin"><\/div>Loading dashboard data/.test(IX));
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
