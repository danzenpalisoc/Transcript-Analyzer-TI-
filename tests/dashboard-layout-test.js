/**
 * Dashboard layout fixes found when each tab was reviewed (code review of the new look):
 *  1. The KPI tiles sat in a 900 px wrapper and wrapped 5 + 1 (Overview) or 5 + 3 (Repeats, Sales). They now share a row as wide as the page.
 *  2. The "Flags / Violation Breakdown" doughnut gave every flag text its own slice, so the top 8 were tiny and "Other" was ~99%. It is now
 *     "Critical Flags per Audit": how many audits carry 0, 1, 2, 3 or 4+ flags.
 *  3. Wide tables: a scrollbar that can be seen, and the first column stays while the others scroll.
 *  4. No chart text under 12 px.
 *
 * Usage: node tests/dashboard-layout-test.js
 */
const fs = require('fs');
const path = require('path');
const { load } = require('./lib/page-harness');
const { makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const IX = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

console.log('\nTest 1 — the KPI tiles (static)');
{
  const wrap = /<!-- Shared stats row[\s\S]*?<div style="([^"]*)">([\s\S]*?)<\/div>\s*\n\s*\n/.exec(IX);
  const block = IX.slice(IX.indexOf('<!-- Shared stats row'), IX.indexOf('id="al-stats"') + 80);
  check('the row is as wide as the page content (up to 1280 px, like .db-body) and takes the full width (a flex item with auto margins would shrink to its content), not 900 px', /width:100%;max-width:1280px/.test(block) && !/max-width:900px/.test(block), block.slice(0, 160));
  const rows = ['ov-stats', 'r-stats', 's-stats', 'al-stats'].map(id => new RegExp('<div class="db-stats-row" id="' + id + '"([^>]*)></div>').exec(IX));
  check('the four rows have no inline display:flex, so the grid rule applies (the hidden ones are only display:none)', rows.every(m => m) && rows.every(m => !/display:flex/.test(m[1])) && rows[1][1].trim() === 'style="display:none"' && rows[0][1].trim() === '', rows.map(m => m && m[1]).join('|'));
  const css = /\.db-stats-row \{([^}]*)\}/.exec(IX)[1], tile = /\.db-stat \{([^}]*)\}/.exec(IX)[1];
  check('the row wraps and is centred, and the tiles grow from 140 px to at most 230 px to share it (an auto-fit grid with a fixed maximum counts its columns by the maximum, so only 5 fitted)', /display: flex;/.test(css) && /flex-wrap: wrap;/.test(css) && /justify-content: center;/.test(css) && /flex: 1 1 140px;/.test(tile) && /max-width: 230px;/.test(tile), css + ' || ' + tile);
  // the maths of the grid for the real tile counts at the content width
  const fit = (w, n, basis, gap) => n * basis + (n - 1) * gap <= w;
  check('at the full 1232 px of content a flex line holds 8 tiles of 140 px with 12 px gaps (Repeats and Sales), and so 6 (Overview) and 5 (Audit Log)', fit(1232, 8, 140, 12) && fit(1232, 6, 140, 12) && !fit(1232, 9, 140, 12));
  check('dbSwitchTab shows and hides the rows with style.display = "" or "none" only, which the grid keeps working with', /getElementById\('ov-stats'\)\.style\.display = tab === 'overview'\s*\? '' : 'none'/.test(IX));
}

console.log('\nTest 2 — Critical Flags per Audit (static)');
{
  check('the three doughnuts are titled Critical Flags per Audit and the old title is gone', (IX.match(/Critical Flags per Audit<\/div>/g) || []).length === 3 && IX.indexOf('Flags / Violation Breakdown') === -1);
  check('no chart groups flag texts into "Other" any more', !/addRow\(\['Other'/.test(IX));
  check('the three doughnuts call the one function', (IX.match(/dbDrawFlagsPerAudit\('/g) || []).length === 1 && /dbDrawFlagsPerAudit\(pfx \+ '-chart-flagdist', rows\)/.test(IX) && /dbDrawFlagsPerAudit\('ov-chart-rca', all\)/.test(IX));
  check('the helper sits after the shared dashboard helper block (that block must stay identical to Dashboard.html)', IX.indexOf('function _dbFlagsPerAudit') > IX.indexOf('// ── end of dashboard helpers ──'));
  check('no chart text is under 12 px (legend, tooltip, axes)', !/fontSize:\s*(?:[0-9]|10|11)\b/.test(IX), (IX.match(/fontSize:\s*\d+/g) || []).filter(x => Number(x.replace(/\D/g, '')) < 12).join());
}

console.log('\nTest 3 — wide tables (static)');
{
  const css = IX.slice(IX.indexOf('.db-tbl-wrap {'), IX.indexOf('.db-tbl-wrap {') + 1100);
  check('the scrollbar is visible and styled (thin, purple), so people can see the table scrolls', /scrollbar-width: thin;/.test(css) && /\.db-tbl-wrap::-webkit-scrollbar \{ height: 10px; \}/.test(css) && /::-webkit-scrollbar-thumb \{ background: var\(--c-nav\)/.test(css));
  check('the sticky column has a soft edge on its right, so it is clear that the other columns slide under it', /\.db-tbl th:first-child, \.db-tbl tbody td:first-child \{ box-shadow: 2px 0 4px rgba\(58,31,87,\.14\); \}/.test(IX));
  check('the first column stays while the rest scrolls, on a solid background (header, rows, even rows, hover)', /\.db-tbl th:first-child \{ position: sticky; left: 0; z-index: 3; background: var\(--c-deep\); \}/.test(css) && /\.db-tbl tbody td:first-child \{ position: sticky; left: 0; z-index: 2; background: #fff; \}/.test(css) && /nth-child\(even\) td:first-child \{ background: #FAFAFA; \}/.test(css) && /tr:hover td:first-child \{ background: #FAF0FF; \}/.test(css));
}

console.log('\nTest 4 — the numbers');
{
  const h = load();
  const f = rows => h.ctx._dbFlagsPerAudit(rows);
  const r = flags => ({ 'Critical Flags': flags });
  const rows = [r('None'), r(''), r(' none '), r(undefined), r('A'), r('A | B'), r('A | B'), r('A | B | C'), r('A | B | C | D'), r('A | B | C | D | E | F')];
  const got = f(rows);
  check('0 flags: "None", blank, spaces and missing all count as no flag (4 audits)', got[0][0] === 'No flags' && got[0][1] === 4, JSON.stringify(got));
  check('1, 2, 3 and 4 or more flags are counted: 1, 2, 1, 2', got[1][1] === 1 && got[2][1] === 2 && got[3][1] === 1 && got[4][1] === 2 && got[4][0] === '4 or more', JSON.stringify(got));
  check('every audit is in exactly one bucket (the buckets add up to the audits)', got.reduce((s, x) => s + x[1], 0) === rows.length);
  check('a list of nothing gives five zeros, not an error', f([]).every(x => x[1] === 0) && f(null).every(x => x[1] === 0));
  // the same counting as the tiles: for audits with fewer than 4 flags the flags in the buckets equal the flags the tiles add up
  const few = [r('A'), r('A | B'), r('A | B | C'), r('X | Y'), r('')];
  const fromBuckets = f(few).slice(0, 4).reduce((s, x, i) => s + i * x[1], 0);
  const fromTiles = few.reduce((s, x) => s + h.ctx._dbSplitFlags(x['Critical Flags']).length, 0);
  check('the flags in the buckets add up to the flags the tiles count (same _dbSplitFlags)', fromBuckets === fromTiles && fromTiles === 8, fromBuckets + ' vs ' + fromTiles);
  check('a comma-separated older flag cell is read the same way as in the tiles', f([r('Wrong LOB, handled a TV call, No save attempt')])[3][1] === 0 && f([r('Wrong LOB, handled a TV call, No save attempt')]).reduce((s, x) => s + x[1], 0) === 1);
}

console.log('\nTest 4b — the charts are drawn again when the window changes size');
{
  const h = load(), ctx = h.ctx, calls = [];
  ctx.dbSwitchTab = (name, btn) => calls.push(name + ':' + (btn && btn.tab));
  h.document.querySelector = sel => { const m = /data-tab="(\w+)"/.exec(sel); return m ? { tab: m[1] } : null; };
  const boxes = ['ov-chart-flags', 'r-chart-drivers', 's-chart-flagdist'].map(id => ({ id, innerHTML: '<svg width=600></svg>' }));
  h.document.querySelectorAll = sel => (sel === '[id*="-chart-"]' ? boxes : []);
  ctx._dbLoaded = true; ctx._currentView = 'repeats'; ctx._tabRenderHash = { r: 'x' };
  ctx.dbRedrawActiveView();
  check('on a dashboard view the memo is cleared and the open tab is drawn again', calls.join() === 'repeats:repeats' && Object.keys(ctx._tabRenderHash).length === 0, calls.join());
  check('the old drawings are removed first (they keep their old width and stop the columns from getting narrower)', boxes.every(b => b.innerHTML === ''));
  check('and the columns of the chart rows may shrink (min-width: 0)', /\.ov-two-col > \*, \.db-chart-row > \* \{ min-width: 0; \}/.test(IX));
  calls.length = 0; ctx._currentView = 'home'; ctx._tabRenderHash = { r: 'x' }; boxes.forEach(b => { b.innerHTML = 'kept'; });
  ctx.dbRedrawActiveView();
  check('on Home or the form nothing is drawn or cleared (the dashboard is hidden and its charts would have no size)', calls.length === 0 && ctx._tabRenderHash.r === 'x' && boxes.every(b => b.innerHTML === 'kept'));
  calls.length = 0; ctx._currentView = 'overview'; ctx._dbLoaded = false;
  ctx.dbRedrawActiveView();
  check('before the data has loaded nothing is drawn', calls.length === 0);
  check('the listener waits until the size has settled (400 ms after the last change) and is registered on the window', /var _resizeTimer = null;[\s\S]*?setTimeout\(dbRedrawActiveView, 400\)/.test(IX) && /window\.addEventListener\('resize', _dbOnResize\)/.test(IX));
}

console.log('\nTest 5 — the doughnut that is drawn');
{
  const h = load();
  const drawn = {}, opts = {};
  function DataTable() { this.rows = []; }
  DataTable.prototype.addColumn = function () {};
  DataTable.prototype.addRow = function (x) { this.rows.push(x); };
  DataTable.prototype.addRows = function (xs) { xs.forEach(x => this.rows.push(x)); };
  DataTable.prototype.getNumberOfRows = function () { return this.rows.length; };
  h.ctx.google.visualization = { DataTable,
    PieChart: function (el) { this.draw = (dt, o) => { drawn[el.id] = dt.rows.slice(); opts[el.id] = o; }; },
    BarChart: function (el) { this.draw = (dt, o) => { drawn['bar:' + el.id] = dt.rows.slice(); }; } };
  const base = i => ({ 'Audit Ref': 'R' + i, 'Timestamp': '2026-10-12T10:00:00', 'Team Member': 'A' + (i % 7), 'Analysis Type': 'Repeats & Transfer Analyzer', 'Critical Flags': '' });
  // 200 audits whose flag texts are all different: one slice per text made "Other" ~96%
  const many = Array.from({ length: 200 }, (_, i) => Object.assign(base(i), { 'Critical Flags': i % 2 ? 'Unique flag ' + i + ' | Another unique ' + i : 'Only unique ' + i }));
  h.ctx._dbRFiltered = many; h.ctx.dbDrawChart('r', many);
  const rows = drawn['r-chart-flagdist'];
  check('200 audits with all-different flags: two slices, 100 audits with 1 flag and 100 with 2 (no "Other")', JSON.stringify(rows) === JSON.stringify([['1 flag', 100], ['2 flags', 100]]), JSON.stringify(rows));
  const o = opts['r-chart-flagdist'];
  check('each slice has its own colour, in the order of the rows (1 flag amber, 2 flags orange), and the legend is 12 px', o.colors.length === 2 && o.colors[0] === '#B7791F' && o.colors[1] === '#E65100' && o.legend.textStyle.fontSize === 12, JSON.stringify(o.colors));
  const withNone = many.concat([base(900), base(901)]);
  h.ctx.dbDrawChart('r', withNone);
  check('audits with no flag get the green slice, listed first', drawn['r-chart-flagdist'][0][0] === 'No flags' && drawn['r-chart-flagdist'][0][1] === 2 && opts['r-chart-flagdist'].colors[0] === '#248700', JSON.stringify(drawn['r-chart-flagdist']));
  h.ctx.dbDrawChart('s', []);
  check('no audits at all draws a single grey "No audits" slice, not an error', JSON.stringify(drawn['s-chart-flagdist']) === JSON.stringify([['No audits', 1]]), JSON.stringify(drawn['s-chart-flagdist']));
  // the Overview does the same with both analyzers together
  h.ctx._dbRFiltered = many.slice(0, 20); h.ctx._dbSFiltered = [Object.assign(base(500), { 'Analysis Type': 'Sales Analyzer', 'Critical Flags': 'A | B | C | D | E' })];
  h.ctx.renderOverviewTab();
  const ov = drawn['ov-chart-rca'];
  check('the Overview doughnut counts Repeats and Sales together (21 audits) and shows the 4-or-more bucket', Array.isArray(ov) && ov.reduce((s, x) => s + x[1], 0) === 21 && ov.some(x => x[0] === '4 or more' && x[1] === 1), JSON.stringify(ov));
}

console.log('\nTest — the title row under the navigation bar has room');
{
  const tb = /#db-topbar \{([^}]*)\}/.exec(IX);
  check('the title row has at least 12 px of padding above and below and a minimum height of 60 px (it sat tight under the bar)',
    !!tb && /padding:\s*(1[2-9]|[2-9]\d)px\s+24px/.test(tb[1]) && /min-height:\s*(6\d|[7-9]\d)px/.test(tb[1]), tb ? tb[1].replace(/\s+/g, ' ') : 'no rule');
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
