/**
 * Dashboard loading and refresh (code review batch 3).
 *
 *  1. A Refresh (or the background reload) put every filter pill back on All, undoing what the person had chosen.
 *  2. A slow older reply could overwrite a newer one.
 *  3. A failed Refresh left the spinner spinning for ever.
 *  4. One failed call out of three was counted twice, so the page could render before the others had answered.
 *  5. Repeats and Sales shared one counter for their AI insights, so starting one cancelled the other for good ("Generating AI insights..." never ended).
 *  6. Every tick of a filter box sent its own AI requests (three per tick); now the last one is sent after the filters are still.
 *
 * Usage: node tests/dashboard-refresh-test.js
 */
const { load } = require('./lib/page-harness');
const { makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const h = load();
h.ctx.alert = () => {};

const mk = (i, name, type) => ({ 'Audit Ref': 'NHA-' + i, 'Timestamp': '2026-10-13T10:00:00', 'Team Member': name, 'Team Leader': 'TL', 'Line of Business': 'MOB', 'Locale': 'PH',
  'Analysis Type': type || 'Repeats & Transfer Analyzer', 'SAP ID': '1' + i });
const ROWS = [mk(1, 'Maria Santos'), mk(2, 'Aisha Khan'), mk(3, 'Carlos Mendoza')];
const callsOf = m => h.pending.filter(p => p.method === m);

// The harness has no real checkboxes, so the pill stores are small fakes that behave like the page's: dbPopulateFilters rebuilds them from the
// data (everything checked is what "All" looks like), and querySelectorAll answers for the two selectors the page uses.
const FIELD = { member: 'Team Member', leader: 'Team Leader', ops: 'Operations Manager', lob: 'Line of Business', locale: 'Locale', vtid: 'VTID', atype: 'Analysis Type', ref: 'Audit Ref', qa: 'Observer Name' };
function fakeStores() {
  Object.keys(FIELD).forEach(key => {
    const store = h.document.getElementById('sh-ms-' + key);
    const values = [...new Set((h.ctx._dbAllData || []).map(r => r[FIELD[key]]).filter(Boolean))];
    store.boxes = values.map(v => ({ value: v, checked: true }));
    store.querySelectorAll = sel => (/:checked/.test(sel) ? store.boxes.filter(b => b.checked) : store.boxes);
  });
}
h.ctx.dbPopulateFilters = fakeStores;

(async () => {
  console.log('\nTest 1 — a refresh keeps the filters the person chose');
  {
    h.ctx._dbAllData = ROWS; h.ctx._dbAuditLog = ROWS; h.ctx._dbQAUsers = []; h.ctx._dbRoster = [];
    h.ctx._dbLoaded = false;
    h.ctx.dbFinishRender();                                   // the first load: everything is on All
    check('first load: nothing is narrowed', h.ctx.dbGetSharedFilters().members.length === 0);
    h.ctx.dbPillOnly('member', 'Aisha Khan');
    check('the person narrows to one agent', JSON.stringify(h.ctx.dbGetSharedFilters().members) === JSON.stringify(['Aisha Khan']), JSON.stringify(h.ctx.dbGetSharedFilters().members));
    h.ctx.dbFinishRender();                                   // what a Refresh or the background reload does
    check('after the refresh the filter is still Aisha Khan', JSON.stringify(h.ctx.dbGetSharedFilters().members) === JSON.stringify(['Aisha Khan']), JSON.stringify(h.ctx.dbGetSharedFilters().members));
    check('and the rows shown are still only hers', h.ctx._dbRFiltered.length === 1 && h.ctx._dbRFiltered[0]['Team Member'] === 'Aisha Khan', JSON.stringify(h.ctx._dbRFiltered.map(r => r['Team Member'])));
    // the agent disappears from the new data: the pill goes back to All instead of showing nothing
    h.ctx._dbAllData = ROWS.filter(r => r['Team Member'] !== 'Aisha Khan'); h.ctx._dbAuditLog = h.ctx._dbAllData;
    h.ctx.dbFinishRender();
    check('if the chosen agent is no longer in the data the pill goes back to All', h.ctx.dbGetSharedFilters().members.length === 0 && h.ctx._dbRFiltered.length === 2, JSON.stringify(h.ctx.dbGetSharedFilters().members));
    h.ctx._dbAllData = ROWS; h.ctx._dbAuditLog = ROWS;
    h.ctx.dbFinishRender();
    h.ctx.dbPillOnly('member', 'Maria Santos'); h.ctx.dbPillOnly('member', 'Maria Santos');
    h.ctx._dbAllData = ROWS.concat([mk(4, 'New Person')]); h.ctx._dbAuditLog = h.ctx._dbAllData;
    h.ctx.dbFinishRender();
    check('a new agent in the data does not break a narrowed filter', JSON.stringify(h.ctx.dbGetSharedFilters().members) === JSON.stringify(['Maria Santos']));
  }

  console.log('\nTest 2 — a slow older reply does not overwrite a newer one');
  {
    h.reset();
    h.ctx._dbRoster = [{ 'SAP ID': '1' }];
    h.ctx.dbHandleData({ rows: [], ts: 1 }, false);           // call 1 (slow)
    h.ctx.dbHandleData({ rows: [], ts: 2 }, false);           // call 2 (newer)
    const reqs = callsOf('getAuditLogData');
    check('two loads are waiting for the audit log', reqs.length === 2, String(reqs.length));
    const A = [mk(7, 'Old Reply')], B = [mk(8, 'New Reply')];
    reqs[1].success(B);                                        // the newer answers first
    reqs[0].success(A);                                        // the older arrives late
    check('the newer data is kept', h.ctx._dbAllData.length === 1 && h.ctx._dbAllData[0]['Team Member'] === 'New Reply', JSON.stringify(h.ctx._dbAllData.map(r => r['Team Member'])));
    h.ctx.dbHandleData({ rows: [], ts: 3 }, false);
    const r3 = callsOf('getAuditLogData').pop();
    r3.success([mk(9, 'Third')]);
    check('a single load still works', h.ctx._dbAllData[0]['Team Member'] === 'Third');
    h.ctx.dbHandleData({ rows: [], ts: 4 }, false); h.ctx.dbHandleData({ rows: [], ts: 5 }, false);
    const rs = callsOf('getAuditLogData').slice(-2);
    rs[0].failure(new Error('old failed'));
    check('an older FAILURE after a newer load started is ignored too', h.ctx._dbAllData[0]['Team Member'] === 'Third', JSON.stringify(h.ctx._dbAllData.map(r => r['Team Member'])));
  }

  console.log('\nTest 3 — a failed Refresh stops the spinner');
  {
    h.reset();
    h.ctx.dbForceRefresh();
    const load = h.els['db-loading'];
    check('the spinner is shown while refreshing', load && load.style.display === 'block');
    callsOf('getDashboardDataFresh')[0].failure(new Error('boom'));
    check('after a failure the spinner is hidden', h.els['db-loading'].style.display === 'none', h.els['db-loading'].style.display);
    check('and the button works again', h.els['db-refresh-btn'].disabled === false && /Refresh$/.test(h.els['db-refresh-btn'].textContent), h.els['db-refresh-btn'].textContent);
  }

  console.log('\nTest 4 — one failed call is counted once');
  {
    h.reset();
    let done = 0;
    h.ctx._fetchAllParallel(() => { done++; });
    const dash = callsOf('getDashboardData')[0], roster = callsOf('getRosterAll')[0], audit = callsOf('getAuditLogData')[0];
    roster.failure(new Error('roster down'));
    dash.success({ rows: [], ts: 1 });
    check('after 2 of the 3 calls have answered the page has NOT finished loading yet', done === 0, String(done));
    audit.success([mk(1, 'Maria Santos')]);
    check('after the third it finishes, exactly once', done === 1, String(done));
  }

  console.log('\nTest 5 — Repeats and Sales insights do not cancel each other');
  {
    h.reset();
    ['r-highlights', 'r-lowlights', 's-highlights', 's-lowlights'].forEach(id => h.document.getElementById(id));
    h.ctx.renderTabAIHighlights('r', ROWS);
    h.ctx.renderTabAIHighlights('s', [mk(5, 'Sales Person', 'Sales Analyzer')]);
    await h.sleep(1000);
    const calls = callsOf('getOverviewHighlightsLowlights');
    check('both tabs sent their request', calls.length === 2, String(calls.length));
    calls[0].success({ success: true, highlights: ['Good greeting'], lowlights: ['Long holds'] });
    calls[1].success({ success: true, highlights: ['Offer made'], lowlights: [] });
    check('the Repeats tab got its answer (it used to stay on "Generating AI insights..." for ever)', /Good greeting/.test(h.els['r-highlights'].innerHTML) && /Long holds/.test(h.els['r-lowlights'].innerHTML), h.els['r-highlights'].innerHTML);
    check('and so did Sales', /Offer made/.test(h.els['s-highlights'].innerHTML));
  }

  console.log('\nTest 6 — ticking filter boxes sends one AI request, after the filters are still');
  {
    h.reset();
    ['r-highlights', 'r-lowlights', 'ov-highlights', 'ov-lowlights'].forEach(id => h.document.getElementById(id));
    for (let i = 0; i < 6; i++) { h.ctx.renderTabAIHighlights('r', ROWS.slice(0, 1 + (i % 3))); h.ctx._dbRFiltered = ROWS; h.ctx._dbSFiltered = []; h.ctx.renderOverviewAIHighlights(); }
    check('right away nothing has been sent (6 renders)', callsOf('getOverviewHighlightsLowlights').length === 0, String(callsOf('getOverviewHighlightsLowlights').length));
    check('the person sees that insights are coming', /Generating AI insights/.test(h.els['r-highlights'].innerHTML));
    await h.sleep(1000);
    const calls = callsOf('getOverviewHighlightsLowlights');
    check('after the filters are still, ONE request per area is sent (Repeats tab and Overview), not 12', calls.length === 2, String(calls.length));
    const tabCall = calls.find(c => c.args[1] && c.args[1].analysisType === 'repeats');
    check('and the Repeats request carries the LAST selection (3 rows... index 5 -> 3 rows)', tabCall && tabCall.args[0].length === 3, tabCall && String(tabCall.args[0].length));
    // a filter box ticked, then ticked again 500 ms later (before the first wait is over): still nothing sent; one request after the last tick
    h.reset();
    h.ctx.renderTabAIHighlights('r', ROWS);
    await h.sleep(500);
    h.ctx.renderTabAIHighlights('r', ROWS.slice(0, 2));
    await h.sleep(500);
    check('1000 ms after the first tick but only 500 ms after the second: nothing sent yet (the wait restarts at each tick)', callsOf('getOverviewHighlightsLowlights').length === 0, String(callsOf('getOverviewHighlightsLowlights').length));
    await h.sleep(600);
    const late = callsOf('getOverviewHighlightsLowlights');
    check('then one request, for the last selection', late.length === 1 && late[0].args[0].length === 2, String(late.length));
    h.reset();
    h.ctx.renderTabAIHighlights('r', ROWS);
    h.ctx.renderTabAIHighlights('r', []);
    await h.sleep(1000);
    check('if the filters end with no records, the waiting request is cancelled', callsOf('getOverviewHighlightsLowlights').length === 0);
  }

  console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
  process.exit(check.state.failures ? 1 : 0);
})();
