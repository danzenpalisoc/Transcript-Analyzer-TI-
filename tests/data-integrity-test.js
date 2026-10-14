/**
 * Data that must not be corrupted or lost on its way to a sheet or a cache.
 *
 * 1. Cache chunks. A report over 48k characters is stored across several rows. The
 *    chunk boundary can fall anywhere, and Google Sheets reads what appendRow writes
 *    as if typed: a chunk that starts with = + - @ becomes a formula (an error
 *    value), and a chunk that starts with an apostrophe loses it. Reassembling then
 *    returned a damaged evaluation, and it was served to the agent.
 * 2. Replacing an agent's cached rows reads the sheet, deletes rows by index, then
 *    appends. Without a lock, two saves at once delete the wrong rows.
 * 3. The dashboard cache. If writing the cache fails (the 100 KB limit is in bytes,
 *    and accented or Filipino text is more than one byte per character), the
 *    function returned an EMPTY dashboard although the sheet had been read fine.
 *
 * Usage: node tests/data-integrity-test.js
 */
const vm = require('vm');
const { read, extractFunction, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const codeSrc = read('Code.gs');

function makeEnv(opts) {
  opts = opts || {};
  const store = {};
  const lock = { waits: 0, releases: 0, held: false, heldWhileWriting: false };
  const cs = {
    get: k => (k in store ? store[k] : null),
    put: (k, v) => {
      if (opts.putFails) throw new Error('Service error: CacheService');
      if (opts.maxBytes && Buffer.byteLength(String(v), 'utf8') > opts.maxBytes) throw new Error('Argument too large: value');
      store[k] = v;
    },
    remove: k => { delete store[k]; },
    removeAll: ks => ks.forEach(k => { delete store[k]; }),
  };
  const logs = [];
  const sandbox = {
    Logger: { log: m => logs.push(String(m)) },
    CacheService: { getScriptCache: () => cs },
    LockService: { getScriptLock: () => ({
      waitLock() { lock.waits++; lock.held = true; if (opts.lockFails) { lock.held = false; throw new Error('Could not obtain lock'); } },
      releaseLock() { lock.releases++; lock.held = false; },
    }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Utilities: { base64Encode: s => Buffer.from(String(s)).toString('base64'), sleep() {} },
    ScriptApp: { getScriptId: () => 'x' },
    sharedCSS: () => '', fixBadgeClasses: h => h,
  };
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('Config.gs'), ctx);
  vm.runInContext(read('HtmlSanitizer.gs'), ctx);
  vm.runInContext(read('SheetService.gs'), ctx);
  const sheets = {};
  sheets[ctx.AUDIT_LOG_SHEET] = new FakeSheet(ctx.AUDIT_LOG_HEADERS, []);
  sheets[ctx.CACHE_SHEET]     = new FakeSheet(ctx.CACHE_HEADERS, []);
  sheets[ctx.DASHBOARD_DATA_SHEET] = new FakeSheet(ctx.DASHBOARD_HEADERS, []);
  // note whether the lock is held at the moment rows are written or deleted
  const cache = sheets[ctx.CACHE_SHEET];
  const origAppend = cache.appendRow.bind(cache), origDelete = cache.deleteRow.bind(cache);
  cache.appendRow = a => { if (lock.held) lock.heldWhileWriting = true; return origAppend(a); };
  cache.deleteRow = n => { if (lock.held) lock.heldWhileWriting = true; return origDelete(n); };
  const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new FakeSheet([], [])), getUrl: () => 'u' };
  ctx.getOrCreateSpreadsheet = () => ss;
  return { ctx, sheets, store, lock, logs, cs };
}

const INT = 'int-1', SAP = '373106';

console.log('\nTest 1 — a report whose chunk boundary starts with = + - @ or an apostrophe comes back intact');
{
  const env = makeEnv();
  const CH = env.ctx._CACHE_CHUNK_SIZE;
  // 'x' is the control: an ordinary character at the boundary must of course be fine
  for (const lead of ['=', '+', '-', '@', "'", '"', 'x']) {
    const e = makeEnv();
    // > 99000 characters so only the sheet holds it (the 6-hour script cache skips it)
    let html = '<div>' + 'a'.repeat(CH - 5) + '</div>';          // exactly fills chunk 1
    html = html.slice(0, CH) + lead + 'b'.repeat(CH) + lead + 'c'.repeat(CH) + '</div>';
    if (html.length <= 99000) html += 'z'.repeat(99001 - html.length);
    e.ctx.saveCachedResult_(INT, 'repeats', html, SAP);
    const rows = e.sheets[e.ctx.CACHE_SHEET].grid.slice(1);
    check('"' + lead + '": stored in ' + rows.length + ' chunks', rows.length >= 3);
    const back = e.ctx.findCachedResult_(INT, 'repeats', SAP);
    check('"' + lead + '": reassembled report is identical', back === html,
          back === null ? 'got null' : 'differs at ' + (function () { for (let i = 0; i < html.length; i++) if (back[i] !== html[i]) return i; return back.length; })());
    check('"' + lead + '": no chunk was turned into an error value', !rows.some(r => String(r[3]) === '#ERROR!'));
  }
}

console.log('\nTest 2 — an Interaction ID or SAP ID that looks like a formula is not executed in the Cache sheet');
{
  const e = makeEnv();
  e.ctx.saveCachedResult_('=HYPERLINK("http://evil","x")', 'repeats', '<div>ok</div>', '@SUM(1)');
  const row = e.sheets[e.ctx.CACHE_SHEET].grid[1] || [];
  check('Interaction ID cell is not a formula', String(row[0]) !== '#ERROR!' && /HYPERLINK/.test(String(row[0])), JSON.stringify(row[0]));
  check('SAP ID cell is not a formula', String(row[4]) !== '#ERROR!' && /SUM/.test(String(row[4])), JSON.stringify(row[4]));
}

console.log('\nTest 3 — replacing an agent\'s cached rows happens under the script lock, and the lock is always released');
{
  const e = makeEnv();
  e.ctx.saveCachedResult_(INT, 'repeats', '<div>first</div>', SAP);
  e.ctx.saveCachedResult_(INT, 'repeats', '<div>second</div>', SAP);
  check('the lock was taken for each save', e.lock.waits === 2, 'waits=' + e.lock.waits);
  check('and released each time', e.lock.releases === e.lock.waits, 'releases=' + e.lock.releases);
  check('rows were written / deleted WHILE the lock was held', e.lock.heldWhileWriting === true);
  check('the latest save wins', e.ctx.findCachedResult_(INT, 'repeats', SAP) === '<div>second</div>');

  const f = makeEnv();
  f.ctx.saveCachedResult_(INT, 'repeats', '<div>first</div>', SAP);
  f.sheets[f.ctx.CACHE_SHEET].deleteRow = () => { throw new Error('sheet exploded'); };
  f.ctx.saveCachedResult_(INT, 'repeats', '<div>second</div>', SAP);     // swallows its own error
  check('the lock is released even when the write throws', f.lock.releases === f.lock.waits && !f.lock.held, JSON.stringify(f.lock));

  const g = makeEnv({ lockFails: true });
  g.ctx.saveCachedResult_(INT, 'repeats', '<div>x</div>', SAP);
  check('if the lock cannot be had, the evaluation is STILL saved (availability over perfect ordering)',
        g.ctx.findCachedResult_(INT, 'repeats', SAP) === '<div>x</div>');
  check('...and that is logged', g.logs.some(l => /lock/i.test(l)), g.logs.join(' | '));
}

console.log('\nTest 4 — a failing dashboard-cache write never blanks the dashboard');
{
  // multi-byte text: each "é" / "ñ" / "ñ" is 2-3 bytes, so a 90,000-character chunk is far over 100 KB
  const rows = [];
  for (let i = 0; i < 600; i++) rows.push({ 'Audit Ref': 'NHA-' + i, 'Call Summary': 'Ñandú niño canción ' + 'é'.repeat(300) + ' 日本語'.repeat(60) });
  const rawBytes = Buffer.byteLength(JSON.stringify(rows), 'utf8');

  const env = makeEnv({ maxBytes: 100000 });
  const ctx = env.ctx;
  ['writeDashboardCache_', 'readDashboardCache_', 'invalidateDashboardCache_', 'getDashboardData', 'getDashboardDataFresh', 'readDashboardSheet_']
    .forEach(n => vm.runInContext(extractFunction(codeSrc, n), ctx));
  vm.runInContext(codeSrc.match(/^var DB_CACHE_KEY[^\n]*/m)[0] + '\n' + codeSrc.match(/^var DB_CACHE_META_KEY[^\n]*/m)[0] + '\n' +
                  codeSrc.match(/^var DB_CACHE_TTL[^\n]*/m)[0] + '\n' + codeSrc.match(/^var DB_CACHE_CHUNK[^\n]*/m)[0], ctx);
  ctx.readDashboardSheet_ = () => rows;                   // the sheet read succeeds
  check('the fixture really exceeds one cache value (so the old code would throw)', rawBytes > 100000, 'bytes=' + rawBytes);

  const r1 = ctx.getDashboardData();
  check('getDashboardData returns the rows it read', r1.rows.length === rows.length, 'rows=' + r1.rows.length);
  check('...with the right count', r1.rowCount === rows.length);

  const r2 = ctx.getDashboardDataFresh();
  check('getDashboardDataFresh returns the rows too', r2.rows.length === rows.length, 'rows=' + r2.rows.length);

  const again = ctx.getDashboardData();
  check('a second call is served correctly (from cache or the sheet), never blank', again.rows.length === rows.length);
  check('a cached copy, if one was kept, round-trips exactly',
        !again.fromCache || JSON.stringify(again.rows) === JSON.stringify(rows));

  const small = [{ 'Audit Ref': 'A', 'Call Summary': 'plain' }];
  const e2 = makeEnv({ maxBytes: 100000 });
  ['writeDashboardCache_', 'readDashboardCache_', 'invalidateDashboardCache_', 'getDashboardData']
    .forEach(n => vm.runInContext(extractFunction(codeSrc, n), e2.ctx));
  vm.runInContext(codeSrc.match(/^var DB_CACHE_KEY[^\n]*/m)[0] + '\n' + codeSrc.match(/^var DB_CACHE_META_KEY[^\n]*/m)[0] + '\n' +
                  codeSrc.match(/^var DB_CACHE_TTL[^\n]*/m)[0] + '\n' + codeSrc.match(/^var DB_CACHE_CHUNK[^\n]*/m)[0], e2.ctx);
  e2.ctx.readDashboardSheet_ = () => small;
  e2.ctx.getDashboardData();
  const hit = e2.ctx.getDashboardData();
  check('a small dashboard is still cached and served from the cache', hit.fromCache === true && hit.rows.length === 1, JSON.stringify(hit).slice(0, 80));
}

console.log('\nTest 5 — a partly written dashboard cache is never read as if complete');
{
  const env = makeEnv({ maxBytes: 100000 });
  const ctx = env.ctx;
  ['writeDashboardCache_', 'readDashboardCache_'].forEach(n => vm.runInContext(extractFunction(codeSrc, n), ctx));
  vm.runInContext(codeSrc.match(/^var DB_CACHE_KEY[^\n]*/m)[0] + '\n' + codeSrc.match(/^var DB_CACHE_META_KEY[^\n]*/m)[0] + '\n' +
                  codeSrc.match(/^var DB_CACHE_TTL[^\n]*/m)[0] + '\n' + codeSrc.match(/^var DB_CACHE_CHUNK[^\n]*/m)[0], ctx);
  const rows = [];
  for (let i = 0; i < 400; i++) rows.push({ a: 'ñ'.repeat(400) + i });
  let threw = false;
  try { ctx.writeDashboardCache_(rows); } catch (e) { threw = true; }
  check('writeDashboardCache_ does not throw', !threw);
  const back = ctx.readDashboardCache_();
  check('what is read back is either a miss or the complete data', back === null || JSON.stringify(back.rows) === JSON.stringify(rows));
}

console.log('\nTest 6 — if the cache service itself is down, the dashboard is still served from the sheet');
{
  const env = makeEnv({ putFails: true });
  const ctx = env.ctx;
  ['writeDashboardCache_', 'readDashboardCache_', 'invalidateDashboardCache_', 'getDashboardData', 'getDashboardDataFresh']
    .forEach(n => vm.runInContext(extractFunction(codeSrc, n), ctx));
  vm.runInContext(codeSrc.match(/^var DB_CACHE_KEY[^\n]*/m)[0] + '\n' + codeSrc.match(/^var DB_CACHE_META_KEY[^\n]*/m)[0] + '\n' +
                  codeSrc.match(/^var DB_CACHE_TTL[^\n]*/m)[0] + '\n' + codeSrc.match(/^var DB_CACHE_CHUNK[^\n]*/m)[0], ctx);
  const rows = [{ 'Audit Ref': 'A1' }, { 'Audit Ref': 'A2' }];
  ctx.readDashboardSheet_ = () => rows;
  check('getDashboardData returns the rows although every cache write fails', ctx.getDashboardData().rows.length === 2);
  check('getDashboardDataFresh does too', ctx.getDashboardDataFresh().rows.length === 2);
  check('the failure is logged, not silent', env.logs.some(l => /cache write failed/i.test(l)), env.logs.join(' | '));
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
