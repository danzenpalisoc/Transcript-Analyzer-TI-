/**
 * The SAP ID field must never fail silently.
 *
 * Reported: an analyst types a valid SAP ID, nothing auto-fills, no error, no
 * hint -- across hard refreshes and a fresh deploy. Every failure path in the
 * lookup (null reply, source error, failed call) ended in a bare `return`, so a
 * server that could not read the roster looked identical to a typo.
 *
 * Also: a reply that carries no participant must never be cached. The cache is
 * read BEFORE the server, so one empty row would be replayed for two hours and
 * survive hard refreshes and redeploys (it lives in localStorage).
 *
 * Usage: node tests/sap-lookup-feedback-test.js [path-to-index.html]
 */
const fs = require('fs');
const vm = require('vm');
const path = require('path');
const htmlPath = process.argv[2] || path.join(__dirname, '..', 'index.html');
const html = fs.readFileSync(htmlPath, 'utf8');
const blocks = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi)].map(m => m[1]);
const script = blocks.sort((a, b) => b.length - a.length)[0];

function makeEl(id) {
  return {
    id, value: '', textContent: '', innerHTML: '', className: '', style: {}, hidden: false, options: [], selectedIndex: -1,
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    addEventListener(){}, removeEventListener(){}, setAttribute(){}, removeAttribute(){},
    getAttribute(){ return null; }, appendChild(){}, removeChild(){}, remove(){},
    querySelector(){ return null; }, querySelectorAll(){ return []; }, focus(){}, click(){},
  };
}
const els = {};
const document = {
  getElementById(id) { return els[id] || (els[id] = makeEl(id)); },
  querySelector(){ return null; }, querySelectorAll(){ return []; },
  createElement(){ return makeEl('created'); },
  addEventListener(){}, body: makeEl('body'), documentElement: makeEl('html'),
};
const pending = [];
function makeRunner() {
  const h = {};
  const runner = new Proxy({}, {
    get(_, prop) {
      if (prop === 'withSuccessHandler') return fn => { h.success = fn; return runner; };
      if (prop === 'withFailureHandler') return fn => { h.failure = fn; return runner; };
      return (...args) => { pending.push({ method: String(prop), args, success: h.success, failure: h.failure }); };
    }
  });
  return runner;
}
const google = { script: { get run() { return makeRunner(); }, host: { close(){} } } };
const store = {};
const localStorage = {
  getItem: k => (k in store ? store[k] : null),
  setItem: (k, v) => { store[k] = String(v); },
  removeItem: k => { delete store[k]; },
  clear: () => { for (const k in store) delete store[k]; },
};
const sandbox = {
  document, google, localStorage, sessionStorage: localStorage,
  window: { addEventListener(){}, location: { href: '' }, matchMedia: () => ({ matches: false, addEventListener(){} }) },
  navigator: { userAgent: 'node', clipboard: { writeText: async () => {} } },
  setTimeout, clearTimeout, setInterval, clearInterval, console,
  alert(){}, confirm: () => true, Date, Math, JSON, RegExp, parseInt, parseFloat, isNaN,
  encodeURIComponent, decodeURIComponent, Object, Array, String, Number, Boolean, Error, Promise,
};
sandbox.globalThis = sandbox;
sandbox.window.document = document;
const ctx = vm.createContext(sandbox);
try { vm.runInContext(script, ctx, { timeout: 15000 }); } catch (e) { /* page init in a stub DOM */ }

const sleep = ms => new Promise(r => setTimeout(r, ms));
const AGENT = { participant: 'Real Person', teamLeader: 'TL', opsManager: 'OM', lineOfBusiness: 'LOB', locale: 'Site', vtid: 'V1', agentEmail: 'r@telus.com' };
const status = () => document.getElementById('sapStatus').textContent;
const field = id => document.getElementById(id).value;
function reset() {
  for (const k in els) delete els[k];
  pending.length = 0; localStorage.clear();
  if ('_sapResolvedParticipant' in ctx) ctx._sapResolvedParticipant = false;
}
async function type(sap) {
  document.getElementById('sapId').value = sap;
  ctx.debounceSapLookup();
  await sleep(700);
}
const take = () => pending.shift() || null;
let failures = 0;
function check(name, ok, detail) {
  if (!ok) failures++;
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}`);
  if (!ok && detail !== undefined) console.log('          got: ' + detail);
}
const cached = () => localStorage.getItem('nh_sap_cache_v1');

async function main() {
  console.log('\nTest 1 — a null reply says the agent was not found');
  reset(); await type('243141');
  let req = take();
  check('status shows "Looking up" while waiting', /looking up/i.test(status()), status());
  req.success(null);
  check('status says no agent found', /no agent found/i.test(status()), status());
  check('status names the SAP ID', /243141/.test(status()), status());
  check('nothing cached', cached() === null, cached());

  console.log('\nTest 2 — an unreadable source is reported, and is not "not found"');
  reset(); await type('243141'); req = take();
  req.success({ _error: 'source_unavailable', sources: ['Roster'] });
  check('says it could not be read', /could not be read/i.test(status()), status());
  check('names the source', /Roster/.test(status()), status());
  check('does NOT claim the agent is missing', !/no agent found/i.test(status()), status());
  check('Team Member not polluted by the error object', field('participant') === '', field('participant'));
  check('nothing cached', cached() === null, cached());

  console.log('\nTest 3 — a lookup_failed reply asks the analyst to retry');
  reset(); await type('243141'); req = take();
  req.success({ _error: 'lookup_failed' });
  check('mentions trying again', /try again/i.test(status()), status());
  check('nothing cached', cached() === null, cached());

  console.log('\nTest 4 — a failed google.script.run call is reported');
  reset(); await type('243141'); req = take();
  req.failure(new Error('server hiccup'));
  check('says lookup failed', /failed/i.test(status()), status());

  console.log('\nTest 5 — success clears the status and fills the form');
  reset(); await type('243141'); req = take();
  req.success(AGENT);
  check('status cleared', status() === '', status());
  check('Team Member filled', field('participant') === 'Real Person', field('participant'));
  check('cached', cached() !== null);

  console.log('\nTest 6 — a reply with no participant is never cached');
  reset(); await type('243141'); req = take();
  req.success({ participant: '', teamLeader: '', opsManager: '', lineOfBusiness: '', locale: '', vtid: '', agentEmail: '' });
  check('not cached', cached() === null, cached());

  console.log('\nTest 7 — an empty row already in the cache is ignored and the server is asked');
  reset();
  localStorage.setItem('nh_sap_cache_v1', JSON.stringify({ savedAt: Date.now(), data: { '243141': { participant: '' } } }));
  await type('243141');
  check('server was called despite the cached entry', pending.length === 1, 'pending=' + pending.length);
  req = take();
  if (req) { req.success(AGENT); check('form filled from the server reply', field('participant') === 'Real Person', field('participant')); }

  console.log('\nTest 8 — a good cached row still short-circuits the server');
  reset();
  localStorage.setItem('nh_sap_cache_v1', JSON.stringify({ savedAt: Date.now(), data: { '243141': AGENT } }));
  await type('243141');
  check('no server call', pending.length === 0, 'pending=' + pending.length);
  check('form filled from cache', field('participant') === 'Real Person', field('participant'));

  console.log('\nTest 9 — a stale reply must not touch the status');
  reset(); await type('2007'); const stale = take();
  await type('2007888'); const fresh = take();
  fresh.success(AGENT);
  stale.success(null);
  check('status not overwritten by the stale not-found', status() === '', status());
  check('name stays correct', field('participant') === 'Real Person', field('participant'));

  console.log('\nTest 10 — clearing the field clears the status');
  reset(); await type('243141'); take().success(null);
  await type('');
  check('status cleared', status() === '', status());

  console.log('\nTest 11 — the status line exists in the page and is announced to screen readers');
  check('element #sapStatus in markup', /id="sapStatus"/.test(html));
  check('aria-live set', /id="sapStatus"[^>]*aria-live="polite"/.test(html) || /aria-live="polite"[^>]*id="sapStatus"/.test(html));

  console.log(failures ? '\nFAILED' : '\nALL TESTS PASSED');
  process.exit(failures ? 1 : 0);
}
main();
