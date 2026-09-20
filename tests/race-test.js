/**
 * Race-condition test for the NH Analyzer SAP autofill.
 *
 * Loads the real <script> block out of index.html into a sandbox with a stubbed
 * DOM and a stubbed google.script.run, so the ORDER in which replies arrive can
 * be controlled. That ordering is the whole bug: a lookup for a 4-digit prefix
 * of a 7-digit SAP ID returns a different real agent, and if its reply lands
 * last it overwrites the correct name.
 *
 * Usage:  node tests/race-test.js [path-to-index.html]
 *         (defaults to ../index.html)
 * Exits non-zero if any test fails.
 */
const fs = require('fs');
const vm = require('vm');

const path = require('path');
const htmlPath = process.argv[2] || path.join(__dirname, '..', 'index.html');

const html = fs.readFileSync(htmlPath, 'utf8');
const blocks = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi)].map(m => m[1]);
const script = blocks.sort((a, b) => b.length - a.length)[0];   // the app block

// ── Stub DOM ────────────────────────────────────────────────────────────────
function makeEl(id) {
  return {
    id, value: '', textContent: '', innerHTML: '', style: {}, options: [], selectedIndex: -1,
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

// ── Stub google.script.run ──────────────────────────────────────────────────
// Real GAS returns a fresh runner per chain, so concurrent chains must not
// share handler state — modelling that faithfully matters for this test.
const pending = [];
function makeRunner() {
  const h = {};
  const runner = new Proxy({}, {
    get(_, prop) {
      if (prop === 'withSuccessHandler') return fn => { h.success = fn; return runner; };
      if (prop === 'withFailureHandler') return fn => { h.failure = fn; return runner; };
      return (...args) => { pending.push({ method: String(prop), args, success: h.success }); };
    }
  });
  return runner;
}
const google = { script: { get run() { return makeRunner(); }, host: { close(){} } } };

// ── Other globals the page expects ──────────────────────────────────────────
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
try {
  vm.runInContext(script, ctx, { timeout: 15000 });
} catch (e) {
  // Top-level page-init code may touch DOM shapes we did not stub. The functions
  // under test are hoisted declarations, so they exist regardless.
  console.log('  (page init threw, expected in a stub DOM: ' + String(e.message).slice(0, 80) + ')');
}

// ── Helpers ─────────────────────────────────────────────────────────────────
const sleep = ms => new Promise(r => setTimeout(r, ms));
const AGENT_PREFIX = { participant: 'Prefix Person',  teamLeader: 'TL-A', opsManager: 'OM-A', lineOfBusiness: 'LOB-A', locale: 'Site-A', vtid: 'V-A', agentEmail: 'a@telus.com' };
const AGENT_FULL   = { participant: 'Correct Person', teamLeader: 'TL-B', opsManager: 'OM-B', lineOfBusiness: 'LOB-B', locale: 'Site-B', vtid: 'V-B', agentEmail: 'b@telus.com' };

function reset() {
  for (const k in els) delete els[k];
  pending.length = 0;
  localStorage.clear();
  if ('_sapResolvedParticipant' in ctx) ctx._sapResolvedParticipant = false;
}
function take(method, arg) {
  const i = pending.findIndex(p => p.method === method && String(p.args[0]) === String(arg));
  if (i === -1) return null;
  return pending.splice(i, 1)[0];
}

let failures = 0;
function check(name, actual, expected) {
  const ok = actual === expected;
  if (!ok) failures++;
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}`);
  if (!ok) console.log(`          expected "${expected}", got "${actual}"`);
}

// ── Test 1: the reported bug ────────────────────────────────────────────────
// Analyst types 2007, pauses (lookup fires for another real agent), types 888,
// second lookup fires. The CORRECT reply arrives first, the stale prefix reply
// arrives last. The prefix must not win.
async function testStaleReplyLosesRace() {
  console.log('\nTest 1 — stale prefix reply must not overwrite the correct name');
  reset();

  document.getElementById('sapId').value = '2007';
  ctx.debounceSapLookup();
  await sleep(700);

  document.getElementById('sapId').value = '2007888';
  ctx.debounceSapLookup();
  await sleep(700);

  const reqPrefix = take('getRosterBySapId', '2007');
  const reqFull   = take('getRosterBySapId', '2007888');

  if (!reqPrefix) { console.log('  SKIP  no lookup fired for the 4-digit prefix'); return; }
  if (!reqFull)   { console.log('  FAIL  no lookup fired for the full SAP ID'); failures++; return; }

  reqFull.success(AGENT_FULL);       // correct reply lands first
  reqPrefix.success(AGENT_PREFIX);   // stale reply lands last — must be ignored

  check('Team Member name', document.getElementById('participant').value, 'Correct Person');
  check('Team Leader',      document.getElementById('teamLeader').value,  'TL-B');
  check('Ops Manager',      document.getElementById('opsManager').value,  'OM-B');
  check('VTID',             document.getElementById('vtid').value,        'V-B');
}

// ── Test 2: normal case still works ─────────────────────────────────────────
async function testNormalLookupStillWorks() {
  console.log('\nTest 2 — an ordinary single lookup still fills the fields');
  reset();
  document.getElementById('sapId').value = '2007888';
  ctx.debounceSapLookup();
  await sleep(700);
  const req = take('getRosterBySapId', '2007888');
  if (!req) { console.log('  FAIL  no lookup fired'); failures++; return; }
  req.success(AGENT_FULL);
  check('Team Member name', document.getElementById('participant').value, 'Correct Person');
  check('Locale',           document.getElementById('locale').value,      'Site-B');
}

// ── Test 3: field changed after dispatch ────────────────────────────────────
// The analyst clears and retypes while a reply is in flight; that reply is for
// a SAP ID no longer in the box and must be dropped.
async function testFieldChangedUnderneath() {
  console.log('\nTest 3 — a reply for a SAP ID no longer in the field is dropped');
  reset();
  document.getElementById('sapId').value = '2007888';
  ctx.debounceSapLookup();
  await sleep(700);
  const req = take('getRosterBySapId', '2007888');
  if (!req) { console.log('  FAIL  no lookup fired'); failures++; return; }

  document.getElementById('sapId').value = '3115222';  // analyst retyped
  req.success(AGENT_FULL);                              // old reply arrives

  check('Team Member left empty', document.getElementById('participant').value, '');
}

// ── Test 4: the unvalidated name -> SAP cache must be gone ──────────────────
// It bound a client-side regex guess at the speaker's name to a real SAP ID
// with no server validation, then replayed it for two hours with no network
// call that could disprove it. One misparsed name meant hours of confidently
// wrong team members.
function testNoNameSapCache() {
  console.log('\nTest 4 — no unvalidated name -> SAP ID cache is read or written');
  let bad = 0;
  if (/lsGet\(\s*LS_NAME_SAP_KEY/.test(script)) { console.log('  FAIL  source still READS the name -> SAP cache');  bad++; }
  if (/lsSet\(\s*LS_NAME_SAP_KEY/.test(script)) { console.log('  FAIL  source still WRITES the name -> SAP cache'); bad++; }
  for (const key of Object.keys(store)) {
    if (/name_sap/i.test(key)) { console.log(`  FAIL  runtime wrote cache key "${key}"`); bad++; }
  }
  failures += bad;
  if (bad === 0) console.log('  PASS  name -> SAP cache is neither read nor written');
}

(async () => {
  console.log('NH Analyzer — SAP autofill race tests');
  console.log('file: ' + htmlPath);
  await testStaleReplyLosesRace();
  await testNormalLookupStillWorks();
  await testFieldChangedUnderneath();
  testNoNameSapCache();
  console.log(failures === 0 ? '\nALL TESTS PASSED' : `\n${failures} CHECK(S) FAILED`);
  process.exit(failures === 0 ? 0 : 1);
})();
