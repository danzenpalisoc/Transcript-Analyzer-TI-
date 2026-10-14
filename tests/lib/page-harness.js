/**
 * Loads the real <script> block out of index.html into a sandbox with a stubbed
 * DOM and a stubbed google.script.run, so the ORDER in which server replies
 * arrive can be controlled by the test.
 *
 * Real GAS returns a fresh runner per chain, so concurrent chains must not share
 * handler state; the stub models that. Each call is recorded in `pending` with
 * the success/failure handlers that were attached to it.
 */
const fs = require('fs');
const vm = require('vm');
const path = require('path');

// opts.preset: { key: value } put into localStorage BEFORE the page script runs (for example a saved draft).
function load(htmlPath, opts) {
  const html = fs.readFileSync(htmlPath || path.join(__dirname, '..', '..', 'index.html'), 'utf8');
  const blocks = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi)].map(m => m[1]);
  const script = blocks.sort((a, b) => b.length - a.length)[0];

  function makeEl(id) {
    return {
      id, value: '', textContent: '', innerHTML: '', className: '', style: {}, hidden: false,
      disabled: false, checked: false, options: [], selectedIndex: -1,
      classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
      addEventListener(){}, removeEventListener(){}, setAttribute(){}, removeAttribute(){},
      getAttribute(){ return null; }, appendChild(){}, removeChild(){}, remove(){},
      insertAdjacentElement(){}, querySelector(){ return null; }, querySelectorAll(){ return []; },
      focus(){}, click(){},
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
  Object.keys((opts && opts.preset) || {}).forEach(k => { store[k] = String(opts.preset[k]); });
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
  function reset() {
    for (const k in els) delete els[k];
    pending.length = 0;
    localStorage.clear();
    if ('_sapResolvedParticipant' in ctx) ctx._sapResolvedParticipant = false;
  }
  const field = id => document.getElementById(id).value;
  const setField = (id, v) => { document.getElementById(id).value = v; };
  // Type into the SAP field and wait out the debounce.
  async function typeSap(sap) {
    setField('sapId', sap);
    ctx.debounceSapLookup();
    await sleep(700);
  }
  // Take the oldest pending call, optionally for one method.
  function take(method) {
    const i = method ? pending.findIndex(p => p.method === method) : 0;
    return i === -1 ? null : pending.splice(i, 1)[0];
  }

  return { html, script, ctx, els, document, pending, localStorage, store, sleep, reset, field, setField, typeSap, take };
}

function makeChecker() {
  const s = { failures: 0 };
  function check(name, ok, detail) {
    if (!ok) s.failures++;
    console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}`);
    if (!ok && detail !== undefined) console.log('          got: ' + detail);
  }
  check.state = s;
  return check;
}

module.exports = { load, makeChecker };
