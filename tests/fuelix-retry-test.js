/**
 * Regression test for callFuelIX() (AIService.gs) error tagging + retry.
 *
 * Reported symptom: "A lookup source could not be opened... the file is
 * shared with a different account" shown to analysts, audit not saved.
 * Root cause: submitTranscript() (Code.gs) calls analyzeTranscript() ->
 * callFuelIX() completely unguarded. A FuelIX auth/outage error's response
 * body often contains wording like "permission" or "access", which
 * submitTranscript's generic substring check misclassified as a Sheets/Drive
 * lookup failure — sending everyone hunting through Roster sharing for a bug
 * that was actually the AI call.
 *
 * Fix: every error callFuelIX() throws is tagged 'AI_SERVICE_ERROR' so the
 * caller can classify it correctly, and transient-looking failures (429,
 * 5xx, network-level) are retried — a bad/expired key (401/403) or malformed
 * request (400) is not, since retrying those just delays an identical result.
 *
 * Usage: node tests/fuelix-retry-test.js [path-to-AIService.gs]
 * Exits non-zero if any check fails.
 */
const fs = require('fs');
const vm = require('vm');
const path = require('path');

const gsPath = process.argv[2] || path.join(__dirname, '..', 'AIService.gs');
const src = fs.readFileSync(gsPath, 'utf8');

function extractFunction(source, name) {
  const startMatch = source.match(new RegExp('function\\s+' + name + '\\s*\\('));
  if (!startMatch) throw new Error('function ' + name + ' not found in ' + gsPath);
  let i = startMatch.index;
  const braceStart = source.indexOf('{', i);
  let depth = 0, end = -1;
  for (let j = braceStart; j < source.length; j++) {
    if (source[j] === '{') depth++;
    else if (source[j] === '}') { depth--; if (depth === 0) { end = j + 1; break; } }
  }
  if (end === -1) throw new Error('could not find end of function ' + name);
  return source.slice(i, end);
}

const fnSource = extractFunction(src, 'callFuelIX');

function makeResponse(code, bodyObj) {
  const text = typeof bodyObj === 'string' ? bodyObj : JSON.stringify(bodyObj);
  return { getResponseCode: () => code, getContentText: () => text };
}

function run(fetchImpl) {
  const sleeps = [];
  const logs = [];
  const sandbox = {
    FUELIX_CONFIG: { baseUrl: 'https://api.fuelix.ai', model: 'claude-sonnet-4', apiKey: 'test-key' },
    UrlFetchApp: { fetch: fetchImpl },
    Utilities: { sleep(ms) { sleeps.push(ms); } },
    Logger: { log(msg) { logs.push(String(msg)); } },
    JSON,
  };
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(fnSource, ctx);
  return { callFuelIX: ctx.callFuelIX, sleeps, logs };
}

let failures = 0;
function check(name, cond) {
  console.log('  ' + (cond ? 'PASS' : 'FAIL') + '  ' + name);
  if (!cond) failures++;
}

console.log('NH Analyzer — callFuelIX retry/tagging regression test');
console.log('file: ' + gsPath);

console.log('\nTest 1 — succeeds immediately on 200, no retry');
{
  let calls = 0;
  const fetch = () => { calls++; return makeResponse(200, { choices: [{ message: { content: '  hello  ' }, finish_reason: 'stop' }] }); };
  const { callFuelIX, sleeps } = run(fetch);
  const result = callFuelIX('prompt');
  check('trims and returns content', result === 'hello');
  check('called fetch exactly once', calls === 1);
  check('never slept', sleeps.length === 0);
}

console.log('\nTest 2 — retries a 429 and succeeds once it clears');
{
  let calls = 0;
  const fetch = () => {
    calls++;
    if (calls < 3) return makeResponse(429, { error: 'rate limited' });
    return makeResponse(200, { choices: [{ message: { content: 'ok' }, finish_reason: 'stop' }] });
  };
  const { callFuelIX, sleeps } = run(fetch);
  const result = callFuelIX('prompt');
  check('returned content on the 3rd attempt', result === 'ok');
  check('called fetch exactly 3 times', calls === 3);
  check('slept between retries (2 sleeps)', sleeps.length === 2);
}

console.log('\nTest 3 — retries a 500 and gives up after 3 attempts, tagged error');
{
  let calls = 0;
  const fetch = () => { calls++; return makeResponse(500, 'Internal Server Error'); };
  const { callFuelIX, sleeps } = run(fetch);
  let threw = null;
  try { callFuelIX('prompt'); } catch (e) { threw = e; }
  check('threw after exhausting retries', !!threw);
  check('error is tagged AI_SERVICE_ERROR', threw && threw.message.indexOf('AI_SERVICE_ERROR') === 0);
  check('called fetch exactly 3 times', calls === 3);
  check('slept between attempts (2 sleeps)', sleeps.length === 2);
}

console.log('\nTest 4 — does NOT retry a 401 (bad/expired key) — fails fast');
{
  let calls = 0;
  const fetch = () => { calls++; return makeResponse(401, '{"error":"You do not have permission to access this resource"}'); };
  const { callFuelIX, sleeps } = run(fetch);
  let threw = null;
  try { callFuelIX('prompt'); } catch (e) { threw = e; }
  check('threw', !!threw);
  check('error is tagged AI_SERVICE_ERROR (not a lookup-source error)', threw && threw.message.indexOf('AI_SERVICE_ERROR') === 0);
  check('called fetch exactly once (no wasted retries on a bad key)', calls === 1);
  check('never slept', sleeps.length === 0);
}

console.log('\nTest 5 — a network-level exception (timeout/DNS) is retried like a 5xx');
{
  let calls = 0;
  const fetch = () => {
    calls++;
    if (calls < 2) throw new Error('Exception: Timeout');
    return makeResponse(200, { choices: [{ message: { content: 'recovered' }, finish_reason: 'stop' }] });
  };
  const { callFuelIX, sleeps } = run(fetch);
  const result = callFuelIX('prompt');
  check('recovered after a network error', result === 'recovered');
  check('called fetch exactly twice', calls === 2);
  check('slept once', sleeps.length === 1);
}

console.log('\nTest 6 — a malformed 200 response is tagged and not silently swallowed');
{
  const fetch = () => makeResponse(200, { choices: [] });
  const { callFuelIX } = run(fetch);
  let threw = null;
  try { callFuelIX('prompt'); } catch (e) { threw = e; }
  check('threw', !!threw);
  check('error is tagged AI_SERVICE_ERROR', threw && threw.message.indexOf('AI_SERVICE_ERROR') === 0);
}

console.log(failures === 0 ? '\nALL TESTS PASSED' : '\n' + failures + ' CHECK(S) FAILED');
process.exit(failures === 0 ? 0 : 1);
