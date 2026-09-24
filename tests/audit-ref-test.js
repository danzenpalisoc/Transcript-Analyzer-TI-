/**
 * Regression test for the audit reference collision bug.
 *
 * Reported symptom: an observer submits an evaluation, gets a confirmation
 * email with the correct audit ref (e.g. NHA-20260921-0003), but clicking
 * "view evaluation" shows a DIFFERENT team member and a different observer.
 *
 * Root cause: generateAuditRef() (Code.gs) sizes the CacheService TTL as
 * "seconds until midnight". CacheService.put() hard-caps expirationInSeconds
 * at 21600 (6 hours) and throws above that — so any call made more than 6
 * hours before midnight (i.e. almost every daytime submission) throws,
 * falls into the catch block, and returns a *random* 4-digit suffix with no
 * collision check. Two unrelated submissions on the same day can land on
 * the same ref. getEvalViewData()/getEvaluationByAuditRef() then return the
 * FIRST row matching that ref, so the later (colliding) submission is
 * permanently shadowed by the earlier one whenever anyone views it.
 *
 * This harness extracts generateAuditRef() out of Code.gs by source text
 * (same technique tests/race-test.js uses for index.html) and runs it
 * against stub GAS services that faithfully reproduce the real 21600s cap,
 * so it exercises the actual production logic, not a re-implementation.
 *
 * Usage: node tests/audit-ref-test.js [path-to-Code.gs]
 * Exits non-zero if any check fails.
 */
const fs = require('fs');
const vm = require('vm');
const path = require('path');

const gsPath = process.argv[2] || path.join(__dirname, '..', 'Code.gs');
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

const fnSource = extractFunction(src, 'generateAuditRef') + '\n' +
                 extractFunction(src, '_uniqueFallbackAuditRef');

// ── Stub GAS services ──────────────────────────────────────────────────────
// Real CacheService.put() throws "Argument too large: expirationInSeconds"
// outside [1, 21600] — https://developers.google.com/apps-script/reference/cache/cache
function makeCacheService() {
  const store = {};
  return {
    getScriptCache() {
      return {
        get(key) { return key in store ? store[key] : null; },
        put(key, value, expirationInSeconds) {
          if (expirationInSeconds < 1 || expirationInSeconds > 21600) {
            throw new Error('Argument too large: expirationInSeconds');
          }
          store[key] = value;
        },
      };
    },
  };
}
function makePropertiesService() {
  const store = {};
  return {
    getScriptProperties() {
      return {
        getProperty(key) { return key in store ? store[key] : null; },
        setProperty(key, value) { store[key] = value; },
      };
    },
  };
}
function makeLockService() {
  return { getScriptLock() { return { waitLock() {}, releaseLock() {} }; } };
}
function makeThrowingLockService(message) {
  return { getScriptLock() { return { waitLock() { throw new Error(message || 'Lock timeout: could not acquire lock within 5000ms.'); }, releaseLock() {} }; } };
}
// Stub for the Audit_Log sheet _uniqueFallbackAuditRef() consults — mirrors
// range.createTextFinder(text).matchEntireCell(true).findNext() from the real
// Sheets API closely enough to exercise the actual collision-check logic.
function makeAuditLogStub(existingRefs) {
  const refs = new Set(existingRefs || []);
  const sheet = {
    getRange() {
      return {
        createTextFinder(text) {
          return {
            matchEntireCell() { return this; },
            findNext() { return refs.has(text) ? {} : null; },
          };
        },
      };
    },
  };
  return {
    AUDIT_LOG_SHEET: 'Audit_Log',
    getOrCreateSpreadsheet: () => ({}),
    getOrCreateSheet: () => sheet,
  };
}

function runAt(fixedNow, extraGlobals) {
  const RealDate = Date;
  class FixedDate extends RealDate {
    constructor(...args) {
      if (args.length === 0) return new RealDate(fixedNow.getTime());
      return new RealDate(...args);
    }
    static now() { return fixedNow.getTime(); }
  }
  const logs = [];
  const sandbox = Object.assign({
    Date: FixedDate,
    Math, String, parseInt, Logger: { log(msg) { logs.push(String(msg)); } },
  }, makeAuditLogStub(), extraGlobals);
  sandbox.globalThis = sandbox;
  const ctx = vm.createContext(sandbox);
  vm.runInContext(fnSource, ctx);
  return { generateAuditRef: ctx.generateAuditRef, logs };
}

let failures = 0;
function check(name, cond) {
  console.log('  ' + (cond ? 'PASS' : 'FAIL') + '  ' + name);
  if (!cond) failures++;
}

// 11:02:41 local time, matching the reported bug's timestamp — 13+ hours
// before midnight, i.e. the TTL-cap bug fires on essentially every daytime
// submission, not some rare edge case.
const DAYTIME = new Date(2026, 8, 21, 11, 2, 41);

console.log('NH Analyzer — audit ref collision regression test');
console.log('file: ' + gsPath);

console.log('\nTest 1 — a daytime call must not silently fall back to an unchecked random suffix');
{
  const { generateAuditRef, logs } = runAt(DAYTIME, {
    CacheService: makeCacheService(),
    PropertiesService: makePropertiesService(),
    LockService: makeLockService(),
  });
  const ref = generateAuditRef();
  const hitFallback = logs.some(l => l.indexOf('generateAuditRef error') !== -1);
  check('ref generated (' + ref + ')', !!ref);
  check('did NOT hit the TTL-exception fallback path', !hitFallback);
}

console.log('\nTest 2 — two unrelated same-day submissions must not get the same ref');
{
  // Math.random pinned: proves the fallback path has NO collision protection —
  // if generateAuditRef() ever falls back to Math.random(), two calls with the
  // same random draw MUST collide. A correct implementation never depends on
  // Math.random() for uniqueness in the first place.
  const pinnedMath = Object.create(Math);
  pinnedMath.random = () => 0.42;

  const { generateAuditRef } = runAt(DAYTIME, {
    CacheService: makeCacheService(),
    PropertiesService: makePropertiesService(),
    LockService: makeLockService(),
    Math: pinnedMath,
  });

  const ref1 = generateAuditRef(); // e.g. Francisco auditing Keiry Marisol
  const ref2 = generateAuditRef(); // e.g. Danzen auditing Kenji

  check('refs are non-empty', !!ref1 && !!ref2);
  check('refs are different for different submissions (ref1=' + ref1 + ', ref2=' + ref2 + ')', ref1 !== ref2);
}

console.log('\nTest 3 — sequence must increment even though it is called well before 6pm (TTL cap zone)');
{
  const { generateAuditRef } = runAt(DAYTIME, {
    CacheService: makeCacheService(),
    PropertiesService: makePropertiesService(),
    LockService: makeLockService(),
  });
  const refs = [generateAuditRef(), generateAuditRef(), generateAuditRef()];
  const suffixes = refs.map(r => r.split('-').pop());
  check('three daytime calls produce three distinct sequential suffixes (got ' + suffixes.join(', ') + ')',
    new Set(suffixes).size === 3);
}

console.log('\nTest 4 — a lock timeout must not hand out a ref already used that day');
{
  // Two draws queued: the first matches an already-used ref (must be
  // rejected), the second is free and must be the one actually returned.
  const draws = [0, 0.1];
  const queuedMath = Object.create(Math);
  queuedMath.random = () => (draws.length > 1 ? draws.shift() : draws[0]);

  const { generateAuditRef } = runAt(DAYTIME, {
    CacheService: makeCacheService(),
    PropertiesService: makePropertiesService(),
    LockService: makeThrowingLockService(),
    Math: queuedMath,
    ...makeAuditLogStub(['NHA-20260921-0000']), // suffix Math.floor(0*9999) would collide
  });

  const ref = generateAuditRef();
  check('did not return the already-used ref (got ' + ref + ')', ref !== 'NHA-20260921-0000');
  check('returned the next free candidate instead', ref === 'NHA-20260921-0999');
}

console.log('\nTest 5 — if every guess collides, fall back to a suffix that cannot collide');
{
  const pinnedMath = Object.create(Math);
  pinnedMath.random = () => 0.42; // always drafts the same 4-digit suffix

  const { generateAuditRef } = runAt(DAYTIME, {
    CacheService: makeCacheService(),
    PropertiesService: makePropertiesService(),
    LockService: makeThrowingLockService(),
    Math: pinnedMath,
    ...makeAuditLogStub(['NHA-20260921-4199']), // the only suffix pinnedMath can ever draw
  });

  const ref = generateAuditRef();
  check('did not return the known-colliding ref (got ' + ref + ')', ref !== 'NHA-20260921-4199');
  check('used the timestamp escape hatch', /^NHA-20260921-T\d+$/.test(ref));
}

console.log(failures === 0 ? '\nALL TESTS PASSED' : '\n' + failures + ' CHECK(S) FAILED');
process.exit(failures === 0 ? 0 : 1);
