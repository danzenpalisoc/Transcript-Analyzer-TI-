/**
 * getQAUsersFromSheet is callable by every signed-in user, and the page only needs a name from it
 * (the QA filter). It must return names only, never the address or role of a person, and must not
 * serve the old cached list (key v4) that held them.
 *
 * Usage: node tests/qa-users-directory-test.js
 */
const vm = require('vm');
const { read, FakeSheet, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

function makeEnv() {
  const store = {}, logs = [];
  const sandbox = {
    Logger: { log: m => logs.push(String(m)) }, JSON, Array, Object, String, Error,
    CacheService: { getScriptCache: () => ({ get: k => (k in store ? store[k] : null), put: (k, v) => { store[k] = v; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
  };
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('Config.gs'), ctx);
  vm.runInContext(read('ObserverHelper.gs'), ctx);
  const users = new FakeSheet(['Name', 'Email Address', 'Role'], [
    ['Ana Cruz', 'ana.cruz@telus.com', 'QA'],
    ['Ben Reyes', 'ben.reyes@telus.com', 'Admin'],
    ['', 'no.name@telus.com', 'QA'],
    ['Cara Diaz', '', ''],
  ]);
  users.getDataRange = () => ({ getValues: () => users.grid.map(r => r.slice()) });
  users.getName = () => 'Users';
  users.getSheetId = () => ctx.USERS_SHEET_GID;
  const ss = { getSheets: () => [users], getSheetByName: () => users };
  ctx.openSheetWithRetry = () => ss;
  return { ctx, store, logs };
}

console.log('\nTest 1 — names only');
{
  const e = makeEnv();
  const list = e.ctx.getQAUsersFromSheet();
  check('three people with a name are listed', list.length === 3, JSON.stringify(list));
  check('each entry has the name', list.map(u => u.name).join('|') === 'Ana Cruz|Ben Reyes|Cara Diaz');
  check('no entry has an email or a role key', list.every(u => Object.keys(u).join() === 'name'), JSON.stringify(list));
  const text = JSON.stringify(list);
  check('no address and no role appears anywhere in what is returned', !/@|Admin|QA/.test(text), text);
}

console.log('\nTest 2 — what is cached holds names only, and the old cached list is not served');
{
  const e = makeEnv();
  e.store['qa_users_list_v4'] = JSON.stringify([{ name: 'Old', email: 'old@telus.com', role: 'Admin' }]);
  const list = e.ctx.getQAUsersFromSheet();
  check('the old v4 cache entry is ignored', !list.some(u => u.name === 'Old') && list.length === 3, JSON.stringify(list));
  const keys = Object.keys(e.store);
  check('the new list is cached under v5', keys.indexOf('qa_users_list_v5') !== -1, keys.join());
  check('the cached copy holds no address or role', !/@|Admin/.test(e.store['qa_users_list_v5']), e.store['qa_users_list_v5']);
  const again = e.ctx.getQAUsersFromSheet();
  check('a second call is served from the cache with the same names', JSON.stringify(again) === JSON.stringify(list));
}

console.log('\nTest 3 — the page uses only the name');
{
  const html = read('index.html');
  const uses = (html.match(/_dbQAUsers[^;\n]*/g) || []).join('\n');
  check('index.html never reads .email or .role of a Users-sheet entry', !/\bu\.(email|role)\b/.test(uses), uses);
}

console.log('\n' + (check.state.failures === 0 ? 'ALL PASSED' : check.state.failures + ' FAILED'));
process.exit(check.state.failures === 0 ? 0 : 1);
