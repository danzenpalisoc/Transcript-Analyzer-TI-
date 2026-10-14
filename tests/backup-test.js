/**
 * backupSpreadsheet(): a copy of the data spreadsheet, made by an admin before a repair writes thousands of cells.
 * It must be admin-only, make exactly one copy next to the original, name it so it cannot be mistaken for the live file, and never touch the original.
 *
 * Usage: node tests/backup-test.js
 */
const vm = require('vm');
const { read, makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();

function makeEnv(opts) {
  opts = opts || {};
  const who = { email: 'danzen.palisoc@telus.com' }, logs = [], copies = [], touched = [];
  const folder = { id: 'FOLDER' };
  const original = { getId: () => 'ORIG', getParents: () => ({ hasNext: () => !opts.noParent, next: () => folder }),
    makeCopy: (name, parent) => { if (opts.copyFails) throw new Error('Service invoked too many times'); const c = { name, parent, getName: () => name, getUrl: () => 'https://docs.google.com/spreadsheets/d/COPY/edit', getId: () => 'COPY' }; copies.push(c); return c; },
    setSharing() { touched.push('sharing'); }, setTrashed() { touched.push('trashed'); }, setName() { touched.push('rename'); } };
  const sandbox = { Logger: { log: m => logs.push(String(m)) }, JSON, Math, Date, String, Error,
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Session: { getActiveUser: () => ({ getEmail: () => who.email }), getScriptTimeZone: () => 'UTC' },
    Utilities: { formatDate: (d, tz, f) => f.replace('yyyy', d.getUTCFullYear()).replace('MM', String(d.getUTCMonth() + 1).padStart(2, '0')).replace('dd', String(d.getUTCDate()).padStart(2, '0')).replace('HH', String(d.getUTCHours()).padStart(2, '0')).replace('mm', String(d.getUTCMinutes()).padStart(2, '0')) },
    DriveApp: { getFileById: id => { if (id !== 'ORIG') throw new Error('wrong file ' + id); return original; } } };
  const ctx = vm.createContext(sandbox);
  vm.runInContext(read('Config.gs'), ctx);
  vm.runInContext(read('DataRepair.gs'), ctx);
  ctx.getOrCreateSpreadsheet = () => ({ getId: () => 'ORIG', getName: () => 'NH Call Analyzer Data' });
  return { ctx, who, logs, copies, touched, folder };
}

console.log('\nTest 1 — an admin gets one copy, named as a backup, next to the original');
{
  const e = makeEnv();
  const r = e.ctx.backupSpreadsheet();
  check('one copy is made', e.copies.length === 1);
  check('its name says it is a backup and when (so it cannot be mistaken for the live file)', /^NH Call Analyzer Data - BACKUP \d{4}-\d{2}-\d{2} \d{2}-\d{2}$/.test(e.copies[0].name), e.copies[0].name);
  check('it is put in the same folder as the original', e.copies[0].parent === e.folder);
  check('the result gives the link and the name', r.success === true && /spreadsheets/.test(r.url) && r.name === e.copies[0].name, JSON.stringify(r));
  check('the log says where it is', e.logs.some(l => /Backup made/.test(l) && /COPY/.test(l)), e.logs.join(' | '));
  check('the original is not renamed, shared, or trashed', e.touched.length === 0, e.touched.join(','));
}

console.log('\nTest 2 — only an admin');
{
  const e = makeEnv(); e.who.email = 'some.analyst@telus.com';
  let err = ''; try { e.ctx.backupSpreadsheet(); } catch (x) { err = x.message; }
  check('an analyst is refused and nothing is copied', /admin/i.test(err) && e.copies.length === 0, err);
}

console.log('\nTest 3 — odd cases');
{
  const e = makeEnv({ noParent: true });
  e.ctx.backupSpreadsheet();
  check('a file with no folder is still copied (to the top of My Drive)', e.copies.length === 1 && e.copies[0].parent === undefined);
  const f = makeEnv({ copyFails: true });
  let err = ''; try { f.ctx.backupSpreadsheet(); } catch (x) { err = x.message; }
  check('if Drive refuses, the error is raised (so nobody goes on to repair without a backup)', /too many times/.test(err) && f.copies.length === 0, err);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
