/**
 * Every .gs file must parse. clasp push refuses a file with a syntax error, so one stray quote in a
 * long text (it happened in CreateActionRegistry.gs) leaves Apps Script on the OLD code while the
 * commit looks fine. This is the check that would have caught it.
 *
 * Usage: node tests/gs-syntax-test.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { makeCheck } = require('./lib/gas-fakes');
const check = makeCheck();
const root = path.join(__dirname, '..');

console.log('\nTest 1 — every .gs file parses');
for (const f of fs.readdirSync(root).filter(x => x.endsWith('.gs')).sort()) {
  let err = '';
  try { new vm.Script(fs.readFileSync(path.join(root, f), 'utf8'), { filename: f }); } catch (e) { err = e.message; }
  check(f + ' parses', !err, err);
}

console.log('\nTest 2 — the manifest is valid and asks for the permissions the code uses');
{
  let manifest = null, err = '';
  try { manifest = JSON.parse(fs.readFileSync(path.join(root, 'appsscript.json'), 'utf8')); } catch (e) { err = e.message; }
  check('appsscript.json is valid JSON', !!manifest, err);
  const scopes = (manifest && manifest.oauthScopes) || [];
  const all = fs.readdirSync(root).filter(x => x.endsWith('.gs')).map(x => fs.readFileSync(path.join(root, x), 'utf8')).join('\n');
  const need = [
    ['UrlFetchApp','https://www.googleapis.com/auth/script.external_request', 'UrlFetchApp'],
    ['DriveApp', 'https://www.googleapis.com/auth/drive', 'DriveApp'],
    ['SpreadsheetApp', 'https://www.googleapis.com/auth/spreadsheets', 'SpreadsheetApp'],
  ];
  need.forEach(([uses, scope, what]) => { if (all.indexOf(uses) !== -1) check('the manifest has the permission for ' + what, scopes.indexOf(scope) !== -1, scope + ' is missing'); });
  check('the scorecard deck needs no Google Slides permission (it is a PowerPoint file converted by Drive), so none is asked for', scopes.indexOf('https://www.googleapis.com/auth/presentations') === -1);
  const drv = manifest && manifest.dependencies && manifest.dependencies.enabledAdvancedServices || [];
  check('the advanced Drive service the conversion uses is enabled in the manifest', all.indexOf('Drive.Files.insert') === -1 || drv.some(s => s.userSymbol === 'Drive'));
  check('every scope is listed once', new Set(scopes).size === scopes.length);
}

console.log(check.state.failures ? '\nFAILED' : '\nALL TESTS PASSED');
process.exit(check.state.failures ? 1 : 0);
