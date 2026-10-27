/**
 * ScorecardTargets.gs
 * The LOB scorecards (2026) as far as the app can measure them from an audit. Pure data and pure functions: no Sheets, no Slides.
 *
 * The production scorecards of each LOB (FCR, Transfers, RCT, RGU, Save Rate, Repeats, VTR, EWT, T2B, LCR) come from call
 * production data, which the app does not have. Only two of their KPIs have a counterpart in an audit:
 *   - FCR       <- the Resolution rate of the audit (a QA verdict, not the production FCR: the slide says "audit");
 *   - Transfers <- the Transfer rate of the Repeats audits (MOB CxSS, Q1 to Q3 only; lower is better).
 * Each is [Threshold, Target, Stretch]. Numbers were read from the screenshots of the admin's scorecards (Agent Scorecard).
 *
 * Levels: a value below the Threshold is "below"; from the Threshold up to the Target is "threshold"; from the Target up to the
 * Stretch is "target"; the Stretch or better is "stretch". For a lower-is-better KPI the same, the other way round.
 */

var SCORECARD_LOB_YEAR = 2026;

// FCR, higher is better. { quarter: [thresh, target, stretch] }
var SCORECARD_LOB_FCR = {
  'PF CxSS':  { 1: [74.5, 78.0, 81.0], 2: [74.5, 78.0, 81.0], 3: [74.0, 78.0, 80.5], 4: [77.0, 79.5, 82.0] },
  'PF CLS':   { 1: [74.0, 78.0, 83.0], 2: [74.0, 78.0, 83.0], 3: [74.0, 79.5, 82.0], 4: [76.5, 79.5, 82.0] },
  'PF TS':    { 1: [70.0, 76.0, 79.0], 2: [70.0, 76.0, 79.0], 3: [73.5, 78.0, 80.5], 4: [76.0, 78.5, 80.5] },
  'SHS CxSS': { 1: [75.0, 79.0, 84.0], 2: [75.0, 79.0, 84.0], 3: [76.5, 80.5, 83.5], 4: [78.5, 81.0, 83.5] },
  'SHS CLS':  { 1: [75.0, 80.0, 85.0], 2: [75.0, 80.0, 85.0], 3: [77.5, 81.5, 84.0], 4: [79.0, 81.5, 84.0] },
  'SHS TS':   { 1: [74.5, 78.5, 82.5], 2: [74.5, 80.0, 87.0], 3: [76.5, 81.5, 84.0], 4: [79.0, 81.5, 84.0] },
  'MOB CLS':  { 1: [72.0, 77.0, 82.0], 2: [72.5, 77.5, 82.5], 3: [77.5, 81.0, 84.0], 4: [78.5, 81.0, 84.0] },
  'MOB CxSS': { 1: [77.0, 80.0, 83.0], 2: [77.0, 80.0, 83.0], 3: [77.5, 81.0, 83.5], 4: [80.0, 82.5, 85.0] },
  'One CAM':  {             2: [77.5, 81.0, 86.0], 3: [81.5, 86.0, 88.5], 4: [81.5, 86.0, 88.5] }   // the scorecard has no Q1
};

// Transfers, LOWER is better (MOB CxSS only; the KPI is gone in Q4). { quarter: [thresh, target, stretch] }
var SCORECARD_LOB_TRANSFERS = {
  'MOB CxSS': { 1: [34.0, 28.0, 20.0], 2: [34.0, 28.0, 20.0], 3: [34.0, 28.0, 20.0] }
};

// What the Line of Business of an audit is called in the data (28 spellings) -> the scorecard it uses. Keys are lower case with
// single spaces. Decided with the admin: a name with MOB uses the MOB scorecards, the others use PF; Koodo is MOB CxSS.
var SCORECARD_LOB_ALIASES = {
  'pf cxss': 'PF CxSS', 'pf cls': 'PF CLS', 'pf ts': 'PF TS',
  'shs cxss': 'SHS CxSS', 'shs cls': 'SHS CLS', 'shs ts': 'SHS TS',
  'mob cls': 'MOB CLS', 'mob cxss': 'MOB CxSS',
  'mob ts': 'MOB CxSS', 'koodo': 'MOB CxSS', 'koodo ts': 'MOB CxSS', 'kdo cxss': 'MOB CxSS',
  'mob one cls': 'MOB CLS', 'mob onr cls': 'MOB CLS', 'cls2 access + mob': 'MOB CLS',
  'one cls': 'PF CLS',
  'one ts': 'PF TS', 'ohe ts': 'PF TS', 'ffh_fbr_ts_tiin': 'PF TS', 'ffh_fbr_ts_tip': 'PF TS', 'copper / pf ts': 'PF TS',
  'shs / vivint ts': 'SHS TS',
  'cxss': 'PF CxSS', 'ohe cxss': 'PF CxSS', 'ffh sales': 'PF CxSS',
  'one cam': 'One CAM'
};

// The scorecard a Line of Business uses, or '' when it has none (the agent then goes to the "Other LOB" section with the app's targets).
function scorecardLobKey_(lob) {
  var k = String(lob === null || lob === undefined ? '' : lob).replace(/\s+/g, ' ').trim().toLowerCase();
  return SCORECARD_LOB_ALIASES.hasOwnProperty(k) ? SCORECARD_LOB_ALIASES[k] : '';
}

// The targets of a scorecard at the quarter of a time (ms). When the scorecard has none for that quarter (One CAM Q1, or a date after
// 2026), the closest quarter that has one is used (the later one on a tie) and `exact` is false so the slide can say so.
// Returns { key, quarter, label: 'Q4 2026', exact, fcr: { thresh, target, stretch }, transfers: { ... } | null } or null.
function scorecardLobTargets_(key, ms) {
  var fcr = key && SCORECARD_LOB_FCR[key];
  if (!fcr) return null;
  var want;
  if (ms === null || ms === undefined || isNaN(ms)) want = SCORECARD_LOB_YEAR * 4 + 3;
  else { var p = _scParts_(ms); want = p.y * 4 + Math.floor(p.m / 3); }
  var best = null, bestD = null;
  Object.keys(fcr).forEach(function(q) {
    var idx = SCORECARD_LOB_YEAR * 4 + (Number(q) - 1), d = Math.abs(idx - want);
    if (best === null || d < bestD || (d === bestD && idx > best)) { best = idx; bestD = d; }
  });
  var quarter = best - SCORECARD_LOB_YEAR * 4 + 1;
  var tr = SCORECARD_LOB_TRANSFERS[key] && SCORECARD_LOB_TRANSFERS[key][quarter];
  var pack = function(a) { return { thresh: a[0], target: a[1], stretch: a[2] }; };
  return { key: key, quarter: quarter, label: 'Q' + quarter + ' ' + SCORECARD_LOB_YEAR, exact: best === want,
           fcr: pack(fcr[quarter]), transfers: tr ? pack(tr) : null };
}

// The level of a number against [thresh, target, stretch]: 'none' (no number), 'below', 'threshold', 'target' or 'stretch'.
function scorecardLevel_(value, t, higherIsBetter) {
  if (value === null || value === undefined || isNaN(value)) return 'none';
  if (higherIsBetter) return value >= t.stretch ? 'stretch' : (value >= t.target ? 'target' : (value >= t.thresh ? 'threshold' : 'below'));
  return value <= t.stretch ? 'stretch' : (value <= t.target ? 'target' : (value <= t.thresh ? 'threshold' : 'below'));
}
