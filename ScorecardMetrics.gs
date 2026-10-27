/**
 * ScorecardMetrics.gs
 * The numbers and wording on the agent scorecard slides. Pure functions: no Sheets, no Slides.
 *
 * Every rule here is a copy of a rule the dashboard uses (the helper block in index.html and
 * Dashboard.html: _dbNum, _dbResolution, _dbSplitFlags, _dbKey, _dbName), so a scorecard and the
 * dashboard always show the same number for the same audits. tests/scorecard-metrics-test.js
 * runs both on the same examples and fails if they ever disagree.
 *
 * A "row" is one merged audit, with the dashboard's field names ('Team Member', 'Issue Resolved',
 * 'Repeat Projection %', 'Critical Flags', ...).
 */

// A number from a sheet cell. "40%", 40 and " 40 " are 40; blank is null (not 0); a bare number above
// 0 and up to 1 is a percent cell (0.78 is 78); a sentence gives the first "NN%" in it.
function _scNum_(v) {
  if (v === null || v === undefined) return null;
  var s = String(v).trim();
  if (s === '') return null;
  var plain = s.replace('%', '').trim();
  if (plain !== '' && isFinite(Number(plain))) {
    var num = Number(plain);
    return (s.indexOf('%') === -1 && num > 0 && num <= 1) ? Math.round(num * 1000) / 10 : num;
  }
  var rng = /(\d+(?:\.\d+)?)\s*[-–]\s*(\d+(?:\.\d+)?)\s*%/.exec(s);
  if (rng) return (Number(rng[1]) + Number(rng[2])) / 2;
  var m = /(\d+(?:\.\d+)?)\s*%/.exec(s);
  if (m && Number(m[1]) <= 100) return Number(m[1]);
  return null;
}

// Issue Resolved is a verdict first, so the START of the text decides. 'yes' | 'partial' | 'no' | 'unknown'.
function _scResolution_(v) {
  var s = String(v === null || v === undefined ? '' : v).trim().toLowerCase().replace(/^[^a-z]+/, '');
  if (!s) return 'unknown';
  if (/^(partial|partially|in part|mostly)\b/.test(s)) return 'partial';
  if (/^(no|not|unresolved|completely unresolved|fully unresolved|never)\b/.test(s)) return 'no';
  if (/^(yes|resolved|completely resolved|fully resolved)\b/.test(s)) return 'yes';
  // A verdict written as a short sentence: "Issue not resolved. ...", "Issue resolved — ...", "Neither issue was resolved". A yes is not taken
  // from a sentence that goes on to qualify it (but / however / remains), so that one is left for a person to read.
  if (/^(?:the |this )?(?:primary |main |original )?(?:issue|problem|request)s?\s+(?:was |is |were |are |has been |have been )?(?:ultimately |eventually |still )?(?:not|never)\s+(?:been )?(?:fully |completely )?(?:resolved|fixed)\b/.test(s)) return 'no';
  if (/^neither (?:issue|problem|request) was (?:resolved|fixed)\b/.test(s)) return 'no';
  if (/^(?:the |this )?(?:primary |main |original )?(?:issue|problem|request)\s+(?:was |is |has been )?(?:ultimately |eventually |finally |successfully |fully |completely )*(?:resolved|fixed)\b/.test(s) && !/\b(but|however|although|except|remains?|remained|secondary|second issue|other issue)\b/.test(s.slice(0, 160))) return 'yes';
  var m = /^(?:the |this )?(?:customer'?s )?(?:issue|problem|concern|request)\b[^.]{0,25}?\b(was|is|has been|wasn'?t|isn'?t)\s+(not |never |partially |only partially )?(?:fully |completely )?(?:resolved|fixed)/.exec(s);
  if (m) return /partial/.test(m[2] || '') ? 'partial' : (/n'?t$/.test(m[1]) || /^(not|never)/.test(m[2] || '')) ? 'no' : 'yes';
  return 'unknown';
}

function _scKey_(s) { return String(s === null || s === undefined ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase(); }
function _scName_(s) { return String(s === null || s === undefined ? '' : s).replace(/\s+/g, ' ').trim(); }

// Critical Flags are stored in one cell: " | " between flags in new audits, ", " in older ones (a piece that
// starts in lower case belongs to the flag before it). "None" is no flag.
function _scSplitFlags_(s) {
  var raw = String(s === null || s === undefined ? '' : s).trim();
  if (!raw || /^none$/i.test(raw)) return [];
  var piped = raw.indexOf(' | ') !== -1;
  var parts = piped ? raw.split(' | ') : raw.split(',');
  var out = [];
  parts.forEach(function(p) {
    p = p.trim();
    if (!p || /^none$/i.test(p)) return;
    if (!piped && out.length && /^[a-z]/.test(p)) out[out.length - 1] += ', ' + p;
    else out.push(p);
  });
  // One flag written twice in one audit is one flag (the same capitals and spacing are the same flag), so every count agrees.
  var seen = {};
  return out.filter(function(f) { var k = f.replace(/\s+/g, ' ').trim().toLowerCase(); if (seen[k]) return false; seen[k] = true; return true; });
}

// Yes / No answers (Transfer Occurred, Sales Attempted). Anything else, such as a sentence or N/A, is 'unknown'
// and is not counted either way.
function _scYesNo_(v) {
  var s = String(v === null || v === undefined ? '' : v).trim().toLowerCase();
  if (/^yes\b/.test(s)) return 'yes';
  if (/^no\b/.test(s)) return 'no';
  return 'unknown';
}

// Cut at the end of a word, with an ellipsis, so a long sentence never overflows its box.
function _scTruncate_(text, max) {
  var s = String(text === null || text === undefined ? '' : text).replace(/\s+/g, ' ').trim();
  if (s.length <= max) return s;
  var cut = s.substring(0, max - 1);
  var sp = cut.lastIndexOf(' ');
  if (sp > max * 0.5) cut = cut.substring(0, sp);
  return cut.replace(/[\s,;:.\-–]+$/, '') + '…';
}

// The n most common values, ignoring capitals and spacing; ties keep the order they were first seen.
function _scTopLabels_(values, n, maxLen) {
  var counts = {}, labels = {}, order = [];
  values.forEach(function(v) {
    var label = _scName_(v);
    if (!label) return;
    var k = _scKey_(label);
    if (!labels[k]) { labels[k] = label; order.push(k); }
    counts[k] = (counts[k] || 0) + 1;
  });
  return order
    .map(function(k, i) { return { k: k, i: i }; })
    .sort(function(a, b) { return (counts[b.k] - counts[a.k]) || (a.i - b.i); })
    .slice(0, n)
    .map(function(o) { return _scTruncate_(labels[o.k], maxLen); });
}

// "What's Working" items from a text report: the report is saved in the Analysis sheet as PLAIN TEXT (htmlToPlainText), where the
// column is a heading line followed by bullet lines ("• item"). Reads the bullets after the heading and stops at the first line
// that is not a bullet (the next column heading).
function _scWorkingFromText_(text) {
  var lines = String(text).split(/\r?\n/), out = [], i = -1;
  for (var k = 0; k < lines.length; k++) { if (/what.{0,3}s\s+working/i.test(lines[k])) { i = k; break; } }
  if (i === -1) return [];
  for (var j = i + 1; j < lines.length; j++) {
    var ln = lines[j].replace(/^\s+|\s+$/g, '');
    if (!ln) continue;   // the saved text has blank lines between the bullets
    var m = /^[\u2022*-]\s*(.+)$/.exec(ln);
    if (!m) break;
    var item = m[1].replace(/\s+/g, ' ').trim();
    if (item) out.push(item);
  }
  return out.slice(0, 3);
}

// "What's Working" items from a saved report: the <li> items of the report-col-working column when the report is HTML, or the
// bullets under the heading when it is the plain text the Analysis sheet holds. Never throws; returns [] when it is not there.
function extractWhatsWorking_(html) {
  try {
    var h = String(html === null || html === undefined ? '' : html);
    var start = h.indexOf('report-col-working');
    if (start === -1) return _scWorkingFromText_(h);
    var end = h.indexOf('report-col ', start + 20);
    var block = h.substring(start, end === -1 ? Math.min(h.length, start + 6000) : end);
    var out = [], re = /<li[^>]*>([\s\S]*?)<\/li>/gi, m;
    while ((m = re.exec(block)) !== null) {
      var t = m[1].replace(/<[^>]*>/g, ' ')
        .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"').replace(/&#39;|&#x27;/g, "'")
        .replace(/\s+/g, ' ').trim();
      if (t) out.push(t);
    }
    return out.slice(0, 3);
  } catch (e) { return []; }
}

var _SC_MONTHS_ = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
// Year, month (0-11) and day of a time in the script's time zone (the audits are made and read in local time); UTC only when
// Apps Script is not there (the tests).
function _scParts_(ms) {
  try {
    if (typeof Utilities !== 'undefined' && typeof Session !== 'undefined') {
      var p = Utilities.formatDate(new Date(ms), Session.getScriptTimeZone(), 'yyyy-M-d').split('-');
      return { y: Number(p[0]), m: Number(p[1]) - 1, d: Number(p[2]) };
    }
  } catch (e) {}
  var u = new Date(ms);
  return { y: u.getUTCFullYear(), m: u.getUTCMonth(), d: u.getUTCDate() };
}
function _scDate_(ms) { var p = _scParts_(ms); return _SC_MONTHS_[p.m] + ' ' + p.d; }
// "Oct 1 - Oct 13, 2026" (or "Oct 1, 2026" for one day).
function _scDateRange_(minMs, maxMs) {
  if (minMs === null || maxMs === null || isNaN(minMs) || isNaN(maxMs)) return '';
  var y1 = _scParts_(minMs).y, y2 = _scParts_(maxMs).y;
  if (_scDate_(minMs) === _scDate_(maxMs) && y1 === y2) return _scDate_(minMs) + ', ' + y1;
  return _scDate_(minMs) + (y1 !== y2 ? ', ' + y1 : '') + ' - ' + _scDate_(maxMs) + ', ' + y2;
}

function _scRound1_(x) { return Math.round(x * 10) / 10; }
// A number for a target on the slide: 78 for 78.0, 79.5 for 79.5.
function _scF1_(n) { return String(Math.round(n * 10) / 10); }

// status of one tile against its threshold, comparing the number as it is displayed
function _scStatus_(value, th) {
  if (value === null || value === undefined) return 'none';
  var ok = th.direction === 'max' ? value <= th.target : value >= th.target;
  return ok ? 'target' : 'below';
}

function _scTargetText_(th, unit) {
  return 'Target ' + (th.direction === 'max' ? '<= ' : '>= ') + th.target + unit;
}

// The words for a status. 'threshold' and 'stretch' only exist for the KPIs that use a LOB scorecard (four levels).
function _scStatusText_(status) {
  return status === 'target' ? 'On target' : (status === 'below' ? 'Below threshold' : (status === 'threshold' ? 'At threshold' : (status === 'stretch' ? 'Stretch' : 'No data')));
}

// One card per agent. rows: merged audits. extras.whatsWorking: { auditRef: [text, ...] } (optional).
// Returns { cards: [...], skipped: [{ ref, reason }], range: 'Oct 1 - Oct 13, 2026', total, sections: [...] }.
// sections: the cards grouped by LOB scorecard, in alphabetical order with "Other LOB" last: { name, key, scorecard, cards }.
// With ScorecardTargets.gs loaded an agent is judged against the scorecard of the LOB they were audited in most (the FCR and Transfers
// targets of the quarter of the newest audit of the whole section); without it there is one section and the app's own targets.
function buildAgentCards_(rows, extras) {
  var TH = (typeof SCORECARD_THRESHOLDS !== 'undefined') ? SCORECARD_THRESHOLDS : null;
  if (!TH) throw new Error('SCORECARD_THRESHOLDS is not defined');
  var ww = (extras && extras.whatsWorking) || {};
  var byAgent = {}, order = [], skipped = [], minMs = null, maxMs = null;

  (rows || []).forEach(function(r) {
    var raw = String(r['Team Member'] === null || r['Team Member'] === undefined ? '' : r['Team Member']);
    var name = _scName_(raw);
    if (!name) { skipped.push({ ref: r['Audit Ref'] || '', reason: 'no Team Member' }); return; }
    if (name.length > 80 || raw.indexOf('\n') !== -1) { skipped.push({ ref: r['Audit Ref'] || '', reason: 'Team Member is not a name (a column shifted)' }); return; }
    var k = _scKey_(name);
    if (!byAgent[k]) { byAgent[k] = { name: name, rows: [] }; order.push(k); }
    byAgent[k].rows.push(r);
    var t = Date.parse(String(r['Timestamp'] === undefined ? '' : r['Timestamp']));
    if (!isNaN(t)) { if (minMs === null || t < minMs) minMs = t; if (maxMs === null || t > maxMs) maxMs = t; }
  });

  // Which scorecard each agent is judged by, and the quarter of each section (the newest audit in it)
  var haveLob = (typeof scorecardLobKey_ === 'function' && typeof scorecardLobTargets_ === 'function');
  var agentSection = {}, sectionKey = {}, sectionMs = {};
  order.forEach(function(k) {
    var a = byAgent[k], lobKey = haveLob ? scorecardLobKey_(_scMostCommon_(a.rows.map(function(r) { return r['Line of Business']; }))) : '';
    var sec = haveLob ? (lobKey || 'Other LOB') : '';
    agentSection[k] = sec; sectionKey[sec] = lobKey;
    a.rows.forEach(function(r) {
      var t = Date.parse(String(r['Timestamp'] === undefined ? '' : r['Timestamp']));
      if (!isNaN(t) && (sectionMs[sec] === undefined || t > sectionMs[sec])) sectionMs[sec] = t;
    });
  });
  var ctxOf = {};
  Object.keys(sectionKey).forEach(function(sec) { ctxOf[sec] = sectionKey[sec] ? scorecardLobTargets_(sectionKey[sec], sectionMs[sec] === undefined ? null : sectionMs[sec]) : null; });

  var cards = order.map(function(k) {
    var card = _scBuildCard_(byAgent[k], TH, ww, ctxOf[agentSection[k]]);
    card.section = agentSection[k];
    return card;
  });
  cards.sort(function(a, b) { return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : (a.name.toLowerCase() > b.name.toLowerCase() ? 1 : 0); });

  var names = Object.keys(sectionKey).sort(function(a, b) {
    if (a === 'Other LOB') return b === 'Other LOB' ? 0 : 1;
    if (b === 'Other LOB') return -1;
    return a.toLowerCase() < b.toLowerCase() ? -1 : (a.toLowerCase() > b.toLowerCase() ? 1 : 0);
  });
  if (!names.length) names = [''];
  var sections = names.map(function(name) {
    return { name: name, key: sectionKey[name] || '', scorecard: ctxOf[name] || null, cards: cards.filter(function(c) { return c.section === name; }) };
  });
  return { cards: cards, skipped: skipped, range: _scDateRange_(minMs, maxMs), total: (rows || []).length, sections: sections };
}

function _scMostCommon_(values) { var t = _scTopLabels_(values, 1, 200); return t.length ? t[0] : ''; }

// lob: the scorecard targets this agent is judged by (scorecardLobTargets_), or null for the app's own targets.
function _scBuildCard_(agent, TH, ww, lob) {
  var rows = agent.rows, n = rows.length;
  var isRepeats = function(r) { return String(r['Analysis Type'] || '').indexOf('Repeats') !== -1; };
  var isSales = function(r) { return String(r['Analysis Type'] || '').indexOf('Sales') !== -1; };
  var repeatsRows = rows.filter(isRepeats), salesRows = rows.filter(isSales);

  // Avg Repeat Risk: the audits that have a number (a real 0% counts, blank does not)
  var sum = 0, cnt = 0;
  rows.forEach(function(r) { var v = _scNum_(r['Repeat Projection %']); if (v !== null) { sum += v; cnt++; } });
  var repeatRisk = cnt ? Math.round(sum / cnt) : null;

  // Resolution rate: Resolved of the audits that have a clear verdict
  var rc = { yes: 0, partial: 0, no: 0, unknown: 0 };
  rows.forEach(function(r) { rc[_scResolution_(r['Issue Resolved'])]++; });
  var den = rc.yes + rc.partial + rc.no;
  var resolution = den ? Math.round(rc.yes / den * 100) : null;

  // Critical flags: per audit
  var flagTotal = 0;
  rows.forEach(function(r) { flagTotal += _scSplitFlags_(r['Critical Flags']).length; });
  var flagsRaw = n ? flagTotal / n : null;                        // judged as it really is: 1 flag in 21 audits is NOT 0
  var flagsPer = flagsRaw === null ? null : _scRound1_(flagsRaw);   // shown to one decimal

  // Transfer rate (Repeats audits) and Sales attempted rate (Sales audits)
  var tc = { yes: 0, no: 0 };
  repeatsRows.forEach(function(r) { var a = _scYesNo_(r['Transfer Occurred']); if (a !== 'unknown') tc[a]++; });
  var transfer = (tc.yes + tc.no) ? Math.round(tc.yes / (tc.yes + tc.no) * 100) : null;
  var sc = { yes: 0, no: 0 };
  salesRows.forEach(function(r) { var a = _scYesNo_(r['Sales Attempted']); if (a !== 'unknown') sc[a]++; });
  var salesAttempt = (sc.yes + sc.no) ? Math.round(sc.yes / (sc.yes + sc.no) * 100) : null;

  var salesOnly = salesRows.length > 0 && repeatsRows.length === 0;
  var tiles = [
    { key: 'repeatRisk', label: 'Avg Repeat Risk', value: repeatRisk, display: repeatRisk === null ? 'N/A' : repeatRisk + '%',
      status: _scStatus_(repeatRisk, TH.repeatRisk), target: _scTargetText_(TH.repeatRisk, '%'),
      sub: cnt === n ? n + ' audit' + (n === 1 ? '' : 's') : cnt + ' of ' + n + ' audits' },
    { key: 'resolution', label: 'Resolution rate', value: resolution, display: resolution === null ? 'N/A' : resolution + '%',
      status: _scStatus_(resolution, TH.resolution), target: _scTargetText_(TH.resolution, '%'),
      sub: den ? 'Partial ' + rc.partial + ' | Unresolved ' + rc.no : 'No verdict recorded' },
    { key: 'flags', label: 'Critical flags', value: flagsPer, display: flagsPer === null ? 'N/A' : (flagsRaw > 0 && flagsPer === 0 ? '<0.1' : flagsPer.toFixed(1)),
      status: _scStatus_(flagsRaw, TH.flags), target: _scTargetText_(TH.flags, ' per audit'),
      sub: flagTotal + ' flag' + (flagTotal === 1 ? '' : 's') + ' in ' + n + ' audit' + (n === 1 ? '' : 's') },
    salesOnly
      ? { key: 'salesAttempt', label: 'Sales attempted', value: salesAttempt, display: salesAttempt === null ? 'N/A' : salesAttempt + '%',
          status: _scStatus_(salesAttempt, TH.salesAttempt), target: _scTargetText_(TH.salesAttempt, '%'),
          sub: (sc.yes + sc.no) + ' of ' + salesRows.length + ' Sales audits' }
      : { key: 'transfer', label: 'Transfer rate', value: transfer, display: transfer === null ? 'N/A' : transfer + '%',
          status: _scStatus_(transfer, TH.transfer), target: _scTargetText_(TH.transfer, '%'),
          sub: (tc.yes + tc.no) + ' of ' + repeatsRows.length + ' Repeats audits' }
  ];
  if (lob) {
    // The KPIs that have a counterpart on the LOB scorecard use its four levels; the others keep the app's target and say so.
    tiles.forEach(function(t) {
      if (t.key === 'resolution') {
        t.label = 'Resolution (audit)'; t.status = scorecardLevel_(t.value, lob.fcr, true);
        t.target = 'Target >= ' + _scF1_(lob.fcr.target) + '% (LOB)'; t.lobKpi = 'FCR';
      } else if (t.key === 'transfer' && lob.transfers) {
        t.label = 'Transfer rate (audit)'; t.status = scorecardLevel_(t.value, lob.transfers, false);
        t.target = 'Target <= ' + _scF1_(lob.transfers.target) + '% (LOB)'; t.lobKpi = 'Transfers';
      } else if (t.key !== 'flags') {
        t.target += ' (audit)';
      }
    });
  }
  tiles.forEach(function(t) { t.statusText = _scStatusText_(t.status); });

  var emailed = rows.filter(function(r) { return r['Email Status'] === 'Sent'; }).length;
  var supporting = [
    { label: 'Audits', display: String(n) },
    { label: 'Repeats', display: String(repeatsRows.length) },
    { label: 'Sales', display: String(salesRows.length) },
    { label: 'Emailed', display: n ? Math.round(emailed / n * 100) + '%' : '—' }
  ];

  var below = tiles.filter(function(t) { return t.status === 'below'; });
  var measured = tiles.filter(function(t) { return t.status !== 'none'; });
  var badge = !measured.length ? { level: 'nodata', text: 'No data' }
    : below.length === 0 ? { level: 'ontrack', text: 'On track' }
    : below.length === 1 ? { level: 'attention', text: 'Needs attention' }
    : { level: 'priority', text: 'Priority support' };

  // Gaps: the flags this agent gets most, then RCA sub-parameters when there are fewer than three flags
  var allFlags = [];
  rows.forEach(function(r) { _scSplitFlags_(r['Critical Flags']).forEach(function(f) { allFlags.push(f); }); });
  var gaps = _scTopLabels_(allFlags, 3, 70);
  if (gaps.length < 3) {
    var rca = [];
    rows.forEach(function(r) { String(r['RCA Sub-Parameter'] || '').split(/\s*;\s*/).forEach(function(x) { rca.push(x); }); });
    _scTopLabels_(rca, 3, 70).forEach(function(g) { if (gaps.length < 3 && gaps.map(_scKey_).indexOf(_scKey_(g)) === -1) gaps.push(g); });
  }

  // Strengths: the "What's Working" items of the agent's audits, else the areas that meet their target
  var working = [];
  rows.forEach(function(r) { (ww[r['Audit Ref']] || []).forEach(function(x) { working.push(x); }); });
  var strengths = _scTopLabels_(working, 2, 120);
  if (!strengths.length) {
    strengths = tiles.filter(function(t) { return t.status === 'target'; }).slice(0, 2)
      .map(function(t) { return 'Meets the target for ' + t.label.toLowerCase() + ' (' + t.display + ').'; });
  }

  // Also noted: the areas below target, in one or two sentences
  var noted = below.slice(0, 2).map(function(t) {
    var tgt = t.target.replace(/ \((LOB|audit)\)$/, '');   // the "(LOB)" and "(audit)" marks are on the tile, not in the sentence
    return t.label + ': ' + t.display + (t.key === 'flags' ? ' per audit' : '') + ' (' + tgt.charAt(0).toLowerCase() + tgt.slice(1) + ').';
  }).join(' ');
  if (!noted) noted = measured.length ? 'All measured areas are within target.' : '';
  if (!salesOnly && salesAttempt !== null) {
    var line = 'Sales attempted on ' + salesAttempt + '% of Sales audits.';
    noted = noted ? noted + ' ' + line : line;
  }

  var coaching = _scTopLabels_(rows.map(function(r) { return r['Top SMART Opportunity']; }), 3, 110);

  var times = rows.map(function(r) { return Date.parse(String(r['Timestamp'] === undefined ? '' : r['Timestamp'])); }).filter(function(x) { return !isNaN(x); });
  var vtid = '';
  rows.forEach(function(r) { var v = _scName_(r['VTID']); if (v) vtid = v; });
  var sap = _scMostCommon_(rows.map(function(r) { return r['SAP ID']; }));

  return {
    name: agent.name, sap: sap, audits: n,
    teamLeader: _scMostCommon_(rows.map(function(r) { return r['Team Leader']; })),
    lob: _scMostCommon_(rows.map(function(r) { return r['Line of Business']; })),
    locale: _scMostCommon_(rows.map(function(r) { return r['Locale']; })),
    vtid: vtid,
    range: times.length ? _scDateRange_(Math.min.apply(null, times), Math.max.apply(null, times)) : '',
    tiles: tiles, supporting: supporting, badge: badge, belowCount: below.length,
    strengths: strengths, gaps: gaps, alsoNoted: noted, coaching: coaching,
    scorecard: lob ? { key: lob.key, label: lob.label, exact: lob.exact } : null
  };
}

// ── The slides about the whole group ──────────────────────────────────────────

var _SC_KPI_ORDER_ = ['repeatRisk', 'resolution', 'flags', 'transfer', 'salesAttempt'];
var _SC_KPI_SHORT_ = { repeatRisk: 'Repeat', resolution: 'Resolve', flags: 'Flags', transfer: 'Transfer', salesAttempt: 'Sales' };

// Target as a short phrase for a table cell: "<= 30%", ">= 80%", "<= 0 / audit".
function _scTargetShort_(th, key) {
  return (th.direction === 'max' ? '<= ' : '>= ') + th.target + (key === 'flags' ? '/audit' : '%');
}

// The numbers and sentences of the group slides of ONE section (the agents of one LOB scorecard), from the agent cards
// (buildAgentCards_). Pure: no Sheets, no Slides.
// Averages are simple means of the agents' own numbers. A KPI that no agent has a number for is left out; an agent that has no tile for a
// KPI (Transfer for a Sales-only agent, Sales attempted for the others) is not counted in it. In a section with a LOB scorecard the
// Resolution (FCR) and, for MOB CxSS in Q1 to Q3, the Transfer KPI have four levels (below, threshold, target, stretch); the others two.
//   { label, range, agents, audits, filters, section: { name, quarterLabel, exact, hasLob, key, fcr, transfers },
//     kpis: [{ key, label, avg, display, status, statusText, target, targetShort, targetValue, onTarget, below, threshold, stretch,
//              none, applicable, measured, short, isPercent, lobLevels }],
//     strengths: [text], gaps: [{ label, count }], levels: { ontrack|attention|priority|nodata: [{ name, reason }] }, counts: { ... },
//     nextSteps: [{ title, body }] }
// section: one of built.sections (optional: without it all the cards of `built` are one section without a name).
function buildGroupSummary_(built, label, filtersText, audits, section) {
  var TH = SCORECARD_THRESHOLDS, cards = (section ? section.cards : built.cards) || [], N = cards.length, sc = section && section.scorecard ? section.scorecard : null;

  var kpis = [];
  _SC_KPI_ORDER_.forEach(function(key) {
    var th = TH[key], vals = [], counts = { below: 0, threshold: 0, target: 0, stretch: 0, none: 0 };
    cards.forEach(function(c) {
      var t = c.tiles.filter(function(x) { return x.key === key; })[0];
      if (!t) return;
      counts[counts.hasOwnProperty(t.status) ? t.status : 'none']++;
      if (t.value !== null && t.value !== undefined) vals.push(t.value);
    });
    if (!vals.length) return;
    var sum = 0; vals.forEach(function(v) { sum += v; });
    var avg = sum / vals.length, isFlags = key === 'flags', rounded = _scRound1_(avg), avgNum = isFlags ? avg : Math.round(avg);
    var shown = isFlags ? (avg > 0 && rounded === 0 ? '<0.1' : rounded.toFixed(1)) : Math.round(avg) + '%';

    // the LOB scorecard that applies to this KPI in this section, if any
    var lobT = null, lobName = '';
    if (sc && key === 'resolution') { lobT = sc.fcr; lobName = 'Resolution (audit)'; }
    else if (sc && key === 'transfer' && sc.transfers) { lobT = sc.transfers; lobName = 'Transfer rate (audit)'; }
    var higher = key !== 'transfer';
    var status = lobT ? scorecardLevel_(avgNum, lobT, higher) : _scStatus_(avgNum, th);
    var onTarget = counts.target + counts.stretch, measured = counts.below + counts.threshold + counts.target + counts.stretch;
    kpis.push({ key: key, label: lobT ? lobName : th.label, avg: avgNum, display: shown, status: status, statusText: _scStatusText_(status),
      target: lobT ? 'Target ' + (higher ? '>= ' : '<= ') + _scF1_(lobT.target) + '% (LOB)' : _scTargetText_(th, isFlags ? ' per audit' : '%') + (sc && !isFlags ? ' (audit)' : ''),
      targetShort: lobT ? (higher ? '>= ' : '<= ') + _scF1_(lobT.target) + '%' : _scTargetShort_(th, key),
      targetValue: lobT ? lobT.target : th.target, threshValue: lobT ? lobT.thresh : th.target,
      onTarget: onTarget, below: counts.below, threshold: counts.threshold, stretch: counts.stretch, none: counts.none,
      applicable: measured + counts.none, measured: measured, short: _SC_KPI_SHORT_[key], isPercent: !isFlags, lobLevels: !!lobT });
  });

  var levels = { ontrack: [], attention: [], priority: [], nodata: [] };
  cards.forEach(function(c) {
    var lvl = levels[c.badge.level] ? c.badge.level : 'nodata';
    var belowNames = c.tiles.filter(function(t) { return t.status === 'below'; }).map(function(t) { return _SC_KPI_SHORT_[t.key] || t.label; });
    var reason = lvl === 'ontrack' ? 'All measured KPIs on target' : (lvl === 'nodata' ? 'No KPI could be measured' : 'Below: ' + belowNames.join(', '));
    levels[lvl].push({ name: c.name, reason: reason });
  });
  var counts = { ontrack: levels.ontrack.length, attention: levels.attention.length, priority: levels.priority.length, nodata: levels.nodata.length };

  var allStrengths = [];
  cards.forEach(function(c) { (c.strengths || []).forEach(function(s) { allStrengths.push(s); }); });
  var strengths = _scTopLabels_(allStrengths, 4, 90);

  var gapCount = {}, gapLabel = {}, gapOrder = [];
  cards.forEach(function(c) {
    var seen = {};
    (c.gaps || []).forEach(function(g) {
      var k = _scKey_(g); if (!k || seen[k]) return;
      seen[k] = true;
      if (!gapLabel[k]) { gapLabel[k] = g; gapOrder.push(k); }
      gapCount[k] = (gapCount[k] || 0) + 1;
    });
  });
  var gaps = gapOrder.map(function(k, i) { return { k: k, i: i }; })
    .sort(function(a, b) { return (gapCount[b.k] - gapCount[a.k]) || (a.i - b.i); })
    .slice(0, 5).map(function(o) { return { label: gapLabel[o.k], count: gapCount[o.k] }; });

  // Next steps: one card for each KPI with agents below its threshold or target (the most agents first), then the priority-support agents.
  var steps = kpis.filter(function(k) { return k.below > 0; })
    .sort(function(a, b) { return b.below - a.below; })
    .map(function(k) {
      var isMax = k.key !== 'resolution' && k.key !== 'salesAttempt', unit = k.key === 'flags' ? ' per audit' : '%';
      var word = k.lobLevels ? 'threshold' : (isMax ? 'limit' : 'target');
      return { title: k.label, body: k.below + ' of ' + k.measured + ' agent' + (k.measured === 1 ? ' is ' : 's are ') + (isMax ? 'above the ' : 'below the ') +
        _scF1_(k.threshValue) + unit + ' ' + word + '.' };
    });
  if (steps.length > 5) steps = steps.slice(0, 5);
  if (levels.priority.length) {
    var names = levels.priority.map(function(p) { return p.name; }), shownNames = names.slice(0, 4);
    steps.push({ title: 'Priority support', body: shownNames.join(', ') + (names.length > shownNames.length ? ' and ' + (names.length - shownNames.length) + ' more.' : '.') });
  }
  if (!steps.length) steps.push({ title: 'All on target', body: kpis.length ? 'Every measured KPI is on target for every agent.' : 'No KPI could be measured for these audits.' });

  return { label: label || '', range: built.range || '', agents: N, audits: audits, filters: filtersText || '',
    section: section ? { name: section.name, key: section.key, hasLob: !!sc, quarterLabel: sc ? sc.label : '', exact: sc ? sc.exact : true,
                         fcr: sc ? sc.fcr : null, transfers: sc ? sc.transfers : null } : null,
    kpis: kpis, strengths: strengths, gaps: gaps, levels: levels, counts: counts, nextSteps: steps };
}
