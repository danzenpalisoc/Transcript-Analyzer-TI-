/**
 * ScorecardLayout.gs
 * The look of the scorecard slides, as plain data: where every box goes, its colour, its text and size.
 * Pure functions, no Slides calls. The HTML preview (tests/lib/scorecard-preview.js) and the real Google
 * Slides builder (ScorecardSlides.gs) both read THIS, so what was approved on the preview is what the deck gets.
 *
 * Page: 16:9, 720 x 405 pt (the Google Slides default). Font: Arial (the safe stand-in for TELUS Helvetica:
 * it is on every machine and survives a PowerPoint export). Colours are the TELUS palette the app already uses.
 * Status is never colour alone: every tile says "On target" or "Below threshold" in words.
 */

var SCORECARD_THEME = {
  font: 'Arial',
  purple: '#4B286D',       // headings, accent, dark panel
  lilac: '#F5F0FF',        // card background
  white: '#FFFFFF',
  ink: '#2B2B2B',          // body text on light
  grey: '#54595F',         // secondary text
  green: '#2A7A00',        // on target; a shade darker than the app green so it passes 4.5:1 on lilac too
  brightGreen: '#66CC00',  // only on the purple panel
  red: '#C12335',          // below threshold, priority support
  amber: '#8C4A00',        // needs attention
  line: '#D8CCE6',         // thin tile border
  violet: '#8A63B2',       // the large circle on the cover (no text on it)
  onDark: '#EDE6F5'        // soft text on the purple cover (contrast about 9:1 on purple)
};

var SCORECARD_PAGE = { w: 720, h: 405 };

// Smallest font on a slide: 11 pt, except the footer (9 pt).
var SCORECARD_MIN_FONT = 11;
var SCORECARD_FOOTER_FONT = 9;

// Average glyph width of Arial as a share of the font size; a little high on purpose so a box never overflows.
var SCORECARD_CHAR_W = 0.52;
var SCORECARD_LINE_H = 1.2;

function _slCharsPerLine_(widthPt, size) { return Math.max(4, Math.floor(widthPt / (size * SCORECARD_CHAR_W))); }
function _slLines_(text, widthPt, size) {
  var cpl = _slCharsPerLine_(widthPt, size), words = String(text).split(/\s+/), lines = 1, cur = 0;
  words.forEach(function(w) {
    var len = w.length;
    if (cur === 0) { cur = len; }
    else if (cur + 1 + len <= cpl) { cur += 1 + len; }
    else { lines++; cur = len; }
    while (cur > cpl) { lines++; cur -= cpl; }
  });
  return lines;
}
// Height in pt that a paragraph needs.
function _slHeight_(lines, size, spacing) { return lines * size * SCORECARD_LINE_H * (spacing || 1); }

// Cut text so it fits in `maxLines` lines of the given width and size (at a word end, with an ellipsis).
function _slFitText_(text, widthPt, size, maxLines) {
  var t = String(text === null || text === undefined ? '' : text).replace(/\s+/g, ' ').trim();
  if (_slLines_(t, widthPt, size) <= maxLines) return t;
  var words = t.split(' '), out = '';
  for (var i = 0; i < words.length; i++) {
    var next = out ? out + ' ' + words[i] : words[i];
    if (_slLines_(next + '…', widthPt, size) > maxLines) break;
    out = next;
  }
  if (!out) out = t.substring(0, Math.max(1, _slCharsPerLine_(widthPt, size) * maxLines - 1));
  return out.replace(/[\s,;:.\-–]+$/, '') + '…';
}

// Fit a paragraph that starts with a bold label: the label takes room too, so fit label + body and give back the body.
function _slFitWithLabel_(label, body, widthPt, size, maxLines) {
  var fitted = _slFitText_(label + body, widthPt, size, maxLines);
  return fitted.substring(label.length);
}

// The four levels of a LOB KPI have their own colours: stretch purple, target green, threshold amber, below red; no data grey.
function _slStatusColor_(status) {
  var T = SCORECARD_THEME;
  return status === 'target' ? T.green : (status === 'below' ? T.red : (status === 'threshold' ? T.amber : (status === 'stretch' ? T.purple : T.grey)));
}
function _slBadgeColor_(level) {
  var T = SCORECARD_THEME;
  return level === 'ontrack' ? T.green : (level === 'attention' ? T.amber : (level === 'priority' ? T.red : T.grey));
}

function _slRun_(t, size, bold, color) { return { t: String(t), size: size, bold: !!bold, color: color }; }
function _slShape_(id, type, x, y, w, h, o) {
  o = o || {};
  return { id: id, type: type, x: x, y: y, w: w, h: h, fill: o.fill || null, line: o.line || null, radius: o.radius || 0,
           paras: o.paras || [], valign: o.valign || 'TOP', pad: o.pad === undefined ? 0 : o.pad };
}
function _slPara_(runs, o) { o = o || {}; return { runs: runs, align: o.align || 'START', bullet: !!o.bullet, spacing: o.spacing || 1 }; }

// The shapes of one agent's slide, in drawing order. pageNo is 1-based.
function scorecardSlideShapes_(card, pageNo, totalPages) {
  var T = SCORECARD_THEME, P = SCORECARD_PAGE, S = [];
  var nameSize = card.name.length <= 28 ? 30 : (card.name.length <= 36 ? 26 : 22);
  var name = _slFitText_(card.name, 480, nameSize, 1);

  S.push(_slShape_('bg', 'RECTANGLE', 0, 0, P.w, P.h, { fill: T.white }));
  S.push(_slShape_('accent', 'RECTANGLE', 0, 0, P.w, 8, { fill: T.purple }));
  S.push(_slShape_('name', 'TEXT_BOX', 36, 22, 490, 44, { valign: 'MIDDLE', paras: [_slPara_([_slRun_(name, nameSize, true, T.purple)])] }));
  S.push(_slShape_('badge', 'ROUND_RECTANGLE', 546, 28, 138, 30, { fill: _slBadgeColor_(card.badge.level), radius: 15, valign: 'MIDDLE',
    paras: [_slPara_([_slRun_(card.badge.text, 12, true, T.white)], { align: 'CENTER' })] }));
  var sub = ['SAP ID ' + (card.sap || '—'), card.audits + ' audited interaction' + (card.audits === 1 ? '' : 's')];
  if (card.range) sub.push(card.range);
  if (card.scorecard) sub.push(card.scorecard.key + ' ' + card.scorecard.label + ' targets' + (card.scorecard.exact ? '' : ' (closest)'));
  S.push(_slShape_('sub', 'TEXT_BOX', 36, 68, 648, 20, { valign: 'MIDDLE', paras: [_slPara_([_slRun_(_slFitText_(sub.join('  |  '), 648, 12, 1), 12, false, T.grey)])] }));

  // left card: scorecard
  S.push(_slShape_('card-score', 'ROUND_RECTANGLE', 36, 100, 330, 268, { fill: T.lilac, radius: 10 }));
  S.push(_slShape_('card-score-title', 'TEXT_BOX', 48, 106, 306, 22, { valign: 'MIDDLE', paras: [_slPara_([_slRun_('Audit scorecard', 14, true, T.purple)])] }));
  var tilePos = [[48, 130], [207, 130], [48, 224], [207, 224]];
  card.tiles.forEach(function(t, i) {
    var x = tilePos[i][0], y = tilePos[i][1], col = _slStatusColor_(t.status);
    S.push(_slShape_('tile-' + i, 'ROUND_RECTANGLE', x, y, 147, 88, { fill: T.white, line: T.line, radius: 8, pad: 8,
      paras: [
        _slPara_([_slRun_(t.label, 11, false, T.grey)]),
        _slPara_([_slRun_(_slFitText_(t.display, 131, 26, 1), 26, true, col)]),
        _slPara_([_slRun_(t.statusText, 11, true, col)]),
        _slPara_([_slRun_(_slFitText_(t.target, 131, 11, 1), 11, false, T.grey)])
      ] }));
  });
  var supX = [48, 126, 204, 282];
  card.supporting.forEach(function(s, i) {
    S.push(_slShape_('sup-' + i, 'ROUND_RECTANGLE', supX[i], 320, 72, 42, { fill: T.white, line: T.line, radius: 8, pad: 4, valign: 'MIDDLE',
      paras: [_slPara_([_slRun_(s.display, 16, true, T.purple)], { align: 'CENTER' }), _slPara_([_slRun_(s.label, 11, false, T.grey)], { align: 'CENTER' })] }));
  });

  // right top: audit feedback (7 lines of 11 pt)
  var bodyW = 282, size = 12, spacing = 1.1;
  var sTxt = _slFitWithLabel_('Strengths: ', card.strengths.length ? card.strengths.join('; ') : 'No report details available.', bodyW, size, 2);
  var gTxt = _slFitWithLabel_('Gaps: ', card.gaps.length ? card.gaps.join('; ') : 'No repeated gaps in these audits.', bodyW, size, 2);
  var nTxt = _slFitWithLabel_('Also noted: ', card.alsoNoted || 'Nothing else to note.', bodyW, size, 3);
  S.push(_slShape_('card-feedback', 'ROUND_RECTANGLE', 378, 100, 306, 146, { fill: T.lilac, radius: 10 }));
  S.push(_slShape_('card-feedback-title', 'TEXT_BOX', 390, 106, 282, 22, { valign: 'MIDDLE', paras: [_slPara_([_slRun_('Audit feedback', 14, true, T.purple)])] }));
  S.push(_slShape_('feedback', 'TEXT_BOX', 390, 130, bodyW, 112, { paras: [
    _slPara_([_slRun_('Strengths: ', size, true, T.green), _slRun_(sTxt, size, false, T.ink)], { spacing: spacing }),
    _slPara_([_slRun_('Gaps: ', size, true, card.gaps.length ? T.red : T.grey), _slRun_(gTxt, size, false, T.ink)], { spacing: spacing }),
    _slPara_([_slRun_('Also noted: ', size, true, T.purple), _slRun_(nTxt, size, false, T.ink)], { spacing: spacing })
  ] }));

  // right bottom: coaching advice (5 lines of 11 pt), bullets added while they still fit
  S.push(_slShape_('card-coaching', 'ROUND_RECTANGLE', 378, 254, 306, 114, { fill: T.purple, radius: 10 }));
  S.push(_slShape_('card-coaching-title', 'TEXT_BOX', 390, 258, 282, 20, { valign: 'MIDDLE', paras: [_slPara_([_slRun_('Coaching advice', 14, true, T.brightGreen)])] }));
  var bulletW = 268, linesLeft = 5, bullets = [];
  (card.coaching.length ? card.coaching : ['No coaching points recorded for these audits.']).forEach(function(c) {
    if (linesLeft < 1) return;
    var fitted = _slFitText_(c, bulletW, size, Math.min(2, linesLeft));
    linesLeft -= _slLines_(fitted, bulletW, size);
    bullets.push(_slPara_([_slRun_(fitted, size, false, T.white)], { bullet: true, spacing: spacing }));
  });
  S.push(_slShape_('coaching', 'TEXT_BOX', 390, 280, 282, 84, { paras: bullets }));

  // footer
  var foot = [card.teamLeader, card.lob, card.locale, card.vtid, 'Audit Scorecard'].filter(function(x) { return x; }).join('  |  ');
  S.push(_slShape_('footer', 'TEXT_BOX', 36, 378, 560, 18, { valign: 'MIDDLE', paras: [_slPara_([_slRun_(_slFitText_(foot, 560, SCORECARD_FOOTER_FONT, 1), SCORECARD_FOOTER_FONT, false, T.grey)])] }));
  S.push(_slShape_('page', 'TEXT_BOX', 620, 378, 64, 18, { valign: 'MIDDLE', paras: [_slPara_([_slRun_(String(pageNo) + ' / ' + totalPages, SCORECARD_FOOTER_FONT, false, T.grey)], { align: 'END' })] }));
  return S;
}

// The cover slide: dark purple, the title, the period, the counts and the filters, with two circles on the right.
// summary: { label, range, agents, audits, filters }. (The old `thresholds` list is no longer shown here: the targets
// are on the "Scorecard results" slide.)
function scorecardCoverShapes_(summary, totalPages) {
  var T = SCORECARD_THEME, P = SCORECARD_PAGE, S = [];
  S.push(_slShape_('bg', 'RECTANGLE', 0, 0, P.w, P.h, { fill: T.purple }));
  S.push(_slShape_('circle-large', 'ELLIPSE', 440, 40, 240, 240, { fill: T.violet }));
  S.push(_slShape_('circle-small', 'ELLIPSE', 560, 220, 112, 112, { fill: T.brightGreen }));
  S.push(_slShape_('title', 'TEXT_BOX', 48, 120, 380, 56, { valign: 'MIDDLE', paras: [_slPara_([_slRun_('Audit Scorecards', 40, true, T.white)])] }));
  var line2 = [summary.label, summary.range].filter(function(x) { return x; }).join('  |  ');
  S.push(_slShape_('subtitle', 'TEXT_BOX', 48, 184, 380, 44, { paras: [_slPara_([_slRun_(_slFitText_(line2, 380, 16, 2), 16, false, T.onDark)])] }));
  S.push(_slShape_('counts', 'TEXT_BOX', 48, 234, 380, 24, { valign: 'MIDDLE', paras: [_slPara_([_slRun_(summary.agents + ' agent' + (summary.agents === 1 ? '' : 's') + '  |  ' + summary.audits + ' audit' + (summary.audits === 1 ? '' : 's'), 14, true, T.white)])] }));
  if (summary.filters) S.push(_slShape_('filters', 'TEXT_BOX', 48, 264, 380, 36, { paras: [_slPara_([_slRun_('Filters: ', 12, true, T.white), _slRun_(_slFitText_(summary.filters, 320, 12, 2), 12, false, T.onDark)])] }));
  if (summary.lobs) S.push(_slShape_('lobs', 'TEXT_BOX', 48, 304, 380, 36, { paras: [_slPara_([_slRun_('LOBs: ', 12, true, T.white), _slRun_(_slFitText_(summary.lobs, 330, 12, 2), 12, false, T.onDark)])] }));
  S.push(_slShape_('footer', 'TEXT_BOX', 36, 378, 560, 18, { valign: 'MIDDLE', paras: [_slPara_([_slRun_('Audit Scorecard', SCORECARD_FOOTER_FONT, false, T.onDark)])] }));
  S.push(_slShape_('page', 'TEXT_BOX', 620, 378, 64, 18, { valign: 'MIDDLE', paras: [_slPara_([_slRun_('1 / ' + totalPages, SCORECARD_FOOTER_FONT, false, T.onDark)], { align: 'END' })] }));
  return S;
}

// ── The slides about the whole group (the numbers come from buildGroupSummary_ in ScorecardMetrics.gs) ──

function _slFoot_(g) {
  var s = g.section, q = s && s.hasLob ? s.quarterLabel + ' targets' + (s.exact ? '' : ' (closest quarter)') : '';
  return [g.label, g.range, q, 'Audit Scorecard'].filter(function(x) { return x; }).join('  |  ');
}

// Every group slide starts the same way: white page, purple strip, a title, the footer and the page number. A deck with LOB
// sections also puts the name of the LOB in a purple tag at the right of the title (so the title is a little smaller and narrower).
function _slGroupBase_(title, pageNo, totalPages, g) {
  var T = SCORECARD_THEME, P = SCORECARD_PAGE, S = [], tag = g.section && g.section.name ? g.section.name : '';
  S.push(_slShape_('bg', 'RECTANGLE', 0, 0, P.w, P.h, { fill: T.white }));
  S.push(_slShape_('accent', 'RECTANGLE', 0, 0, P.w, 8, { fill: T.purple }));
  if (tag) {
    S.push(_slShape_('title', 'TEXT_BOX', 36, 22, 470, 44, { valign: 'MIDDLE', paras: [_slPara_([_slRun_(_slFitText_(title, 470, 24, 1), 24, true, T.purple)])] }));
    S.push(_slShape_('lob-tag', 'ROUND_RECTANGLE', 520, 28, 164, 30, { fill: T.purple, radius: 15, pad: 6, valign: 'MIDDLE',
      paras: [_slPara_([_slRun_(_slFitText_(tag, 152, 12, 1), 12, true, T.white)], { align: 'CENTER' })] }));
  } else {
    S.push(_slShape_('title', 'TEXT_BOX', 36, 22, 648, 44, { valign: 'MIDDLE', paras: [_slPara_([_slRun_(_slFitText_(title, 648, 28, 1), 28, true, T.purple)])] }));
  }
  S.push(_slShape_('footer', 'TEXT_BOX', 36, 378, 560, 18, { valign: 'MIDDLE', paras: [_slPara_([_slRun_(_slFitText_(_slFoot_(g), 560, SCORECARD_FOOTER_FONT, 1), SCORECARD_FOOTER_FONT, false, T.grey)])] }));
  S.push(_slShape_('page', 'TEXT_BOX', 620, 378, 64, 18, { valign: 'MIDDLE', paras: [_slPara_([_slRun_(pageNo + ' / ' + totalPages, SCORECARD_FOOTER_FONT, false, T.grey)], { align: 'END' })] }));
  return S;
}

function _slText_(id, x, y, w, h, runs, o) { o = o || {}; return _slShape_(id, 'TEXT_BOX', x, y, w, h, { valign: o.valign || 'TOP', paras: [_slPara_(runs, { align: o.align })] }); }
function _slLegendItem_(S, id, x, y, color, text, textW) {
  var T = SCORECARD_THEME;
  S.push(_slShape_(id + '-box', 'RECTANGLE', x, y + 4, 12, 12, { fill: color }));
  S.push(_slText_(id, x + 18, y, textW, 20, [_slRun_(text, 11, false, T.grey)], { valign: 'MIDDLE' }));
}

// 1. Executive summary: one tile per KPI (the average of the agents), and three panels.
function scorecardSummaryShapes_(g, pageNo, totalPages) {
  var T = SCORECARD_THEME, S = _slGroupBase_('Executive summary', pageNo, totalPages, g);
  var k = g.kpis.slice(0, 5), n = k.length;
  if (n) {
    var gap = 10, w = (648 - (n - 1) * gap) / n, inner = w - 20;
    k.forEach(function(t, i) {
      S.push(_slShape_('kpi-' + i, 'ROUND_RECTANGLE', 36 + i * (w + gap), 76, w, 108, { fill: T.lilac, radius: 10, pad: 10, paras: [
        _slPara_([_slRun_(_slFitText_(t.display, inner, 28, 1), 28, true, _slStatusColor_(t.status))]),
        _slPara_([_slRun_(_slFitText_(t.label, inner, 11, 2), 11, false, T.ink)]),
        _slPara_([_slRun_(_slFitText_(t.target, inner, 11, 2), 11, false, T.grey)])
      ] }));
    });
  } else {
    S.push(_slShape_('kpi-none', 'ROUND_RECTANGLE', 36, 76, 648, 108, { fill: T.lilac, radius: 10, pad: 14, valign: 'MIDDLE',
      paras: [_slPara_([_slRun_('No KPI could be measured for these audits.', 14, true, T.grey)])] }));
  }
  S.push(_slText_('averages-note', 36, 187, 648, 16, [_slRun_('Averages are simple means of the ' + g.agents + ' agent' + (g.agents === 1 ? '' : 's') + '.', 11, false, T.grey)], { valign: 'MIDDLE' }));

  var ready = [];
  [['ontrack', 'On track', T.green], ['attention', 'Needs attention', T.amber], ['priority', 'Priority support', T.red]].forEach(function(l) {
    ready.push(_slPara_([_slRun_(g.counts[l[0]] + ' ' + l[1], 12, true, l[2])], { bullet: true, spacing: 1.1 }));
  });
  if (g.counts.nodata) ready.push(_slPara_([_slRun_(g.counts.nodata + ' No data', 12, true, T.grey)], { bullet: true, spacing: 1.1 }));
  var bulletW = 170;
  var well = (g.strengths.length ? g.strengths.slice(0, 3) : ['No report details available.']).map(function(s) { return _slPara_([_slRun_(_slFitText_(s, bulletW, 12, 2), 12, false, T.ink)], { bullet: true, spacing: 1.1 }); });
  var work = (g.gaps.length ? g.gaps.slice(0, 3).map(function(x) { return x.label + ' (' + x.count + ' of ' + g.agents + ')'; }) : ['No repeated gaps in these audits.'])
    .map(function(s) { return _slPara_([_slRun_(_slFitText_(s, bulletW, 12, 2), 12, false, T.ink)], { bullet: true, spacing: 1.1 }); });
  var panels = [['well', 'What went well', T.green, well], ['work', 'What needs work', T.amber, work], ['ready', 'Readiness', T.purple, ready]], px = [36, 256, 476];
  panels.forEach(function(p, i) {
    S.push(_slShape_('panel-' + p[0], 'ROUND_RECTANGLE', px[i], 208, 208, 162, { fill: T.lilac, radius: 10 }));
    S.push(_slShape_('panel-' + p[0] + '-dot', 'ELLIPSE', px[i] + 12, 218, 22, 22, { fill: p[2] }));
    S.push(_slText_('panel-' + p[0] + '-title', px[i] + 42, 216, 154, 26, [_slRun_(p[1], 14, true, T.purple)], { valign: 'MIDDLE' }));
    S.push(_slShape_('panel-' + p[0] + '-list', 'TEXT_BOX', px[i] + 12, 250, 184, 114, { paras: p[3] }));
  });
  return S;
}

// 2. Scorecard results: the targets and the group average in a table, a bar chart of the percent KPIs, and how the badge is read.
function scorecardResultsShapes_(g, pageNo, totalPages) {
  var T = SCORECARD_THEME, S = _slGroupBase_('Scorecard results', pageNo, totalPages, g);
  var cols = [['KPI', 132], ['Target', 72], ['Average', 56], ['On target', 64]], x0 = 36;
  var xs = []; cols.reduce(function(x, c) { xs.push(x); return x + c[1]; }, x0);
  cols.forEach(function(c, i) {
    S.push(_slShape_('th-' + i, 'RECTANGLE', xs[i], 76, c[1], 26, { fill: T.purple, pad: 4, valign: 'MIDDLE', paras: [_slPara_([_slRun_(c[0], 11, true, T.white)])] }));
  });
  g.kpis.slice(0, 5).forEach(function(k, r) {
    var y = 102 + r * 28, cells = [
      [_slFitText_(k.label, 124, 11, 1), true, T.ink, 11], [_slFitText_(k.targetShort, 64, 11, 1), false, T.ink, 11],
      [k.display, true, _slStatusColor_(k.status), 12], [k.onTarget + ' of ' + k.measured, false, T.ink, 11]];
    cells.forEach(function(c, i) {
      S.push(_slShape_('td-' + r + '-' + i, 'RECTANGLE', xs[i], y, cols[i][1], 28, { fill: T.white, line: T.line, pad: 4, valign: 'MIDDLE', paras: [_slPara_([_slRun_(c[0], c[3], c[1], c[2])])] }));
    });
  });
  if (!g.kpis.length) S.push(_slText_('no-kpi', 36, 110, 324, 40, [_slRun_('No KPI could be measured for these audits.', 12, false, T.grey)]));
  var s = g.section, noteParas = [
    _slPara_([_slRun_('Badge: no KPI below threshold = On track, one = Needs attention, two or more = Priority support.', 11, false, T.ink)], { bullet: true, spacing: 1.1 }),
    _slPara_([_slRun_('Averages are simple means of the ' + g.agents + ' agent' + (g.agents === 1 ? '' : 's') + '. Critical flags are not in the chart.', 11, false, T.ink)], { bullet: true, spacing: 1.1 })
  ];
  if (s && s.hasLob) {
    var f = s.fcr, lobNote = 'Resolution (audit) vs the FCR scorecard of ' + s.key + ' ' + s.quarterLabel + ': Threshold ' + _scF1_(f.thresh) + '%, Target ' + _scF1_(f.target) + '%, Stretch ' + _scF1_(f.stretch) + '%.';
    noteParas.push(_slPara_([_slRun_(lobNote, 11, false, T.ink)], { bullet: true, spacing: 1.1 }));
  }
  S.push(_slShape_('note', 'ROUND_RECTANGLE', 36, 250, 324, 122, { fill: T.lilac, radius: 10, pad: 10, paras: noteParas }));

  // bar chart: the percent KPIs, the average beside the target
  var pk = g.kpis.filter(function(k) { return k.isPercent; }), cx0 = 400, cw = 272, base = 282, scale = 1.7;
  S.push(_slText_('chart-title', 376, 76, 308, 22, [_slRun_('Average vs target (%)', 12, true, T.purple)], { valign: 'MIDDLE' }));
  if (pk.length) {
    [0, 50, 100].forEach(function(v) {
      S.push(_slShape_('grid-' + v, 'RECTANGLE', cx0, base - v * scale, cw, 1, { fill: v === 0 ? T.grey : T.line }));
      S.push(_slText_('grid-' + v + '-label', cx0 - 28, base - v * scale - 8, 24, 16, [_slRun_(String(v), 11, false, T.grey)], { valign: 'MIDDLE', align: 'END' }));
    });
    var gw = cw / pk.length, bw = Math.min(26, gw / 2 - 6);
    pk.forEach(function(k, i) {
      var cx = cx0 + i * gw + gw / 2, bars = [['avg', k.avg, T.purple, cx - bw - 2], ['target', k.targetValue, T.green, cx + 2]];
      bars.forEach(function(b) {
        var v = Math.max(0, Math.min(100, b[1])), h = Math.max(1, v * scale);
        S.push(_slShape_('bar-' + i + '-' + b[0], 'RECTANGLE', b[3], base - h, bw, h, { fill: b[2] }));
        S.push(_slText_('bar-' + i + '-' + b[0] + '-value', b[3] - 8, base - h - 16, bw + 16, 14, [_slRun_(String(Math.round(b[1])), 11, true, T.ink)], { valign: 'MIDDLE', align: 'CENTER' }));
      });
      S.push(_slText_('cat-' + i, cx0 + i * gw, base + 4, gw, 16, [_slRun_(_slFitText_(k.short, gw - 4, 11, 1), 11, false, T.grey)], { valign: 'MIDDLE', align: 'CENTER' }));
    });
    _slLegendItem_(S, 'legend-avg', cx0, 312, T.purple, 'Average', 60);
    _slLegendItem_(S, 'legend-target', cx0 + 100, 312, T.green, 'Target', 60);
  } else {
    S.push(_slText_('no-chart', 376, 120, 308, 40, [_slRun_('No percentage KPI to chart.', 12, false, T.grey)]));
  }
  return S;
}

// 3. Where the agents landed: one bar per KPI, split into below threshold, on target and no data; three notes on the right.
function scorecardLandedShapes_(g, pageNo, totalPages) {
  var T = SCORECARD_THEME, S = _slGroupBase_('Where the agents landed on each KPI', pageNo, totalPages, g);
  var bx = 164, bwTotal = 296;
  g.kpis.slice(0, 5).forEach(function(k, i) {
    var y = 84 + i * 48, x = bx;
    S.push(_slText_('lk-' + i, 36, y, 122, 34, [_slRun_(_slFitText_(k.label, 122, 12, 1), 12, true, T.ink)], { valign: 'MIDDLE' }));
    // the parts: below threshold, at threshold, on target, stretch, no data (the two middle levels exist only for a LOB KPI)
    var target = k.lobLevels ? k.onTarget - k.stretch : k.onTarget;
    [['below', k.below, T.red], ['thr', k.threshold, T.amber], ['on', target, T.green], ['str', k.stretch, T.purple], ['none', k.none, T.grey]].forEach(function(seg) {
      if (!seg[1] || !k.applicable) return;
      var w = bwTotal * seg[1] / k.applicable;
      S.push(_slShape_('lb-' + i + '-' + seg[0], 'RECTANGLE', x, y, w, 34, { fill: seg[2], valign: 'MIDDLE', paras: w >= 24 ? [_slPara_([_slRun_(String(seg[1]), 12, true, T.white)], { align: 'CENTER' })] : [] }));
      x += w;
    });
  });
  if (!g.kpis.length) S.push(_slText_('no-kpi', 36, 90, 420, 30, [_slRun_('No KPI could be measured for these audits.', 12, false, T.grey)]));
  if (g.section && g.section.hasLob) {
    _slLegendItem_(S, 'legend-below', bx, 330, T.red, 'Below threshold', 100);
    _slLegendItem_(S, 'legend-thr', bx + 132, 330, T.amber, 'At threshold', 80);
    _slLegendItem_(S, 'legend-on', bx + 230, 330, T.green, 'On target', 62);
    _slLegendItem_(S, 'legend-str', bx + 324, 330, T.purple, 'Stretch', 50);
    _slLegendItem_(S, 'legend-none', bx + 408, 330, T.grey, 'No data', 50);
  } else {
    _slLegendItem_(S, 'legend-below', bx, 330, T.red, 'Below threshold', 100);
    _slLegendItem_(S, 'legend-on', bx + 130, 330, T.green, 'On target', 70);
    _slLegendItem_(S, 'legend-none', bx + 230, 330, T.grey, 'No data', 60);
  }
  g.kpis.slice().sort(function(a, b) { return b.below - a.below; }).slice(0, 3).forEach(function(k, i) {
    var body = _slFitWithLabel_(k.label + ': ', k.onTarget + ' of ' + k.measured + ' on target; ' + k.below + ' below threshold.', 188, 12, 3);
    S.push(_slShape_('lnote-' + i, 'ROUND_RECTANGLE', 476, 84 + i * 82, 208, 72, { fill: T.lilac, radius: 10, pad: 10, valign: 'MIDDLE',
      paras: [_slPara_([_slRun_(k.label + ': ', 12, true, T.purple), _slRun_(body, 12, false, T.ink)], { spacing: 1.1 })] }));
  });
  return S;
}

// 4. Heatmap: every agent, a coloured cell per KPI. 14 agents on a slide, in two tables of 7.
var SCORECARD_HEATMAP_PER_SLIDE = 14;
function scorecardHeatmapShapes_(cards, pageIdx, pageCount, pageNo, totalPages, g) {
  var T = SCORECARD_THEME, S = _slGroupBase_('Scorecard heatmap by agent' + (pageCount > 1 ? ' (' + (pageIdx + 1) + ' of ' + pageCount + ')' : ''), pageNo, totalPages, g);
  var chunk = cards.slice(pageIdx * SCORECARD_HEATMAP_PER_SLIDE, (pageIdx + 1) * SCORECARD_HEATMAP_PER_SLIDE);
  var heads = ['Repeat risk', 'Resolve rate', 'Flags / audit', 'Xfer / Sales'];
  [['l', 36, chunk.slice(0, 7)], ['r', 368, chunk.slice(7)]].forEach(function(t) {
    var id = t[0], x = t[1], list = t[2];
    if (!list.length) return;
    S.push(_slShape_(id + '-th-name', 'RECTANGLE', x, 76, 124, 32, { fill: T.purple, pad: 4, valign: 'MIDDLE', paras: [_slPara_([_slRun_('Agent', 11, true, T.white)])] }));
    heads.forEach(function(h, c) {
      S.push(_slShape_(id + '-th-' + c, 'RECTANGLE', x + 124 + c * 48, 76, 48, 32, { fill: T.purple, pad: 2, valign: 'MIDDLE', paras: [_slPara_([_slRun_(h, 11, true, T.white)], { align: 'CENTER' })] }));
    });
    list.forEach(function(card, r) {
      var y = 108 + r * 30;
      // the name may take two lines (a long name is cut with an ellipsis only after that)
      S.push(_slShape_(id + '-name-' + r, 'RECTANGLE', x, y, 124, 30, { fill: T.white, line: T.line, pad: 2, valign: 'MIDDLE', paras: [_slPara_([_slRun_(_slFitText_(card.name, 120, 11, 2), 11, true, T.ink)])] }));
      card.tiles.slice(0, 4).forEach(function(tile, c) {
        S.push(_slShape_(id + '-c-' + r + '-' + c, 'RECTANGLE', x + 124 + c * 48, y, 48, 30, { fill: _slStatusColor_(tile.status), line: T.white, pad: 2, valign: 'MIDDLE',
          paras: [_slPara_([_slRun_(_slFitText_(tile.display, 44, 12, 1), 12, true, T.white)], { align: 'CENTER' })] }));
      });
    });
  });
  var sec = g.section, note;
  if (sec && sec.hasLob) {
    _slLegendItem_(S, 'legend-below', 36, 326, T.red, 'Below threshold', 100);
    _slLegendItem_(S, 'legend-thr', 168, 326, T.amber, 'At threshold', 80);
    _slLegendItem_(S, 'legend-on', 280, 326, T.green, 'On target', 62);
    _slLegendItem_(S, 'legend-str', 374, 326, T.purple, 'Stretch', 50);
    _slLegendItem_(S, 'legend-none', 456, 326, T.grey, 'No data', 50);
    note = 'Resolve rate: the audit Resolution against the FCR scorecard of ' + sec.key + ' ' + sec.quarterLabel + '. Xfer / Sales: Transfer rate (Repeats agents) or Sales attempted (Sales agents). Alphabetical order.';
  } else {
    _slLegendItem_(S, 'legend-on', 36, 326, T.green, 'On target', 70);
    _slLegendItem_(S, 'legend-below', 140, 326, T.red, 'Below threshold', 100);
    _slLegendItem_(S, 'legend-none', 276, 326, T.grey, 'No data', 60);
    note = 'Xfer / Sales: Transfer rate for Repeats agents, Sales attempted for Sales agents. Agents are listed in alphabetical order.';
  }
  S.push(_slText_('heat-note', 36, 348, 648, 28, [_slRun_(note, 11, false, T.grey)]));
  return S;
}

// 5. Audit findings: the strengths that come up most, and a bar chart of how many agents have each gap.
function scorecardFindingsShapes_(g, pageNo, totalPages) {
  var T = SCORECARD_THEME, S = _slGroupBase_('Audit findings: strengths and gaps', pageNo, totalPages, g);
  S.push(_slShape_('card-strengths', 'ROUND_RECTANGLE', 36, 84, 300, 284, { fill: T.lilac, radius: 10 }));
  S.push(_slShape_('strengths-dot', 'ELLIPSE', 48, 96, 22, 22, { fill: T.green }));
  S.push(_slText_('strengths-title', 78, 94, 246, 26, [_slRun_('Strengths', 14, true, T.purple)], { valign: 'MIDDLE' }));
  var items = g.strengths.length ? g.strengths.slice(0, 4) : ['No report details available.'];
  S.push(_slShape_('strengths-list', 'TEXT_BOX', 48, 130, 276, 230, { paras: items.map(function(s) { return _slPara_([_slRun_(_slFitText_(s, 262, 12, 3), 12, false, T.ink)], { bullet: true, spacing: 1.1 }); }) }));
  S.push(_slText_('gaps-title', 352, 80, 332, 22, [_slRun_(_slFitText_('Agents flagged in audits, by gap (of ' + g.agents + ')', 332, 12, 1), 12, true, T.purple)], { valign: 'MIDDLE' }));
  if (!g.gaps.length) {
    S.push(_slText_('no-gaps', 352, 120, 332, 40, [_slRun_('No repeated gaps in these audits.', 12, false, T.grey)]));
  } else {
    var max = Math.max.apply(null, g.gaps.map(function(x) { return x.count; }));
    g.gaps.slice(0, 5).forEach(function(gap, i) {
      var y = 112 + i * 50, w = Math.max(4, 110 * gap.count / max);
      S.push(_slText_('gap-' + i + '-label', 352, y, 170, 36, [_slRun_(_slFitText_(gap.label, 170, 11, 2), 11, false, T.ink)], { valign: 'MIDDLE', align: 'END' }));
      S.push(_slShape_('gap-' + i + '-bar', 'RECTANGLE', 530, y + 7, w, 22, { fill: T.purple }));
      S.push(_slText_('gap-' + i + '-value', 530 + w + 6, y + 7, 36, 22, [_slRun_(String(gap.count), 12, true, T.ink)], { valign: 'MIDDLE' }));
    });
  }
  return S;
}

// 6. Readiness: the agents in three columns, by badge, with a line of reason under each name.
function scorecardReadinessShapes_(g, pageNo, totalPages) {
  var T = SCORECARD_THEME, S = _slGroupBase_('Readiness', pageNo, totalPages, g);
  [['ontrack', 'On track', 36], ['attention', 'Needs attention', 256], ['priority', 'Priority support', 476]].forEach(function(col) {
    var list = g.levels[col[0]], x = col[2], paras = [];
    S.push(_slShape_('col-' + col[0], 'ROUND_RECTANGLE', x, 82, 208, 262, { fill: T.lilac, radius: 10 }));
    S.push(_slShape_('col-' + col[0] + '-head', 'ROUND_RECTANGLE', x + 12, 94, 184, 32, { fill: _slBadgeColor_(col[0]), radius: 16, valign: 'MIDDLE',
      paras: [_slPara_([_slRun_(col[1] + ' (' + list.length + ')', 13, true, T.white)], { align: 'CENTER' })] }));
    var shown = list.length > 7 ? list.slice(0, 6) : list;
    shown.forEach(function(a) {
      paras.push(_slPara_([_slRun_(_slFitText_(a.name, 184, 12, 1), 12, true, T.ink)]));
      paras.push(_slPara_([_slRun_(_slFitText_(a.reason, 184, 11, 1), 11, false, T.grey)]));
    });
    if (list.length > shown.length) paras.push(_slPara_([_slRun_('+ ' + (list.length - shown.length) + ' more', 11, true, T.purple)]));
    if (!list.length) paras.push(_slPara_([_slRun_('No agents', 11, false, T.grey)]));
    S.push(_slShape_('col-' + col[0] + '-list', 'TEXT_BOX', x + 12, 136, 184, 200, { paras: paras }));
  });
  S.push(_slText_('ready-note', 36, 348, 648, 28, [_slRun_('Badge: no KPI below threshold = On track, one = Needs attention, two or more = Priority support.' + (g.counts.nodata ? ' No data: ' + g.counts.nodata + ' agent' + (g.counts.nodata === 1 ? '' : 's') + '.' : ''), 11, false, T.grey)]));
  return S;
}

// 7. Next steps: up to six cards, one for each KPI with agents below its target, then the agents that need priority support.
function scorecardNextStepsShapes_(g, pageNo, totalPages) {
  var T = SCORECARD_THEME, S = _slGroupBase_('Next steps', pageNo, totalPages, g);
  g.nextSteps.slice(0, 6).forEach(function(step, i) {
    var x = 36 + (i % 3) * 220, y = 82 + Math.floor(i / 3) * 142;
    S.push(_slShape_('step-' + i, 'ROUND_RECTANGLE', x, y, 208, 130, { fill: T.lilac, radius: 10 }));
    S.push(_slShape_('step-' + i + '-dot', 'ELLIPSE', x + 12, y + 12, 28, 28, { fill: T.purple }));
    S.push(_slText_('step-' + i + '-num', x + 12, y + 12, 28, 28, [_slRun_(String(i + 1), 13, true, T.white)], { valign: 'MIDDLE', align: 'CENTER' }));
    S.push(_slText_('step-' + i + '-title', x + 48, y + 12, 148, 28, [_slRun_(_slFitText_(step.title, 148, 14, 1), 14, true, T.purple)], { valign: 'MIDDLE' }));
    S.push(_slText_('step-' + i + '-body', x + 12, y + 50, 184, 72, [_slRun_(_slFitText_(step.body, 184, 12, 4), 12, false, T.ink)]));
  });
  return S;
}

// The whole deck as layout shapes, in order: the cover, then one section for each LOB scorecard (its group slides, the slide of each of its
// agents, its next steps). built: buildAgentCards_ result (built.sections). Returns an array of slides (each an array of shapes);
// page numbers are 1-based. Without sections (no LOB scorecards) the whole deck is one section without a tag.
function scorecardDeckShapes_(built, label, filtersText, audits, thresholds) {
  var sections = (built.sections && built.sections.length) ? built.sections : [{ name: '', key: '', scorecard: null, cards: built.cards }];
  sections = sections.filter(function(s) { return s.cards.length; });
  var heatOf = function(s) { return Math.max(1, Math.ceil(s.cards.length / SCORECARD_HEATMAP_PER_SLIDE)); };
  var total = 1, p = 1, slides = [];
  sections.forEach(function(s) { total += 6 + heatOf(s) + s.cards.length; });
  var lobs = sections.length > 1 || (sections[0] && sections[0].name) ? sections.map(function(s) { return s.name + ' (' + s.cards.length + ')'; }).join(', ') : '';
  slides.push(scorecardCoverShapes_({ label: label, range: built.range, agents: built.cards.length, audits: audits, filters: filtersText, lobs: lobs, thresholds: thresholds }, total));
  sections.forEach(function(s) {
    var g = buildGroupSummary_(built, label, filtersText, audits, s), heatPages = heatOf(s);
    slides.push(scorecardSummaryShapes_(g, ++p, total));
    slides.push(scorecardResultsShapes_(g, ++p, total));
    slides.push(scorecardLandedShapes_(g, ++p, total));
    for (var h = 0; h < heatPages; h++) slides.push(scorecardHeatmapShapes_(s.cards, h, heatPages, ++p, total, g));
    slides.push(scorecardFindingsShapes_(g, ++p, total));
    slides.push(scorecardReadinessShapes_(g, ++p, total));
    s.cards.forEach(function(c) { slides.push(scorecardSlideShapes_(c, ++p, total)); });
    slides.push(scorecardNextStepsShapes_(g, ++p, total));
  });
  return slides;
}

// WCAG contrast ratio of two #RRGGBB colours.
function _slContrast_(a, b) {
  var lum = function(hex) {
    var c = [1, 3, 5].map(function(i) { var v = parseInt(hex.substr(i, 2), 16) / 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); });
    return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
  };
  var l1 = lum(a), l2 = lum(b), hi = Math.max(l1, l2), lo = Math.min(l1, l2);
  return (hi + 0.05) / (lo + 0.05);
}
