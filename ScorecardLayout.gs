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
  line: '#D8CCE6'          // thin tile border
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

function _slStatusColor_(status) {
  var T = SCORECARD_THEME;
  return status === 'target' ? T.green : (status === 'below' ? T.red : T.grey);
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
  S.push(_slShape_('sub', 'TEXT_BOX', 36, 68, 648, 20, { valign: 'MIDDLE', paras: [_slPara_([_slRun_(sub.join('  |  '), 12, false, T.grey)])] }));

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

// The cover slide: title, the period, the filters, and a legend of the targets so "Below threshold" is clear.
// summary: { label, range, agents, audits, filters, thresholds:[{label,target}] }
function scorecardCoverShapes_(summary, totalPages) {
  var T = SCORECARD_THEME, P = SCORECARD_PAGE, S = [];
  S.push(_slShape_('bg', 'RECTANGLE', 0, 0, P.w, P.h, { fill: T.lilac }));
  S.push(_slShape_('accent', 'RECTANGLE', 0, 0, P.w, 8, { fill: T.purple }));
  S.push(_slShape_('title', 'TEXT_BOX', 48, 70, 624, 56, { valign: 'MIDDLE', paras: [_slPara_([_slRun_('Audit Scorecards', 40, true, T.purple)])] }));
  var line2 = [summary.label, summary.range].filter(function(x) { return x; }).join('  |  ');
  S.push(_slShape_('subtitle', 'TEXT_BOX', 48, 130, 624, 28, { valign: 'MIDDLE', paras: [_slPara_([_slRun_(_slFitText_(line2, 624, 18, 1), 18, false, T.grey)])] }));
  S.push(_slShape_('counts', 'TEXT_BOX', 48, 162, 624, 24, { valign: 'MIDDLE', paras: [_slPara_([_slRun_(summary.agents + ' agent' + (summary.agents === 1 ? '' : 's') + '  |  ' + summary.audits + ' audit' + (summary.audits === 1 ? '' : 's'), 14, true, T.purple)])] }));
  if (summary.filters) S.push(_slShape_('filters', 'TEXT_BOX', 48, 190, 624, 34, { paras: [_slPara_([_slRun_('Filters: ', 12, true, T.purple), _slRun_(_slFitText_(summary.filters, 560, 12, 2), 12, false, T.ink)])] }));
  S.push(_slShape_('legend', 'ROUND_RECTANGLE', 48, 226, 624, 142, { fill: T.white, line: T.line, radius: 10, pad: 10,
    paras: [_slPara_([_slRun_('How to read the scorecard', 14, true, T.purple)])].concat(
      (summary.thresholds || []).map(function(t) { return _slPara_([_slRun_(t.label + ': ', 12, true, T.ink), _slRun_(t.target, 12, false, T.ink)], { spacing: 1.05 }); }),
      [_slPara_([_slRun_('Badge: no tile below threshold = On track, one = Needs attention, two or more = Priority support.', 11, false, T.grey)], { spacing: 1.05 })]
    ) }));
  S.push(_slShape_('footer', 'TEXT_BOX', 36, 378, 560, 18, { valign: 'MIDDLE', paras: [_slPara_([_slRun_('Audit Scorecard', SCORECARD_FOOTER_FONT, false, T.grey)])] }));
  S.push(_slShape_('page', 'TEXT_BOX', 620, 378, 64, 18, { valign: 'MIDDLE', paras: [_slPara_([_slRun_('1 / ' + totalPages, SCORECARD_FOOTER_FONT, false, T.grey)], { align: 'END' })] }));
  return S;
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
