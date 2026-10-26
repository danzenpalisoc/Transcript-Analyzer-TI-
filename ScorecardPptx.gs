/**
 * ScorecardPptx.gs
 * Writes the scorecard slides as a PowerPoint (.pptx) package: plain XML parts, no Google API. ScorecardSlides.gs zips the parts
 * and hands the file to Drive, which converts it to Google Slides (the same Drive conversion the PDF export already uses), so
 * nothing has to be switched on in Google Cloud and no new permission is needed.
 *
 * Input is the layout data from ScorecardLayout.gs (the same data the approved preview was drawn from). Everything the layout
 * sets is written exactly: position and size, fill, outline, corner radius, inner margin, vertical alignment, font, size, colour,
 * line spacing, bullets. Page: 16:9, 720 x 405 pt = 9,144,000 x 5,143,500 EMU (1 pt = 12,700 EMU).
 */

var _PX_NS_A_ = 'http://schemas.openxmlformats.org/drawingml/2006/main';
var _PX_NS_R_ = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
var _PX_NS_P_ = 'http://schemas.openxmlformats.org/presentationml/2006/main';
var _PX_REL_ = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
var _PX_DECL_ = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
var _PX_XMLNS_ = ' xmlns:a="' + _PX_NS_A_ + '" xmlns:r="' + _PX_NS_R_ + '" xmlns:p="' + _PX_NS_P_ + '"';

function _pxEmu_(pt) { return Math.round(pt * 12700); }
function _pxHex_(c) { return String(c).replace('#', '').toUpperCase(); }

// Text that is legal in XML 1.0: control characters and lone surrogates are dropped, line breaks and tabs become spaces,
// and the five special characters are escaped.
function _pxText_(s) {
  s = String(s === null || s === undefined ? '' : s).replace(/[\t\r\n]+/g, ' ');
  var out = '';
  for (var i = 0; i < s.length; i++) {
    var c = s.charCodeAt(i);
    if (c >= 0xD800 && c <= 0xDBFF) {
      var d = s.charCodeAt(i + 1);
      if (d >= 0xDC00 && d <= 0xDFFF) { out += s.charAt(i) + s.charAt(i + 1); i++; }
      continue;
    }
    if (c >= 0xDC00 && c <= 0xDFFF) continue;
    if (c >= 32 && c !== 0xFFFE && c !== 0xFFFF) out += s.charAt(i);
  }
  return out.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function _pxFill_(hex) { return '<a:solidFill><a:srgbClr val="' + _pxHex_(hex) + '"/></a:solidFill>'; }

function _pxPara_(p) {
  var font = SCORECARD_THEME.font;
  var algn = p.align === 'CENTER' ? 'ctr' : (p.align === 'END' ? 'r' : 'l');
  var bulletAttrs = p.bullet ? ' marL="' + _pxEmu_(14) + '" indent="-' + _pxEmu_(14) + '"' : '';
  var bullet = p.bullet
    ? '<a:buClr><a:srgbClr val="' + _pxHex_(p.runs[0].color) + '"/></a:buClr><a:buFont typeface="' + _pxText_(font) + '"/><a:buChar char="&#8226;"/>'
    : '<a:buNone/>';
  var pPr = '<a:pPr algn="' + algn + '"' + bulletAttrs + '><a:lnSpc><a:spcPct val="' + Math.round((p.spacing || 1) * 100000) + '"/></a:lnSpc>' +
    '<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft>' + bullet + '</a:pPr>';
  var biggest = 0;
  var runs = p.runs.map(function(r) {
    if (r.size > biggest) biggest = r.size;
    if (!String(r.t).length) return '';
    return '<a:r><a:rPr lang="en-US" sz="' + Math.round(r.size * 100) + '" b="' + (r.bold ? 1 : 0) + '" i="0" dirty="0">' + _pxFill_(r.color) +
      '<a:latin typeface="' + _pxText_(font) + '"/><a:cs typeface="' + _pxText_(font) + '"/></a:rPr><a:t>' + _pxText_(r.t) + '</a:t></a:r>';
  }).join('');
  return '<a:p>' + pPr + runs + '<a:endParaRPr lang="en-US" sz="' + Math.round((biggest || 12) * 100) + '" dirty="0"/></a:p>';
}

function _pxShape_(s, id) {
  var min = Math.min(s.w, s.h);
  var adj = Math.min(50000, Math.round((s.radius || 0) / min * 100000));
  var geom = (s.type === 'ROUND_RECTANGLE' && s.radius)
    ? '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ' + adj + '"/></a:avLst></a:prstGeom>'
    : (s.type === 'ELLIPSE' ? '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' : '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>');
  var fill = s.fill ? _pxFill_(s.fill) : '<a:noFill/>';
  var line = s.line ? '<a:ln w="9525">' + _pxFill_(s.line) + '</a:ln>' : '<a:ln><a:noFill/></a:ln>';
  var pad = _pxEmu_(s.pad || 0);
  var body = '';
  if (s.paras && s.paras.length) {
    body = '<p:txBody><a:bodyPr wrap="square" lIns="' + pad + '" tIns="' + pad + '" rIns="' + pad + '" bIns="' + pad + '" rtlCol="0" anchor="' +
      (s.valign === 'MIDDLE' ? 'ctr' : 't') + '"><a:noAutofit/></a:bodyPr><a:lstStyle/>' + s.paras.map(_pxPara_).join('') + '</p:txBody>';
  }
  return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + _pxText_(s.id) + '"/><p:cNvSpPr' + (s.type === 'TEXT_BOX' ? ' txBox="1"' : '') + '/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="' + _pxEmu_(s.x) + '" y="' + _pxEmu_(s.y) + '"/><a:ext cx="' + _pxEmu_(s.w) + '" cy="' + _pxEmu_(s.h) + '"/></a:xfrm>' +
    geom + fill + line + '</p:spPr>' + body + '</p:sp>';
}

var _PX_GROUP_ = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>' +
  '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

function _pxSlide_(shapes) {
  var bg = null, rest = [];
  shapes.forEach(function(s) { if (s.id === 'bg') bg = s; else rest.push(s); });
  var bgXml = bg ? '<p:bg><p:bgPr>' + _pxFill_(bg.fill) + '<a:effectLst/></p:bgPr></p:bg>' : '';
  return _PX_DECL_ + '<p:sld' + _PX_XMLNS_ + '><p:cSld>' + bgXml + '<p:spTree>' + _PX_GROUP_ +
    rest.map(function(s, i) { return _pxShape_(s, i + 2); }).join('') + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
}

function _pxRels_(items) {
  return _PX_DECL_ + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    items.map(function(r) { return '<Relationship Id="' + r.id + '" Type="' + _PX_REL_ + r.type + '" Target="' + r.target + '"/>'; }).join('') + '</Relationships>';
}

function _pxTheme_() {
  var T = SCORECARD_THEME, c = function(n, hex) { return '<a:' + n + '><a:srgbClr val="' + _pxHex_(hex) + '"/></a:' + n + '>'; };
  var font = function(n) { return '<a:' + n + '><a:latin typeface="' + _pxText_(T.font) + '"/><a:ea typeface=""/><a:cs typeface=""/></a:' + n + '>'; };
  var three = function(x) { return x + x + x; };
  return _PX_DECL_ + '<a:theme xmlns:a="' + _PX_NS_A_ + '" name="Scorecard"><a:themeElements><a:clrScheme name="Scorecard">' +
    c('dk1', T.ink) + c('lt1', T.white) + c('dk2', T.purple) + c('lt2', T.lilac) + c('accent1', T.purple) + c('accent2', T.green) + c('accent3', T.red) +
    c('accent4', T.amber) + c('accent5', T.grey) + c('accent6', T.brightGreen) + c('hlink', T.purple) + c('folHlink', T.grey) +
    '</a:clrScheme><a:fontScheme name="Scorecard">' + font('majorFont') + font('minorFont') + '</a:fontScheme><a:fmtScheme name="Scorecard">' +
    '<a:fillStyleLst>' + three('<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>') + '</a:fillStyleLst>' +
    '<a:lnStyleLst>' + three('<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>') + '</a:lnStyleLst>' +
    '<a:effectStyleLst>' + three('<a:effectStyle><a:effectLst/></a:effectStyle>') + '</a:effectStyleLst>' +
    '<a:bgFillStyleLst>' + three('<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>') + '</a:bgFillStyleLst>' +
    '</a:fmtScheme></a:themeElements></a:theme>';
}

function _pxMaster_() {
  var lvl = '<a:lvl1pPr algn="l"><a:defRPr sz="1800"/></a:lvl1pPr>';
  return _PX_DECL_ + '<p:sldMaster' + _PX_XMLNS_ + '><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>' + _PX_GROUP_ + '</p:spTree></p:cSld>' +
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
    '<p:txStyles><p:titleStyle>' + lvl + '</p:titleStyle><p:bodyStyle>' + lvl + '</p:bodyStyle><p:otherStyle>' + lvl + '</p:otherStyle></p:txStyles></p:sldMaster>';
}

function _pxLayout_() {
  return _PX_DECL_ + '<p:sldLayout' + _PX_XMLNS_ + ' type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>' + _PX_GROUP_ +
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
}

var _PX_MIME_ = 'application/vnd.openxmlformats-officedocument.presentationml.';

// The parts of the package, in the order they are zipped. slides: an array of layout-shape arrays. Returns [{ path, xml }].
function buildPptxParts_(slides) {
  var n = slides.length, parts = [];
  var slideOverrides = slides.map(function(s, i) {
    return '<Override PartName="/ppt/slides/slide' + (i + 1) + '.xml" ContentType="' + _PX_MIME_ + 'slide+xml"/>';
  }).join('');
  parts.push({ path: '[Content_Types].xml', xml: _PX_DECL_ + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/ppt/presentation.xml" ContentType="' + _PX_MIME_ + 'presentation.main+xml"/>' +
    '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="' + _PX_MIME_ + 'slideMaster+xml"/>' +
    '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="' + _PX_MIME_ + 'slideLayout+xml"/>' +
    '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' + slideOverrides + '</Types>' });
  parts.push({ path: '_rels/.rels', xml: _pxRels_([{ id: 'rId1', type: 'officeDocument', target: 'ppt/presentation.xml' }]) });

  var sldIds = slides.map(function(s, i) { return '<p:sldId id="' + (256 + i) + '" r:id="rId' + (i + 2) + '"/>'; }).join('');
  parts.push({ path: 'ppt/presentation.xml', xml: _PX_DECL_ + '<p:presentation' + _PX_XMLNS_ + '><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
    '<p:sldIdLst>' + sldIds + '</p:sldIdLst><p:sldSz cx="' + _pxEmu_(SCORECARD_PAGE.w) + '" cy="' + _pxEmu_(SCORECARD_PAGE.h) + '"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>' });
  var presRels = [{ id: 'rId1', type: 'slideMaster', target: 'slideMasters/slideMaster1.xml' }];
  slides.forEach(function(s, i) { presRels.push({ id: 'rId' + (i + 2), type: 'slide', target: 'slides/slide' + (i + 1) + '.xml' }); });
  presRels.push({ id: 'rId' + (n + 2), type: 'theme', target: 'theme/theme1.xml' });
  parts.push({ path: 'ppt/_rels/presentation.xml.rels', xml: _pxRels_(presRels) });

  parts.push({ path: 'ppt/slideMasters/slideMaster1.xml', xml: _pxMaster_() });
  parts.push({ path: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', xml: _pxRels_([
    { id: 'rId1', type: 'slideLayout', target: '../slideLayouts/slideLayout1.xml' }, { id: 'rId2', type: 'theme', target: '../theme/theme1.xml' }]) });
  parts.push({ path: 'ppt/slideLayouts/slideLayout1.xml', xml: _pxLayout_() });
  parts.push({ path: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', xml: _pxRels_([{ id: 'rId1', type: 'slideMaster', target: '../slideMasters/slideMaster1.xml' }]) });
  parts.push({ path: 'ppt/theme/theme1.xml', xml: _pxTheme_() });
  slides.forEach(function(shapes, i) {
    parts.push({ path: 'ppt/slides/slide' + (i + 1) + '.xml', xml: _pxSlide_(shapes) });
    parts.push({ path: 'ppt/slides/_rels/slide' + (i + 1) + '.xml.rels', xml: _pxRels_([{ id: 'rId1', type: 'slideLayout', target: '../slideLayouts/slideLayout1.xml' }]) });
  });
  return parts;
}
