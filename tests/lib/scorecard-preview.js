/**
 * Draws the scorecard layout (ScorecardLayout.gs shapes) as an HTML page, so the look can be seen in a browser
 * before it is ever sent to Google Slides. It reads the SAME shapes the real Slides builder reads: positions, sizes,
 * colours and text come straight from the layout data, in points, scaled by `k` pixels per point.
 */
function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

function shapeHtml(s, k, font) {
  const style = [
    'position:absolute', 'box-sizing:border-box',
    'left:' + s.x * k + 'px', 'top:' + s.y * k + 'px', 'width:' + s.w * k + 'px', 'height:' + s.h * k + 'px',
    s.fill ? 'background:' + s.fill : '',
    s.line ? 'border:' + Math.max(1, k * 0.75) + 'px solid ' + s.line : '',
    s.radius ? 'border-radius:' + s.radius * k + 'px' : '',
    'padding:' + s.pad * k + 'px', 'overflow:visible',
    'display:flex', 'flex-direction:column',
    'justify-content:' + (s.valign === 'MIDDLE' ? 'center' : 'flex-start'),
    'font-family:' + font + ', Helvetica, sans-serif',
  ].filter(Boolean).join(';');
  const paras = s.paras.map(p => {
    const size = Math.max.apply(null, p.runs.map(r => r.size));
    const align = p.align === 'CENTER' ? 'center' : (p.align === 'END' ? 'right' : 'left');
    const runs = p.runs.map(r => '<span style="font-size:' + r.size * k + 'px;font-weight:' + (r.bold ? 700 : 400) + ';color:' + r.color + '">' + esc(r.t) + '</span>').join('');
    const line = 'line-height:' + (1.2 * p.spacing) + ';text-align:' + align + ';';
    if (p.bullet) {
      const col = p.runs[0].color;
      return '<div style="display:flex;' + line + '"><span style="flex:0 0 ' + 14 * k + 'px;font-size:' + size * k + 'px;color:' + col + '">&#8226;</span><span style="flex:1">' + runs + '</span></div>';
    }
    return '<div style="' + line + '">' + runs + '</div>';
  }).join('');
  return '<div data-id="' + s.id + '" style="' + style + '">' + paras + '</div>';
}

// slides: array of shape arrays. Returns a full HTML page, each slide in its own 16:9 frame.
function renderDeckHtml(slides, theme, page, k, titles) {
  k = k || 1.6;
  const frames = slides.map((shapes, i) =>
    '<div class="label">' + esc((titles && titles[i]) || 'Slide ' + (i + 1)) + '</div>' +
    '<div class="slide" style="width:' + page.w * k + 'px;height:' + page.h * k + 'px">' + shapes.map(s => shapeHtml(s, k, theme.font)).join('') + '</div>').join('\n');
  return '<!doctype html><html><head><meta charset="utf-8"><title>Scorecard preview</title><style>' +
    'body{margin:0;padding:24px;background:#E6E1EE;font-family:Arial,sans-serif}' +
    '.label{font:12px Arial;color:#54595F;margin:18px 0 6px}' +
    '.slide{position:relative;background:#fff;box-shadow:0 2px 10px rgba(0,0,0,.25);overflow:hidden}' +
    '</style></head><body>' + frames + '</body></html>';
}

module.exports = { renderDeckHtml };
