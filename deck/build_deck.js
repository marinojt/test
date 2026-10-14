// Optum LS Conversational AI deck — Spec A (Optum Brand). Slide copy is verbatim from the locked draft.
const path = require('path');
const pptxgen = require('pptxgenjs');
const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const sharp = require('sharp');
const Md = require('react-icons/md');
const { applyTheme } = require(process.env.PPTX_SKILL + '/scripts/apply_theme.js');

const HEX = { navy: '002677', orange: 'FF612B', body: '4B4D4F', muted: '7A7C7E', canvas: 'FAF8F2', neutral: 'D9D4C7', tint: 'E6ECF5', white: 'FFFFFF' };
const THEME = {
  name: 'Optum LS Strategy', headFontFace: 'Arial', bodyFontFace: 'Arial',
  colors: { dk1: HEX.body, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.canvas, accent1: HEX.navy, accent2: HEX.orange, accent3: HEX.neutral, accent4: HEX.muted, accent5: HEX.tint, accent6: '196ECF', hlink: '196ECF', folHlink: HEX.muted },
};

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = 'Optum Life Sciences — Conversational AI';
const C = pres.SchemeColor;
const COL = { navy: C.accent1, orange: C.accent2, body: C.text1, muted: C.accent4, neutral: C.accent3, tint: C.accent5, white: C.background1, canvas: C.background2 };
const L = 0.6, W = 12.13;

pres.defineSlideMaster({
  title: 'CONTENT', background: { color: COL.canvas },
  objects: [
    { line: { x: L, y: 6.82, w: W, h: 0, line: { color: COL.neutral, width: 0.75 } } },
    { text: { text: 'Candidate work sample — outside-in research', options: { x: L, y: 6.9, w: 8, h: 0.22, fontSize: 9, color: COL.muted, margin: 0 } } },
    { placeholder: { options: { name: 'kicker', type: 'body', x: L, y: 0.42, w: 9, h: 0.24, fontSize: 10, bold: true, color: COL.orange, charSpacing: 1, margin: 0, valign: 'top' }, text: '' } },
    { placeholder: { options: { name: 'title', type: 'title', x: L, y: 0.7, w: W, h: 0.82, fontSize: 22, bold: true, color: COL.navy, margin: 0, valign: 'top', align: 'left' }, text: '' } },
  ],
  slideNumber: { x: 11.73, y: 6.9, w: 1.0, h: 0.22, fontSize: 9, color: COL.muted, align: 'right', margin: 0 },
});

// helpers
const tx = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: 'top', fontSize: 11, color: COL.body }, o));
const label = (s, text, x, y, w, color) => tx(s, text, { x, y, w, h: 0.22, fontSize: 10, bold: true, color: color || COL.orange, charSpacing: 1, objectName: 'label' });
const card = (s, x, y, w, h, name) => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: COL.white }, line: { type: 'none' }, objectName: name || 'card',
  shadow: { type: 'outer', color: '000000', opacity: 0.1, blur: 8, offset: 2, angle: 90 } });
const iconCache = {};
async function icon(name) {
  if (!iconCache[name]) {
    const svg = renderToStaticMarkup(React.createElement(Md[name], { color: '#' + HEX.navy, size: 256 }));
    iconCache[name] = 'image/png;base64,' + (await sharp(Buffer.from(svg)).png().toBuffer()).toString('base64');
  }
  return iconCache[name];
}
async function iconDot(s, name, x, y, d) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: COL.tint }, line: { type: 'none' }, objectName: 'icon circle' });
  s.addImage({ data: await icon(name), x: x + d * 0.22, y: y + d * 0.22, w: d * 0.56, h: d * 0.56, objectName: 'icon ' + name });
}
const title = (s, t) => s.addText(t.split('|').map((p, i) => ({ text: p, options: i ? { softBreakBefore: true } : {} })), { placeholder: 'title' });
const brk = t => t.split('|').map((p, i) => ({ text: p, options: i ? { softBreakBefore: true } : {} }));
const B = (text, extra) => ({ text, options: Object.assign({ bold: true }, extra) });
const R = (text, extra) => ({ text, options: Object.assign({}, extra) });

// Workflow chip — the five decision workflows are the deck's core taxonomy; same chip everywhere.
const chipW = (t, fs) => 0.24 + t.length * fs * 0.0078;
function chip(s, t, x, y, o = {}) {
  const fs = o.fontSize || 10, h = o.h || 0.32, w = o.w || chipW(t, fs);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: h / 2, fill: { color: COL.tint }, line: { color: COL.navy, width: 0.75 }, objectName: 'workflow chip ' + t });
  tx(s, t, { x, y, w, h, fontSize: fs, bold: true, color: COL.navy, align: 'center', valign: 'middle' });
  return w;
}
const cardLabel = (s, t, x, y, w) => tx(s, t, { x, y, w, h: Math.ceil(t.length * 0.118 / w) * 0.26, fontSize: 13, bold: true, color: COL.orange, charSpacing: 1.5, valign: 'middle', objectName: 'card label' });

// ---------- SLIDE 1 — EXECUTIVE SUMMARY ----------
async function slide1() {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Executive summary' });
  s.addText('EXECUTIVE SUMMARY', { placeholder: 'kicker' });
  title(s, 'Conversational AI is being marketed across life sciences workflows, creating|new ways to access information, analytics and decision support');

  // Outside-in research lens (full width) with the five workflows as chips
  const ly = 1.68, lh = 1.68;
  card(s, L, ly, W, lh, 'lens card');
  await iconDot(s, 'MdTravelExplore', L + 0.25, ly + 0.22, 0.52);
  cardLabel(s, 'OUTSIDE-IN RESEARCH LENS', L + 0.95, ly + 0.2, 6);
  tx(s, [B('This study is an assessment of '), B('publicly-marketed', { color: COL.orange }), B(' conversational AI products across five Optum Life Sciences-relevant business domains:')], { x: L + 0.95, y: ly + 0.56, w: 10.9, h: 0.22, fontSize: 11 });
  let x = L + 0.95; const chy = ly + 0.86;
  for (const [i, t] of WF.entries()) {
    if (i === 4) { tx(s, 'and', { x, y: chy, w: 0.34, h: 0.27, fontSize: 10, color: COL.muted, align: 'center', valign: 'middle' }); x += 0.4; }
    x += chip(s, t, x, chy, { h: 0.27, fontSize: 9.5 }) + 0.1;
  }
  tx(s, 'The U.S. competitive landscape is the primary focus; selected international offerings provide context; and Optum LS is the strategic reference point.', { x: L + 0.95, y: ly + 1.22, w: 10.9, h: 0.22, fontSize: 11 });

  const cy = 3.52, ch = 3.1, cw = 3.91, gap = 0.2;
  const cx = [L, L + cw + gap, L + 2 * (cw + gap)];
  const head = async (i, ic, t) => { card(s, cx[i], cy, cw, ch, t.toLowerCase() + ' card'); await iconDot(s, ic, cx[i] + 0.25, cy + 0.22, 0.5); cardLabel(s, t, cx[i] + 0.88, cy + 0.32, cw - 1.0); };

  // Key findings — number and sentence share a top line and a fixed gutter
  await head(0, 'MdInsights', 'KEY FINDINGS');
  const kf = [['19', [R('The U.S. screen identified '), B('19 conversational AI offerings'), R(' across five business domains, '), B('concentrated in RWE and Commercial Intelligence'), R('.')]],
              ['10', [R('The selected international comparison identified '), B('10 conversational AI offerings,'), R(' providing additional examples of how vendors position conversational interaction.')]]];
  kf.forEach(([n, t], i) => {
    const y = cy + 1.0 + i * 1.08;
    if (i) s.addShape(pres.shapes.LINE, { x: cx[0] + 0.25, y: y - 0.12, w: cw - 0.5, h: 0, line: { color: COL.neutral, width: 0.75 } });
    tx(s, n, { x: cx[0] + 0.25, y: y - 0.07, w: 0.68, h: 0.55, fontSize: 34, bold: true, color: COL.orange, objectName: 'hero ' + n });
    tx(s, t, { x: cx[0] + 0.98, y, w: cw - 1.2, h: 0.82, fontSize: 10.5 });
  });

  // Optum positioning — three sentences as three paragraphs
  await head(1, 'MdOutlineLayers', 'OPTUM POSITIONING');
  tx(s, [
    R('Optum LS publicly emphasizes linked healthcare data, AI-enabled research tools and scientific expertise.', { breakLine: true }),
    R('Several competitors prominently market conversational interfaces as the entry point to comparable customer tasks.', { breakLine: true }),
    B('This reflects differences in public positioning, not necessarily in underlying functionality.', { color: COL.navy }),
  ], { x: cx[1] + 0.25, y: cy + 0.95, w: cw - 0.5, h: 1.8, fontSize: 11, paraSpaceAfter: 9 });

  // Strategic implication
  await head(2, 'MdOutlineLightbulb', 'STRATEGIC IMPLICATION');
  tx(s, 'The competitive significance of conversational AI depends on the customer task, underlying data and reliability of the results.',
    { x: cx[2] + 0.25, y: cy + 0.95, w: cw - 0.5, h: 0.75, fontSize: 11 });
  tx(s, "Where could conversational interaction add value to Optum LS's existing offerings?",
    { x: cx[2] + 0.25, y: cy + 1.75, w: cw - 0.5, h: 0.95, fontSize: 14, bold: true, color: COL.navy, objectName: 'strategic question' });

}

// ---------- SLIDE 3 — U.S. COMPETITIVE LANDSCAPE ----------
async function slide3() {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'U.S. competitive landscape' });
  s.addText('U.S. COMPETITIVE LANDSCAPE', { placeholder: 'kicker' });
  title(s, 'Among U.S. competitors, conversational AI is most prominent in RWE; Optum LS has relevant offerings in all five domains');

  // U.S. competitive screen — KPI tiles linked by arrows
  const tiles = [['45', 'competitors screened', COL.navy], ['17', 'competitors who market conversational AI', COL.orange]];
  const aw = 0.8, tw = (W - aw) / 2, ty = 1.66, th = 0.6;
  tiles.forEach(([n, t, c], i) => {
    const x = L + i * (tw + aw), nw = 0.46, tww = t.length * 0.085 + 0.1, gx = x + (tw - (nw + 0.12 + tww)) / 2;
    s.addShape(pres.shapes.RECTANGLE, { x, y: ty, w: tw, h: th, fill: { color: COL.tint }, line: { type: 'none' }, objectName: 'kpi tile ' + n });
    tx(s, n, { x: gx, y: ty, w: nw, h: th, fontSize: 26, bold: true, color: c, align: 'right', valign: 'middle' });
    tx(s, t, { x: gx + nw + 0.12, y: ty, w: tww, h: th, fontSize: 11, bold: true, color: COL.navy, valign: 'middle' });
    if (i === 0) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + tw + 0.2, y: ty + th / 2 - 0.13, w: aw - 0.4, h: 0.26, fill: { color: COL.muted }, line: { type: 'none' }, objectName: 'flow arrow' });
  });

  // Competitor coverage and Optum LS positioning — mirrored chart + table
  label(s, 'COMPETITOR COVERAGE AND OPTUM LS POSITIONING', L, 2.44, 7);
  const rows = [
    ['RWE', 9, 34, 'Komodo Marmot and MapAI; Truveta Intelligence; Flatiron Telescope; Tempus Lens; and more', 'Analytics on Demand; epidemiology research'],
    ['Commercial Intelligence', 3, 23, 'IQVIA Field Force Agent and GMI Agent; ZS Conversational AI for Commercial; Trinity InsightsEDGE Digital Twins', 'Market Clarity: portfolio and patient/prescriber analytics'],
    ['Medical / Regulatory Research', 3, 18, 'IQVIA Med-R1; ZS Conversational AI for Medical; Veeva Vault AI Chat', 'Postapproval safety studies; Market Clarity medical affairs insights'],
    ['Clinical Trials', 1, 25, 'TriNetX Terrain AI Assistant', 'Prospector; Digital Research Network; True Source'],
    ['Market Access', 1, 24, 'EVERSANA Ask NAVLIN', 'HEOR/value evidence; Evidence Engine'],
  ];
  const hy = 2.7, hh = 0.5, ry = hy + hh, rh = 0.51;
  const cX = { bars: L, wf: 3.8, rep: 6.0, opt: 9.3, chat: 11.75 }, cW = { bars: 3.15, wf: 2.15, rep: 3.15, opt: 2.3, chat: 0.98 };
  s.addShape(pres.shapes.RECTANGLE, { x: L, y: hy, w: W, h: hh, fill: { color: COL.navy }, line: { type: 'none' }, objectName: 'table header' });
  const hdr = { y: hy, h: hh, fontSize: 9, bold: true, color: COL.white, valign: 'middle', align: 'center' };
  tx(s, [R('Competitors marketing conversational AI /', { breakLine: true }), R('Total comparable competitors')], { ...hdr, x: cX.bars + 0.45, w: 2.65 });
  tx(s, 'Domain', { ...hdr, x: cX.wf, w: cW.wf });
  tx(s, 'Conversational AI offerings', { ...hdr, x: cX.rep, w: cW.rep });
  tx(s, [R('Relevant Optum LS', { breakLine: true }), R('offerings')], { ...hdr, x: cX.opt, w: cW.opt });
  tx(s, [R('Optum LS', { breakLine: true }), R('chat claim')], { ...hdr, x: cX.chat - 0.1, w: cW.chat + 0.1 });

  // Native stacked bar chart extending left toward the workflow spine
  const chartX = cX.bars + 0.68, chartW = 2.42, maxV = 34;
  s.addChart(pres.charts.BAR, [
    { name: 'Marketing qualifying chat', labels: rows.map(r => r[0]), values: rows.map(r => r[1]) },
    { name: 'Remaining comparable', labels: rows.map(r => r[0]), values: rows.map(r => r[2] - r[1]) },
  ], {
    x: chartX, y: ry, w: chartW, h: rh * rows.length, barDir: 'bar', barGrouping: 'stacked', barGapWidthPct: 85,
    chartColors: [HEX.navy, HEX.neutral], valAxisMinVal: 0, valAxisMaxVal: maxV, valAxisOrientation: 'maxMin', catAxisOrientation: 'maxMin',
    valAxisHidden: true, catAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
    showLegend: false, showValue: false, layout: { x: 0, y: 0, w: 1, h: 1 }, objectName: 'competitor coverage chart',
  });
  rows.forEach(([wf, n, d, rep, opt], i) => {
    const y = ry + i * rh;
    const barLeft = chartX + chartW * (1 - d / maxV);
    tx(s, [B(`${n}`, { color: COL.orange }), R(` / ${d}`)], { x: barLeft - 0.66, y, w: 0.6, h: rh, fontSize: 10, align: 'right', valign: 'middle', objectName: 'fraction ' + wf });
    chip(s, wf, cX.wf + 0.08, y + (rh - 0.34) / 2, { w: cW.wf - 0.16, h: 0.34, fontSize: 9 });
    tx(s, rep, { x: cX.rep + 0.1, y, w: cW.rep - 0.2, h: rh, fontSize: 9, valign: 'middle' });
    tx(s, opt, { x: cX.opt + 0.1, y, w: cW.opt - 0.2, h: rh, fontSize: 9, valign: 'middle' });
    s.addShape(pres.shapes.OVAL, { x: cX.chat + cW.chat / 2 - 0.09, y: y + rh / 2 - 0.09, w: 0.18, h: 0.18, fill: { type: 'none' }, line: { color: COL.muted, width: 1.25 }, objectName: 'chat status ' + wf });
    s.addShape(pres.shapes.LINE, { x: L, y: y + rh, w: W, h: 0, line: { color: COL.neutral, width: 0.75 } });
  });
  const ty2 = ry + rh * rows.length + 0.07;
  tx(s, 'Scope note: Offerings are organized by their primary customer decision workflow; supporting data and analytics capabilities may serve more than one workflow.',
    { x: L, y: ty2, w: 9.6, h: 0.2, fontSize: 8, bold: true, color: COL.muted, objectName: 'scope note' });
  tx(s, '○  Not publicly established in this review', { x: 10.3, y: ty2, w: W + L - 10.3, h: 0.2, fontSize: 8, italic: true, color: COL.muted, align: 'right', objectName: 'chat claim legend' });

  // Key takeaway — full-width strip
  const ky = 6.12, kh = 0.6;
  card(s, L, ky, W, kh, 'key takeaway card');
  tx(s, 'KEY TAKEAWAY', { x: L + 0.2, y: ky, w: 2.1, h: kh, fontSize: 13, bold: true, color: COL.orange, charSpacing: 1.5, valign: 'middle', objectName: 'label' });
  tx(s, [B('A publicly-marketed customer-facing conversational interface was not established in the reviewed materials for Optum LS', { breakLine: true }), B('across the five relevant business domains.')],
    { x: L + 2.6, y: ky, w: W - 2.8, h: kh, fontSize: 10.5, valign: 'middle' });
}

// ---------- shared helpers for slides 2, 4–11 ----------
const WF = ['RWE', 'Commercial Intelligence', 'Medical / Regulatory Research', 'Clinical Trials', 'Market Access'];
const COUNTRY = /(\((?:Israel|Finland|China|Australia|Argentina|France|Spain)\))/;
// "**bold**" spans; "(Country)" tags flag international examples in orange
function rich(str, base = {}, plain = false) {
  const out = [];
  str.split('**').forEach((seg, i) => {
    seg.split(COUNTRY).forEach(p => {
      if (!p) return;
      const c = !plain && COUNTRY.test(p) && /^\(.*\)$/.test(p);
      out.push({ text: p, options: Object.assign({}, base, (i % 2 === 1 || c) ? { bold: true } : {}, c ? { color: COL.orange } : {}) });
    });
  });
  return out;
}
function paras(list, base = {}, gap = true, plain = false) {
  const runs = [];
  list.forEach((str, i) => {
    const r = rich(str, base, plain);
    if (i < list.length - 1) { r[r.length - 1].options.breakLine = true; }
    if (gap && i > 0) r[0].options.paraSpaceBefore = 4;
    runs.push(...r);
  });
  return runs;
}
// Workflow tracker: the five-workflow taxonomy, active workflow(s) filled navy
function tracker(s, active) {
  const fs = 7.5, h = 0.26, g = 0.07, ws = WF.map(t => 0.24 + t.length * 0.061);
  let x = L + W - ws.reduce((a, b) => a + b, 0) - g * 4;
  WF.forEach((t, i) => {
    const on = active.includes(t);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 0.4, w: ws[i], h, rectRadius: h / 2, fill: { color: on ? COL.navy : COL.tint }, line: { type: 'none' }, objectName: 'tracker ' + t });
    tx(s, t, { x, y: 0.4, w: ws[i], h, fontSize: fs, bold: true, color: on ? COL.white : COL.muted, align: 'center', valign: 'middle' });
    x += ws[i] + g;
  });
}
const lines = (str, cpl) => str.split('\n').reduce((n, p) => n + Math.max(1, Math.ceil(p.replace(/\*\*/g, '').length / cpl)), 0);
function footnote(s, runs, y) {
  tx(s, runs, { x: L, y: y || 6.3, w: W, h: 0.42, fontSize: 9, italic: true, color: COL.muted, objectName: 'evidence boundary' });
}
const ebLabel = text => [B('EVIDENCE BOUNDARY  ', { italic: false, charSpacing: 1 }), R(text)];

// Three-column task table (Slides 5–10). rowHeights() measures without drawing.
const TT_COLW = [2.5, 6.3, 3.33];
const estH = (str, w, f) => lines(str, Math.floor((w - 0.24) / (f * 0.0062))) * f * 1.22 / 72;
function rowHeights(rows, opt = {}) {
  const colW = opt.colW || TT_COLW, fs = opt.fontSize || 9.5;
  return rows.map(([task, comp, optum]) => Math.max(estH(task.replace('|', '\n'), colW[0], 11.5), comp.reduce((n, c) => n + estH(c, colW[1], fs), 0) + (comp.length - 1) * 0.06, estH(optum, colW[2], fs)) + 0.2 + (opt.rowPad || 0));
}
const TT_HEAD = 0.36;
function taskTable(s, y, rows, opt = {}) {
  const colW = opt.colW || TT_COLW, fs = opt.fontSize || 9.5;
  const none = { type: 'none' }, hair = { type: 'solid', pt: 0.75, color: HEX.neutral }, white = { type: 'solid', pt: 2, color: HEX.white };
  const hd = { fill: { color: COL.navy }, color: COL.white, bold: true, fontSize: 9, valign: 'middle', charSpacing: 1, border: [none, none, none, none] };
  const head = (opt.head || ['CUSTOMER TASK', 'CONVERSATIONAL AI OFFERINGS', 'RELEVANT OPTUM LS OFFERINGS']).map(t => ({ text: t, options: Object.assign({}, hd, opt.headAlign ? { align: opt.headAlign } : {}) }));
  const rowH = [opt.headH || TT_HEAD, ...rowHeights(rows, opt)];
  const body = rows.map(([task, comp, optum], i) => {
    const bot = opt.groupAfter === i ? { type: 'solid', pt: 1.5, color: HEX.navy } : hair;
    return [
      { text: brk(task), options: { fill: { color: opt.taskFill ? opt.taskFill(i) : COL.tint }, color: opt.taskColor ? opt.taskColor(i) : COL.navy, bold: true, fontSize: 10.5, valign: 'middle', align: opt.taskAlign || 'left', border: [white, white, white, none] } },
      { text: paras(comp, {}, true, !!opt.plainCountry), options: { fontSize: fs, color: COL.body, valign: 'middle', border: [none, none, bot, none] } },
      { text: rich(optum), options: { fontSize: fs, color: COL.body, valign: 'middle', border: [none, none, bot, none] } },
    ];
  });
  s.addTable([head, ...body], { x: L, y, w: W, colW, rowH, margin: [0.07, 0.12, 0.07, 0.12], objectName: 'task table' });
  return y + rowH.reduce((a, b) => a + b, 0);
}
// Cards with an orange label + text; the label/text group is vertically centered in each card.
const runText = t => (typeof t === 'string' ? [t] : t.reduce((ps, r) => { ps[ps.length - 1] += r.text; if (r.options && r.options.breakLine) ps.push(''); return ps; }, ['']));
function cardMetrics(items, w, fs) {
  return items.map(([lab, runs]) => {
    const lh = Math.ceil(lab.length * 0.118 / (w - 0.44)) * 0.25 + 0.03;
    const cpl = Math.floor((w - 0.44) / (fs * 0.0075));
    const ps = runText(runs);
    const th = ps.reduce((n, p) => n + Math.max(1, Math.ceil(p.length / cpl)), 0) * fs * 1.22 / 72 + (ps.length - 1) * 0.08;
    return { lh, th, group: lh + 0.08 + th };
  });
}
const cardWidth = n => (W - 0.2 * (n - 1)) / n;
const cardsHeight = (items, fs = 10.5) => Math.max(...cardMetrics(items, cardWidth(items.length), fs).map(m => m.group)) + 0.4;
function cardRow(s, y, bottom, items, fs = 10.5) {
  const w = cardWidth(items.length), h = bottom - y, ms = cardMetrics(items, w, fs);
  const gmax = Math.max(...ms.map(m => m.group)); // shared top line across cards
  items.forEach(([lab, runs], i) => {
    const x = L + i * (w + 0.2), m = ms[i], gy = y + (h - gmax) / 2;
    card(s, x, y, w, h, lab.toLowerCase() + ' card');
    tx(s, lab, { x: x + 0.22, y: gy, w: w - 0.44, h: m.lh, fontSize: 13, bold: true, color: COL.orange, charSpacing: 1.5, objectName: 'label' });
    const body = typeof runs === 'string' ? [B(runs)] : runs.map(r => ({ text: r.text, options: Object.assign({}, r.options, { bold: true }) }));
    tx(s, body, { x: x + 0.22, y: gy + m.lh + 0.08, w: w - 0.44, h: m.th + 0.06, fontSize: fs });
  });
}
const sectionLabel = (s, t, y, right) => {
  label(s, t, L, y, right ? 6.0 : W);
  if (right) tx(s, right, { x: L + 6.1, y, w: W - 6.1, h: 0.22, fontSize: 10, bold: true, color: COL.navy, align: 'right', objectName: 'screen counts' });
};

// ---------- SLIDE 2 — RESEARCH CATALYST ----------
async function slide2() {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Research catalyst' });
  s.addText('RESEARCH CATALYST', { placeholder: 'kicker' });
  title(s, 'Recent product launches in RWE prompted a broader examination|of conversational AI across life sciences');

  label(s, 'SELECTED PUBLIC PRODUCT ANNOUNCEMENTS', L, 1.68, 8);
  const ms = [
    ['2024', 'Earlier precedents', ['Tempus’s **Tempus One / Lens** — AI-assisted research and cohort exploration (February)', 'Komodo Health’s **MapAI** — conversational healthcare analytics (September)']],
    ['August 2025', 'Platform expansion', ['Komodo Health’s **Marmot** — AI-powered healthcare analytics and research workflows']],
    ['April–May 2026', 'Research-interface launches', ['Truveta’s **Truveta Intelligence** — natural-language RWE research questions', 'Flatiron Health’s **Telescope** — conversational oncology cohort exploration']],
    ['September 2026', 'New releases', ['Verana Health’s **Agent Claire** — conversational specialty-data research', 'Norstella’s **Atlas Patient Analytics** — agentic, managed cohort analytics']],
  ];
  const colw = W / 4, ly = 2.68;
  s.addShape(pres.shapes.LINE, { x: L, y: ly, w: W, h: 0, line: { color: COL.navy, width: 1.5 }, objectName: 'timeline' });
  ms.forEach(([date, sub, items], i) => {
    const x = L + i * colw;
    tx(s, date, { x, y: 2.0, w: colw - 0.2, h: 0.3, fontSize: 15, bold: true, color: COL.navy });
    tx(s, sub, { x, y: 2.3, w: colw - 0.2, h: 0.22, fontSize: 10, color: COL.muted, italic: true });
    s.addShape(pres.shapes.OVAL, { x, y: ly - 0.09, w: 0.18, h: 0.18, fill: { color: i === 3 ? COL.orange : COL.navy }, line: { color: COL.canvas, width: 1.5 }, objectName: 'milestone' });
    tx(s, paras(items, {}, false).map((r, j) => (j && r.text.startsWith('Komodo') || j && /^(Flatiron|Norstella)/.test(r.text) ? { text: r.text, options: Object.assign({}, r.options, { paraSpaceBefore: 8 }) } : r)),
      { x, y: 2.86, w: colw - 0.25, h: 1.5, fontSize: 10.5 });
  });

  const items = [
    ['THE BROADER COMPETITIVE QUESTION', [{ text: 'Where else are life sciences vendors marketing conversational AI, what tasks does it serve, and how does that positioning compare with Optum LS?', options: { color: COL.navy } }]],
    ['OUR RESEARCH RESPONSE', 'Expand beyond RWE to assess publicly-marketed conversational AI across five Optum LS-aligned domains, prioritizing the U.S. competitive landscape and using selected international offerings for context.'],
  ];
  const bottom = 6.6, top = Math.min(4.5, bottom - cardsHeight(items, 12) - 0.1);
  cardRow(s, top, bottom, items, 12);
}

// ---------- SLIDE 4 — U.S. BASELINE AND INTERNATIONAL SCREENS ----------
async function slide4() {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'International competitor screens' });
  s.addText('U.S. BASELINE AND INTERNATIONAL COMPETITOR SCREENS', { placeholder: 'kicker' });
  title(s, 'Selected international screens identify additional competitors marketing conversational AI across Optum LS-relevant domains');
  label(s, 'COMPARABLE COMPANIES MARKETING CONVERSATIONAL AI', L, 1.62, 8.4);

  const data = [['United States', 15, 43, '34.9%'], ['China', 5, 20, '25.0%'], ['Spain', 3, 17, '17.6%'], ['Israel', 2, 15, '13.3%'], ['France', 2, 17, '11.8%'], ['Nordics', 2, 17, '11.8%'],
    ['India', 3, 32, '9.4%'], ['Netherlands', 1, 11, '9.1%'], ['Australia', 1, 14, '7.1%'], ['South Korea', 1, 17, '5.9%'], ['Belgium', 1, 18, '5.6%'], ['United Kingdom', 1, 18, '5.6%'],
    ['Canada', 0, 11, '0.0%'], ['Germany', 0, 13, '0.0%'], ['Japan', 0, 15, '0.0%']];
  const cw = 8.35, hy = 1.96, hh = 0.56, ry = hy + hh + 0.14, rh = 0.25, maxV = 0.4;
  const lab = { x: L, w: 1.3 }, ch = { x: L + 1.42, w: 5.0 };
  const usEnd = ch.x + ch.w * (15 / 43) / maxV + 0.07 + 0.62; // right edge of the "15 / 43" label
  const pc = { x: usEnd + 0.12, w: L + cw - usEnd - 0.12 };
  s.addShape(pres.shapes.RECTANGLE, { x: L, y: hy, w: cw, h: hh, fill: { color: COL.navy }, line: { type: 'none' }, objectName: 'table header' });
  const hd = { y: hy, h: hh, fontSize: 9, bold: true, color: COL.white, valign: 'middle', align: 'center' };
  tx(s, [R('International', { breakLine: true }), R('Screen')], { ...hd, x: lab.x, w: lab.w });
  tx(s, [R('Companies with conversational AI /', { breakLine: true }), R('Total comparable companies')], { ...hd, x: ch.x, w: usEnd - ch.x });
  tx(s, [R('Screened-', { breakLine: true }), R('company share')], { ...hd, x: pc.x, w: pc.w });

  const plotH = rh * data.length, frameH = plotH + 0.28;
  const rev = [...data].reverse(); // bar charts plot the first category at the bottom
  s.addChart(pres.charts.BAR, [{ name: 'Screened-company share', labels: rev.map(d => d[0]), values: rev.map(d => d[1] / d[2]) }], {
    x: ch.x, y: ry, w: ch.w, h: frameH, barDir: 'bar', barGapWidthPct: 55,
    chartColors: rev.map(d => (d[0] === 'United States' ? HEX.navy : '8FA3C4')),
    valAxisMinVal: 0, valAxisMaxVal: maxV, valAxisMajorUnit: 0.1, valAxisLabelFormatCode: '0%', valAxisLabelFontSize: 8, valAxisLabelColor: HEX.muted, valAxisLabelFontFace: '+mn-lt', valAxisLineShow: false,
    catAxisHidden: true, valGridLine: { color: 'E3DED3', size: 0.5 }, catGridLine: { style: 'none' }, showLegend: false, showValue: false,
    layout: { x: 0, y: 0, w: 1, h: plotH / frameH }, objectName: 'screened-company share chart',
  });
  data.forEach(([name, n, d, pct], i) => {
    const y = ry + i * rh, us = name === 'United States';
    tx(s, name, { x: lab.x, y, w: lab.w - 0.05, h: rh, fontSize: 9, bold: us, color: us ? COL.navy : COL.body, align: 'right', valign: 'middle' });
    const end = ch.x + ch.w * (n / d) / maxV;
    tx(s, [B(`${n}`, { color: COL.navy }), R(` / ${d}`)], { x: end + 0.07, y, w: 0.7, h: rh, fontSize: 9, valign: 'middle', objectName: 'fraction ' + name });
    tx(s, pct, { x: pc.x, y, w: pc.w, h: rh, fontSize: 9, bold: us, color: us ? COL.navy : COL.muted, align: 'center', valign: 'middle' });
  });

  // Right panel: each card's label + text is centered as a group; the arrow sits beside the NEXT label
  const rx = L + cw + 0.25, rw = W - cw - 0.25, iw = rw - 0.44, lab13 = { fontSize: 13, bold: true, color: COL.orange, charSpacing: 1.5, objectName: 'label' };
  const h1 = 1.95, l1 = 0.52, t1 = 0.9, g1 = hy + (h1 - (l1 + 0.1 + t1)) / 2;
  card(s, rx, hy, rw, h1, 'how to read card');
  tx(s, 'HOW TO READ THE COMPARISON', { ...lab13, x: rx + 0.22, y: g1, w: iw, h: l1 });
  tx(s, 'Each bar shows the share of comparable companies we screened that market conversational AI. International counts exclude competitors already represented in the U.S. inventory.',
    { x: rx + 0.22, y: g1 + l1 + 0.1, w: iw, h: t1, fontSize: 10 });
  const ny = hy + h1 + 0.2, nh = 6.68 - ny, l2 = 0.8, t2 = 0.9, g2 = ny + (nh - (l2 + 0.16 + t2)) / 2;
  card(s, rx, ny, rw, nh, 'next card');
  await iconDot(s, 'MdArrowForward', rx + 0.22, g2 + l2 / 2 - 0.22, 0.44);
  tx(s, 'NEXT: FROM COMPANY SCREENS TO ACTUAL PRODUCTS', { ...lab13, x: rx + 0.8, y: g2, w: iw - 0.58, h: l2, valign: 'middle' });
  tx(s, 'The following slides examine the specific marketed conversational tools and corresponding Optum LS offerings – beginning with RWE.',
    { x: rx + 0.22, y: g2 + l2 + 0.16, w: iw, h: t2, fontSize: 11.5, bold: true, color: COL.navy });
}

// ---------- SLIDES 5–9 — workflow task tables ----------
// Bottom cards are sized to their content; spare height goes into table row padding.
const QL = 'BEYOND THE PUBLIC EVIDENCE';
async function workflowSlide(o) {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: o.section });
  s.addText(o.kicker, { placeholder: 'kicker' });
  tracker(s, o.active);
  title(s, o.title);
  const topt = Object.assign({ headAlign: 'center', taskAlign: 'center', plainCountry: true }, o.table || {});
  const tables = o.tables || [{ label: o.label, rows: o.rows }];
  const TGAP = 0.16, top = 1.62, cardBottom = 6.62;
  const lblH = t => (t.label ? 0.3 : 0);
  const cardTop = cardBottom - cardsHeight(o.cards);
  const natural = tables.reduce((n, t) => n + lblH(t) + TT_HEAD + rowHeights(t.rows, topt).reduce((a, b) => a + b, 0), 0) + TGAP * (tables.length - 1);
  const nRows = tables.reduce((n, t) => n + t.rows.length, 0);
  const pad = Math.max(0, Math.min(0.4, (cardTop - 0.2 - top - natural) / nRows));
  let y = top;
  tables.forEach(t => { if (t.label) sectionLabel(s, t.label, y, t.counts); y = taskTable(s, y + lblH(t), t.rows, Object.assign({}, topt, { rowPad: pad })) + TGAP; });
  const ct = Math.max(cardTop, y - TGAP + 0.18);
  cardRow(s, ct, cardBottom, o.cards);
  return s;
}
const slide5 = () => workflowSlide({
  section: 'Workflow deep dives', kicker: 'CONVERSATIONAL AI IN RWE', active: ['RWE'],
  title: 'Competitors are marketing conversational tools to assess healthcare|datasets, explore cohorts and answer RWE questions',
  rows: [
    ['Identify and assess|suitable data', ['Merative’s **MarketScan Snapshot**: natural-language assessment of dataset suitability;', 'HealthVerity’s **eXOs**: conversational exploration of connected healthcare data'], 'Market Clarity Data; linked claims/EHR data; Enriched Clinical Data'],
    ['Explore patient|populations and cohorts', ['Komodo Health’s **Marmot** and **MapAI**: conversational healthcare analytics and cohort exploration;', 'Truveta’s **Truveta Intelligence**: natural-language research questions over longitudinal clinical data;', 'Briya’s **AIRE Population Agent** (Israel): natural-language cohort eligibility and feasibility assessment'], 'Analytics on Demand; Clinical Notes Lab; Prospector'],
    ['Investigate broader|RWE questions', ['Norstella’s **Ella AI**: conversational RWE analytics;', 'BC Platforms’ **BC Catalyst** (Finland): prompt-based cohort building and clinical/genomic research exploration'], 'Analytics on Demand; epidemiology research services'],
  ],
  cards: [[QL, 'What information can researchers query directly? What outputs can they generate, inspect and reuse? Where do dataset permissions, cohort definitions or specialist involvement constrain the experience?']],
});
const slide6 = () => workflowSlide({
  section: 'Workflow deep dives', kicker: 'SPECIALIZED RWE AND EVIDENCE GENERATION', active: ['RWE'],
  title: 'Specialty-data platforms bring conversational AI into oncology|research and more narrowly defined evidence questions',
  rows: [
    ['Explore oncology and|specialty cohorts', ['Flatiron Health’s **Telescope**: conversational oncology cohort exploration;', 'Tempus’s **Lens**: clinical and molecular research exploration'], 'Enriched Clinical Data; clinicogenomics datasets'],
    ['Investigate specialty-|disease evidence', ['ConcertAI’s **Precision Explorer**: oncology research analytics;', 'Verana Health’s **Agent Claire**: conversational specialty-data research;', 'OrigiMed’s **OMSmartData** (China): AI-supported analysis of oncology real-world data'], 'Clinical Notes Lab; epidemiology and observational research'],
    ['Analyze curated|research datasets', ['MDClone’s **ADAMS Copilot** (Israel): conversational statistical analysis and visualization within controlled research datasets'], 'Analytics on Demand; research data and study environments'],
  ],
  cards: [
    ['WHERE DOES CONVERSATIONAL ANALYSIS END AND SCIENTIFIC EVIDENCE GENERATION BEGIN?', 'Optum also markets epidemiology studies and Evidence Engine services. The conversational products above are not automatically equivalent to those scientific services.'],
    [QL, 'Which tools support exploratory questions versus reproducible study analyses? What supporting methods, provenance and review can customers inspect? Which capabilities are publicly demonstrated, and which require further validation?'],
  ],
});
const slide7 = () => workflowSlide({
  section: 'Workflow deep dives', kicker: 'CONVERSATIONAL AI IN COMMERCIAL INTELLIGENCE', active: ['Commercial Intelligence'],
  title: 'Commercial Intelligence AI products extend conversational interaction into market intelligence, field decisions and simulated stakeholder research',
  rows: [
    ['Pharmaceutical market and portfolio intelligence', ['IQVIA’s **GMI Agent**: global market insights;', 'GBI Health’s **GBI BOT** (China): pharmaceutical database Q&A and competitive intelligence;', 'Prospection’s **ProGPT** (Australia): conversational patient and market insights for brand planning'], 'Market Clarity: market attractiveness and portfolio research'],
    ['Commercial field and prescribing decisions', ['IQVIA’s **Field Force Agent**: commercial field support;', 'ZS’s **ZAIDYN Conversational AI for Commercial**: commercial information and decision support;', 'Close-Up International’s **Xplora AI** (Spain): AI-enabled pharmaceutical CRM and commercial intelligence'], 'Market Clarity: patient and prescriber analytics'],
    ['Simulated stakeholder research', ['Trinity Life Sciences’ **InsightsEDGE Digital Twins**: conversational simulated personas;', 'Ipsos’s **Healthcare PersonaBot** (France): interaction with simulated healthcare stakeholders'], 'Market Clarity provides market insights; a directly comparable simulated-persona offering was not established in the Optum review'],
  ],
  cards: [
    [QL, 'Which decisions are these products designed to support? What underlying commercial information can users access? Where does an assistant move beyond analytics into field execution or simulated research?'],
    ['TRANSITION TO OTHER WORKFLOWS', 'How does the comparison change for medical affairs, clinical development and market access, where the intended users, evidence requirements and decisions differ?'],
  ],
});
const slide8 = () => workflowSlide({
  section: 'Workflow deep dives', kicker: 'CONVERSATIONAL AI IN MEDICAL / REGULATORY RESEARCH', active: ['Medical / Regulatory Research'],
  title: 'Competitors market conversational tools for scientific evidence,|medical insights and regulated content workflows',
  rows: [
    ['Search and synthesize|scientific evidence', ["IQVIA's **Med-R1 Deep Research Agent**: synthesizes biomedical, clinical and regulatory sources to answer complex research questions"], 'Epidemiology research; postapproval safety and comparative-effectiveness studies'],
    ['Explore medical insights and|support scientific engagement', ["ZS's **ZAIDYN Conversational AI for Medical**: natural-language analysis of care gaps, scientific insights and opinion-leader information;", "Evidenze's **AI medical-education chatbots** (Spain): conversational medical education and knowledge support"], 'Market Clarity medical affairs evidence and MSL insights'],
    ['Question and navigate|regulated documents', ["Veeva's **Vault AI Chat**: questions, summaries and follow-up interactions within authorized Vault documents and records"], "Optum's safety and evidence-research services are adjacent; a directly comparable document-chat offering was not established in the review"],
  ],
  cards: [[QL, 'Which offerings help users find evidence, and which help them interpret or act on it? What source traceability, permissions and scientific review are built into the marketed experience? Where are these fundamentally different services?']],
});
const slide9 = () => workflowSlide({
  section: 'Workflow deep dives', kicker: 'CONVERSATIONAL AI IN CLINICAL TRIALS AND MARKET ACCESS', active: ['Clinical Trials', 'Market Access'],
  title: 'Conversational offerings address specific trial operations and pricing questions, not the entire clinical or market-access workflow',
  tables: [
    { label: 'CLINICAL TRIALS — COMPETITIVE EXAMPLES', rows: [
      ['Trial feasibility, site and investigator identification', ["TriNetX's **Terrain AI Assistant**: natural-language questions about trial sites, investigator experience and patient populations"], 'Prospector Data Mining Studio; Optum Digital Research Network'],
      ['Clinical-trial document operations', ["Taimei's **Wiz-iCTA** (China): conversational assistant supporting trial project teams and document-management work"], 'Optum True Source supports clinical-trial source-data capture; a directly equivalent document assistant was not established'],
    ] },
    { label: 'MARKET ACCESS — COMPETITIVE EXAMPLES', rows: [
      ['Pharmaceutical pricing and market-access intelligence', ["EVERSANA's **Ask NAVLIN**: conversational access to pricing, reimbursement and market-access information"], 'Optum HEOR and value evidence; Evidence Engine; health economic and budget impact modeling'],
    ] },
  ],
  cards: [[QL, [
    R('For Clinical Trials, does the interface help identify viable sites, evaluate patients or manage trial documents—and how far does it go?', { breakLine: true }),
    R('For Market Access, is the customer asking for pricing intelligence, reimbursement information or new economic evidence? Where do the vendor offerings overlap with Optum, and where do they solve a different problem?', { paraSpaceBefore: 6 }),
  ]]],
});

// ---------- SLIDE 10 — AI SUPPLIERS ----------
async function slide10() {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'AI suppliers' });
  s.addText('ADDITIONAL COMPETITIVE DYNAMIC: AI SUPPLIERS EXPAND INTO LIFE SCIENCES', { placeholder: 'kicker' });
  s.addText('Conversational AI providers are expanding into pharmaceutical workflows through partnerships, specialized offerings and acquisitions', { placeholder: 'title' });
  label(s, 'THREE OBSERVED APPROACHES TO ENTERING LIFE SCIENCES', L, 1.62, 8);
  // Row 1 (partnering through vendors) tinted; rows 2–3 (more direct routes to pharma customers) navy
  const bottom = taskTable(s, 1.9, [
    ['Partner with established life-sciences vendors', ['**Anthropic + Komodo Health:** A documented partnership supporting AI-enabled healthcare analytics;', "**Anthropic + Veeva:** Claude incorporated into Veeva's industry-specific AI offerings;", '**AWS + Datavant:** Explore Assistant, publicly described in February 2026, uses Bedrock and AWS Clean Rooms to support natural-language assessment of healthcare datasets and research feasibility.'],
      'AI suppliers can contribute underlying technology to existing life-sciences platforms while those companies retain specialized data, applications and customer relationships. Komodo and Veeva already appear in our competitive analysis.'],
    ['Develop their own life-sciences-specific products and tools', ["**Anthropic's Claude for Life Sciences:** Specialized scientific tools and integrations supporting research, clinical trials and regulatory workflows;", "**OpenAI's GPT-Rosalind:** A purpose-built model for biology, drug discovery and translational research, with access expanded to eligible organizations in September 2026."],
      "Rather than acting only as technology suppliers, AI companies are building offerings aimed more directly at life-sciences users. These products are not necessarily direct substitutes for Optum's RWE and healthcare-data offerings, but they illustrate potential expansion into adjacent customer workflows."],
    ['Acquire specialized life-sciences capabilities', ['**Cohere + Reliant AI:** In May 2026, Cohere acquired a biopharma AI company offering literature review, competitive landscaping and scientific/regulatory information extraction. The acquisition supports development of North for Pharma, a planned specialized agentic offering.'],
      'Acquisitions provide another way for AI companies to obtain industry-specific technology, expertise and pharmaceutical customer relationships—potentially strengthening their ability to supply or compete with established life-sciences platforms.'],
  ], {
    head: ['AI SUPPLIER APPROACH', 'SELECTED PUBLIC DEVELOPMENTS', 'COMPETITIVE SIGNIFICANCE FOR OPTUM LS'], groupAfter: 0, colW: [2.15, 5.85, 4.13], fontSize: 9, headH: 0.44,
    taskFill: i => (i === 0 ? COL.tint : COL.navy), taskColor: i => (i === 0 ? COL.navy : COL.white),
  });
  const qy = bottom + 0.14, qh = 6.7 - qy;
  if (qh < 0.8) console.warn('slide 10 question card too short:', qh.toFixed(2));
  card(s, L, qy, W, qh, 'strategic question card');
  const gq = qy + Math.max(0.08, (qh - (0.44 + 0.08 + 0.85)) / 2);
  await iconDot(s, 'MdOutlineLightbulb', L + 0.22, gq, 0.44);
  tx(s, 'STRATEGIC QUESTIONS FOR OPTUM LS', { x: L + 0.8, y: gq, w: 8, h: 0.44, fontSize: 13, bold: true, color: COL.orange, charSpacing: 1.5, valign: 'middle', objectName: 'label' });
  tx(s, [
    B('As conversational AI providers expand beyond supplying general-purpose models, where will they remain technology partners to established life-sciences vendors—and where might they increasingly compete for the same pharmaceutical customers and workflows?', { color: COL.navy, breakLine: true }),
    B('What could these developments mean for the competitive positions of companies whose offerings combine differentiated healthcare data, research tools and scientific expertise?', { color: COL.navy, paraSpaceBefore: 6 }),
  ], { x: L + 0.8, y: gq + 0.52, w: W - 1.05, h: 0.85, fontSize: 11 });
}

// ---------- SLIDES 11–12 — METHODOLOGY NOTE (long notes and evidence boundaries moved off their slides) ----------
const NOTES = [
    ['SLIDE 1 — EVIDENCE BOUNDARY', 'Public product marketing and announcements reviewed as of October 13, 2026. The findings describe selected screened inventories—not a market census, adoption rates, verified product performance or undisclosed capabilities.'],
    ['SLIDE 2 — EVIDENCE BOUNDARY', 'Dates reflect selected public announcements or releases, not necessarily first feature availability. The timeline illustrates a research catalyst, not measured market-wide growth. AI-enabled or agentic launches are not automatically qualifying conversational products; the Norstella example is excluded from the final conversational count.'],
    ['SLIDE 3 — METHODOLOGY NOTE', [R('Comparable competitors include '), B('direct and specialized/partial offering-level competitors'), R(" serving similar customer tasks; adjacent participants are excluded. Counts represent distinct company groups per workflow, not product totals. A company may appear in multiple workflows. Optum's ○ indicates that a qualifying conversational interface was "), R('not publicly established in this review', { italic: true }), R(', not that the capability is absent. These are selected screened companies, not market-wide adoption rates.')]],
    ['SLIDE 4 — INDIA RECONCILIATION', 'Of 37 screened companies, 36 qualify as direct/specialized comparators in at least one of the five decision workflows; excluding four companies already represented in the qualifying U.S. chat cohort leaves 32. Three additional chat-marketing companies qualify: Axtria, Doceree and Genpact. Company presence in an India screen does not prove the specific product is available or deployed in India.'],
    ['SLIDE 4 — EVIDENCE BOUNDARY', 'Selected country/region screens differ in research depth and company composition and are not representative national market samples. The Nordics screen groups Denmark, Finland, Norway and Sweden. A multinational may appear in multiple country screens; the counts are not additive across countries. A globally marketed product does not establish locally verified use. Percentages show shares within screened comparable companies, not regional adoption, customer penetration or product deployment rates.'],
    ['SLIDE 5 — EVIDENCE BOUNDARY', 'Scope: Six qualifying U.S. conversational offerings plus two selected international examples. Optum offerings are task-level comparators, not established conversational equivalents. Descriptions reflect public marketing, not independently verified performance or data access.'],
    ['SLIDE 6 — EVIDENCE BOUNDARY', 'Scope: Four qualifying U.S. conversational offerings plus two selected international examples. Product descriptions are based on public marketing; study-grade reliability, data entitlements and actual deployment are not independently established.'],
    ['SLIDE 7 — EVIDENCE BOUNDARY', 'Scope: All four qualifying U.S. commercial conversational offerings plus four selected international examples. Customer tasks are not interchangeable; Optum comparisons are based on public offerings, not feature-for-feature validation. International examples are illustrative, not a geographic prevalence comparison.'],
    ['SLIDE 8 — EVIDENCE BOUNDARY', 'Three qualifying U.S. offerings plus one selected Spanish example. The examples span distinct tasks; scientific research, medical education, document assistance and pharmacovigilance are not interchangeable. Optum comparisons reflect publicly marketed capabilities, not verified conversational equivalence.'],
    ['SLIDE 9 — EVIDENCE BOUNDARY', 'Clinical trials — one qualifying U.S. offering and one selected Chinese example. Market access — one qualifying U.S. offering. No additional qualifying international market-access offering was identified in the selected inventory. The Optum comparisons are based on customer-task overlap, not identical product functionality.'],
    ['SLIDE 10 — EVIDENCE BOUNDARY', "These are selected publicly documented developments, not an exhaustive assessment of AI suppliers. The three approaches are neither mutually exclusive nor predictions of future competition. They do not establish that AI suppliers caused the recent RWE launch cluster, are displacing established vendors, or expose an Optum capability gap. Not all supplier offerings qualify as conversational products; they are not additions to the deck's product counts. North for Pharma was described as planned; Explore Assistant's general commercial availability was not established in our review."],
  ];
async function notesSlide(list) {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Methodology note' });
  s.addText('METHODOLOGY NOTE', { placeholder: 'kicker' });
  // 3-column grid; each row as tall as its longest note
  const gap = 0.18, w = (W - 2 * gap) / 3, fs = 10, cpl = Math.floor((w - 0.4) / (fs * 0.0057));
  const hOf = n => 0.85 + runText(n[1]).reduce((k, p) => k + Math.ceil(p.length / cpl), 0) * fs * 1.22 / 72;
  let y = 0.9;
  for (let r = 0; r * 3 < list.length; r++) {
    const row = list.slice(r * 3, r * 3 + 3), h = Math.max(...row.map(hOf));
    row.forEach(([lab, body], c) => {
      const x = L + c * (w + gap);
      card(s, x, y, w, h, lab.toLowerCase());
      tx(s, lab, { x: x + 0.2, y: y + 0.16, w: w - 0.4, h: 0.24, fontSize: 10, bold: true, color: COL.orange, charSpacing: 1, objectName: 'label' });
      tx(s, body, { x: x + 0.2, y: y + 0.46, w: w - 0.4, h: h - 0.54, fontSize: fs });
    });
    y += h + gap;
  }
  if (y > 6.95) console.warn('notes overflow:', y.toFixed(2));
}
const slide11 = () => notesSlide(NOTES.slice(0, 6));
const slide12 = () => notesSlide(NOTES.slice(6));

(async () => {

  const all = [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10, slide11, slide12];
  const only = (process.env.SLIDES || '1,2,3,4,5,6,7,8,9,10,11,12').split(',').map(Number);
  const sections = { 1: 'Executive summary', 2: 'Research catalyst', 3: 'U.S. competitive landscape', 4: 'International competitor screens', 5: 'Workflow deep dives', 10: 'AI suppliers', 11: 'Methodology note' };
  for (const [i, fn] of all.entries()) {
    if (sections[i + 1]) pres.addSection({ title: sections[i + 1] });
    if (only.includes(i + 1)) await fn();
  }
  const out = path.resolve(process.argv[2] || 'deck.pptx');
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log('wrote', out);
})();
