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
const B = (text, extra) => ({ text, options: Object.assign({ bold: true }, extra) });
const R = (text, extra) => ({ text, options: Object.assign({}, extra) });

// Workflow chip — the five decision workflows are the deck's core taxonomy; same chip everywhere.
const chipW = (t, fs) => 0.32 + t.length * fs * 0.0078;
function chip(s, t, x, y, o = {}) {
  const fs = o.fontSize || 10, h = o.h || 0.32, w = o.w || chipW(t, fs);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: h / 2, fill: { color: COL.tint }, line: { color: COL.navy, width: 0.75 }, objectName: 'workflow chip ' + t });
  tx(s, t, { x, y, w, h, fontSize: fs, bold: true, color: COL.navy, align: 'center', valign: 'middle' });
  return w;
}
const cardLabel = (s, t, x, y, w) => tx(s, t, { x, y, w, h: 0.3, fontSize: 13, bold: true, color: COL.orange, charSpacing: 1.5, valign: 'middle', objectName: 'card label' });

// ---------- SLIDE 1 — EXECUTIVE SUMMARY ----------
async function slide1() {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Executive summary' });
  s.addText('EXECUTIVE SUMMARY', { placeholder: 'kicker' });
  s.addText('Conversational AI is being marketed across life sciences workflows, creating new ways to access information, analytics and decision support', { placeholder: 'title' });

  // Outside-in research lens (full width) with the five workflows as chips
  const ly = 1.68, lh = 1.5;
  card(s, L, ly, W, lh, 'lens card');
  await iconDot(s, 'MdTravelExplore', L + 0.25, ly + 0.22, 0.52);
  cardLabel(s, 'OUTSIDE-IN RESEARCH LENS', L + 0.95, ly + 0.2, 6);
  tx(s, 'An assessment of publicly marketed conversational AI across five Optum Life Sciences decision workflows:', { x: L + 0.95, y: ly + 0.56, w: 10.9, h: 0.22, fontSize: 11 });
  let x = L + 0.95; const chy = ly + 0.84;
  for (const [i, t] of ['RWE', 'Commercial intelligence', 'Medical / regulatory research', 'Clinical trials', 'Market access'].entries()) {
    if (i === 4) { tx(s, 'and', { x, y: chy, w: 0.36, h: 0.32, fontSize: 10, color: COL.muted, align: 'center', valign: 'middle' }); x += 0.42; }
    x += chip(s, t, x, chy) + 0.12;
  }
  tx(s, 'The U.S. competitive landscape is the primary focus; selected international offerings provide context; and Optum LS is the strategic reference point.', { x: L + 0.95, y: ly + 1.22, w: 10.9, h: 0.22, fontSize: 11 });

  const cy = 3.32, ch = 2.85, cw = 3.91, gap = 0.2;
  const cx = [L, L + cw + gap, L + 2 * (cw + gap)];
  const head = async (i, ic, t) => { card(s, cx[i], cy, cw, ch, t.toLowerCase() + ' card'); await iconDot(s, ic, cx[i] + 0.25, cy + 0.22, 0.5); cardLabel(s, t, cx[i] + 0.88, cy + 0.32, cw - 1.0); };

  // Key findings — number and sentence share a top line and a fixed gutter
  await head(0, 'MdInsights', 'KEY FINDINGS');
  const kf = [['19', 'The U.S. screen identified 19 qualifying conversational AI offerings across five decision workflows, concentrated in RWE and commercial intelligence.'],
              ['10', 'The selected international comparison identified 10 qualifying offerings, providing additional examples of how vendors position conversational interaction.']];
  kf.forEach(([n, t], i) => {
    const y = cy + 0.95 + i * 0.98;
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
  tx(s, "Where could conversational interaction add value to Optum's existing offerings?",
    { x: cx[2] + 0.25, y: cy + 1.75, w: cw - 0.5, h: 0.95, fontSize: 14, bold: true, color: COL.navy, objectName: 'strategic question' });

  tx(s, 'Evidence boundary: Public product marketing and announcements reviewed as of October 13, 2026. The findings describe selected screened inventories—not a market census, adoption rates, verified product performance or undisclosed capabilities.',
    { x: L, y: 6.3, w: W, h: 0.4, fontSize: 9, italic: true, color: COL.muted, objectName: 'evidence boundary' });
}

// ---------- SLIDE 3 — U.S. COMPETITIVE LANDSCAPE ----------
async function slide3() {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'U.S. competitive landscape' });
  s.addText('U.S. COMPETITIVE LANDSCAPE', { placeholder: 'kicker' });
  s.addText('Conversational AI is most prominent in RWE among screened U.S. competitors; Optum has relevant offerings across five decision workflows', { placeholder: 'title' });

  // U.S. competitive screen — KPI tiles linked by arrows
  label(s, 'U.S. COMPETITIVE SCREEN', L, 1.6, 6);
  const tiles = [['45', 'competitor groups initially screened'], ['15', 'distinct groups marketing qualifying conversational AI'], ['19', 'qualifying conversational AI offerings across five decision workflows']];
  const aw = 0.62, tw = (W - 2 * aw) / 3, ty = 1.86, th = 0.58;
  tiles.forEach(([n, t], i) => {
    const x = L + i * (tw + aw);
    s.addShape(pres.shapes.RECTANGLE, { x, y: ty, w: tw, h: th, fill: { color: COL.tint }, line: { type: 'none' }, objectName: 'kpi tile ' + n });
    tx(s, n, { x: x + 0.18, y: ty, w: 0.75, h: th, fontSize: 26, bold: true, color: i === 2 ? COL.orange : COL.navy, valign: 'middle' });
    tx(s, t, { x: x + 0.95, y: ty, w: tw - 1.08, h: th, fontSize: 10, color: COL.navy, valign: 'middle' });
    if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + tw + 0.14, y: ty + th / 2 - 0.13, w: aw - 0.28, h: 0.26, fill: { color: COL.muted }, line: { type: 'none' }, objectName: 'flow arrow' });
  });

  // Competitor coverage and Optum LS positioning — mirrored chart + table
  label(s, 'COMPETITOR COVERAGE AND OPTUM LS POSITIONING', L, 2.6, 7);
  const rows = [
    ['RWE', 9, 34, 'Komodo Marmot and MapAI; Truveta Intelligence; Flatiron Telescope; Tempus Lens', 'Analytics on Demand; epidemiology research'],
    ['Commercial intelligence', 3, 23, 'IQVIA Field Force Agent and GMI Agent; ZS Conversational AI for Commercial; Trinity InsightsEDGE Digital Twins', 'Market Clarity: portfolio and patient/prescriber analytics'],
    ['Medical / regulatory research', 3, 18, 'IQVIA Med-R1; ZS Conversational AI for Medical; Veeva Vault AI Chat', 'Postapproval safety studies; Market Clarity medical affairs insights'],
    ['Clinical trials', 1, 25, 'TriNetX Terrain AI Assistant', 'Prospector; Digital Research Network; True Source'],
    ['Market access', 1, 24, 'EVERSANA Ask NAVLIN', 'HEOR/value evidence; Evidence Engine'],
  ];
  const hy = 2.86, hh = 0.48, ry = hy + hh, rh = 0.49;
  const cX = { bars: L, wf: 3.8, rep: 6.0, opt: 9.3, chat: 11.75 }, cW = { bars: 3.15, wf: 2.15, rep: 3.15, opt: 2.3, chat: 0.98 };
  s.addShape(pres.shapes.RECTANGLE, { x: L, y: hy, w: W, h: hh, fill: { color: COL.navy }, line: { type: 'none' }, objectName: 'table header' });
  const hdr = { y: hy, h: hh, fontSize: 9, bold: true, color: COL.white, valign: 'middle', align: 'center' };
  tx(s, [R('U.S. competitors marketing chat /', { breakLine: true }), R('all comparable competitors')], { ...hdr, x: cX.bars + 0.6, w: 2.45 });
  tx(s, 'Workflow', { ...hdr, x: cX.wf, w: cW.wf });
  tx(s, 'Representative conversational AI offerings', { ...hdr, x: cX.rep, w: cW.rep });
  tx(s, [R("Optum's publicly", { breakLine: true }), R('marketed offerings')], { ...hdr, x: cX.opt, w: cW.opt });
  tx(s, 'Optum public chat claim', { ...hdr, x: cX.chat - 0.1, w: cW.chat + 0.1 });

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
    tx(s, [B(`${n}`, { color: COL.navy }), R(` / ${d}`)], { x: barLeft - 0.66, y, w: 0.6, h: rh, fontSize: 10, align: 'right', valign: 'middle', objectName: 'fraction ' + wf });
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
  tx(s, 'KEY TAKEAWAY', { x: L + 0.2, y: ky, w: 1.7, h: kh, fontSize: 10, bold: true, color: COL.orange, charSpacing: 1, valign: 'middle', objectName: 'label' });
  tx(s, 'RWE has the largest number of screened U.S. companies marketing qualifying conversational AI. Optum offers products and services across the five decision workflows shown, but a qualifying customer-facing conversational interface was not established in the reviewed public materials.',
    { x: L + 1.95, y: ky, w: W - 2.15, h: kh, fontSize: 10.5, valign: 'middle' });
}

// ---------- shared helpers for slides 2, 4–11 ----------
const WF = ['RWE', 'Commercial intelligence', 'Medical / regulatory research', 'Clinical trials', 'Market access'];
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
  const fs = 7.5, h = 0.26, g = 0.07, ws = WF.map(t => 0.24 + t.length * 0.058);
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

// Three-column task table (Slides 5–9): returns bottom y
function taskTable(s, y, rows, opt = {}) {
  const colW = opt.colW || [2.5, 6.3, 3.33], fs = opt.fontSize || 9.5;
  const est = (str, w, f) => lines(str, Math.floor((w - 0.24) / (f * 0.0062))) * f * 1.22 / 72;
  const none = { type: 'none' }, hair = { type: 'solid', pt: 0.75, color: HEX.neutral }, white = { type: 'solid', pt: 2, color: HEX.white };
  const hd = { fill: { color: COL.navy }, color: COL.white, bold: true, fontSize: 9, valign: 'middle', charSpacing: 1, border: [none, none, none, none] };
  const head = (opt.head || ['CUSTOMER TASK', 'COMPETITOR CONVERSATIONAL AI OFFERINGS', 'RELEVANT OPTUM OFFERINGS']).map(t => ({ text: t, options: Object.assign({}, hd, opt.headAlign ? { align: opt.headAlign } : {}) }));
  const rowH = [0.36];
  const body = rows.map(([task, comp, optum], i) => {
    const h = Math.max(est(task, colW[0], 11.5), comp.reduce((n, c) => n + est(c, colW[1], fs), 0) + (comp.length - 1) * 0.06, est(optum, colW[2], fs)) + 0.2;
    rowH.push(h);
    const bot = opt.groupAfter === i ? { type: 'solid', pt: 1.5, color: HEX.navy } : hair;
    return [
      { text: task, options: { fill: { color: opt.taskFill ? opt.taskFill(i) : COL.tint }, color: opt.taskColor ? opt.taskColor(i) : COL.navy, bold: true, fontSize: 10.5, valign: 'middle', align: opt.taskAlign || 'left', border: [white, white, white, none] } },
      { text: paras(comp, {}, true, !!opt.plainCountry), options: { fontSize: fs, color: COL.body, valign: 'middle', border: [none, none, bot, none] } },
      { text: rich(optum), options: { fontSize: fs, color: COL.body, valign: 'middle', border: [none, none, bot, none] } },
    ];
  });
  s.addTable([head, ...body], { x: L, y, w: W, colW, rowH, margin: [0.07, 0.12, 0.07, 0.12], objectName: 'task table' });
  return y + rowH.reduce((a, b) => a + b, 0);
}
// Row of white cards with orange labels under a table
function cardRow(s, y, items, bottom = 6.18) {
  const gap = 0.2, w = (W - gap * (items.length - 1)) / items.length, h = bottom - y;
  items.forEach(([lab, runs, o], i) => {
    const x = L + i * (w + gap);
    card(s, x, y, w, h, lab.toLowerCase() + ' card');
    const lh = Math.ceil(lab.length * 0.092 / (w - 0.44)) * 0.2 + 0.04;
    tx(s, lab, { x: x + 0.22, y: y + 0.15, w: w - 0.44, h: lh, fontSize: 10, bold: true, color: COL.orange, charSpacing: 1, objectName: 'label' });
    tx(s, runs, Object.assign({ x: x + 0.22, y: y + 0.24 + lh, w: w - 0.44, h: h - 0.35 - lh, fontSize: 11 }, o || {}));
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
  s.addText('Recent RWE product launches prompted a broader examination of conversational AI across life sciences', { placeholder: 'title' });
  tx(s, 'A series of public announcements brought greater attention to natural-language access to healthcare data, cohort exploration and research workflows.', { x: L, y: 1.52, w: W, h: 0.25, fontSize: 12 });

  label(s, 'SELECTED PUBLIC PRODUCT ANNOUNCEMENTS', L, 1.92, 8);
  const ms = [
    ['2024', 'Earlier precedents', ['**February:** Tempus One / Lens — AI-assisted research and cohort exploration', '**September:** Komodo MapAI — conversational healthcare analytics']],
    ['August 2025', 'Platform expansion', ['**Komodo Marmot** — AI-powered healthcare analytics and research workflows']],
    ['April–May 2026', 'Research-interface launches', ['**Truveta Intelligence** — natural-language RWE research questions', '**Flatiron Telescope** — conversational oncology cohort exploration']],
    ['September 2026', 'New releases', ['**Verana Agent Claire** — conversational specialty-data research', '**Norstella Atlas Patient Analytics** — agentic, managed cohort analytics']],
  ];
  const colw = W / 4, ly = 2.86;
  s.addShape(pres.shapes.LINE, { x: L, y: ly, w: W, h: 0, line: { color: COL.navy, width: 1.5 }, objectName: 'timeline' });
  ms.forEach(([date, sub, items], i) => {
    const x = L + i * colw;
    tx(s, date, { x, y: 2.2, w: colw - 0.2, h: 0.3, fontSize: 15, bold: true, color: COL.navy });
    tx(s, sub, { x, y: 2.5, w: colw - 0.2, h: 0.22, fontSize: 10, color: COL.muted, italic: true });
    s.addShape(pres.shapes.OVAL, { x, y: ly - 0.09, w: 0.18, h: 0.18, fill: { color: i === 3 ? COL.orange : COL.navy }, line: { color: COL.canvas, width: 1.5 }, objectName: 'milestone' });
    // Norstella is excluded from the conversational count (see evidence boundary) — shown muted
    const runs = [];
    items.forEach((it, j) => {
      const r = rich(it, /Norstella/.test(it) ? { color: COL.muted } : {});
      if (j < items.length - 1) r[r.length - 1].options.breakLine = true;
      if (j > 0) r[0].options.paraSpaceBefore = 6;
      runs.push(...r);
    });
    tx(s, runs, { x, y: 3.05, w: colw - 0.25, h: 1.15, fontSize: 10 });
  });

  const cy = 4.38;
  const items = [
    ['BROADER COMPETITIVE SIGNAL', 'AI model providers are actively expanding enterprise sales into life sciences, offering models, integrations and implementation partnerships that may accelerate AI adoption among established healthcare-data vendors.'],
    ['THE BROADER COMPETITIVE QUESTION', [B('Where else are life sciences vendors marketing conversational AI, what tasks does it serve, and how does that positioning compare with Optum Life Sciences?', { color: COL.navy })]],
    ['OUR RESEARCH RESPONSE', 'Expand beyond RWE to assess publicly marketed conversational AI across five Optum-aligned decision workflows, prioritizing the U.S. competitive landscape and using selected international offerings for context.'],
  ];
  cardRow(s, cy, items.map(([a, b]) => [a, b]), 6.12);
  footnote(s, 'Evidence boundary: Dates reflect selected public announcements or releases, not necessarily first feature availability. The timeline illustrates a research catalyst, not measured market-wide growth. AI-enabled or agentic launches are not automatically qualifying conversational products; the Norstella example is excluded from the final conversational count.', 6.24);
}

// ---------- SLIDE 4 — U.S. BASELINE AND INTERNATIONAL SCREENS ----------
async function slide4() {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'International competitor screens' });
  s.addText('U.S. BASELINE AND INTERNATIONAL COMPETITOR SCREENS', { placeholder: 'kicker' });
  s.addText('Selected international screens identify additional competitors marketing conversational AI across Optum-relevant workflows', { placeholder: 'title' });
  label(s, 'COMPARABLE COMPETITOR GROUPS MARKETING QUALIFYING CHAT — FIVE DECISION WORKFLOWS', L, 1.62, 8.4);

  const data = [['United States (reference)', 15, 43, '34.9%'], ['China', 5, 20, '25.0%'], ['Spain', 3, 17, '17.6%'], ['Israel', 2, 15, '13.3%'], ['France', 2, 17, '11.8%'], ['Nordics', 2, 17, '11.8%'],
    ['India', 3, 32, '9.4%'], ['Netherlands', 1, 11, '9.1%'], ['Australia', 1, 14, '7.1%'], ['South Korea', 1, 17, '5.9%'], ['Belgium', 1, 18, '5.6%'], ['United Kingdom', 1, 18, '5.6%'],
    ['Canada', 0, 11, '0.0%'], ['Germany', 0, 13, '0.0%'], ['Japan', 0, 15, '0.0%']];
  const bold = ['United States (reference)', 'India'];
  const cw = 8.35, hy = 1.9, hh = 0.42, ry = hy + hh, rh = 0.245;
  const lab = { x: L, w: 1.85 }, ch = { x: L + 1.95, w: 4.95 }, pc = { x: L + 7.45, w: 0.9 };
  s.addShape(pres.shapes.RECTANGLE, { x: L, y: hy, w: cw, h: hh, fill: { color: COL.navy }, line: { type: 'none' }, objectName: 'table header' });
  const hd = { y: hy, h: hh, fontSize: 8.5, bold: true, color: COL.white, valign: 'middle' };
  tx(s, 'Screen', { ...hd, x: lab.x + 0.1, w: lab.w });
  tx(s, 'Distinct chat-marketing groups / distinct comparable competitor groups', { ...hd, x: ch.x, w: 5.3, align: 'center' });
  tx(s, 'Screened-company share', { ...hd, x: pc.x - 0.15, w: pc.w + 0.1, align: 'right' });

  const plotH = rh * data.length, frameH = plotH + 0.28, maxV = 0.4;
  const rev = [...data].reverse(); // bar charts plot the first category at the bottom
  s.addChart(pres.charts.BAR, [{ name: 'Screened-company share', labels: rev.map(d => d[0]), values: rev.map(d => d[1] / d[2]) }], {
    x: ch.x, y: ry, w: ch.w, h: frameH, barDir: 'bar', barGapWidthPct: 55,
    chartColors: rev.map(d => d[0].startsWith('United States') ? HEX.navy : '8FA3C4'),
    valAxisMinVal: 0, valAxisMaxVal: maxV, valAxisMajorUnit: 0.1, valAxisLabelFormatCode: '0%', valAxisLabelFontSize: 8, valAxisLabelColor: HEX.muted, valAxisLabelFontFace: '+mn-lt', valAxisLineShow: false,
    catAxisHidden: true, valGridLine: { color: 'E3DED3', size: 0.5 }, catGridLine: { style: 'none' }, showLegend: false, showValue: false,
    layout: { x: 0, y: 0, w: 1, h: plotH / frameH }, objectName: 'screened-company share chart',
  });
  data.forEach(([name, n, d, pct], i) => {
    const y = ry + i * rh, b = bold.includes(name);
    tx(s, name, { x: lab.x + 0.1, y, w: lab.w, h: rh, fontSize: 9, bold: b, color: b ? COL.navy : COL.body, valign: 'middle' });
    const end = ch.x + ch.w * (n / d) / maxV;
    tx(s, [B(`${n}`, { color: COL.navy }), R(` / ${d}`)], { x: end + 0.07, y, w: 0.7, h: rh, fontSize: 9, valign: 'middle', objectName: 'fraction ' + name });
    tx(s, pct, { x: pc.x, y, w: pc.w - 0.05, h: rh, fontSize: 8.5, bold: b, color: b ? COL.navy : COL.muted, align: 'right', valign: 'middle' });
  });

  // Right panel
  const rx = L + cw + 0.25, rw = W - cw - 0.25;
  card(s, rx, 1.62, rw, 2.75, 'how to read card');
  tx(s, 'HOW TO READ THE COMPARISON', { x: rx + 0.22, y: 1.77, w: rw - 0.44, h: 0.24, fontSize: 10, bold: true, color: COL.orange, charSpacing: 1, objectName: 'label' });
  tx(s, 'Each bar represents distinct company groups publicly marketing at least one qualifying conversational offering divided by distinct direct or specialized/partial competitor groups comparable in at least one of the five decision workflows. Count each group once per screen. International numerators and denominators exclude companies already represented in the qualifying U.S. chat inventory.',
    { x: rx + 0.22, y: 2.06, w: rw - 0.44, h: 2.2, fontSize: 9.5 });
  card(s, rx, 4.52, rw, 1.86, 'next card');
  await iconDot(s, 'MdArrowForward', rx + 0.22, 4.68, 0.4);
  tx(s, 'NEXT: FROM COMPANY SCREENS TO THE ACTUAL PRODUCTS', { x: rx + 0.75, y: 4.66, w: rw - 0.95, h: 0.44, fontSize: 10, bold: true, color: COL.orange, charSpacing: 1, valign: 'middle', objectName: 'label' });
  tx(s, 'The following slides examine the specific marketed conversational tools and corresponding Optum Life Sciences offerings, beginning with RWE dataset assessment, cohort exploration and research analysis.',
    { x: rx + 0.22, y: 5.2, w: rw - 0.44, h: 1.1, fontSize: 9.5 });
}

// ---------- SLIDES 5–9 — workflow task tables ----------
async function workflowSlide(o) {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: o.section });
  s.addText(o.kicker, { placeholder: 'kicker' });
  tracker(s, o.active);
  s.addText(o.title, { placeholder: 'title' });
  sectionLabel(s, o.label, 1.62, o.counts);
  const bottom = taskTable(s, 1.92, o.rows, o.table || {});
  cardRow(s, bottom + 0.2, o.cards, 6.18);
  footnote(s, o.eb, 6.3);
  return s;
}
const slide5 = () => workflowSlide({
  section: 'Workflow deep dives', kicker: 'CONVERSATIONAL AI IN RWE', active: ['RWE'],
  title: 'Competitors are marketing conversational tools to assess healthcare datasets, explore cohorts and answer RWE questions',
  label: 'RWE DATA ACCESS AND ANALYSIS — COMPETITIVE EXAMPLES',
  rows: [
    ['Identify and assess suitable data', ['Merative’s **MarketScan Snapshot**: natural-language assessment of dataset suitability;', 'HealthVerity’s **eXOs**: conversational exploration of connected healthcare data'], 'Market Clarity Data; linked claims/EHR data; Enriched Clinical Data'],
    ['Explore patient populations and cohorts', ['Komodo Health’s **Marmot** and **MapAI**: conversational healthcare analytics and cohort exploration;', 'Truveta’s **Truveta Intelligence**: natural-language research questions over longitudinal clinical data;', 'Briya’s **AIRE Population Agent** (Israel): natural-language cohort eligibility and feasibility assessment'], 'Analytics on Demand; Clinical Notes Lab; Prospector'],
    ['Investigate broader RWE questions', ['Norstella’s **Ella AI**: conversational RWE analytics;', 'BC Platforms’ **BC Catalyst** (Finland): prompt-based cohort building and clinical/genomic research exploration'], 'Analytics on Demand; epidemiology research services'],
  ],
  cards: [['QUESTIONS FOR COMPARISON', 'What information can researchers query directly? What outputs can they generate, inspect and reuse? Where do dataset permissions, cohort definitions or specialist involvement constrain the experience?']],
  eb: ebLabel('Scope: Six qualifying U.S. conversational offerings plus two selected international examples. Optum offerings are task-level comparators, not established conversational equivalents. Descriptions reflect public marketing, not independently verified performance or data access.'),
});
const slide6 = () => workflowSlide({
  section: 'Workflow deep dives', kicker: 'SPECIALIZED RWE AND EVIDENCE GENERATION', active: ['RWE'],
  title: 'Specialty-data platforms bring conversational AI into oncology research and more narrowly defined evidence questions',
  label: 'SPECIALIZED RWE — COMPETITIVE EXAMPLES',
  rows: [
    ['Explore oncology and specialty cohorts', ['Flatiron Health’s **Telescope**: conversational oncology cohort exploration;', 'Tempus’s **Lens**: clinical and molecular research exploration'], 'Enriched Clinical Data; clinicogenomics datasets'],
    ['Investigate specialty-disease evidence', ['ConcertAI’s **Precision Explorer**: oncology research analytics;', 'Verana Health’s **Agent Claire**: conversational specialty-data research;', 'OrigiMed’s **OMSmartData** (China): AI-supported analysis of oncology real-world data'], 'Clinical Notes Lab; epidemiology and observational research'],
    ['Analyze curated research datasets', ['MDClone’s **ADAMS Copilot** (Israel): conversational statistical analysis and visualization within controlled research datasets'], 'Analytics on Demand; research data and study environments'],
  ],
  cards: [
    ['WHERE DOES CONVERSATIONAL ANALYSIS END AND SCIENTIFIC EVIDENCE GENERATION BEGIN?', 'Optum also markets epidemiology studies and Evidence Engine services. The conversational products above are not automatically equivalent to those scientific services.'],
    ['QUESTIONS FOR COMPARISON', 'Which tools support exploratory questions versus reproducible study analyses? What supporting methods, provenance and review can customers inspect? Which capabilities are publicly demonstrated, and which require further validation?'],
  ],
  eb: ebLabel('Scope: Four qualifying U.S. conversational offerings plus two selected international examples. Product descriptions are based on public marketing; study-grade reliability, data entitlements and actual deployment are not independently established.'),
});
const slide7 = () => workflowSlide({
  section: 'Workflow deep dives', kicker: 'CONVERSATIONAL AI IN COMMERCIAL INTELLIGENCE', active: ['Commercial intelligence'],
  title: 'Commercial AI products extend conversational interaction into market intelligence, field decisions and simulated stakeholder research',
  label: 'COMMERCIAL INTELLIGENCE — COMPETITIVE EXAMPLES',
  table: { head: ['CUSTOMER TASK', 'CONVERSATIONAL AI OFFERINGS', 'RELEVANT OPTUM OFFERINGS'], headAlign: 'center', taskAlign: 'center', plainCountry: true },
  rows: [
    ['Pharmaceutical market and portfolio intelligence', ['IQVIA’s **GMI Agent**: global market insights;', 'GBI Health’s **GBI BOT** (China): pharmaceutical database Q&A and competitive intelligence;', 'Prospection’s **ProGPT** (Australia): conversational patient and market insights for brand planning'], 'Market Clarity: market attractiveness and portfolio research'],
    ['Commercial field and prescribing decisions', ['IQVIA’s **Field Force Agent**: commercial field support;', 'ZS’s **ZAIDYN Conversational AI for Commercial**: commercial information and decision support;', 'Close-Up International’s **Xplora AI** (Spain): AI-enabled pharmaceutical CRM and commercial intelligence'], 'Market Clarity: patient and prescriber analytics'],
    ['Simulated stakeholder research', ['Trinity Life Sciences’ **InsightsEDGE Digital Twins**: conversational simulated personas;', 'Ipsos’s **Healthcare PersonaBot** (France): interaction with simulated healthcare stakeholders'], 'Market Clarity provides market insights; a directly comparable simulated-persona offering was not established in the Optum review'],
  ],
  cards: [
    ['QUESTIONS FOR COMPARISON', 'Which decisions are these products designed to support? What underlying commercial information can users access? Where does an assistant move beyond analytics into field execution or simulated research?', { bold: true }],
    ['TRANSITION TO OTHER WORKFLOWS', 'How does the comparison change for medical affairs, clinical development and market access, where the intended users, evidence requirements and decisions differ?', { bold: true }],
  ],
  eb: ebLabel('Scope: All four qualifying U.S. commercial conversational offerings plus four selected international examples. Customer tasks are not interchangeable; Optum comparisons are based on public offerings, not feature-for-feature validation. International examples are illustrative, not a geographic prevalence comparison.'),
});
const slide8 = () => workflowSlide({
  section: 'Workflow deep dives', kicker: 'CONVERSATIONAL AI IN MEDICAL / REGULATORY RESEARCH', active: ['Medical / regulatory research'],
  title: 'Competitors market conversational tools for scientific evidence, medical insights and regulated content workflows',
  label: 'U.S. SCREEN: 3 COMPANIES | 3 QUALIFYING OFFERINGS',
  rows: [
    ['Search and synthesize scientific evidence', ["IQVIA's **Med-R1 Deep Research Agent**: synthesizes biomedical, clinical and regulatory sources to answer complex research questions"], 'Epidemiology research; postapproval safety and comparative-effectiveness studies'],
    ['Explore medical insights and support scientific engagement', ["ZS's **ZAIDYN Conversational AI for Medical**: natural-language analysis of care gaps, scientific insights and opinion-leader information;", "Evidenze's **AI medical-education chatbots** (Spain): conversational medical education and knowledge support"], 'Market Clarity medical affairs evidence and MSL insights'],
    ['Question and navigate regulated documents', ["Veeva's **Vault AI Chat**: questions, summaries and follow-up interactions within authorized Vault documents and records"], "Optum's safety and evidence-research services are adjacent; a directly comparable document-chat offering was not established in the review"],
  ],
  cards: [['QUESTIONS FOR COMPARISON', 'Which offerings help users find evidence, and which help them interpret or act on it? What source traceability, permissions and scientific review are built into the marketed experience? Where are these fundamentally different services?']],
  eb: 'Evidence boundary: Three qualifying U.S. offerings plus one selected Spanish example. The examples span distinct tasks; scientific research, medical education, document assistance and pharmacovigilance are not interchangeable. Optum comparisons reflect publicly marketed capabilities, not verified conversational equivalence.',
});
const slide9 = () => workflowSlide({
  section: 'Workflow deep dives', kicker: 'CONVERSATIONAL AI IN CLINICAL TRIALS AND MARKET ACCESS', active: ['Clinical trials', 'Market access'],
  title: 'Conversational offerings address specific trial operations and pricing questions—not the entire clinical or market-access workflow',
  label: 'TWO DISTINCT WORKFLOWS — COMPETITIVE EXAMPLES',
  table: { groupAfter: 1 },
  rows: [
    ['Trial feasibility, site and investigator identification', ["TriNetX's **Terrain AI Assistant**: natural-language questions about trial sites, investigator experience and patient populations"], 'Prospector Data Mining Studio; Optum Digital Research Network'],
    ['Clinical-trial document operations', ["Taimei's **Wiz-iCTA** (China): conversational assistant supporting trial project teams and document-management work"], 'Optum True Source supports clinical-trial source-data capture; a directly equivalent document assistant was not established'],
    ['Pharmaceutical pricing and market-access intelligence', ["EVERSANA's **Ask NAVLIN**: conversational access to pricing, reimbursement and market-access information"], 'Optum HEOR and value evidence; Evidence Engine; health economic and budget impact modeling'],
  ],
  cards: [['QUESTIONS FOR COMPARISON', [
    R('For trials, does the interface help identify viable sites, evaluate patients or manage trial documents—and how far does it go?', { breakLine: true }),
    R('For market access, is the customer asking for pricing intelligence, reimbursement information or new economic evidence? Where do the vendor offerings overlap with Optum, and where do they solve a different problem?', { paraSpaceBefore: 6 }),
  ]]],
  eb: 'Evidence boundary: Clinical trials — one qualifying U.S. offering and one selected Chinese example. Market access — one qualifying U.S. offering. No additional qualifying international market-access offering was identified in the selected inventory. The Optum comparisons are based on customer-task overlap, not identical product functionality.',
});

// ---------- SLIDE 10 — AI SUPPLIERS ----------
async function slide10() {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'AI suppliers' });
  s.addText('AI SUPPLIERS EXPAND INTO LIFE SCIENCES', { placeholder: 'kicker' });
  s.addText('AI technology providers are expanding into pharmaceutical workflows through partnerships, specialized offerings and acquisitions', { placeholder: 'title' });
  label(s, 'A SECOND COMPETITIVE DYNAMIC', L, 1.55, 6);
  tx(s, 'The preceding slides examine conversational AI products marketed by established life-sciences data, technology and services companies. At the same time, AI model and infrastructure providers are pursuing life-sciences customers through several distinct approaches—some complementary to incumbent platforms, others potentially competitive.',
    { x: L, y: 1.8, w: W, h: 0.4, fontSize: 10 });
  label(s, 'THREE OBSERVED APPROACHES TO ENTERING LIFE SCIENCES', L, 2.28, 8);
  // Row 1 (partnering through vendors) tinted; rows 2–3 (more direct routes to pharma customers) navy
  const bottom = taskTable(s, 2.53, [
    ['Partner with established life-sciences vendors', ['**Anthropic + Komodo Health:** A documented partnership supporting AI-enabled healthcare analytics;', "**Anthropic + Veeva:** Claude incorporated into Veeva's industry-specific AI offerings;", '**AWS + Datavant:** Explore Assistant, publicly described in February 2026, uses Bedrock and AWS Clean Rooms to support natural-language assessment of healthcare datasets and research feasibility.'],
      'AI suppliers can contribute underlying technology to existing life-sciences platforms while those companies retain specialized data, applications and customer relationships. Komodo and Veeva already appear in our competitive analysis.'],
    ['Develop their own life-sciences-specific products and tools', ["**Anthropic's Claude for Life Sciences:** Specialized scientific tools and integrations supporting research, clinical trials and regulatory workflows;", "**OpenAI's GPT-Rosalind:** A purpose-built model for biology, drug discovery and translational research, with access expanded to eligible organizations in September 2026."],
      "Rather than acting only as technology suppliers, AI companies are building offerings aimed more directly at life-sciences users. These products are not necessarily direct substitutes for Optum's RWE and healthcare-data offerings, but they illustrate potential expansion into adjacent customer workflows."],
    ['Acquire specialized life-sciences capabilities', ['**Cohere + Reliant AI:** In May 2026, Cohere acquired a biopharma AI company offering literature review, competitive landscaping and scientific/regulatory information extraction. The acquisition supports development of North for Pharma, a planned specialized agentic offering.'],
      'Acquisitions provide another way for AI companies to obtain industry-specific technology, expertise and pharmaceutical customer relationships—potentially strengthening their ability to supply or compete with established life-sciences platforms.'],
  ], {
    head: ['AI SUPPLIER APPROACH', 'SELECTED PUBLIC DEVELOPMENTS', 'COMPETITIVE SIGNIFICANCE FOR OPTUM LIFE SCIENCES'], groupAfter: 0, colW: [1.9, 5.85, 4.38], fontSize: 9,
    taskFill: i => (i === 0 ? COL.tint : COL.navy), taskColor: i => (i === 0 ? COL.navy : COL.white),
  });
  const qy = bottom + 0.14, qh = 6.7 - qy;
  if (qh < 0.8) console.warn('slide 10 question card too short:', qh.toFixed(2));
  card(s, L, qy, W, qh, 'strategic question card');
  await iconDot(s, 'MdOutlineLightbulb', L + 0.2, qy + qh / 2 - 0.21, 0.42);
  tx(s, 'THE STRATEGIC QUESTION FOR OPTUM LIFE SCIENCES', { x: L + 0.75, y: qy + 0.1, w: 1.4, h: qh - 0.2, fontSize: 8.5, bold: true, color: COL.orange, charSpacing: 1, valign: 'middle', objectName: 'label' });
  tx(s, [
    B('As AI providers expand beyond supplying general-purpose models, where will they remain technology partners to established life-sciences vendors—and where might they increasingly compete for the same pharmaceutical customers and workflows?', { color: COL.navy, breakLine: true }),
    B('What could these developments mean for the competitive positions of companies whose offerings combine differentiated healthcare data, research tools and scientific expertise?', { color: COL.navy, paraSpaceBefore: 6 }),
  ], { x: L + 2.3, y: qy + 0.08, w: W - 2.5, h: qh - 0.16, fontSize: 10.5, valign: 'middle' });
}

// ---------- SLIDE 11 — METHODOLOGY NOTE (long notes moved off their slides) ----------
async function slide11() {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Methodology note' });
  s.addText('METHODOLOGY NOTE', { placeholder: 'kicker' });
  const notes = [
    ['SLIDE 3 — METHODOLOGY NOTE', [R('Comparable competitors include '), B('direct and specialized/partial offering-level competitors'), R(" serving similar customer tasks; adjacent participants are excluded. Counts represent distinct company groups per workflow, not product totals. A company may appear in multiple workflows. Optum's ○ indicates that a qualifying conversational interface was "), R('not publicly established in this review', { italic: true }), R(', not that the capability is absent. These are selected screened companies, not market-wide adoption rates.')]],
    ['SLIDE 4 — INDIA RECONCILIATION', 'Of 37 screened companies, 36 qualify as direct/specialized comparators in at least one of the five decision workflows; excluding four companies already represented in the qualifying U.S. chat cohort leaves 32. Three additional chat-marketing companies qualify: Axtria, Doceree and Genpact. Company presence in an India screen does not prove the specific product is available or deployed in India.'],
    ['SLIDE 4 — EVIDENCE BOUNDARY', 'Selected country/region screens differ in research depth and company composition and are not representative national market samples. The Nordics screen groups Denmark, Finland, Norway and Sweden. A multinational may appear in multiple country screens; the counts are not additive across countries. A globally marketed product does not establish locally verified use. Percentages show shares within screened comparable companies, not regional adoption, customer penetration or product deployment rates.'],
    ['SLIDE 10 — EVIDENCE BOUNDARY', [R("These are selected publicly documented developments, not an exhaustive assessment of AI suppliers. The three approaches are neither mutually exclusive nor predictions of future competition. They do not establish that AI suppliers caused the recent RWE launch cluster, are displacing established vendors, or expose an Optum capability gap. Not all supplier offerings qualify as conversational products; they are not additions to the deck's product counts. North for Pharma was described as planned; Explore Assistant's general commercial availability was not established in our review.", { italic: true })]],
  ];
  const gap = 0.2, w = (W - gap) / 2, h = (6.6 - 0.9 - gap) / 2;
  notes.forEach(([lab, body], i) => {
    const x = L + (i % 2) * (w + gap), y = 0.9 + Math.floor(i / 2) * (h + gap);
    card(s, x, y, w, h, lab.toLowerCase());
    tx(s, lab, { x: x + 0.25, y: y + 0.22, w: w - 0.5, h: 0.24, fontSize: 10, bold: true, color: COL.orange, charSpacing: 1, objectName: 'label' });
    tx(s, body, { x: x + 0.25, y: y + 0.55, w: w - 0.5, h: h - 0.7, fontSize: 10.5 });
  });
}

(async () => {
  const all = [slide1, slide2, slide3, slide4, slide5, slide6, slide7, slide8, slide9, slide10, slide11];
  const only = (process.env.SLIDES || '1,2,3,4,5,6,7,8,9,10,11').split(',').map(Number);
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
