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
    s.addShape(pres.shapes.RECTANGLE, { x, y: ty, w: tw, h: th, fill: { color: COL.navy }, line: { type: 'none' }, objectName: 'kpi tile ' + n });
    tx(s, n, { x: x + 0.18, y: ty, w: 0.75, h: th, fontSize: 26, bold: true, color: i === 2 ? COL.orange : COL.white, valign: 'middle' });
    tx(s, t, { x: x + 0.95, y: ty, w: tw - 1.08, h: th, fontSize: 10, color: COL.white, valign: 'middle' });
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
  s.addShape(pres.shapes.RECTANGLE, { x: L, y: hy, w: W, h: hh, fill: { color: COL.tint }, line: { type: 'none' }, objectName: 'table header' });
  const hdr = { y: hy, h: hh, fontSize: 9, bold: true, color: COL.navy, valign: 'middle', align: 'center' };
  tx(s, [R('U.S. competitors marketing chat /', { breakLine: true }), R('comparable competitors')], { ...hdr, x: cX.bars + 0.6, w: 2.45 });
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

(async () => {
  const only = (process.env.SLIDES || '1,3').split(',').map(Number);
  pres.addSection({ title: 'Executive summary' });
  if (only.includes(1)) await slide1();
  pres.addSection({ title: 'U.S. competitive landscape' });
  if (only.includes(3)) await slide3();
  const out = path.resolve(process.argv[2] || 'deck.pptx');
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log('wrote', out);
})();
