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

// ---------- SLIDE 1 — EXECUTIVE SUMMARY ----------
async function slide1() {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'Executive summary' });
  s.addText('EXECUTIVE SUMMARY', { placeholder: 'kicker' });
  s.addText('Conversational AI is being marketed across life sciences workflows, creating new ways to access information, analytics and decision support', { placeholder: 'title' });

  // Outside-in research lens (full width)
  card(s, L, 1.72, W, 1.18, 'lens card');
  await iconDot(s, 'MdTravelExplore', L + 0.25, 1.95, 0.5);
  label(s, 'OUTSIDE-IN RESEARCH LENS', L + 0.95, 1.9, 6);
  tx(s, [
    R('An assessment of publicly-marketed conversational AI products across six core Optum Life Sciences workflows: data access, RWE, market access, clinical trials, commercial intelligence, and medical/regulatory research.', { breakLine: true }),
    R('The U.S. competitive landscape is the primary focus; selected international offerings provide context; and Optum LS is the strategic reference point.'),
  ], { x: L + 0.95, y: 2.15, w: 10.9, h: 0.65, fontSize: 11, paraSpaceAfter: 3 });

  const cy = 3.05, ch = 3.08, cw = 3.91, gap = 0.2;
  const cx = [L, L + cw + gap, L + 2 * (cw + gap)];

  // Key findings
  card(s, cx[0], cy, cw, ch, 'key findings card');
  await iconDot(s, 'MdInsights', cx[0] + 0.25, cy + 0.25, 0.46);
  label(s, 'KEY FINDINGS', cx[0] + 0.85, cy + 0.37, 2.8);
  tx(s, '19', { x: cx[0] + 0.25, y: cy + 0.86, w: 0.85, h: 0.6, fontSize: 36, bold: true, color: COL.orange, objectName: 'hero 19' });
  tx(s, [R('The U.S. screen identified '), B('19 qualifying conversational AI offerings across five of six domains'), R(', concentrated in RWE and commercial intelligence.')],
    { x: cx[0] + 1.12, y: cy + 0.9, w: 2.58, h: 0.98, fontSize: 10.5 });
  s.addShape(pres.shapes.LINE, { x: cx[0] + 0.25, y: cy + 1.97, w: cw - 0.5, h: 0, line: { color: COL.neutral, width: 0.75 } });
  tx(s, '10', { x: cx[0] + 0.25, y: cy + 2.05, w: 0.85, h: 0.6, fontSize: 36, bold: true, color: COL.orange, objectName: 'hero 10' });
  tx(s, [R('The selected international comparison identified '), B('10 qualifying offerings'), R(', providing additional examples of how vendors position conversational interaction.')],
    { x: cx[0] + 1.12, y: cy + 2.09, w: 2.58, h: 0.9, fontSize: 10.5 });

  // Optum positioning
  card(s, cx[1], cy, cw, ch, 'optum positioning card');
  await iconDot(s, 'MdOutlineLayers', cx[1] + 0.25, cy + 0.25, 0.46);
  label(s, 'OPTUM POSITIONING', cx[1] + 0.85, cy + 0.37, 2.8);
  tx(s, [
    R('Optum LS publicly emphasizes linked healthcare data, AI-enabled research tools and scientific expertise. Several competitors prominently market conversational interfaces as the entry point to comparable customer tasks. '),
    B('This reflects differences in public positioning, not necessarily in underlying functionality.', { color: COL.navy }),
  ], { x: cx[1] + 0.25, y: cy + 0.9, w: cw - 0.5, h: 2.0, fontSize: 11 });

  // Strategic implication
  card(s, cx[2], cy, cw, ch, 'strategic implication card');
  await iconDot(s, 'MdOutlineLightbulb', cx[2] + 0.25, cy + 0.25, 0.46);
  label(s, 'STRATEGIC IMPLICATION', cx[2] + 0.85, cy + 0.37, 2.8);
  tx(s, 'The competitive significance of conversational AI depends on the customer task, underlying data and reliability of the results.',
    { x: cx[2] + 0.25, y: cy + 0.9, w: cw - 0.5, h: 0.8, fontSize: 11 });
  tx(s, "Where could conversational interaction add value to Optum's existing offerings?",
    { x: cx[2] + 0.25, y: cy + 1.78, w: cw - 0.5, h: 1.0, fontSize: 14, bold: true, color: COL.navy, objectName: 'strategic question' });

  tx(s, 'Evidence boundary: Public product marketing and announcements reviewed as of October 13, 2026. The findings describe selected screened inventories—not a market census, adoption rates, verified product performance or undisclosed capabilities.',
    { x: L, y: 6.27, w: W, h: 0.4, fontSize: 9, italic: true, color: COL.muted, objectName: 'evidence boundary' });
}

// ---------- SLIDE 3 — U.S. COMPETITIVE LANDSCAPE ----------
async function slide3() {
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: 'U.S. competitive landscape' });
  s.addText('U.S. COMPETITIVE LANDSCAPE', { placeholder: 'kicker' });
  s.addText('Conversational AI is most prominent in RWE among screened U.S. competitors; Optum has relevant offerings across five decision workflows', { placeholder: 'title' });

  const LW = 8.3; // left (chart) zone width
  // U.S. competitive screen — KPI tiles
  label(s, 'U.S. COMPETITIVE SCREEN', L, 1.66, 6);
  const tiles = [['45', 'competitor groups initially screened'], ['15', 'distinct groups marketing qualifying conversational AI'], ['19', 'qualifying conversational AI offerings across five of six workflows']];
  const tw = (LW - 0.3) / 3;
  tiles.forEach(([n, t], i) => {
    const x = L + i * (tw + 0.15);
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.95, w: tw, h: 0.72, fill: { color: COL.navy }, line: { type: 'none' }, objectName: 'kpi tile ' + n });
    tx(s, n, { x: x + 0.15, y: 1.95, w: 0.75, h: 0.72, fontSize: 28, bold: true, color: i === 2 ? COL.orange : COL.white, valign: 'middle' });
    tx(s, t, { x: x + 0.92, y: 1.95, w: tw - 1.02, h: 0.72, fontSize: 9.5, color: COL.white, valign: 'middle' });
  });

  // Competitor coverage and Optum LS positioning — mirrored chart
  label(s, 'COMPETITOR COVERAGE AND OPTUM LS POSITIONING', L, 2.86, 7);
  const rows = [
    ['RWE', 9, 34, 'Analytics on Demand; epidemiology research'],
    ['Commercial intelligence', 3, 23, 'Market Clarity: portfolio and patient/prescriber analytics'],
    ['Medical / regulatory research', 3, 18, 'Postapproval safety studies; Market Clarity medical affairs insights'],
    ['Clinical trials', 1, 25, 'Prospector; Digital Research Network; True Source'],
    ['Market access', 1, 24, 'HEOR/value evidence; Evidence Engine'],
  ];
  const hy = 3.13, hh = 0.4, ry = hy + hh, rh = 0.46;
  const colX = { bars: L, wf: 4.0, opt: 5.95, chat: 8.0 }, colW = { bars: 3.3, wf: 1.85, opt: 2.0, chat: 0.9 };
  s.addShape(pres.shapes.RECTANGLE, { x: L, y: hy, w: LW, h: hh, fill: { color: COL.navy }, line: { type: 'none' }, objectName: 'table header' });
  const hdr = { y: hy, h: hh, fontSize: 8.5, bold: true, color: COL.white, valign: 'middle' };
  tx(s, 'U.S. competitors marketing chat / comparable competitors', { ...hdr, x: colX.bars + 0.1, w: colW.bars - 0.2, align: 'right' });
  tx(s, 'Workflow', { ...hdr, x: colX.wf, w: colW.wf, align: 'center' });
  tx(s, "Optum's publicly marketed offerings", { ...hdr, x: colX.opt, w: colW.opt });
  tx(s, 'Optum public chat claim', { ...hdr, x: colX.chat, w: colW.chat, align: 'center' });

  // Native stacked bar chart extending left from the workflow spine
  const chartX = 1.3, chartW = 2.6, maxV = 34;
  s.addChart(pres.charts.BAR, [
    { name: 'Marketing qualifying chat', labels: rows.map(r => r[0]), values: rows.map(r => r[1]) },
    { name: 'Remaining comparable', labels: rows.map(r => r[0]), values: rows.map(r => r[2] - r[1]) },
  ], {
    x: chartX, y: ry, w: chartW, h: rh * rows.length, barDir: 'bar', barGrouping: 'stacked', barGapWidthPct: 85,
    chartColors: [HEX.navy, HEX.neutral], valAxisMinVal: 0, valAxisMaxVal: maxV, valAxisOrientation: 'maxMin', catAxisOrientation: 'maxMin',
    valAxisHidden: true, catAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
    showLegend: false, showValue: false, layout: { x: 0, y: 0, w: 1, h: 1 }, objectName: 'competitor coverage chart',
  });
  rows.forEach(([wf, n, d, opt], i) => {
    const y = ry + i * rh;
    const barLeft = chartX + chartW * (1 - d / maxV);
    tx(s, [B(`${n}`, { color: COL.navy }), R(` / ${d}`)], { x: barLeft - 0.7, y, w: 0.62, h: rh, fontSize: 10, align: 'right', valign: 'middle', objectName: 'fraction ' + wf });
    tx(s, wf, { x: colX.wf + 0.05, y, w: colW.wf - 0.1, h: rh, fontSize: 10, bold: true, color: COL.navy, align: 'center', valign: 'middle' });
    tx(s, opt, { x: colX.opt, y, w: colW.opt, h: rh, fontSize: 9, valign: 'middle' });
    s.addShape(pres.shapes.OVAL, { x: colX.chat + colW.chat / 2 - 0.09, y: y + rh / 2 - 0.09, w: 0.18, h: 0.18, fill: { type: 'none' }, line: { color: COL.navy, width: 1.25 }, objectName: 'chat status ' + wf });
    s.addShape(pres.shapes.LINE, { x: L, y: y + rh, w: LW, h: 0, line: { color: COL.neutral, width: 0.75 } });
  });

  // Right column: representative offerings + key takeaway
  const rx = L + LW + 0.25, rw = W - LW - 0.25;
  card(s, rx, 1.66, rw, 2.6, 'representative offerings card');
  tx(s, 'REPRESENTATIVE CONVERSATIONAL AI OFFERINGS', { x: rx + 0.2, y: 1.82, w: rw - 0.4, h: 0.4, fontSize: 10, bold: true, color: COL.orange, charSpacing: 1, objectName: 'label' });
  const reps = [
    ['RWE:', ' Komodo Marmot and MapAI; Truveta Intelligence; Flatiron Telescope; Tempus Lens'],
    ['Commercial intelligence:', ' IQVIA Field Force Agent and GMI Agent; ZS Conversational AI for Commercial; Trinity InsightsEDGE Digital Twins'],
    ['Medical / regulatory:', ' IQVIA Med-R1; ZS Conversational AI for Medical; Veeva Vault AI Chat'],
    ['Clinical trials:', ' TriNetX Terrain AI Assistant'],
    ['Market access:', ' EVERSANA Ask NAVLIN'],
  ];
  const repRuns = [];
  reps.forEach(([k, v], i) => { repRuns.push(B(k, { color: COL.navy })); repRuns.push(R(v, { breakLine: i < reps.length - 1 })); });
  tx(s, repRuns, { x: rx + 0.2, y: 2.3, w: rw - 0.4, h: 1.85, fontSize: 9, paraSpaceAfter: 3 });

  card(s, rx, 4.38, rw, 1.52, 'key takeaway card');
  label(s, 'KEY TAKEAWAY', rx + 0.2, 4.53, rw - 0.4);
  tx(s, 'RWE has the largest number of screened U.S. companies marketing qualifying conversational AI. Optum offers products and services across the five decision workflows shown, but a qualifying customer-facing conversational interface was not established in the reviewed public materials.',
    { x: rx + 0.2, y: 4.8, w: rw - 0.4, h: 1.0, fontSize: 9.5 });

  // Scope note + methodology note (footnotes)
  tx(s, [
    B('Scope note: Data access/enrichment is a cross-workflow enabling category. Zero qualifying chat offerings assigned primarily to this domain does not imply an absence of conversational data-access capabilities in products classified under RWE or other workflows.', { breakLine: true }),
    B('METHODOLOGY NOTE  '),
    R('Comparable competitors include '), B('direct and specialized/partial offering-level competitors'),
    R(" serving similar customer tasks; adjacent participants are excluded. Counts represent distinct company groups per workflow, not product totals. A company may appear in multiple workflows. Optum's ○ indicates that a qualifying conversational interface was "),
    R('not publicly established in this review', { italic: true }),
    R(', not that the capability is absent. These are selected screened companies, not market-wide adoption rates.'),
  ], { x: L, y: 5.97, w: W, h: 0.75, fontSize: 8, color: COL.muted, paraSpaceAfter: 2, objectName: 'scope and methodology notes' });
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
