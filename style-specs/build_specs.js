const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5

const SPECS = [
  {
    id: 'A', name: 'Optum Brand', cue: 'Cue: UnitedHealth Group Investor Conference 2024 book — clean, brand-forward, big numbers, generous space',
    bg: 'FAF8F2', head: 'Arial', headColor: '002677', label: 'FF612B', body: '4B4D4F', muted: '7A7C7E',
    chat: '002677', rest: 'D9D4C7', accent: 'FF612B',
    palette: [['FAF8F2','Canvas (warm white)'],['002677','Navy · headlines, chat bars'],['FF612B','Optum orange · labels, key numbers'],['4B4D4F','Body text'],['D9D4C7','Neutral · remaining bars']],
    tableHead: { fill: '002677', color: 'FFFFFF' }, rowLine: 'D9D4C7',
    metric: 'bigOrange', callout: 'card',
    rules: ['Warm-white canvas; navy headlines; orange only for labels and key numbers', 'White cards with soft shadow group each content block', 'Hero numbers where the locked copy has them (45 · 15 · 19)'],
  },
  {
    id: 'B', name: 'Strategy Consulting', cue: 'Cue: L.E.K. China Clinical CRO — dense, pre-read ready, hairline tables, tracker + draft/confidential footer',
    bg: 'FFFFFF', head: 'Cambria', headColor: '002677', label: '002677', body: '222222', muted: '6B6B6B',
    chat: '002677', rest: 'C9CDD3', accent: 'FF612B',
    palette: [['FFFFFF','Canvas (white)'],['002677','Navy · headlines, labels, chat bars'],['222222','Body text'],['C9CDD3','Hairlines · remaining bars'],['FF612B','Orange · one highlight per slide max']],
    tableHead: { fill: null, color: '002677' }, rowLine: 'C9CDD3',
    metric: 'inline', callout: 'gray',
    rules: ['Serif action titles; every slide reads standalone as a pre-read', 'Hairline tables, no fills; tracker label top-left on every slide', 'Orange used at most once per slide, for the single key point'],
  },
  {
    id: 'C', name: 'Data-Forward', cue: 'Cue: IQVIA Q2 2025, Tempus Q2 2026, Certara (William Blair 2024) — KPI tiles, chart-led, navy statement blocks',
    bg: 'FFFFFF', head: 'Arial', headColor: '1A1A1A', label: '196ECF', body: '333333', muted: '6E7681',
    chat: '196ECF', rest: 'E1E4E8', accent: 'FF612B',
    palette: [['002677','Navy · KPI tiles, statement blocks'],['196ECF','Bright blue · labels, chat bars'],['1A1A1A','Headlines'],['E1E4E8','Light gray · remaining bars'],['FF612B','Orange · highlight']],
    tableHead: { fill: 'E8F1FB', color: '002677' }, rowLine: 'E1E4E8',
    metric: 'tile', callout: 'navy',
    rules: ['Lead with the number: KPI tiles open every data slide', 'Charts dominate; text supports; navy blocks for the takeaway', 'Navy cover and section slides, light content slides'],
  },
];

function txt(slide, text, o) { slide.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: 'top' }, o)); }

for (const s of SPECS) {
  const sl = pres.addSlide();
  sl.background = { color: s.bg };
  const L = 0.6, R = 7.0, CW = 5.75;
  const f = { fontFace: 'Arial' };

  // Board header
  txt(sl, `STYLE SPEC ${s.id}`, { ...f, x: L, y: 0.4, w: 6, h: 0.25, fontSize: 10, bold: true, color: s.label, charSpacing: 1 });
  txt(sl, s.name, { x: L, y: 0.68, w: 9, h: 0.5, fontFace: s.head, fontSize: 26, bold: true, color: s.headColor });
  txt(sl, s.cue, { ...f, x: L, y: 1.2, w: 12.1, h: 0.25, fontSize: 11, color: s.muted });

  const sec = (t, x, y) => txt(sl, t, { ...f, x, y, w: 5, h: 0.22, fontSize: 9, bold: true, color: s.muted, charSpacing: 1 });

  // Palette
  sec('PALETTE', L, 1.7);
  s.palette.forEach(([hex, role], i) => {
    const x = L + i * 1.17;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.97, w: 1.02, h: 0.5, fill: { color: hex }, line: { color: 'BFBFBF', width: 0.5 } });
    txt(sl, hex, { ...f, x, y: 2.52, w: 1.1, h: 0.18, fontSize: 9, bold: true, color: s.body });
    txt(sl, role, { ...f, x, y: 2.71, w: 1.1, h: 0.5, fontSize: 8, color: s.muted });
  });

  // Type
  sec('TYPE', L, 3.35);
  const note = (t, y) => txt(sl, t, { ...f, x: L, y, w: CW, h: 0.18, fontSize: 8, italic: true, color: s.muted });
  txt(sl, 'Recent RWE product launches prompted a broader examination of conversational AI across life sciences', { x: L, y: 3.62, w: CW, h: 1.0, fontFace: s.head, fontSize: 20, bold: true, color: s.headColor });
  note(`Headline · ${s.head} Bold 20–22pt · action title, top-left, fixed position`, 4.66);
  txt(sl, 'KEY FINDINGS', { ...f, x: L, y: 4.95, w: CW, h: 0.2, fontSize: 10, bold: true, color: s.label, charSpacing: 1 });
  txt(sl, 'Optum LS publicly emphasizes linked healthcare data, AI-enabled research tools and scientific expertise.', { ...f, x: L, y: 5.18, w: CW, h: 0.42, fontSize: 12, color: s.body });
  note('Section label · Arial Bold 10pt caps  |  Body · Arial 12pt', 5.62);
  txt(sl, 'Evidence boundary: Public product marketing and announcements reviewed as of October 13, 2026.', { ...f, x: L, y: 5.92, w: CW, h: 0.2, fontSize: 9, italic: true, color: s.muted });
  note('Evidence boundary / footnote · Arial Italic 9pt, bottom of slide', 6.14);

  // Components
  sec('COMPONENTS', R, 1.7);
  // Table snippet
  const cols = [1.25, 1.15, 2.6, 0.75];
  const hdr = ['Workflow', 'Chat / comparable', "Optum's publicly marketed offerings", 'Chat claim'];
  const row = ['RWE', '9 / 34', 'Analytics on Demand; epidemiology research', '○'];
  const th = { fontFace: 'Arial', fontSize: 9, bold: true, color: s.tableHead.color, valign: 'middle' };
  if (s.tableHead.fill) th.fill = { color: s.tableHead.fill };
  const line = { type: 'solid', pt: 0.75, color: s.rowLine };
  const headCells = hdr.map((h, i) => ({ text: h, options: { ...th, align: i >= 1 && i !== 2 ? 'center' : 'left', border: [{ type: 'none' }, { type: 'none' }, s.tableHead.fill ? { type: 'none' } : { type: 'solid', pt: 1.25, color: s.headColor }, { type: 'none' }] } }));
  const bodyCells = row.map((c, i) => ({ text: c, options: { fontFace: 'Arial', fontSize: 10, color: i === 1 ? s.chat : s.body, bold: i <= 1, align: i >= 1 && i !== 2 ? 'center' : 'left', valign: 'middle', border: [{ type: 'none' }, { type: 'none' }, line, { type: 'none' }] } }));
  sl.addTable([headCells, bodyCells], { x: R, y: 1.97, w: 5.75, colW: cols, rowH: [0.32, 0.4], margin: [0, 0.06, 0, 0.06] });

  // Bar snippet
  txt(sl, 'RWE', { ...f, x: R, y: 2.95, w: 0.9, h: 0.3, fontSize: 10, bold: true, color: s.body, valign: 'middle' });
  const bw = 3.9, bx = R + 1.0;
  sl.addShape(pres.shapes.RECTANGLE, { x: bx, y: 2.98, w: bw * 9 / 34, h: 0.24, fill: { color: s.chat }, line: { type: 'none' } });
  sl.addShape(pres.shapes.RECTANGLE, { x: bx + bw * 9 / 34, y: 2.98, w: bw * 25 / 34, h: 0.24, fill: { color: s.rest }, line: { type: 'none' } });
  txt(sl, '9 / 34', { ...f, x: bx + bw + 0.12, y: 2.95, w: 0.8, h: 0.3, fontSize: 10, bold: true, color: s.chat, valign: 'middle' });
  txt(sl, 'Bar · blue = marketing qualifying chat, neutral = remaining comparable companies', { ...f, x: R, y: 3.3, w: 5.75, h: 0.18, fontSize: 8, italic: true, color: s.muted });

  // Headline metric
  if (s.metric === 'bigOrange') {
    txt(sl, '19', { ...f, x: R, y: 3.58, w: 1.2, h: 0.8, fontSize: 48, bold: true, color: s.accent, valign: 'middle' });
    txt(sl, 'qualifying conversational AI offerings across five of six workflows', { ...f, x: R + 1.25, y: 3.68, w: 4.4, h: 0.6, fontSize: 12, color: s.body, valign: 'middle' });
  } else if (s.metric === 'inline') {
    txt(sl, [
      { text: '45', options: { fontFace: 'Cambria', fontSize: 22, bold: true, color: s.headColor } },
      { text: ' competitor groups initially screened  |  ', options: { fontSize: 11, color: s.body } },
      { text: '15', options: { fontFace: 'Cambria', fontSize: 22, bold: true, color: s.headColor } },
      { text: ' distinct groups  |  ', options: { fontSize: 11, color: s.body } },
      { text: '19', options: { fontFace: 'Cambria', fontSize: 22, bold: true, color: s.accent } },
      { text: ' qualifying offerings', options: { fontSize: 11, color: s.body } },
    ], { ...f, x: R, y: 3.62, w: 5.75, h: 0.7, valign: 'middle' });
  } else {
    [['45', 'competitor groups initially screened'], ['15', 'distinct groups marketing qualifying conversational AI'], ['19', 'qualifying conversational AI offerings']].forEach(([n, l], i) => {
      const x = R + i * 1.95;
      sl.addShape(pres.shapes.RECTANGLE, { x, y: 3.58, w: 1.8, h: 0.85, fill: { color: '002677' }, line: { type: 'none' } });
      txt(sl, n, { ...f, x: x + 0.12, y: 3.6, w: 1.6, h: 0.45, fontSize: 26, bold: true, color: i === 2 ? s.accent : 'FFFFFF', valign: 'middle' });
      txt(sl, l, { ...f, x: x + 0.12, y: 4.03, w: 1.6, h: 0.36, fontSize: 8, color: 'FFFFFF' });
    });
  }
  txt(sl, 'Headline metric', { ...f, x: R, y: 4.48, w: 5.75, h: 0.18, fontSize: 8, italic: true, color: s.muted });

  // Callout
  const cy = 4.78, ch = 0.82;
  const ctext = 'RWE has the largest number of screened U.S. companies marketing qualifying conversational AI.';
  if (s.callout === 'card') {
    sl.addShape(pres.shapes.RECTANGLE, { x: R, y: cy, w: 5.75, h: ch, fill: { color: 'FFFFFF' }, line: { type: 'none' }, shadow: { type: 'outer', color: '000000', opacity: 0.12, blur: 6, offset: 2, angle: 90 } });
  } else if (s.callout === 'gray') {
    sl.addShape(pres.shapes.RECTANGLE, { x: R, y: cy, w: 5.75, h: ch, fill: { color: 'F2F3F5' }, line: { type: 'none' } });
  } else {
    sl.addShape(pres.shapes.RECTANGLE, { x: R, y: cy, w: 5.75, h: ch, fill: { color: '002677' }, line: { type: 'none' } });
  }
  const onDark = s.callout === 'navy';
  txt(sl, 'KEY TAKEAWAY', { ...f, x: R + 0.18, y: cy + 0.12, w: 5.4, h: 0.2, fontSize: 9, bold: true, color: onDark ? 'FFFFFF' : s.label, charSpacing: 1 });
  txt(sl, ctext, { ...f, x: R + 0.18, y: cy + 0.36, w: 5.4, h: 0.4, fontSize: 11, color: onDark ? 'FFFFFF' : s.body });

  // Rules
  txt(sl, s.rules.map((r, i) => ({ text: r, options: { bullet: true, breakLine: i < s.rules.length - 1 } })), { ...f, x: R, y: 5.8, w: 5.75, h: 0.75, fontSize: 9.5, color: s.body, paraSpaceAfter: 3 });

  // Page furniture (as it would appear on every slide)
  sl.addShape(pres.shapes.LINE, { x: L, y: 6.75, w: 12.13, h: 0, line: { color: s.rowLine, width: 0.75 } });
  const left = s.id === 'B' ? 'Draft — for discussion  |  Candidate work sample — outside-in research' : 'Candidate work sample — outside-in research';
  txt(sl, left, { ...f, x: L, y: 6.85, w: 8, h: 0.2, fontSize: 9, color: s.muted });
  txt(sl, s.id === 'B' ? 'U.S. COMPETITIVE LANDSCAPE  |  3' : '3', { ...f, x: 8.73, y: 6.85, w: 4, h: 0.2, fontSize: 9, color: s.muted, align: 'right' });
}

pres.writeFile({ fileName: process.argv[2] }).then(() => console.log('ok'));
