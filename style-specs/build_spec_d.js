const pptxgen = require('pptxgenjs');
const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const sharp = require('sharp');
const { MdTravelExplore } = require('react-icons/md');

const NAVY = '002677', ORANGE = 'FF612B', BODY = '333333', MUTED = '6B6E73', LINE = 'C9CDD3', TINT = 'E6ECF5';
async function icon(C, color) {
  const svg = renderToStaticMarkup(React.createElement(C, { color: '#' + color, size: 256 }));
  return 'image/png;base64,' + (await sharp(Buffer.from(svg)).png().toBuffer()).toString('base64');
}

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  const sl = pres.addSlide();
  sl.background = { color: 'FFFFFF' };
  const A = { fontFace: 'Arial', isTextBox: true, margin: 0, valign: 'top' };
  const t = (text, o) => sl.addText(text, Object.assign({}, A, o));
  const L = 0.6, R = 7.0, CW = 5.75;
  const sec = (s, x, y) => t(s, { x, y, w: 5, h: 0.22, fontSize: 9, bold: true, color: MUTED, charSpacing: 1 });
  const note = (s, x, y, w) => t(s, { x, y, w: w || CW, h: 0.18, fontSize: 8, italic: true, color: MUTED });

  // Board header
  t('STYLE SPEC D', { x: L, y: 0.4, w: 6, h: 0.25, fontSize: 10, bold: true, color: ORANGE, charSpacing: 1 });
  t('Lookbook Synthesis', { x: L, y: 0.68, w: 9, h: 0.5, fontSize: 26, bold: true, color: NAVY });
  t('Draws on: L.E.K. (action titles, hairline, filled row labels, Note footer) · UHG IC24 (orange kicker, icon rows, hero numbers) · IQVIA (labeled bars, hairline tables)', { x: L, y: 1.2, w: 12.1, h: 0.25, fontSize: 11, color: MUTED });

  // Palette
  sec('PALETTE', L, 1.7);
  [['FFFFFF', 'Canvas (white)'], [NAVY, 'Navy · titles, row labels, chat bars'], [ORANGE, 'Orange · section kicker only'], [BODY, 'Body text'], [LINE, 'Hairlines · remaining bars'], [TINT, 'Tint · icon circles, panels']].forEach(([hex, role], i) => {
    const x = L + i * 0.97;
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.97, w: 0.85, h: 0.5, fill: { color: hex }, line: { color: 'BFBFBF', width: 0.5 } });
    t(hex, { x, y: 2.52, w: 0.95, h: 0.18, fontSize: 9, bold: true, color: BODY });
    t(role, { x, y: 2.71, w: 0.92, h: 0.5, fontSize: 8, color: MUTED });
  });

  // Type: a mini slide header
  sec('TYPE', L, 3.35);
  t('RESEARCH CATALYST', { x: L, y: 3.62, w: CW, h: 0.2, fontSize: 10, bold: true, color: ORANGE, charSpacing: 1 });
  t('Recent RWE product launches prompted a broader examination of conversational AI across life sciences', { x: L, y: 3.88, w: CW, h: 1.05, fontSize: 20, bold: true, color: NAVY });
  sl.addShape(pres.shapes.LINE, { x: L, y: 5.0, w: CW, h: 0, line: { color: LINE, width: 0.75 } });
  note('Kicker · Arial Bold 10pt orange caps  |  Action title · Arial Bold 20–22pt navy, max 2 lines  |  full-width gray hairline', L, 5.06);
  t('Optum LS publicly emphasizes linked healthcare data, AI-enabled research tools and scientific expertise.', { x: L, y: 5.35, w: CW, h: 0.45, fontSize: 12, color: BODY });
  note('Body · Arial 12pt', L, 5.82);
  t('Evidence boundary: Dates reflect selected public announcements or releases, not necessarily first feature availability.', { x: L, y: 6.1, w: CW, h: 0.3, fontSize: 8, color: MUTED });
  note('Evidence boundary · Arial 8–9pt, Note-style above footer (L.E.K. / IQVIA)', L, 6.44);

  // Components
  sec('COMPONENTS', R, 1.7);
  // L.E.K.-style task table
  const hdr = { fontFace: 'Arial', fontSize: 8, bold: true, color: MUTED, valign: 'bottom', border: [{ type: 'none' }, { type: 'none' }, { type: 'solid', pt: 0.75, color: LINE }, { type: 'none' }] };
  const cell = { fontFace: 'Arial', fontSize: 9, color: BODY, valign: 'middle', border: [{ type: 'none' }, { type: 'none' }, { type: 'solid', pt: 0.75, color: LINE }, { type: 'none' }] };
  sl.addTable([
    [{ text: 'CUSTOMER TASK', options: hdr }, { text: 'COMPETITOR CONVERSATIONAL AI OFFERINGS', options: hdr }, { text: 'RELEVANT OPTUM OFFERINGS', options: hdr }],
    [{ text: 'Explore patient populations and cohorts', options: { ...cell, bold: true, color: 'FFFFFF', fill: { color: NAVY }, border: [{ type: 'solid', pt: 2, color: 'FFFFFF' }, { type: 'solid', pt: 2, color: 'FFFFFF' }, { type: 'solid', pt: 2, color: 'FFFFFF' }, { type: 'none' }] } },
     { text: [{ text: "Komodo Health's " }, { text: 'Marmot', options: { bold: true } }, { text: ' and ' }, { text: 'MapAI', options: { bold: true } }, { text: ': conversational healthcare analytics and cohort exploration' }], options: cell },
     { text: 'Analytics on Demand; Clinical Notes Lab; Prospector', options: cell }],
  ], { x: R, y: 1.97, w: CW, colW: [1.55, 2.6, 1.6], rowH: [0.3, 0.62], margin: [0.03, 0.08, 0.03, 0.08] });
  note('Task table · navy filled row labels, hairline rows, no column fills (L.E.K.)', R, 2.95);

  // IQVIA-style labeled bar
  t('RWE', { x: R, y: 3.25, w: 0.9, h: 0.3, fontSize: 10, bold: true, color: BODY, valign: 'middle' });
  const bx = R + 1.0, bw = 3.9;
  sl.addShape(pres.shapes.RECTANGLE, { x: bx, y: 3.28, w: bw * 9 / 34, h: 0.24, fill: { color: NAVY }, line: { type: 'none' } });
  sl.addShape(pres.shapes.RECTANGLE, { x: bx + bw * 9 / 34, y: 3.28, w: bw * 25 / 34, h: 0.24, fill: { color: LINE }, line: { type: 'none' } });
  t('9 / 34', { x: bx + bw + 0.12, y: 3.25, w: 0.8, h: 0.3, fontSize: 10, bold: true, color: NAVY, valign: 'middle' });
  note('Bar · navy = marketing qualifying chat, gray = remaining comparable; label every bar (IQVIA)', R, 3.6);

  // IC24-style icon row
  sl.addShape(pres.shapes.OVAL, { x: R, y: 3.95, w: 0.46, h: 0.46, fill: { color: TINT }, line: { type: 'none' } });
  sl.addImage({ data: await icon(MdTravelExplore, NAVY), x: R + 0.1, y: 4.05, w: 0.26, h: 0.26 });
  t('OUTSIDE-IN RESEARCH LENS', { x: R + 0.62, y: 3.95, w: 5.1, h: 0.2, fontSize: 9, bold: true, color: NAVY, charSpacing: 1 });
  t('The U.S. competitive landscape is the primary focus; selected international offerings provide context; and Optum LS is the strategic reference point.', { x: R + 0.62, y: 4.17, w: 5.1, h: 0.4, fontSize: 9.5, color: BODY });
  sl.addShape(pres.shapes.LINE, { x: R, y: 4.67, w: CW, h: 0, line: { color: LINE, width: 0.75 } });
  note('Icon row · navy icon in tint circle, hairline separators (IC24)', R, 4.73);

  // IC24-style hero numbers
  [['45', 'competitor groups initially screened'], ['15', 'distinct groups marketing qualifying conversational AI'], ['19', 'qualifying conversational AI offerings across five of six workflows']].forEach(([n, l], i) => {
    const x = R + i * 1.95;
    if (i) sl.addShape(pres.shapes.LINE, { x: x - 0.12, y: 5.05, w: 0, h: 0.62, line: { color: LINE, width: 0.75 } });
    t(n, { x, y: 5.0, w: 1.8, h: 0.42, fontSize: 26, bold: true, color: NAVY });
    t(l, { x, y: 5.42, w: 1.75, h: 0.32, fontSize: 8, color: BODY });
  });
  note('Hero numbers · bold navy, small label, hairline dividers (IC24)', R, 5.9);

  // Rules
  t([
    { text: 'White canvas throughout; navy carries data; orange appears once per slide, as the kicker', options: { bullet: true, breakLine: true } },
    { text: 'Every slide: kicker → action title → hairline → body → evidence boundary → footer', options: { bullet: true } },
  ], { x: R, y: 6.16, w: CW, h: 0.5, fontSize: 9, color: BODY, paraSpaceAfter: 2 });

  // Footer (L.E.K. ordering)
  sl.addShape(pres.shapes.LINE, { x: L, y: 6.78, w: 12.13, h: 0, line: { color: LINE, width: 0.75 } });
  t('3  |  Candidate work sample — outside-in research', { x: L, y: 6.88, w: 8, h: 0.2, fontSize: 9, color: MUTED });

  await pres.writeFile({ fileName: 'spec-d.pptx' });
  console.log('ok');
})();
