'use strict';
const P = require('pptxgenjs');
const pptx = new P();
pptx.defineLayout({name:'W',width:13.33,height:7.5});
pptx.layout='W';

// ===== DESIGN SYSTEM =====
const C = {
  navy: '1E2A38',    // primary deep navy
  blue: '3A6EA5',    // secondary policy blue
  orange: 'D98E04',  // accent rational orange
  bg: 'F7F9FC',      // background light gray
  white: 'FFFFFF',
  text: '1E2A38',
  gray: '8899AA',
  lg: 'B0B8C4',
  lighter: 'E5E9F0',
  dark: '0D1B2A',
  b1: '3A6EA5', b2: '5C9BD4', b3: '6DA8E0', b4: '7BB5E6',
};

const TF = 'Source Han Sans SC';
const BF = 'Source Han Sans SC';

const D = require('./data.json');

// Grid constants
const ML = 0.65, MR = 0.65, MT = 0.4, MB = 0.4;
const CW = (13.33 - ML - MR) / 12; // column width

function col(n, s) { return { x: ML + n * CW, w: (s || 1) * CW }; }

function bg(s) { s.background = { fill: C.bg }; }

function pageTitle(s, t, sub) {
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 0.03, fill: { color: C.blue } });
  s.addText(t, { x: ML, y: 0.2, w: 12.03, h: 0.55, fontSize: 24, fontFace: TF, color: C.navy, bold: true });
  if (sub) {
    s.addText(sub, { x: ML, y: 0.72, w: 12.03, h: 0.22, fontSize: 10, fontFace: BF, color: C.gray });
  }
}

function footer(s, n) {
  s.addText(`哈尔滨市第三中学模拟政协提案 · ${n}/23`, {
    x: ML, y: 7.05, w: 12.03, h: 0.3, fontSize: 7.5, fontFace: BF, color: C.lg, align: 'right'
  });
}

function sourceLine(s, text) {
  s.addText(text, { x: ML, y: 6.55, w: 12.03, h: 0.3, fontSize: 7.5, fontFace: BF, color: C.lg });
}

function card(s, x, y, w, h) {
  s.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, fill: { color: C.white },
    line: { color: C.lighter, width: 0.5 },
    rectRadius: 0.05
  });
}

function thinRule(s, x, y, w, color) {
  s.addShape(pptx.ShapeType.rect, { x, y, w: w || 12.03, h: 0.01, fill: { color: color || C.lighter } });
}

const D1 = 'Source Han Sans SC';

// ===== P1: COVER =====
(function() {
  const s = pptx.addSlide();
  s.background = { fill: C.white };
  // Large faint circle top-right
  s.addShape(pptx.ShapeType.ellipse, { x: 8.5, y: -3, w: 8, h: 8, fill: { color: C.navy, transparency: 96 } });
  s.addShape(pptx.ShapeType.ellipse, { x: 10, y: 4, w: 6, h: 6, fill: { color: C.blue, transparency: 94 } });
  // Left accent bar
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.06, h: 7.5, fill: { color: C.blue } });
  // City silhouette - faint buildings at bottom
  const bld = [0.6, 1.0, 0.7, 1.4, 0.5, 1.1, 0.8, 1.3, 0.6, 0.9, 0.4, 1.2, 0.7, 1.0, 0.5, 0.8, 0.6, 1.1];
  bld.forEach((h, i) => {
    s.addShape(pptx.ShapeType.rect, {
      x: 0.3 + i * 0.72, y: 6.8 - h * 0.5,
      w: 0.6, h: h * 0.5,
      fill: { color: C.navy, transparency: 92 }
    });
  });
  // Horizontal accent
  s.addShape(pptx.ShapeType.rect, { x: ML, y: 1.6, w: 2.5, h: 0.03, fill: { color: C.orange } });
  // Title
  s.addText(D.p1.title, { x: ML, y: 1.9, w: 9, h: 1.6, fontSize: 36, fontFace: TF, color: C.navy, bold: true, lineSpacing: 42 });
  // Subtitle
  s.addText(D.p1.subtitle, { x: ML, y: 3.6, w: 10, h: 0.35, fontSize: 14, fontFace: BF, color: C.blue });
  // Info
  s.addText(D.p1.group, { x: ML, y: 4.2, w: 8, h: 0.3, fontSize: 12, fontFace: BF, color: C.text });
  s.addText(D.p1.members, { x: ML, y: 4.55, w: 8, h: 0.3, fontSize: 11, fontFace: BF, color: C.gray });
  s.addText(D.p1.teacher, { x: ML, y: 4.9, w: 5, h: 0.3, fontSize: 11, fontFace: BF, color: C.gray });
  s.addText(D.p1.date, { x: ML, y: 5.3, w: 3, h: 0.3, fontSize: 11, fontFace: BF, color: C.lg });
  // Emblem
  s.addShape(pptx.ShapeType.ellipse, { x: 10.8, y: 1.0, w: 1.8, h: 1.8, fill: { color: C.white }, line: { color: C.blue, width: 1.5 } });
  s.addText('模拟\n政协', { x: 10.8, y: 1.15, w: 1.8, h: 1.5, fontSize: 16, fontFace: TF, color: C.blue, align: 'center', bold: true, valign: 'middle' });
  s.addNotes(D.p1.note);
})();

// ===== BUILD HELPERS =====
const SLIDE_IDS = ['p2','p3','p4','p5','p6','p7','p8','p9','p10','p11','p12','p13','p14','p15','p16','p17','p18','p19','p20','p21','p22','p23'];

function buildSlide(id) {
  const d = D[id];
  const s = pptx.addSlide();
  bg(s);
  pageTitle(s, d.title, d.sub);

  switch (id) {
    case 'p2': buildP2(s, d); break;
    case 'p3': buildP3(s, d); break;
    case 'p4': buildP4(s, d); break;
    case 'p5': buildP5(s, d); break;
    case 'p6': buildP6(s, d); break;
    case 'p7': buildP7(s, d); break;
    case 'p8': buildP8(s, d); break;
    case 'p9': buildP9(s, d); break;
    case 'p10': buildP10(s, d); break;
    case 'p11': buildP11(s, d); break;
    case 'p12': buildP12(s, d); break;
    case 'p13': buildP13(s, d); break;
    case 'p14': buildP14(s, d); break;
    case 'p15': buildP15(s, d); break;
    case 'p16': buildP16(s, d); break;
    case 'p17': buildP17(s, d); break;
    case 'p18': buildP18(s, d); break;
    case 'p19': buildP19(s, d); break;
    case 'p20': buildP20(s, d); break;
    case 'p21': buildP21(s, d); break;
    case 'p22': buildP22(s, d); break;
    case 'p23': buildP23(s, d); break;
  }

  const pn = parseInt(id.substring(1));
  footer(s, pn);
  if (d.note) s.addNotes(d.note);
}

// ===== P2: Policy Background =====
function buildP2(s, d) {
  // Four policy timeline cards
  d.items.forEach((it, i) => {
    const x = ML + i * 3.0;
    // Year badge
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.15, w: 0.9, h: 0.3, fill: { color: C.blue }, rectRadius: 0.04 });
    s.addText(it.y, { x, y: 1.15, w: 0.9, h: 0.3, fontSize: 10, fontFace: BF, color: C.white, align: 'center', bold: true, valign: 'middle' });
    if (i < 3) {
      s.addText('→', { x: x + 0.9, y: 1.15, w: 0.5, h: 0.3, fontSize: 12, color: C.lg, align: 'center', valign: 'middle' });
    }
    // Card
    card(s, x, 1.6, 2.7, 2.2);
    s.addText(it.t, { x: x + 0.12, y: 1.68, w: 2.46, h: 0.25, fontSize: 11, fontFace: BF, color: C.navy, bold: true });
    s.addText(it.d, { x: x + 0.12, y: 1.95, w: 2.46, h: 1.0, fontSize: 8, fontFace: BF, color: C.text, lineSpacing: 12 });
    s.addText(it.n, { x: x + 0.12, y: 2.95, w: 2.46, h: 0.3, fontSize: 8, fontFace: BF, color: C.blue, bold: true });
    s.addText(it.r, { x: x + 0.12, y: 3.25, w: 2.46, h: 0.25, fontSize: 6.5, fontFace: BF, color: C.lg });
  });
  // Tags
  thinRule(s, ML, 4.05);
  s.addText('政策关键词', { x: ML, y: 4.15, w: 3, h: 0.25, fontSize: 10, fontFace: BF, color: C.navy, bold: true });
  d.tags.forEach((t, i) => {
    const x = ML + i * 1.55;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 4.45, w: 1.4, h: 0.3, fill: { color: C.white }, line: { color: C.blue, width: 0.3 }, rectRadius: 0.04 });
    s.addText(t, { x, y: 4.45, w: 1.4, h: 0.3, fontSize: 7.5, fontFace: BF, color: C.blue, align: 'center', valign: 'middle' });
  });
  // Note at bottom
  s.addText(d.footer, { x: ML, y: 5.4, w: 12.03, h: 0.7, fontSize: 8.5, fontFace: BF, color: C.gray });
}

// ===== P3: Epidemiology =====
function buildP3(s, d) {
  // 4 stat items on left
  d.items.forEach((item, i) => {
    const y = 1.15 + i * 1.25;
    card(s, ML, y, 6.8, 1.05);
    // Big number
    s.addText(item.p, { x: ML + 0.15, y: y + 0.08, w: 1.2, h: 0.55, fontSize: 26, fontFace: TF, color: C.blue, align: 'center', bold: true });
    // Label and description
    s.addText(item.l, { x: ML + 1.4, y: y + 0.05, w: 5.1, h: 0.3, fontSize: 11, fontFace: BF, color: C.navy, bold: true });
    s.addText(item.n, { x: ML + 1.4, y: y + 0.35, w: 5.1, h: 0.25, fontSize: 8, fontFace: BF, color: C.text });
    s.addText(item.r, { x: ML + 1.4, y: y + 0.65, w: 5.1, h: 0.22, fontSize: 6.5, fontFace: BF, color: C.lg });
  });
  // Right side: policy timeline
  card(s, 7.7, 1.15, 5.0, 5.0);
  s.addText('政策关注度演变', { x: 7.85, y: 1.25, w: 4.7, h: 0.3, fontSize: 11, fontFace: BF, color: C.navy, bold: true });
  d.right.forEach((t, i) => {
    const y = 1.7 + i * 0.85;
    s.addShape(pptx.ShapeType.ellipse, { x: 7.95, y: y + 0.08, w: 0.12, h: 0.12, fill: { color: C.blue } });
    if (i < 4) s.addShape(pptx.ShapeType.rect, { x: 8.01, y: y + 0.2, w: 0.01, h: 0.6, fill: { color: C.lighter } });
    s.addText(t.y, { x: 8.2, y: y - 0.02, w: 0.55, h: 0.28, fontSize: 8, fontFace: BF, color: C.blue, bold: true });
    s.addText(t.t, { x: 8.2, y: y + 0.25, w: 4.2, h: 0.45, fontSize: 7.5, fontFace: BF, color: C.text, lineSpacing: 11 });
  });
}

// ===== P4: Methodology =====
function buildP4(s, d) {
  // 5 steps in a row
  d.steps.forEach((st, i) => {
    const x = ML + i * 2.4;
    // Step circle
    s.addShape(pptx.ShapeType.ellipse, { x: x + 0.6, y: 1.15, w: 0.85, h: 0.85, fill: { color: i === 0 ? C.blue : i < 3 ? C.b2 : C.b3 } });
    s.addText(String(i + 1), { x: x + 0.6, y: 1.2, w: 0.85, h: 0.75, fontSize: 18, color: C.white, align: 'center', bold: true, valign: 'middle', fontFace: BF });
    if (i < 4) s.addText('→', { x: x + 1.45, y: 1.35, w: 0.4, h: 0.45, fontSize: 13, color: C.lg, align: 'center', valign: 'middle' });
    s.addText(st.l, { x: x, y: 2.1, w: 2.0, h: 0.25, fontSize: 10, fontFace: BF, color: C.navy, align: 'center', bold: true });
    s.addText(st.t, { x: x, y: 2.35, w: 2.0, h: 0.2, fontSize: 7.5, fontFace: BF, color: C.gray, align: 'center' });
    s.addText(st.d, { x: x + 0.1, y: 2.55, w: 1.8, h: 0.4, fontSize: 7.5, fontFace: BF, color: C.text, align: 'center', lineSpacing: 11 });
  });
  // Stats row
  thinRule(s, ML, 3.2);
  d.stats.forEach((st, i) => {
    const x = ML + i * 3.0;
    s.addText(st.n, { x, y: 3.35, w: 2.5, h: 0.4, fontSize: 22, fontFace: TF, color: C.blue, align: 'center', bold: true });
    s.addText(st.l + (st.s ? ' | ' + st.s : ''), { x, y: 3.75, w: 2.5, h: 0.25, fontSize: 8, fontFace: BF, color: C.gray, align: 'center' });
  });
  // Fieldwork & limitations side by side
  card(s, ML, 4.2, 6.0, 2.2);
  s.addShape(pptx.ShapeType.rect, { x: ML, y: 4.2, w: 0.05, h: 2.2, fill: { color: C.blue } });
  s.addText('调研现场记录', { x: ML + 0.2, y: 4.28, w: 3, h: 0.25, fontSize: 9.5, fontFace: BF, color: C.navy, bold: true });
  s.addText(d.fieldwork, { x: ML + 0.2, y: 4.58, w: 5.5, h: 1.6, fontSize: 7.5, fontFace: BF, color: C.text, lineSpacing: 13 });

  card(s, 6.9, 4.2, 5.8, 2.2);
  s.addShape(pptx.ShapeType.rect, { x: 6.9, y: 4.2, w: 0.05, h: 2.2, fill: { color: C.orange } });
  s.addText('调研局限性说明', { x: 7.1, y: 4.28, w: 5.4, h: 0.25, fontSize: 9.5, fontFace: BF, color: C.navy, bold: true });
  s.addText(d.limitations, { x: 7.1, y: 4.58, w: 5.4, h: 1.6, fontSize: 7.5, fontFace: BF, color: C.gray, lineSpacing: 13 });
}

// ===== P5: Findings =====
function buildP5(s, d) {
  // 3 key findings as large stat cards
  d.cards.forEach((item, i) => {
    const x = ML + i * 4.1;
    card(s, x, 1.15, 3.8, 1.7);
    s.addText(item.p, { x, y: 1.2, w: 3.8, h: 0.6, fontSize: 28, fontFace: TF, color: C.blue, align: 'center', bold: true });
    s.addText(item.l, { x: x + 0.15, y: 1.8, w: 3.5, h: 0.25, fontSize: 11, fontFace: BF, color: C.navy, align: 'center', bold: true });
    s.addText(item.d, { x: x + 0.15, y: 2.08, w: 3.5, h: 0.45, fontSize: 7.5, fontFace: BF, color: C.gray, align: 'center', lineSpacing: 11 });
    s.addText(item.r, { x: x + 0.15, y: 2.55, w: 3.5, h: 0.2, fontSize: 6.5, fontFace: BF, color: C.lg, align: 'center' });
  });
  // Bar chart - clean horizontal bars
  thinRule(s, ML, 3.1);
  s.addText('主要心理困扰分布（多选）', { x: ML, y: 3.2, w: 5, h: 0.25, fontSize: 10, fontFace: BF, color: C.navy, bold: true });
  const barColors = [C.blue, C.b2, C.b3, C.b4, C.navy, C.orange, C.gray];
  d.bars.forEach((b, i) => {
    const y = 3.55 + i * 0.36;
    s.addText(b.l, { x: ML, y: y, w: 1.5, h: 0.28, fontSize: 8, fontFace: BF, color: C.text, align: 'right' });
    // Background bar
    s.addShape(pptx.ShapeType.rect, { x: 2.3, y: y + 0.06, w: 5.5, h: 0.16, fill: { color: C.lighter } });
    // Data bar
    s.addShape(pptx.ShapeType.rect, { x: 2.3, y: y + 0.06, w: 5.5 * (b.p / 100), h: 0.16, fill: { color: barColors[i % barColors.length] }, rectRadius: 0.02 });
    s.addText(b.p + '%', { x: 2.3 + 5.5 * (b.p / 100) + 0.08, y: y, w: 0.6, h: 0.28, fontSize: 7.5, fontFace: BF, color: C.navy, bold: true });
  });
  sourceLine(s, '来源：本次调研 N=672');
  // Validation note
  s.addShape(pptx.ShapeType.roundRect, { x: ML, y: 6.05, w: 12.03, h: 0.35, fill: { color: C.white }, line: { color: C.lighter, width: 0.3 }, rectRadius: 0.04 });
  s.addText(d.validation, { x: ML + 0.15, y: 6.05, w: 11.73, h: 0.35, fontSize: 7.5, fontFace: BF, color: C.navy, valign: 'middle' });
}

// ===== P6: Student Voices =====
function buildP6(s, d) {
  s.background = { fill: C.dark };
  // Thin accent line
  s.addShape(pptx.ShapeType.rect, { x: ML, y: 0.6, w: 0.02, h: 5.5, fill: { color: C.orange, transparency: 30 } });
  s.addText('来自学生的真实声音', { x: ML + 0.3, y: 0.25, w: 10, h: 0.35, fontSize: 14, fontFace: TF, color: C.orange, bold: true });
  d.quotes.forEach((q, i) => {
    const y = 0.85 + i * 1.05;
    s.addShape(pptx.ShapeType.roundRect, { x: ML + 0.3, y, w: 11.2, h: 0.8, fill: { color: '162230' }, line: { color: C.orange, width: 0.3, transparency: 60 }, rectRadius: 0.06 });
    s.addText('“' + q + '”', { x: ML + 0.6, y: y + 0.05, w: 10.6, h: 0.7, fontSize: 12, fontFace: BF, color: C.white, valign: 'middle', lineSpacing: 18 });
  });
  s.addText(d.footer, { x: ML, y: 6.1, w: 12.03, h: 0.3, fontSize: 7.5, fontFace: BF, color: C.gray, align: 'right' });
  // Custom footer for dark slide
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 7.1, w: 13.33, h: 0.4, fill: { color: '162230' } });
  s.addText('哈尔滨市第三中学模拟政协提案 · 6/23', { x: ML, y: 7.1, w: 12.03, h: 0.4, fontSize: 7.5, color: C.gray, fontFace: BF, align: 'right', valign: 'middle' });
}

// ===== P7: Harbin Context =====
function buildP7(s, d) {
  // 4 cards in 2x2
  d.cards.forEach((c, i) => {
    const x = ML + (i % 2) * 6.2;
    const y = 1.15 + Math.floor(i / 2) * 2.1;
    card(s, x, y, 5.9, 1.9);
    s.addShape(pptx.ShapeType.rect, { x, y, w: 5.9, h: 0.04, fill: { color: c.c } });
    s.addText(c.t, { x: x + 0.15, y: y + 0.12, w: 5.6, h: 0.28, fontSize: 10, fontFace: BF, color: c.c, bold: true });
    s.addText(c.d, { x: x + 0.15, y: y + 0.44, w: 5.6, h: 1.3, fontSize: 8, fontFace: BF, color: C.text, lineSpacing: 12 });
  });
  // Local stats
  thinRule(s, ML, 5.4);
  d.stats.forEach((st, i) => {
    const x = ML + i * 3.0;
    s.addText(st.v, { x, y: 5.5, w: 2.7, h: 0.35, fontSize: 17, fontFace: TF, color: C.blue, align: 'center', bold: true });
    s.addText(st.l, { x, y: 5.85, w: 2.7, h: 0.3, fontSize: 7, fontFace: BF, color: C.gray, align: 'center' });
  });
  // Summary
  s.addShape(pptx.ShapeType.roundRect, { x: ML, y: 6.35, w: 12.03, h: 0.55, fill: { color: C.navy }, rectRadius: 0.04 });
  s.addText(d.summary, { x: ML + 0.2, y: 6.35, w: 11.63, h: 0.55, fontSize: 8, fontFace: BF, color: C.white, valign: 'middle', lineSpacing: 12 });
}

// ===== P8: Case Studies =====
function buildP8(s, d) {
  d.cases.forEach((c, i) => {
    const x = ML + i * 6.2;
    card(s, x, 1.15, 5.9, 5.0);
    // Header
    s.addShape(pptx.ShapeType.rect, { x, y: 1.15, w: 5.9, h: 0.35, fill: { color: C.navy } });
    s.addText(c.t, { x: x + 0.15, y: 1.15, w: 5.6, h: 0.35, fontSize: 9.5, fontFace: BF, color: C.white, bold: true, valign: 'middle' });
    // Context
    s.addText('情境', { x: x + 0.15, y: 1.6, w: 1.2, h: 0.2, fontSize: 7.5, fontFace: BF, color: C.blue, bold: true });
    s.addText(c.ct, { x: x + 0.15, y: 1.8, w: 5.6, h: 0.9, fontSize: 7.5, fontFace: BF, color: C.text, lineSpacing: 11 });
    // Data
    thinRule(s, x + 0.1, 2.85, 5.7);
    s.addText('关联数据', { x: x + 0.15, y: 2.9, w: 1.5, h: 0.2, fontSize: 7.5, fontFace: BF, color: C.orange, bold: true });
    s.addShape(pptx.ShapeType.roundRect, { x: x + 0.15, y: 3.1, w: 5.6, h: 0.35, fill: { color: 'FFF8F0' }, line: { color: C.orange, width: 0.3 }, rectRadius: 0.04 });
    s.addText(c.d, { x: x + 0.3, y: 3.1, w: 5.3, h: 0.35, fontSize: 7, fontFace: BF, color: C.text, valign: 'middle' });
    // Analysis
    thinRule(s, x + 0.1, 3.6, 5.7);
    s.addText('分析', { x: x + 0.15, y: 3.65, w: 1, h: 0.2, fontSize: 7.5, fontFace: BF, color: C.navy, bold: true });
    s.addShape(pptx.ShapeType.roundRect, { x: x + 0.15, y: 3.85, w: 5.6, h: 0.65, fill: { color: C.bg }, line: { color: C.lighter, width: 0.3 }, rectRadius: 0.04 });
    s.addText(c.an, { x: x + 0.3, y: 3.85, w: 5.3, h: 0.65, fontSize: 7.5, fontFace: BF, color: C.navy, valign: 'middle', lineSpacing: 11 });
    s.addText(c.r, { x: x + 0.15, y: 4.6, w: 5.6, h: 0.2, fontSize: 6.5, fontFace: BF, color: C.lg });
  });
}

// ===== P9: Structural Tension =====
function buildP9(s, d) {
  // 4 quadrant cards
  d.quads.forEach(it => {
    card(s, it.x, it.y, it.w, it.h);
    s.addShape(pptx.ShapeType.rect, { x: it.x, y: it.y, w: 0.05, h: it.h, fill: { color: it.c } });
    s.addText(it.r, { x: it.x + 0.2, y: it.y + 0.1, w: 2, h: 0.3, fontSize: 12, fontFace: TF, color: it.c, bold: true });
    s.addText(it.t, { x: it.x + 0.2, y: it.y + 0.45, w: it.w - 0.4, h: 0.7, fontSize: 9.5, fontFace: BF, color: C.text, lineSpacing: 14 });
  });
  // Center tension circle
  s.addShape(pptx.ShapeType.ellipse, { x: 5.5, y: 2.5, w: 2.2, h: 0.85, fill: { color: C.orange } });
  s.addText('张力', { x: 5.5, y: 2.55, w: 2.2, h: 0.75, fontSize: 14, fontFace: TF, color: C.white, align: 'center', bold: true, valign: 'middle' });
  // Conclusion
  card(s, ML, 4.6, 12.03, 1.6);
  s.addShape(pptx.ShapeType.rect, { x: ML, y: 4.6, w: 0.05, h: 1.6, fill: { color: C.orange } });
  s.addText('结构性矛盾的本质', { x: ML + 0.2, y: 4.68, w: 5, h: 0.25, fontSize: 10, fontFace: BF, color: C.navy, bold: true });
  s.addText(d.conclusion, { x: ML + 0.2, y: 5.0, w: 11.63, h: 1.0, fontSize: 8.5, fontFace: BF, color: C.text, lineSpacing: 14 });
}

// ===== P10: Problem 1 =====
function buildP10(s, d) {
  // Left: factors
  card(s, ML, 1.15, 6.5, 2.4);
  s.addShape(pptx.ShapeType.rect, { x: ML, y: 1.15, w: 6.5, h: 0.3, fill: { color: C.blue } });
  s.addText('学业压力的制度性来源', { x: ML, y: 1.15, w: 6.5, h: 0.3, fontSize: 9, fontFace: BF, color: C.white, align: 'center', bold: true, valign: 'middle' });
  d.factors.forEach((f, i) => {
    const y = 1.6 + i * 0.63;
    s.addShape(pptx.ShapeType.rect, { x: ML + 0.15, y: y, w: 0.04, h: 0.5, fill: { color: f.c } });
    s.addText(f.l, { x: ML + 0.35, y: y - 0.02, w: 2, h: 0.25, fontSize: 9, fontFace: BF, color: f.c, bold: true });
    s.addText(f.d, { x: ML + 0.35, y: y + 0.23, w: 5.8, h: 0.28, fontSize: 7.5, fontFace: BF, color: C.text });
  });
  // Left bottom: pressure chain
  card(s, ML, 3.75, 6.5, 2.4);
  s.addText('压力传导闭环', { x: ML + 0.2, y: 3.82, w: 3, h: 0.25, fontSize: 9, fontFace: BF, color: C.navy, bold: true });
  d.chain.forEach((c, i) => {
    const cx = ML + 0.15 + i * 1.0;
    s.addShape(pptx.ShapeType.roundRect, { x: cx, y: 4.15, w: 0.85, h: 0.4, fill: { color: i < 3 ? C.blue : C.orange }, rectRadius: 0.04 });
    s.addText(c, { x: cx, y: 4.15, w: 0.85, h: 0.4, fontSize: 6.5, fontFace: BF, color: C.white, align: 'center', valign: 'middle' });
    if (i < 5) s.addText('→', { x: cx + 0.82, y: 4.18, w: 0.25, h: 0.35, fontSize: 9, color: C.lg, align: 'center', valign: 'middle' });
  });
  s.addText('需通过结构性干预打破这一闭环', { x: ML + 0.2, y: 4.65, w: 6, h: 0.25, fontSize: 7.5, fontFace: BF, color: C.orange });
  // Right: supporting data
  card(s, 7.4, 1.15, 5.3, 5.0);
  s.addShape(pptx.ShapeType.rect, { x: 7.4, y: 1.15, w: 5.3, h: 0.3, fill: { color: C.orange } });
  s.addText('关键支撑数据', { x: 7.4, y: 1.15, w: 5.3, h: 0.3, fontSize: 9, fontFace: BF, color: C.white, align: 'center', bold: true, valign: 'middle' });
  d.data.forEach((item, i) => {
    const y = 1.6 + i * 0.85;
    s.addShape(pptx.ShapeType.roundRect, { x: 7.55, y, w: 5.0, h: 0.65, fill: { color: C.white }, line: { color: C.lighter, width: 0.3 }, rectRadius: 0.04 });
    s.addText(item.t, { x: 7.7, y: y + 0.03, w: 4.7, h: 0.32, fontSize: 8, fontFace: BF, color: C.text, valign: 'middle' });
    s.addText(item.s, { x: 7.7, y: y + 0.35, w: 4.7, h: 0.22, fontSize: 6.5, fontFace: BF, color: C.lg });
  });
}

// ===== P11: Problem 2 =====
function buildP11(s, d) {
  // Left: reality
  card(s, ML, 1.15, 5.9, 5.0);
  s.addShape(pptx.ShapeType.rect, { x: ML, y: 1.15, w: 5.9, h: 0.35, fill: { color: C.orange } });
  s.addText('现实困境', { x: ML, y: 1.15, w: 5.9, h: 0.35, fontSize: 10, fontFace: BF, color: C.white, align: 'center', bold: true, valign: 'middle' });
  d.left.forEach((it, i) => {
    const y = 1.65 + i * 0.82;
    s.addShape(pptx.ShapeType.rect, { x: ML + 0.15, y, w: 0.04, h: 0.6, fill: { color: C.orange } });
    s.addText(it.q, { x: ML + 0.35, y, w: 5.3, h: 0.25, fontSize: 8.5, fontFace: BF, color: C.navy, bold: true });
    s.addText(it.a, { x: ML + 0.35, y: y + 0.28, w: 5.3, h: 0.32, fontSize: 7, fontFace: BF, color: C.gray });
  });
  // Right: standard
  card(s, 6.8, 1.15, 5.9, 5.0);
  s.addShape(pptx.ShapeType.rect, { x: 6.8, y: 1.15, w: 5.9, h: 0.35, fill: { color: C.blue } });
  s.addText('政策要求与理想标准', { x: 6.8, y: 1.15, w: 5.9, h: 0.35, fontSize: 10, fontFace: BF, color: C.white, align: 'center', bold: true, valign: 'middle' });
  d.right.forEach((it, i) => {
    const y = 1.65 + i * 0.82;
    s.addShape(pptx.ShapeType.rect, { x: 6.95, y, w: 0.04, h: 0.6, fill: { color: C.blue } });
    s.addText(it.q, { x: 7.15, y, w: 5.3, h: 0.25, fontSize: 8.5, fontFace: BF, color: C.navy, bold: true });
    s.addText(it.a, { x: 7.15, y: y + 0.28, w: 5.3, h: 0.32, fontSize: 7, fontFace: BF, color: C.gray });
  });
  // Vertical divider label
  s.addText('← 现实 vs 理想 →', { x: 5.3, y: 3.2, w: 2.7, h: 0.3, fontSize: 10, fontFace: BF, color: C.gray, align: 'center', bold: true });
}

// ===== P12: International Comparison =====
function buildP12(s, d) {
  d.countries.forEach((c, i) => {
    const x = ML + i * 3.05;
    card(s, x, 1.15, 2.85, 5.0);
    // Country header
    s.addShape(pptx.ShapeType.rect, { x, y: 1.15, w: 2.85, h: 0.38, fill: { color: i === 3 ? C.orange : C.blue } });
    s.addText(c.n, { x, y: 1.15, w: 2.85, h: 0.38, fontSize: 10, fontFace: BF, color: C.white, align: 'center', bold: true, valign: 'middle' });
    // Mode
    s.addText('模式', { x: x + 0.12, y: 1.65, w: 2.6, h: 0.18, fontSize: 6.5, fontFace: BF, color: C.lg });
    s.addText(c.md, { x: x + 0.12, y: 1.8, w: 2.6, h: 0.22, fontSize: 8, fontFace: BF, color: C.navy, bold: true });
    // Features
    s.addText('特点', { x: x + 0.12, y: 2.1, w: 2.6, h: 0.18, fontSize: 6.5, fontFace: BF, color: C.lg });
    s.addText(c.fe, { x: x + 0.12, y: 2.28, w: 2.6, h: 1.2, fontSize: 6.5, fontFace: BF, color: C.text, lineSpacing: 10 });
    // Ratio
    thinRule(s, x + 0.1, 3.55, 2.65);
    s.addText('师生比', { x: x + 0.12, y: 3.6, w: 2.6, h: 0.18, fontSize: 6.5, fontFace: BF, color: C.lg });
    s.addText(c.r, { x: x + 0.12, y: 3.78, w: 2.6, h: 0.22, fontSize: 9, fontFace: BF, color: C.blue, bold: true });
    // Lessons
    thinRule(s, x + 0.1, 4.1, 2.65);
    s.addText('启示', { x: x + 0.12, y: 4.18, w: 2.6, h: 0.18, fontSize: 6.5, fontFace: BF, color: C.orange });
    s.addShape(pptx.ShapeType.roundRect, { x: x + 0.12, y: 4.38, w: 2.6, h: 0.85, fill: { color: C.bg }, line: { color: C.lighter, width: 0.3 }, rectRadius: 0.04 });
    s.addText(c.l, { x: x + 0.22, y: 4.4, w: 2.4, h: 0.8, fontSize: 6.5, fontFace: BF, color: C.navy, valign: 'middle', lineSpacing: 10 });
  });
}

// ===== P13: Root Causes =====
function buildP13(s, d) {
  d.layers.forEach((ly, li) => {
    const x = ML + li * 3.05;
    // Layer header
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.15, w: 2.85, h: 0.35, fill: { color: ly.c }, rectRadius: 0.04 });
    s.addText(ly.t, { x, y: 1.15, w: 2.85, h: 0.35, fontSize: 10, fontFace: BF, color: C.white, align: 'center', bold: true, valign: 'middle' });
    // Items
    ly.its.forEach((it, ii) => {
      const y = 1.65 + ii * 0.6;
      card(s, x + 0.05, y, 2.75, 0.48);
      s.addText((ii + 1) + '. ' + it, { x: x + 0.15, y: y + 0.02, w: 2.55, h: 0.44, fontSize: 6.5, fontFace: BF, color: C.text, valign: 'middle', lineSpacing: 10 });
    });
  });
  // Conclusions
  card(s, ML, 3.6, 12.03, 1.1);
  s.addShape(pptx.ShapeType.rect, { x: ML, y: 3.6, w: 0.05, h: 1.1, fill: { color: C.blue } });
  s.addText(d.conclusion1, { x: ML + 0.2, y: 3.65, w: 11.63, h: 0.35, fontSize: 9, fontFace: BF, color: C.blue, align: 'center', bold: true });
  s.addText(d.conclusion2, { x: ML + 0.2, y: 4.0, w: 11.63, h: 0.55, fontSize: 8, fontFace: BF, color: C.gray, align: 'center', lineSpacing: 13 });
  // Directional arrows
  const arrowY = 4.9;
  s.addText('↑ 宏观', { x: 1.0, y: arrowY, w: 2, h: 0.3, fontSize: 8, color: C.gray, align: 'center' });
  s.addText('↑ 中观', { x: 4.5, y: arrowY, w: 2, h: 0.3, fontSize: 8, color: C.gray, align: 'center' });
  s.addText('↑ 微观', { x: 8.5, y: arrowY, w: 2, h: 0.3, fontSize: 8, color: C.gray, align: 'center' });
}

// ===== P14: Proposal 1 =====
function buildP14(s, d) {
  d.items.forEach((it, i) => {
    const y = 1.15 + i * 1.55;
    // Number badge
    s.addShape(pptx.ShapeType.roundRect, { x: ML, y: y + 0.02, w: 0.4, h: 0.32, fill: { color: C.blue }, rectRadius: 0.04 });
    s.addText(it.n, { x: ML, y: y + 0.02, w: 0.4, h: 0.32, fontSize: 13, fontFace: BF, color: C.white, align: 'center', bold: true, valign: 'middle' });
    // Title
    s.addText(it.t, { x: ML + 0.55, y: y + 0.02, w: 11.3, h: 0.32, fontSize: 11, fontFace: BF, color: C.navy, bold: true, valign: 'middle' });
    // Content
    card(s, ML, y + 0.4, 12.03, 1.0);
    s.addText(it.b, { x: ML + 0.2, y: y + 0.45, w: 11.63, h: 0.9, fontSize: 7.5, fontFace: BF, color: C.text, lineSpacing: 13 });
  });
}

// ===== P15: Proposal 2 =====
function buildP15(s, d) {
  // Left column
  card(s, ML, 1.15, 5.9, 4.0);
  s.addShape(pptx.ShapeType.rect, { x: ML, y: 1.15, w: 5.9, h: 0.35, fill: { color: C.blue } });
  s.addText('心理健康课程体系建设', { x: ML, y: 1.15, w: 5.9, h: 0.35, fontSize: 10, fontFace: BF, color: C.white, align: 'center', bold: true, valign: 'middle' });
  d.left.forEach((it, i) => {
    const y = 1.65 + i * 0.7;
    s.addShape(pptx.ShapeType.ellipse, { x: ML + 0.2, y: y + 0.08, w: 0.08, h: 0.08, fill: { color: C.blue } });
    s.addText(it, { x: ML + 0.4, y, w: 5.3, h: 0.55, fontSize: 7.5, fontFace: BF, color: C.text, valign: 'middle', lineSpacing: 11 });
  });
  // Right column
  card(s, 6.8, 1.15, 5.9, 4.0);
  s.addShape(pptx.ShapeType.rect, { x: 6.8, y: 1.15, w: 5.9, h: 0.35, fill: { color: C.blue } });
  s.addText('班级心理委员与朋辈支持', { x: 6.8, y: 1.15, w: 5.9, h: 0.35, fontSize: 10, fontFace: BF, color: C.white, align: 'center', bold: true, valign: 'middle' });
  d.right.forEach((it, i) => {
    const y = 1.65 + i * 0.7;
    s.addShape(pptx.ShapeType.ellipse, { x: 7.0, y: y + 0.08, w: 0.08, h: 0.08, fill: { color: C.blue } });
    s.addText(it, { x: 7.2, y, w: 5.3, h: 0.55, fontSize: 7.5, fontFace: BF, color: C.text, valign: 'middle', lineSpacing: 11 });
  });
  // Footer
  s.addShape(pptx.ShapeType.roundRect, { x: ML, y: 5.4, w: 12.03, h: 0.5, fill: { color: C.navy }, rectRadius: 0.04 });
  s.addText(d.footer, { x: ML + 0.2, y: 5.4, w: 11.63, h: 0.5, fontSize: 8.5, fontFace: BF, color: C.white, align: 'center', valign: 'middle' });
}

// ===== P16: Proposal 3 =====
function buildP16(s, d) {
  d.cols.forEach((se, i) => {
    const x = ML + i * 4.1;
    card(s, x, 1.15, 3.85, 4.2);
    s.addShape(pptx.ShapeType.rect, { x, y: 1.15, w: 3.85, h: 0.38, fill: { color: se.c } });
    s.addText(se.t, { x, y: 1.15, w: 3.85, h: 0.38, fontSize: 10, fontFace: BF, color: C.white, align: 'center', bold: true, valign: 'middle' });
    se.its.forEach((it, j) => {
      const y = 1.7 + j * 0.75;
      s.addShape(pptx.ShapeType.roundRect, { x: x + 0.15, y, w: 3.55, h: 0.58, fill: { color: C.white }, line: { color: se.c, width: 0.3 }, rectRadius: 0.04 });
      s.addText((j + 1) + '. ' + it, { x: x + 0.25, y: y + 0.03, w: 3.35, h: 0.52, fontSize: 7.5, fontFace: BF, color: C.text, valign: 'middle', lineSpacing: 11 });
    });
  });
  // Footer bar
  s.addShape(pptx.ShapeType.roundRect, { x: ML, y: 5.55, w: 12.03, h: 0.55, fill: { color: C.navy }, rectRadius: 0.04 });
  s.addText(d.footer, { x: ML + 0.2, y: 5.55, w: 11.63, h: 0.55, fontSize: 8.5, fontFace: BF, color: C.white, align: 'center', valign: 'middle', lineSpacing: 13 });
}

// ===== P17: Proposal 4 =====
function buildP17(s, d) {
  // Indicator targets
  card(s, ML, 1.15, 12.03, 1.5);
  s.addShape(pptx.ShapeType.rect, { x: ML, y: 1.15, w: 0.05, h: 1.5, fill: { color: C.blue } });
  s.addText('评估指标体系', { x: ML + 0.2, y: 1.2, w: 3, h: 0.28, fontSize: 10, fontFace: BF, color: C.navy, bold: true });
  d.index.forEach((it, i) => {
    const x = ML + 0.2 + i * 2.45;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.55, w: 2.25, h: 0.9, fill: { color: C.bg }, line: { color: C.lighter, width: 0.3 }, rectRadius: 0.04 });
    s.addText(it.l, { x: x + 0.1, y: 1.58, w: 2.05, h: 0.22, fontSize: 7, fontFace: BF, color: C.navy, bold: true, align: 'center' });
    s.addText(it.t, { x: x + 0.1, y: 1.82, w: 2.05, h: 0.5, fontSize: 6.5, fontFace: BF, color: C.gray, align: 'center', valign: 'middle', lineSpacing: 10 });
  });
  // Supervision
  card(s, ML, 2.9, 12.03, 2.2);
  s.addShape(pptx.ShapeType.rect, { x: ML, y: 2.9, w: 0.05, h: 2.2, fill: { color: C.orange } });
  s.addText('监督机制设计', { x: ML + 0.2, y: 2.95, w: 3, h: 0.28, fontSize: 10, fontFace: BF, color: C.navy, bold: true });
  d.supervision.forEach((it, i) => {
    const y = 3.35 + i * 0.5;
    s.addText(it.w, { x: ML + 0.2, y, w: 1.2, h: 0.3, fontSize: 8, fontFace: BF, color: C.orange, bold: true, valign: 'middle' });
    s.addText(it.v, { x: ML + 1.5, y, w: 10.33, h: 0.3, fontSize: 8, fontFace: BF, color: C.text, valign: 'middle' });
  });
  // Governance logic
  card(s, ML, 5.35, 12.03, 0.85);
  s.addShape(pptx.ShapeType.rect, { x: ML, y: 5.35, w: 0.05, h: 0.85, fill: { color: C.blue } });
  s.addText('治理逻辑', { x: ML + 0.2, y: 5.4, w: 3, h: 0.25, fontSize: 9, fontFace: BF, color: C.navy, bold: true });
  s.addText(d.conclusion, { x: ML + 0.2, y: 5.65, w: 11.63, h: 0.4, fontSize: 8, fontFace: BF, color: C.text, lineSpacing: 13 });
}

// ===== P18: Pilot Model =====
function buildP18(s, d) {
  // Top center node
  s.addShape(pptx.ShapeType.roundRect, { x: 4.2, y: 1.15, w: 4.8, h: 0.45, fill: { color: C.navy }, rectRadius: 0.05 });
  s.addText('心理健康教育委员会（校级统筹）', { x: 4.2, y: 1.15, w: 4.8, h: 0.45, fontSize: 9, fontFace: BF, color: C.white, align: 'center', bold: true, valign: 'middle' });
  // Connector lines
  s.addShape(pptx.ShapeType.rect, { x: 6.6, y: 1.6, w: 0.02, h: 0.3, fill: { color: C.lighter } });
  s.addShape(pptx.ShapeType.rect, { x: 2.0, y: 1.9, w: 9.2, h: 0.02, fill: { color: C.lighter } });
  [3.5, 6.6, 9.7].forEach(x => {
    s.addShape(pptx.ShapeType.rect, { x, y: 1.9, w: 0.02, h: 0.2, fill: { color: C.lighter } });
  });
  // Three pillars
  d.pillars.forEach(p => {
    card(s, p.x, 2.2, 3.0, 2.8);
    s.addShape(pptx.ShapeType.rect, { x: p.x, y: 2.2, w: 3.0, h: 0.35, fill: { color: p.c } });
    s.addText(p.t, { x: p.x, y: 2.2, w: 3.0, h: 0.35, fontSize: 8.5, fontFace: BF, color: C.white, align: 'center', bold: true, valign: 'middle' });
    p.its.forEach((it, j) => {
      const y = 2.65 + j * 0.5;
      s.addShape(pptx.ShapeType.roundRect, { x: p.x + 0.12, y, w: 2.76, h: 0.38, fill: { color: C.bg }, line: { color: p.c, width: 0.3 }, rectRadius: 0.03 });
      s.addText(it, { x: p.x + 0.22, y, w: 2.56, h: 0.38, fontSize: 7.5, fontFace: BF, color: C.text, valign: 'middle' });
    });
  });
  // Bottom guarantee bar
  s.addShape(pptx.ShapeType.roundRect, { x: ML, y: 5.3, w: 12.03, h: 0.7, fill: { color: C.navy }, rectRadius: 0.04 });
  s.addText('运行保障机制', { x: ML + 0.2, y: 5.33, w: 3, h: 0.25, fontSize: 9, fontFace: BF, color: C.orange, bold: true });
  s.addText(d.guarantee, { x: ML + 0.2, y: 5.58, w: 11.63, h: 0.35, fontSize: 8, fontFace: BF, color: C.white, align: 'center', lineSpacing: 12 });
}

// ===== P19: Roadmap =====
function buildP19(s, d) {
  d.phases.forEach((ph, i) => {
    const y = 1.2 + i * 1.25;
    // Phase badge
    s.addShape(pptx.ShapeType.roundRect, { x: ML, y, w: 1.4, h: 0.3, fill: { color: ph.c }, rectRadius: 0.04 });
    s.addText(ph.p, { x: ML, y, w: 1.4, h: 0.3, fontSize: 8.5, fontFace: BF, color: C.white, align: 'center', bold: true, valign: 'middle' });
    // Time
    s.addText(ph.t, { x: ML + 1.55, y, w: 1.7, h: 0.3, fontSize: 8, fontFace: BF, color: C.gray, valign: 'middle' });
    // Content
    card(s, ML, y + 0.38, 12.03, 0.7);
    s.addText(ph.ts, { x: ML + 0.2, y: y + 0.42, w: 11.63, h: 0.62, fontSize: 8, fontFace: BF, color: C.text, valign: 'middle', lineSpacing: 12 });
  });
  // Footer
  s.addShape(pptx.ShapeType.roundRect, { x: ML, y: 6.35, w: 12.03, h: 0.5, fill: { color: C.bg }, line: { color: C.lighter, width: 0.3 }, rectRadius: 0.04 });
  s.addText(d.footer, { x: ML + 0.2, y: 6.35, w: 11.63, h: 0.5, fontSize: 8, fontFace: BF, color: C.navy, bold: true, valign: 'middle' });
}

// ===== P20: Outcomes =====
function buildP20(s, d) {
  d.cards.forEach((o, i) => {
    const x = ML + i * 4.1;
    card(s, x, 1.15, 3.85, 3.3);
    // Top color bar
    s.addShape(pptx.ShapeType.rect, { x, y: 1.15, w: 3.85, h: 0.04, fill: { color: o.c } });
    s.addText(o.t, { x, y: 1.3, w: 3.85, h: 0.3, fontSize: 10, fontFace: BF, color: C.navy, align: 'center', bold: true });
    o.its.forEach((it, j) => {
      const y = 1.7 + j * 0.55;
      s.addShape(pptx.ShapeType.roundRect, { x: x + 0.15, y, w: 3.55, h: 0.4, fill: { color: C.bg }, line: { color: C.lighter, width: 0.3 }, rectRadius: 0.04 });
      s.addText(it, { x: x + 0.25, y, w: 3.35, h: 0.4, fontSize: 7.5, fontFace: BF, color: C.text, valign: 'middle' });
    });
  });
  // Social value bar
  s.addShape(pptx.ShapeType.roundRect, { x: ML, y: 4.7, w: 12.03, h: 1.2, fill: { color: C.navy }, rectRadius: 0.05 });
  s.addText('社会价值', { x: ML + 0.2, y: 4.75, w: 5, h: 0.25, fontSize: 9.5, fontFace: BF, color: C.orange, bold: true });
  s.addText(d.social, { x: ML + 0.2, y: 5.05, w: 11.63, h: 0.7, fontSize: 8, fontFace: BF, color: C.white, lineSpacing: 13 });
}

// ===== P21: References =====
function buildP21(s, d) {
  const half = Math.ceil(d.refs.length / 2);
  d.refs.forEach((r, i) => {
    const colIdx = i < half ? 0 : 1;
    const x = colIdx === 0 ? ML : 6.7;
    const y = 1.15 + (i - colIdx * half) * 0.43;
    s.addText(r.n, { x, y, w: 0.45, h: 0.36, fontSize: 6.5, fontFace: BF, color: C.blue, bold: true, valign: 'middle' });
    s.addText(r.t, { x: x + 0.45, y, w: 5.6, h: 0.36, fontSize: 6.5, fontFace: BF, color: C.text, valign: 'middle', lineSpacing: 9 });
  });
  s.addText(d.footer, { x: ML, y: 6.2, w: 12.03, h: 0.3, fontSize: 6.5, fontFace: BF, color: C.lg });
}

// ===== P22: Summary =====
function buildP22(s, d) {
  s.background = { fill: C.navy };
  s.addShape(pptx.ShapeType.ellipse, { x: -2, y: -2, w: 5, h: 5, fill: { color: C.blue, transparency: 70 } });
  s.addShape(pptx.ShapeType.ellipse, { x: 10, y: 4, w: 6, h: 6, fill: { color: C.blue, transparency: 80 } });
  s.addShape(pptx.ShapeType.rect, { x: 4.5, y: 0.7, w: 4, h: 0.03, fill: { color: C.orange } });
  s.addText('总结与展望', { x: ML, y: 1.0, w: 12.03, h: 0.5, fontSize: 22, fontFace: TF, color: C.white, align: 'center', bold: true });
  d.lines.forEach((l, i) => {
    s.addText(l, { x: 1.0, y: 1.7 + i * 0.38, w: 11.33, h: 0.35, fontSize: l ? 9.5 : 4, fontFace: BF, color: l ? C.white : 'transparent', align: 'center', lineSpacing: 14 });
  });
  s.addShape(pptx.ShapeType.rect, { x: 3.5, y: 5.8, w: 6.33, h: 0.02, fill: { color: C.orange, transparency: 40 } });
  s.addText('“' + d.quote + '”', { x: 1.0, y: 6.0, w: 11.33, h: 0.4, fontSize: 13, fontFace: TF, color: C.orange, align: 'center' });
  // Custom footer
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 7.1, w: 13.33, h: 0.4, fill: { color: C.dark } });
  s.addText('哈尔滨市第三中学模拟政协提案 · 22/23', { x: ML, y: 7.1, w: 12.03, h: 0.4, fontSize: 7.5, color: C.gray, fontFace: BF, align: 'right', valign: 'middle' });
  s.addNotes(d.note);
}

// ===== P23: Thank You =====
function buildP23(s, d) {
  s.background = { fill: C.white };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 3.5, fill: { color: C.navy } });
  s.addShape(pptx.ShapeType.ellipse, { x: 10, y: -1, w: 4, h: 4, fill: { color: C.blue, transparency: 60 } });
  s.addShape(pptx.ShapeType.ellipse, { x: -1, y: 2, w: 3, h: 3, fill: { color: C.blue, transparency: 70 } });
  s.addShape(pptx.ShapeType.rect, { x: 4.5, y: 0.8, w: 4, h: 0.03, fill: { color: C.orange } });
  s.addText('感谢各位评委老师聆听', { x: ML, y: 1.1, w: 12.03, h: 0.65, fontSize: 28, fontFace: TF, color: C.white, align: 'center', bold: true });
  s.addText('Thank you for your attention', { x: ML, y: 1.85, w: 12.03, h: 0.3, fontSize: 12, fontFace: BF, color: C.gray, align: 'center' });
  s.addText('心晴调研小组 · 哈尔滨市第三中学', { x: ML, y: 4.0, w: 12.03, h: 0.35, fontSize: 12, fontFace: BF, color: C.navy, align: 'center' });
  s.addText('2026年5月', { x: ML, y: 4.45, w: 12.03, h: 0.3, fontSize: 10, fontFace: BF, color: C.gray, align: 'center' });
  s.addText(d.footer, { x: ML, y: 4.85, w: 12.03, h: 0.25, fontSize: 8, fontFace: BF, color: C.lg, align: 'center' });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 7.1, w: 13.33, h: 0.4, fill: { color: C.navy } });
  s.addText('哈尔滨市第三中学模拟政协提案 · 23/23', { x: ML, y: 7.1, w: 12.03, h: 0.4, fontSize: 7.5, color: C.gray, fontFace: BF, align: 'right', valign: 'middle' });
  s.addNotes(d.note);
}

// ===== GENERATE =====
SLIDE_IDS.forEach(id => buildSlide(id));

const out = "C:\\Users\\ROG\\OneDrive\\Desktop\\模拟政协\\模拟政协提案PPT.pptx";
pptx.writeFile({ fileName: out }).then(() => console.log('DONE')).catch(e => console.error(e));
