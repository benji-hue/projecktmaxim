// PowerPoint-версия презентации «Эффект относительного возраста» (pptxgenjs).
//   node build_deck.js [light|dark]   →  out/deck-<theme>.pptx + out/anim-<theme>.json
// Затем postprocess.py добавляет анимации по щелчку, переходы и финальные правки XML.
const path = require('path');
const fs = require('fs');
const pptxgen = require('pptxgenjs');

const SKILL = '/root/.claude/skills/synced/2d64efaf-fc81-4729-a14f-24728a20f0e8_08b58b4f-6847-4399-ba9d-bded480b3897/pptx';
const { applyTheme } = require(SKILL + '/scripts/apply_theme.js');

const THEMEKEY = process.argv[2] || 'light';
const A = (f) => path.join(__dirname, 'assets', THEMEKEY, f);
const OUTDIR = path.join(__dirname, 'out');
fs.mkdirSync(OUTDIR, { recursive: true });

// ---------- палитра (синий = «ранние», оранжевый = «поздние»; проверено на дальтонизм) ----------
const PAL = {
  light: {
    INK: '0B1B2E', INK2: '3B4F66', INK3: '5F738A', BLUE: '2A78D6', BLUE_INK: '1F5FC4', BLUE_MUTE: 'A9C6EC',
    BLUE_SOFT: 'D6E6FA', ORANGE: 'E2622F', ORANGE_INK: 'B8461A', ORANGE_SOFT: 'FBE0D3', CARD: 'FFFFFF',
    ICE: 'EEF3F8', GRID: 'D3DCE6', SHADOW: '0B1B2E', SHADOW_OP: 0.16, CARD_T: 26, CHIP_T: 38,
  },
  dark: {
    INK: 'EAF2FB', INK2: 'A9BDD3', INK3: '7F93AB', BLUE: '3987E5', BLUE_INK: '8DBCF4', BLUE_MUTE: '2C5482',
    BLUE_SOFT: '1D3557', ORANGE: 'D95926', ORANGE_INK: 'F39A6F', ORANGE_SOFT: '4A2618', CARD: '1A2B47',
    ICE: '0F1A2C', GRID: '2A3A52', SHADOW: '000000', SHADOW_OP: 0.45, CARD_T: 30, CHIP_T: 55,
  },
}[THEMEKEY];

const THEME = {
  name: THEMEKEY === 'dark' ? 'Арена' : 'Лёд',
  headFontFace: 'Arial',
  bodyFontFace: 'Arial',
  colors: {
    dk1: PAL.INK, lt1: PAL.CARD, dk2: PAL.INK2, lt2: PAL.ICE,
    accent1: PAL.BLUE, accent2: PAL.ORANGE, accent3: PAL.BLUE_MUTE, accent4: PAL.INK3,
    accent5: PAL.BLUE_SOFT, accent6: PAL.ORANGE_SOFT, hlink: PAL.BLUE_INK, folHlink: PAL.INK3,
  },
};

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13,333 × 7,5 дюйма = холст 1920 × 1080 px
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = 'Эффект относительного возраста';
pres.subject = 'Почему хоккеисты рождаются в январе. Марафон лекториев «Поток»';
pres.author = 'Имя Фамилия';
const C = pres.SchemeColor;

// px холста 1920×1080 → дюймы; px → пункты (1 px = 0,5 pt)
const X = (px) => Math.round((px / 144) * 10000) / 10000;
const PT = (px) => Math.round(px) / 2;
const MONO = 'Consolas';
const BLACK = 'Arial Black';
const LANG = 'ru-RU';
const WHITE = 'FFFFFF';
const MONTHS = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];

// ---------- реестр анимаций: slideNo → [[{name, fx, delay}], ...] по шагам ----------
const ANIM = {};
const TRANS = {};
let slideNo = 0;
const anim = (step, name, fx = 'fade', delay = 0) => {
  const s = (ANIM[slideNo] = ANIM[slideNo] || []);
  (s[step - 1] = s[step - 1] || []).push({ name, fx, delay });
};

// ---------- примитивы ----------
const names = new Set();
function nm(n) {
  const key = slideNo + '|' + n;
  if (names.has(key)) throw new Error('повтор имени объекта: ' + key);
  names.add(key);
  return n;
}
function shadow() {
  return { type: 'outer', color: PAL.SHADOW, opacity: PAL.SHADOW_OP, blur: 14, offset: 5, angle: 90 };
}
function glass(s, name, x, y, w, h, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: X(x), y: X(y), w: X(w), h: X(h), rectRadius: X(o.r ?? 30),
    fill: { color: o.fill ?? C.background1, transparency: o.t ?? PAL.CARD_T },
    line: o.line ?? { color: C.background1, transparency: THEMEKEY === 'dark' ? 70 : 15, width: 1.25 },
    shadow: o.noShadow ? undefined : shadow(),
    objectName: nm(name),
  });
}
function rect(s, name, x, y, w, h, o = {}) {
  s.addShape(o.round === false ? pres.shapes.RECTANGLE : pres.shapes.ROUNDED_RECTANGLE, {
    x: X(x), y: X(y), w: X(w), h: X(h), rectRadius: o.round === false ? undefined : X(o.r ?? 12),
    fill: o.fill ?? { color: C.accent1 }, line: o.line ?? { type: 'none' }, objectName: nm(name),
  });
}
function line(s, name, x1, y1, x2, y2, o = {}) {
  s.addShape(pres.shapes.LINE, {
    x: X(Math.min(x1, x2)), y: X(Math.min(y1, y2)), w: X(Math.abs(x2 - x1)) || 0.0001, h: X(Math.abs(y2 - y1)) || 0.0001,
    flipH: x2 < x1, flipV: y2 < y1,
    line: { color: o.color ?? C.accent4, width: o.width ?? 1.5, dashType: o.dash ?? 'solid', endArrowType: o.arrow ? 'triangle' : undefined, transparency: o.t },
    objectName: nm(name),
  });
}
function text(s, name, runs, x, y, w, h, o = {}) {
  const base = {
    x: X(x), y: X(y), w: X(w), h: X(h), margin: 0, isTextBox: true, valign: o.valign ?? 'top', align: o.align ?? 'left',
    fontFace: o.font ?? 'Arial', fontSize: o.size ?? 14, color: o.color ?? C.text1, bold: !!o.bold, lang: LANG,
    lineSpacingMultiple: o.ls ?? 1.0, charSpacing: o.cs, italic: o.italic, objectName: nm(name), fit: 'none', wrap: o.wrap ?? true,
    paraSpaceAfter: o.psa,
  };
  s.addText(runs, base);
}
function img(s, name, file, x, y, w, h, o = {}) {
  s.addImage({ path: A(file), x: X(x), y: X(y), w: X(w), h: X(h), objectName: nm(name), altText: o.alt ?? '' });
}
const R = (t, o = {}) => ({ text: t, options: { lang: LANG, ...o } });
function kicker(s, num, label) {
  s.addText([R(num, { color: C.accent1, bold: true }), R('  —  ' + label.toUpperCase())], { placeholder: 'kicker' });
}

// ---------- макеты (layouts) ----------
pres.defineSlideMaster({
  title: 'Титул',
  background: { color: PAL.ICE },
  objects: [
    { placeholder: { options: { name: 'kicker', type: 'body', x: X(126), y: X(80), w: X(1100), h: X(36), fontFace: MONO, fontSize: 10.5, color: C.accent4, charSpacing: 3, margin: 0, valign: 'middle', align: 'left', lang: LANG }, text: '' } },
    { placeholder: { options: { name: 'title', type: 'title', x: X(92), y: X(146), w: X(1740), h: X(176), fontFace: BLACK, fontSize: 40, color: C.text1, margin: 0, valign: 'top', align: 'left', lineSpacingMultiple: 0.95, lang: LANG }, text: '' } },
    { placeholder: { options: { name: 'subtitle', type: 'body', x: X(96), y: X(344), w: X(1300), h: X(60), fontFace: 'Arial', fontSize: 23, color: C.text2, margin: 0, valign: 'top', align: 'left', lang: LANG }, text: '' } },
    { placeholder: { options: { name: 'author', type: 'body', x: X(96), y: X(444), w: X(1100), h: X(46), fontFace: 'Arial', fontSize: 17, bold: true, color: C.text1, margin: 0, valign: 'top', align: 'left', lang: LANG }, text: 'Имя Фамилия' } },
    { placeholder: { options: { name: 'direction', type: 'body', x: X(96), y: X(494), w: X(1100), h: X(40), fontFace: 'Arial', fontSize: 13, color: C.text2, margin: 0, valign: 'top', align: 'left', lang: LANG }, text: 'Образовательное направление' } },
  ],
});
pres.defineSlideMaster({
  title: 'Контент',
  background: { color: PAL.ICE },
  objects: [
    { placeholder: { options: { name: 'kicker', type: 'body', x: X(96), y: X(52), w: X(1300), h: X(36), fontFace: MONO, fontSize: 10.5, color: C.accent4, charSpacing: 3, margin: 0, valign: 'middle', align: 'left', lang: LANG }, text: '' } },
    { placeholder: { options: { name: 'title', type: 'title', x: X(96), y: X(98), w: X(1640), h: X(150), fontFace: 'Arial', fontSize: 30, bold: true, color: C.text1, margin: 0, valign: 'top', align: 'left', lineSpacingMultiple: 0.98, lang: LANG }, text: '' } },
    { placeholder: { options: { name: 'source', type: 'body', x: X(96), y: X(1020), w: X(1560), h: X(34), fontFace: MONO, fontSize: 9.5, color: C.accent4, margin: 0, valign: 'middle', align: 'left', lang: LANG }, text: '' } },
  ],
  slideNumber: { x: X(1700), y: X(1020), w: X(124), h: X(34), fontFace: MONO, fontSize: 10.5, color: C.accent4, align: 'right' },
});

function newSlide(layout, section, bgIndex) {
  slideNo++;
  if (section) pres.addSection({ title: section });
  const s = pres.addSlide({ masterName: layout, sectionTitle: CUR_SECTION = section || CUR_SECTION });
  s.background = { path: A(`bg-${String(bgIndex).padStart(2, '0')}.jpg`) };
  TRANS[slideNo] = 'fade';
  return s;
}
let CUR_SECTION = null;

// ---------- общие компоненты ----------
function monthChips(s, prefix, x0, y, w, h, o = {}) {
  const gap = o.gap ?? 10;
  const n = 12;
  const cw = (w - gap * (n - 1)) / n;
  const start = o.start ?? 1;
  const out = [];
  for (let k = 0; k < n; k++) {
    const m = ((start - 1 + k) % 12) + 1;
    const x = x0 + k * (cw + gap);
    const st = o.style ? o.style(m, k) : {};
    const fill = st.fill ?? { color: C.background1, transparency: PAL.CHIP_T };
    s.addText(o.labels === false ? '' : MONTHS[m - 1], {
      shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: X(o.r ?? 16),
      x: X(x), y: X(y), w: X(cw), h: X(h), margin: 0, align: 'center', valign: st.valign ?? 'middle',
      fontFace: 'Arial', fontSize: o.size ?? 12.5, bold: true, color: st.color ?? C.text2, lang: LANG,
      fill, line: st.line ?? { color: C.text1, transparency: THEMEKEY === 'dark' ? 80 : 86, width: 0.75 },
      objectName: nm(`${prefix} ${MONTHS[m - 1]}`),
    });
    out.push({ m, x, cw });
  }
  return out;
}

// ============================================================================
// 1 · ТИТУЛЬНЫЙ
// ============================================================================
{
  const s = newSlide('Титул', 'Введение', 1);
  s.addShape(pres.shapes.OVAL, { x: X(96), y: X(91), w: X(14), h: X(14), fill: { color: C.accent1 }, line: { type: 'none' }, objectName: nm('Точка') });
  s.addText([R('Марафон лекториев «Поток»'.toUpperCase())], { placeholder: 'kicker' });
  s.addText([R('Эффект '), R('относительного', { color: C.accent1, breakLine: true }), R('возраста')], { placeholder: 'title' });
  s.addText([R('Почему хоккеисты рождаются в январе')], { placeholder: 'subtitle' });
  s.addText([R('Имя Фамилия')], { placeholder: 'author' });
  s.addText([R('Образовательное направление')], { placeholder: 'direction' });

  const Y = 872, H = 76, X0 = 96, W = 1728;
  const chips = monthChips(s, 'Месяц', X0, Y, W, H);
  // шаг 1: январь–март + руки; шаг 2: октябрь–декабрь + руки
  for (const [step, months, col, hand] of [[1, [1, 2, 3], C.accent1, 'hand-blue.png'], [2, [10, 11, 12], C.accent2, 'hand-orange.png']]) {
    months.forEach((m, j) => {
      const c = chips.find((q) => q.m === m);
      const nmChip = `Шаг ${step} · ${MONTHS[m - 1]} цвет`;
      s.addText(MONTHS[m - 1], {
        shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: X(16), x: X(c.x), y: X(Y), w: X(c.cw), h: X(H), margin: 0,
        align: 'center', valign: 'middle', fontFace: 'Arial', fontSize: 12.5, bold: true, color: WHITE, lang: LANG,
        fill: { color: col }, line: { type: 'none' }, objectName: nm(nmChip),
      });
      anim(step, nmChip, 'fade', j * 80);
      const hn = `Шаг ${step} · Рука ${j + 1}`;
      img(s, hn, hand, c.x + c.cw / 2 - 30, Y - 78, 60, 60, { alt: 'Поднятая рука' });
      anim(step, hn, 'float', 120 + j * 110);
    });
  }
  text(s, 'Шаг 3 · Подводка', [R('Будь мы канадскими хоккеистами: у первой группы шансов попасть в НХЛ больше чем в 2 раза'.toUpperCase())],
    1030, 592, 794, 70, { font: MONO, size: 10, color: C.accent4, align: 'right', cs: 1, ls: 1.15 });
  text(s, 'Шаг 3 · Вывод', [R('Дело не в генах, '), R('а в календаре', { color: C.accent1 })], 830, 672, 994, 70, { size: 27, bold: true, align: 'right' });
  anim(3, 'Шаг 3 · Подводка', 'fade');
  anim(3, 'Шаг 3 · Вывод', 'float', 150);

  s.addNotes([
    'СЛАЙД 1 · титульный (≈ 0:33)',
    '',
    'Добрый день! Давайте начнём с эксперимента.',
    '[КЛИК 1 — подсвечиваются январь–март] Поднимите руку те, кто родился в январе, феврале или марте. (пауза, оглядеть зал)',
    '[КЛИК 2 — октябрь–декабрь] А теперь те, кто родился в октябре, ноябре или декабре. (пауза)',
    '[КЛИК 3] Будь мы канадскими хоккеистами, у первой группы шансов попасть в НХЛ было бы больше чем в два раза. И дело не в генах, а в календаре.',
  ].join('\n'));
}

// ============================================================================
// 2 · 1984
// ============================================================================
{
  const s = newSlide('Контент', null, 2);
  kicker(s, '02', 'Открытие');
  s.addText([R('1984: закономерность нашли в программке матча')], { placeholder: 'title' });
  s.addText([R('Barnsley, Thompson, Barnsley · CAHPER Journal, 1985 · программка — иллюстрация')], { placeholder: 'source' });
  const meta = JSON.parse(fs.readFileSync(A('meta.json'), 'utf8'));
  const p = meta.program;
  img(s, 'Программка матча', 'program.png', p.x, p.y, p.width, p.height, { alt: 'Иллюстрация: программка юниорского матча 1984 года, отмечены месяцы рождения игроков' });
  img(s, 'Шаг 1 · Обводка январь–март', 'program-circle.png', p.x, p.y, p.width, p.height, { alt: 'Обведены январь–март' });
  anim(1, 'Шаг 1 · Обводка январь–март', 'wheel', 0);

  text(s, 'История', [R('Психолог '), R('Роджер Барнсли', { bold: true, color: C.text1 }), R(' с женой '), R('Паулой', { bold: true, color: C.text1 }), R(' пришли на юниорский хоккейный матч в Летбридже, Канада.')],
    880, 296, 944, 110, { size: 17, color: C.text2, ls: 1.08 });

  glass(s, 'Шаг 1 · Карточка цитаты', 880, 428, 944, 218);
  text(s, 'Шаг 1 · Цитата', [R('«Смотри, они почти все родились в январе, феврале и марте»')], 924, 462, 856, 110, { size: 20, bold: true, ls: 1.05 });
  text(s, 'Шаг 1 · Автор цитаты', [R('— Паула Барнсли')], 924, 590, 600, 30, { font: MONO, size: 10.5, color: C.accent4, cs: 1 });
  text(s, 'Шаг 1 · Определение', [R('Так открыли '), R('эффект относительного возраста', { bold: true, color: C.text1 }), R(': в одной возрастной группе выигрывают те, кто старше.')],
    880, 672, 944, 80, { size: 14, color: C.text2, ls: 1.08 });
  anim(1, 'Шаг 1 · Карточка цитаты', 'float', 150);
  anim(1, 'Шаг 1 · Цитата', 'float', 150);
  anim(1, 'Шаг 1 · Автор цитаты', 'float', 150);
  anim(1, 'Шаг 1 · Определение', 'fade', 600);

  glass(s, 'Шаг 2 · Карточка цели', 880, 786, 340, 214);
  text(s, 'Шаг 2 · Цель заголовок', [R('ЦЕЛЬ')], 914, 816, 280, 26, { font: MONO, size: 10.5, color: C.accent1, bold: true, cs: 2 });
  text(s, 'Шаг 2 · Цель', [R('Проверить, жив ли эффект сегодня')], 914, 852, 280, 120, { size: 15, bold: true, ls: 1.05 });
  glass(s, 'Шаг 2 · Карточка гипотезы', 1246, 786, 578, 214);
  text(s, 'Шаг 2 · Гипотеза заголовок', [R('ГИПОТЕЗА')], 1280, 816, 500, 26, { font: MONO, size: 10.5, color: C.accent1, bold: true, cs: 2 });
  text(s, 'Шаг 2 · Гипотеза', [R('Среди игроков НХЛ родившихся в начале года значимо больше, чем родившихся в конце')], 1280, 852, 510, 130, { size: 15, bold: true, ls: 1.05 });
  ['Шаг 2 · Карточка цели', 'Шаг 2 · Цель заголовок', 'Шаг 2 · Цель'].forEach((n) => anim(2, n, 'float'));
  ['Шаг 2 · Карточка гипотезы', 'Шаг 2 · Гипотеза заголовок', 'Шаг 2 · Гипотеза'].forEach((n) => anim(2, n, 'float', 200));

  s.addNotes([
    'СЛАЙД 2 · 1984: открытие (≈ 0:34)',
    '',
    'В 1984 году канадский психолог Роджер Барнсли с женой Паулой пришёл на юниорский хоккейный матч.',
    '[КЛИК 1 — обводка и цитата] Паула листала программку и вдруг сказала: «Смотри, они почти все родились в январе, феврале и марте». Барнсли проверил профессионалов: та же картина. Так открыли эффект относительного возраста.',
    '[КЛИК 2 — цель и гипотеза] Моя цель — проверить, жив ли этот эффект сегодня. Гипотеза: среди игроков НХЛ родившихся в начале года значимо больше, чем родившихся в конце.',
    '',
    'Справка: матч был в Летбридже (Альберта, Канада); статья вышла в 1985 году. Программка на слайде — иллюстрация, не подлинный документ.',
  ].join('\n'));
}

// ============================================================================
// 3 · СНЕЖНЫЙ КОМ
// ============================================================================
{
  const s = newSlide('Контент', 'Механизм и масштаб', 3);
  kicker(s, '03', 'Механизм');
  s.addText([R('Почти год разницы в одной команде запускает снежный ком')], { placeholder: 'title' });
  s.addText([R('Rosenthal, Jacobson, 1968 · Merton, 1968 · Lemez et al., 2014')], { placeholder: 'source' });

  glass(s, 'Карточка команды', 96, 290, 720, 710);
  text(s, 'Подпись команды', [R('ОДНА КОМАНДА · ОДИН ГОД РОЖДЕНИЯ')], 134, 318, 640, 26, { font: MONO, size: 10, color: C.accent4, cs: 2 });
  img(s, 'Дети', 'kids.png', 134, 352, 644, 330, { alt: 'Два мальчика из одной команды: родившийся 1 января выше родившегося 31 декабря' });
  monthChips(s, 'Год', 134, 702, 644, 30, {
    gap: 6, r: 8, labels: false,
    style: (m) => (m === 1 ? { fill: { color: C.accent1 }, line: { type: 'none' } } : m === 12 ? { fill: { color: C.accent2 }, line: { type: 'none' } } : {}),
  });
  img(s, 'Скобка', 'brace.png', 133, 738, 646, 28, { alt: '' });
  text(s, '364 дня', [R('364 дня')], 134, 768, 644, 30, { font: MONO, size: 11.5, bold: true, align: 'center' });
  s.addChart(pres.charts.DOUGHNUT, [{ name: 'Доля жизни', labels: ['364 дня', 'остальная жизнь'], values: [11.1, 88.9] }], {
    x: X(124), y: X(826), w: X(150), h: X(150), holeSize: 66, chartColors: [PAL.ORANGE, PAL.GRID], showLegend: false,
    showValue: false, showLabel: false, showPercent: false, showTitle: false, firstSliceAng: 0, dataBorder: { pt: 0, color: PAL.GRID },
    plotArea: { fill: { color: 'FFFFFF', transparency: 100 } }, chartArea: { fill: { color: 'FFFFFF', transparency: 100 } },
    objectName: nm('Кольцо 11 процентов'), altText: '11 % — доля 364 дней в жизни девятилетнего',
  });
  text(s, '11 процентов', [R('11 %')], 296, 836, 400, 70, { font: BLACK, size: 30 });
  text(s, 'Пояснение 11 процентов', [R('жизни девятилетнего — больше десятой части')], 296, 912, 480, 60, { size: 13, color: C.text2, ls: 1.05 });

  // петля: Отбор → Ресурсы → Ожидания ↺, из «Отбора» — уход
  const node = (name, cx, cy, w, h, num, title, desc, step, exit) => {
    const x = cx - w / 2, y = cy - h / 2;
    glass(s, `Шаг ${step} · ${name} карточка`, x, y, w, h, exit ? { line: { color: C.accent2, transparency: 20, width: 2 } } : {});
    text(s, `Шаг ${step} · ${name} номер`, [R(num)], x + 26, y + 22, w - 52, 24, { font: MONO, size: 10, bold: true, color: exit ? C.accent2 : C.accent1, cs: 1.5 });
    text(s, `Шаг ${step} · ${name} заголовок`, [R(title)], x + 26, y + 52, w - 52, 40, { size: 18, bold: true });
    text(s, `Шаг ${step} · ${name} текст`, [R(desc)], x + 24, y + 96, w - 44, 76, { size: 11.5, color: C.text2, ls: 1.05 });
    [' карточка', ' номер', ' заголовок', ' текст'].forEach((p) => anim(step, `Шаг ${step} · ${name}${p}`, 'float'));
  };
  const NY = 545;
  node('Отбор', 1012, NY, 296, 184, '01', 'Отбор', '«Сильнее сейчас» ≠ талантливее', 1);
  node('Ресурсы', 1344, NY, 296, 184, '02', 'Ресурсы', 'Лучшие тренеры и больше льда', 2);
  node('Ожидания', 1676, NY, 296, 184, '03', 'Ожидания', 'Пигмалион: в кого верят, тот растёт', 3);
  node('Уход', 1012, 784, 296, 176, '↓ НЕ ОТОБРАЛИ', 'Уход', 'Младшие на скамейке бросают спорт', 4, true);
  line(s, 'Шаг 2 · Стрелка 1–2', 1164, NY, 1192, NY, { arrow: true, width: 2 });
  anim(2, 'Шаг 2 · Стрелка 1–2', 'fade', 200);
  line(s, 'Шаг 3 · Стрелка 2–3', 1496, NY, 1524, NY, { arrow: true, width: 2 });
  img(s, 'Шаг 3 · Дуга', 'loop-arc.png', 984, 283, 720, 176, { alt: 'Круг повторяется' });
  text(s, 'Шаг 3 · Подпись дуги', [R('по кругу — разрыв растёт')], 1144, 372, 400, 28, { font: MONO, size: 10, color: C.accent4, align: 'center' });
  anim(3, 'Шаг 3 · Стрелка 2–3', 'fade', 200);
  anim(3, 'Шаг 3 · Дуга', 'wipeLeft', 350);
  anim(3, 'Шаг 3 · Подпись дуги', 'fade', 700);
  line(s, 'Шаг 4 · Стрелка ухода', 1012, 639, 1012, 694, { arrow: true, width: 2, color: C.accent2 });
  anim(4, 'Шаг 4 · Стрелка ухода', 'fade');

  glass(s, 'Шаг 5 · Пророчество карточка', 1196, 696, 628, 176, { fill: C.accent5, t: THEMEKEY === 'dark' ? 10 : 25 });
  text(s, 'Шаг 5 · Пророчество заголовок', [R('САМОИСПОЛНЯЮЩЕЕСЯ ПРОРОЧЕСТВО')], 1232, 724, 560, 24, { font: MONO, size: 10, bold: true, color: C.accent1, cs: 1.5 });
  text(s, 'Шаг 5 · Пророчество текст', [R('Решили, что январские сильнее, — и сами сделали их сильнее')], 1232, 762, 560, 92, { size: 16, bold: true, ls: 1.05 });
  ['Шаг 5 · Пророчество карточка', 'Шаг 5 · Пророчество заголовок', 'Шаг 5 · Пророчество текст'].forEach((n) => anim(5, n, 'float'));

  s.addNotes([
    'СЛАЙД 3 · снежный ком (≈ 0:44)',
    '',
    'В детском хоккее команды собирают по году рождения. Мальчик, родившийся 1 января, и мальчик, родившийся 31 декабря, играют в одной команде, хотя между ними почти год. Для девятилетнего это больше десятой части жизни.',
    '[КЛИК 1 — Отбор] Дальше всё катится как снежный ком. Тренер берёт того, кто сильнее сейчас, и путает зрелость с талантом.',
    '[КЛИК 2 — Ресурсы] Отобранные получают лучших тренеров и больше льда.',
    '[КЛИК 3 — Ожидания] Работает эффект Пигмалиона: тот, в кого верят, растёт быстрее.',
    '[КЛИК 4 — Уход] А младшие сидят на скамейке и бросают спорт.',
    '[КЛИК 5] Получается самоисполняющееся пророчество: мы решили, что январские сильнее, и сами сделали их сильнее.',
  ].join('\n'));
}

// ============================================================================
// 4 · МАСШТАБ
// ============================================================================
{
  const s = newSlide('Контент', null, 4);
  kicker(s, '04', 'Масштаб');
  s.addText([R('Эффект находят почти везде — и через 40 лет он на месте')], { placeholder: 'title' });
  s.addText([R('Cobley, Baker, Wattie, McKenna · Sports Medicine, 2009 · Barnsley · Frontiers in Sports and Active Living, 2025')], { placeholder: 'source' });
  [['38', 'исследований'], ['14', 'видов спорта'], ['16', 'стран']].forEach(([n, l], k) => {
    const x = 96 + k * 588;
    glass(s, `Число ${n} карточка`, x, 290, 552, 250);
    text(s, `Число ${n}`, [R(n)], x + 40, 312, 480, 150, { font: BLACK, size: 66, valign: 'middle' });
    text(s, `Число ${n} подпись`, [R(l)], x + 40, 470, 480, 44, { size: 17, bold: true, color: C.text2 });
  });
  text(s, 'Подпись метаанализа', [R('Метаанализ Cobley et al., 2009 · 253 выборки за 1984–2007 годы')], 96, 558, 1200, 30, { font: MONO, size: 10.5, color: C.accent4 });

  glass(s, 'Шаг 1 · Барнсли карточка', 96, 622, 1000, 378);
  text(s, 'Шаг 1 · Барнсли заголовок', [R('Барнсли повторил своё исследование через 40 лет')], 136, 652, 920, 40, { size: 16, bold: true });
  text(s, 'Шаг 1 · Барнсли подзаголовок', [R('Лучшие юниоры по кварталу рождения — почти как в 1985 году')], 136, 696, 920, 32, { size: 12, color: C.text2 });
  const bars = [['I квартал', 'январь – март', 40, C.accent1, '≈ 40 %'], ['IV квартал', 'октябрь – декабрь', 10, C.accent2, '≈ 10 %']];
  bars.forEach(([t, sub, v, col, lbl], k) => {
    const y = 776 + k * 104;
    text(s, `Шаг 1 · ${t}`, [R(t)], 136, y, 230, 30, { size: 14, bold: true });
    text(s, `Шаг 1 · ${t} месяцы`, [R(sub)], 136, y + 32, 230, 24, { font: MONO, size: 9.5, color: C.accent4 });
    rect(s, `Шаг 1 · ${t} столбик`, 380, y + 6, (v / 40) * 470, 48, { fill: { color: col }, r: 10 });
    text(s, `Шаг 1 · ${t} значение`, [R(lbl)], 870, y + 4, 200, 54, { font: BLACK, size: 20, valign: 'middle' });
  });
  ['Шаг 1 · Барнсли карточка', 'Шаг 1 · Барнсли заголовок', 'Шаг 1 · Барнсли подзаголовок'].forEach((n) => anim(1, n, 'float'));
  bars.forEach(([t], k) => {
    anim(1, `Шаг 1 · ${t}`, 'fade', 300 + k * 200);
    anim(1, `Шаг 1 · ${t} месяцы`, 'fade', 300 + k * 200);
    anim(1, `Шаг 1 · ${t} столбик`, 'wipeRight', 400 + k * 200);
    anim(1, `Шаг 1 · ${t} значение`, 'fade', 900 + k * 200);
  });

  glass(s, 'Шаг 2 · Цитата карточка', 1136, 622, 688, 378);
  text(s, 'Шаг 2 · Цитата', [R('Life is not fair'), R('!', { color: C.accent2 })], 1176, 700, 620, 76, { font: BLACK, size: 30, valign: 'middle' });
  text(s, 'Шаг 2 · Перевод', [R('«Жизнь несправедлива»')], 1176, 790, 620, 44, { size: 17, bold: true, color: C.text2 });
  text(s, 'Шаг 2 · Подпись цитаты', [R('Название статьи Р. Барнсли · Frontiers, 2025')], 1176, 852, 620, 30, { font: MONO, size: 10, color: C.accent4 });
  ['Шаг 2 · Цитата карточка', 'Шаг 2 · Цитата', 'Шаг 2 · Перевод', 'Шаг 2 · Подпись цитаты'].forEach((n) => anim(2, n, 'float'));

  s.addNotes([
    'СЛАЙД 4 · масштаб (≈ 0:24)',
    '',
    'Метаанализ 38 исследований из 14 видов спорта и 16 стран нашёл эффект почти везде.',
    '[КЛИК 1] А в 2025 году Барнсли сам повторил своё исследование через сорок лет. Результат почти тот же: около 40 % лучших юниоров родились в первом квартале и только 10 % в последнем.',
    '[КЛИК 2] Статью он назвал «Жизнь несправедлива».',
  ].join('\n'));
}

// ============================================================================
// 5 · МОЁ ИССЛЕДОВАНИЕ
// ============================================================================
const QLAB = ['I кв.\nянв–мар', 'II кв.\nапр–июн', 'III кв.\nиюл–сен', 'IV кв.\nокт–дек'];
const EXPECT = [24.7, 24.9, 25.2, 25.2];
function chartOpts(extra) {
  return Object.assign({
    showLegend: false, showTitle: false, barGapWidthPct: 70,
    valAxisMinVal: 0, valAxisMaxVal: 40, valAxisMajorUnit: 10, valAxisLabelFormatCode: '0',
    valAxisLabelColor: PAL.INK3, valAxisLabelFontSize: 11, valAxisLabelFontFace: '+mn-lt', valAxisLineShow: false,
    valGridLine: { color: PAL.GRID, size: 0.75 }, catGridLine: { style: 'none' },
    catAxisLabelColor: PAL.INK, catAxisLabelFontSize: 13, catAxisLabelFontFace: '+mn-lt', catAxisLineShow: true,
    catAxisLineColor: PAL.INK3, catAxisLineSize: 1,
    plotArea: { fill: { color: 'FFFFFF', transparency: 100 } }, chartArea: { fill: { color: 'FFFFFF', transparency: 100 } },
  }, extra);
}
{
  const s = newSlide('Контент', 'Исследование', 5);
  kicker(s, '05', 'Моё исследование');
  s.addText([R('853 игрока НХЛ: через 40 лет эффект на месте')], { placeholder: 'title' });
  s.addText([R('Данные: составы НХЛ 2024/25 (api-web.nhle.com) · χ²(3) = 29,9 для канадцев · расчёт — analysis.py')], { placeholder: 'source' });

  glass(s, 'Метод карточка', 96, 268, 520, 732);
  const items = [['853', 'игрока', 'составы всех клубов НХЛ, сезон 2024/25'], ['4', 'квартала рождения', 'янв–мар · апр–июн · июл–сен · окт–дек'], ['χ²', 'критерий хи-квадрат', 'ожидание — по числу дней в квартале']];
  items.forEach(([v, t, sub], k) => {
    const y = 300 + k * 128;
    if (v === 'χ²') text(s, 'Метод χ²', [R('χ', { italic: true }), R('²')], 134, y - 6, 110, 60, { font: 'Cambria', size: 28, valign: 'middle' });
    else text(s, `Метод ${v}`, [R(v)], 134, y, 120, 50, { font: BLACK, size: 22 });
    text(s, `Метод ${t}`, [R(t)], 258, y + 2, 330, 30, { size: 14, bold: true });
    text(s, `Метод ${t} пояснение`, [R(sub)], 258, y + 36, 330, 56, { size: 11, color: C.text2, ls: 1.05 });
    if (k < 2) line(s, `Метод разделитель ${k + 1}`, 134, y + 108, 578, y + 108, { color: C.text1, width: 0.75, t: 85 });
  });
  rect(s, 'Шаг 2 · Вывод фон', 126, 832, 460, 140, { fill: { color: C.accent5, transparency: THEMEKEY === 'dark' ? 0 : 10 }, line: { color: C.accent1, transparency: 50, width: 1 }, r: 20 });
  img(s, 'Шаг 2 · Галочка', 'check.png', 150, 856, 34, 34, { alt: 'Подтверждено' });
  text(s, 'Шаг 2 · Вывод', [R('Гипотеза подтвердилась')], 196, 856, 380, 36, { size: 15, bold: true, valign: 'middle' });
  text(s, 'Шаг 2 · p-значение', [R('p < 0,001: шанс случайности меньше 1 из 1000')], 154, 904, 410, 54, { font: MONO, size: 10.5, color: C.text2, ls: 1.1 });
  ['Шаг 2 · Вывод фон', 'Шаг 2 · Галочка', 'Шаг 2 · Вывод', 'Шаг 2 · p-значение'].forEach((n) => anim(2, n, 'float'));

  glass(s, 'График карточка', 652, 268, 1172, 732);
  text(s, 'График заголовок', [R('Канадцы в НХЛ по кварталу рождения'), R('   n = 342 · %', { fontFace: MONO, fontSize: 10.5, bold: false, color: C.accent4 })], 690, 298, 820, 34, { size: 14, bold: true });
  text(s, 'График легенда', [R('- - - ожидание ≈ 25 %')], 1500, 298, 290, 34, { font: MONO, size: 10.5, color: C.accent4, align: 'right' });
  s.addChart([
    { type: pres.charts.BAR, data: [{ name: 'Канада, %', labels: QLAB, values: [34.5, 28.1, 22.5, 14.9] }],
      options: { barDir: 'col', chartColors: [PAL.BLUE, PAL.BLUE_MUTE, PAL.BLUE_MUTE, PAL.BLUE], showValue: true, dataLabelPosition: 'outEnd',
        dataLabelFormatCode: '0.0" %"', dataLabelFontSize: 14, dataLabelFontBold: true, dataLabelFontFace: '+mn-lt', dataLabelColor: PAL.INK } },
    { type: pres.charts.LINE, data: [{ name: 'Ожидание, %', labels: QLAB, values: EXPECT }],
      options: { chartColors: [PAL.INK3], lineSize: 2, lineDash: 'dash', lineDataSymbol: 'none', showValue: false } },
  ], chartOpts({ x: X(676), y: X(344), w: X(800), h: X(636), objectName: nm('График Канада'), altText: 'Канадцы в НХЛ 2024/25 по кварталу рождения: 34,5 %, 28,1 %, 22,5 %, 14,9 %; ожидание около 25 %' }));
  text(s, 'Шаг 1 · 2,3', [R('2,3'), R('×', { color: C.accent1 })], 1494, 404, 320, 120, { font: BLACK, size: 50, valign: 'middle' });
  text(s, 'Шаг 1 · Пояснение 2,3', [R('Рождённых в январе – марте больше, чем в октябре – декабре: 34,5 % против 14,9 %')], 1494, 536, 300, 150, { size: 13, color: C.text2, ls: 1.08 });
  anim(1, 'Шаг 1 · 2,3', 'zoom');
  anim(1, 'Шаг 1 · Пояснение 2,3', 'fade', 300);

  s.addNotes([
    'СЛАЙД 5 · моё исследование (≈ 0:32)',
    '',
    'Я проверил это сам. Взял составы всех клубов НХЛ сезона 2024/25, это 853 игрока, распределил их по кварталам рождения и применил критерий хи-квадрат.',
    '[КЛИК 1 — «2,3×»] Посмотрите на синие столбики. Среди канадцев в январе–марте родились 35 % игроков, а в октябре–декабре всего 15 %. Разница в 2,3 раза.',
    '[КЛИК 2 — «Гипотеза подтвердилась»] Вероятность получить такое случайно меньше одной тысячной. Гипотеза подтвердилась: через сорок лет эффект никуда не делся.',
  ].join('\n'));
}

// ============================================================================
// 6 · ЗАГАДКА
// ============================================================================
const FRAME = { x: 26, y: 80, w: 840, h: 1000 }; // PNG c полями под тень; сама рамка 700×860 в точке (96;150)
{
  const s = newSlide('Контент', null, 6);
  kicker(s, '06', 'Загадка');
  img(s, '!!Шестёркин', 'frame.png', FRAME.x, FRAME.y, FRAME.w, FRAME.h, { alt: 'Игорь Шестёркин, вратарь, номер 31. Чтобы поставить фото: правый клик → Изменить рисунок' });
  text(s, 'Подпись фото имя', [R('Игорь Шестёркин')], 132, 878, 620, 40, { size: 18, bold: true, color: WHITE });
  text(s, 'Подпись фото роль', [R('ВРАТАРЬ · № 31')], 132, 924, 620, 28, { font: MONO, size: 10, color: WHITE, cs: 2 });

  text(s, 'Вопрос вступление', [R('Это '), R('Игорь Шестёркин', { bold: true, color: C.text1 }), R(' — один из лучших вратарей мира.')], 880, 178, 944, 50, { size: 17, color: C.text2 });
  text(s, 'Вопрос', [R('В каком месяце он родился?')], 880, 246, 944, 170, { font: BLACK, size: 34, ls: 0.98 });

  const GX = 880, GY = 452, GW = 944, gap = 14, cw = (GW - gap * 5) / 6, ch = 120;
  for (let k = 0; k < 12; k++) {
    const col = k % 6, row = Math.floor(k / 6);
    s.addText(MONTHS[k], {
      shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: X(20), x: X(GX + col * (cw + gap)), y: X(GY + row * (ch + gap)), w: X(cw), h: X(ch),
      margin: 0, align: 'center', valign: 'middle', fontFace: 'Arial', fontSize: 15, bold: true, color: C.text1, lang: LANG,
      fill: { color: C.background1, transparency: PAL.CARD_T + 4 }, line: { color: C.text1, transparency: THEMEKEY === 'dark' ? 80 : 86, width: 0.75 },
      shadow: shadow(), objectName: nm(`Месяц ${MONTHS[k]}`),
    });
  }
  // шаг 1: ответ — декабрь
  const dx = GX + 5 * (cw + gap), dy = GY + 1 * (ch + gap);
  s.addText([R('Дек', { breakLine: true }), R('30', { fontFace: MONO, fontSize: 11, bold: false })], {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: X(20), x: X(dx), y: X(dy), w: X(cw), h: X(ch), margin: 0,
    align: 'center', valign: 'middle', fontFace: 'Arial', fontSize: 15, bold: true, color: WHITE, lang: LANG,
    fill: { color: C.accent2 }, line: { type: 'none' }, shadow: shadow(), objectName: nm('Шаг 1 · Декабрь ответ'),
  });
  text(s, 'Шаг 1 · Ответ', [R('30 '), R('декабря', { color: C.accent2 })], 880, 748, 944, 120, { font: BLACK, size: 54, valign: 'middle' });
  text(s, 'Шаг 1 · Ответ пояснение', [R('1995 год · '), R('«Везина Трофи» 2022', { bold: true, color: C.text1 }), R(' — лучший вратарь НХЛ')], 880, 880, 944, 44, { size: 15, color: C.text2 });
  anim(1, 'Шаг 1 · Декабрь ответ', 'zoom');
  anim(1, 'Шаг 1 · Ответ', 'float', 250);
  anim(1, 'Шаг 1 · Ответ пояснение', 'fade', 600);

  s.addNotes([
    'СЛАЙД 6 · загадка (≈ 0:25 вместе с ответами зала)',
    '',
    'А теперь вопрос залу. Это Игорь Шестёркин, один из лучших вратарей мира. Как думаете, в каком месяце он родился?',
    '(выслушать 2–3 ответа)',
    '[КЛИК 1 — подсвечивается декабрь] 30 декабря.',
    '',
    'Фото: на месте джерси можно поставить фотографию — правый клик по картинке → «Изменить рисунок» (на слайде 7 сделать то же самое).',
  ].join('\n'));
}

// ============================================================================
// 7 · РОССИЯНЕ: КАРТИНА ОБРАТНАЯ  (переход «Трансформация» со слайда 6)
// ============================================================================
{
  const s = newSlide('Контент', null, 7);
  TRANS[slideNo] = 'morph';
  kicker(s, '07', 'Эффект аутсайдера');
  s.addText([R('У российских игроков картина '), R('обратная', { color: C.accent2 })], { placeholder: 'title' });
  s.addText([R('Fumarco, Gibbs, Jarvis, Rossi · PLoS ONE, 2017 · Gibbs, Jarvis, Dufur, 2012')], { placeholder: 'source' });
  const k = 300 / 700;
  img(s, '!!Шестёркин', 'frame.png', 96 - 70 * k, 256 - 70 * k, FRAME.w * k, FRAME.h * k, { alt: 'Игорь Шестёркин' });
  text(s, 'Шестёркин имя', [R('Игорь Шестёркин')], 96, 638, 320, 34, { size: 13.5, bold: true });
  s.addShape(pres.shapes.OVAL, { x: X(98), y: X(684), w: X(14), h: X(14), fill: { color: C.accent2 }, line: { type: 'none' }, objectName: nm('Шестёркин точка') });
  text(s, 'Шестёркин дата', [R('30.12.1995')], 122, 676, 260, 30, { font: MONO, size: 12, bold: true });

  glass(s, 'График карточка', 432, 252, 820, 520);
  text(s, 'График заголовок', [R('Игроки НХЛ по кварталу рождения, %')], 462, 276, 760, 30, { size: 13, bold: true });
  s.addShape(pres.shapes.OVAL, { x: X(464), y: X(321), w: X(13), h: X(13), fill: { color: C.accent2 }, line: { type: 'none' }, objectName: nm('Легенда Россия точка') });
  text(s, 'Легенда Россия', [R('Россия · n = 52')], 484, 314, 220, 26, { font: MONO, size: 10 });
  s.addShape(pres.shapes.OVAL, { x: X(704), y: X(321), w: X(13), h: X(13), fill: { color: C.accent3 }, line: { type: 'none' }, objectName: nm('Легенда Канада точка') });
  text(s, 'Легенда Канада', [R('Канада · n = 342')], 724, 314, 230, 26, { font: MONO, size: 10 });
  text(s, 'Легенда ожидание', [R('- - - ожидание ≈ 25 %')], 960, 314, 262, 26, { font: MONO, size: 10, color: C.accent4, align: 'right' });
  s.addChart([
    { type: pres.charts.BAR, data: [
      { name: 'Россия', labels: QLAB, values: [23.1, 21.2, 21.2, 34.6] },
      { name: 'Канада', labels: QLAB, values: [34.5, 28.1, 22.5, 14.9] }],
      options: { barDir: 'col', barGrouping: 'clustered', chartColors: [PAL.ORANGE, PAL.BLUE_MUTE], showValue: true, dataLabelPosition: 'outEnd',
        dataLabelFormatCode: '0.0', dataLabelFontSize: 12, dataLabelFontBold: true, dataLabelFontFace: '+mn-lt', dataLabelColor: PAL.INK } },
    { type: pres.charts.LINE, data: [{ name: 'Ожидание, %', labels: QLAB, values: EXPECT }],
      options: { chartColors: [PAL.INK3], lineSize: 1.5, lineDash: 'dash', lineDataSymbol: 'none', showValue: false } },
  ], chartOpts({ x: X(452), y: X(348), w: X(780), h: X(412), barGapWidthPct: 55, catAxisLabelFontSize: 12, objectName: nm('График Россия и Канада'),
    altText: 'Россияне в НХЛ: 23,1 %, 21,2 %, 21,2 %, 34,6 % по кварталам; канадцы: 34,5 %, 28,1 %, 22,5 %, 14,9 %' }));

  const players = [['Артемий Панарин', '30.10', 'Не выбран на драфте никем — сегодня звезда лиги', 'strip-oct.png'],
    ['Матвей Мичков', '09.12', 'Родился в декабре', 'strip-dec.png'], ['Иван Демидов', '10.12', 'Родился в декабре', 'strip-dec.png']];
  players.forEach(([n, d, t, strip], j) => {
    const y = 252 + j * 176;
    glass(s, `Шаг 1 · ${n} карточка`, 1288, y, 536, 160);
    text(s, `Шаг 1 · ${n}`, [R(n)], 1314, y + 20, 360, 32, { size: 15, bold: true });
    text(s, `Шаг 1 · ${n} дата`, [R(d)], 1680, y + 22, 118, 30, { font: MONO, size: 13, bold: true, align: 'right' });
    text(s, `Шаг 1 · ${n} текст`, [R(t)], 1314, y + 60, 484, 50, { size: 11.5, color: C.text2, ls: 1.05 });
    img(s, `Шаг 1 · ${n} месяцы`, strip, 1314, y + 122, 484, 16, { alt: 'Месяц рождения на ленте года' });
    [' карточка', '', ' дата', ' текст', ' месяцы'].forEach((p) => anim(1, `Шаг 1 · ${n}${p}`, 'float', j * 220));
  });

  glass(s, 'Шаг 2 · Честно карточка', 96, 800, 1728, 200);
  text(s, 'Шаг 2 · Честно заголовок', [R('ЧЕСТНО')], 136, 828, 800, 24, { font: MONO, size: 10, bold: true, color: C.accent4, cs: 2 });
  text(s, 'Шаг 2 · Честно текст', [R('Россиян всего 52: p = 0,47 — статистически это не доказано')], 136, 862, 790, 90, { size: 14, ls: 1.06 });
  line(s, 'Шаг 2 · Разделитель', 960, 828, 960, 972, { color: C.text1, width: 0.75, t: 82 });
  text(s, 'Шаг 2 · Аутсайдер заголовок', [R('ЭФФЕКТ АУТСАЙДЕРА')], 1000, 828, 790, 24, { font: MONO, size: 10, bold: true, color: C.accent2, cs: 2 });
  text(s, 'Шаг 2 · Аутсайдер текст', [R('В элите НХЛ поздних больше, и зарабатывают они больше')], 1000, 862, 790, 72, { size: 14, ls: 1.06 });
  text(s, 'Шаг 2 · Аутсайдер источник', [R('+51 % к зарплате у топ-10 % игроков · Fumarco et al., 2017')], 1000, 944, 790, 30, { font: MONO, size: 10, color: C.accent4 });
  ['Шаг 2 · Честно карточка', 'Шаг 2 · Честно заголовок', 'Шаг 2 · Честно текст', 'Шаг 2 · Разделитель', 'Шаг 2 · Аутсайдер заголовок', 'Шаг 2 · Аутсайдер текст', 'Шаг 2 · Аутсайдер источник']
    .forEach((n) => anim(2, n, 'float'));

  s.addNotes([
    'СЛАЙД 7 · эффект аутсайдера (≈ 0:33). Переход «Трансформация»: фото уменьшается, появляются данные.',
    '',
    'И это не исключение. Посмотрите на оранжевые столбики: у российских игроков картина обратная, больше всего родились в конце года.',
    '[КЛИК 1 — игроки] Панарин родился 30 октября, на драфте его не выбрал никто, а сегодня он звезда лиги. Мичков и Демидов родились в декабре.',
    '[КЛИК 2] Честно: россиян всего 52, и статистически это не доказано. Но это совпадает с эффектом аутсайдера: среди самой элиты НХЛ поздних больше, и зарабатывают они больше. Кто всё детство был самым маленьким, учился играть умнее.',
    '',
    'Справка: Панарин с февраля 2026 года играет за «Лос-Анджелес Кингз» (поэтому клуб на слайде не указан). Мичков — «Филадельфия», Демидов — «Монреаль».',
  ].join('\n'));
}

// ============================================================================
// 8 · НЕ ТОЛЬКО СПОРТ
// ============================================================================
{
  const s = newSlide('Контент', 'Шире спорта', 8);
  kicker(s, '08', 'Не только спорт');
  s.addText([R('В школе тот же механизм: самым младшим в классе СДВГ ставят на 34 % чаще')], { placeholder: 'title' });
  s.addText([R('Layton, Barnett, Hicks, Jena · New England Journal of Medicine, 2018 · ≈ 407 000 детей · штаты с границей набора 1 сентября')], { placeholder: 'source' });
  text(s, 'Подпись учебного года', [R('УЧЕБНЫЙ ГОД В США: КЛАСС НАБИРАЮТ ПО ДАТЕ РОЖДЕНИЯ · ГРАНИЦА — 1 СЕНТЯБРЯ')], 96, 286, 1728, 28, { font: MONO, size: 10, color: C.accent4, cs: 1.5 });
  monthChips(s, 'Школа', 96, 332, 1728, 84, {
    start: 9, size: 13,
    style: (m) => (m === 9 ? { fill: { color: C.accent1 }, color: WHITE, line: { type: 'none' } } : m === 8 ? { fill: { color: C.accent2 }, color: WHITE, line: { type: 'none' } } : {}),
  });
  text(s, 'Старшие', [R('← самые старшие')], 96, 428, 500, 30, { font: MONO, size: 10.5, color: C.text2 });
  text(s, 'Младшие', [R('самые младшие →')], 1324, 428, 500, 30, { font: MONO, size: 10.5, color: C.text2, align: 'right' });

  text(s, 'Шаг 1 · 34 процента', [R('+34'), R(' %', { color: C.accent2 })], 96, 556, 780, 220, { font: BLACK, size: 88, valign: 'middle' });
  text(s, 'Шаг 1 · Пояснение 34', [R('диагнозов СДВГ у рождённых в августе — по сравнению с сентябрём')], 96, 790, 760, 90, { size: 16, color: C.text2, ls: 1.08 });
  anim(1, 'Шаг 1 · 34 процента', 'zoom');
  anim(1, 'Шаг 1 · Пояснение 34', 'fade', 300);

  glass(s, 'Шаг 1 · СДВГ карточка', 900, 560, 924, 440);
  text(s, 'Шаг 1 · СДВГ подпись', [R('ДИАГНОЗОВ СДВГ НА 10 000 ДЕТЕЙ')], 940, 592, 840, 26, { font: MONO, size: 10, color: C.accent4, cs: 1.5 });
  const rows = [['Август', 'самые младшие', 85.1, C.accent2], ['Сентябрь', 'самые старшие', 63.6, C.accent1]];
  rows.forEach(([m, sub, v, col], j) => {
    const y = 646 + j * 96;
    text(s, `Шаг 1 · ${m}`, [R(m)], 940, y, 220, 30, { size: 14, bold: true });
    text(s, `Шаг 1 · ${m} пояснение`, [R(sub)], 940, y + 32, 220, 24, { font: MONO, size: 9.5, color: C.accent4 });
    s.addText(String(v).replace('.', ','), {
      shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: X(10), x: X(1170), y: X(y + 2), w: X((v / 85.1) * 610), h: X(56),
      margin: [0, X(16) * 72, 0, 0], align: 'right', valign: 'middle', fontFace: BLACK, fontSize: 14, color: WHITE, lang: LANG,
      fill: { color: col }, line: { type: 'none' }, objectName: nm(`Шаг 1 · ${m} столбик`),
    });
  });
  line(s, 'Шаг 1 · СДВГ разделитель', 940, 862, 1784, 862, { color: C.text1, width: 0.75, t: 85 });
  text(s, 'Шаг 1 · СДВГ вывод', [R('Обычную незрелость принимают за расстройство')], 940, 884, 840, 90, { size: 17, bold: true, ls: 1.05 });
  ['Шаг 1 · СДВГ карточка', 'Шаг 1 · СДВГ подпись'].forEach((n) => anim(1, n, 'float', 150));
  rows.forEach(([m], j) => {
    anim(1, `Шаг 1 · ${m}`, 'fade', 350 + j * 150);
    anim(1, `Шаг 1 · ${m} пояснение`, 'fade', 350 + j * 150);
    anim(1, `Шаг 1 · ${m} столбик`, 'wipeRight', 450 + j * 150);
  });
  ['Шаг 1 · СДВГ разделитель', 'Шаг 1 · СДВГ вывод'].forEach((n) => anim(1, n, 'fade', 900));

  s.addNotes([
    'СЛАЙД 8 · не только спорт (≈ 0:12)',
    '',
    'И это не только про спорт.',
    '[КЛИК 1] В США самым младшим в классе диагноз СДВГ ставят на 34 % чаще: обычную незрелость принимают за расстройство.',
    '',
    'Справка: 85,1 против 63,6 диагноза на 10 000 детей; в штатах с другой датой набора такой разницы нет.',
  ].join('\n'));
}

// ============================================================================
// 9 · ЧТО ДЕЛАТЬ
// ============================================================================
{
  const s = newSlide('Контент', null, 9);
  kicker(s, '09', 'Что делать');
  s.addText([R('Перекос можно ослабить: сравнивать зрелость, а не дату рождения')], { placeholder: 'title' });
  const cards = [
    ['01 · БИО-БЭНДИНГ', 'Группы по зрелости', [R('В английских футбольных академиях детей делят '), R('по биологической зрелости', { bold: true, color: C.text1 }), R(', а не по году рождения.')], 'Cumming et al. · J. Sports Sciences, 2018', 'fix-1.png', 'Дети разного роста сгруппированы по зрелости'],
    ['02 · НОМЕРА ПО ВОЗРАСТУ', 'Скаут видит возраст', [R('Номер на футболке = место по возрасту в команде: '), R('перекос в отборе почти исчезает', { bold: true, color: C.text1 }), R('.')], 'Mann, van Ginneken · J. Sports Sciences, 2017', 'fix-2.png', 'Футболки с номерами от старшего к младшему'],
    ['03 · СРАВНЕНИЕ С СОБОЙ', 'Прогресс, а не рейтинг', [R('Оценивать, '), R('как ребёнок вырос сам', { bold: true, color: C.text1 }), R(', а не какое место он занял среди ровесников.')], 'Принцип личного прогресса', 'fix-3.png', 'Линия личного прогресса ребёнка'],
  ];
  cards.forEach(([no, title, body, src, pic, alt], k) => {
    const x = 96 + k * 584, y = 268, st = k + 1;
    glass(s, `Шаг ${st} · Карточка`, x, y, 560, 732);
    img(s, `Шаг ${st} · Иллюстрация`, pic, x + 28, y + 28, 504, 330, { alt });
    text(s, `Шаг ${st} · Номер`, [R(no)], x + 32, y + 384, 496, 26, { font: MONO, size: 10, bold: true, color: C.accent1, cs: 1.5 });
    text(s, `Шаг ${st} · Заголовок`, [R(title)], x + 32, y + 418, 496, 52, { size: 21, bold: true });
    text(s, `Шаг ${st} · Текст`, body, x + 32, y + 482, 496, 180, { size: 14, color: C.text2, ls: 1.1 });
    text(s, `Шаг ${st} · Источник`, [R(src)], x + 32, y + 676, 496, 28, { font: MONO, size: 9, color: C.accent4 });
    ['Карточка', 'Иллюстрация', 'Номер', 'Заголовок', 'Текст', 'Источник'].forEach((p) => anim(st, `Шаг ${st} · ${p}`, 'float'));
  });
  s.addNotes([
    'СЛАЙД 9 · что делать (≈ 0:18)',
    '',
    'Решения есть.',
    '[КЛИК 1] В английских футбольных академиях детей делят по биологической зрелости, а не по году рождения.',
    '[КЛИК 2] А если дать скауту номера на футболках по возрасту, перекос в отборе почти исчезает.',
    '[КЛИК 3] (по желанию) И ещё — сравнивать ребёнка с ним самим, а не с ровесниками.',
    '',
    'Осторожно на вопросах: эксперимент с номерами проводили со скаутами футбольного клуба; в хоккейной проверке предварительные результаты неоднозначны.',
  ].join('\n'));
}

// ============================================================================
// 10 · ВЫВОДЫ
// ============================================================================
{
  const s = newSlide('Контент', 'Итоги', 10);
  kicker(s, '10', 'Выводы');
  s.addText([R('Итак')], { placeholder: 'title' });
  const items = [
    ['1', [R('Эффект реален '), R('и жив в НХЛ до сих пор', { bold: false, color: C.text2 })], 'канадцы: 2,3× · p < 0,001', C.accent1, C.accent5],
    ['2', [R('Причина не в таланте, '), R('а в системе отбора и ожиданиях взрослых', { bold: false, color: C.text2 })], 'Пигмалион · эффект Матфея', C.accent1, C.accent5],
    ['3', [R('Его можно '), R('ослабить', { bold: false, color: C.text2 })], 'био-бэндинг · номера по возрасту', C.accent2, C.accent6],
  ];
  items.forEach(([n, body, chip, col, soft], k) => {
    const x = 96 + k * 584, y = 220, st = k + 1;
    glass(s, `Шаг ${st} · Вывод карточка`, x, y, 560, 300);
    text(s, `Шаг ${st} · Вывод номер`, [R(n)], x + 32, y + 14, 120, 80, { font: BLACK, size: 38, color: col, valign: 'middle' });
    text(s, `Шаг ${st} · Вывод текст`, body, x + 32, y + 100, 496, 136, { size: 17, bold: true, ls: 1.04 });
    s.addText(chip, {
      shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: X(20), x: X(x + 32), y: X(y + 244), w: X(Math.min(496, 40 + chip.length * 12.5)), h: X(38),
      margin: 0, align: 'center', valign: 'middle', fontFace: MONO, fontSize: 10, color: C.text1, lang: LANG,
      fill: { color: soft, transparency: THEMEKEY === 'dark' ? 0 : 15 }, line: { type: 'none' }, objectName: nm(`Шаг ${st} · Вывод метка`),
    });
    ['Вывод карточка', 'Вывод номер', 'Вывод текст', 'Вывод метка'].forEach((p) => anim(st, `Шаг ${st} · ${p}`, 'float'));
  });
  text(s, 'Шаг 4 · Финал 1', [R('Календарь не определяет талант.')], 96, 566, 1728, 96, { font: BLACK, size: 36, valign: 'middle' });
  text(s, 'Шаг 4 · Финал 2', [R('Он определяет только, '), R('кого мы успеваем заметить', { bold: true, color: C.accent1 }), R('.')], 96, 668, 1728, 56, { size: 24, color: C.text2 });
  anim(4, 'Шаг 4 · Финал 1', 'float');
  anim(4, 'Шаг 4 · Финал 2', 'float', 300);
  const lit = (m) => ({ fill: { color: C.accent5, transparency: THEMEKEY === 'dark' ? 0 : 5 }, color: C.text1, line: { color: C.accent1, transparency: 55, width: 1 } });
  monthChips(s, 'Шаг 4 · Месяц', 96, 762, 1728, 72, { style: lit, size: 12.5 });
  MONTHS.forEach((m, j) => anim(4, `Шаг 4 · Месяц ${m}`, 'zoom', 600 + j * 70));
  text(s, 'Шаг 4 · Любой месяц', [R('ХОККЕИСТЫ РОЖДАЮТСЯ В ЛЮБОЙ МЕСЯЦ')], 96, 846, 1728, 28, { font: MONO, size: 10, color: C.accent4, align: 'center', cs: 2 });
  text(s, 'Шаг 4 · Спасибо', [R('Спасибо!')], 96, 896, 700, 64, { font: BLACK, size: 28, valign: 'middle' });
  text(s, 'Шаг 4 · Вопросы', [R('Готов ответить на вопросы')], 96, 962, 700, 40, { size: 15, color: C.text2 });
  anim(4, 'Шаг 4 · Любой месяц', 'fade', 1500);
  anim(4, 'Шаг 4 · Спасибо', 'float', 1700);
  anim(4, 'Шаг 4 · Вопросы', 'float', 1800);

  s.addNotes([
    'СЛАЙД 10 · выводы (≈ 0:25)',
    '',
    'Итак:',
    '[КЛИК 1] эффект реален и жив в НХЛ до сих пор.',
    '[КЛИК 2] Его причина не в таланте, а в системе отбора и ожиданиях взрослых.',
    '[КЛИК 3] И его можно ослабить.',
    '[КЛИК 4] Хоккеисты рождаются не в январе. Хоккеисты рождаются в любой месяц, просто январским чаще дают шанс. Календарь не определяет талант. Он определяет только, кого мы успеваем заметить.',
    'Спасибо! Готов ответить на вопросы.',
  ].join('\n'));
}

// ============================================================================
// 11 · ПРИЛОЖЕНИЕ: ДАННЫЕ
// ============================================================================
const DATA = [
  ['Вся лига', 853, [33.3, 25.1, 21.9, 19.7], '1,69', '39,3', '< 0,001', true],
  ['Канада', 342, [34.5, 28.1, 22.5, 14.9], '2,31', '29,9', '< 0,001', true],
  ['США', 259, [35.1, 22.4, 20.1, 22.4], '1,57', '15,5', '0,001', true],
  ['Швеция', 80, [26.2, 20.0, 27.5, 26.2], '1,00', '1,1', '0,79', false],
  ['Россия', 52, [23.1, 21.2, 21.2, 34.6], '0,67', '2,5', '0,47', false],
];
const QA = [
  'ОТВЕТЫ НА ВОПРОСЫ КОМИССИИ',
  '',
  '1. Может, в начале года просто рождается больше детей? — Нет. В Канаде и России рождаемость по месяцам почти равномерна, а летом детей рождается даже немного больше. С поправкой на это эффект получается ещё сильнее.',
  '2. Почему у шведов эффекта нет? — В Швеции детский хоккей строится на принципе «спорт для всех»: жёсткий отбор начинается позже. Это косвенно подтверждает, что отложенный отбор работает. Но 80 человек — небольшая выборка, поэтому это предположение.',
  '3. Почему результат по россиянам не доказан? — 52 человека — слишком мало: p = 0,47, примерно в половине случаев такое распределение получилось бы случайно. Следующий этап — все российские игроки КХЛ и молодёжных сборных.',
  '4. Значит, родителям стоит планировать рождение детей? — Нет. Эффект ослабевает с возрастом, а у «поздних» есть плюс — эффект аутсайдера. Правильнее менять систему отбора, а не дату рождения.',
  '5. При чём тут психология, если дело в физиологии? — Физиология даёт стартовую разницу в несколько месяцев. В разницу на всю карьеру её превращают психологические механизмы: ожидания тренера (эффект Пигмалиона), накопление преимуществ (эффект Матфея) и мотивация младших, которые чаще бросают спорт.',
  '6. Почему НХЛ, а не КХЛ? — У НХЛ открытые и полные данные по составам, в лиге играют представители многих стран — можно сравнить страны. Анализ КХЛ — продолжение работы.',
  '',
  'Методика: данные — составы всех клубов НХЛ 2024/25 (api-web.nhle.com, выгрузка github.com/mforrest13/NHL-stats); кварталы рождения I–IV; ожидание пропорционально числу дней в квартале (90,25 · 91 · 92 · 92); χ²-критерий согласия, 3 степени свободы.',
  '+51 % — Fumarco et al., 2017: среди 10 % самых высокооплачиваемых игроков НХЛ рождённые в IV квартале получают на 51 % больше рождённых в I квартале.',
].join('\n');
{
  const s = newSlide('Контент', 'Приложения', 11);
  kicker(s, 'П1', 'Для вопросов');
  s.addText([R('Данные по странам: НХЛ, сезон 2024/25')], { placeholder: 'title' });
  s.addText([R('Ожидаемые доли по числу дней: 24,7 · 24,9 · 25,2 · 25,2 % · χ²-критерий согласия, df = 3, α = 0,05 · расчёт — analysis.py')], { placeholder: 'source' });
  text(s, 'Метка приложения', [R('ПРИЛОЖЕНИЕ')], 1620, 56, 204, 30, { font: MONO, size: 10, bold: true, color: C.accent4, align: 'right', cs: 2 });
  glass(s, 'Таблица карточка', 96, 236, 1728, 600);
  const H = (t, al = 'right') => ({ text: t, options: { fontFace: MONO, fontSize: 10, color: C.accent4, bold: true, align: al, valign: 'bottom', lang: LANG, border: [{ type: 'none' }, { type: 'none' }, { pt: 1.25, color: PAL.GRID }, { type: 'none' }] } });
  const rows = [[H('Группа', 'left'), H('I кв., %'), H('II кв., %'), H('III кв., %'), H('IV кв., %'), H('I / IV'), H('χ²(3)'), H('p'), H('Вывод', 'left')]];
  DATA.forEach(([g, n, q, ratio, chi, p, sig]) => {
    const fill = g === 'Канада' ? { color: PAL.BLUE_SOFT, transparency: THEMEKEY === 'dark' ? 0 : 35 } : g === 'Россия' ? { color: PAL.ORANGE_SOFT, transparency: THEMEKEY === 'dark' ? 0 : 35 } : undefined;
    const cell = (t, o = {}) => ({ text: t, options: { fontFace: 'Arial', fontSize: 15, color: C.text1, align: 'right', valign: 'middle', lang: LANG, fill, border: [{ type: 'none' }, { type: 'none' }, { pt: 0.75, color: PAL.GRID }, { type: 'none' }], ...o } });
    rows.push([
      { text: [R(g, { bold: true, breakLine: true }), R(`n = ${n}`, { fontFace: MONO, fontSize: 10, color: C.accent4 })], options: { fontFace: 'Arial', fontSize: 15, color: C.text1, align: 'left', valign: 'middle', lang: LANG, fill, border: [{ type: 'none' }, { type: 'none' }, { pt: 0.75, color: PAL.GRID }, { type: 'none' }] } },
      ...q.map((v) => cell(String(v.toFixed(1)).replace('.', ','))),
      cell(ratio), cell(chi), cell(p, { fontFace: MONO, fontSize: 13 }),
      cell(sig ? 'значимо' : 'не значимо', { align: 'left', bold: sig, color: sig ? C.text1 : C.accent4, fontSize: 13 }),
    ]);
  });
  s.addTable(rows, { x: X(126), y: X(262), w: X(1668), colW: [300, 160, 160, 160, 160, 140, 140, 190, 258].map(X), rowH: [X(60), X(96), X(96), X(96), X(96), X(96)], margin: [0.03, 0.1, 0.03, 0.1], objectName: nm('Таблица данных') });
  text(s, 'Пояснение таблицы', [R('Канада и США — значимый перекос в пользу ранних; Швеция — эффекта нет; Россия — обратная картина, но выборка мала.')], 96, 868, 1728, 60, { size: 14, color: C.text2 });
  s.addNotes(QA);
}

// ============================================================================
// 12 · ПРИЛОЖЕНИЕ: ИСТОЧНИКИ
// ============================================================================
{
  const s = newSlide('Контент', null, 12);
  kicker(s, 'П2', 'Источники');
  s.addText([R('Источники')], { placeholder: 'title' });
  text(s, 'Метка приложения', [R('ПРИЛОЖЕНИЕ')], 1620, 56, 204, 30, { font: MONO, size: 10, bold: true, color: C.accent4, align: 'right', cs: 2 });
  const refs = [
    ['1', 'Barnsley R. H., Thompson A. H., Barnsley P. E.', ' Hockey success and birthdate: The relative age effect // CAHPER Journal. 1985.'],
    ['1а', 'Barnsley R. H.', ' 40-years of relative age effects: life is not fair! // Frontiers in Sports and Active Living. 2025. doi:10.3389/fspor.2024.1516173'],
    ['2', 'Gladwell M.', ' Outliers: The Story of Success. 2008 (рус. изд. «Гении и аутсайдеры»).'],
    ['3', 'Cobley S., Baker J., Wattie N., McKenna J.', ' Annual age-grouping and athlete development: a meta-analytical review of relative age effects in sport // Sports Medicine. 2009. Vol. 39(3). P. 235–256.'],
    ['4', 'Deaner R. O., Lowen A., Cobley S.', ' Born at the wrong time: selection bias in the NHL draft // PLoS ONE. 2013.'],
    ['5', 'Gibbs B. G., Jarvis J. A., Dufur M. J.', ' The rise of the underdog? The relative age effect reversal among Canadian-born NHL hockey players // International Review for the Sociology of Sport. 2012.'],
    ['6', 'Fumarco L., Gibbs B. G., Jarvis J. A., Rossi G.', ' The relative age effect reversal among the National Hockey League elite // PLoS ONE. 2017.'],
    ['7', 'Lemez S., Baker J., Horton S., Wattie N., Weir P.', ' Examining the relationship between relative age, competition level, and dropout rates in male youth ice-hockey players // Scand. J. Med. Sci. Sports. 2014. Vol. 24(6).'],
    ['8', 'Bedard K., Dhuey E.', ' The persistence of early childhood maturity: international evidence of long-run age effects // Quarterly Journal of Economics. 2006. Vol. 121(4).'],
    ['9', 'Morrow R. L. et al.', ' Influence of relative age on diagnosis and treatment of ADHD in children // CMAJ. 2012. Vol. 184(7).'],
    ['10', 'Layton T. J., Barnett M. L., Hicks T. R., Jena A. B.', ' Attention deficit–hyperactivity disorder and month of school enrollment // NEJM. 2018. Vol. 379(22).'],
    ['11', 'Mann D. L., van Ginneken P. J. M. A.', ' Age-ordered shirt numbering reduces the selection bias associated with the relative age effect // Journal of Sports Sciences. 2017. Vol. 35(8).'],
    ['12', 'Cumming S. P. et al.', ' Premier League academy soccer players’ experiences of competing in a tournament bio-banded for biological maturation // Journal of Sports Sciences. 2018.'],
    ['13', 'Rosenthal R., Jacobson L.', ' Pygmalion in the Classroom. 1968.'],
    ['14', 'Merton R. K.', ' The Matthew Effect in Science // Science. 1968.'],
    ['15', 'Данные:', ' составы клубов НХЛ сезона 2024/25 (api-web.nhle.com), выгрузка github.com/mforrest13/NHL-stats.'],
  ];
  const col = (list) => list.flatMap(([n, a, rest], i) => [
    R(n + '  ', { fontFace: MONO, bold: true, color: C.accent1, fontSize: 10 }),
    R(a, { bold: true, color: C.text1 }),
    R(rest, { breakLine: i < list.length - 1 }),
  ]);
  text(s, 'Источники 1–7', col(refs.slice(0, 8)), 96, 236, 840, 770, { size: 11, color: C.text2, ls: 1.06, psa: 7 });
  text(s, 'Источники 8–15', col(refs.slice(8)), 984, 236, 840, 770, { size: 11, color: C.text2, ls: 1.06, psa: 7 });
  s.addNotes(QA);
}

// ---------- запись ----------
(async () => {
  const file = path.join(OUTDIR, `deck-${THEMEKEY}.pptx`);
  await pres.writeFile({ fileName: file });
  await applyTheme(file, THEME);
  fs.writeFileSync(path.join(OUTDIR, `anim-${THEMEKEY}.json`), JSON.stringify({ anim: ANIM, trans: TRANS }, null, 1));
  console.log('OK', file, 'слайдов:', slideNo);
})();
