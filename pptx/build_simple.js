// Простая версия презентации: белый фон, Calibri, стандартные фигуры и диаграммы PowerPoint.
//   node build_simple.js  →  out/deck-simple.pptx + out/anim-simple.json  →  python3 postprocess.py simple
// Цвета: синий — «ранние» (январь–март), оранжевый — «поздние» (октябрь–декабрь), всё остальное серое.
const path = require('path');
const fs = require('fs');
const pptxgen = require('pptxgenjs');
const SKILL = '/root/.claude/skills/synced/2d64efaf-fc81-4729-a14f-24728a20f0e8_08b58b4f-6847-4399-ba9d-bded480b3897/pptx';
const { applyTheme } = require(SKILL + '/scripts/apply_theme.js');

const P = { INK: '1F2933', GRAY: '616E7C', LIGHT: 'E4E9F0', BLUE: '2A78D6', BLUE_L: 'DCE9F8', ORANGE: 'E2622F', ORANGE_L: 'FBE3D8', BG: 'FFFFFF', SOFT: 'F3F6FA' };
const THEME = {
  name: 'Простая', headFontFace: 'Calibri', bodyFontFace: 'Calibri',
  colors: { dk1: P.INK, lt1: P.BG, dk2: P.GRAY, lt2: P.SOFT, accent1: P.BLUE, accent2: P.ORANGE, accent3: P.BLUE_L, accent4: P.GRAY, accent5: P.LIGHT, accent6: P.ORANGE_L, hlink: P.BLUE, folHlink: P.GRAY },
};
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.theme = { headFontFace: 'Calibri', bodyFontFace: 'Calibri' };
pres.title = 'Эффект относительного возраста';
pres.author = 'Имя Фамилия';
const C = pres.SchemeColor;
const X = (px) => Math.round((px / 144) * 10000) / 10000;
const LANG = 'ru-RU';
const R = (t, o = {}) => ({ text: t, options: { lang: LANG, ...o } });
const M = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];

const ANIM = {}, TRANS = {};
let n = 0;
const anim = (step, name) => { const a = (ANIM[n] = ANIM[n] || []); (a[step - 1] = a[step - 1] || []).push({ name, fx: 'fade', delay: 0 }); };

pres.defineSlideMaster({
  title: 'Слайд',
  background: { color: P.BG },
  objects: [
    { placeholder: { options: { name: 'title', type: 'title', x: X(96), y: X(64), w: X(1728), h: X(140), fontFace: 'Calibri', fontSize: 36, bold: true, color: C.text1, margin: 0, valign: 'top', align: 'left', lang: LANG }, text: '' } },
    { placeholder: { options: { name: 'source', type: 'body', x: X(96), y: X(1010), w: X(1560), h: X(40), fontFace: 'Calibri', fontSize: 11, color: C.accent4, margin: 0, valign: 'middle', align: 'left', lang: LANG }, text: '' } },
  ],
  slideNumber: { x: X(1720), y: X(1010), w: X(104), h: X(40), fontFace: 'Calibri', fontSize: 12, color: C.accent4, align: 'right' },
});
pres.defineSlideMaster({ title: 'Титул', background: { color: P.BG }, objects: [] });

function slide(layout, title, source) {
  n++;
  const s = pres.addSlide({ masterName: layout });
  if (title) s.addText(Array.isArray(title) ? title : [R(title)], { placeholder: 'title' });
  if (source) s.addText([R(source)], { placeholder: 'source' });
  TRANS[n] = 'fade';
  return s;
}
const names = new Set();
const nm = (x) => { const k = n + '|' + x; if (names.has(k)) throw new Error('повтор: ' + k); names.add(k); return x; };
function t(s, name, runs, x, y, w, h, o = {}) {
  s.addText(typeof runs === 'string' ? [R(runs)] : runs, {
    x: X(x), y: X(y), w: X(w), h: X(h), margin: 0, isTextBox: true, fontFace: 'Calibri', fontSize: o.size ?? 20, bold: !!o.bold,
    color: o.color ?? C.text1, align: o.align ?? 'left', valign: o.valign ?? 'top', lang: LANG, lineSpacingMultiple: o.ls ?? 1.05,
    objectName: nm(name), paraSpaceAfter: o.psa,
  });
}
function box(s, name, x, y, w, h, fill = P.SOFT, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: X(x), y: X(y), w: X(w), h: X(h), rectRadius: X(14), fill: { color: fill }, line: o.line ?? { type: 'none' }, objectName: nm(name) });
}
function months(s, prefix, x0, y, w, h, color, opts = {}) {
  const gap = 8, cw = (w - gap * 11) / 12, start = opts.start ?? 1;
  for (let k = 0; k < 12; k++) {
    const m = ((start - 1 + k) % 12) + 1;
    const c = color(m);
    s.addText(M[m - 1], {
      shape: pres.shapes.RECTANGLE, x: X(x0 + k * (cw + gap)), y: X(y), w: X(cw), h: X(h), margin: 0, align: 'center', valign: 'middle',
      fontFace: 'Calibri', fontSize: opts.size ?? 16, bold: true, color: c === P.LIGHT ? P.GRAY : 'FFFFFF', fill: { color: c }, line: { type: 'none' },
      lang: LANG, objectName: nm(`${prefix} ${M[m - 1]}`),
    });
  }
}
const QLAB = ['I кв.\nянв–мар', 'II кв.\nапр–июн', 'III кв.\nиюл–сен', 'IV кв.\nокт–дек'];
const chartBase = (o) => Object.assign({
  showLegend: false, showTitle: false, barGapWidthPct: 60, valAxisMinVal: 0, valAxisMajorUnit: 10, valAxisLabelFormatCode: '0',
  valAxisLabelColor: P.GRAY, valAxisLabelFontSize: 12, valAxisLabelFontFace: '+mn-lt', valAxisLineShow: false,
  valGridLine: { color: P.LIGHT, size: 0.75 }, catGridLine: { style: 'none' }, catAxisLabelColor: P.INK, catAxisLabelFontSize: 14,
  catAxisLabelFontFace: '+mn-lt', showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 14, dataLabelFontBold: true,
  dataLabelFontFace: '+mn-lt', dataLabelColor: P.INK,
}, o);

// 1 · Титул
{
  const s = slide('Титул');
  t(s, 'Название', 'Эффект относительного возраста', 96, 250, 1728, 110, { size: 48, bold: true });
  t(s, 'Подзаголовок', [R('Почему хоккеисты '), R('рождаются в январе', { color: C.accent1, bold: true })], 96, 370, 1728, 70, { size: 30, color: C.accent4 });
  t(s, 'Автор', 'Имя Фамилия', 96, 520, 1200, 50, { size: 24, bold: true });
  t(s, 'Направление', 'Образовательное направление', 96, 572, 1200, 44, { size: 20, color: C.accent4 });
  months(s, 'Месяц', 96, 800, 1728, 80, (m) => (m <= 3 ? P.BLUE : m >= 10 ? P.ORANGE : P.LIGHT));
  t(s, 'Подпись месяцев', 'Марафон лекториев «Поток»', 96, 910, 1728, 40, { size: 14, color: C.accent4 });
  s.addNotes(['Добрый день! Давайте начнём с эксперимента.', 'Поднимите руку те, кто родился в январе, феврале или марте. (пауза, оглядеть зал)', 'А теперь те, кто родился в октябре, ноябре или декабре. (пауза)', 'Будь мы канадскими хоккеистами, у первой группы шансов попасть в НХЛ было бы больше чем в два раза. И дело не в генах, а в календаре.'].join('\n'));
}

// 2 · 1984
{
  const s = slide('Слайд', '1984: как открыли эффект', 'Barnsley, Thompson, Barnsley // CAHPER Journal, 1985');
  t(s, 'История', 'Психолог Роджер Барнсли с женой Паулой пришли на юниорский хоккейный матч в Канаде. Паула листала программку матча и заметила:', 96, 220, 900, 200, { size: 21, color: C.accent4, ls: 1.1 });
  box(s, 'Цитата фон', 96, 440, 900, 210, P.BLUE_L);
  t(s, 'Цитата', '«Смотри, они почти все родились в январе, феврале и марте»', 136, 470, 820, 150, { size: 26, bold: true });
  t(s, 'Определение', [R('Так открыли '), R('эффект относительного возраста', { bold: true, color: C.text1 }), R(': в одной возрастной группе выигрывают те, кто старше.')], 96, 690, 900, 110, { size: 21, color: C.accent4, ls: 1.1 });
  box(s, 'Цель фон', 1080, 230, 744, 260, P.SOFT);
  t(s, 'Цель заголовок', 'Цель', 1120, 260, 660, 40, { size: 20, bold: true, color: C.accent1 });
  t(s, 'Цель', 'Проверить, жив ли этот эффект сегодня', 1120, 310, 660, 150, { size: 26, bold: true });
  box(s, 'Гипотеза фон', 1080, 520, 744, 300, P.SOFT);
  t(s, 'Гипотеза заголовок', 'Гипотеза', 1120, 550, 660, 40, { size: 20, bold: true, color: C.accent1 });
  t(s, 'Гипотеза', 'Среди игроков НХЛ родившихся в начале года значимо больше, чем родившихся в конце', 1120, 600, 660, 200, { size: 24, bold: true });
  ['Цель фон', 'Цель заголовок', 'Цель', 'Гипотеза фон', 'Гипотеза заголовок', 'Гипотеза'].forEach((x) => anim(1, x));
  s.addNotes('В 1984 году канадский психолог Роджер Барнсли с женой Паулой пришёл на юниорский хоккейный матч. Паула листала программку и вдруг сказала: «Смотри, они почти все родились в январе, феврале и марте». Барнсли проверил профессионалов: та же картина. Так открыли эффект относительного возраста.\n[КЛИК] Моя цель — проверить, жив ли этот эффект сегодня. Гипотеза: среди игроков НХЛ родившихся в начале года значимо больше, чем родившихся в конце.');
}

// 3 · Снежный ком
{
  const s = slide('Слайд', 'Почему так происходит: снежный ком', 'Rosenthal, Jacobson, 1968; Merton, 1968; Lemez et al., 2014');
  t(s, 'Вступление', [R('Дети в одной команде по году рождения: '), R('1 января', { bold: true, color: C.accent1 }), R(' и '), R('31 декабря', { bold: true, color: C.accent2 }), R(' — почти год разницы.')], 96, 220, 1728, 50, { size: 24 });
  box(s, '11 фон', 96, 300, 1728, 150, P.SOFT);
  t(s, '11 процентов', '11 %', 136, 315, 260, 120, { size: 66, bold: true, color: C.accent2, valign: 'middle' });
  t(s, '11 пояснение', '364 дня — это больше десятой части жизни девятилетнего', 420, 315, 1360, 120, { size: 26, valign: 'middle' });
  const steps = [['1. Отбор', 'Тренер берёт того, кто сильнее сейчас'], ['2. Ресурсы', 'Отобранным — лучшие тренеры и больше льда'], ['3. Ожидания', 'Эффект Пигмалиона: в кого верят, тот растёт'], ['4. Уход', 'Младшие сидят на скамейке и бросают спорт']];
  steps.forEach(([h, d], k) => {
    const x = 96 + k * 444;
    const col = k === 3 ? P.ORANGE : P.BLUE;
    s.addText([R(h, { bold: true, fontSize: 24, breakLine: true }), R(d, { fontSize: 18 })], {
      shape: pres.shapes.RECTANGLE, x: X(x), y: X(500), w: X(400), h: X(230), margin: 14, fontFace: 'Calibri', color: 'FFFFFF', fill: { color: col },
      line: { type: 'none' }, valign: 'top', lang: LANG, objectName: nm(`Шаг ${k + 1}`), lineSpacingMultiple: 1.05,
    });
    if (k < 3) s.addShape(pres.shapes.RIGHT_ARROW, { x: X(x + 404), y: X(595), w: X(36), h: X(40), fill: { color: P.GRAY }, line: { type: 'none' }, objectName: nm(`Стрелка ${k + 1}`) });
    anim(k + 1, `Шаг ${k + 1}`);
    if (k > 0) anim(k + 1, `Стрелка ${k}`);
  });
  t(s, 'Пророчество', [R('Самоисполняющееся пророчество: ', { bold: true, color: C.accent1 }), R('решили, что январские сильнее, — и сами сделали их сильнее.')], 96, 790, 1728, 100, { size: 24 });
  anim(5, 'Пророчество');
  s.addNotes('В детском хоккее команды собирают по году рождения. Мальчик, родившийся 1 января, и мальчик, родившийся 31 декабря, играют в одной команде, хотя между ними почти год. Для девятилетнего это больше десятой части жизни.\n[КЛИК] Дальше всё катится как снежный ком. Тренер берёт того, кто сильнее сейчас, и путает зрелость с талантом.\n[КЛИК] Отобранные получают лучших тренеров и больше льда.\n[КЛИК] Работает эффект Пигмалиона: тот, в кого верят, растёт быстрее.\n[КЛИК] А младшие сидят на скамейке и бросают спорт.\n[КЛИК] Получается самоисполняющееся пророчество: мы решили, что январские сильнее, и сами сделали их сильнее.');
}

// 4 · Масштаб
{
  const s = slide('Слайд', 'Эффект находят почти везде', 'Cobley et al. // Sports Medicine, 2009; Barnsley // Frontiers in Sports and Active Living, 2025');
  [['38', 'исследований'], ['14', 'видов спорта'], ['16', 'стран']].forEach(([v, l], k) => {
    t(s, `Число ${v}`, v, 96 + k * 380, 210, 340, 170, { size: 80, bold: true, color: C.accent1 });
    t(s, `Подпись ${v}`, l, 96 + k * 380, 380, 340, 40, { size: 22, color: C.accent4 });
  });
  t(s, 'Метаанализ', 'Метаанализ 2009 года, 253 выборки', 96, 430, 1100, 40, { size: 18, color: C.accent4 });
  t(s, 'Барнсли заголовок', 'Барнсли повторил исследование через 40 лет (2025):', 96, 530, 1000, 50, { size: 24, bold: true });
  s.addChart(pres.charts.BAR, [{ name: 'Лучшие юниоры, %', labels: ['I квартал (янв–мар)', 'IV квартал (окт–дек)'], values: [40, 10] }],
    chartBase({ x: X(96), y: X(590), w: X(1000), h: X(380), barDir: 'bar', catAxisOrientation: 'maxMin', chartColors: [P.BLUE, P.ORANGE], valAxisMaxVal: 50, valAxisHidden: true, valGridLine: { style: 'none' }, dataLabelFormatCode: '"≈ "0" %"', catAxisLabelFontSize: 16, objectName: nm('Барнсли диаграмма') }));
  box(s, 'Цитата фон', 1180, 520, 644, 330, P.ORANGE_L);
  t(s, 'Цитата', '«Life is not fair!»', 1220, 570, 580, 90, { size: 40, bold: true });
  t(s, 'Перевод', '«Жизнь несправедлива» — так Барнсли назвал свою статью', 1220, 680, 580, 130, { size: 22, color: C.accent4 });
  ['Барнсли заголовок', 'Барнсли диаграмма'].forEach((x) => anim(1, x));
  ['Цитата фон', 'Цитата', 'Перевод'].forEach((x) => anim(2, x));
  s.addNotes('Метаанализ 38 исследований из 14 видов спорта и 16 стран нашёл эффект почти везде.\n[КЛИК] А в 2025 году Барнсли сам повторил своё исследование через сорок лет. Результат почти тот же: около 40 % лучших юниоров родились в первом квартале и только 10 % в последнем.\n[КЛИК] Статью он назвал «Жизнь несправедлива».');
}

// 5 · Моё исследование
{
  const s = slide('Слайд', 'Моё исследование: 853 игрока НХЛ', 'Составы всех клубов НХЛ, сезон 2024/25 (api-web.nhle.com); критерий хи-квадрат');
  t(s, 'Метод', [R('Что сделал:', { bold: true, breakLine: true }), R('• взял составы всех клубов НХЛ 2024/25 — 853 игрока', { breakLine: true }), R('• разделил по кварталам рождения', { breakLine: true }), R('• проверил критерием хи-квадрат')], 96, 220, 560, 340, { size: 19, ls: 1.15 });
  box(s, 'Итог фон', 96, 590, 560, 400, P.BLUE_L);
  t(s, 'Итог число', '2,3 раза', 136, 610, 500, 130, { size: 60, bold: true, color: C.accent1 });
  t(s, 'Итог текст', 'больше канадцев из января–марта, чем из октября–декабря', 136, 750, 490, 130, { size: 19 });
  t(s, 'Итог p', 'p < 0,001 — гипотеза подтвердилась', 136, 900, 490, 60, { size: 20, bold: true });
  t(s, 'Диаграмма заголовок', 'Канадцы в НХЛ по кварталу рождения, % (342 игрока)', 720, 220, 1104, 40, { size: 20, bold: true });
  s.addChart(pres.charts.BAR, [{ name: 'Канада, %', labels: QLAB, values: [34.5, 28.1, 22.5, 14.9] }],
    chartBase({ x: X(720), y: X(270), w: X(1104), h: X(700), barDir: 'col', chartColors: [P.BLUE, P.LIGHT, P.LIGHT, P.BLUE], valAxisMaxVal: 40, dataLabelFormatCode: '0.0', objectName: nm('Канада диаграмма') }));
  ['Итог фон', 'Итог число', 'Итог текст'].forEach((x) => anim(1, x));
  anim(2, 'Итог p');
  s.addNotes('Я проверил это сам. Взял составы всех клубов НХЛ сезона 2024/25, это 853 игрока, распределил их по кварталам рождения и применил критерий хи-квадрат.\n[КЛИК] Посмотрите на синие столбики. Среди канадцев в январе–марте родились 35 % игроков, а в октябре–декабре всего 15 %. Разница в 2,3 раза.\n[КЛИК] Вероятность получить такое случайно меньше одной тысячной. Гипотеза подтвердилась: через сорок лет эффект никуда не делся.');
}

// 6 · Загадка
{
  const s = slide('Слайд', 'Загадка');
  t(s, 'Вопрос вступление', 'Игорь Шестёркин — один из лучших вратарей мира.', 96, 230, 1728, 60, { size: 28, color: C.accent4 });
  t(s, 'Вопрос', 'В каком месяце он родился?', 96, 300, 1728, 110, { size: 54, bold: true });
  months(s, 'Месяц', 96, 500, 1728, 110, () => P.LIGHT, { size: 22 });
  const cw = (1728 - 88) / 12;
  s.addText('Дек', { shape: pres.shapes.RECTANGLE, x: X(96 + 11 * (cw + 8)), y: X(500), w: X(cw), h: X(110), margin: 0, align: 'center', valign: 'middle', fontFace: 'Calibri', fontSize: 22, bold: true, color: 'FFFFFF', fill: { color: P.ORANGE }, line: { type: 'none' }, lang: LANG, objectName: nm('Ответ декабрь') });
  t(s, 'Ответ', [R('30 декабря', { color: C.accent2, bold: true }), R(' 1995 года', { color: C.text1 })], 96, 670, 1728, 120, { size: 54 });
  t(s, 'Ответ пояснение', 'Обладатель «Везина Трофи» 2022 — приза лучшему вратарю НХЛ', 96, 810, 1728, 50, { size: 22, color: C.accent4 });
  ['Ответ декабрь', 'Ответ', 'Ответ пояснение'].forEach((x) => anim(1, x));
  s.addNotes('А теперь вопрос залу. Это Игорь Шестёркин, один из лучших вратарей мира. Как думаете, в каком месяце он родился? (выслушать 2–3 ответа)\n[КЛИК] 30 декабря.\n\nМожно вставить фото: Вставка → Рисунки, справа от вопроса.');
}

// 7 · Россияне
{
  const s = slide('Слайд', [R('У российских игроков картина '), R('обратная', { color: C.accent2 })], 'Fumarco, Gibbs, Jarvis, Rossi // PLoS ONE, 2017');
  t(s, 'Диаграмма заголовок', 'Игроки НХЛ по кварталу рождения, %', 96, 210, 1100, 40, { size: 20, bold: true });
  s.addChart(pres.charts.BAR, [{ name: 'Россия (52 игрока)', labels: QLAB, values: [23.1, 21.2, 21.2, 34.6] }, { name: 'Канада (342 игрока)', labels: QLAB, values: [34.5, 28.1, 22.5, 14.9] }],
    chartBase({ x: X(96), y: X(260), w: X(1100), h: X(560), barDir: 'col', barGrouping: 'clustered', chartColors: [P.ORANGE, P.BLUE_L], valAxisMaxVal: 40, dataLabelFormatCode: '0.0', dataLabelFontSize: 12, showLegend: true, legendPos: 't', legendFontSize: 14, legendFontFace: '+mn-lt', legendColor: P.INK, objectName: nm('Россия диаграмма') }));
  t(s, 'Игроки', [R('Поздние звёзды:', { bold: true, breakLine: true }), R('Артемий Панарин — 30 октября', { breakLine: true }), R('(не выбран на драфте, сейчас звезда лиги)', { color: C.accent4, fontSize: 16, breakLine: true }), R('Матвей Мичков — 9 декабря', { breakLine: true }), R('Иван Демидов — 10 декабря')], 1260, 260, 564, 330, { size: 22, ls: 1.25 });
  box(s, 'Честно фон', 96, 850, 1728, 130, P.ORANGE_L);
  t(s, 'Честно', [R('Честно: ', { bold: true }), R('россиян всего 52 (p = 0,47) — это не доказано. Но совпадает с эффектом аутсайдера: в элите НХЛ поздние зарабатывают на 51 % больше.')], 136, 870, 1648, 90, { size: 20, valign: 'middle' });
  anim(1, 'Игроки');
  ['Честно фон', 'Честно'].forEach((x) => anim(2, x));
  s.addNotes('И это не исключение. Посмотрите на оранжевые столбики: у российских игроков картина обратная, больше всего родились в конце года.\n[КЛИК] Панарин родился 30 октября, на драфте его не выбрал никто, а сегодня он звезда лиги. Мичков и Демидов родились в декабре.\n[КЛИК] Честно: россиян всего 52, и статистически это не доказано. Но это совпадает с эффектом аутсайдера: среди самой элиты НХЛ поздних больше, и зарабатывают они больше. Кто всё детство был самым маленьким, учился играть умнее.');
}

// 8 · Не только спорт
{
  const s = slide('Слайд', 'Не только спорт: школа', 'Layton et al. // New England Journal of Medicine, 2018 (≈ 407 000 детей)');
  t(s, 'Пояснение года', 'В США в класс набирают по дате рождения, граница — 1 сентября:', 96, 220, 1728, 40, { size: 20, color: C.accent4 });
  months(s, 'Школа', 96, 280, 1728, 80, (m) => (m === 9 ? P.BLUE : m === 8 ? P.ORANGE : P.LIGHT), { start: 9 });
  t(s, 'Старшие', '← самые старшие', 96, 372, 600, 36, { size: 16, color: C.accent4 });
  t(s, 'Младшие', 'самые младшие →', 1224, 372, 600, 36, { size: 16, color: C.accent4, align: 'right' });
  t(s, '34', '+34 %', 96, 440, 760, 210, { size: 96, bold: true, color: C.accent2 });
  t(s, '34 текст', 'чаще ставят диагноз СДВГ самым младшим в классе', 96, 660, 700, 100, { size: 24 });
  t(s, '34 вывод', 'Обычную незрелость принимают за расстройство', 96, 800, 700, 100, { size: 24, bold: true });
  s.addChart(pres.charts.BAR, [{ name: 'Диагнозов на 10 000 детей', labels: ['Рождённые в августе', 'Рождённые в сентябре'], values: [85.1, 63.6] }],
    chartBase({ x: X(880), y: X(470), w: X(944), h: X(420), barDir: 'bar', catAxisOrientation: 'maxMin', chartColors: [P.ORANGE, P.BLUE], valAxisMaxVal: 100, valAxisHidden: true, valGridLine: { style: 'none' }, dataLabelFormatCode: '0.0', catAxisLabelFontSize: 16, objectName: nm('СДВГ диаграмма') }));
  t(s, 'Подпись диаграммы', 'Диагнозов СДВГ на 10 000 детей', 880, 900, 944, 40, { size: 16, color: C.accent4 });
  ['34', '34 текст', '34 вывод', 'СДВГ диаграмма', 'Подпись диаграммы'].forEach((x) => anim(1, x));
  s.addNotes('И это не только про спорт.\n[КЛИК] В США самым младшим в классе диагноз СДВГ ставят на 34 % чаще: обычную незрелость принимают за расстройство.');
}

// 9 · Что делать
{
  const s = slide('Слайд', 'Что можно сделать', 'Cumming et al., 2018; Mann, van Ginneken, 2017');
  const items = [['1', 'Делить по зрелости', 'В английских футбольных академиях детей делят по биологической зрелости, а не по году рождения'], ['2', 'Номер = возраст', 'Если скаут видит возраст по номеру на футболке, перекос в отборе почти исчезает'], ['3', 'Сравнивать с собой', 'Оценивать, как ребёнок вырос сам, а не какое место он занял среди ровесников']];
  items.forEach(([num, h, d], k) => {
    const x = 96 + k * 584;
    box(s, `Карточка ${num}`, x, 240, 560, 640, k === 2 ? P.SOFT : P.BLUE_L);
    t(s, `Номер ${num}`, num, x + 40, 260, 200, 160, { size: 72, bold: true, color: C.accent1 });
    t(s, `Заголовок ${num}`, h, x + 40, 430, 480, 100, { size: 28, bold: true });
    t(s, `Текст ${num}`, d, x + 40, 540, 480, 300, { size: 22, color: C.accent4, ls: 1.15 });
    ['Карточка', 'Номер', 'Заголовок', 'Текст'].forEach((p) => anim(k + 1, `${p} ${num}`));
  });
  s.addNotes('Решения есть.\n[КЛИК] В английских футбольных академиях детей делят по биологической зрелости, а не по году рождения.\n[КЛИК] А если дать скауту номера на футболках по возрасту, перекос в отборе почти исчезает.\n[КЛИК] (по желанию) И ещё — сравнивать ребёнка с ним самим, а не с ровесниками.');
}

// 10 · Выводы
{
  const s = slide('Слайд', 'Выводы');
  const items = [['Эффект реален и жив в НХЛ до сих пор', P.BLUE], ['Причина — не талант, а система отбора и ожидания взрослых', P.BLUE], ['Его можно ослабить', P.ORANGE]];
  items.forEach(([txt, col], k) => {
    const y = 220 + k * 110;
    s.addShape(pres.shapes.OVAL, { x: X(96), y: X(y + 6), w: X(64), h: X(64), fill: { color: col }, line: { type: 'none' }, objectName: nm(`Круг ${k + 1}`) });
    t(s, `Номер ${k + 1}`, String(k + 1), 96, y + 6, 64, 64, { size: 26, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    t(s, `Вывод ${k + 1}`, txt, 196, y + 8, 1628, 64, { size: 26, valign: 'middle' });
    [`Круг ${k + 1}`, `Номер ${k + 1}`, `Вывод ${k + 1}`].forEach((x) => anim(k + 1, x));
  });
  months(s, 'Финал месяц', 96, 600, 1728, 70, () => P.BLUE);
  t(s, 'Финал', [R('Календарь не определяет талант.', { bold: true, breakLine: true }), R('Он определяет только, кого мы успеваем заметить.', { color: C.accent4 })], 96, 700, 1728, 130, { size: 32 });
  t(s, 'Спасибо', 'Спасибо за внимание! Готов ответить на вопросы.', 96, 880, 1728, 60, { size: 26, bold: true, color: C.accent1 });
  [...M.map((m) => `Финал месяц ${m}`), 'Финал', 'Спасибо'].forEach((x) => anim(4, x));
  s.addNotes('Итак:\n[КЛИК] эффект реален и жив в НХЛ до сих пор.\n[КЛИК] Его причина не в таланте, а в системе отбора и ожиданиях взрослых.\n[КЛИК] И его можно ослабить.\n[КЛИК] Хоккеисты рождаются не в январе. Хоккеисты рождаются в любой месяц, просто январским чаще дают шанс. Календарь не определяет талант. Он определяет только, кого мы успеваем заметить. Спасибо! Готов ответить на вопросы.');
}

// 11 · Данные (для вопросов)
{
  const s = slide('Слайд', 'Приложение: данные по странам (НХЛ 2024/25)', 'Ожидание по числу дней в квартале; критерий хи-квадрат, 3 степени свободы');
  const H = (x) => ({ text: x, options: { bold: true, color: 'FFFFFF', fill: { color: P.GRAY }, align: 'center', lang: LANG } });
  const rows = [[H('Группа'), H('Игроков'), H('I кв., %'), H('II кв., %'), H('III кв., %'), H('IV кв., %'), H('I / IV'), H('p')]];
  [['Вся лига', 853, '33,3', '25,1', '21,9', '19,7', '1,69', '< 0,001'], ['Канада', 342, '34,5', '28,1', '22,5', '14,9', '2,31', '< 0,001'], ['США', 259, '35,1', '22,4', '20,1', '22,4', '1,57', '0,001'], ['Швеция', 80, '26,2', '20,0', '27,5', '26,2', '1,00', '0,79'], ['Россия', 52, '23,1', '21,2', '21,2', '34,6', '0,67', '0,47']]
    .forEach((r, k) => rows.push(r.map((v, j) => ({ text: String(v), options: { align: j === 0 ? 'left' : 'center', bold: j === 0, lang: LANG, fill: { color: k % 2 ? P.SOFT : P.BG } } }))));
  s.addTable(rows, { x: X(96), y: X(220), w: X(1728), colW: [300, 200, 200, 200, 200, 200, 200, 228].map(X), fontFace: 'Calibri', fontSize: 20, color: P.INK, rowH: X(80), border: { type: 'solid', pt: 0.75, color: P.LIGHT }, margin: [0.05, 0.1, 0.05, 0.1], objectName: nm('Таблица') });
  t(s, 'Пояснение', 'Канада и США — значимый перекос в пользу ранних. Швеция — эффекта нет. Россия — обратная картина, но выборка мала.', 96, 760, 1728, 80, { size: 20, color: C.accent4 });
  s.addNotes('Резерв для вопросов.\n1. Больше ли рождается детей в начале года? — Нет, рождаемость почти равномерная, летом даже немного больше; с поправкой эффект сильнее.\n2. Почему нет эффекта у шведов? — Там жёсткий отбор начинается позже («спорт для всех»); но выборка 80 человек — это предположение.\n3. Почему россияне не доказаны? — 52 человека мало, p = 0,47. Дальше — проверить КХЛ.\n4. Планировать ли рождение детей? — Нет, правильнее менять систему отбора.\n5. При чём психология? — Физиология даёт стартовую разницу, психология (Пигмалион, эффект Матфея, мотивация) превращает её в разницу на всю карьеру.\n6. Почему НХЛ, а не КХЛ? — Открытые полные данные и много стран.');
}

// 12 · Источники
{
  const s = slide('Слайд', 'Источники');
  const refs = ['Barnsley R. H., Thompson A. H., Barnsley P. E. Hockey success and birthdate: The relative age effect // CAHPER Journal. 1985.', 'Barnsley R. H. 40-years of relative age effects: life is not fair! // Frontiers in Sports and Active Living. 2025.', 'Gladwell M. Outliers: The Story of Success. 2008 («Гении и аутсайдеры»).', 'Cobley S. et al. Annual age-grouping and athlete development: a meta-analytical review // Sports Medicine. 2009.', 'Deaner R. O. et al. Born at the wrong time: selection bias in the NHL draft // PLoS ONE. 2013.', 'Gibbs B. G. et al. The rise of the underdog? // International Review for the Sociology of Sport. 2012.', 'Fumarco L. et al. The relative age effect reversal among the NHL elite // PLoS ONE. 2017.', 'Lemez S. et al. Relative age, competition level, and dropout rates in youth ice-hockey // Scand. J. Med. Sci. Sports. 2014.', 'Bedard K., Dhuey E. The persistence of early childhood maturity // Quarterly Journal of Economics. 2006.', 'Morrow R. L. et al. Influence of relative age on diagnosis and treatment of ADHD // CMAJ. 2012.', 'Layton T. J. et al. ADHD and month of school enrollment // NEJM. 2018.', 'Mann D. L., van Ginneken P. Age-ordered shirt numbering reduces the selection bias // Journal of Sports Sciences. 2017.', 'Cumming S. P. et al. Bio-banded tournament in Premier League academies // Journal of Sports Sciences. 2018.', 'Rosenthal R., Jacobson L. Pygmalion in the Classroom. 1968.', 'Merton R. K. The Matthew Effect in Science // Science. 1968.', 'Данные: составы клубов НХЛ 2024/25 (api-web.nhle.com).'];
  const runs = refs.map((r, i) => R(`${i + 1}. ${r}`, { breakLine: i < refs.length - 1 }));
  t(s, 'Список', runs, 96, 200, 1728, 820, { size: 14, color: C.text1, psa: 4 });
}

(async () => {
  const file = path.join(__dirname, 'out', 'deck-simple.pptx');
  fs.mkdirSync(path.dirname(file), { recursive: true });
  await pres.writeFile({ fileName: file });
  await applyTheme(file, THEME);
  fs.writeFileSync(path.join(__dirname, 'out', 'anim-simple.json'), JSON.stringify({ anim: ANIM, trans: TRANS }));
  console.log('OK', file, 'слайдов:', n);
})();
