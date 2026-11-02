// Простая версия: мягкий фон, скруглённые полупрозрачные карточки, Arial, стандартные фигуры и диаграммы.
// Принцип надёжности: каждая карточка — ОДНА фигура с текстом внутри (ничего не лежит слоями друг на друге).
//   node build_simple.js  →  out/deck-simple.pptx + out/anim-simple.json  →  python3 postprocess.py simple
const path = require('path');
const fs = require('fs');
const pptxgen = require('pptxgenjs');
const SKILL = '/root/.claude/skills/synced/2d64efaf-fc81-4729-a14f-24728a20f0e8_08b58b4f-6847-4399-ba9d-bded480b3897/pptx';
const { applyTheme } = require(SKILL + '/scripts/apply_theme.js');

// синий — «ранние», доверие и спокойствие; оранжевый — «поздние», внимание; остальное нейтральное
const P = { INK: '1B2A3A', GRAY: '56657A', LINE: 'D5DEE9', CHIP: 'E3E9F1', BLUE: '2A78D6', BLUE_L: 'DCE9F8', ORANGE: 'E2622F', ORANGE_L: 'FCE6DC', WHITE: 'FFFFFF' };
const THEME = {
  name: 'Простая', headFontFace: 'Arial', bodyFontFace: 'Arial',
  colors: { dk1: P.INK, lt1: P.WHITE, dk2: P.GRAY, lt2: 'F2F5F9', accent1: P.BLUE, accent2: P.ORANGE, accent3: P.BLUE_L, accent4: P.GRAY, accent5: P.CHIP, accent6: P.ORANGE_L, hlink: P.BLUE, folHlink: P.GRAY },
};
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.theme = { headFontFace: 'Arial', bodyFontFace: 'Arial' };
pres.title = 'Эффект относительного возраста';
pres.author = 'Имя Фамилия';
const C = pres.SchemeColor;
const X = (px) => Math.round((px / 144) * 10000) / 10000;
const LANG = 'ru-RU';
const R = (t, o = {}) => ({ text: t, options: { lang: LANG, ...o } });
const BR = (t, o = {}) => R(t, { ...o, breakLine: true });
const M = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];
const BG = path.join(__dirname, 'assets', 'simple', 'bg.jpg');

const ANIM = {}, TRANS = {};
let n = 0;
const anim = (step, name) => { const a = (ANIM[n] = ANIM[n] || []); (a[step - 1] = a[step - 1] || []).push({ name, fx: 'fade', delay: 0 }); };

pres.defineSlideMaster({ title: 'Титул', background: { path: BG }, objects: [] });
pres.defineSlideMaster({
  title: 'Слайд',
  background: { path: BG },
  objects: [
    { placeholder: { options: { name: 'title', type: 'title', x: X(96), y: X(60), w: X(1728), h: X(110), fontFace: 'Arial', fontSize: 34, bold: true, color: C.text1, margin: 0, valign: 'middle', align: 'left', lang: LANG }, text: '' } },
    { placeholder: { options: { name: 'source', type: 'body', x: X(96), y: X(1012), w: X(1560), h: X(40), fontFace: 'Arial', fontSize: 11, color: C.accent4, margin: 0, valign: 'middle', align: 'left', lang: LANG }, text: '' } },
  ],
  slideNumber: { x: X(1704), y: X(1000), w: X(120), h: X(64), fontFace: 'Arial', fontSize: 12, color: C.accent4, align: 'right' },
});

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
const shadow = () => ({ type: 'outer', color: '1B2A3A', opacity: 0.12, blur: 12, offset: 4, angle: 90 });

// карточка: скруглённая, полупрозрачная, текст внутри
function card(s, name, x, y, w, h, runs, o = {}) {
  s.addText(runs, {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: X(o.r ?? 26), x: X(x), y: X(y), w: X(w), h: X(h),
    fill: { color: o.fill ?? P.WHITE, transparency: o.t ?? 22 }, line: o.line ?? { color: P.WHITE, width: 1, transparency: 10 },
    shadow: o.noShadow ? undefined : shadow(), margin: o.pad ?? 16, fontFace: 'Arial', fontSize: o.size ?? 20, color: o.color ?? C.text1,
    bold: !!o.bold, align: o.align ?? 'left', valign: o.valign ?? 'top', lang: LANG, lineSpacingMultiple: o.ls ?? 1.05,
    paraSpaceAfter: o.psa ?? 0, objectName: nm(name),
  });
}
function text(s, name, runs, x, y, w, h, o = {}) {
  s.addText(typeof runs === 'string' ? [R(runs)] : runs, {
    x: X(x), y: X(y), w: X(w), h: X(h), margin: 0, isTextBox: true, fontFace: 'Arial', fontSize: o.size ?? 20, bold: !!o.bold,
    color: o.color ?? C.text1, align: o.align ?? 'left', valign: o.valign ?? 'top', lang: LANG, objectName: nm(name),
  });
}
function chips(s, prefix, x0, y, w, h, color, opts = {}) {
  const gap = 10, cw = (w - gap * 11) / 12, start = opts.start ?? 1;
  for (let k = 0; k < 12; k++) {
    const m = ((start - 1 + k) % 12) + 1;
    const c = color(m);
    card(s, `${prefix} ${M[m - 1]}`, x0 + k * (cw + gap), y, cw, h, [R(M[m - 1])], {
      r: 16, fill: c, t: c === P.CHIP ? 15 : 0, line: { type: 'none' }, noShadow: c === P.CHIP, pad: 2,
      size: opts.size ?? 16, bold: true, color: c === P.CHIP ? P.GRAY : P.WHITE, align: 'center', valign: 'middle',
    });
  }
  return { xOf: (m) => x0 + ((m - start + 12) % 12) * (cw + gap), cw };
}
const QLAB = ['I кв.\nянв–мар', 'II кв.\nапр–июн', 'III кв.\nиюл–сен', 'IV кв.\nокт–дек'];
const chartBase = (o) => Object.assign({
  showLegend: false, showTitle: false, barGapWidthPct: 60, valAxisMinVal: 0, valAxisMajorUnit: 10, valAxisLabelFormatCode: '0',
  valAxisLabelColor: P.GRAY, valAxisLabelFontSize: 12, valAxisLabelFontFace: 'Arial', valAxisLineShow: false,
  valGridLine: { color: P.LINE, size: 0.75 }, catGridLine: { style: 'none' }, catAxisLabelColor: P.INK, catAxisLabelFontSize: 14,
  catAxisLabelFontFace: 'Arial', showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 14, dataLabelFontBold: true,
  dataLabelFontFace: 'Arial', dataLabelColor: P.INK,
  plotArea: { fill: { color: 'FFFFFF', transparency: 100 } }, chartArea: { fill: { color: 'FFFFFF', transparency: 100 } },
}, o);

// 1 · Титул
{
  const s = slide('Титул');
  text(s, 'Конкурс', 'МАРАФОН ЛЕКТОРИЕВ «ПОТОК»', 96, 140, 1200, 40, { size: 14, color: C.accent4, bold: true });
  text(s, 'Название', 'Эффект относительного возраста', 96, 186, 1728, 124, { size: 48, bold: true });
  text(s, 'Подзаголовок', [R('Почему хоккеисты '), R('рождаются в январе', { color: C.accent1, bold: true })], 96, 316, 1728, 76, { size: 28, color: C.accent4 });
  card(s, 'Автор', 96, 440, 900, 210, [BR('Имя Фамилия', { fontSize: 26, bold: true }), R('Образовательное направление', { fontSize: 20, color: P.GRAY })], { pad: 20, valign: 'middle' });
  chips(s, 'Месяц', 96, 800, 1728, 96, (m) => (m <= 3 ? P.BLUE : m >= 10 ? P.ORANGE : P.CHIP));
  s.addNotes(['Добрый день! Давайте начнём с эксперимента.', 'Поднимите руку те, кто родился в январе, феврале или марте. (пауза, оглядеть зал)', 'А теперь те, кто родился в октябре, ноябре или декабре. (пауза)', 'Будь мы канадскими хоккеистами, у первой группы шансов попасть в НХЛ было бы больше чем в два раза. И дело не в генах, а в календаре.'].join('\n'));
}

// 2 · 1984
{
  const s = slide('Слайд', '1984: как открыли эффект', 'Barnsley, Thompson, Barnsley // CAHPER Journal, 1985');
  card(s, 'История', 96, 200, 980, 790, [
    BR('Психолог Роджер Барнсли с женой Паулой пришли на юниорский хоккейный матч в Канаде. Паула листала программку матча и заметила:', { color: P.GRAY }),
    BR(' ', { fontSize: 14 }),
    BR('«Смотри, они почти все родились в январе, феврале и марте»', { fontSize: 28, bold: true }),
    BR('— Паула Барнсли', { fontSize: 16, color: P.GRAY }),
    BR(' ', { fontSize: 14 }),
    R('Так открыли '), R('эффект относительного возраста', { bold: true }), R(': в одной возрастной группе выигрывают те, кто старше.'),
  ], { pad: 20, size: 19, ls: 1.1 });
  card(s, 'Цель', 1116, 200, 708, 290, [BR('ЦЕЛЬ', { fontSize: 16, bold: true, color: P.BLUE }), R('Проверить, жив ли этот эффект сегодня', { fontSize: 26, bold: true })], { pad: 22, fill: P.BLUE_L, t: 10 });
  card(s, 'Гипотеза', 1116, 520, 708, 450, [BR('ГИПОТЕЗА', { fontSize: 16, bold: true, color: P.BLUE }), R('Среди игроков НХЛ родившихся в начале года значимо больше, чем родившихся в конце', { fontSize: 24, bold: true })], { pad: 22, fill: P.BLUE_L, t: 10 });
  anim(1, 'Цель'); anim(1, 'Гипотеза');
  s.addNotes('В 1984 году канадский психолог Роджер Барнсли с женой Паулой пришёл на юниорский хоккейный матч. Паула листала программку и вдруг сказала: «Смотри, они почти все родились в январе, феврале и марте». Барнсли проверил профессионалов: та же картина. Так открыли эффект относительного возраста.\n[КЛИК] Моя цель — проверить, жив ли этот эффект сегодня. Гипотеза: среди игроков НХЛ родившихся в начале года значимо больше, чем родившихся в конце.');
}

// 3 · Снежный ком
{
  const s = slide('Слайд', 'Почему так происходит: снежный ком', 'Rosenthal, Jacobson, 1968; Merton, 1968; Lemez et al., 2014');
  card(s, '11 процентов', 96, 190, 1728, 210, [R('11 %   ', { fontSize: 44, bold: true, color: P.ORANGE }), R('364 дня разницы — больше десятой части жизни девятилетнего', { fontSize: 22 })], { pad: 22, valign: 'middle' });
  const steps = [['1. Отбор', 'Тренер берёт того, кто сильнее сейчас'], ['2. Ресурсы', 'Отобранным — лучшие тренеры и больше льда'], ['3. Ожидания', 'Эффект Пигмалиона: в кого верят, тот растёт'], ['4. Уход', 'Младшие сидят на скамейке и бросают спорт']];
  steps.forEach(([h, d], k) => {
    const x = 96 + k * 446;
    card(s, `Шаг ${k + 1}`, x, 430, 390, 340, [BR(h, { fontSize: 24, bold: true }), BR(' ', { fontSize: 8 }), R(d, { fontSize: 18 })],
      { pad: 16, fill: k === 3 ? P.ORANGE : P.BLUE, t: 8, color: P.WHITE, line: { type: 'none' } });
    if (k < 3) s.addShape(pres.shapes.RIGHT_ARROW, { x: X(x + 398), y: X(580), w: X(40), h: X(40), fill: { color: P.GRAY }, line: { type: 'none' }, objectName: nm(`Стрелка ${k + 1}`) });
    anim(k + 1, `Шаг ${k + 1}`);
    if (k > 0) anim(k + 1, `Стрелка ${k}`);
  });
  card(s, 'Пророчество', 96, 800, 1728, 180, [R('Самоисполняющееся пророчество: ', { bold: true, color: P.BLUE }), R('решили, что январские сильнее, — и сами сделали их сильнее.')], { pad: 16, size: 22, valign: 'middle' });
  anim(5, 'Пророчество');
  s.addNotes('В детском хоккее команды собирают по году рождения. Мальчик, родившийся 1 января, и мальчик, родившийся 31 декабря, играют в одной команде, хотя между ними почти год. Для девятилетнего это больше десятой части жизни.\n[КЛИК] Дальше всё катится как снежный ком. Тренер берёт того, кто сильнее сейчас, и путает зрелость с талантом.\n[КЛИК] Отобранные получают лучших тренеров и больше льда.\n[КЛИК] Работает эффект Пигмалиона: тот, в кого верят, растёт быстрее.\n[КЛИК] А младшие сидят на скамейке и бросают спорт.\n[КЛИК] Получается самоисполняющееся пророчество: мы решили, что январские сильнее, и сами сделали их сильнее.');
}

// 4 · Масштаб
{
  const s = slide('Слайд', 'Эффект находят почти везде', 'Метаанализ: Cobley et al. // Sports Medicine, 2009 (253 выборки); Barnsley // Frontiers in Sports and Active Living, 2025');
  [['38', 'исследований'], ['14', 'видов спорта'], ['16', 'стран']].forEach(([v, l], k) => {
    card(s, `Число ${v}`, 96 + k * 588, 190, 552, 280, [BR(v, { fontSize: 54, bold: true, color: P.BLUE }), R(l, { fontSize: 22, color: P.GRAY })], { pad: 16, valign: 'middle' });
  });
  card(s, 'Барнсли', 96, 510, 1100, 470, [R('Барнсли повторил опыт через 40 лет (2025)', { bold: true })], { pad: 22, size: 22 });
  s.addChart(pres.charts.BAR, [{ name: 'Лучшие юниоры, %', labels: ['I квартал (янв–мар)', 'IV квартал (окт–дек)'], values: [40, 10] }],
    chartBase({ x: X(130), y: X(610), w: X(1030), h: X(350), barDir: 'bar', catAxisOrientation: 'maxMin', chartColors: [P.BLUE, P.ORANGE], valAxisMaxVal: 50, valAxisHidden: true, valGridLine: { style: 'none' }, dataLabelFormatCode: '"≈ "0" %"', catAxisLabelFontSize: 16, dataLabelFontSize: 18, objectName: nm('Барнсли диаграмма') }));
  card(s, 'Цитата', 1236, 510, 588, 470, [BR('«Life is not fair!»', { fontSize: 32, bold: true }), BR(' ', { fontSize: 10 }), R('«Жизнь несправедлива» — так Барнсли назвал свою статью', { fontSize: 20, color: P.GRAY })], { pad: 20, fill: P.ORANGE_L, t: 10, valign: 'middle' });
  anim(1, 'Барнсли'); anim(1, 'Барнсли диаграмма'); anim(2, 'Цитата');
  s.addNotes('Метаанализ 38 исследований из 14 видов спорта и 16 стран нашёл эффект почти везде.\n[КЛИК] А в 2025 году Барнсли сам повторил своё исследование через сорок лет. Результат почти тот же: около 40 % лучших юниоров родились в первом квартале и только 10 % в последнем.\n[КЛИК] Статью он назвал «Жизнь несправедлива».');
}

// 5 · Моё исследование
{
  const s = slide('Слайд', 'Моё исследование: 853 игрока НХЛ', 'Составы всех клубов НХЛ, сезон 2024/25 (api-web.nhle.com); критерий хи-квадрат');
  card(s, 'Метод', 96, 190, 640, 420, [BR('Что я сделал', { bold: true, fontSize: 22 }),
    BR('• взял составы всех клубов НХЛ 2024/25 — 853 игрока'), BR('• разделил игроков по кварталам рождения'), R('• проверил результат критерием хи-квадрат')], { pad: 18, size: 18, ls: 1.1 });
  card(s, 'Итог', 96, 630, 640, 360, [BR('в 2,3 раза', { fontSize: 36, bold: true, color: P.BLUE }), BR('больше канадцев, рождённых в январе–марте', { fontSize: 18 }), R('p < 0,001 — гипотеза подтвердилась', { fontSize: 18, bold: true })], { pad: 18, fill: P.BLUE_L, t: 10 });
  card(s, 'Диаграмма карточка', 776, 200, 1048, 770, [R('Канадцы в НХЛ по кварталу рождения, %', { bold: true })], { pad: 22, size: 19 });
  s.addChart(pres.charts.BAR, [{ name: 'Канада, %', labels: QLAB, values: [34.5, 28.1, 22.5, 14.9] }],
    chartBase({ x: X(806), y: X(290), w: X(990), h: X(670), barDir: 'col', chartColors: [P.BLUE, P.CHIP, P.CHIP, P.BLUE], valAxisMaxVal: 40, dataLabelFormatCode: '0.0', objectName: nm('Канада диаграмма') }));
  anim(1, 'Итог');
  s.addNotes('Я проверил это сам. Взял составы всех клубов НХЛ сезона 2024/25, это 853 игрока, распределил их по кварталам рождения и применил критерий хи-квадрат. Посмотрите на синие столбики. Среди канадцев в январе–марте родились 35 % игроков, а в октябре–декабре всего 15 %.\n[КЛИК] Разница в 2,3 раза. Вероятность получить такое случайно меньше одной тысячной. Гипотеза подтвердилась: через сорок лет эффект никуда не делся.');
}

// 6 · Загадка
{
  const s = slide('Слайд', 'Загадка');
  card(s, 'Вопрос', 96, 190, 1728, 280, [BR('Игорь Шестёркин — один из лучших вратарей мира.', { fontSize: 24, color: P.GRAY }), R('В каком месяце он родился?', { fontSize: 46, bold: true })], { pad: 20, valign: 'middle' });
  const g = chips(s, 'Месяц', 96, 500, 1728, 110, () => P.CHIP, { size: 22 });
  card(s, 'Ответ декабрь', g.xOf(12), 500, g.cw, 110, [R('Дек')], { r: 16, fill: P.ORANGE, t: 0, line: { type: 'none' }, pad: 2, size: 22, bold: true, color: P.WHITE, align: 'center', valign: 'middle' });
  card(s, 'Ответ', 96, 650, 1728, 280, [BR('30 декабря 1995 года', { fontSize: 46, bold: true, color: P.ORANGE }), R('Обладатель «Везина Трофи» 2022 — приза лучшему вратарю НХЛ', { fontSize: 22, color: P.GRAY })], { pad: 20, fill: P.ORANGE_L, t: 10, valign: 'middle' });
  anim(1, 'Ответ декабрь'); anim(1, 'Ответ');
  s.addNotes('А теперь вопрос залу. Это Игорь Шестёркин, один из лучших вратарей мира. Как думаете, в каком месяце он родился? (выслушать 2–3 ответа)\n[КЛИК] 30 декабря.');
}

// 7 · Россияне
{
  const s = slide('Слайд', [R('У российских игроков картина '), R('обратная', { color: C.accent2 })], 'Fumarco, Gibbs, Jarvis, Rossi // PLoS ONE, 2017');
  card(s, 'Диаграмма карточка', 96, 190, 1100, 630, [R('Игроки НХЛ по кварталу рождения, %', { bold: true })], { pad: 22, size: 19 });
  s.addChart(pres.charts.BAR, [{ name: 'Россия (52 игрока)', labels: QLAB, values: [23.1, 21.2, 21.2, 34.6] }, { name: 'Канада (342 игрока)', labels: QLAB, values: [34.5, 28.1, 22.5, 14.9] }],
    chartBase({ x: X(120), y: X(290), w: X(1050), h: X(510), barDir: 'col', barGrouping: 'clustered', chartColors: [P.ORANGE, P.BLUE_L], valAxisMaxVal: 40, dataLabelFormatCode: '0.0', dataLabelFontSize: 12, showLegend: true, legendPos: 't', legendFontSize: 14, legendFontFace: 'Arial', legendColor: P.INK, objectName: nm('Россия диаграмма') }));
  card(s, 'Игроки', 1236, 190, 588, 630, [BR('Поздние звёзды', { bold: true, fontSize: 22 }), BR(' ', { fontSize: 8 }),
    BR('Артемий Панарин — 30 октября'), BR('не выбран на драфте, сейчас звезда лиги', { fontSize: 16, color: P.GRAY }), BR(' ', { fontSize: 8 }),
    BR('Матвей Мичков — 9 декабря'), BR(' ', { fontSize: 8 }), R('Иван Демидов — 10 декабря')], { pad: 20, size: 20, ls: 1.05 });
  card(s, 'Честно', 96, 842, 1728, 150, [R('Честно: ', { bold: true }), R('россиян всего 52 (p = 0,47) — это не доказано. Но совпадает с эффектом аутсайдера: в элите НХЛ поздние зарабатывают на 51 % больше.')], { pad: 12, size: 19, fill: P.ORANGE_L, t: 10, valign: 'middle' });
  anim(1, 'Игроки'); anim(2, 'Честно');
  s.addNotes('И это не исключение. Посмотрите на оранжевые столбики: у российских игроков картина обратная, больше всего родились в конце года.\n[КЛИК] Панарин родился 30 октября, на драфте его не выбрал никто, а сегодня он звезда лиги. Мичков и Демидов родились в декабре.\n[КЛИК] Честно: россиян всего 52, и статистически это не доказано. Но это совпадает с эффектом аутсайдера: среди самой элиты НХЛ поздних больше, и зарабатывают они больше. Кто всё детство был самым маленьким, учился играть умнее.');
}

// 8 · Не только спорт
{
  const s = slide('Слайд', 'Не только спорт: школа', 'Layton et al. // New England Journal of Medicine, 2018 (≈ 407 000 детей)');
  text(s, 'Пояснение года', 'В США в класс набирают по дате рождения, граница — 1 сентября', 96, 190, 1728, 52, { size: 18, color: C.accent4 });
  chips(s, 'Школа', 96, 250, 1728, 86, (m) => (m === 9 ? P.BLUE : m === 8 ? P.ORANGE : P.CHIP), { start: 9 });
  text(s, 'Старшие', '← самые старшие', 96, 344, 600, 46, { size: 16, color: C.accent4 });
  text(s, 'Младшие', 'самые младшие →', 1224, 344, 600, 46, { size: 16, color: C.accent4, align: 'right' });
  card(s, '34', 96, 410, 760, 560, [BR('+34 %', { fontSize: 72, bold: true, color: P.ORANGE }), BR('чаще ставят диагноз СДВГ самым младшим в классе', { fontSize: 22 }), BR(' ', { fontSize: 10 }), R('Обычную незрелость принимают за расстройство', { fontSize: 22, bold: true })], { pad: 24, valign: 'middle' });
  card(s, 'Диаграмма карточка', 896, 410, 928, 560, [R('Диагнозов СДВГ на 10 000 детей', { bold: true })], { pad: 22, size: 19 });
  s.addChart(pres.charts.BAR, [{ name: 'Диагнозов на 10 000 детей', labels: ['Рождённые в августе', 'Рождённые в сентябре'], values: [85.1, 63.6] }],
    chartBase({ x: X(920), y: X(480), w: X(880), h: X(460), barDir: 'bar', catAxisOrientation: 'maxMin', chartColors: [P.ORANGE, P.BLUE], valAxisMaxVal: 100, valAxisHidden: true, valGridLine: { style: 'none' }, dataLabelFormatCode: '0.0', catAxisLabelFontSize: 16, dataLabelFontSize: 18, objectName: nm('СДВГ диаграмма') }));
  anim(1, '34'); anim(1, 'Диаграмма карточка'); anim(1, 'СДВГ диаграмма');
  s.addNotes('И это не только про спорт.\n[КЛИК] В США самым младшим в классе диагноз СДВГ ставят на 34 % чаще: обычную незрелость принимают за расстройство.');
}

// 9 · Что делать
{
  const s = slide('Слайд', 'Что можно сделать', 'Cumming et al. // Journal of Sports Sciences, 2018; Mann, van Ginneken // Journal of Sports Sciences, 2017');
  const items = [['1', 'Делить по зрелости', 'В английских футбольных академиях детей делят по биологической зрелости, а не по году рождения'], ['2', 'Номер = возраст', 'Если скаут видит возраст по номеру на футболке, перекос в отборе почти исчезает'], ['3', 'Сравнивать с собой', 'Оценивать, как ребёнок вырос сам, а не какое место он занял среди ровесников']];
  items.forEach(([num, h, d], k) => {
    card(s, `Решение ${num}`, 96 + k * 584, 200, 560, 720, [BR(num, { fontSize: 54, bold: true, color: P.BLUE }), BR(h, { fontSize: 26, bold: true }), BR(' ', { fontSize: 10 }), R(d, { fontSize: 20, color: P.GRAY })], { pad: 24, ls: 1.1, fill: k === 2 ? P.WHITE : P.BLUE_L, t: k === 2 ? 22 : 10 });
    anim(k + 1, `Решение ${num}`);
  });
  s.addNotes('Решения есть.\n[КЛИК] В английских футбольных академиях детей делят по биологической зрелости, а не по году рождения.\n[КЛИК] А если дать скауту номера на футболках по возрасту, перекос в отборе почти исчезает.\n[КЛИК] (по желанию) И ещё — сравнивать ребёнка с ним самим, а не с ровесниками.');
}

// 10 · Выводы
{
  const s = slide('Слайд', 'Выводы');
  const items = [['1', 'Эффект реален и жив в НХЛ до сих пор', P.BLUE], ['2', 'Причина — не талант, а система отбора и ожидания взрослых', P.BLUE], ['3', 'Его можно ослабить', P.ORANGE]];
  items.forEach(([num, txt, col], k) => {
    card(s, `Вывод ${num}`, 96 + k * 584, 190, 560, 340, [BR(num, { fontSize: 36, bold: true, color: col }), R(txt, { fontSize: 20, bold: true })], { pad: 18 });
    anim(k + 1, `Вывод ${num}`);
  });
  chips(s, 'Финал месяц', 96, 562, 1728, 70, () => P.BLUE);
  card(s, 'Финал', 96, 660, 1728, 230, [BR('Календарь не определяет талант.', { fontSize: 30, bold: true }), R('Он определяет только, кого мы успеваем заметить.', { fontSize: 24, color: P.GRAY })], { pad: 18, valign: 'middle' });
  text(s, 'Спасибо', 'Спасибо за внимание! Готов ответить на вопросы.', 96, 908, 1728, 64, { size: 24, bold: true, color: C.accent1 });
  [...M.map((m) => `Финал месяц ${m}`), 'Финал', 'Спасибо'].forEach((x) => anim(4, x));
  s.addNotes('Итак:\n[КЛИК] эффект реален и жив в НХЛ до сих пор.\n[КЛИК] Его причина не в таланте, а в системе отбора и ожиданиях взрослых.\n[КЛИК] И его можно ослабить.\n[КЛИК] Хоккеисты рождаются не в январе. Хоккеисты рождаются в любой месяц, просто январским чаще дают шанс. Календарь не определяет талант. Он определяет только, кого мы успеваем заметить. Спасибо! Готов ответить на вопросы.');
}

// 11 · Данные (для вопросов)
{
  const s = slide('Слайд', 'Приложение: данные по странам (НХЛ 2024/25)', 'Ожидание — по числу дней в квартале; критерий хи-квадрат, 3 степени свободы');
  const H = (x) => ({ text: x, options: { bold: true, color: P.WHITE, fill: { color: P.GRAY }, align: 'center', valign: 'middle', lang: LANG } });
  const rows = [[H('Группа'), H('Игроков'), H('I кв., %'), H('II кв., %'), H('III кв., %'), H('IV кв., %'), H('I / IV'), H('p')]];
  [['Вся лига', 853, '33,3', '25,1', '21,9', '19,7', '1,69', '< 0,001'], ['Канада', 342, '34,5', '28,1', '22,5', '14,9', '2,31', '< 0,001'], ['США', 259, '35,1', '22,4', '20,1', '22,4', '1,57', '0,001'], ['Швеция', 80, '26,2', '20,0', '27,5', '26,2', '1,00', '0,79'], ['Россия', 52, '23,1', '21,2', '21,2', '34,6', '0,67', '0,47']]
    .forEach((r, k) => rows.push(r.map((v, j) => ({ text: String(v), options: { align: j === 0 ? 'left' : 'center', valign: 'middle', bold: j === 0, lang: LANG, fill: { color: k % 2 ? 'F2F5F9' : P.WHITE } } }))));
  s.addTable(rows, { x: X(96), y: X(210), w: X(1728), colW: [304, 200, 204, 204, 204, 204, 204, 204].map(X), fontFace: 'Arial', fontSize: 20, color: P.INK, rowH: X(84), border: { type: 'solid', pt: 0.75, color: P.LINE }, margin: [0.05, 0.12, 0.05, 0.12], objectName: nm('Таблица') });
  card(s, 'Пояснение', 96, 790, 1728, 170, [R('Канада и США — значимый перекос в пользу ранних. Швеция — эффекта нет. Россия — обратная картина, но выборка мала.')], { pad: 14, size: 20, valign: 'middle' });
  s.addNotes('Резерв для вопросов.\n1. Больше ли рождается детей в начале года? — Нет, рождаемость почти равномерная, летом даже немного больше; с поправкой эффект сильнее.\n2. Почему нет эффекта у шведов? — Там жёсткий отбор начинается позже («спорт для всех»); но 80 человек — это предположение.\n3. Почему россияне не доказаны? — 52 человека мало, p = 0,47. Дальше — проверить КХЛ.\n4. Планировать ли рождение детей? — Нет, правильнее менять систему отбора.\n5. При чём психология? — Физиология даёт стартовую разницу, психология (Пигмалион, эффект Матфея, мотивация) превращает её в разницу на всю карьеру.\n6. Почему НХЛ, а не КХЛ? — Открытые полные данные и много стран.');
}

// 12 · Источники
{
  const s = slide('Слайд', 'Источники');
  const refs = ['Barnsley R. H., Thompson A. H., Barnsley P. E. Hockey success and birthdate: The relative age effect // CAHPER Journal. 1985.', 'Barnsley R. H. 40-years of relative age effects: life is not fair! // Frontiers in Sports and Active Living. 2025.', 'Gladwell M. Outliers: The Story of Success. 2008 («Гении и аутсайдеры»).', 'Cobley S. et al. Annual age-grouping and athlete development: a meta-analytical review // Sports Medicine. 2009.', 'Deaner R. O. et al. Born at the wrong time: selection bias in the NHL draft // PLoS ONE. 2013.', 'Gibbs B. G. et al. The rise of the underdog? // International Review for the Sociology of Sport. 2012.', 'Fumarco L. et al. The relative age effect reversal among the NHL elite // PLoS ONE. 2017.', 'Lemez S. et al. Relative age, competition level, and dropout rates in youth ice-hockey // Scand. J. Med. Sci. Sports. 2014.', 'Bedard K., Dhuey E. The persistence of early childhood maturity // Quarterly Journal of Economics. 2006.', 'Morrow R. L. et al. Influence of relative age on diagnosis and treatment of ADHD // CMAJ. 2012.', 'Layton T. J. et al. ADHD and month of school enrollment // NEJM. 2018.', 'Mann D. L., van Ginneken P. Age-ordered shirt numbering reduces the selection bias // Journal of Sports Sciences. 2017.', 'Cumming S. P. et al. Bio-banded tournament in Premier League academies // Journal of Sports Sciences. 2018.', 'Rosenthal R., Jacobson L. Pygmalion in the Classroom. 1968.', 'Merton R. K. The Matthew Effect in Science // Science. 1968.', 'Данные: составы клубов НХЛ 2024/25 (api-web.nhle.com).'];
  card(s, 'Список', 96, 190, 1728, 800, refs.map((r, i) => R(`${i + 1}. ${r}`, { breakLine: i < refs.length - 1 })), { pad: 20, size: 14, ls: 1.1, psa: 3 });
}

(async () => {
  const file = path.join(__dirname, 'out', 'deck-simple.pptx');
  fs.mkdirSync(path.dirname(file), { recursive: true });
  await pres.writeFile({ fileName: file });
  await applyTheme(file, THEME);
  fs.writeFileSync(path.join(__dirname, 'out', 'anim-simple.json'), JSON.stringify({ anim: ANIM, trans: TRANS }));
  console.log('OK', file, 'слайдов:', n);
})();
