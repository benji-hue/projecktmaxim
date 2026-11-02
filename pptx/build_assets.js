// Рендер графических ассетов для PowerPoint-версии из HTML-презентации (Playwright).
//   node build_assets.js [light|dark|all]
// Фоны слайдов — JPG 1920×1080; иллюстрации — PNG с прозрачностью, плотность ×2.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const HTML = 'file://' + path.resolve(__dirname, '..', 'presentation.html');
const OUT = path.resolve(__dirname, 'assets');

// Свет «арены» для каждого слайда PowerPoint (как data-glow в HTML)
const GLOWS = [
  [84, 4, 4, 98],   // 1 титул
  [12, 10, 94, 92], // 2 1984
  [70, 30, 10, 90], // 3 снежный ком
  [50, 0, 90, 100], // 4 масштаб
  [64, 22, 8, 96],  // 5 исследование
  [30, 10, 92, 90], // 6 загадка
  [40, 14, 90, 92], // 7 россияне
  [90, 40, 10, 10], // 8 школа
  [20, 0, 80, 100], // 9 что делать
  [50, 0, 50, 110], // 10 выводы
  [70, 10, 10, 90], // 11 данные
  [80, 90, 10, 0],  // 12 источники
];

const BASE_CSS = `
  *, *::before, *::after { transition: none !important; animation: none !important; }
  .toast, .controls, .goto, .tip, .notes-bar, .help, .blackout { display: none !important; }
`;

async function prepare(page, theme) {
  await page.goto(HTML);
  await page.addStyleTag({ content: BASE_CSS });
  await page.evaluate((t) => { localStorage.setItem('rae-deck-v1', JSON.stringify({ seen: 1 })); window.deck.setTheme(t); }, theme);
  await page.waitForTimeout(400);
}

async function isolate(page, css) {
  await page.evaluate((c) => {
    let s = document.getElementById('iso');
    if (!s) { s = document.createElement('style'); s.id = 'iso'; document.head.appendChild(s); }
    s.textContent = c;
  }, css);
  await page.waitForTimeout(150);
}

async function shotEl(page, selector, file, pad = 0) {
  const el = await page.$(selector);
  if (!el) throw new Error('нет элемента ' + selector);
  const b = await el.boundingBox();
  const clip = { x: Math.max(0, b.x - pad), y: Math.max(0, b.y - pad), width: b.width + pad * 2, height: b.height + pad * 2 };
  await page.screenshot({ path: file, clip, omitBackground: true });
  return clip;
}

const TRANSPARENT = `
  html, body, .stage, .deck { background: transparent !important; }
  .bg, .ice, .hud { display: none !important; }
`;

async function renderTheme(browser, theme) {
  const dir = path.join(OUT, theme);
  fs.mkdirSync(dir, { recursive: true });
  const meta = {};

  // ---------- фоны 1920×1080 ----------
  {
    const page = await browser.newPage({ viewport: { width: 1920, height: 1080 }, deviceScaleFactor: 1 });
    await prepare(page, theme);
    for (let k = 0; k < GLOWS.length; k++) {
      const g = GLOWS[k];
      await page.evaluate((g) => {
        const st = document.querySelector('.stage');
        ['--g1x', '--g1y', '--g2x', '--g2y'].forEach((v, j) => st.style.setProperty(v, g[j] + '%'));
      }, g);
      if (k === 0) {
        // титул: лёд + каток (линии прорисованы)
        await page.evaluate(() => window.deck.go(0, 0));
        await page.waitForTimeout(300);
        await isolate(page, `
          .hud { display: none !important; }
          #sl-title > *:not(.rink-wrap) { visibility: hidden !important; }
          .rink [pathLength] { stroke-dashoffset: 0 !important; }
          .rink .rk-fill { opacity: 1 !important; }
          #sl-title .rink-wrap { opacity: .62 !important; }
          .slide.is-active { animation: none !important; opacity: 1 !important; transform: none !important; }`);
      } else {
        await isolate(page, `.deck, .hud { display: none !important; }`);
      }
      await page.waitForTimeout(250);
      await page.screenshot({ path: path.join(dir, `bg-${String(k + 1).padStart(2, '0')}.jpg`), type: 'jpeg', quality: 88 });
    }
    await page.close();
  }

  // ---------- иллюстрации (×2, прозрачный фон) ----------
  const page = await browser.newPage({ viewport: { width: 1920, height: 1080 }, deviceScaleFactor: 2 });
  await prepare(page, theme);
  const settle = async (i, s) => {
    await page.evaluate(([i, s]) => window.deck.go(i, s), [i, s]);
    await page.waitForTimeout(500);
  };

  // программка 1984: основа и отдельно рукописная обводка
  await settle(1, 2);
  await isolate(page, TRANSPARENT + `
    #sl-1984 > *:not([data-anim]) , #sl-1984 .story, #sl-1984 .kicker, #sl-1984 .headline, #sl-1984 .src { visibility: hidden !important; }
    .program .pg-circle { visibility: hidden !important; }`);
  meta.program = await shotEl(page, '#sl-1984 .program', path.join(dir, 'program.png'), 60);
  await isolate(page, TRANSPARENT + `
    #sl-1984 * { visibility: hidden !important; }
    .program { background: none !important; box-shadow: none !important; }
    .program::after { display: none !important; }
    #sl-1984 .program .pg-circle, #sl-1984 .program .pg-circle * { visibility: visible !important; }
    .program .pg-circle path { stroke-dashoffset: 0 !important; }`);
  await shotEl(page, '#sl-1984 .program', path.join(dir, 'program-circle.png'), 60);

  // дети (слайд 3)
  await settle(2, 5);
  await isolate(page, TRANSPARENT + `#sl-snow * { visibility: hidden !important; } #sl-snow .kids, #sl-snow .kids * { visibility: visible !important; }`);
  meta.kids = await shotEl(page, '#sl-snow .kids svg', path.join(dir, 'kids.png'), 0);

  // фигурная скобка «364 дня»
  await isolate(page, TRANSPARENT + `#sl-snow * { visibility: hidden !important; } #sl-snow .brace svg, #sl-snow .brace svg * { visibility: visible !important; }`);
  await shotEl(page, '#sl-snow .brace svg', path.join(dir, 'brace.png'), 2);

  // джерси в рамке (слайды 6–7): без подписи, с тенью
  await settle(5, 0);
  await isolate(page, TRANSPARENT + `
    #sl-riddle > *:not(.frame) { visibility: hidden !important; }
    #sl-riddle .cap b, #sl-riddle .cap span, #sl-riddle .drop-hint { visibility: hidden !important; }
    #sl-riddle .frame { transform: none !important; }`);
  meta.frame = await shotEl(page, '#sl-riddle .frame', path.join(dir, 'frame.png'), 70);

  // иллюстрации «что делать»
  await settle(7, 3);
  for (const k of [1, 2, 3]) {
    await isolate(page, TRANSPARENT + `#sl-fix * { visibility: hidden !important; } #sl-fix .fix-${k} .ill, #sl-fix .fix-${k} .ill * { visibility: visible !important; }`);
    await shotEl(page, `#sl-fix .fix-${k} .ill`, path.join(dir, `fix-${k}.png`), 0);
  }
  await page.close();

  // ---------- векторные мелочи: дуга петли, иконки ----------
  const ink3 = theme === 'dark' ? '#7f93ab' : '#5f738a';
  const blue = theme === 'dark' ? '#3987e5' : '#2a78d6';
  const orange = theme === 'dark' ? '#d95926' : '#e2622f';
  const vec = await browser.newPage({ viewport: { width: 900, height: 400 }, deviceScaleFactor: 2 });
  const svgShot = async (svg, w, h, file) => {
    await vec.setContent(`<html><body style="margin:0;background:transparent">${svg}</body></html>`);
    await vec.screenshot({ path: path.join(dir, file), clip: { x: 0, y: 0, width: w, height: h }, omitBackground: true });
  };
  // дуга обратной связи: от верха «Ожидания» к верху «Отбора», стрелка вниз
  await svgShot(`<svg width="720" height="176" viewBox="0 0 720 176" xmlns="http://www.w3.org/2000/svg">
      <path d="M692,170 C692,6 28,6 28,156" fill="none" stroke="${ink3}" stroke-width="3" stroke-linecap="round"/>
      <path d="M17,148 L28,170 L39,148 Z" fill="${ink3}"/>
      <circle cx="360" cy="47" r="14" fill="${blue}"/></svg>`, 720, 176, 'loop-arc.png');
  const hand = (c) => `<svg width="128" height="128" viewBox="0 0 24 24" fill="none" stroke="${c}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" xmlns="http://www.w3.org/2000/svg">
      <path d="M18 11V6a2 2 0 0 0-2-2a2 2 0 0 0-2 2"/><path d="M14 10V4a2 2 0 0 0-2-2a2 2 0 0 0-2 2v2"/>
      <path d="M10 10.5V6a2 2 0 0 0-2-2a2 2 0 0 0-2 2v8"/>
      <path d="M18 8a2 2 0 1 1 4 0v6a8 8 0 0 1-8 8h-2c-2.8 0-4.5-.86-5.99-2.34l-3.6-3.6a2 2 0 0 1 2.83-2.82L7 15"/></svg>`;
  await svgShot(hand(blue), 128, 128, 'hand-blue.png');
  await svgShot(hand(orange), 128, 128, 'hand-orange.png');
  await svgShot(`<svg width="96" height="96" viewBox="0 0 24 24" fill="none" stroke="${theme === 'dark' ? '#8dbcf4' : '#1f5fc4'}" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" xmlns="http://www.w3.org/2000/svg"><path d="M20 6 9 17l-5-5"/></svg>`, 96, 96, 'check.png');
  await vec.close();

  fs.writeFileSync(path.join(dir, 'meta.json'), JSON.stringify(meta, null, 2));
  console.log('ассеты:', theme, Object.keys(meta).join(', '));
}

(async () => {
  const which = process.argv[2] || 'all';
  const browser = await chromium.launch();
  for (const t of which === 'all' ? ['light', 'dark'] : [which]) await renderTheme(browser, t);
  await browser.close();
})();
