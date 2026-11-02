/* Эффект относительного возраста — движок презентации. Без зависимостей, работает офлайн. */
(() => {
  'use strict';

  // Снимок исходного документа до любых изменений — для «Сохранить файл» (Ctrl+S)
  const PRISTINE = '<!DOCTYPE html>\n' + document.documentElement.outerHTML;

  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const root = document.documentElement;
  const body = document.body;
  const stage = $('.stage');
  const slides = $$('.deck > .slide');
  const MAIN = slides.filter((s) => !s.hasAttribute('data-appendix'));
  const N_MAIN = MAIN.length;
  const reduceMotion = () =>
    body.classList.contains('reduce-motion') || matchMedia('(prefers-reduced-motion: reduce)').matches;

  /* ---------- хранилище: только удобства этого зрителя (тема, правки, фото) ---------- */
  const KEY = 'rae-deck-v1';
  let store = {};
  try { store = JSON.parse(localStorage.getItem(KEY) || '{}') || {}; } catch (e) { store = {}; }
  const persist = () => {
    try { localStorage.setItem(KEY, JSON.stringify(store)); return true; } catch (e) { return false; }
  };

  /* ---------- подготовка слайдов ---------- */
  slides.forEach((s, i) => {
    s._i = i;
    s._frags = $$('[data-step]', s);
    const maxFrag = s._frags.reduce((m, el) => Math.max(m, +el.dataset.step || 0), 0);
    s._steps = Math.max(+s.dataset.steps || 0, maxFrag);
    s._notes = $$('aside.notes', s).map((a) => ({ step: +a.dataset.step || 0, html: a.innerHTML.trim() }));
    s.setAttribute('role', 'group');
    s.setAttribute('aria-roledescription', 'слайд');
    s.setAttribute('aria-label', label(s));
  });
  function appNo(s) { return s._i - N_MAIN + 1; }
  function label(s) {
    return (s.hasAttribute('data-appendix') ? 'Приложение ' + appNo(s) : 'Слайд ' + (s._i + 1) + ' из ' + N_MAIN) +
      ': ' + (s.dataset.title || '');
  }

  /* ---------- состояние и шаги ---------- */
  const state = { i: 0, step: 0 };

  function applyStep(s, step) {
    for (let k = 1; k <= s._steps; k++) s.classList.toggle('st-' + k, k <= step);
    s._frags.forEach((el) => el.classList.toggle('is-shown', (+el.dataset.step || 0) <= step));
    runCounts(s);
  }

  function enter(s) {
    clearTimeout(s._t);
    s.classList.remove('is-leaving', 'is-in');
    s.classList.add('is-active');
    void s.offsetWidth;
    requestAnimationFrame(() => requestAnimationFrame(() => {
      if (s.classList.contains('is-active')) { s.classList.add('is-in'); runCounts(s); }
    }));
  }
  function leave(s) {
    s.classList.remove('is-active', 'is-in');
    s.classList.add('is-leaving');
    clearTimeout(s._t);
    s._t = setTimeout(() => { s.classList.remove('is-leaving'); resetCounts(s); applyStep(s, 0); }, 430);
  }

  function go(i, step = 0) {
    i = clamp(i, 0, slides.length - 1);
    const to = slides[i];
    step = clamp(step, 0, to._steps);
    if (i === state.i && to.classList.contains('is-active')) {
      state.step = step;
      applyStep(to, step);
      sync();
      return;
    }
    const from = slides[state.i];
    stage.style.setProperty('--dir', i >= state.i ? 1 : -1);
    if (from && from !== to && from.classList.contains('is-active')) leave(from);
    if (i > state.i && step === 0) $$('.is-guess', to).forEach((b) => b.classList.remove('is-guess'));
    state.i = i;
    state.step = step;
    applyStep(to, step);
    enter(to);
    glow(to);
    sync();
  }
  function next() {
    const s = slides[state.i];
    if (!timer.running && timer.acc === 0 && state.i === 0) startTimer();
    if (state.step < s._steps) { state.step++; applyStep(s, state.step); sync(); }
    else if (state.i < slides.length - 1) go(state.i + 1, 0);
  }
  function prev() {
    const s = slides[state.i];
    if (state.step > 0) { state.step--; applyStep(s, state.step); sync(); }
    else if (state.i > 0) go(state.i - 1, slides[state.i - 1]._steps);
  }

  function glow(s) {
    const g = (s.dataset.glow || '84 6 6 96').split(/\s+/).map(Number);
    ['--g1x', '--g1y', '--g2x', '--g2y'].forEach((k, j) => stage.style.setProperty(k, (g[j] || 0) + '%'));
  }

  /* ---------- счётчики чисел ---------- */
  const fmt = (v, dec) => (dec ? v.toFixed(dec).replace('.', ',') : String(Math.round(v)));
  function runCounts(s) {
    if (!s.classList.contains('is-in')) return;
    $$('[data-count]', s).forEach((el) => {
      const frag = el.closest('[data-step]');
      if ((frag && !frag.classList.contains('is-shown')) || el._counted) return;
      el._counted = true;
      const target = parseFloat(el.dataset.count);
      const dec = +(el.dataset.dec || 0);
      if (reduceMotion()) { el.textContent = fmt(target, dec); return; }
      const t0 = performance.now();
      const tick = (t) => {
        const p = clamp((t - t0 - 300) / 1300, 0, 1);
        const e = p >= 1 ? 1 : 1 - Math.pow(2, -10 * p);
        el.textContent = fmt(target * e, dec);
        if (p < 1) requestAnimationFrame(tick);
      };
      el.textContent = fmt(0, dec);
      requestAnimationFrame(tick);
    });
  }
  function resetCounts(s) {
    $$('[data-count]', s).forEach((el) => {
      el._counted = false;
      el.textContent = fmt(parseFloat(el.dataset.count), +(el.dataset.dec || 0));
    });
  }

  /* ---------- HUD: счётчик, прогресс, адрес, заметки ---------- */
  const counter = $('.counter');
  const progress = $('.progress');
  const live = $('#live');
  progress.innerHTML = MAIN.map((s) => `<i style="--w:${Math.max(4, +s.dataset.duration || 4)}"></i>`).join('');
  const segs = $$('i', progress);

  function sync() {
    const s = slides[state.i];
    const app = s.hasAttribute('data-appendix');
    counter.innerHTML = app
      ? `<b>П${appNo(s)}</b> · приложение`
      : `<b>${String(state.i + 1).padStart(2, '0')}</b> / ${String(N_MAIN).padStart(2, '0')}`;
    segs.forEach((seg, k) => {
      const p = app || k < state.i ? 1 : k > state.i ? 0 : (state.step + 1) / (s._steps + 1);
      seg.style.setProperty('--p', p);
    });
    progress.classList.toggle('is-appendix', app);
    const h = '#' + (state.i + 1) + (state.step ? '.' + state.step : '');
    if (location.hash !== h) { try { history.replaceState(null, '', h); } catch (e) { /* file:// в старых браузерах */ } }
    live.textContent = label(s);
    renderNotes();
    updatePresenter();
  }

  const notesBody = $('.notes-bar .nb-body');
  function renderNotes() {
    if (!body.classList.contains('show-notes')) return;
    const s = slides[state.i];
    notesBody.innerHTML = s._notes.map((n) =>
      `<p class="${n.step === state.step ? 'is-now' : n.step === state.step + 1 ? 'is-next' : ''}">${n.html}</p>`
    ).join('') || '<p class="is-now">—</p>';
    const cur = $('.is-now', notesBody);
    if (cur) cur.scrollIntoView({ block: 'nearest' });
  }

  /* ---------- масштаб сцены под любой экран ---------- */
  function fit() {
    const s = Math.min(innerWidth / 1920, innerHeight / 1080);
    stage.style.setProperty('--s', s);
  }
  addEventListener('resize', fit);

  /* ---------- текстура льда: следы коньков, стружка, иней ---------- */
  function mulberry32(a) {
    return function () {
      a |= 0; a = (a + 0x6d2b79f5) | 0;
      let t = Math.imul(a ^ (a >>> 15), 1 | a);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }
  function paintIce() {
    const c = $('.ice');
    if (!c || !c.getContext) return;
    const g = c.getContext('2d');
    const W = 1920, H = 1080;
    const rnd = mulberry32(19840101);
    const dark = root.dataset.theme === 'dark';
    g.clearRect(0, 0, W, H);
    g.lineCap = 'round';
    g.lineJoin = 'round';
    const stroke = (pts, w, col) => {
      g.beginPath();
      g.moveTo(pts[0][0], pts[0][1]);
      for (let k = 1; k < pts.length; k++) g.lineTo(pts[k][0], pts[k][1]);
      g.lineWidth = w;
      g.strokeStyle = col;
      g.stroke();
    };
    for (let i = 0; i < 95; i++) {
      let x = rnd() * W, y = rnd() * H;
      let a = (rnd() < 0.5 ? 0 : Math.PI) + (rnd() - 0.5) * 0.9;
      const len = 240 + rnd() * 1100, curv = (rnd() - 0.5) * 0.004, step = 10;
      const pts = [[x, y]];
      for (let d = 0; d < len; d += step) { a += curv * step; x += Math.cos(a) * step; y += Math.sin(a) * step; pts.push([x, y]); }
      const w = 0.5 + rnd() * 1.6, al = 0.18 + rnd() * 0.4;
      stroke(pts, w, dark ? `rgba(190,220,255,${al * 0.5})` : `rgba(255,255,255,${Math.min(1, al + 0.2)})`);
      if (!dark) stroke(pts.map(([px, py]) => [px + 0.8, py + 1.4]), 0.7, `rgba(60,95,140,${0.04 + rnd() * 0.05})`);
    }
    for (let i = 0; i < 14; i++) {
      const cx = rnd() * W, cy = rnd() * H, base = rnd() * Math.PI * 2;
      for (let j = 0; j < 42; j++) {
        const a = base + (rnd() - 0.5) * 0.7, r0 = rnd() * 40, r1 = r0 + 20 + rnd() * 70;
        stroke([[cx + Math.cos(a) * r0, cy + Math.sin(a) * r0], [cx + Math.cos(a) * r1, cy + Math.sin(a) * r1]],
          0.4 + rnd() * 0.8, dark ? `rgba(200,225,255,${0.05 + rnd() * 0.1})` : `rgba(255,255,255,${0.3 + rnd() * 0.4})`);
      }
    }
    for (let i = 0; i < 6500; i++) {
      g.fillStyle = dark ? `rgba(200,225,255,${rnd() * 0.08})` : `rgba(255,255,255,${rnd() * 0.6})`;
      const sz = rnd() < 0.9 ? 1 : 2;
      g.fillRect(rnd() * W, rnd() * H, sz, sz);
    }
  }

  /* ---------- тема: светлая (проектор в светлом зале) / тёмная (тёмный зал) ---------- */
  function setTheme(t, origin) {
    const apply = () => {
      root.dataset.theme = t;
      store.theme = t;
      persist();
      paintIce();
      updatePresenter();
    };
    if (document.startViewTransition && !reduceMotion() && origin) {
      const x = origin.x, y = origin.y;
      const r = Math.hypot(Math.max(x, innerWidth - x), Math.max(y, innerHeight - y));
      const vt = document.startViewTransition(apply);
      vt.ready.then(() => {
        root.animate(
          { clipPath: [`circle(0px at ${x}px ${y}px)`, `circle(${r}px at ${x}px ${y}px)`] },
          { duration: 800, easing: 'cubic-bezier(.65,0,.35,1)', pseudoElement: '::view-transition-new(root)' }
        );
      }).catch(() => {});
    } else apply();
  }
  const toggleTheme = (origin) => setTheme(root.dataset.theme === 'dark' ? 'light' : 'dark', origin);

  /* ---------- полноэкранный режим, чёрный экран, справка, заметки ---------- */
  function toggleFS() {
    const fs = document.fullscreenElement || document.webkitFullscreenElement;
    try {
      if (!fs) {
        const r = (root.requestFullscreen || root.webkitRequestFullscreen).call(root);
        if (r && r.catch) r.catch(() => {});
      } else (document.exitFullscreen || document.webkitExitFullscreen).call(document);
    } catch (e) { /* нет поддержки */ }
  }
  const toggleBlack = (f) => body.classList.toggle('is-black', f);
  const toggleHelp = (f) => body.classList.toggle('show-help', f);
  function toggleNotes(f) { body.classList.toggle('show-notes', f); renderNotes(); }
  function toggleLite() {
    const on = body.classList.toggle('reduce-motion');
    body.classList.toggle('lite', on);
    toast(on ? 'Облегчённый режим: без размытия и анимаций — для слабого компьютера. <b>M</b> — вернуть.' : 'Полный режим с анимациями.');
  }

  /* ---------- обзор всех слайдов ---------- */
  let overview = false;
  function toggleOverview(force) {
    overview = typeof force === 'boolean' ? force : !overview;
    body.classList.toggle('is-overview', overview);
    if (overview) {
      const cols = 4, gap = 36, pad = 56;
      const sc = (1920 - pad * 2 - gap * (cols - 1)) / cols / 1920;
      const rows = Math.ceil(slides.length / cols);
      const top = Math.max(pad, (1080 - (rows * 1080 * sc + (rows - 1) * gap)) / 2);
      slides.forEach((s, k) => {
        const c = k % cols, r = Math.floor(k / cols);
        s.style.setProperty('--ox', pad + c * (1920 * sc + gap) + 'px');
        s.style.setProperty('--oy', top + r * (1080 * sc + gap) + 'px');
        s.style.setProperty('--os', sc);
        clearTimeout(s._t);
        s.classList.remove('is-leaving');
        s.classList.add('is-in');
        applyStep(s, s._steps);
        s.classList.toggle('is-current-ov', k === state.i);
      });
    } else {
      slides.forEach((s, k) => {
        s.classList.remove('is-current-ov');
        if (k !== state.i) { s.classList.remove('is-in', 'is-active'); applyStep(s, 0); resetCounts(s); }
      });
      const s = slides[state.i];
      applyStep(s, state.step);
      s.classList.add('is-active', 'is-in');
    }
  }

  /* ---------- уведомления ---------- */
  const toastEl = $('.toast');
  let toastT;
  function toast(html, ms = 4600) {
    toastEl.innerHTML = html;
    toastEl.classList.add('is-on');
    clearTimeout(toastT);
    toastT = setTimeout(() => toastEl.classList.remove('is-on'), ms);
  }

  /* ---------- правка текста: имя и направление ---------- */
  let editing = false;
  const editables = $$('[data-editable]');
  if (store.edits) {
    editables.forEach((el) => {
      const v = store.edits[el.dataset.editable];
      if (typeof v === 'string' && v.trim()) el.textContent = v;
    });
  }
  function mirror() {
    $$('[data-mirror]').forEach((m) => {
      const src = $(`[data-editable="${m.dataset.mirror}"]`);
      if (src) m.textContent = src.textContent;
    });
  }
  mirror();
  editables.forEach((el) => el.addEventListener('input', () => {
    store.edits = store.edits || {};
    store.edits[el.dataset.editable] = el.textContent.trim();
    persist();
    mirror();
  }));
  function toggleEdit(force) {
    editing = typeof force === 'boolean' ? force : !editing;
    body.classList.toggle('is-edit', editing);
    editables.forEach((el) => {
      if (editing) { try { el.contentEditable = 'plaintext-only'; } catch (e) { el.contentEditable = 'true'; } el.spellcheck = false; }
      else el.removeAttribute('contenteditable');
    });
    if (editing) {
      toast('Режим правки. Слайд 1: кликните по имени или направлению и введите текст. Слайд 6: перетащите фото на рамку (двойной клик — выбрать файл, перетаскивание фото — сдвиг кадра). <b>Ctrl+S</b> — сохранить файл, <b>E</b> — выйти.', 9000);
    } else {
      if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
      toast('Правка завершена. Не забудьте <b>Ctrl+S</b>, чтобы правки и фото оказались внутри файла.');
    }
  }

  /* ---------- фото: перетащить на рамку → уменьшить → встроить ---------- */
  const slotsEls = $$('[data-photo-slot]');
  function setPhoto(slot, src, pos) {
    const img = $('img.photo', slot);
    if (!img || !src) return;
    img.src = src;
    slot.classList.add('has-photo');
    if (pos) slot.style.setProperty('--pos', pos);
  }
  slotsEls.forEach((slot) => {
    const key = slot.dataset.photoSlot;
    const fromStore = store.photos && store.photos[key];
    const fromFile = slot.getAttribute('data-photo');
    setPhoto(slot, fromStore || fromFile, (store.pos && store.pos[key]) || slot.getAttribute('data-photo-pos'));
    slot.addEventListener('dragover', (e) => { e.preventDefault(); slot.classList.add('is-over'); });
    slot.addEventListener('dragleave', () => slot.classList.remove('is-over'));
    slot.addEventListener('drop', (e) => {
      e.preventDefault();
      e.stopPropagation();
      slot.classList.remove('is-over');
      body.classList.remove('is-dragging');
      const f = Array.from(e.dataTransfer.files || []).find((x) => /^image\//.test(x.type));
      if (f) acceptFile(slot, f);
      else toast('Нужен файл изображения: JPG, PNG или WebP.');
    });
    slot.addEventListener('dblclick', (e) => {
      if (!editing) return;
      e.stopPropagation();
      const inp = document.createElement('input');
      inp.type = 'file';
      inp.accept = 'image/*';
      inp.onchange = () => inp.files[0] && acceptFile(slot, inp.files[0]);
      inp.click();
    });
    // сдвиг кадра в режиме правки
    let drag = null;
    slot.addEventListener('pointerdown', (e) => {
      if (!editing || !slot.classList.contains('has-photo')) return;
      const cs = getComputedStyle(slot).getPropertyValue('--pos').trim() || '50% 30%';
      const [px, py] = cs.split(/\s+/).map(parseFloat);
      drag = { x: e.clientX, y: e.clientY, px: isNaN(px) ? 50 : px, py: isNaN(py) ? 30 : py };
      slot.setPointerCapture(e.pointerId);
      e.preventDefault();
    });
    slot.addEventListener('pointermove', (e) => {
      if (!drag) return;
      const nx = clamp(drag.px - (e.clientX - drag.x) * 0.25, 0, 100);
      const ny = clamp(drag.py - (e.clientY - drag.y) * 0.25, 0, 100);
      const v = `${nx.toFixed(1)}% ${ny.toFixed(1)}%`;
      slot.style.setProperty('--pos', v);
      store.pos = store.pos || {};
      store.pos[key] = v;
    });
    const end = () => { if (drag) { drag = null; persist(); } };
    slot.addEventListener('pointerup', end);
    slot.addEventListener('pointercancel', end);
  });
  const hasFiles = (e) => e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');
  let dragDepth = 0;
  document.addEventListener('dragenter', (e) => { if (hasFiles(e)) { dragDepth++; body.classList.add('is-dragging'); } });
  document.addEventListener('dragleave', () => { dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) body.classList.remove('is-dragging'); });
  // не дать браузеру открыть картинку вместо презентации, если промахнулись мимо рамки
  document.addEventListener('dragover', (e) => e.preventDefault());
  document.addEventListener('drop', (e) => { e.preventDefault(); dragDepth = 0; body.classList.remove('is-dragging'); });

  function acceptFile(slot, file) {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const max = 1600;
        const k = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
        const cv = document.createElement('canvas');
        cv.width = Math.round(img.naturalWidth * k);
        cv.height = Math.round(img.naturalHeight * k);
        cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height);
        const out = cv.toDataURL('image/jpeg', 0.86);
        setPhoto(slot, out);
        store.photos = store.photos || {};
        store.photos[slot.dataset.photoSlot] = out;
        const ok = persist();
        toast('Фото добавлено. Нажмите <b>Ctrl+S</b> — сохранится новый файл презентации с фото внутри.' +
          (ok ? '' : '<br>Браузер не дал запомнить фото — сохраните файл сейчас.'), 7000);
      };
      img.onerror = () => toast('Не удалось прочитать изображение.');
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  }

  /* ---------- сохранить самодостаточный файл с правками и фото ---------- */
  async function saveFile() {
    const doc = new DOMParser().parseFromString(PRISTINE, 'text/html');
    editables.forEach((el) => {
      const t = doc.querySelector(`[data-editable="${el.dataset.editable}"]`);
      if (t) t.textContent = el.textContent.trim();
    });
    $$('[data-mirror]', doc).forEach((m) => {
      const src = doc.querySelector(`[data-editable="${m.dataset.mirror}"]`);
      if (src) m.textContent = src.textContent;
    });
    slotsEls.forEach((slot) => {
      const key = slot.dataset.photoSlot;
      const t = doc.querySelector(`[data-photo-slot="${key}"]`);
      const img = $('img.photo', slot);
      if (t && slot.classList.contains('has-photo') && img && img.src) {
        t.setAttribute('data-photo', img.src);
        const pos = slot.style.getPropertyValue('--pos');
        if (pos) t.setAttribute('data-photo-pos', pos.trim());
      }
    });
    const html = '<!DOCTYPE html>\n' + doc.documentElement.outerHTML.replace(/^<!DOCTYPE html>\n?/i, '');
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const name = 'presentation.html';
    if (window.showSaveFilePicker) {
      try {
        const h = await window.showSaveFilePicker({ suggestedName: name, types: [{ description: 'Презентация HTML', accept: { 'text/html': ['.html'] } }] });
        const w = await h.createWritable();
        await w.write(blob);
        await w.close();
        toast('Сохранено: <b>' + h.name + '</b>. Фото и правки уже внутри файла — интернет не нужен.');
        return;
      } catch (e) {
        if (e && e.name === 'AbortError') return;
      }
    }
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    toast('Файл <b>' + name + '</b> сохранён в «Загрузки». Открывайте на защите именно его — фото и правки внутри.', 7000);
  }

  /* ---------- окно докладчика: текст, таймер, хронометраж, ответы на вопросы ---------- */
  let pw = null;
  let timer = { start: 0, acc: 0, running: false };
  const plan = MAIN.map((s) => +s.dataset.duration || 0);
  const planTotal = plan.reduce((a, b) => a + b, 0);
  const elapsed = () => timer.acc + (timer.running ? Date.now() - timer.start : 0);
  function startTimer() { if (!timer.running) { timer.start = Date.now(); timer.running = true; } }
  function pauseTimer() { if (timer.running) { timer.acc += Date.now() - timer.start; timer.running = false; } }
  function resetTimer() { timer = { start: 0, acc: 0, running: false }; }
  const mmss = (ms) => {
    const t = Math.round(Math.abs(ms) / 1000);
    return Math.floor(t / 60) + ':' + String(t % 60).padStart(2, '0');
  };
  function plannedNow() {
    const s = slides[state.i];
    if (s.hasAttribute('data-appendix')) return planTotal * 1000;
    let t = 0;
    for (let k = 0; k < state.i; k++) t += plan[k];
    t += plan[state.i] * (state.step / (s._steps + 1));
    return t * 1000;
  }

  const PRESENTER_HTML = `<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8"><title>Окно докладчика</title><style>
  :root{color-scheme:dark}*{box-sizing:border-box}
  body{margin:0;height:100vh;display:grid;grid-template-rows:auto 1fr auto;background:#0a1322;color:#eaf2fb;font:16px/1.5 "Segoe UI",system-ui,-apple-system,Roboto,sans-serif}
  header{display:flex;align-items:center;gap:22px;padding:16px 26px;border-bottom:1px solid rgba(255,255,255,.08)}
  .tm{font:600 56px/1 Consolas,ui-monospace,monospace;letter-spacing:-.02em;font-variant-numeric:tabular-nums}
  .tm.paused{opacity:.5}
  .pace{padding:9px 14px;border-radius:12px;background:rgba(255,255,255,.06);font:600 15px/1.35 Consolas,ui-monospace,monospace}
  .pace.ok{color:#8fe39b}.pace.late{color:#ffa27a}.pace.early{color:#9cc9ff}
  .ck{color:#7f93ab;font:500 15px Consolas,ui-monospace,monospace}
  .meta{margin-left:auto;text-align:right}.meta b{display:block;font-size:24px}.meta span{color:#a9bdd3}
  main{display:grid;grid-template-columns:1.65fr 1fr;min-height:0}
  .now{padding:22px 30px;overflow:auto}
  .now p{margin:0 0 14px;font-size:25px;line-height:1.45;opacity:.36;transition:opacity .2s}
  .now p.cur{opacity:1}.now p.cur::before{content:"▶ ";color:#3987e5}
  .now i{color:#ffb38f;font-style:normal}
  .side{padding:22px 26px;border-left:1px solid rgba(255,255,255,.08);overflow:auto;display:flex;flex-direction:column;gap:18px}
  h4{margin:0 0 8px;font:600 12px/1 Consolas,ui-monospace,monospace;letter-spacing:.14em;text-transform:uppercase;color:#8dbcf4}
  .nx{padding:14px 16px;border-radius:14px;background:rgba(255,255,255,.05);font-size:17px;color:#cfe0f5}
  .bar{display:flex;gap:3px;height:10px}.bar i{flex:var(--w);border-radius:3px;background:rgba(255,255,255,.1);position:relative;overflow:hidden}
  .bar i::after{content:"";position:absolute;inset:0;background:#3987e5;transform-origin:0 50%;transform:scaleX(var(--p,0))}
  .bar i.cur{outline:2px solid #e2622f;outline-offset:1px}
  .legend{margin-top:6px;color:#7f93ab;font-size:13px}
  details{border-top:1px solid rgba(255,255,255,.08);padding:9px 0}summary{cursor:pointer;font-weight:600}details p{margin:6px 0 0;color:#cfe0f5;font-size:15px}
  footer{display:flex;flex-wrap:wrap;gap:10px;padding:12px 26px;border-top:1px solid rgba(255,255,255,.08)}
  button{font:600 15px/1 "Segoe UI",system-ui,sans-serif;padding:12px 18px;border:0;border-radius:12px;background:rgba(255,255,255,.08);color:#fff;cursor:pointer}
  button:hover{background:rgba(255,255,255,.16)}button.pri{background:#2a78d6}button.pri:hover{background:#3987e5}
  .hint{margin-left:auto;align-self:center;color:#7f93ab;font-size:13px}
  </style></head><body>
  <header><div class="tm" id="tm">0:00</div><div class="pace" id="pc"></div><div class="ck" id="ck"></div><div class="meta"><b id="sl"></b><span id="st"></span></div></header>
  <main><section class="now" id="now"></section>
  <aside class="side"><div><h4>Дальше</h4><div class="nx" id="nx"></div></div>
  <div><h4>Хронометраж · план 4:40</h4><div class="bar" id="bar"></div><div class="legend">Ширина сегмента — запланированное время слайда</div></div>
  <div><h4>Вопросы комиссии</h4><div id="qa"></div></div></aside></main>
  <footer><button data-p="prev">← Назад</button><button class="pri" data-p="next">Вперёд →</button><button data-p="timer" id="tb">Пауза</button><button data-p="reset">Сбросить таймер</button><button data-p="black">Чёрный экран</button><button data-p="theme">Тема</button><span class="hint">Работают клавиши: → ← пробел B T F</span></footer>
  </body></html>`;

  let pTick = null;
  function openPresenter() {
    if (pw && !pw.closed) { pw.focus(); return; }
    pw = window.open('', 'rae-presenter', 'popup=yes,width=1280,height=820');
    if (!pw) {
      toast('Браузер заблокировал окно докладчика. Разрешите всплывающие окна для этого файла — или нажмите <b>N</b>, и текст появится внизу экрана.', 8000);
      return;
    }
    const d = pw.document;
    d.open();
    d.write(PRESENTER_HTML);
    d.close();
    $('#qa', d).innerHTML = $('#qa').innerHTML;
    $('#bar', d).innerHTML = MAIN.map((s) => `<i style="--w:${Math.max(4, +s.dataset.duration || 4)}"></i>`).join('');
    d.addEventListener('keydown', (e) => onKey(e));
    d.addEventListener('click', (e) => {
      const b = e.target.closest('[data-p]');
      if (!b) return;
      const a = b.dataset.p;
      if (a === 'next') next();
      else if (a === 'prev') prev();
      else if (a === 'timer') { timer.running ? pauseTimer() : startTimer(); updatePresenter(); }
      else if (a === 'reset') { resetTimer(); updatePresenter(); }
      else if (a === 'black') toggleBlack();
      else if (a === 'theme') toggleTheme();
    });
    clearInterval(pTick);
    pTick = setInterval(() => { if (!pw || pw.closed) { clearInterval(pTick); pw = null; return; } updateClock(); }, 500);
    updatePresenter();
  }
  function updateClock() {
    if (!pw || pw.closed) return;
    const d = pw.document;
    const tm = $('#tm', d);
    if (!tm) return;
    const el = elapsed();
    tm.textContent = mmss(el);
    tm.classList.toggle('paused', !timer.running);
    const delta = el - plannedNow();
    const pc = $('#pc', d);
    const app = slides[state.i].hasAttribute('data-appendix');
    if (!timer.running && el === 0) { pc.className = 'pace'; pc.textContent = 'таймер стартует с первого клика'; }
    else if (app) { pc.className = 'pace'; pc.textContent = 'вопросы комиссии'; }
    else if (Math.abs(delta) < 8000) { pc.className = 'pace ok'; pc.textContent = 'по плану'; }
    else if (delta > 0) { pc.className = 'pace late'; pc.textContent = 'отстаёте на ' + mmss(delta); }
    else { pc.className = 'pace early'; pc.textContent = 'опережаете на ' + mmss(delta); }
    const now = new Date();
    $('#ck', d).textContent = now.getHours() + ':' + String(now.getMinutes()).padStart(2, '0');
    $('#tb', d).textContent = timer.running ? 'Пауза' : 'Старт';
  }
  function updatePresenter() {
    if (!pw || pw.closed) return;
    const d = pw.document;
    const s = slides[state.i];
    const app = s.hasAttribute('data-appendix');
    const now = $('#now', d);
    if (!now) return;
    $('#sl', d).textContent = app ? 'Приложение ' + appNo(s) : 'Слайд ' + (state.i + 1) + ' / ' + N_MAIN;
    $('#st', d).textContent = (s.dataset.title || '') + (s._steps ? ' · шаг ' + state.step + ' из ' + s._steps : '');
    now.innerHTML = s._notes.map((n) => `<p class="${n.step === state.step ? 'cur' : ''}">${n.html}</p>`).join('');
    const cur = $('.cur', now);
    if (cur) cur.scrollIntoView({ block: 'center' });
    let nx = '';
    const nn = s._notes.find((n) => n.step === state.step + 1);
    if (nn) nx = nn.html;
    else if (state.i < slides.length - 1) {
      const ns = slides[state.i + 1];
      nx = '<b>Следующий слайд:</b> ' + (ns.dataset.title || '') + (ns._notes[0] ? '<br>' + ns._notes[0].html : '');
    } else nx = 'Конец презентации.';
    $('#nx', d).innerHTML = nx;
    $$('#bar i', d).forEach((seg, k) => {
      const p = app || k < state.i ? 1 : k > state.i ? 0 : (state.step + 1) / (s._steps + 1);
      seg.style.setProperty('--p', p);
      seg.classList.toggle('cur', k === state.i && !app);
    });
    updateClock();
  }

  /* ---------- подсказки на графиках ---------- */
  const tip = $('.tip');
  document.addEventListener('pointerover', (e) => {
    const h = e.target.closest && e.target.closest('[data-tip]');
    if (!h || overview) return;
    tip.textContent = h.dataset.tip;
    tip.classList.add('is-on');
  });
  document.addEventListener('pointermove', (e) => {
    if (!tip.classList.contains('is-on')) return;
    tip.style.left = e.clientX + 'px';
    tip.style.top = e.clientY + 'px';
  }, { passive: true });
  document.addEventListener('pointerout', (e) => {
    if (e.target.closest && e.target.closest('[data-tip]')) tip.classList.remove('is-on');
  });

  /* ---------- загадка: зал называет месяцы — можно отмечать ---------- */
  $$('#sl-riddle .picker button').forEach((b) => b.addEventListener('click', (e) => {
    e.stopPropagation();
    b.classList.toggle('is-guess');
  }));

  /* ---------- мышь, тач, интерфейс ---------- */
  let swipeAt = 0;
  stage.addEventListener('click', (e) => {
    if (overview) {
      const s = e.target.closest('.slide');
      if (s) { toggleOverview(false); go(s._i, 0); }
      return;
    }
    if (editing || Date.now() - swipeAt < 400) return;
    if (e.target.closest('button, a, input, [contenteditable]')) return;
    if (String(window.getSelection ? window.getSelection() : '')) return;
    next();
    wake();
  });
  let sx = null, sy = null;
  stage.addEventListener('pointerdown', (e) => { if (e.pointerType === 'touch') { sx = e.clientX; sy = e.clientY; } });
  stage.addEventListener('pointerup', (e) => {
    if (sx === null) return;
    const dx = e.clientX - sx, dy = e.clientY - sy;
    sx = null;
    if (Math.abs(dx) > 50 && Math.abs(dx) > Math.abs(dy) * 1.2) { swipeAt = Date.now(); dx < 0 ? next() : prev(); }
  });

  let idleT;
  function poke() {
    body.classList.add('show-ui');
    body.classList.remove('is-idle');
    clearTimeout(idleT);
    idleT = setTimeout(() => {
      body.classList.remove('show-ui');
      if (document.fullscreenElement || document.webkitFullscreenElement) body.classList.add('is-idle');
    }, 2400);
  }
  addEventListener('mousemove', poke, { passive: true });

  const ACT = {
    prev, next,
    overview: () => toggleOverview(),
    presenter: openPresenter,
    theme: () => {
      const b = $('.controls [data-act="theme"]').getBoundingClientRect();
      toggleTheme({ x: b.left + b.width / 2, y: b.top + b.height / 2 });
    },
    fullscreen: toggleFS,
    help: () => toggleHelp(),
  };
  $('.controls').addEventListener('click', (e) => {
    const b = e.target.closest('[data-act]');
    if (!b) return;
    e.stopPropagation();
    ACT[b.dataset.act]();
  });
  $('.help').addEventListener('click', () => toggleHelp(false));

  /* ---------- клавиатура (по физическим клавишам — работает и в русской раскладке) ---------- */
  let digits = '';
  let digitsT;
  const gotoEl = $('.goto');
  function onKey(e) {
    const t = e.target;
    const inField = t && (t.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName || ''));
    const code = e.code || '';
    const key = e.key;
    if ((e.ctrlKey || e.metaKey) && code === 'KeyS') { e.preventDefault(); saveFile(); return; }
    if (inField) {
      if (key === 'Escape' || key === 'Enter') { e.preventDefault(); t.blur(); }
      return;
    }
    if (e.altKey || e.ctrlKey || e.metaKey) return;
    wake();

    const d = /^(Digit|Numpad)(\d)$/.exec(code);
    if (d) {
      digits = (digits + d[2]).slice(-2);
      gotoEl.textContent = '→ ' + digits;
      gotoEl.classList.add('is-on');
      clearTimeout(digitsT);
      digitsT = setTimeout(() => { digits = ''; gotoEl.classList.remove('is-on'); }, 1800);
      e.preventDefault();
      return;
    }
    if ((key === 'Enter' || code === 'NumpadEnter') && digits) {
      const n = clamp(parseInt(digits, 10), 1, slides.length);
      digits = '';
      gotoEl.classList.remove('is-on');
      if (overview) toggleOverview(false);
      go(n - 1, 0);
      e.preventDefault();
      return;
    }
    if (body.classList.contains('show-help') && key !== 'Escape' && code !== 'KeyH') { toggleHelp(false); }

    let handled = true;
    switch (key) {
      case 'ArrowRight': case 'ArrowDown': case 'PageDown': case ' ': case 'Enter':
        if (overview) toggleOverview(false); else next();
        break;
      case 'ArrowLeft': case 'ArrowUp': case 'PageUp': case 'Backspace':
        if (overview) toggleOverview(false); else prev();
        break;
      case 'Home': go(0, 0); break;
      case 'End': go(N_MAIN - 1, slides[N_MAIN - 1]._steps); break;
      case 'F5': toggleFS(); break;
      case 'Escape':
        if (body.classList.contains('show-help')) toggleHelp(false);
        else if (overview) toggleOverview(false);
        else if (editing) toggleEdit(false);
        else if (body.classList.contains('is-black')) toggleBlack(false);
        else if (body.classList.contains('show-notes')) toggleNotes(false);
        else handled = false;
        break;
      default: handled = false;
    }
    if (!handled) {
      handled = true;
      switch (code) {
        case 'KeyF': toggleFS(); break;
        case 'KeyB': case 'Period': case 'KeyW': toggleBlack(); break;
        case 'KeyO': case 'KeyG': toggleOverview(); break;
        case 'KeyT': toggleTheme({ x: innerWidth - 60, y: innerHeight - 40 }); break;
        case 'KeyS': case 'KeyP': openPresenter(); break;
        case 'KeyN': toggleNotes(); break;
        case 'KeyH': case 'Slash': toggleHelp(); break;
        case 'KeyE': toggleEdit(); break;
        case 'KeyM': toggleLite(); break;
        default: handled = false;
      }
    }
    if (handled) e.preventDefault();
  }
  document.addEventListener('keydown', onKey);
  addEventListener('hashchange', () => {
    const m = /^#(\d+)(?:\.(\d+))?$/.exec(location.hash);
    if (m) go(+m[1] - 1, m[2] ? +m[2] : 0);
  });

  /* ---------- экран не гаснет во время доклада ---------- */
  let wl = null;
  async function wake() {
    try {
      if ('wakeLock' in navigator && !wl && document.visibilityState === 'visible') {
        wl = await navigator.wakeLock.request('screen');
        wl.addEventListener('release', () => { wl = null; });
      }
    } catch (e) { /* не поддерживается */ }
  }
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && wl === null && timer.running) wake(); });

  /* ---------- печать / PDF: все слайды в финальном состоянии ---------- */
  function preparePrint() {
    slides.forEach((s) => { clearTimeout(s._t); s.classList.add('is-in'); applyStep(s, s._steps); resetCounts(s); });
  }
  function restorePrint() {
    slides.forEach((s, k) => { if (k !== state.i) { s.classList.remove('is-in', 'is-active', 'is-leaving'); applyStep(s, 0); } });
    applyStep(slides[state.i], state.step);
  }
  addEventListener('beforeprint', preparePrint);
  addEventListener('afterprint', restorePrint);

  /* ---------- старт ---------- */
  function start() {
    root.dataset.theme = store.theme === 'dark' ? 'dark' : root.dataset.theme || 'light';
    paintIce();
    fit();
    const m = /^#(\d+)(?:\.(\d+))?$/.exec(location.hash);
    const i = m ? clamp(+m[1] - 1, 0, slides.length - 1) : 0;
    state.i = i;
    go(i, m && m[2] ? +m[2] : 0);
    if (!store.seen) {
      setTimeout(() => toast('<b>F</b> — во весь экран · <b>→</b> или клик — дальше · <b>S</b> — окно докладчика · <b>H</b> — справка', 6500), 1600);
      store.seen = 1;
      persist();
    }
    poke();
  }
  fit();
  const fontsReady = document.fonts && document.fonts.ready
    ? Promise.race([document.fonts.ready, new Promise((r) => setTimeout(r, 1500))])
    : Promise.resolve();
  fontsReady.then(start);

  // для отладки и автотестов
  window.deck = { go, next, prev, state, slides, toggleTheme: () => toggleTheme(), toggleOverview, setTheme };
})();
