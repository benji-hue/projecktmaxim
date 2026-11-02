#!/usr/bin/env python3
"""Сборка презентации в один самодостаточный HTML-файл.

    python3 src/build.py            ->  presentation.html

Что делает:
  * разворачивает макросы <x-...> в разметку/SVG (лента месяцев, графики по данным,
    каток по разметке НХЛ, джерси, программка 1984, иллюстрации);
  * расставляет неразрывные пробелы по правилам русской типографики;
  * вшивает шрифты (woff2, base64), стили и скрипт — файл работает без интернета.
"""
from __future__ import annotations

import base64
import html
import pathlib
import re

SRC = pathlib.Path(__file__).resolve().parent
OUT = SRC.parent / "presentation.html"

MONTHS = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]
MONTHS_FULL = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август",
               "сентябрь", "октябрь", "ноябрь", "декабрь"]

# ---------------------------------------------------------------- данные (НХЛ 2024/25)
# доли в процентах по кварталам I..IV, n, χ², p — из таблицы автора, пересчитано
DATA = [
    # группа, n, [I, II, III, IV] %, counts, chi2, p-текст
    ("Вся лига", 853, [33.3, 25.1, 21.9, 19.7], [284, 214, 187, 168], 39.3, "< 0,001"),
    ("Канада", 342, [34.5, 28.1, 22.5, 14.9], [118, 96, 77, 51], 29.9, "< 0,001"),
    ("США", 259, [35.1, 22.4, 20.1, 22.4], [91, 58, 52, 58], 15.5, "0,001"),
    ("Швеция", 80, [26.2, 20.0, 27.5, 26.2], [21, 16, 22, 21], 1.1, "0,79"),
    ("Россия", 52, [23.1, 21.2, 21.2, 34.6], [12, 11, 11, 18], 2.5, "0,47"),
]
QN = ["I", "II", "III", "IV"]
QM = ["янв–мар", "апр–июн", "июл–сен", "окт–дек"]


def ru(x: float, d: int = 1) -> str:
    return f"{x:.{d}f}".replace(".", ",")


# ---------------------------------------------------------------- иконки (stroke 24×24)
ICONS = {
    # «рука» — по мотивам Lucide (ISC)
    "hand": '<path d="M18 11V6a2 2 0 0 0-2-2a2 2 0 0 0-2 2"/><path d="M14 10V4a2 2 0 0 0-2-2a2 2 0 0 0-2 2v2"/>'
            '<path d="M10 10.5V6a2 2 0 0 0-2-2a2 2 0 0 0-2 2v8"/>'
            '<path d="M18 8a2 2 0 1 1 4 0v6a8 8 0 0 1-8 8h-2c-2.8 0-4.5-.86-5.99-2.34l-3.6-3.6a2 2 0 0 1 2.83-2.82L7 15"/>',
    "check": '<path d="M20 6 9 17l-5-5"/>',
}


def icon(name: str) -> str:
    return f'<svg class="ico" viewBox="0 0 24 24" aria-hidden="true">{ICONS[name]}</svg>'


# ---------------------------------------------------------------- лента месяцев
def months(attrs: dict) -> str:
    start = int(attrs.get("start", 1))
    mini = "mini" in attrs
    hl = int(attrs.get("hl", 0))
    ends = "ends" in attrs
    vals = [float(v) for v in attrs["vals"].split(",")] if "vals" in attrs else None
    cls = "months" + (" months--mini" if mini else "")
    cells = []
    for k in range(12):
        m = (start - 1 + k) % 12 + 1
        q = (m - 1) // 3 + 1
        style = f"--k:{k}" + (f";--v:{vals[q - 1]}" if vals else "")
        c = "m"
        if hl and m == hl:
            c += " is-hl"
        if ends and k == 0:
            c += " is-b"
        if ends and k == 11:
            c += " is-o"
        label = "" if mini else MONTHS[m - 1]
        cells.append(f'<div class="{c}" data-m="{m}" data-q="{q}" style="{style}"><span>{label}</span></div>')
    return f'<div class="{cls}" aria-hidden="true">{"".join(cells)}</div>'


def picker(_attrs: dict) -> str:
    out = []
    for m in range(1, 13):
        ans = " is-answer" if m == 12 else ""
        day = '<span class="d">30</span>' if m == 12 else ""
        out.append(
            f'<button type="button" class="pk{ans}" data-m="{m}" aria-label="{MONTHS_FULL[m - 1]}">'
            f'<span class="g">{icon("hand")}</span>{MONTHS[m - 1]}{day}</button>'
        )
    return "".join(out)


# ---------------------------------------------------------------- каток (разметка НХЛ, 1 фут = 10 ед.)
def circ(cx, cy, r):
    return f"M{cx - r},{cy} a{r},{r} 0 1,0 {2 * r},0 a{r},{r} 0 1,0 {-2 * r},0"


def rink(_attrs: dict) -> str:
    W, H, R = 2000, 850, 280
    board = f"M{R},0 H{W - R} A{R},{R} 0 0 1 {W},{R} V{H - R} A{R},{R} 0 0 1 {W - R},{H} H{R} A{R},{R} 0 0 1 0,{H - R} V{R} A{R},{R} 0 0 1 {R},0 Z"
    p = []
    d = 0

    def path(cls, dd, draw=True):
        nonlocal d
        d += 1
        pl = ' pathLength="1"' if draw else ""
        p.append(f'<path class="{cls}" d="{dd}"{pl} style="--d:{d}"/>')

    p.append(f'<path class="rk-ice rk-fill" d="{board}"/>')
    path("rk-board", board)
    # центральная красная и синие линии
    path("rk-red-w", f"M1000,0 V{H}")
    path("rk-blue-w", f"M750,0 V{H}")
    path("rk-blue-w", f"M1250,0 V{H}")
    # линии ворот (обрезаны скруглением бортов)
    gy = 57.5
    path("rk-red", f"M110,{gy} V{H - gy}")
    path("rk-red", f"M1890,{gy} V{H - gy}")
    # центральный круг
    path("rk-blue", circ(1000, 425, 150))
    # круги вбрасывания
    for cx, cy in [(310, 205), (310, 645), (1690, 205), (1690, 645)]:
        path("rk-red", circ(cx, cy, 150))
        # «хэш-марки»
        for sx in (-1, 1):
            for sy in (-1, 1):
                x = cx + sx * 30
                y0 = cy + sy * 148
                path("rk-red", f"M{x},{y0} V{y0 + sy * 22}")
    # вратарские площадки и ворота
    p.append('<path class="rk-crease rk-fill" d="M110,365 A60,60 0 0 1 110,485 Z"/>')
    p.append('<path class="rk-crease rk-fill" d="M1890,365 A60,60 0 0 0 1890,485 Z"/>')
    path("rk-red", "M110,395 H76 Q66,395 66,405 V445 Q66,455 76,455 H110")
    path("rk-red", "M1890,395 H1924 Q1934,395 1934,405 V445 Q1934,455 1924,455 H1890")
    # судейская зона
    path("rk-red", f"M900,{H} A100,100 0 0 1 1100,{H}")
    # точки
    dots = ['<circle class="rk-dot-b rk-fill" cx="1000" cy="425" r="7"/>']
    for cx, cy in [(310, 205), (310, 645), (1690, 205), (1690, 645), (800, 205), (800, 645), (1200, 205), (1200, 645)]:
        dots.append(f'<circle class="rk-dot-r rk-fill" cx="{cx}" cy="{cy}" r="11"/>')
    defs = (
        '<defs><linearGradient id="rkSheen" x1="0" y1="0" x2="1" y2="1">'
        '<stop offset="0" style="stop-color:var(--ice-a)"/><stop offset=".55" style="stop-color:var(--ice-b)"/>'
        '<stop offset="1" style="stop-color:var(--ice-a)"/></linearGradient></defs>'
    )
    return (f'<svg class="rink" viewBox="-12 -12 {W + 24} {H + 24}" preserveAspectRatio="xMidYMid meet" aria-hidden="true">'
            f'{defs}{"".join(p)}{"".join(dots)}</svg>')


# ---------------------------------------------------------------- джерси «Шестёркин 31» (вид со спины)
def jersey(_attrs: dict) -> str:
    body = ("M228,60 C262,76 338,76 372,60 L452,92 L578,300 L520,352 L446,252 L462,640 "
            "Q300,664 138,640 L154,252 L80,352 L22,300 L148,92 Z")

    def band(o1, o2, mirror=False):
        ax, ay = 0.545, -0.839
        outer, inner = (22, 300), (80, 352)
        pts = [
            (outer[0] + ax * o1, outer[1] + ay * o1), (inner[0] + ax * o1, inner[1] + ay * o1),
            (inner[0] + ax * o2, inner[1] + ay * o2), (outer[0] + ax * o2, outer[1] + ay * o2),
        ]
        if mirror:
            pts = [(600 - x, y) for x, y in pts]
        return "M" + " L".join(f"{x:.1f},{y:.1f}" for x, y in pts) + " Z"

    stripes = "".join([
        f'<path class="j-st1" d="{band(22, 50)}"/>', f'<path class="j-st2" d="{band(58, 70)}"/>',
        f'<path class="j-st1" d="{band(22, 50, True)}"/>', f'<path class="j-st2" d="{band(58, 70, True)}"/>',
        '<rect class="j-st1" x="0" y="548" width="600" height="30"/>',
        '<rect class="j-st2" x="0" y="588" width="600" height="13"/>',
    ])
    return (
        '<svg class="jersey" viewBox="0 0 600 700" aria-hidden="true">'
        '<defs><linearGradient id="jGrad" x1="0" y1="0" x2="0" y2="1">'
        '<stop offset="0" style="stop-color:var(--jersey-a)"/><stop offset="1" style="stop-color:var(--jersey-b)"/></linearGradient>'
        f'<clipPath id="jClip"><path d="{body}"/></clipPath></defs>'
        f'<path class="j-body" d="{body}"/>'
        f'<g clip-path="url(#jClip)">{stripes}</g>'
        '<path class="j-seam" d="M154,252 C200,240 400,240 446,252"/>'
        '<path id="jArc" d="M150,206 Q300,150 450,206" fill="none"/>'
        '<text class="j-name"><textPath href="#jArc" startOffset="50%" text-anchor="middle">ШЕСТЁРКИН</textPath></text>'
        '<text class="j-num" x="300" y="478" text-anchor="middle">31</text>'
        '</svg>'
    )


# ---------------------------------------------------------------- программка 1984 (иллюстрация)
def program(_attrs: dict) -> str:
    # 16 игроков; распределение ≈ как у Барнсли: ~40 % I кв. … ~10 % IV кв.
    roster = [(4, 1), (7, 2), (8, 1), (10, 5), (11, 3), (12, 2), (14, 4), (16, 1),
              (17, 7), (18, 3), (19, 12), (21, 2), (22, 8), (24, 6), (26, 9), (27, 10)]
    head = '<div class="pg-head"><span>№</span>' + "".join(f"<span>{m[0]}</span>" for m in MONTHS) + "</div>"
    rows = []
    for num, mon in roster:
        cells = "".join(f'<i class="on"></i>' if k + 1 == mon else "<i></i>" for k in range(12))
        rows.append(f'<div class="pg-row"><span>{num:02d}</span>{cells}</div>')
    circle = ('<svg class="pg-circle" viewBox="0 0 132 459" preserveAspectRatio="none" aria-hidden="true">'
              '<path pathLength="1" '
              'd="M72,6 C113,5 127,72 126,232 C125,392 111,454 64,453 C19,452 6,386 6,228 C6,76 23,9 75,9 C93,10 106,18 115,32"/></svg>')
    return (
        '<div class="program" aria-hidden="true">'
        '<div class="pg-top"><span>Юниорский хоккей</span><span>1984</span></div>'
        '<div class="pg-title">Программа<br>матча</div>'
        '<div class="pg-sub">Летбридж · Альберта · Канада</div>'
        '<div class="pg-rule"></div>'
        f'<div class="pg-roster">{head}{"".join(rows)}{circle}</div>'
        '<div class="pg-foot"><span>Состав · месяц рождения</span><span>иллюстрация</span></div>'
        '</div>'
    )


# ---------------------------------------------------------------- дети (слайд 3)
def kid(cx, ground, h, cls, stick=True):
    r = round(h * 0.128)
    bw = round(h * 0.35)
    by = ground - h + 2 * r + 6
    s = (f'<circle class="{cls}" cx="{cx}" cy="{ground - h + r}" r="{r}"/>'
         f'<rect class="{cls}" x="{cx - bw / 2}" y="{by}" width="{bw}" height="{ground - by}" rx="{bw / 2}"/>')
    if stick:
        x0, y0 = cx + bw * 0.42, by + (ground - by) * 0.32
        x1, y1 = cx + bw * 1.05, ground - 4
        s += f'<path class="kid-stick" d="M{x0:.0f},{y0:.0f} L{x1:.0f},{y1:.0f} H{x1 + bw * 0.55:.0f}"/>'
    return s


def kids(_attrs: dict) -> str:
    return (
        '<svg viewBox="0 0 644 340" aria-hidden="true">'
        '<path class="kid-ground" d="M20,300 H624"/>'
        + kid(190, 300, 250, "kid-b") + kid(468, 300, 192, "kid-o") +
        '<text class="kid-txt" x="190" y="336" text-anchor="middle">1 января</text>'
        '<text class="kid-txt" x="468" y="336" text-anchor="middle">31 декабря</text>'
        '</svg>'
    )


# ---------------------------------------------------------------- петля «снежного кома» (слайд 3)
def loop(_attrs: dict) -> str:
    y = 262
    return (
        '<svg class="arcs" viewBox="0 0 944 600" aria-hidden="true">'
        f'<path class="track" d="M804,172 C804,30 140,30 140,172"/>'
        f'<path class="arc arc-12" pathLength="1" d="M284,{y} H322"/>'
        f'<path class="arrowhead ah-12" d="M316,{y - 9} L330,{y} L316,{y + 9} Z"/>'
        f'<path class="arc arc-23" pathLength="1" d="M616,{y} H654"/>'
        f'<path class="arrowhead ah-23" d="M648,{y - 9} L662,{y} L648,{y + 9} Z"/>'
        '<path class="arc arc-31" pathLength="1" d="M804,174 C804,30 140,30 140,166"/>'
        '<path class="arrowhead ah-31" d="M131,160 L140,176 L149,160 Z"/>'
        '<text class="exit-lbl" x="472" y="122" text-anchor="middle" style="fill:var(--ink-3)">по кругу — разрыв растёт</text>'
        '<path class="arc arc-exit" pathLength="1" d="M140,352 V378"/>'
        '<path class="arrowhead arrowhead-o ah-exit" d="M131,374 L140,390 L149,374 Z"/>'
        '<g class="ball"><circle r="12">'
        '<animateMotion dur="3.4s" repeatCount="indefinite" path="M804,172 C804,30 140,30 140,172" keyPoints="0;1" keyTimes="0;1" calcMode="spline" keySplines=".5 0 .5 1"/>'
        '<animate attributeName="r" values="7;20;20" keyTimes="0;.85;1" dur="3.4s" repeatCount="indefinite"/>'
        '<animate attributeName="opacity" values="0;1;1;0" keyTimes="0;.12;.88;1" dur="3.4s" repeatCount="indefinite"/>'
        '</circle></g>'
        '</svg>'
    )


# ---------------------------------------------------------------- графики
def bar_path(x, y, w, h, r):
    r = min(r, w / 2, h)
    return (f"M{x:.1f},{y + h:.1f} V{y + r:.1f} A{r},{r} 0 0 1 {x + r:.1f},{y:.1f} "
            f"H{x + w - r:.1f} A{r},{r} 0 0 1 {x + w:.1f},{y + r:.1f} V{y + h:.1f} Z")


def chart(attrs: dict) -> str:
    name = attrs["name"]
    ca = DATA[1]
    rus = DATA[4]
    if name == "ca":
        W, H, x0, x1, base, top = 760, 600, 64, 750, 500, 30
    else:
        W, H, x0, x1, base, top = 760, 398, 64, 750, 326, 14
    k = (base - top) / 40
    band = (x1 - x0) / 4
    g = []
    for v in (10, 20, 30, 40):
        yy = base - v * k
        g.append(f'<line class="gl" x1="{x0}" y1="{yy:.1f}" x2="{x1}" y2="{yy:.1f}"/>')
        g.append(f'<text class="ytick" x="{x0 - 14}" y="{yy + 6:.1f}" text-anchor="end">{v}{" %" if v == 40 else ""}</text>')
    g.append(f'<text class="ytick" x="{x0 - 14}" y="{base + 6}" text-anchor="end">0</text>')
    ye = base - 25 * k
    g.append(f'<line class="exp" x1="{x0}" y1="{ye:.1f}" x2="{x1}" y2="{ye:.1f}"/>')
    bars, labels, hits = [], [], []
    bi = 0
    for q in range(4):
        c = x0 + band * (q + 0.5)
        if name == "ca":
            v = ca[2][q]
            w = 104
            h = v * k
            x = c - w / 2
            yb = base - h
            bars.append(f'<path class="bar bar-ca q{q + 1}" style="--b:{bi}" d="{bar_path(x, yb, w, h, 12)}"/>')
            cls = "val q%d" % (q + 1) + ("" if q in (0, 3) else " val-s")
            labels.append(f'<text class="{cls}" style="--b:{bi}" x="{c:.1f}" y="{yb - 16:.1f}" text-anchor="middle">{ru(v)} %</text>')
            hits.append(f'<rect class="hit" x="{c - band / 2 + 6:.1f}" y="{top}" width="{band - 12:.1f}" height="{base - top}" '
                        f'data-tip="{QN[q]} квартал ({QM[q]}) · Канада: {ru(v)} % — {ca[3][q]} из {ca[1]}"/>')
            bi += 1
        else:
            w, gap = 66, 8
            for j, (row, cl) in enumerate([(rus, "bar-ru"), (ca, "bar-ctx")]):
                v = row[2][q]
                h = v * k
                x = c - (2 * w + gap) / 2 + j * (w + gap)
                yb = base - h
                bars.append(f'<path class="bar {cl} q{q + 1}" style="--b:{bi}" d="{bar_path(x, yb, w, h, 10)}"/>')
                if j == 0 and q in (0, 3):
                    labels.append(f'<text class="val" style="--b:{bi};font-size:26px" x="{x + w / 2:.1f}" y="{yb - 14:.1f}" text-anchor="middle">{ru(v)}</text>')
                hits.append(f'<rect class="hit" x="{x:.1f}" y="{top}" width="{w:.1f}" height="{base - top}" '
                            f'data-tip="{QN[q]} квартал ({QM[q]}) · {row[0]}: {ru(v)} % — {row[3][q]} из {row[1]}"/>')
                bi += 1
        dy = 38 if name == "ca" else 34
        labels.append(f'<text class="xq" x="{c:.1f}" y="{base + dy}" text-anchor="middle">{QN[q]} кв.</text>')
        labels.append(f'<text class="xm" x="{c:.1f}" y="{base + dy + 25}" text-anchor="middle">{QM[q]}</text>')
    g.append(f'<line class="base" x1="{x0}" y1="{base}" x2="{x1}" y2="{base}"/>')
    extra = ""
    if name == "ca":
        c1 = x0 + band * 0.5
        c4 = x0 + band * 3.5
        t1 = base - ca[2][0] * k
        t4 = base - ca[2][3] * k
        extra = (f'<path class="bracket" pathLength="1" d="M{c1 + 60:.1f},{t1:.1f} H{c4:.1f} V{t4 - 52:.1f}"/>'
                 f'<path class="bracket" pathLength="1" d="M{c4 - 9:.1f},{t4 - 63:.1f} L{c4:.1f},{t4 - 50:.1f} L{c4 + 9:.1f},{t4 - 63:.1f}"/>')
    title = "Канадцы в НХЛ 2024/25 по кварталу рождения" if name == "ca" else "Россияне и канадцы в НХЛ 2024/25 по кварталу рождения"
    return (f'<svg class="chart chart-{name}" viewBox="0 0 {W} {H}" role="img" aria-label="{title}">'
            f'{"".join(g)}{"".join(bars)}{extra}{"".join(labels)}{"".join(hits)}</svg>')


# ---------------------------------------------------------------- таблица (приложение)
def table(_attrs: dict) -> str:
    rows = []
    for name, n, pc, cnt, chi, p in DATA:
        ratio = pc[0] / pc[3]
        sig = '<span class="sig y">значимо</span>' if not p.startswith("0,4") and not p.startswith("0,7") else '<span class="sig n">не значимо</span>'
        seg = "".join(f'<i style="flex:{v}"></i>' for v in pc)
        hl = ' class="hl"' if name == "Канада" else (' class="hl-o"' if name == "Россия" else "")
        rows.append(
            f'<tr{hl}><td class="g">{name}<small>n = {n}</small></td>'
            + "".join(f"<td>{ru(v)}</td>" for v in pc)
            + f"<td>{ru(ratio, 2)}</td><td>{ru(chi)}</td><td class=\"p\">{p}{sig}</td>"
            f'<td class="dist"><div class="segbar" title="I–IV кварталы">{seg}</div></td></tr>'
        )
    head = ("<thead><tr><th>Группа</th><th>I кв., %</th><th>II кв., %</th><th>III кв., %</th>"
            "<th>IV кв., %</th><th>I / IV</th><th>χ²(3)</th><th>p</th><th>I → IV</th></tr></thead>")
    return f'<table class="dtable">{head}<tbody>{"".join(rows)}</tbody></table>'


# ---------------------------------------------------------------- иллюстрации «что делать»
def fix(attrs: dict) -> str:
    name = attrs["name"]
    if name == "bio":
        s = ['<svg viewBox="0 0 520 270" aria-hidden="true">',
             '<rect class="band" x="18" y="22" width="232" height="226" rx="22"/>',
             '<rect class="band" x="270" y="22" width="232" height="226" rx="22"/>',
             '<text class="band-l" x="134" y="54" text-anchor="middle">85–90 % роста</text>',
             '<text class="band-l" x="386" y="54" text-anchor="middle">90–95 % роста</text>']
        for cx, h, c in [(62, 118, "b"), (134, 124, "o"), (206, 112, "o"), (314, 158, "o"), (386, 168, "b"), (458, 150, "b")]:
            s.append(kid(cx, 236, h, f"kid {c}", stick=False))
        s.append("</svg>")
        return "".join(s)
    if name == "num":
        s = ['<svg viewBox="0 0 520 270" aria-hidden="true">']
        xs = [62, 161, 260, 359, 458]
        for i, x in enumerate(xs):
            c = " b" if i == 0 else (" o" if i == 4 else "")
            sc = 1 - i * 0.04
            w = 82 * sc
            top = 70 + i * 6
            d = (f"M{x - w * .32:.1f},{top} L{x - w * .62:.1f},{top + 16} L{x - w * .74:.1f},{top + 50} L{x - w * .48:.1f},{top + 58} "
                 f"L{x - w * .46:.1f},{top + 128 * sc} L{x + w * .46:.1f},{top + 128 * sc} L{x + w * .48:.1f},{top + 58} L{x + w * .74:.1f},{top + 50} "
                 f"L{x + w * .62:.1f},{top + 16} L{x + w * .32:.1f},{top} Q{x:.1f},{top + 16} {x - w * .32:.1f},{top} Z")
            s.append(f'<path class="shirt{c}" d="{d}"/>')
            s.append(f'<text class="shirt-n{" w" if c else ""}" x="{x}" y="{top + 92 * sc:.1f}">{i + 1}</text>')
        s.append('<path class="age-ax" d="M40,236 H480"/>')
        s.append('<path class="age-ax" d="M470,228 L482,236 L470,244" fill="none"/>')
        s.append('<text class="age-l" x="40" y="262">старше</text><text class="age-l" x="480" y="262" text-anchor="end">младше</text>')
        s.append("</svg>")
        return "".join(s)
    if name == "self":
        return (
            '<svg viewBox="0 0 520 270" aria-hidden="true">'
            '<path class="age-ax" d="M50,226 H486"/><path class="age-ax" d="M50,226 V36"/>'
            '<path class="peer-line" d="M60,92 C180,86 320,82 476,76"/>'
            '<text class="age-l" x="476" y="64" text-anchor="end">ровесники постарше</text>'
            '<path class="self-line" d="M60,196 C150,186 210,160 270,140 C340,116 400,104 470,92"/>'
            '<circle class="self-dot" cx="60" cy="196" r="9"/><circle class="self-dot" cx="270" cy="140" r="9"/><circle class="self-dot" cx="470" cy="92" r="11"/>'
            '<text class="age-l" x="300" y="172">сам ребёнок: рост</text>'
            '<text class="age-l" x="50" y="256">сезон 1</text><text class="age-l" x="486" y="256" text-anchor="end">сезон 3</text>'
            '</svg>'
        )
    raise ValueError(name)


MACROS = {
    "x-months": months, "x-picker": picker, "x-rink": rink, "x-jersey": jersey, "x-program": program,
    "x-kids": kids, "x-loop": loop, "x-chart": chart, "x-table": table, "x-fix": fix,
    "x-icon": lambda a: icon(a["name"]),
}


def parse_attrs(s: str) -> dict:
    out = {}
    for m in re.finditer(r'([\w-]+)(?:="([^"]*)")?', s):
        out[m.group(1)] = m.group(2) if m.group(2) is not None else True
    return out


def expand(markup: str) -> str:
    def rep(m):
        return MACROS[m.group(1)](parse_attrs(m.group(2) or ""))
    return re.sub(r"<(x-[\w-]+)([^>]*)></\1>", rep, markup)


# ---------------------------------------------------------------- русская типографика
NB = " "
SHORT = (r"в|во|с|со|к|ко|у|о|об|от|до|на|за|из|по|не|ни|но|и|а|я|мы|он|она|они|без|для|при|про|под|над|"
         r"что|как|или|это|их|его|её|все|всё|так|кто|чем|уже|ещё|же|бы|ли|то|те|та|тот|эти|этот|эта")
SHORT_RE = re.compile(r"(?<![\w\-])(" + SHORT + r") (?=\S)", re.IGNORECASE)


def typo(text: str) -> str:
    if not text.strip():
        return text
    for _ in range(2):
        text = SHORT_RE.sub(lambda m: m.group(1) + NB, text)
    text = re.sub(r" (—|–)(?=\s)", NB + r"\1", text)                  # тире не в начале строки
    text = re.sub(r"(\d) (?=[%‰×]|[а-яёa-z«(])", r"\1" + NB, text, flags=re.I)  # число + слово/знак
    text = re.sub(r"(\d) (?=\d{3}\b)", r"\1" + NB, text)              # 10 000
    text = re.sub(r"([№§]) ", r"\1" + NB, text)
    text = re.sub(r" (ли|же|бы)(?=[\s.,!?:;»)]|$)", NB + r"\1", text)
    text = re.sub(r"(\b[IV]+) (?=кв)", r"\1" + NB, text)
    return text


def typograph(markup: str) -> str:
    parts = re.split(r"(<[^>]+>)", markup)
    skip = 0
    for i, part in enumerate(parts):
        if part.startswith("<"):
            tag = part[1:].split()[0].lower() if len(part) > 1 else ""
            if tag in ("script", "style"):
                skip += 1
            elif tag in ("/script", "/style"):
                skip -= 1
            continue
        if not skip:
            parts[i] = typo(part)
    return "".join(parts)


# ---------------------------------------------------------------- шрифты
FONT_FACES = [
    # семейство, файл, вес — порядок: latin последним (проверяется первым при пересечении диапазонов)
    ("Unbounded", "Unbounded-cyrillic", "200 900"),
    ("Unbounded", "Unbounded-latin", "200 900"),
    ("Onest", "Onest-cyrillic", "100 900"),
    ("Onest", "Onest-math", "100 900"),
    ("Onest", "Onest-latin", "100 900"),
    ("JetBrains Mono", "JetBrainsMono-cyrillic", "100 800"),
    ("JetBrains Mono", "JetBrainsMono-greek", "100 800"),
    ("JetBrains Mono", "JetBrainsMono-latin", "100 800"),
]


def fonts_css() -> str:
    out = []
    for fam, file, weight in FONT_FACES:
        data = base64.b64encode((SRC / "fonts" / f"{file}.woff2").read_bytes()).decode()
        rng = (SRC / "fonts" / f"{file}.woff2.range").read_text().strip()
        out.append(
            f"@font-face{{font-family:'{fam}';font-style:normal;font-weight:{weight};font-display:block;"
            f"src:url(data:font/woff2;base64,{data}) format('woff2');unicode-range:{rng};}}"
        )
    return "\n".join(out)


# ---------------------------------------------------------------- сборка
def main() -> None:
    css = (SRC / "styles.css").read_text(encoding="utf-8").replace("/*@FONTS@*/", fonts_css())
    js = (SRC / "engine.js").read_text(encoding="utf-8")
    slides = (SRC / "slides.html").read_text(encoding="utf-8")
    slides = re.sub(r"<!--.*?-->", "", slides, flags=re.S)
    slides = typograph(expand(slides))
    shell = (SRC / "shell.html").read_text(encoding="utf-8")
    shell = typograph(shell.replace("<!--@SLIDES@-->", "\x00SLIDES\x00"))
    page = (shell.replace("/*@CSS@*/", css)
                 .replace("/*@JS@*/", js)
                 .replace("\x00SLIDES\x00", slides))
    OUT.write_text(page, encoding="utf-8")
    print(f"OK → {OUT.relative_to(SRC.parent)}  ({OUT.stat().st_size / 1024:.0f} КБ)")


if __name__ == "__main__":
    main()
