#!/usr/bin/env python3
"""Проверка вёрстки .pptx без PowerPoint: влезает ли текст в фигуры, нет ли наложений и выхода за слайд.

    python3 check_fit.py ../presentation-simple.pptx

Ширина текста меряется шрифтом Liberation Sans — он метрически совпадает с Arial.
Высота строки = 1,2 × кегль × межстрочный интервал (как у Arial в PowerPoint), запас 3 %.
"""
import sys
from itertools import combinations

from PIL import ImageFont
from pptx import Presentation
from pptx.util import Emu

FONTS = {False: "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
         True: "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"}
_cache = {}


def font(size_pt, bold):
    key = (round(size_pt * 4), bold)
    if key not in _cache:
        _cache[key] = ImageFont.truetype(FONTS[bold], size=max(1, round(size_pt * 4)))  # 1 pt = 4 ед.
    return _cache[key]


def width_pt(txt, size, bold):
    return font(size, bold).getlength(txt) / 4


def para_height(par, avail_w, default_size):
    """Перенос по словам с учётом разных кеглей внутри абзаца."""
    runs = [(r.text, (r.font.size.pt if r.font.size else default_size), bool(r.font.bold)) for r in par.runs]
    if not runs or all(not t for t, _, _ in runs):
        sz = runs[0][1] if runs else default_size
        return sz * 1.2, 0
    words = []
    for t, sz, b in runs:
        for i, w in enumerate(t.split(" ")):
            words.append((("" if i == 0 else " ") + w, sz, b))
    lines, cur_w, cur_max, line_sizes, widest = 0, 0.0, 0, [], 0.0
    for w, sz, b in words:
        ww = width_pt(w, sz, b)
        widest = max(widest, width_pt(w.strip(), sz, b))
        if cur_w + ww > avail_w and cur_w > 0:
            line_sizes.append(cur_max)
            cur_w, cur_max = width_pt(w.lstrip(), sz, b), sz
        else:
            cur_w += ww
            cur_max = max(cur_max, sz)
    line_sizes.append(cur_max)
    ls = par.line_spacing if isinstance(par.line_spacing, float) else 1.0
    after = par.space_after.pt if par.space_after is not None else 0
    return sum(s * 1.2 * ls for s in line_sizes) + after, widest


def main(path):
    prs = Presentation(path)
    SW, SH = prs.slide_width, prs.slide_height
    problems = 0
    for no, slide in enumerate(prs.slides, 1):
        boxes = []
        for sh in slide.shapes:
            x, y, w, h = sh.left, sh.top, sh.width, sh.height
            if x < 0 or y < 0 or x + w > SW + 5000 or y + h > SH + 5000:
                print(f"слайд {no}: «{sh.name}» выходит за край слайда"); problems += 1
            if sh.has_text_frame and sh.text_frame.text.strip():
                tf = sh.text_frame
                ml, mr = (tf.margin_left or 0), (tf.margin_right or 0)
                mt, mb = (tf.margin_top or 0), (tf.margin_bottom or 0)
                avail_w = Emu(w - ml - mr).pt
                avail_h = Emu(h - mt - mb).pt
                need, widest = 0.0, 0.0
                for p in tf.paragraphs:
                    ph, pw = para_height(p, avail_w, 18)
                    need += ph
                    widest = max(widest, pw)
                if need > avail_h * 1.0 + 1:
                    print(f"слайд {no}: «{sh.name}» — текст не влезает по высоте: нужно {need:.0f} pt, есть {avail_h:.0f} pt"); problems += 1
                if widest > avail_w + 1:
                    print(f"слайд {no}: «{sh.name}» — слово шире фигуры ({widest:.0f} > {avail_w:.0f} pt)"); problems += 1
                boxes.append((sh.name, x, y, x + w, y + h))
        # наложения фигур с текстом друг на друга (кроме намеренных цветных «ответов» поверх месяцев)
        for a, b in combinations(boxes, 2):
            if "Ответ декабрь" in (a[0], b[0]):
                continue
            ix = min(a[3], b[3]) - max(a[1], b[1])
            iy = min(a[4], b[4]) - max(a[2], b[2])
            if ix > 20000 and iy > 20000:
                print(f"слайд {no}: наложение «{a[0]}» и «{b[0]}»"); problems += 1
    print("ПРОБЛЕМ:", problems)
    return problems


if __name__ == "__main__":
    sys.exit(1 if main(sys.argv[1]) else 0)
