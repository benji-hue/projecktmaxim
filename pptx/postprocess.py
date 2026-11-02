#!/usr/bin/env python3
"""Постобработка .pptx после pptxgenjs.

    python3 postprocess.py [light|dark]   →  ../presentation.pptx | ../presentation-dark.pptx

* анимации по щелчку (p:timing) — по реестру out/anim-<theme>.json, привязка к объектам по имени;
* переходы: «Выцветание» на всех слайдах, «Трансформация» (Morph) на слайде 7
  с автоматическим откатом на «Выцветание» в старых версиях PowerPoint;
* график «Россия и Канада»: подписи значений только у ряда «Россия».
"""
from __future__ import annotations

import html
import json
import pathlib
import re
import sys
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent
THEME = sys.argv[1] if len(sys.argv) > 1 else "light"
SRC = ROOT / "out" / f"deck-{THEME}.pptx"
OUT = ROOT.parent / ("presentation.pptx" if THEME == "light" else "presentation-dark.pptx")
SPEC = json.loads((ROOT / "out" / f"anim-{THEME}.json").read_text(encoding="utf-8"))

KINDS = ("sp", "pic", "graphicFrame", "cxnSp")


def element_map(xml: str) -> dict[str, tuple[int, str]]:
    """Имя объекта → (id, тип элемента)."""
    out: dict[str, tuple[int, str]] = {}
    for m in re.finditer(r'<p:cNvPr id="(\d+)" name="([^"]*)"', xml):
        head = xml[: m.start()]
        best, kind = -1, None
        for k in KINDS:
            i = max(head.rfind(f"<p:{k}>"), head.rfind(f"<p:{k} "))
            if i > best:
                best, kind = i, k
        name = html.unescape(m.group(2))
        out[name] = (-1, kind) if name in out else (int(m.group(1)), kind)  # -1 = имя не уникально
    return out


class Ids:
    def __init__(self) -> None:
        self.n = 2

    def __call__(self) -> int:
        self.n += 1
        return self.n


def tgt(spid: int) -> str:
    return f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'


def set_visible(ids: Ids, spid: int) -> str:
    return (f'<p:set><p:cBhvr><p:cTn id="{ids()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            f'{tgt(spid)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:to><p:strVal val="visible"/></p:to></p:set>')


def anim_effect(ids: Ids, spid: int, flt: str, dur: int) -> str:
    return (f'<p:animEffect transition="in" filter="{flt}"><p:cBhvr><p:cTn id="{ids()}" dur="{dur}"/>{tgt(spid)}</p:cBhvr></p:animEffect>')


def anim_prop(ids: Ids, spid: int, attr: str, v0: str, v1: str, dur: int) -> str:
    def val(v: str) -> str:
        return f'<p:fltVal val="{v}"/>' if re.fullmatch(r"-?\d+(\.\d+)?", v) else f'<p:strVal val="{v}"/>'
    return (f'<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="{ids()}" dur="{dur}" fill="hold"/>{tgt(spid)}'
            f'<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:tavLst><p:tav tm="0"><p:val>{val(v0)}</p:val></p:tav><p:tav tm="100000"><p:val>{val(v1)}</p:val></p:tav></p:tavLst></p:anim>')


# эффект → (presetID, presetSubtype, поведение)
def behaviours(fx: str, ids: Ids, spid: int) -> tuple[int, int, str]:
    if fx == "fade":
        return 10, 0, set_visible(ids, spid) + anim_effect(ids, spid, "fade", 500)
    if fx == "float":  # «Всплывание»
        return 42, 0, (set_visible(ids, spid) + anim_effect(ids, spid, "fade", 700)
                       + anim_prop(ids, spid, "ppt_x", "#ppt_x", "#ppt_x", 700)
                       + anim_prop(ids, spid, "ppt_y", "#ppt_y+.04", "#ppt_y", 700))
    if fx == "zoom":  # «Масштабирование»
        return 53, 16, (set_visible(ids, spid)
                        + anim_prop(ids, spid, "ppt_w", "0", "#ppt_w", 500)
                        + anim_prop(ids, spid, "ppt_h", "0", "#ppt_h", 500)
                        + anim_effect(ids, spid, "fade", 500))
    if fx == "wipeRight":  # «Вытеснение» слева направо
        return 22, 8, set_visible(ids, spid) + anim_effect(ids, spid, "wipe(right)", 600)
    if fx == "wipeLeft":  # «Вытеснение» справа налево
        return 22, 2, set_visible(ids, spid) + anim_effect(ids, spid, "wipe(left)", 700)
    if fx == "wheel":  # «Колесо» — рукописная обводка «прорисовывается»
        return 21, 1, set_visible(ids, spid) + anim_effect(ids, spid, "wheel(1)", 1000)
    raise SystemExit(f"неизвестный эффект {fx}")


def timing_xml(steps: list[list[dict]], emap: dict[str, tuple[int, str]], slide_no: int) -> str:
    ids = Ids()
    clicks, built = [], []
    for k, group in enumerate(steps, start=1):
        if not group:
            raise SystemExit(f"слайд {slide_no}: пустой шаг {k}")
        par_outer, par_inner = ids(), ids()
        effects = []
        for j, it in enumerate(group):
            if it["name"] not in emap:
                raise SystemExit(f"слайд {slide_no}: нет объекта «{it['name']}»")
            spid, kind = emap[it["name"]]
            if spid < 0:
                raise SystemExit(f"слайд {slide_no}: имя «{it['name']}» не уникально")
            ctn = ids()
            preset, sub, body = behaviours(it["fx"], ids, spid)
            node = "clickEffect" if j == 0 else "withEffect"
            effects.append(
                f'<p:par><p:cTn id="{ctn}" presetID="{preset}" presetClass="entr" presetSubtype="{sub}" fill="hold"'
                f' grpId="0" nodeType="{node}">'
                f'<p:stCondLst><p:cond delay="{int(it.get("delay", 0))}"/></p:stCondLst><p:childTnLst>{body}</p:childTnLst></p:cTn></p:par>')
            if kind == "sp":
                built.append(f'<p:bldP spid="{spid}" grpId="0" animBg="1"/>')
        clicks.append(
            f'<p:par><p:cTn id="{par_outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
            f'<p:par><p:cTn id="{par_inner}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            + "".join(effects) + "</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>")
    bld = f"<p:bldLst>{''.join(built)}</p:bldLst>" if built else ""
    return ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            + "".join(clicks) +
            '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
            f'</p:childTnLst></p:cTn></p:par></p:tnLst>{bld}</p:timing>')


FADE = '<p:transition spd="med"><p:fade/></p:transition>'
MORPH = ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
         '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
         '<p:transition spd="slow"><p159:morph option="byObject"/></p:transition></mc:Choice>'
         f'<mc:Fallback>{FADE}</mc:Fallback></mc:AlternateContent>')


def patch_slide(xml: str, slide_no: int) -> str:
    anchor = "</p:clrMapOvr>"
    if xml.count(anchor) != 1 or "<p:timing>" in xml or "<p:transition" in xml:
        raise SystemExit(f"слайд {slide_no}: неожиданная структура XML")
    trans = MORPH if SPEC["trans"].get(str(slide_no)) == "morph" else FADE
    steps = SPEC["anim"].get(str(slide_no))
    timing = timing_xml(steps, element_map(xml), slide_no) if steps else ""
    return xml.replace(anchor, anchor + trans + timing)


def patch_chart(xml: str) -> str:
    """Подписи значений только у ряда «Россия» (у «Канады» — фон для сравнения)."""
    if "<c:v>Россия</c:v>" not in xml:
        return xml

    def fix(m: re.Match) -> str:
        ser = m.group(0)
        if re.search(r"<c:tx>.*?<c:v>Канада</c:v>", ser, flags=re.S):
            ser = ser.replace('<c:showVal val="1"/>', '<c:showVal val="0"/>')
        return ser

    return re.sub(r"<c:ser>.*?</c:ser>", fix, xml, flags=re.S)


def main() -> None:
    zin = zipfile.ZipFile(SRC)
    names = zin.namelist()
    order = ["[Content_Types].xml"] + [n for n in names if n != "[Content_Types].xml"]
    tmp = OUT.with_suffix(".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in order:
            data = zin.read(n)
            m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", n)
            if m:
                data = patch_slide(data.decode("utf-8"), int(m.group(1))).encode("utf-8")
            elif re.fullmatch(r"ppt/charts/chart\d+\.xml", n):
                data = patch_chart(data.decode("utf-8")).encode("utf-8")
            zout.writestr(zin.getinfo(n), data)
    tmp.replace(OUT)
    clicks = sum(len(v) for v in SPEC["anim"].values())
    print(f"OK → {OUT.name}: анимаций-щелчков {clicks}, слайдов с анимацией {len(SPEC['anim'])}")


if __name__ == "__main__":
    main()
