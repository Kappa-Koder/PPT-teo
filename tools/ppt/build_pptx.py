#!/usr/bin/env python3
"""Costruisce la versione PowerPoint con testi modificabili.

Ogni slide ha come sfondo il video (senza testi) registrato da record.js; sopra ci sono
caselle di testo vere, con il carattere Calibri, che compaiono in dissolvenza come
nell'HTML. Nel quiz le risposte compaiono a ogni clic.

Uso: python3 tools/ppt/build_pptx.py <cartella-registrazioni> <uscita.pptx>
"""
import json
import pathlib
import re
import subprocess
import sys

from lxml import etree
from PIL import ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

PX = 6350                      # EMU per pixel: 1920 px = 13,333 pollici
FONT = "Calibri"
BG = RGBColor(0x0A, 0x0C, 0x17)
CARLITO = "/usr/share/fonts/truetype/crosextra/Carlito-%s.ttf"
BANDS = [(0, "F8EED9"), (24, "F8EED9"), (24, "E8B97F"), (33, "E8B97F"), (33, "F4E3C6"), (45, "F4E3C6"),
         (45, "C97A45"), (55, "C97A45"), (55, "F2DDB9"), (64, "F2DDB9"), (64, "E36A42"), (71, "E36A42"),
         (71, "EBC38E"), (84, "EBC38E"), (84, "F7EBD3"), (100, "F7EBD3")]
NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' \
     'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'

_fonts = {}

# Ritocchi ai testi: in PowerPoint si clicca, non si tocca la carta.
TEXT_FIX = {
    "Tocca per la risposta": "Clicca per la risposta",
    "poi tocchiamo la carta per scoprire la risposta giusta.": "poi con un clic facciamo comparire la risposta giusta.",
    "(Cliccate su × 1, × 13 e × 75 per mostrare i tre casi; da sola l'animazione li ripete in ciclo.)":
        "(L'animazione mostra da sola i tre casi, uno dopo l'altro.)",
}


def fix(text):
    for a, b in TEXT_FIX.items():
        text = text.replace(a, b)
    return text


def font(bold, size):
    key = (bold, round(size * 4))
    if key not in _fonts:
        _fonts[key] = ImageFont.truetype(CARLITO % ("Bold" if bold else "Regular"), max(1, round(size * 4)))
    return _fonts[key]


def rgb(css):
    m = re.match(r"rgba?\(([^)]+)\)", css or "")
    if not m:
        return None, 0
    v = [float(x) for x in m.group(1).replace("/", ",").split(",") if x.strip()]
    a = v[3] if len(v) > 3 else 1
    return RGBColor(*[max(0, min(255, round(c))) for c in v[:3]]), a


def is_bold(run):
    return run["weight"] >= 600 or run["fam"].startswith("Bricolage") or "Mono" in run["fam"]


def run_text(run):
    return run["t"].upper() if run["upper"] else run["t"]


def measure_lines(par, scale, width):
    """Righe necessarie a Calibri (misurato con Carlito, stesse metriche) per il paragrafo."""
    lines, cur, maxw = 1, 0.0, 0.0
    for run in par["runs"]:
        if run.get("br"):
            maxw = max(maxw, cur); lines += 1; cur = 0.0
            continue
        size = run["size"] * scale
        f = font(is_bold(run), size)
        ls = run["ls"] * scale
        for tok in re.split(r"(\s+)", run_text(run)):
            if not tok:
                continue
            w = f.getlength(tok) / 4 + ls * len(tok)
            if tok.isspace():
                cur += w
                continue
            if cur > 0 and cur + w > width + 0.5:
                maxw = max(maxw, cur); lines += 1; cur = w
            else:
                cur += w
    return lines, max(maxw, cur)


def layout(b):
    """Posizione, dimensione e scala del testo di un blocco."""
    paras = b["paras"]
    single = all(p["lines"] == 1 for p in paras)
    center = any(p["align"] == "center" for p in paras)
    if b["box"] is not None:
        x, y, w, h = b["rect"]["x"], b["rect"]["y"], b["rect"]["w"], b["rect"]["h"]
        return dict(x=x, y=y, w=w, h=h, wrap=True, scale=1.0, inset=b["pad"])
    if b["stamp"]:
        r = b["rect"]
        return dict(x=r["x"], y=r["y"], w=r["w"], h=r["h"], wrap=True, scale=1.0, inset=[v + 5 for v in b["pad"]], middle=True)
    left = min(p["text"]["x"] for p in paras)
    right = max(p["text"]["r"] for p in paras)
    if b.get("content"):
        if center:
            left, right = b["content"]["x"], b["content"]["r"]
        else:
            right = max(right, b["content"]["r"])
    elif center:
        mid = (left + right) / 2
        half = max(right - left, 120) * 0.9
        left, right = mid - half, mid + half
    width = right - left
    # scala più grande (fino a +10%) che non aggiunge righe rispetto all'HTML
    big = paras[0]["runs"][0].get("fam", "").startswith("Bricolage")
    scales = [1.0] if big else ([1.1, 1.06, 1.03, 1.0] if single else [1.06, 1.03, 1.0])
    scales += [0.97, 0.94, 0.9, 0.86]
    chosen = scales[-1]
    for s in scales:
        ok = True
        for p in paras:
            n, used = measure_lines(p, s, width)
            if n > p["lines"] or (single and used > width * (1.12 if not center else 1.0)):
                ok = False
                break
        if ok:
            chosen = s
            break
    first, last = paras[0], paras[-1]
    top = first["text"]["y"] - (first["lh"] - first["runs"][0]["size"] * 1.2) / 2
    bottom = last["text"]["b"] + (last["lh"] - last["runs"][-1].get("size", 20) * 1.2) / 2
    if single and not center:
        width = max(width, max(measure_lines(p, chosen, 1e9)[1] for p in paras) + 8)
    return dict(x=left, y=top, w=width, h=max(bottom - top, 10), wrap=not (single and not center),
                scale=chosen, inset=[0, 0, 0, 0])


def set_spacing(p_el, tag, pts100):
    el = p_el.get_or_add_pPr().find(qn(tag))
    pPr = p_el.get_or_add_pPr()
    if el is not None:
        pPr.remove(el)
    el = etree.SubElement(pPr, qn(tag))
    etree.SubElement(el, qn("a:spcPts")).set("val", str(int(pts100)))
    # ordine richiesto dallo schema: lnSpc, spcBef, spcAft
    order = [qn("a:lnSpc"), qn("a:spcBef"), qn("a:spcAft")]
    kids = sorted([k for k in pPr if k.tag in order], key=lambda k: order.index(k.tag))
    for k in kids:
        pPr.remove(k)
    for i, k in enumerate(kids):
        pPr.insert(i, k)


def add_block(slide, b):
    L = layout(b)
    sc = L["scale"]
    if b["box"] is not None:
        fill, _ = rgb(b["box"]["fill"])
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(round(L["x"] * PX)), Emu(round(L["y"] * PX)),
                                     Emu(round(L["w"] * PX)), Emu(round(L["h"] * PX)))
        shp.adjustments[0] = min(0.5, b["box"]["radius"] / min(L["w"], L["h"]))
        shp.fill.solid(); shp.fill.fore_color.rgb = fill or RGBColor(0xF2, 0xE6, 0xCC)
        shp.line.fill.background()
        shp.shadow.inherit = False
    else:
        shp = slide.shapes.add_textbox(Emu(round(L["x"] * PX)), Emu(round(L["y"] * PX)),
                                       Emu(round(L["w"] * PX)), Emu(round(L["h"] * PX)))
    if b.get("rot"):
        shp.rotation = b["rot"]
    tf = shp.text_frame
    tf.word_wrap = L["wrap"]
    tf.auto_size = MSO_AUTO_SIZE.NONE
    t, r, bt, l = L["inset"]
    tf.margin_top, tf.margin_right, tf.margin_bottom, tf.margin_left = (Emu(round(v * PX)) for v in (t, r, bt, l))
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE if L.get("middle") else MSO_ANCHOR.TOP
    prev_bottom = None
    for i, par in enumerate(b["paras"]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}.get(par["align"], PP_ALIGN.LEFT)
        set_spacing(p._p, "a:lnSpc", round(par["lh"] * sc * 0.5 * 100))
        size0 = par["runs"][0].get("size", 20)
        line_top = par["text"]["y"] - (par["lh"] - size0 * 1.2) / 2
        if prev_bottom is not None:
            set_spacing(p._p, "a:spcBef", max(0, round((line_top - prev_bottom) * 0.5 * 100)))
        prev_bottom = par["text"]["b"] + (par["lh"] - par["runs"][-1].get("size", size0) * 1.2) / 2
        for run in par["runs"]:
            if run.get("br"):
                p.add_line_break()
                continue
            rn = p.add_run()
            rn.text = fix(run["t"])
            f = rn.font
            f.name = FONT
            f.size = Pt(run["size"] * sc * 0.5)
            f.bold = is_bold(run)
            f.italic = run["italic"]
            rPr = rn._r.get_or_add_rPr()
            if run["upper"]:
                rPr.set("cap", "all")
            if run["ls"]:
                rPr.set("spc", str(round(run["ls"] * sc * 0.5 * 100)))
            if run["banded"]:
                grad = etree.SubElement(rPr, qn("a:gradFill"))
                grad.set("rotWithShape", "1")
                gs = etree.SubElement(grad, qn("a:gsLst"))
                for pos, col in BANDS:
                    g = etree.SubElement(gs, qn("a:gs")); g.set("pos", str(pos * 1000))
                    etree.SubElement(g, qn("a:srgbClr")).set("val", col)
                lin = etree.SubElement(grad, qn("a:lin")); lin.set("ang", "5400000"); lin.set("scaled", "0")
                # gradFill deve precedere latin/ea/cs nello schema
                for k in [c for c in rPr if c.tag in (qn("a:latin"), qn("a:ea"), qn("a:cs"))]:
                    rPr.remove(k); rPr.append(k)
            else:
                col, _ = rgb(run["color"])
                if col is not None:
                    f.color.rgb = col
    return shp


def fade(ids, spid, delay, node):
    a, b, c = next(ids), next(ids), next(ids)
    return (f'<p:par><p:cTn id="{a}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="{node}">'
            f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
            f'<p:set><p:cBhvr><p:cTn id="{b}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:to><p:strVal val="visible"/></p:to></p:set>'
            f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{c}" dur="500"/>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>'
            f'</p:childTnLst></p:cTn></p:par>')


def timing_xml(video, auto, clicks):
    """video: (spid, durata_ms, loop) o None; auto: [(spid, ritardo_ms)]; clicks: [spid]."""
    ids = iter(range(3, 10000))
    inner = []
    first = True
    if video:
        a, b = next(ids), next(ids)
        inner.append(f'<p:par><p:cTn id="{a}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect">'
                     f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                     f'<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="{b}" dur="{video[1]}" fill="hold"/>'
                     f'<p:tgtEl><p:spTgt spid="{video[0]}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>')
        first = False
    for spid, delay in auto:
        inner.append(fade(ids, spid, delay, "afterEffect" if first else "withEffect"))
        first = False
    groups = []
    if inner:
        g1, g2 = next(ids), next(ids)
        groups.append(f'<p:par><p:cTn id="{g1}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
                      f'<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
                      f'<p:par><p:cTn id="{g2}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                      + "".join(inner) + '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    for spid in clicks:
        g1, g2 = next(ids), next(ids)
        groups.append(f'<p:par><p:cTn id="{g1}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
                      f'<p:par><p:cTn id="{g2}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                      + fade(ids, spid, 0, "clickEffect") + '</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
    media = ""
    if video:
        rep = ' repeatCount="indefinite"' if video[2] else ""
        media = (f'<p:video><p:cMediaNode vol="80000"><p:cTn id="{next(ids)}" fill="hold" display="0"{rep}>'
                 f'<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn>'
                 f'<p:tgtEl><p:spTgt spid="{video[0]}"/></p:tgtEl></p:cMediaNode></p:video>')
    bld = "".join(f'<p:bldP spid="{s}" grpId="0" animBg="1"/>' for s, _ in auto) + \
          "".join(f'<p:bldP spid="{s}" grpId="0" animBg="1"/>' for s in clicks)
    return (f'<p:timing {NS}><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            + "".join(groups) +
            '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
            + media + '</p:childTnLst></p:cTn></p:par></p:tnLst>'
            + (f"<p:bldLst>{bld}</p:bldLst>" if bld else "") + '</p:timing>')


def finish_slide(slide, timing):
    sld = slide._element
    for old in sld.findall(qn("p:timing")) + sld.findall(qn("p:transition")):
        sld.remove(old)
    clr = sld.find(qn("p:clrMapOvr"))
    idx = list(sld).index(clr) + 1 if clr is not None else 1
    trans = etree.fromstring(f'<p:transition {NS} spd="med"><p:fade/></p:transition>')
    sld.insert(idx, trans)
    tim = etree.fromstring(timing)
    for i, ctn in enumerate(tim.iter(qn("p:cTn")), start=1):   # id in ordine di documento
        ctn.set("id", str(i))
    sld.insert(idx + 1, tim)


def poster(video, at, out):
    subprocess.run(["ffmpeg", "-hide_banner", "-loglevel", "error", "-y", "-sseof" if at < 0 else "-ss", str(at),
                    "-i", str(video), "-frames:v", "1", "-q:v", "3", str(out)], check=True)


def main():
    src, out = pathlib.Path(sys.argv[1]), pathlib.Path(sys.argv[2])
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(1920 * PX), Emu(1080 * PX)
    blank = prs.slide_layouts[6]
    for meta_path in sorted(src.glob("s[0-9][0-9].json")):
        meta = json.loads(meta_path.read_text())
        slide = prs.slides.add_slide(blank)
        bg = slide.background.fill; bg.solid(); bg.fore_color.rgb = BG
        video = None
        if meta.get("video"):
            vpath = src / meta["video"]
            post = src / (meta_path.stem + "-poster.jpg")
            poster(vpath, 0.2 if meta["loop"] else -0.1, post)
            mv = slide.shapes.add_movie(str(vpath), 0, 0, prs.slide_width, prs.slide_height,
                                        poster_frame_image=str(post), mime_type="video/mp4")
            mv.name = "Animazione"
            video = (mv.shape_id, meta["durMs"], meta["loop"])
        else:
            pic = slide.shapes.add_picture(str(src / meta["image"]), 0, 0, prs.slide_width, prs.slide_height)
            pic.name = "Sfondo"
        auto, clicks = [], []
        for b in meta["blocks"]:
            shp = add_block(slide, b)
            if b["click"]:
                clicks.append(shp.shape_id)
            else:
                auto.append((shp.shape_id, round(b["delay"] * 1000)))
        auto.sort(key=lambda a: a[1])
        finish_slide(slide, timing_xml(video, auto, clicks))
        slide.notes_slide.notes_text_frame.text = fix(meta.get("notes", ""))
    prs.save(out)
    print("salvato", out, round(out.stat().st_size / 1e6, 1), "MB,", len(prs.slides), "slide")


if __name__ == "__main__":
    main()
