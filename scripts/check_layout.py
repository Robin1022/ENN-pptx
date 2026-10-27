#!/usr/bin/env python3
"""Geometry checks for slide layout defects that a rendered preview makes easy to miss.

Reads shape coordinates from the slide XML and reports:

  确定(按坐标算出的几何事实):
    - 形状超出画布
    - 图片/图标跨过所在框的边框
    - 箭头过短,只剩箭头头部(带箭头的连接线、块箭头)
  估算(按字宽估算文字占位):
    - 文字可能溢出有边框或底色的框
    - 文字可能跨过所在框的边框
    - 连接线可能穿过文字

几何事实不等于缺陷:有意压在边框上的徽标、刻意画成三角的小箭头也会被列出。
每一条都要对照渲染图判断是不是用户能看到的问题。

Usage:
    python3 scripts/check_layout.py deck.pptx              # 检查全部页面
    python3 scripts/check_layout.py deck.pptx --slides 5,6 # 只查改过的页
    python3 scripts/check_layout.py out.pptx --original backup.pptx  # 原文件已有的不再报

Exit code 1 only when a shape lies outside the slide.
"""

from __future__ import annotations

import argparse
import math
import re
import sys
import zipfile
from dataclasses import dataclass, field

from defusedxml import ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
A = "{%s}" % NS["a"]
P = "{%s}" % NS["p"]

EMU_PER_PT = 12700
EMU_PER_IN = 914400
TOL = 2 * EMU_PER_PT  # 2pt overhang is not visible

BLOCK_ARROWS = {
    "rightArrow": "x", "leftArrow": "x", "leftRightArrow": "x",
    "upArrow": "y", "downArrow": "y", "upDownArrow": "y",
}
ARROW_LEN = {"sm": 2, "med": 3, "lg": 5}  # arrowhead length as a multiple of line width
# Only these outlines match their bounding box; skewed slabs, cans and diamonds do not.
RECT_LIKE = {
    "rect", "roundRect", "snip1Rect", "snip2SameRect", "round1Rect", "round2SameRect",
    "round2DiagRect", "flowChartProcess", "flowChartAlternateProcess",
}
TINY_ARROW = 0.15 * 914400  # block arrows this narrow are drawn as arrowheads on purpose
OVERFLOW_HINT = 10 * 12700


@dataclass
class Box:
    x: float
    y: float
    w: float
    h: float

    @property
    def r(self):
        return self.x + self.w

    @property
    def b(self):
        return self.y + self.h

    @property
    def area(self):
        return max(self.w, 0) * max(self.h, 0)

    def inter(self, o: "Box") -> float:
        w = min(self.r, o.r) - max(self.x, o.x)
        h = min(self.b, o.b) - max(self.y, o.y)
        return w * h if w > 0 and h > 0 else 0.0

    def contains(self, o: "Box", tol: float = TOL) -> bool:
        return (o.x >= self.x - tol and o.y >= self.y - tol
                and o.r <= self.r + tol and o.b <= self.b + tol)


@dataclass
class Shape:
    sid: str
    name: str
    kind: str  # sp | cxn | pic
    box: Box
    rotated: bool = False
    placeholder: bool = False  # position comes from the layout
    flip_h: bool = False
    flip_v: bool = False
    geom: str = ""
    adj: dict = field(default_factory=dict)
    outlined: bool = False
    filled: bool = False
    arrow_ends: list = field(default_factory=list)  # [(end, len, width_emu)]
    line_w: float = 9525
    text: str = ""
    text_box: Box | None = None  # estimated extent of the rendered text
    overflow: float = 0.0  # estimated text height beyond the box, in EMU

    def label(self):
        t = self.text.replace("\n", " ").strip()
        t = f"「{t[:18]}{'…' if len(t) > 18 else ''}」" if t else ""
        return f"{self.name}(id={self.sid}){t}"


def char_w(ch: str, size: float) -> float:
    if ch == " ":
        return 0.3 * size
    if ord(ch) > 0x2E7F:  # CJK, fullwidth punctuation
        return 1.0 * size
    if ch.isupper() or ch in "mwMW@%":
        return 0.68 * size
    return 0.52 * size


def run_size(el, default):
    for tag in ("rPr", "endParaRPr"):
        rpr = el.find(f"{A}{tag}")
        if rpr is not None and rpr.get("sz"):
            return int(rpr.get("sz")) / 100 * EMU_PER_PT
    return default


def estimate_text(sp, box: Box):
    """Return (text, text_box, overflow) from a rough line-wrap estimate."""
    tx = sp.find(f"{P}txBody")
    if tx is None:
        return "", None, 0.0
    body = tx.find(f"{A}bodyPr")
    body = body if body is not None else ET.fromstring(f'<a:bodyPr xmlns:a="{NS["a"]}"/>')
    lI = int(body.get("lIns", 91440)); rI = int(body.get("rIns", 91440))
    tI = int(body.get("tIns", 45720)); bI = int(body.get("bIns", 45720))
    wrap = body.get("wrap", "square") != "none"
    autofit = body.find(f"{A}normAutofit") is not None or body.find(f"{A}spAutoFit") is not None
    scale = 1.0
    nf = body.find(f"{A}normAutofit")
    if nf is not None and nf.get("fontScale"):
        scale = int(nf.get("fontScale")) / 100000
    avail = max(box.w - lI - rI, 1)

    lines_h = 0.0
    widest = 0.0
    texts = []
    algns = []
    for para in tx.findall(f"{A}p"):
        ppr = para.find(f"{A}pPr")
        algns.append(ppr.get("algn", "l") if ppr is not None else "l")
        default = 18 * EMU_PER_PT
        runs = para.findall(f"{A}r") + para.findall(f"{A}fld")
        sizes, width, ptext = [], 0.0, ""
        for r in runs:
            size = run_size(r, default) * scale
            t = "".join(x.text or "" for x in r.findall(f"{A}t"))
            ptext += t
            sizes.append(size)
            width += sum(char_w(c, size) for c in t)
        size = max(sizes) if sizes else run_size(para, default) * scale
        texts.append(ptext)
        n_lines = max(1, math.ceil(width / avail)) if (wrap and width) else 1
        widest = max(widest, min(width, avail) if wrap else width)
        spacing = 1.2
        if ppr is not None:
            pct = ppr.find(f"{A}lnSpc/{A}spcPct")
            if pct is not None:
                spacing = 1.2 * int(pct.get("val")) / 100000
            for tag in ("spcBef", "spcAft"):
                pts = ppr.find(f"{A}{tag}/{A}spcPts")
                if pts is not None:
                    lines_h += int(pts.get("val")) / 100 * EMU_PER_PT
        lines_h += n_lines * size * spacing

    text = "\n".join(texts).strip()
    if not text:
        return "", None, 0.0
    text_h = lines_h + tI + bI
    text_w = widest + lI + rI
    anchor = body.get("anchor", "t")
    if anchor == "ctr":
        ty = box.y + (box.h - text_h) / 2
    elif anchor == "b":
        ty = box.b - text_h
    else:
        ty = box.y
    algn = max(set(algns), key=algns.count)
    if algn == "ctr":
        tx_ = box.x + (box.w - text_w) / 2
    elif algn == "r":
        tx_ = box.r - text_w
    else:
        tx_ = box.x
    overflow = 0.0 if autofit else max(0.0, text_h - box.h)
    if not wrap and not autofit:
        overflow = max(overflow, text_w - box.w)
    return text, Box(tx_, ty, text_w, text_h), overflow


def has_visible_line(sppr, style) -> bool:
    ln = sppr.find(f"{A}ln") if sppr is not None else None
    if ln is not None:
        if ln.find(f"{A}noFill") is not None:
            return False
        if ln.find(f"{A}solidFill") is not None or ln.find(f"{A}gradFill") is not None:
            return True
    if style is not None:
        ref = style.find(f"{A}lnRef")
        return ref is not None and ref.get("idx", "0") != "0"
    return False


def has_fill(sppr, style) -> bool:
    if sppr is not None:
        if sppr.find(f"{A}noFill") is not None:
            return False
        if any(sppr.find(f"{A}{t}") is not None for t in ("solidFill", "gradFill", "blipFill", "pattFill")):
            return True
    if style is not None:
        ref = style.find(f"{A}fillRef")
        return ref is not None and ref.get("idx", "0") != "0"
    return False


def walk(tree, xf, out):
    for el in tree:
        tag = el.tag.split("}")[1]
        if tag == "grpSp":
            g = el.find(f"{P}grpSpPr/{A}xfrm")
            if g is None:
                walk(el, xf, out)
                continue
            off, ext = g.find(f"{A}off"), g.find(f"{A}ext")
            choff, chext = g.find(f"{A}chOff"), g.find(f"{A}chExt")
            ox, oy = xf(int(off.get("x")), int(off.get("y")))
            sx = int(ext.get("cx")) / max(int(chext.get("cx")), 1)
            sy = int(ext.get("cy")) / max(int(chext.get("cy")), 1)
            px0, py0 = xf(0, 0)
            px1, py1 = xf(1, 1)
            psx, psy = px1 - px0, py1 - py0
            cx0, cy0 = int(choff.get("x")), int(choff.get("y"))

            def child(x, y, ox=ox, oy=oy, sx=sx, sy=sy, psx=psx, psy=psy, cx0=cx0, cy0=cy0):
                return ox + (x - cx0) * sx * psx, oy + (y - cy0) * sy * psy

            walk(el, child, out)
            continue
        if tag not in ("sp", "cxnSp", "pic"):
            continue
        sppr = el.find(f"{P}spPr")
        x_ = sppr.find(f"{A}xfrm") if sppr is not None else None
        if x_ is None or x_.find(f"{A}off") is None:
            continue  # inherits placement from the layout
        off, ext = x_.find(f"{A}off"), x_.find(f"{A}ext")
        x0, y0 = xf(int(off.get("x")), int(off.get("y")))
        x1, y1 = xf(int(off.get("x")) + int(ext.get("cx")), int(off.get("y")) + int(ext.get("cy")))
        nv = el.find(f".//{P}cNvPr")
        geom = sppr.find(f"{A}prstGeom")
        style = el.find(f"{P}style")
        s = Shape(
            sid=nv.get("id", "?") if nv is not None else "?",
            name=nv.get("name", "") if nv is not None else "",
            kind={"sp": "sp", "cxnSp": "cxn", "pic": "pic"}[tag],
            box=Box(x0, y0, x1 - x0, y1 - y0),
            rotated=int(x_.get("rot", "0")) % 21600000 != 0,
            placeholder=el.find(f".//{P}nvPr/{P}ph") is not None,
            flip_h=x_.get("flipH") == "1",
            flip_v=x_.get("flipV") == "1",
            geom=geom.get("prst", "") if geom is not None else "custom",
            outlined=has_visible_line(sppr, style),
            filled=has_fill(sppr, style),
        )
        if geom is not None:
            for gd in geom.findall(f"{A}avLst/{A}gd"):
                m = re.match(r"val (-?\d+)", gd.get("fmla", ""))
                if m:
                    s.adj[gd.get("name")] = int(m.group(1))
        ln = sppr.find(f"{A}ln")
        if ln is not None:
            s.line_w = int(ln.get("w", 9525))
            for end in ("headEnd", "tailEnd"):
                e = ln.find(f"{A}{end}")
                if e is not None and e.get("type", "none") != "none":
                    s.arrow_ends.append((end, e.get("len", "med"), s.line_w))
        if tag == "sp":
            s.text, s.text_box, s.overflow = estimate_text(el, s.box)
        out.append(s)


def slide_size(z):
    root = ET.fromstring(z.read("ppt/presentation.xml"))
    sz = root.find(f"{P}sldSz")
    return int(sz.get("cx")), int(sz.get("cy"))


def is_line(s: Shape):
    return s.kind == "cxn" or s.geom in ("line", "straightConnector1")


def is_container(s: Shape):
    return s.kind == "sp" and s.geom in RECT_LIKE and s.outlined and not s.rotated


def segment_hits_box(x0, y0, x1, y1, b: Box) -> bool:
    """Liang-Barsky clip: does the segment pass through the interior of b?"""
    dx, dy = x1 - x0, y1 - y0
    t0, t1 = 0.0, 1.0
    for p, q in ((-dx, x0 - b.x), (dx, b.r - x0), (-dy, y0 - b.y), (dy, b.b - y0)):
        if p == 0:
            if q < 0:
                return False
            continue
        t = q / p
        if p < 0:
            t0 = max(t0, t)
        else:
            t1 = min(t1, t)
        if t0 > t1:
            return False
    return t1 - t0 > 0.05


def line_path(s: Shape):
    """Points along a connector, following elbows instead of the bounding-box diagonal."""
    x0, x1 = (s.box.r, s.box.x) if s.flip_h else (s.box.x, s.box.r)
    y0, y1 = (s.box.b, s.box.y) if s.flip_v else (s.box.y, s.box.b)
    if s.geom == "bentConnector2":
        return [(x0, y0), (x1, y0), (x1, y1)]
    if s.geom == "bentConnector3":
        xm = x0 + (x1 - x0) * s.adj.get("adj1", 50000) / 100000
        return [(x0, y0), (xm, y0), (xm, y1), (x1, y1)]
    if s.geom.startswith(("bentConnector", "curvedConnector")):
        return None  # path not modelled; skip rather than guess
    return [(x0, y0), (x1, y1)]


def check_slide(shapes, size):
    outside, errors, hints = [], [], []
    W, H = size
    canvas = Box(0, 0, W, H)
    for s in shapes:
        if not s.rotated and not s.placeholder and not canvas.contains(s.box):
            outside.append(f"{s.label()} 超出画布")

    # 箭头过短
    for s in shapes:
        if is_line(s) and s.arrow_ends:
            length = math.hypot(s.box.w, s.box.h)
            head = max(ARROW_LEN.get(l, 3) * max(w, 12700) for _, l, w in s.arrow_ends)
            need = head * len(s.arrow_ends) + TOL
            if length < need:
                errors.append(f"{s.label()} 箭头长 {length / EMU_PER_PT:.0f}pt,只够画出箭头头部(至少 {need / EMU_PER_PT:.0f}pt)")
        elif s.geom in BLOCK_ARROWS and not s.rotated:
            axis = BLOCK_ARROWS[s.geom]
            long_side = s.box.w if axis == "x" else s.box.h
            short_side = s.box.h if axis == "x" else s.box.w
            if short_side < TINY_ARROW:
                continue
            heads = 2 if s.geom.endswith(("RightArrow", "DownArrow")) else 1
            head = s.adj.get("adj2", 50000) / 100000 * short_side * heads
            shaft = long_side - head
            if shaft < 0.25 * long_side or shaft < 0.08 * EMU_PER_IN:
                errors.append(f"{s.label()} 块箭头的箭身只剩 {max(shaft, 0) / EMU_PER_PT:.0f}pt,看起来只有一个三角")

    containers = [s for s in shapes if is_container(s)]
    for s in shapes:
        if is_line(s) or s.rotated:
            continue
        # 文字溢出
        if s.overflow > OVERFLOW_HINT and (s.outlined or s.filled):
            hints.append(f"{s.label()} 文字估计超出文本框约 {s.overflow / EMU_PER_PT:.0f}pt")
        # 跨边框
        # (what, extent, exact, lowest inside share that counts)
        probes = []
        if s.kind == "pic":
            probes.append(("图片", s.box, True, 0.15))
        elif s.filled and s.kind == "sp":
            # Tabs and badges straddle a border by design, about half in; a box that is
            # mostly inside and pokes a little past the border is the accident.
            probes.append(("方块", s.box, True, 0.7))
        elif s.text_box is not None:
            probes.append(("文字", s.text_box, False, 0.15))
        for what, b, exact, low in probes:
            if b.area <= 0:
                continue
            for c in containers:
                if c is s or c.box.area < 1.5 * b.area:
                    continue
                share = c.box.inter(b) / b.area
                if low < share and not c.box.contains(b):
                    msg = f"{s.label()} 的{what}跨过了 {c.label()} 的边框"
                    (errors if exact else hints).append(msg)
                    break

    # 连接线穿过文字
    z = {id(s): i for i, s in enumerate(shapes)}
    texts = [s for s in shapes if s.text_box is not None and not s.rotated]
    for s in shapes:
        if not is_line(s) or s.rotated:
            continue
        path = line_path(s)
        if path is None:
            continue
        (x0, y0), (x1, y1) = path[0], path[-1]
        for t in texts:
            if t.filled and z[id(t)] > z[id(s)]:
                continue  # an opaque label drawn above the line masks it
            tb = t.text_box
            mx, my = max(TOL, 0.15 * tb.w), max(TOL, 0.25 * tb.h)
            inner = Box(tb.x + mx, tb.y + my, tb.w - 2 * mx, tb.h - 2 * my)  # strike through the glyphs
            if inner.w <= 0 or inner.h <= 0:
                continue
            if Box(x0 - 1, y0 - 1, 2, 2).inter(t.box) or Box(x1 - 1, y1 - 1, 2, 2).inter(t.box):
                continue  # the line starts or ends on this shape
            if any(segment_hits_box(*a, *b, inner) for a, b in zip(path, path[1:])):
                hints.append(f"{s.label()} 可能穿过文字 {t.label()}")
    return outside, errors, hints


def deck_findings(path, wanted=None):
    """Yield (index, part, outside, errors, hints) for each slide in show order."""
    with zipfile.ZipFile(path) as z:
        size = slide_size(z)
        pres = ET.fromstring(z.read("ppt/presentation.xml"))
        rels = ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
        target = {r.get("Id"): r.get("Target") for r in rels}
        rid = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        order = [target[s.get(rid)] for s in pres.find(f"{P}sldIdLst")]
        for idx, part in enumerate(order, 1):
            if wanted and idx not in wanted:
                continue
            path_ = "ppt/" + part.lstrip("/").removeprefix("ppt/")
            root = ET.fromstring(z.read(path_))
            shapes = []
            walk(root.find(f"{P}cSld/{P}spTree"), lambda x, y: (x, y), shapes)
            yield (idx, path_, *check_slide(shapes, size))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--slides", help="只检查这些页,按放映顺序编号,如 5,6")
    ap.add_argument("--original", help="改动前的文件或模板;它里面已有的同一条结果不再报出")
    args = ap.parse_args()
    wanted = {int(n) for n in args.slides.split(",")} if args.slides else None

    # Shapes keep their name and id when a page is edited or copied with add_slide.py,
    # so an identical message means the defect came with the source, not this edit.
    known = set()
    if args.original:
        for _, _, o, e, h in deck_findings(args.original):
            known.update(o + e + h)

    n_out = n_err = n_hint = n_known = 0
    for idx, path, outside, errors, hints in deck_findings(args.deck, wanted):
        before = len(outside) + len(errors) + len(hints)
        outside, errors, hints = ([m for m in ms if m not in known] for ms in (outside, errors, hints))
        n_known += before - len(outside) - len(errors) - len(hints)
        if not (outside or errors or hints):
            continue
        print(f"第 {idx} 页 ({path.split('/')[-1]})")
        for m in outside + errors:
            print(f"  确定  {m}")
        for m in hints:
            print(f"  估算  {m}")
        n_out += len(outside)
        n_err += len(errors)
        n_hint += len(hints)

    if n_known:
        print(f"(原文件里已有的 {n_known} 条没有列出)")

    if n_out or n_err or n_hint:
        print(f"\n共 {n_out + n_err} 条确定、{n_hint} 条估算。逐条对照渲染图判断是否为可见问题,"
              "有意的设计(压边框的徽标、三角形小箭头)不用改。")
    else:
        print("PASSED - 没有发现越界、跨边框或过短的箭头")
    sys.exit(1 if n_out else 0)


if __name__ == "__main__":
    main()
