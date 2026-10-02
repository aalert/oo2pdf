"""SmartArt diagrams, drawn from the cached drawing Word stores with them.

A SmartArt graphic (a:graphicData uri=".../diagram") references its data,
layout, style and colour parts. Laying those out would mean re-implementing
SmartArt's layout engine; Word however also saves the rendered result as plain
DrawingML shapes (word/diagrams/drawingN.xml, linked from the data part via
dsp:dataModelExt), and that is what Word itself draws. We read those shapes.
"""
from __future__ import annotations

import math

from ..opc import A, R, local
from ..theme import drawingml_color
from .model import Paragraph, RunStyle, Shape, Text
from .props import num

EMU = 12700.0
DGM = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}"
DSP = "{http://schemas.microsoft.com/office/drawing/2008/diagram}"

_ALIGN = {"l": "left", "ctr": "center", "r": "right", "just": "both", "dist": "distribute"}


def diagram_shape(reader, part, graphic_data, width, height) -> Shape | None:
    """The cached drawing of a SmartArt graphic as a group Shape, or None."""
    ids = graphic_data.find(DGM + "relIds")
    if ids is None:
        return None
    rels = reader.pkg.rels(part)
    data_rel = rels.get(ids.get(R + "dm"))
    if not data_rel:
        return None
    data_part = data_rel[1]
    drawing_part = None
    data_root = reader.pkg.xml(data_part)
    if data_root is not None:
        ext = next(data_root.iter(DSP + "dataModelExt"), None)
        if ext is not None and ext.get("relId") in rels:
            drawing_part = rels[ext.get("relId")][1]
    if drawing_part is None:
        return None
    root = reader.pkg.xml(drawing_part) if drawing_part else None
    if root is None:
        return None
    tree = root.find(DSP + "spTree")
    if tree is None:
        return None
    group = Shape(width, height, geom="group")
    for sp in tree.iter(DSP + "sp"):
        child = _shape(reader, sp)
        if child is None:
            continue
        for dx, dy, s in child:
            group.children.append((dx, dy, s))
    return group


def _shape(reader, sp):
    """[(x, y, Shape)] for one cached diagram shape (its outline and its text)."""
    theme = reader.theme
    sppr = sp.find(DSP + "spPr")
    if sppr is None:
        return None
    xfrm = sppr.find(A + "xfrm")
    if xfrm is None:
        return None
    off, ext = xfrm.find(A + "off"), xfrm.find(A + "ext")
    x = num(off.get("x"), 0) / EMU if off is not None else 0.0
    y = num(off.get("y"), 0) / EMU if off is not None else 0.0
    w = num(ext.get("cx"), 0) / EMU if ext is not None else 0.0
    h = num(ext.get("cy"), 0) / EMU if ext is not None else 0.0
    rot = num(xfrm.get("rot"), 0) / 60000.0
    out = []
    s = Shape(w, h)
    geom = sppr.find(A + "prstGeom")
    s.geom = geom.get("prst", "rect") if geom is not None else "rect"
    style = sp.find(DSP + "style")
    # fill / outline: explicit first, then the shape style's references
    if sppr.find(A + "noFill") is not None:
        s.fill = None
    elif sppr.find(A + "solidFill") is not None:
        s.fill = drawingml_color(sppr.find(A + "solidFill"), theme)
    elif style is not None and style.find(A + "fillRef") is not None and style.find(A + "fillRef").get("idx") != "0":
        s.fill = drawingml_color(style.find(A + "fillRef"), theme)
    else:
        s.fill = None
    ln = sppr.find(A + "ln")
    if ln is not None and ln.find(A + "noFill") is not None:
        s.line = None
    elif ln is not None and ln.find(A + "solidFill") is not None:
        s.line = drawingml_color(ln.find(A + "solidFill"), theme)
    elif style is not None and style.find(A + "lnRef") is not None and style.find(A + "lnRef").get("idx") != "0":
        s.line = drawingml_color(style.find(A + "lnRef"), theme)
    else:
        s.line = None
    if ln is not None and ln.get("w"):
        s.line_width = num(ln.get("w"), 9525) / EMU
    cust = sppr.find(A + "custGeom")
    if cust is not None:
        pts, closed = _cust_points(cust, w, h)
        if pts:
            s.geom = "path"
            s.points = pts
            s.closed = closed
    if rot and s.geom in ("path", "line", "straightConnector1"):
        if s.geom != "path":
            s.points, s.geom = [(0.0, 0.0), (w, h)], "path"
        cx, cy = w / 2, h / 2
        a = math.radians(rot)
        ca, sa = math.cos(a), math.sin(a)
        s.points = [(cx + (px - cx) * ca - (py - cy) * sa, cy + (px - cx) * sa + (py - cy) * ca)
                    for px, py in s.points]
    if s.fill is not None or s.line is not None:
        out.append((x, y, s))
    # text, laid out in the text rectangle Word cached (dsp:txXfrm)
    body = sp.find(DSP + "txBody")
    if body is not None:
        blocks = _paragraphs(reader, body, style)
        if any(p.items for p in blocks):
            tx = sp.find(DSP + "txXfrm")
            tx_x, tx_y, tx_w, tx_h = x, y, w, h
            if tx is not None and tx.find(A + "off") is not None and tx.find(A + "ext") is not None:
                o, e = tx.find(A + "off"), tx.find(A + "ext")
                tx_x, tx_y = num(o.get("x"), 0) / EMU, num(o.get("y"), 0) / EMU
                tx_w, tx_h = num(e.get("cx"), 0) / EMU, num(e.get("cy"), 0) / EMU
            t = Shape(tx_w, tx_h, fill=None, line=None)
            t.blocks = blocks
            bp = body.find(A + "bodyPr")
            if bp is not None:
                t.insets = tuple(num(bp.get(k), d) / EMU for k, d in
                                 (("lIns", 91440), ("tIns", 45720), ("rIns", 91440), ("bIns", 45720)))
                t.v_anchor = bp.get("anchor", "t")
            out.append((tx_x, tx_y, t))
    return out


def _cust_points(cust, w, h):
    pts = []
    closed = False
    for path in cust.iter(A + "path"):
        pw = num(path.get("w"), 0) / EMU or w
        ph = num(path.get("h"), 0) / EMU or h
        sx = w / pw if pw else 1.0
        sy = h / ph if ph else 1.0
        for cmd in path:
            name = local(cmd)
            coords = [(num(p.get("x"), 0) / EMU * sx, num(p.get("y"), 0) / EMU * sy) for p in cmd.iter(A + "pt")]
            if name in ("moveTo", "lnTo") and coords:
                pts.append(coords[0])
            elif name in ("cubicBezTo", "quadBezTo") and coords:
                pts.extend(coords)  # control points approximate the curve well enough here
            elif name == "close":
                closed = True
    return pts, closed


def _paragraphs(reader, body, style) -> list:
    theme = reader.theme
    font_ref = style.find(A + "fontRef") if style is not None else None
    ref_font = None
    ref_color = None
    if font_ref is not None:
        ref_font = (theme.major if font_ref.get("idx") == "major" else theme.minor).get("latin")
        ref_color = drawingml_color(font_ref, theme)
    out = []
    for p in body.findall(A + "p"):
        ppr = p.find(A + "pPr")
        jc = _ALIGN.get(ppr.get("algn"), "left") if ppr is not None and ppr.get("algn") else "left"
        pp = {"jc": jc, "sp_before": 0.0, "sp_after": 0.0, "ind_left": 0.0, "ind_first": 0.0}
        if ppr is not None:
            pct = ppr.find(f"{A}lnSpc/{A}spcPct")
            if pct is not None:
                pp["line"] = num(pct.get("val"), 100000) / 100000 * 240
                pp["lineRule"] = "auto"
        items = []
        mark = None
        for r in p:
            if local(r) not in ("r", "endParaRPr", "fld"):
                continue
            rpr = r.find(A + "rPr") if local(r) != "endParaRPr" else r
            rs = _run_style(rpr, ref_font, ref_color, theme)
            if local(r) == "endParaRPr":
                mark = rs
                continue
            t = r.find(A + "t")
            if t is not None and t.text:
                items.append(Text(t.text, rs))
                mark = mark or rs
        out.append(Paragraph(ppr=pp, items=items, mark=mark or _run_style(None, ref_font, ref_color, theme)))
    return out


def _run_style(rpr, ref_font, ref_color, theme) -> RunStyle:
    size = num(rpr.get("sz"), 1800) / 100.0 if rpr is not None and rpr.get("sz") else 18.0
    font = ref_font or theme.minor.get("latin") or "Calibri"
    color = ref_color
    bold = italic = False
    if rpr is not None:
        latin = rpr.find(A + "latin")
        if latin is not None and latin.get("typeface") and not latin.get("typeface").startswith("+"):
            font = latin.get("typeface")
        if rpr.find(A + "solidFill") is not None:
            color = drawingml_color(rpr.find(A + "solidFill"), theme)
        bold = rpr.get("b") in ("1", "true")
        italic = rpr.get("i") in ("1", "true")
    return RunStyle(font_ascii=font, font_hansi=font, font_ea=font, font_cs=font, size=size, size_cs=size,
                    bold=bold, italic=italic, color=color)
