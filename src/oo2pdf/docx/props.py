"""Parsing of WordprocessingML property elements into plain dicts.

Every *Pr element is turned into a flat dict so style inheritance becomes a
simple ordered merge (see :func:`merge`).
"""
from __future__ import annotations

from dataclasses import dataclass

from ..opc import W, local


def attr(el, name, default=None):
    if el is None:
        return default
    v = el.get(W + name)
    return default if v is None else v


def onoff(el) -> bool:
    v = el.get(W + "val")
    return v is None or v.lower() not in ("0", "false", "off", "none")


def num(v, default=None):
    if v is None or v == "":
        return default
    try:
        return float(v)
    except ValueError:
        # universal measure, e.g. "1.5in", "12pt"
        units = {"pt": 20, "in": 1440, "cm": 566.93, "mm": 56.693, "pc": 240, "pi": 240}
        for u, f in units.items():
            if v.endswith(u):
                try:
                    return float(v[:-len(u)]) * f
                except ValueError:
                    return default
        if v.endswith("%"):
            try:
                return float(v[:-1])
            except ValueError:
                return default
        return default


def twips(v, default=None):
    n = num(v)
    return default if n is None else n / 20.0


@dataclass(frozen=True)
class Border:
    style: str
    width: float  # points
    color: tuple | None  # None = auto
    space: float  # points

    @property
    def visible(self):
        return self.style not in ("nil", "none") and self.width > 0


def parse_border(el) -> Border | None:
    if el is None:
        return None
    style = attr(el, "val", "single")
    if style in ("nil", "none"):
        return Border("nil", 0, None, 0)
    sz = num(attr(el, "sz"), 4) / 8.0
    if style in ("thick",):
        sz = max(sz, 1.5)
    sz = max(sz, 0.25)
    color = attr(el, "color")
    from ..theme import hex_color
    col = None if color in (None, "auto") else hex_color(color)
    space = num(attr(el, "space"), 0)
    return Border(style, sz, col, space)


def parse_borders(el) -> dict:
    out = {}
    if el is None:
        return out
    for child in el:
        name = local(child)
        name = {"start": "left", "end": "right"}.get(name, name)
        b = parse_border(child)
        if b is not None:
            out[name] = b
    return out


def parse_shd(el):
    """Return a fill colour tuple, 'none' for explicit clear, or None."""
    from ..theme import hex_color
    if el is None:
        return None
    fill = attr(el, "fill")
    pat = attr(el, "val", "clear")
    color = attr(el, "color")
    base = hex_color(fill) if fill and fill != "auto" else None
    if pat not in ("clear", "nil") and pat:
        # Pattern shading: blend foreground over the fill by coverage.
        cov = None
        if pat == "solid":
            cov = 1.0
        elif pat.startswith("pct"):
            try:
                cov = int(pat[3:]) / 100.0
            except ValueError:
                cov = None
        if cov is not None:
            fg = hex_color(color) if color and color != "auto" else (0.0, 0.0, 0.0)
            bg = base or (1.0, 1.0, 1.0)
            return tuple(bg[i] * (1 - cov) + fg[i] * cov for i in range(3))
    return base if base is not None else "none"


def parse_rpr(el) -> dict:
    d: dict = {}
    if el is None:
        return d
    from ..theme import hex_color
    for c in el:
        t = local(c)
        if t == "rStyle":
            d["rStyle"] = attr(c, "val")
        elif t == "rFonts":
            for a in ("ascii", "hAnsi", "eastAsia", "cs"):
                v = attr(c, a)
                if v is not None:
                    d["font_" + a] = v
                    d.pop("theme_" + a, None)
            for a, key in (("asciiTheme", "ascii"), ("hAnsiTheme", "hAnsi"),
                           ("eastAsiaTheme", "eastAsia"), ("cstheme", "cs")):
                v = attr(c, a)
                if v is not None:
                    d["theme_" + key] = v
            if attr(c, "hint") is not None:
                d["hint"] = attr(c, "hint")
        elif t in ("b", "i", "caps", "smallCaps", "strike", "dstrike", "vanish", "bCs", "iCs",
                   "outline", "shadow", "emboss", "imprint", "specVanish", "rtl", "cs"):
            d[t] = onoff(c)
        elif t == "sz":
            d["sz"] = num(attr(c, "val"), 22) / 2.0
        elif t == "szCs":
            d["szCs"] = num(attr(c, "val"), 22) / 2.0
        elif t == "color":
            v = attr(c, "val")
            d["color"] = None if v in (None, "auto") else hex_color(v)
        elif t == "u":
            d["u"] = attr(c, "val", "single")
            uc = attr(c, "color")
            d["u_color"] = None if uc in (None, "auto") else hex_color(uc)
        elif t == "highlight":
            d["highlight"] = attr(c, "val")
        elif t == "shd":
            d["shd"] = parse_shd(c)
        elif t == "vertAlign":
            d["vertAlign"] = attr(c, "val")
        elif t == "spacing":
            d["spacing"] = twips(attr(c, "val"), 0)
        elif t == "position":
            d["position"] = num(attr(c, "val"), 0) / 2.0
        elif t == "w":
            d["scale"] = num(attr(c, "val"), 100)
        elif t == "kern":
            d["kern"] = num(attr(c, "val"), 0) / 2.0
        elif t == "bdr":
            d["bdr"] = parse_border(c)
        elif t == "lang":
            d["lang"] = attr(c, "val")
    return d


def parse_tabs(el) -> list:
    tabs = []
    for t in el:
        if local(t) != "tab":
            continue
        pos = twips(attr(t, "pos"), 0)
        tabs.append((attr(t, "val", "left"), pos, attr(t, "leader", "none")))
    return tabs


def parse_ppr(el) -> dict:
    d: dict = {}
    if el is None:
        return d
    for c in el:
        t = local(c)
        if t == "pStyle":
            d["pStyle"] = attr(c, "val")
        elif t in ("keepNext", "keepLines", "pageBreakBefore", "widowControl",
                   "contextualSpacing", "suppressAutoHyphens", "bidi", "snapToGrid",
                   "mirrorIndents", "suppressLineNumbers", "wordWrap"):
            d[t] = onoff(c)
        elif t == "numPr":
            ilvl = c.find(W + "ilvl")
            nid = c.find(W + "numId")
            if nid is not None:
                d["numId"] = attr(nid, "val")
            if ilvl is not None:
                d["ilvl"] = int(num(attr(ilvl, "val"), 0))
        elif t == "pBdr":
            d["pBdr"] = parse_borders(c)
        elif t == "shd":
            d["shd"] = parse_shd(c)
        elif t == "tabs":
            d["tabs"] = parse_tabs(c)
        elif t == "spacing":
            for a, key in (("before", "sp_before"), ("after", "sp_after")):
                v = attr(c, a)
                if v is not None:
                    d[key] = twips(v, 0)
            for a, key in (("beforeLines", "sp_before_lines"), ("afterLines", "sp_after_lines")):
                v = attr(c, a)
                if v is not None:
                    d[key] = num(v, 0) / 100.0
            for a, key in (("beforeAutospacing", "sp_before_auto"), ("afterAutospacing", "sp_after_auto")):
                v = attr(c, a)
                if v is not None:
                    d[key] = v.lower() not in ("0", "false", "off")
            if attr(c, "line") is not None:
                d["line"] = num(attr(c, "line"), 240)
                d["lineRule"] = attr(c, "lineRule", "auto")
        elif t == "ind":
            left = attr(c, "left") if attr(c, "left") is not None else attr(c, "start")
            right = attr(c, "right") if attr(c, "right") is not None else attr(c, "end")
            if left is not None:
                d["ind_left"] = twips(left, 0)
            if right is not None:
                d["ind_right"] = twips(right, 0)
            if attr(c, "hanging") is not None:
                d["ind_first"] = -twips(attr(c, "hanging"), 0)
            elif attr(c, "firstLine") is not None:
                d["ind_first"] = twips(attr(c, "firstLine"), 0)
        elif t == "jc":
            d["jc"] = attr(c, "val", "left")
        elif t == "outlineLvl":
            d["outlineLvl"] = int(num(attr(c, "val"), 9))
        elif t == "rPr":
            d["mark_rpr"] = parse_rpr(c)
        elif t == "framePr":
            d["framePr"] = dict((local_attr(k), v) for k, v in c.attrib.items())
        elif t == "textAlignment":
            d["textAlignment"] = attr(c, "val")
    return d


def local_attr(name: str) -> str:
    return name.rsplit("}", 1)[-1]


def parse_cell_margins(el) -> dict:
    out = {}
    if el is None:
        return out
    for c in el:
        name = {"start": "left", "end": "right"}.get(local(c), local(c))
        typ = attr(c, "type", "dxa")
        w = num(attr(c, "w"), 0)
        out[name] = w / 20.0 if typ in ("dxa", None) else 0.0
    return out


def parse_width(el):
    """Return (value, type) where type is dxa|pct|auto|nil; value in pt or percent."""
    if el is None:
        return None
    typ = attr(el, "type", "dxa")
    raw = attr(el, "w", "0")
    if typ == "pct":
        if raw.endswith("%"):
            return (num(raw[:-1], 0), "pct")
        return (num(raw, 0) / 50.0, "pct")
    if typ == "dxa":
        return (num(raw, 0) / 20.0, "dxa")
    return (0.0, typ)


def parse_tblpr(el) -> dict:
    d: dict = {}
    if el is None:
        return d
    for c in el:
        t = local(c)
        if t == "tblStyle":
            d["tblStyle"] = attr(c, "val")
        elif t == "tblW":
            d["tblW"] = parse_width(c)
        elif t == "jc":
            d["jc"] = attr(c, "val")
        elif t == "tblInd":
            d["tblInd"] = parse_width(c)[0] if attr(c, "type", "dxa") == "dxa" else 0.0
        elif t == "tblBorders":
            d["tblBorders"] = parse_borders(c)
        elif t == "shd":
            d["shd"] = parse_shd(c)
        elif t == "tblLayout":
            d["layout"] = attr(c, "type", "autofit")
        elif t == "tblCellMar":
            d["cellMar"] = parse_cell_margins(c)
        elif t == "tblCellSpacing":
            d["cellSpacing"] = twips(attr(c, "w"), 0)
        elif t == "tblLook":
            look = {}
            v = attr(c, "val")
            if v:
                try:
                    bits = int(v, 16)
                    look = {"firstRow": bool(bits & 0x20), "lastRow": bool(bits & 0x40),
                            "firstColumn": bool(bits & 0x80), "lastColumn": bool(bits & 0x100),
                            "noHBand": bool(bits & 0x200), "noVBand": bool(bits & 0x400)}
                except ValueError:
                    pass
            for k in ("firstRow", "lastRow", "firstColumn", "lastColumn", "noHBand", "noVBand"):
                a = attr(c, k)
                if a is not None:
                    look[k] = a.lower() in ("1", "true", "on")
            d["look"] = look
        elif t == "tblStyleRowBandSize":
            d["rowBand"] = int(num(attr(c, "val"), 1))
        elif t == "tblStyleColBandSize":
            d["colBand"] = int(num(attr(c, "val"), 1))
        elif t == "tblpPr":
            fl = {}
            for k, v in c.attrib.items():
                name = local_attr(k)
                if name in ("leftFromText", "rightFromText", "topFromText", "bottomFromText",
                            "tblpX", "tblpY"):
                    fl[name] = twips(v, 0)
                else:
                    fl[name] = v
            d["float"] = fl
        elif t == "bidiVisual":
            d["bidi"] = onoff(c)
    return d


def parse_trpr(el) -> dict:
    d: dict = {}
    if el is None:
        return d
    for c in el:
        t = local(c)
        if t == "trHeight":
            d["height"] = twips(attr(c, "val"), 0)
            d["hRule"] = attr(c, "hRule", "atLeast")
        elif t == "tblHeader":
            d["header"] = onoff(c)
        elif t == "cantSplit":
            d["cantSplit"] = onoff(c)
        elif t == "gridBefore":
            d["gridBefore"] = int(num(attr(c, "val"), 0))
        elif t == "gridAfter":
            d["gridAfter"] = int(num(attr(c, "val"), 0))
        elif t == "jc":
            d["jc"] = attr(c, "val")
        elif t == "hidden":
            d["hidden"] = onoff(c)
    return d


def parse_tcpr(el) -> dict:
    d: dict = {}
    if el is None:
        return d
    for c in el:
        t = local(c)
        if t == "tcW":
            d["tcW"] = parse_width(c)
        elif t == "gridSpan":
            d["gridSpan"] = int(num(attr(c, "val"), 1))
        elif t == "hMerge":
            d["hMerge"] = attr(c, "val", "continue")
        elif t == "vMerge":
            d["vMerge"] = attr(c, "val", "continue")
        elif t == "tcBorders":
            d["tcBorders"] = parse_borders(c)
        elif t == "shd":
            d["shd"] = parse_shd(c)
        elif t == "tcMar":
            d["tcMar"] = parse_cell_margins(c)
        elif t == "vAlign":
            d["vAlign"] = attr(c, "val", "top")
        elif t == "textDirection":
            d["textDirection"] = attr(c, "val")
        elif t == "noWrap":
            d["noWrap"] = onoff(c)
    return d


_DICT_MERGE = {"pBdr", "tblBorders", "tcBorders", "cellMar", "tcMar", "look"}


def merge(base: dict, over: dict) -> dict:
    if not over:
        return base
    if not base:
        return dict(over)
    out = dict(base)
    for k, v in over.items():
        if k == "tabs":
            cur = {pos: (val, pos, leader) for val, pos, leader in out.get("tabs", [])}
            for val, pos, leader in v:
                if val == "clear":
                    # clear removes an inherited stop at (about) this position
                    for p in list(cur):
                        if abs(p - pos) < 0.5:
                            del cur[p]
                else:
                    cur[pos] = (val, pos, leader)
            out["tabs"] = sorted(cur.values(), key=lambda t: t[1])
        elif k in _DICT_MERGE and isinstance(v, dict):
            m = dict(out.get(k) or {})
            m.update(v)
            out[k] = m
        elif k == "mark_rpr":
            out[k] = merge(out.get(k) or {}, v)
        else:
            out[k] = v
    # An explicit font name at a more specific level beats an inherited theme font.
    for slot in ("ascii", "hAnsi", "eastAsia", "cs"):
        if "font_" + slot in over and "theme_" + slot not in over:
            out.pop("theme_" + slot, None)
    return out
