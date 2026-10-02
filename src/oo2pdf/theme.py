"""DrawingML theme parsing and Office colour arithmetic."""
from __future__ import annotations

import colorsys

from .opc import A

Color = tuple[float, float, float]  # r, g, b in 0..1

BLACK: Color = (0.0, 0.0, 0.0)
WHITE: Color = (1.0, 1.0, 1.0)

_THEME_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
                "accent4", "accent5", "accent6", "hlink", "folHlink"]


def hex_color(value: str | None) -> Color | None:
    if not value:
        return None
    v = value.strip().lstrip("#")
    if len(v) == 8:  # ARGB
        v = v[2:]
    if len(v) != 6 or v.lower() == "auto":
        return None
    try:
        return (int(v[0:2], 16) / 255, int(v[2:4], 16) / 255, int(v[4:6], 16) / 255)
    except ValueError:
        return None


def apply_tint(color: Color, tint: float) -> Color:
    """Excel-style tint (-1..1) applied to HLS luminance."""
    if not tint:
        return color
    h, l, s = colorsys.rgb_to_hls(*color)
    if tint < 0:
        l = l * (1 + tint)
    else:
        l = l * (1 - tint) + tint
    return colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)


def apply_lum(color: Color, mod: float | None, off: float | None) -> Color:
    h, l, s = colorsys.rgb_to_hls(*color)
    if mod is not None:
        l *= mod
    if off is not None:
        l += off
    return colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)


class Theme:
    def __init__(self, root=None):
        self.colors: dict[str, Color] = {}
        self.major = {"latin": "Calibri Light", "ea": "", "cs": ""}
        self.minor = {"latin": "Calibri", "ea": "", "cs": ""}
        if root is None:
            self.colors = {
                "dk1": BLACK, "lt1": WHITE, "dk2": hex_color("44546A"), "lt2": hex_color("E7E6E6"),
                "accent1": hex_color("4472C4"), "accent2": hex_color("ED7D31"),
                "accent3": hex_color("A5A5A5"), "accent4": hex_color("FFC000"),
                "accent5": hex_color("5B9BD5"), "accent6": hex_color("70AD47"),
                "hlink": hex_color("0563C1"), "folHlink": hex_color("954F72"),
            }
            return
        scheme = root.find(f".//{A}clrScheme")
        if scheme is not None:
            for child in scheme:
                name = child.tag.rsplit("}", 1)[-1]
                col = None
                for c in child:
                    tag = c.tag.rsplit("}", 1)[-1]
                    if tag == "srgbClr":
                        col = hex_color(c.get("val"))
                    elif tag == "sysClr":
                        col = hex_color(c.get("lastClr")) or (BLACK if c.get("val") == "windowText" else WHITE)
                if col:
                    self.colors[name] = col
        fs = root.find(f".//{A}fontScheme")
        if fs is not None:
            for which, target in (("majorFont", self.major), ("minorFont", self.minor)):
                el = fs.find(f"{A}{which}")
                if el is None:
                    continue
                for slot in ("latin", "ea", "cs"):
                    s = el.find(f"{A}{slot}")
                    if s is not None and s.get("typeface") is not None:
                        target[slot] = s.get("typeface")

    def color(self, name: str | None) -> Color | None:
        if name is None:
            return None
        aliases = {"dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2",
                   "text1": "dk1", "background1": "lt1", "text2": "dk2", "background2": "lt2",
                   "tx1": "dk1", "bg1": "lt1", "tx2": "dk2", "bg2": "lt2",
                   "hyperlink": "hlink", "followedHyperlink": "folHlink"}
        return self.colors.get(aliases.get(name, name))

    def indexed_theme(self, idx: int) -> Color | None:
        """Theme colour by SpreadsheetML index (0=lt1, 1=dk1, ...)."""
        if 0 <= idx < len(_THEME_ORDER):
            return self.colors.get(_THEME_ORDER[idx])
        return None

    def font(self, theme_ref: str | None) -> str | None:
        """Resolve a WordprocessingML theme font reference such as 'minorHAnsi'."""
        if not theme_ref:
            return None
        src = self.major if theme_ref.startswith("major") else self.minor
        ref = theme_ref[5:]
        if ref in ("HAnsi", "Ascii"):
            return src["latin"] or None
        if ref == "EastAsia":
            return src["ea"] or src["latin"] or None
        if ref == "Bidi":
            return src["cs"] or src["latin"] or None
        return src["latin"] or None


def drawingml_color(el, theme: Theme) -> Color | None:
    """Read a DrawingML colour choice (a:srgbClr, a:schemeClr ...) with modifiers."""
    if el is None:
        return None
    for c in el:
        tag = c.tag.rsplit("}", 1)[-1]
        col = None
        if tag == "srgbClr":
            col = hex_color(c.get("val"))
        elif tag == "schemeClr":
            col = theme.color(c.get("val"))
        elif tag == "sysClr":
            col = hex_color(c.get("lastClr"))
        elif tag == "prstClr":
            col = {"black": BLACK, "white": WHITE, "red": (1, 0, 0), "green": (0, 0.5, 0),
                   "blue": (0, 0, 1), "yellow": (1, 1, 0)}.get(c.get("val"), BLACK)
        if col is None:
            continue
        mod = off = None
        alpha = 1.0
        for m in c:
            mt = m.tag.rsplit("}", 1)[-1]
            v = int(m.get("val", "100000")) / 100000
            if mt == "lumMod":
                mod = v
            elif mt == "lumOff":
                off = v
            elif mt == "tint":
                col = tuple(1 - (1 - x) * v for x in col)
            elif mt == "shade":
                col = tuple(x * v for x in col)
            elif mt == "alpha":
                alpha = v
        if mod is not None or off is not None:
            col = apply_lum(col, mod, off)
        if alpha < 0.05:
            return None
        return col
    return None
