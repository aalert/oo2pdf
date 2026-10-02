"""Legacy VML drawings (v:group canvases, v:rect, v:shape, v:line ...).

Word 2007 and earlier store drawing canvases as VML groups: children are
placed in the group's coordinate space (coordorigin/coordsize) which is scaled
onto the group's size.  They are converted into the same Shape model used for
DrawingML so the layout engine renders them.
"""
from __future__ import annotations

import re

from ..opc import R, V, W, local
from ..theme import hex_color
from .model import Picture, Shape

O_NS = "urn:schemas-microsoft-com:office:office"

_NAMED = {
    "black": (0, 0, 0), "white": (1, 1, 1), "red": (1, 0, 0), "green": (0, 0.5, 0),
    "blue": (0, 0, 1), "yellow": (1, 1, 0), "gray": (0.5, 0.5, 0.5), "silver": (0.75, 0.75, 0.75),
    "maroon": (0.5, 0, 0), "navy": (0, 0, 0.5), "olive": (0.5, 0.5, 0), "purple": (0.5, 0, 0.5),
    "teal": (0, 0.5, 0.5), "lime": (0, 1, 0), "aqua": (0, 1, 1), "fuchsia": (1, 0, 1),
}

_DASHES = {"dash": (3, 1), "shortdash": (3, 1), "dot": (1, 1), "shortdot": (1, 1), "1 1": (1, 1),
           "dashdot": (3, 1, 1, 1), "longdash": (8, 3), "longdashdot": (8, 3, 1, 3)}


def css(style: str) -> dict:
    out = {}
    for part in (style or "").split(";"):
        if ":" in part:
            k, v = part.split(":", 1)
            out[k.strip().lower()] = v.strip()
    return out


def length(v, default=0.0) -> float:
    """A CSS length in points; unitless values are returned as numbers."""
    if v is None or v == "":
        return default
    m = re.match(r"\s*(-?[\d.]+)\s*([a-z%]*)", str(v))
    if not m:
        return default
    n = float(m.group(1))
    return n * {"pt": 1.0, "in": 72.0, "cm": 28.3465, "mm": 2.83465, "px": 0.75, "pc": 12.0,
                "emu": 1 / 12700.0, "": 1.0}.get(m.group(2), 1.0)


def color(v, default=None):
    if not v:
        return default
    v = v.split()[0].strip()
    if v.startswith("#"):
        h = v[1:]
        if len(h) == 3:
            h = "".join(c * 2 for c in h)
        return hex_color(h) or default
    return _NAMED.get(v.lower(), default)


def _on(v, default=True):
    if v is None:
        return default
    return v.lower() in ("t", "true", "on", "1")


def _pair(v, default):
    try:
        a, b = (v or "").split(",")
        return float(a), float(b)
    except ValueError:
        return default


def inset(tb) -> tuple:
    vals = [7.2, 3.6, 7.2, 3.6]
    if tb is not None and tb.get("inset"):
        for i, part in enumerate(tb.get("inset").split(",")[:4]):
            if part.strip():
                vals[i] = length(part.strip())
    return tuple(vals)


class VmlReader:
    """Converts VML elements to Shape trees; needs the DocxReader for text and blobs."""

    def __init__(self, reader, part):
        self.reader = reader
        self.part = part

    def group(self, el, width: float, height: float) -> Shape:
        g = Shape(width, height, geom="group")
        ox, oy = _pair(el.get("coordorigin"), (0.0, 0.0))
        cw, ch = _pair(el.get("coordsize"), (1000.0, 1000.0))
        sx = width / cw if cw else 1.0
        sy = height / ch if ch else 1.0
        for c in el:
            name = local(c)
            if name not in ("group", "rect", "roundrect", "oval", "shape", "line", "polyline", "image"):
                continue
            st = css(c.get("style", ""))
            if name == "line":
                x1, y1 = _pair(c.get("from"), (0.0, 0.0))
                x2, y2 = _pair(c.get("to"), (0.0, 0.0))
                shp = self._line_shape(c, abs(x2 - x1) * sx, abs(y2 - y1) * sy)
                shp.flip_x = x2 < x1
                shp.flip_y = y2 < y1
                g.children.append(((min(x1, x2) - ox) * sx, (min(y1, y2) - oy) * sy, shp))
                continue
            left = length(st.get("left"))
            top = length(st.get("top"))
            w = length(st.get("width")) * sx
            h = length(st.get("height")) * sy
            dx, dy = (left - ox) * sx, (top - oy) * sy
            if name == "group":
                child = self.group(c, w, h)
            else:
                child = self.shape(c, name, w, h)
            if child is None:
                continue
            flip = st.get("flip", "")
            if isinstance(child, Shape):
                child.flip_x = "x" in flip
                child.flip_y = "y" in flip
            g.children.append((dx, dy, child))
        return g

    def _line_shape(self, c, w, h) -> Shape:
        s = Shape(w, h, geom="line", fill=None)
        s.line = color(c.get("strokecolor"), (0, 0, 0)) if _on(c.get("stroked")) else None
        s.line_width = length(c.get("strokeweight"), 0.75) if c.get("strokeweight") else 0.75
        self._stroke_details(c, s)
        return s

    def _stroke_details(self, c, s: Shape):
        st = c.find(V + "stroke")
        if st is not None:
            if st.get("dashstyle"):
                pat = _DASHES.get(st.get("dashstyle").lower())
                if pat:
                    s.dash = tuple(x * max(s.line_width, 0.5) for x in pat)
            if st.get("endarrow") and st.get("endarrow") != "none":
                s.end_arrow = st.get("endarrow")
            if st.get("startarrow") and st.get("startarrow") != "none":
                s.start_arrow = st.get("startarrow")
            if st.get("on") is not None and not _on(st.get("on")):
                s.line = None

    def shape(self, c, name, w, h):
        typ = (c.get("type") or "").lstrip("#")
        spt = c.get("{%s}spt" % O_NS)
        img = c.find(V + "imagedata")
        if img is not None:
            rid = img.get(R + "id") or img.get("{%s}relid" % O_NS)
            data = self.reader._blob(self.part, rid)
            if data:
                return Picture(data, w, h, None)
            return None
        if typ.endswith("_t75") or spt == "75":
            return None  # empty picture frame (canvas background)
        if typ.endswith("_t32") or typ.endswith("_t20") or spt in ("32", "20"):
            return self._line_shape(c, w, h)
        geom = {"rect": "rect", "roundrect": "roundRect", "oval": "ellipse"}.get(name, "rect")
        s = Shape(w, h, geom=geom)
        s.fill = color(c.get("fillcolor"), (1, 1, 1)) if _on(c.get("filled")) else None
        fill = c.find(V + "fill")
        if fill is not None and fill.get("on") is not None and not _on(fill.get("on")):
            s.fill = None
        elif fill is not None and fill.get("type") == "frame":
            s.data = self.reader._blob(self.part, fill.get(R + "id") or fill.get("{%s}relid" % O_NS))
        s.line = color(c.get("strokecolor"), (0, 0, 0)) if _on(c.get("stroked")) else None
        s.line_width = length(c.get("strokeweight"), 0.75) if c.get("strokeweight") else 0.75
        self._stroke_details(c, s)
        path = c.get("path")
        if name == "shape" and path and not typ.endswith("_t202"):
            cw, ch = _pair(c.get("coordsize"), (21600.0, 21600.0))
            pts, closed = parse_path(path)
            if pts:
                s.geom = "path"
                s.points = [(x / cw * w, y / ch * h) for x, y in pts]
                s.closed = closed
        tb = c.find(V + "textbox")
        if tb is not None:
            content = tb.find(W + "txbxContent")
            if content is not None:
                s.blocks = self.reader._text_blocks(content, self.part)
                s.insets = inset(tb)
        return s


_CURVE_STEPS = 16  # line segments per cubic Bezier


def _params(s: str) -> list[float]:
    """VML path parameters: comma separated, an empty value means 0
    ("m,1008" is 0,1008 and a trailing "2160," ends with 0)."""
    s = s.strip()
    if not s:
        return []
    return [float(v) if v else 0.0 for v in re.split(r"\s*,\s*|\s+", s)]


def _bezier(p0, p1, p2, p3):
    out = []
    for k in range(1, _CURVE_STEPS + 1):
        t = k / _CURVE_STEPS
        u = 1 - t
        out.append((u * u * u * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t * t * t * p3[0],
                    u * u * u * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t * t * t * p3[1]))
    return out


def parse_path(path: str):
    """Parse a VML path (m/l/c and their relative forms t/r/v, x, e) into
    points; cubic curves are flattened into short segments."""
    pts: list[tuple[float, float]] = []
    closed = False
    x = y = 0.0
    for cmd, args in re.findall(r"([a-z]+)([^a-z]*)", path.lower()):
        nums = _params(args)
        if cmd in ("m", "t"):
            if len(nums) >= 2:
                x, y = (x + nums[0], y + nums[1]) if cmd == "t" else (nums[0], nums[1])
                pts.append((x, y))
        elif cmd in ("l", "r"):
            for i in range(0, len(nums) - 1, 2):
                x, y = (x + nums[i], y + nums[i + 1]) if cmd == "r" else (nums[i], nums[i + 1])
                pts.append((x, y))
        elif cmd in ("c", "v"):
            for i in range(0, len(nums) - 5, 6):
                c = nums[i:i + 6]
                if cmd == "v":
                    c = [c[k] + (x if k % 2 == 0 else y) for k in range(6)]
                if not pts:
                    pts.append((x, y))
                pts.extend(_bezier((x, y), (c[0], c[1]), (c[2], c[3]), (c[4], c[5])))
                x, y = c[4], c[5]
        elif "x" in cmd:
            closed = True
        # e (end), nf/ns (no fill / no stroke) and other flags: nothing to add
    return pts, closed
