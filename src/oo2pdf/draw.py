"""Device-independent drawing operations and the PDF backend.

Layout code emits ops in a top-left origin coordinate system measured in
points; the backend flips them into PDF space.
"""
from __future__ import annotations

import io
import math
from dataclasses import dataclass, field, replace

from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as rl_canvas

from .fonts import FontFace
from .theme import Color


@dataclass(slots=True)
class TextOp:
    x: float
    y: float  # baseline
    text: str
    face: FontFace
    size: float
    color: Color = (0, 0, 0)
    char_space: float = 0.0
    hscale: float = 100.0
    angle: float = 0.0  # degrees, counter-clockwise around (x, y)

    def moved(self, dx, dy):
        return replace(self, x=self.x + dx, y=self.y + dy)


@dataclass(slots=True)
class LineOp:
    x1: float
    y1: float
    x2: float
    y2: float
    width: float = 0.5
    color: Color = (0, 0, 0)
    dash: tuple | None = None
    cap: int = 0

    def moved(self, dx, dy):
        return replace(self, x1=self.x1 + dx, y1=self.y1 + dy, x2=self.x2 + dx, y2=self.y2 + dy)


@dataclass(slots=True)
class RectOp:
    x: float
    y: float
    w: float
    h: float
    fill: Color | None = None
    stroke: Color | None = None
    width: float = 0.5
    radius: float = 0.0
    ellipse: bool = False

    def moved(self, dx, dy):
        return replace(self, x=self.x + dx, y=self.y + dy)


@dataclass(slots=True)
class ImageOp:
    x: float
    y: float
    w: float
    h: float
    data: bytes
    crop: tuple | None = None  # (l, t, r, b) fractions
    angle: float = 0.0

    def moved(self, dx, dy):
        return replace(self, x=self.x + dx, y=self.y + dy)


@dataclass(slots=True)
class LinkOp:
    x: float
    y: float
    w: float
    h: float
    url: str | None = None
    anchor: str | None = None

    def moved(self, dx, dy):
        return replace(self, x=self.x + dx, y=self.y + dy)


@dataclass(slots=True)
class ClipOp:
    x: float
    y: float
    w: float
    h: float
    ops: list = field(default_factory=list)

    def moved(self, dx, dy):
        return ClipOp(self.x + dx, self.y + dy, self.w, self.h, [o.moved(dx, dy) for o in self.ops])


@dataclass(slots=True)
class PolyOp:
    """A polyline/polygon through ``points`` [(x, y), ...]."""
    points: list
    fill: Color | None = None
    stroke: Color | None = None
    width: float = 0.75
    closed: bool = False
    dash: tuple | None = None

    def moved(self, dx, dy):
        return PolyOp([(x + dx, y + dy) for x, y in self.points], self.fill, self.stroke, self.width,
                      self.closed, self.dash)


@dataclass(slots=True)
class TransformOp:
    """Draw ``ops`` through an affine map in top-left page space:
    (u, v) -> (a*u + c*v + e, b*u + d*v + f)."""
    matrix: tuple
    ops: list = field(default_factory=list)

    def moved(self, dx, dy):
        a, b, c, d, e, f = self.matrix
        return TransformOp((a, b, c, d, e + dx, f + dy), self.ops)


@dataclass(slots=True)
class AnchorOp:
    """A named destination (bookmark) for internal links."""
    x: float
    y: float
    name: str

    def moved(self, dx, dy):
        return replace(self, x=self.x + dx, y=self.y + dy)


def move_ops(ops, dx, dy):
    if not dx and not dy:
        return list(ops)
    return [o.moved(dx, dy) for o in ops]


class PdfWriter:
    def __init__(self, target, title: str | None = None, author: str | None = None):
        self.c = rl_canvas.Canvas(target, pageCompression=1)
        self.c.setCreator("oo2pdf")
        if title:
            self.c.setTitle(title)
        if author:
            self.c.setAuthor(author)
        self.h = 0.0
        self._images: dict[int, ImageReader | None] = {}
        self._anchors: set[str] = set()
        self._pending_links: list = []
        self._page_no = 0

    # -- helpers -----------------------------------------------------------
    def _y(self, y):
        return self.h - y

    def _image(self, data: bytes):
        key = id(data)
        if key not in self._images:
            self._images[key] = _load_image(data)
        return self._images[key]

    # -- page lifecycle ------------------------------------------------------
    def page(self, width: float, height: float, ops):
        self.h = height
        self.c.setPageSize((width, height))
        self._page_no += 1
        self._draw(ops)
        self.c.showPage()

    def save(self):
        self.c.save()

    # -- drawing -------------------------------------------------------------
    def _draw(self, ops):
        c = self.c
        for op in ops:
            t = type(op)
            if t is TextOp:
                self._text(op)
            elif t is RectOp:
                if op.fill is None and op.stroke is None:
                    continue
                if op.fill is not None:
                    c.setFillColorRGB(*op.fill)
                if op.stroke is not None:
                    c.setStrokeColorRGB(*op.stroke)
                    c.setLineWidth(op.width)
                    c.setDash()
                if op.ellipse:
                    c.ellipse(op.x, self._y(op.y + op.h), op.x + op.w, self._y(op.y),
                              stroke=int(op.stroke is not None), fill=int(op.fill is not None))
                elif op.radius:
                    c.roundRect(op.x, self._y(op.y + op.h), op.w, op.h, op.radius,
                                stroke=int(op.stroke is not None), fill=int(op.fill is not None))
                else:
                    c.rect(op.x, self._y(op.y + op.h), op.w, op.h,
                           stroke=int(op.stroke is not None), fill=int(op.fill is not None))
            elif t is LineOp:
                c.setStrokeColorRGB(*op.color)
                c.setLineWidth(op.width)
                c.setLineCap(op.cap)
                if op.dash:
                    c.setDash(list(op.dash))
                else:
                    c.setDash()
                c.line(op.x1, self._y(op.y1), op.x2, self._y(op.y2))
            elif t is ImageOp:
                self._draw_image(op)
            elif t is LinkOp:
                rect = (op.x, self._y(op.y + op.h), op.x + op.w, self._y(op.y))
                if op.url:
                    c.linkURL(op.url, rect, relative=0, thickness=0)
                elif op.anchor:
                    self._pending_links.append((op.anchor, rect))
                    c.linkRect("", op.anchor, rect, relative=0, thickness=0)
            elif t is AnchorOp:
                if op.name not in self._anchors:
                    self._anchors.add(op.name)
                    c.bookmarkHorizontal(op.name, op.x, self._y(op.y))
            elif t is ClipOp:
                c.saveState()
                p = c.beginPath()
                p.rect(op.x, self._y(op.y + op.h), op.w, op.h)
                c.clipPath(p, stroke=0, fill=0)
                self._draw(op.ops)
                c.restoreState()
            elif t is PolyOp:
                if len(op.points) < 2 or (op.fill is None and op.stroke is None):
                    continue
                p = c.beginPath()
                x0, y0 = op.points[0]
                p.moveTo(x0, self._y(y0))
                for x, y in op.points[1:]:
                    p.lineTo(x, self._y(y))
                if op.closed:
                    p.close()
                if op.fill is not None:
                    c.setFillColorRGB(*op.fill)
                if op.stroke is not None:
                    c.setStrokeColorRGB(*op.stroke)
                    c.setLineWidth(op.width)
                    c.setDash(list(op.dash) if op.dash else [])
                c.drawPath(p, stroke=int(op.stroke is not None), fill=int(op.fill is not None))
            elif t is TransformOp:
                a, b, cc, d, e, f = op.matrix
                H = self.h
                c.saveState()
                c.transform(a, -b, -cc, d, cc * H + e, H - d * H - f)
                self._draw(op.ops)
                c.restoreState()

    def _text(self, op: TextOp):
        if not op.text:
            return
        c = self.c
        face = op.face
        c.saveState()
        c.setFillColorRGB(*op.color)
        t = c.beginText()
        t.setFont(face.font.name, op.size)
        x, y = op.x, self._y(op.y)
        skew = 0.21 if face.synth_italic else 0.0
        if op.angle:
            a = math.radians(op.angle)
            ca, sa = math.cos(a), math.sin(a)
            t.setTextTransform(ca, sa, -sa + skew * ca, ca + skew * sa, x, y)
        elif skew:
            t.setTextTransform(1, 0, skew, 1, x, y)
        else:
            t.setTextOrigin(x, y)
        if op.char_space:
            t.setCharSpace(op.char_space)
        if op.hscale != 100:
            t.setHorizScale(op.hscale)
        if face.synth_bold:
            t.setTextRenderMode(2)
            c.setStrokeColorRGB(*op.color)
            c.setLineWidth(op.size * 0.03)
        t.textOut(op.text)
        c.drawText(t)
        c.restoreState()

    def _draw_image(self, op: ImageOp):
        img = self._image(op.data)
        if img is None:
            return
        c = self.c
        c.saveState()
        x, y, w, h = op.x, op.y, op.w, op.h
        if op.angle:
            cx, cy = x + w / 2, self._y(y + h / 2)
            c.translate(cx, cy)
            c.rotate(-op.angle)
            c.translate(-cx, -cy)
        if op.crop and any(op.crop):
            l, tp, r, b = op.crop
            p = c.beginPath()
            p.rect(x, self._y(y + h), w, h)
            c.clipPath(p, stroke=0, fill=0)
            fw = 1 - l - r
            fh = 1 - tp - b
            if fw > 0 and fh > 0:
                full_w, full_h = w / fw, h / fh
                x -= l * full_w
                y -= tp * full_h
                w, h = full_w, full_h
        try:
            c.drawImage(img, x, self._y(y + h), w, h, mask="auto")
        except Exception:
            pass
        c.restoreState()


def _load_image(data: bytes):
    from PIL import Image
    try:
        im = Image.open(io.BytesIO(data))
        if im.format in ("WMF", "EMF"):
            try:
                im.load(dpi=144)
            except TypeError:
                im.load()
        if im.mode in ("P", "LA", "PA") or (im.mode == "RGBA"):
            im = im.convert("RGBA")
        elif im.mode not in ("RGB", "L", "RGBA"):
            im = im.convert("RGB")
        if im.format not in ("JPEG", "PNG") or im.mode == "RGBA":
            buf = io.BytesIO()
            im.save(buf, "PNG")
            buf.seek(0)
            return ImageReader(buf)
        return ImageReader(io.BytesIO(data))
    except Exception:
        return None


