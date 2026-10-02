"""Paragraph and table layout following Word's rules.

Blocks are laid out for a given column width into *flows*:

* :class:`ParaFlow` - the lines of one paragraph plus its spacing/keep rules.
* :class:`TableFlow` - the rows of a table (rows can be split across pages).

Lines carry drawing ops relative to (column left, line top); the paginator
positions them on pages.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field

from ..draw import (AnchorOp, ClipOp, ImageOp, LineOp, LinkOp, PolyOp, RectOp, TextOp, TransformOp,
                    move_ops)
from ..fonts import FontFace, FontManager
from .model import (Bookmark, Break, Field, Floating, FootnoteRef, MathItem, Paragraph, Picture,
                    RunStyle, Separator, Shape, Table, Tab, Text)
from .numbering import format_number
from .props import Border

EPS = 0.01
AUTO_SPACING = 14.0
# how far Word's justification may squeeze a line to fit another word
JUSTIFY_MODE = "em"
JUSTIFY_SHRINK = 1.0


# --------------------------------------------------------------------------- context
class LayoutContext:
    def __init__(self, fonts: FontManager, reader):
        self.fonts = fonts
        self.reader = reader
        self.sum_spacing = getattr(reader, "sum_spacing", False)
        self.default_tab = reader.default_tab if reader else 36.0
        self.fields: dict[str, str] = {}
        self.note_labels: dict = {}
        self.note_label: str = ""
        self.dark_bg: bool = False
        self._face_cache: dict = {}

    def face(self, family, bold, italic) -> FontFace:
        key = (family, bold, italic)
        f = self._face_cache.get(key)
        if f is None:
            f = self.fonts.face(family, bold, italic)
            self._face_cache[key] = f
        return f


# --------------------------------------------------------------------------- flows
@dataclass
class Line:
    height: float
    ops: list
    floats: list = field(default_factory=list)  # (Floating, x) pairs
    notes: list = field(default_factory=list)  # FootnoteRef
    break_after: str | None = None  # page | column
    baseline: float = 0.0
    trailing: float = 0.0  # extra leading of "multiple" line spacing below the text
    para: tuple | None = None  # (paragraph id, line index, line count, widow control) in a cell


@dataclass
class ParaFlow:
    lines: list
    space_before: float
    space_after: float
    keep_next: bool = False
    keep_lines: bool = False
    widow: bool = True
    page_break_before: bool = False
    style_id: str | None = None
    contextual: bool = False
    is_empty: bool = False
    relayout: object = None  # callable(bounds) -> ParaFlow, for text wrapping


@dataclass
class CellBox:
    x: float
    w: float
    mar: dict
    lines: list  # (y, Line)
    content_h: float
    valign: str
    fill: tuple | None
    borders: dict  # resolved top/bottom/left/right -> Border|None
    is_continue: bool = False
    span_rows: int = 1
    span_height: float = 0.0  # total height of the vertical span (for vAlign)
    span_last: bool = True
    c0: int = 0
    c1: int = 0
    own: dict | None = None  # the cell's own borders before shared edges were merged


@dataclass
class RowFlow:
    cells: list
    height: float
    header: bool = False
    cant_split: bool = False
    exact: bool = False
    table_x: float = 0.0
    first: bool = False
    last: bool = False
    band_top: float = 0.0  # Word reserves the width of horizontal borders
    band_bottom: float = 0.0
    min_height: float = 0.0  # an "at least" trHeight (with margins/band)

    def as_line(self) -> Line:
        return Line(self.height, self.render(), notes=self.notes(), floats=self.floats())

    def notes(self):
        out = []
        for c in self.cells:
            for _, ln in c.lines:
                out.extend(ln.notes)
        return out

    def floats(self):
        out = []
        for c in self.cells:
            for y, ln in c.lines:
                for fl, x in ln.floats:
                    out.append((fl, x + c.x + c.mar.get("left", 0)))
        return out

    def render(self) -> list:
        ops: list = []
        h = self.height
        for c in self.cells:
            if c.fill is not None:
                if c.is_continue:
                    # a vertically merged cell reads as one area: overlap the piece
                    # above so viewers show no seam between the rows
                    ops.append(RectOp(c.x, -0.5, c.w, h + 0.5, fill=c.fill))
                else:
                    ops.append(RectOp(c.x, 0, c.w, h, fill=c.fill))
        for c in self.cells:
            if c.is_continue:
                continue
            span_h = c.span_height or h
            avail = span_h - self.band_top - c.mar.get("top", 0) - c.mar.get("bottom", 0)
            if c.span_rows <= 1:
                avail -= self.band_bottom
            off = 0.0
            if c.valign == "center":
                off = max(0.0, (avail - c.content_h) / 2)
            elif c.valign == "bottom":
                off = max(0.0, avail - c.content_h)
            cx = c.x + c.mar.get("left", 0)
            cy = self.band_top + c.mar.get("top", 0) + off
            cell_ops = []
            for y, ln in c.lines:
                cell_ops.extend(move_ops(ln.ops, cx, cy + y))
            if self.exact:
                ops.append(ClipOp(c.x, 0, c.w, h, cell_ops))
            else:
                ops.extend(cell_ops)
        for c in self.cells:
            b = c.borders
            x0, x1 = c.x, c.x + c.w
            lb, rb = b.get("left"), b.get("right")
            # Word joins borders like this: a horizontal border runs over the
            # vertical ones it meets (to their outer edges) and a vertical
            # border stops at the horizontal ones (sample-files table-document)
            lw = lb.width / 2 if lb is not None and lb.visible else 0.0
            rw = rb.width / 2 if rb is not None and rb.visible else 0.0
            y0, y1 = 0.0, h
            t = b.get("top")
            if t is not None and t.visible and not c.is_continue:
                _border_line(ops, t, x0 - lw, t.width / 2, x1 + rw, t.width / 2, horizontal=True, cap=0)
                y0 = t.width
            bb = b.get("bottom")
            if self.band_bottom and bb is not None and bb.visible and c.span_last:
                _border_line(ops, bb, x0 - lw, h - bb.width / 2, x1 + rw, h - bb.width / 2, horizontal=True, cap=0)
                y1 = h - bb.width
            _border_line(ops, lb, x0, y0, x0, y1, horizontal=False, cap=0)
            _border_line(ops, rb, x1, y0, x1, y1, horizontal=False, cap=0)
        return ops

    def split(self, avail: float):
        """Split the row so the first part is at most ``avail`` high."""
        if self.cant_split or self.exact:
            return None
        first_cells, rest_cells = [], []
        any_first = False
        any_rest = False
        for c in self.cells:
            top = c.mar.get("top", 0)
            bot = c.mar.get("bottom", 0)
            limit = avail - self.band_top - top - bot
            k = 0
            # a paragraph that ends in the first part needs its space after too
            # (Word moves a row whose one-line cells fit only without it, but
            # carries on a paragraph whose middle line fits: large-document B.3/C.4)
            def after(ln):
                return ln.para[3] if ln.para and ln.para[1] == ln.para[2] - 1 else 0.0
            while k < len(c.lines) and c.lines[k][0] + c.lines[k][1].height                     + after(c.lines[k][1]) <= limit + EPS:
                k += 1
            if k == 0 and c.lines and not c.is_continue:
                return None  # every cell must start on this page, or the row moves
            a, b = list(c.lines[:k]), list(c.lines[k:])
            if a and not c.is_continue:
                any_first = True
            if b:
                any_rest = True
            a_h = (a[-1][0] + a[-1][1].height) if a else 0.0
            shift = b[0][0] if b else 0.0
            b = [(y - shift, ln) for y, ln in b]
            # the rest keeps the remaining cell height, the space after the last
            # paragraph included (Word: large-document.docx 7.2, page 16)
            b_h = max((b[-1][0] + b[-1][1].height), c.content_h - shift) if b else 0.0
            first_cells.append(CellBox(c.x, c.w, c.mar, a, a_h, "top", c.fill, dict(c.borders),
                                       c.is_continue, c0=c.c0, c1=c.c1))
            rest_cells.append(CellBox(c.x, c.w, c.mar, b, b_h, "top", c.fill, dict(c.borders),
                                      c.is_continue, span_last=c.span_last, c0=c.c0, c1=c.c1))
        if not any_first or not any_rest:
            return None
        rest_h = self.band_top + self.band_bottom + max(
            c.mar.get("top", 0) + c.content_h + c.mar.get("bottom", 0) for c in rest_cells)
        first = RowFlow(first_cells, avail, self.header, self.cant_split, False, self.table_x,
                        self.first, False, self.band_top, 0.0)
        rest = RowFlow(rest_cells, rest_h, self.header, self.cant_split, False, self.table_x,
                       False, self.last, self.band_top, self.band_bottom)
        return first, rest


@dataclass
class TableFlow:
    rows: list
    header_rows: int = 0
    keep_next: bool = False
    float: dict | None = None
    x0: float = 0.0
    width: float = 0.0


def _border_line(ops, b: Border | None, x0, y0, x1, y1, horizontal, cap=2):
    """Draw one border edge. Square caps close box corners (paragraph and page
    borders); table cells pass flat caps and join their edges explicitly."""
    if b is None or not b.visible:
        return
    color = b.color or (0, 0, 0)
    w = b.width
    dash = None
    if b.style in ("dotted",):
        dash = (w, w * 2)
    elif b.style in ("dashed", "dashSmallGap"):
        dash = (w * 4, w * 2)
    elif b.style in ("dotDash",):
        dash = (w * 4, w * 2, w, w * 2)
    if b.style in ("double", "thinThickSmallGap", "thickThinSmallGap"):
        lw = max(0.25, w / 3)
        gap = w / 3 + lw / 2
        if horizontal:
            ops.append(LineOp(x0, y0 - gap, x1, y1 - gap, lw, color, dash))
            ops.append(LineOp(x0, y0 + gap, x1, y1 + gap, lw, color, dash))
        else:
            ops.append(LineOp(x0 - gap, y0, x1 - gap, y1, lw, color, dash))
            ops.append(LineOp(x0 + gap, y0, x1 + gap, y1, lw, color, dash))
        return
    ops.append(LineOp(x0, y0, x1, y1, w, color, dash, cap=cap if not dash else 0))


# --------------------------------------------------------------------------- atoms
class Atom:
    __slots__ = ("kind", "text", "face", "size", "style", "width", "asc", "desc", "gap",
                 "shift", "link", "brk", "payload", "char_space", "hscale", "x", "spaces", "kerns",
                 "hang")

    def __init__(self, kind, text="", face=None, size=0.0, style=None, width=0.0, asc=0.0,
                 desc=0.0, gap=0.0, shift=0.0, link=None, brk=False, payload=None,
                 char_space=0.0, hscale=100.0):
        self.kind = kind
        self.text = text
        self.face = face
        self.size = size
        self.style = style
        self.width = width
        self.asc = asc
        self.desc = desc
        self.gap = gap
        self.shift = shift
        self.link = link
        self.brk = brk
        self.payload = payload
        self.char_space = char_space
        self.hscale = hscale
        self.x = 0.0
        self.spaces = 0
        self.kerns = None  # [(char index, dx pt)] pair kerning before that char
        self.hang = 0.0  # trailing width that may run past the margin at a line end


def char_slot(ch: str, hint: str | None) -> str:
    o = ord(ch)
    if o < 0x80:
        return "ascii"
    if (0x1100 <= o <= 0x11FF or 0x2E80 <= o <= 0x9FFF or 0xA960 <= o <= 0xA97F
            or 0xAC00 <= o <= 0xD7FF or 0xF900 <= o <= 0xFAFF or 0xFE30 <= o <= 0xFE4F
            or 0xFF00 <= o <= 0xFFEF or o >= 0x20000):
        return "ea"
    if (0x0590 <= o <= 0x08FF or 0xFB1D <= o <= 0xFDFF or 0xFE70 <= o <= 0xFEFF
            or 0x0900 <= o <= 0x0DFF or 0x0E00 <= o <= 0x0EFF):
        return "cs"
    if hint == "eastAsia" and (0x2000 <= o <= 0x2BFF or 0x00A1 <= o <= 0x00FF):
        return "ea"
    return "hansi"



_NO_BREAK_BEFORE = set("、。，．：；？！）」』】〕〉》”’ー々ゝゞ・")


class AtomBuilder:
    def __init__(self, ctx: LayoutContext):
        self.ctx = ctx

    def style_font(self, style: RunStyle, slot: str):
        if slot == "ascii":
            fam, size, b, i = style.font_ascii, style.size, style.bold, style.italic
        elif slot == "ea":
            fam, size, b, i = style.font_ea, style.size, style.bold, style.italic
        elif slot == "cs":
            fam, size, b, i = style.font_cs, style.size_cs, style.bold_cs, style.italic_cs
        else:
            fam, size, b, i = style.font_hansi, style.size, style.bold, style.italic
        return fam, size, b, i

    def metrics_atom(self, style: RunStyle) -> Atom:
        """Zero-width atom carrying the line metrics of a run style (paragraph mark)."""
        fam, size, b, i = self.style_font(style, "ascii")
        face = self.ctx.face(fam, b, i)
        f = face.font
        return Atom("mark", face=face, size=size, style=style,
                    asc=f.ascent * size, desc=f.descent * size, gap=f.ext_leading * size)

    def text(self, text: str, style: RunStyle, link, out: list):
        if style.caps:
            text = text.upper()
        # group by font slot
        groups: list[tuple[str, str]] = []
        cur_slot = None
        buf: list[str] = []
        for ch in text:
            s = char_slot(ch, style.hint) if not ch.isspace() else (cur_slot or "ascii")
            if s != cur_slot and buf:
                groups.append((cur_slot, "".join(buf)))
                buf = []
            cur_slot = s
            buf.append(ch)
        if buf:
            groups.append((cur_slot, "".join(buf)))
        for slot, chunk in groups:
            fam, size, b, i = self.style_font(style, slot)
            face = self.ctx.face(fam, b, i)
            for seg, sface in self.ctx.fonts.segments(chunk, face, b, i):
                if style.small_caps:
                    for part, is_lower in _split_lower(seg):
                        if is_lower:
                            self._emit(part.upper(), sface, size, style, link, out, size * 0.8)
                        else:
                            self._emit(part, sface, size, style, link, out, size)
                else:
                    self._emit(seg, sface, size, style, link, out, size)

    def _emit(self, text, face, base_size, style, link, out, size):
        f = face.font
        shift = style.position
        draw_size = size
        if style.vert == "superscript":
            draw_size = size * 0.58
            shift += size * 0.33
        elif style.vert == "subscript":
            draw_size = size * 0.58
            shift -= size * 0.14
        asc = f.ascent * draw_size + shift
        desc = f.descent * draw_size - shift
        if style.vert:
            # sub/superscript use the metrics of the full-size font for the line
            asc = max(asc, f.ascent * size) if style.vert == "superscript" else f.ascent * size
            desc = max(desc, f.descent * size)
        gap = f.ext_leading * size
        if style.math_script:
            asc = desc = gap = 0.0  # equation scripts don't size the line
        hs = style.scale
        cs = style.spacing
        kern = bool(style.kern) and size >= style.kern
        prev = None
        for tok, kind in _tokens(text):
            if kind == "space":
                w = f.width(tok, draw_size)
                a = Atom("space", tok, face, draw_size, style, w, asc, desc, gap, shift, link, brk=True)
                a.spaces = len(tok)
            else:
                w = f.width(tok, draw_size)
                a = Atom("text", tok, face, draw_size, style, w, asc, desc, gap, shift, link,
                         brk=kind == "brk")
            if hs != 100:
                a.width *= hs / 100.0
                a.hscale = hs
            if cs:
                a.width += cs * len(tok)
                a.char_space = cs
            if kind == "cjk" and out and out[-1].kind == "text" and tok not in _NO_BREAK_BEFORE:
                out[-1].brk = True
            if kind == "cjk":
                a.brk = True
            if kind == "text" and tok[:1] in _NO_BREAK_BEFORE and out:
                out[-1].brk = False
            if kern:
                ks = []
                for i in range(1, len(tok)):
                    k = f.kern(tok[i - 1], tok[i])
                    if k:
                        ks.append((i, k * draw_size))
                if ks:
                    a.kerns = ks
                    a.width += sum(k for _, k in ks)
                if prev is not None and prev.text:
                    k = f.kern(prev.text[-1], tok[0])
                    if k:
                        # the adjustment belongs to the end of the previous atom
                        prev.kerns = (prev.kerns or []) + [(len(prev.text), k * draw_size)]
                        prev.width += k * draw_size
                prev = a
            out.append(a)


_TOKEN_RE = re.compile(r"( +)|([^ \-‐‒–—/]*[\-‐‒–—]+)|"
                       r"([⺀-鿿가-퟿豈-﫿＀-￯　-〿])|"
                       r"([^ \-‐‒–—⺀-鿿가-퟿豈-﫿＀-￯　-〿]+)")


def _tokens(text: str):
    """Split into (token, kind): kind is space | brk (word ending in hyphen) | cjk | word."""
    for m in _TOKEN_RE.finditer(text):
        if m.group(1):
            yield m.group(1), "space"
        elif m.group(2):
            g = m.group(2)
            # a lone dash between spaces is not a break point before it, but after
            yield g, "brk"
        elif m.group(3):
            yield m.group(3), "cjk"
        elif m.group(4):
            yield m.group(4), "word"


def _split_lower(text: str):
    out = []
    buf = []
    cur = None
    for ch in text:
        low = ch.islower()
        if cur is not None and low != cur:
            out.append(("".join(buf), cur))
            buf = []
        cur = low
        buf.append(ch)
    if buf:
        out.append(("".join(buf), cur))
    return out


# --------------------------------------------------------------------------- paragraph layout
class _LineBuilder:
    def __init__(self, start: float, right: float, first: bool):
        self.start = start
        self.right = right
        self.first = first
        self.atoms: list[Atom] = []
        self.x = start
        self.pending = None  # (tab_atom, stop_pos, type, seg_start_index, tab_x)
        self.skip = 0.0  # vertical space left empty above this line

    @property
    def empty(self):
        return not any(a.kind not in ("mark", "anchor", "float") for a in self.atoms)

    def add(self, a: Atom):
        self.atoms.append(a)
        self.x += a.width
        if self.pending:
            self._update_pending()

    def _seg_width(self, start):
        return sum(a.width for a in self.atoms[start:])

    def _update_pending(self):
        tab, stop, typ, seg_start, tab_x = self.pending
        seg = self._seg_width(seg_start)
        if typ == "right":
            fill = max(0.0, stop - tab_x - seg)
        elif typ == "center":
            fill = max(0.0, stop - tab_x - seg / 2)
        else:  # decimal
            before = 0.0
            for a in self.atoms[seg_start:]:
                if a.kind == "text" and "." in a.text:
                    idx = a.text.index(".")
                    before += a.face.font.width(a.text[:idx], a.size) if a.face else 0
                    break
                before += a.width
            fill = max(0.0, stop - tab_x - before)
        self.x += fill - tab.width
        tab.width = fill

    def end_x_if(self, extra_w: float) -> float:
        """x position after adding content of width ``extra_w``."""
        if not self.pending:
            return self.x + extra_w
        tab, stop, typ, seg_start, tab_x = self.pending
        seg = self._seg_width(seg_start) + extra_w
        if typ == "right":
            return tab_x + max(seg, stop - tab_x)
        if typ == "center":
            return tab_x + max(stop - tab_x - seg / 2, 0) + seg
        return tab_x + max(stop - tab_x, 0) + seg


class ParagraphLayout:
    def __init__(self, para: Paragraph, width: float, ctx: LayoutContext,
                 border_prev_same=False, border_next_same=False):
        self.p = para
        self.width = width
        self.ctx = ctx
        self.bprev = border_prev_same
        self.bnext = border_next_same
        ppr = para.ppr
        self.ind_left = ppr.get("ind_left", 0.0) or 0.0
        self.ind_right = ppr.get("ind_right", 0.0) or 0.0
        self.ind_first = ppr.get("ind_first", 0.0) or 0.0
        self.jc = ppr.get("jc", "left")
        if self.jc == "start":
            self.jc = "left"
        elif self.jc == "end":
            self.jc = "right"
        elif self.jc in ("mediumKashida", "highKashida", "lowKashida", "thaiDistribute"):
            self.jc = "both"
        self.tabs = [t for t in ppr.get("tabs", []) if t[0] not in ("clear", "bar")]
        self.smart_justify = self.jc == "both" and getattr(ctx.reader, "compat_mode", 15) >= 15
        self.builder = AtomBuilder(ctx)

    # ---- tab stops -------------------------------------------------------
    def next_tab(self, x: float, first_line: bool):
        best = None
        for t in self.tabs:
            if t[1] > x + EPS:
                best = t
                break
        if first_line and self.ind_first < 0 and x < self.ind_left - EPS:
            if best is None or self.ind_left < best[1]:
                best = ("left", self.ind_left, "none")
        if best is not None:
            return best
        dt = self.ctx.default_tab or 36.0
        last = max((t[1] for t in self.tabs), default=-1e9)
        pos = (math.floor(x / dt + EPS) + 1) * dt
        while pos <= last + EPS:
            pos += dt
        return ("left", pos, "none")

    # ---- atoms -----------------------------------------------------------
    def build_atoms(self) -> list[Atom]:
        atoms: list[Atom] = []
        b = self.builder
        p = self.p
        if p.label is not None:
            text, lstyle, lvl = p.label
            lab: list[Atom] = []
            numbering = self.ctx.reader.numbering
            pic = numbering.pic_data.get(lvl.pic_id) if lvl.pic_id is not None else None
            if pic:
                # picture bullet: Word scales it to the text size, bottom on the baseline
                pw, ph, _ = numbering.pic_bullets[lvl.pic_id]
                h = lstyle.size * 0.83
                w = h * (pw / ph if ph else 1.0)
                lab.append(Atom("box", width=w, asc=h, desc=0.0, style=lstyle,
                                payload=Picture(pic, w, h, None)))
            elif text:
                b.text(text, lstyle, None, lab)
            for a in lab:
                a.brk = False
                a.kind = "text" if a.kind == "space" else a.kind
                a.payload = "label"
            if pic:
                lab[0].payload = Picture(pic, lab[0].width, lab[0].asc, None)
            elif lab:
                lab[-1].payload = ("label_end", lvl.jc)
                lab[0].payload = ("label_start", lvl.jc)
            atoms.extend(lab)
            if lvl.suff == "tab":
                atoms.append(Atom("tab", style=lstyle, payload="label_tab"))
            elif lvl.suff == "space":
                b.text(" ", lstyle, None, atoms)
        for it in p.items:
            t = type(it)
            if t is Text:
                if it.style.hidden:
                    continue
                txt = it.text.replace("\t", "\x09")
                if "\t" in txt:
                    parts = txt.split("\t")
                    for k, part in enumerate(parts):
                        if k:
                            atoms.append(Atom("tab", style=it.style, link=it.link))
                        if part:
                            b.text(part, it.style, it.link, atoms)
                else:
                    b.text(txt, it.style, it.link, atoms)
            elif t is Tab:
                atoms.append(Atom("tab", style=it.style, link=it.link))
            elif t is Break:
                atoms.append(Atom("break", style=it.style, payload=it.kind))
            elif t is Picture:
                l, tp, r, bt = it.extent
                w = it.width + l + r
                h = it.height + tp + bt
                atoms.append(Atom("box", width=w, asc=h, desc=0.0, style=it.style,
                                  link=it.link, payload=it, brk=True))
                if len(atoms) > 1 and atoms[-2].kind == "text":
                    atoms[-2].brk = True
            elif t is Field:
                text = self.field_text(it)
                b.text(text, it.style, None, atoms)
            elif t is Floating:
                if it.dropcap and not it.width:
                    self._size_dropcap(it)
                atoms.append(Atom("float", payload=it))
            elif t is Bookmark:
                atoms.append(Atom("anchor", payload=it.name))
            elif t is MathItem:
                from .math import MathLayout
                color = it.style.color or ((1, 1, 1) if self.ctx.dark_bg else (0, 0, 0))
                box = MathLayout(self.ctx, color).layout(it.element, it.style.size, 0, it.display)
                if it.display:
                    asc, desc = box.asc, box.desc
                else:
                    # inline equations size the line by their base text only
                    f = MathLayout(self.ctx).font
                    asc, desc = f.ascent * it.style.size, f.descent * it.style.size
                cuts = [] if it.display else sorted(b for b in set(box.breaks) if 0.0 < b[0] < box.w - 0.01)
                pieces = _split_math(box, cuts)
                for k, (piece, hang) in enumerate(pieces):
                    # an inline equation may wrap after relations / before binary operators;
                    # at its edges only a real space is a break opportunity (a space typed
                    # inside the equation, as in "and<m:t> </m:t>S", is not)
                    atom = Atom("math", width=piece.w, asc=asc, desc=desc, style=it.style,
                                payload=piece, brk=k < len(pieces) - 1)
                    atom.hang = hang
                    atoms.append(atom)
            elif t is FootnoteRef:
                if it.id == "__self__":
                    txt = self.ctx.note_label
                else:
                    txt = it.text
                start = len(atoms)
                if txt:
                    b.text(txt, it.style, None, atoms)
                if it.id != "__self__" and start < len(atoms):
                    atoms[start].payload = ("note", it)
                elif it.id != "__self__":
                    atoms.append(Atom("anchor", payload=("note", it)))
            elif t is Separator:
                atoms.append(Atom("sep", width=144.0, asc=0.0, desc=0.0, style=it.style))
        return atoms

    def field_text(self, f: Field) -> str:
        v = self.ctx.fields.get(f.kind)
        if v is None:
            return f.fallback or "1"
        if f.fmt and v.isdigit():
            fmt = {"roman": "lowerRoman", "ROMAN": "upperRoman", "alphabetic": "lowerLetter",
                   "ALPHABETIC": "upperLetter", "Arabic": "decimal", "CardText": "cardinalText",
                   "OrdText": "ordinalText", "Ordinal": "ordinal"}.get(f.fmt)
            if fmt:
                return format_number(int(v), fmt)
        if f.picture and v.isdigit():
            return _numeric_picture(int(v), f.picture)
        return v

    # ---- main --------------------------------------------------------------
    def _squeeze_limit(self, lb) -> float:
        """How far past the right edge Word's justification lets a line run
        (the overflow is then absorbed by narrowing the word spaces)."""
        if JUSTIFY_MODE == "em":
            size = max((a.size for a in lb.atoms if a.kind == "text" and a.face), default=10.0)
            return JUSTIFY_SHRINK * size
        return JUSTIFY_SHRINK * (lb.right - lb.start)

    def _size_dropcap(self, fl):
        """A drop cap frame is as wide as its letters and as tall as the lines
        of this paragraph it spans."""
        cap = ParagraphLayout(fl.shape.blocks[0], 10000.0, self.ctx)
        width = sum(a.width for a in cap.build_atoms() if a.kind in ("text", "space"))
        mark = self.builder.metrics_atom(self.p.mark)
        height = fl.dropcap * self._estimate_line(mark)
        fl.width = fl.shape.width = width
        fl.height = fl.shape.height = height
        if fl.h_align == "outside_margin":  # dropCap="margin": in the left margin
            fl.h_align = None
            fl.h_off = -width - fl.dist[3]

    def _estimate_line(self, mark: Atom) -> float:
        ppr = self.p.ppr
        natural = mark.asc + mark.gap + mark.desc
        line = ppr.get("line")
        rule = ppr.get("lineRule", "auto")
        if line is None:
            return natural
        if rule == "exact":
            return line / 20.0
        if rule == "atLeast":
            return max(natural, line / 20.0)
        return natural * line / 240.0

    def layout(self, bounds=None) -> ParaFlow:
        """Break the paragraph into lines.

        ``bounds(y_offset, height)`` optionally narrows each line (text wrapping
        around floating objects): it returns ``None`` for no constraint, or the
        (left, right) interval available in column coordinates, or a number of
        points to move down when there is no usable room at that height.
        """
        p = self.p
        ppr = p.ppr
        atoms = self.build_atoms()
        mark = self.builder.metrics_atom(p.mark)
        lines_atoms: list[tuple[_LineBuilder, str | None]] = []
        first = True
        right0 = self.width - self.ind_right
        est_h = self._estimate_line(mark)
        y_off = 0.0

        def new_line():
            nonlocal first, y_off
            start = self.ind_left + (self.ind_first if first else 0.0)
            right = right0
            skip = 0.0
            if bounds is not None:
                for _ in range(400):
                    b = bounds(y_off + skip, est_h)
                    if b is None:
                        break
                    if isinstance(b, (int, float)):
                        skip += max(0.25, b)  # no room beside the object: move below it
                        continue
                    if b[0] > 0.01:
                        # beside an object on the left the paragraph's indents
                        # count from the object's edge (Word keeps the first-line
                        # indent there)
                        start = max(start, b[0] + start)
                    right = min(right, b[1])
                    break
            lb = _LineBuilder(start, right, first)
            lb.skip = skip
            y_off += skip + est_h
            first = False
            return lb

        lb = new_line()
        right = lb.right
        i = 0
        n = len(atoms)
        while i < n:
            right = lb.right
            a = atoms[i]
            k = a.kind
            if k == "break":
                a.asc, a.desc = 0, 0
                lb.atoms.append(a)
                kind = a.payload
                lines_atoms.append((lb, kind))
                lb = new_line()
                i += 1
                continue
            if k == "tab":
                x = lb.x
                if lb.pending:
                    lb.pending = None
                stop = self.next_tab(x, lb.first)
                typ, pos, leader = stop
                if a.payload == "label_tab" and self.ind_first < 0 and x < self.ind_left - EPS:
                    # the label tab uses the hanging indent even if other tabs exist
                    pass
                if pos > right + EPS and typ == "left" and not lb.empty and pos > self.width + EPS:
                    lines_atoms.append((lb, None))
                    lb = new_line()
                    continue
                a.payload = (a.payload, leader)
                if typ in ("left", "start", "num"):
                    a.width = max(0.0, pos - x)
                    lb.add(a)
                else:
                    a.width = 0.0
                    lb.atoms.append(a)
                    t = {"end": "right", "right": "right", "center": "center", "decimal": "decimal"}.get(typ, "left")
                    if t == "left":
                        a.width = max(0.0, pos - x)
                        lb.x += a.width
                    else:
                        lb.pending = (a, pos, t, len(lb.atoms), x)
                        lb._update_pending()
                i += 1
                continue
            if k in ("float", "anchor", "mark"):
                lb.atoms.append(a)
                i += 1
                continue
            # gather a word: atoms up to and including a break opportunity
            j = i
            word_w = 0.0
            trail_w = 0.0
            while j < n:
                b = atoms[j]
                if b.kind in ("break", "tab"):
                    break
                if b.kind == "space":
                    trail_w += b.width
                elif b.kind in ("float", "anchor"):
                    pass
                else:
                    if trail_w:
                        break
                    word_w += b.width
                j += 1
                if b.brk and b.kind != "space":
                    # include following spaces
                    while j < n and atoms[j].kind == "space":
                        trail_w += atoms[j].width
                        j += 1
                    break
            if j == i:
                j = i + 1
            end_x = lb.end_x_if(word_w)
            last = next((atoms[q] for q in range(j - 1, i - 1, -1) if atoms[q].kind != "space"), None)
            hang = last.hang if last is not None and not trail_w else 0.0
            fits = end_x - hang <= right + EPS
            if not fits and self.smart_justify and j < n and not lb.empty:
                # Word 2013+ justification may squeeze word spaces to fit one more word
                has_space = any(x.kind == "space" for x in lb.atoms)
                fits = has_space and end_x - right <= self._squeeze_limit(lb)
            if fits or lb.empty:
                if end_x > right + EPS and lb.empty and word_w > 0:
                    # word wider than the line: break it by characters
                    rest = self._split_long(atoms, i, j, right - lb.x, lb)
                    if rest is not None:
                        atoms[i:j] = rest[0]
                        n = len(atoms)
                        j = i + rest[1]
                        for q in range(i, j):
                            lb.add(atoms[q])
                        lines_atoms.append((lb, None))
                        lb = new_line()
                        i = j
                        continue
                for q in range(i, j):
                    lb.add(atoms[q])
                i = j
            else:
                lines_atoms.append((lb, None))
                lb = new_line()
        if lines_atoms and lines_atoms[-1][1] in ("page", "column") and lb.empty                 and not any(x.kind in ("float", "anchor") for x in lb.atoms):
            # a break at the very end of a paragraph: the paragraph mark stays
            # with the break, nothing is carried to the next page
            prev_lb, kind = lines_atoms.pop()
            lb = prev_lb
            lb.atoms = [a for a in lb.atoms if a.kind != "break"]
            lines_atoms.append((lb, "end"))
            lines_atoms[-1] = (lb, "end:" + kind)
        else:
            lines_atoms.append((lb, "end"))

        lines = []
        nlines = len(lines_atoms)
        for idx, (lbx, brk) in enumerate(lines_atoms):
            is_last = idx == nlines - 1
            if brk and brk.startswith("end:"):
                ln = self._finish(lbx, "end", is_last, mark, idx == 0)
                ln.break_after = brk[4:]
            else:
                ln = self._finish(lbx, brk, is_last, mark, idx == 0)
            if lbx.skip:
                ln.ops = move_ops(ln.ops, 0, lbx.skip)
                ln.height += lbx.skip
                ln.baseline += lbx.skip
            lines.append(ln)
        # paragraph borders / shading
        self._decorate(lines)

        sb = ppr.get("sp_before", 0.0) or 0.0
        sa = ppr.get("sp_after", 0.0) or 0.0
        if ppr.get("sp_before_lines"):
            sb = ppr["sp_before_lines"] * 12.0
        if ppr.get("sp_after_lines"):
            sa = ppr["sp_after_lines"] * 12.0
        if ppr.get("sp_before_auto"):
            sb = AUTO_SPACING
        if ppr.get("sp_after_auto"):
            sa = AUTO_SPACING
        is_empty = not any(type(it) in (Text, Picture, Tab, Field) for it in p.items) and p.label is None
        if is_empty and p.sect is not None and not any(type(it) is Floating for it in p.items):
            # an empty paragraph that only carries a section break is not laid out
            lines = []
        return ParaFlow(lines, sb, sa,
                        keep_next=ppr.get("keepNext", False),
                        keep_lines=ppr.get("keepLines", False),
                        widow=ppr.get("widowControl", True),
                        page_break_before=ppr.get("pageBreakBefore", False),
                        style_id=p.style_id,
                        contextual=ppr.get("contextualSpacing", False),
                        is_empty=is_empty)

    def _split_long(self, atoms, i, j, avail, lb):
        """Split atoms[i:j] (one long word) so that its first part fits ``avail``."""
        out = []
        used = 0.0
        for q in range(i, j):
            a = atoms[q]
            if a.kind != "text" or not a.face:
                if used + a.width <= avail or not out:
                    out.append(a)
                    used += a.width
                    continue
                break
            f = a.face.font
            k = 0
            w = 0.0
            for ch in a.text:
                cw = f.char_width(ch, a.size) * a.hscale / 100 + a.char_space
                if used + w + cw > avail and (out or k):
                    break
                w += cw
                k += 1
            if k == len(a.text):
                out.append(a)
                used += a.width
                continue
            if k == 0 and out:
                break
            k = max(1, k)
            left = _clone(a, a.text[:k])
            right = _clone(a, a.text[k:])
            new = atoms[i:q] + [left, right] + atoms[q + 1:j]
            return new, (q - i) + 1
        return None

    def _finish(self, lb: _LineBuilder, brk, is_last: bool, mark: Atom, is_first: bool) -> Line:
        ppr = self.p.ppr
        atoms = lb.atoms
        # metrics
        max_top = 0.0
        max_desc = 0.0
        box_top = 0.0
        label_top = 0.0
        has_content = False
        for a in atoms:
            if a.kind == "box":
                box_top = max(box_top, a.asc)
                has_content = True
            elif a.kind == "text" and a.face and (a.payload == "label" or (
                    isinstance(a.payload, tuple) and a.payload and str(a.payload[0]).startswith("label_"))):
                # a list label (bullet/number) never deepens the line, and only
                # its ascent above the text adds height, unscaled by line spacing
                label_top = max(label_top, a.asc + a.gap)
            elif a.kind == "math" or (a.kind in ("text", "space", "mark") and a.face):
                top = a.asc + a.gap
                if top > max_top:
                    max_top = top
                if a.desc > max_desc:
                    max_desc = a.desc
                has_content = True
        text_on_line = max_top > 0 or max_desc > 0
        if any(a.kind == "math" for a in atoms):
            # lines holding equations are at least as tall as the paragraph's text
            max_top = max(max_top, mark.asc + mark.gap)
            max_desc = max(max_desc, mark.desc)
        if not text_on_line:
            # no text on the line: the paragraph mark's font sets the metrics
            # (Word ignores the mark on lines that carry text)
            max_top = mark.asc + mark.gap
            max_desc = mark.desc
        if not has_content and not is_last:
            # line consisting only of a break: use the break's run metrics
            for a in atoms:
                if a.kind in ("break", "tab") and a.style is not None:
                    m = self.builder.metrics_atom(a.style)
                    max_top = max(max_top, m.asc + m.gap)
                    max_desc = max(max_desc, m.desc)
        text_natural = max_top + max_desc
        label_extra = max(0.0, label_top - max_top) if label_top else 0.0
        if box_top and not text_on_line:
            max_desc = 0.0  # a picture-only line has no descent below the picture
        top = max(max_top, box_top)
        natural = top + max_desc
        rule = ppr.get("lineRule", "auto")
        line = ppr.get("line")
        if line is None:
            h = natural
        elif rule == "exact":
            h = line / 20.0
        elif rule == "atLeast":
            h = max(natural, line / 20.0)
        else:
            # "multiple" spacing adds (factor - 1) x the text height below the
            # line; inline pictures themselves are never scaled
            h = natural + (line / 240.0 - 1.0) * text_natural
        if label_extra and rule != "exact":
            h += label_extra
            top += label_extra
        # Word puts the extra leading of "multiple" spacing below the text; with
        # exact spacing the baseline sits at 80% of the line box for any font.
        baseline = top
        if rule == "exact" and line is not None:
            baseline = h * 0.8
        # alignment
        content_end = lb.start
        last_vis = -1
        for idx, a in enumerate(atoms):
            if a.kind not in ("space", "mark", "float", "anchor", "break"):
                last_vis = idx
        x = lb.start
        for a in atoms:
            a.x = x
            x += a.width
        if last_vis >= 0:
            content_end = atoms[last_vis].x + atoms[last_vis].width - atoms[last_vis].hang
        avail_right = lb.right
        jc = self.jc
        shift = 0.0
        extra_per_space = 0.0
        if jc == "center":
            shift = (avail_right - content_end) / 2
        elif jc == "right":
            shift = avail_right - content_end
        elif jc in ("both", "distribute") and (not is_last or jc == "distribute") and brk not in ("page", "column"):
            last_tab = -1
            for idx, a in enumerate(atoms):
                if a.kind == "tab":
                    last_tab = idx
            nsp = sum(a.spaces for idx, a in enumerate(atoms)
                      if a.kind == "space" and last_tab < idx < last_vis)
            slack = avail_right - content_end
            if nsp and (slack > 0 or (slack < 0 and self.smart_justify)):
                extra_per_space = slack / nsp
            elif jc == "distribute" and slack > 0 and last_vis > 0:
                chars = sum(len(a.text) for a in atoms[last_tab + 1:last_vis + 1] if a.kind == "text")
                if chars > 1:
                    extra = slack / (chars - 1)
                    for a in atoms[last_tab + 1:last_vis + 1]:
                        if a.kind == "text":
                            a.char_space += extra
                            a.width += extra * len(a.text)
                    x = lb.start
                    for a in atoms:
                        a.x = x
                        x += a.width
        if extra_per_space:
            x = lb.start
            last_tab = max((idx for idx, a in enumerate(atoms) if a.kind == "tab"), default=-1)
            for idx, a in enumerate(atoms):
                a.x = x
                w = a.width
                if a.kind == "space" and last_tab < idx < last_vis:
                    w += extra_per_space * a.spaces
                    a.width = w
                x += w
        if shift and jc in ("center", "right"):
            for a in atoms:
                a.x += shift
        # label alignment (lvlJc)
        self._align_label(atoms)
        ops, floats, notes = self._emit(atoms, baseline, h, is_first)
        ln = Line(h, ops, floats, notes, brk if brk in ("page", "column") else None, baseline)
        if rule == "auto" and line is not None and line > 240 and not has_content:
            # an empty paragraph's extra leading may hang below the page bottom
            ln.trailing = (line / 240.0 - 1.0) * text_natural
        return ln

    def _align_label(self, atoms):
        start = end = None
        for idx, a in enumerate(atoms):
            if isinstance(a.payload, tuple) and a.payload and a.payload[0] == "label_start":
                start = idx
            if isinstance(a.payload, tuple) and a.payload and a.payload[0] == "label_end":
                end = idx
        if start is not None and end is None:
            end = start  # a one-atom label carries only the start marker
        if start is None or end is None:
            return
        jc = atoms[start].payload[1]
        if jc not in ("right", "center", "end"):
            return
        width = atoms[end].x + atoms[end].width - atoms[start].x
        delta = -width if jc in ("right", "end") else -width / 2
        for a in atoms[start:end + 1]:
            a.x += delta

    def _emit(self, atoms, baseline, h, is_first):
        ops: list = []
        floats = []
        notes = []
        dark = self.ctx.dark_bg
        run: list[Atom] = []

        def flush():
            if not run:
                return
            a0 = run[0]
            text = "".join(r.text for r in run)
            col = a0.style.color if a0.style and a0.style.color is not None else ((1, 1, 1) if dark else (0, 0, 0))
            kerns = []
            off = 0
            for r in run:
                if r.kerns:
                    kerns.extend((off + i, dx) for i, dx in r.kerns)
                off += len(r.text)
            y = baseline - a0.shift
            if not kerns:
                ops.append(TextOp(a0.x, y, text, a0.face, a0.size, col, a0.char_space, a0.hscale))
            else:
                # split at kerned pairs so every glyph lands where Word puts it
                x = a0.x
                start = 0
                f = a0.face.font
                for i, dx in kerns:
                    seg = text[start:i]
                    if seg:
                        ops.append(TextOp(x, y, seg, a0.face, a0.size, col, a0.char_space, a0.hscale))
                        x += (f.width(seg, a0.size) * a0.hscale / 100 + a0.char_space * len(seg))
                    x += dx
                    start = i
                seg = text[start:]
                if seg:
                    ops.append(TextOp(x, y, seg, a0.face, a0.size, col, a0.char_space, a0.hscale))
            run.clear()

        # shading/highlight and character borders span consecutive atoms of one
        # run as a single box (per-word boxes leave hairline seams)
        boxes = []  # [kind, value, x0, x1, top, bottom]
        for a in atoms:
            st = a.style
            if a.kind not in ("text", "space") or st is None:
                boxes.append(None)
                continue
            for kind, val in (("fill", st.highlight or st.shading), ("border", st.border)):
                if not val:
                    continue
                if kind == "fill":
                    top = baseline - a.asc - a.gap if st.shading and not st.highlight else baseline - a.asc
                else:
                    top = baseline - a.asc
                bottom = baseline + a.desc
                last = next((b for b in reversed(boxes) if b is not None and b[0] == kind), None)
                if last is not None and last[1] == val and abs(last[3] - a.x) < 0.01:
                    last[3] = a.x + a.width
                    last[4] = min(last[4], top)
                    last[5] = max(last[5], bottom)
                else:
                    boxes.append([kind, val, a.x, a.x + a.width, top, bottom])
        for b in boxes:
            if b is None:
                continue
            kind, val, x0, x1, top, bottom = b
            top, bottom = max(0.0, top), min(h, bottom)
            if kind == "fill":
                ops.insert(0, RectOp(x0, top, x1 - x0, bottom - top, fill=val))
            else:
                col = val.color or ((1, 1, 1) if dark else (0, 0, 0))
                pad = val.space
                ops.append(RectOp(x0 - pad, top, x1 - x0 + 2 * pad, bottom - top, stroke=col,
                                  width=max(val.width, 0.25)))

        for a in atoms:
            k = a.kind
            st = a.style
            if k == "text" and a.text.strip("\u200b\u2060\ufeff") == "":
                flush()  # invisible spacer (math spacing): nothing to draw
                continue
            if k == "text":
                if (run and (run[-1].face is not a.face or run[-1].size != a.size or run[-1].style is not a.style
                             or run[-1].char_space != a.char_space or abs(run[-1].x + run[-1].width - a.x) > 0.01)):
                    flush()
                run.append(a)
            elif k == "space":
                if run and run[-1].style is a.style and abs(run[-1].x + run[-1].width - a.x) < 0.01 \
                        and abs(a.width - a.face.font.width(a.text, a.size) * a.hscale / 100
                                - a.char_space * len(a.text) - sum(d for _, d in (a.kerns or ()))) < 1e-6:
                    run.append(a)
                else:
                    flush()
                    if a.face is not None and a.text:
                        # a stretched (justified) or stand-alone space: still write the
                        # character so extracted/copied text keeps its word breaks
                        natural = a.face.font.width(a.text, a.size) * a.hscale / 100
                        ops.append(TextOp(a.x, baseline - a.shift, a.text, a.face, a.size, (0, 0, 0),
                                          (a.width - natural) / len(a.text), a.hscale))
            else:
                flush()
            if k == "math":
                ops.extend(move_ops(a.payload.ops, a.x, baseline))
            if k == "box":
                pic: Picture = a.payload
                l, tp, r, b = pic.extent
                x = a.x + l
                y = baseline - a.asc + tp
                if pic.shape is not None:
                    ops.extend(render_shape(pic.shape, x, y, self.ctx))
                elif pic.data:
                    ops.append(ImageOp(x, y, pic.width, pic.height, pic.data, pic.crop))
                if a.link:
                    ops.append(_link(a.link, a.x, baseline - a.asc, a.width, a.asc))
            elif k == "tab":
                leader = a.payload[1] if isinstance(a.payload, tuple) else "none"
                if leader and leader != "none" and a.width > 0 and a.style is not None:
                    ops.extend(self._leader(a, leader, baseline))
            elif k == "float":
                floats.append((a.payload, a.x))
            elif k == "anchor":
                if isinstance(a.payload, tuple):
                    notes.append(a.payload[1])
                else:
                    ops.append(AnchorOp(a.x, baseline - a.asc, a.payload))
            elif k == "sep":
                ops.append(LineOp(a.x, baseline - 3.0, a.x + a.width, baseline - 3.0, 0.5, (0, 0, 0)))
            if k == "text" and isinstance(a.payload, tuple) and a.payload[0] == "note":
                notes.append(a.payload[1])
        flush()
        # decorations: underline, strike, links
        self._decorations(atoms, baseline, ops)
        return ops, floats, notes

    def _leader(self, a: Atom, leader: str, baseline):
        ch = {"dot": ".", "hyphen": "-", "underscore": "_", "middleDot": "·", "heavy": "_"}.get(leader, ".")
        fam, size, b, i = self.builder.style_font(a.style, "ascii")
        face = self.ctx.face(fam, b, i)
        cw = face.font.char_width(ch, size)
        if cw <= 0:
            return []
        if ch == "_":
            y = baseline - face.font.underline_pos * size
            return [LineOp(a.x, y, a.x + a.width, y, max(0.5, face.font.underline_size * size), a.style.color or (0, 0, 0))]
        # Word snaps leader dots to a grid so consecutive leaders line up
        start = math.ceil((a.x + 1.0) / cw) * cw
        end = a.x + a.width - cw * 0.5
        count = int((end - start) / cw)
        if count <= 0:
            return []
        return [TextOp(start, baseline, ch * count, face, size, a.style.color or (0, 0, 0))]

    def _decorations(self, atoms, baseline, ops):
        n = len(atoms)
        last_vis = -1
        for idx, a in enumerate(atoms):
            if a.kind in ("text", "box", "tab"):
                last_vis = idx
        i = 0
        while i < n:
            a = atoms[i]
            st = a.style
            if a.kind not in ("text", "space", "tab") or st is None or i > last_vis:
                i += 1
                continue
            if st.underline or st.strike or st.dstrike:
                j = i
                while j <= last_vis and atoms[j].kind in ("text", "space", "tab") and atoms[j].style is not None \
                        and atoms[j].style.underline == st.underline and atoms[j].style.strike == st.strike \
                        and atoms[j].style.dstrike == st.dstrike and atoms[j].style.color == st.color:
                    j += 1
                seg = atoms[i:j]
                if not seg:
                    i += 1
                    continue
                x0 = seg[0].x
                x1 = seg[-1].x + seg[-1].width
                ref = max(seg, key=lambda z: z.size if z.face else 0)
                face = ref.face or self.ctx.face(st.font_ascii, st.bold, st.italic)
                size = ref.size if ref.face else st.size
                f = face.font
                color = st.color or (0, 0, 0)
                if st.underline:
                    ucol = st.u_color or color
                    thick = max(0.5, f.underline_size * size)
                    y = baseline - ref.shift - f.underline_pos * size + thick / 2
                    u = st.underline
                    dash = None
                    if u.startswith("dotted") or u == "dottedHeavy":
                        dash = (thick, thick * 2)
                    elif u.startswith("dash"):
                        dash = (thick * 4, thick * 2)
                    if u in ("thick", "dottedHeavy", "dashedHeavy", "wavyHeavy", "dashLongHeavy"):
                        thick *= 2
                    if u == "words":
                        for s in seg:
                            if s.kind == "text":
                                ops.append(LineOp(s.x, y, s.x + s.width, y, thick, ucol))
                    elif u in ("double", "wavyDouble"):
                        ops.append(LineOp(x0, y - thick, x1, y - thick, thick * 0.8, ucol, dash))
                        ops.append(LineOp(x0, y + thick, x1, y + thick, thick * 0.8, ucol, dash))
                    else:
                        ops.append(LineOp(x0, y, x1, y, thick, ucol, dash))
                if st.strike or st.dstrike:
                    y = baseline - ref.shift - f.strike_pos * size
                    t = max(0.5, f.strike_size * size)
                    if st.dstrike:
                        ops.append(LineOp(x0, y - t, x1, y - t, t * 0.8, color))
                        ops.append(LineOp(x0, y + t, x1, y + t, t * 0.8, color))
                    else:
                        ops.append(LineOp(x0, y, x1, y, t, color))
                i = j
            else:
                i += 1
        # links
        i = 0
        while i < n:
            a = atoms[i]
            if a.link and a.kind in ("text", "space", "box"):
                j = i
                while j < n and atoms[j].link == a.link and atoms[j].kind in ("text", "space", "box"):
                    j += 1
                x0 = atoms[i].x
                x1 = atoms[j - 1].x + atoms[j - 1].width
                top = baseline - max(z.asc for z in atoms[i:j])
                bot = baseline + max(z.desc for z in atoms[i:j])
                ops.append(_link(a.link, x0, top, x1 - x0, bot - top))
                i = j
            else:
                i += 1

    def _decorate(self, lines: list[Line]):
        ppr = self.p.ppr
        bdr = ppr.get("pBdr") or {}
        shd = ppr.get("shd")
        fill = shd if isinstance(shd, tuple) else None
        vis = {k: v for k, v in bdr.items() if v is not None and v.visible}
        if not vis and fill is None:
            return
        top = vis.get("top") if not self.bprev else None
        bottom = vis.get("bottom") if not self.bnext else None
        if self.bnext and vis.get("between"):
            bottom = vis.get("between")
        left = vis.get("left")
        right = vis.get("right")
        pad_top = (top.space + top.width) if top else 0.0
        pad_bot = (bottom.space + bottom.width) if bottom else 0.0
        if pad_top:
            ln = lines[0]
            ln.ops = move_ops(ln.ops, 0, pad_top)
            ln.baseline += pad_top
            ln.height += pad_top
        if pad_bot:
            lines[-1].height += pad_bot
        x0 = self.ind_left - (left.space + left.width / 2 if left else 0.0)
        x1 = self.width - self.ind_right + (right.space + right.width / 2 if right else 0.0)
        if left and self.ind_first < 0:
            x0 = self.ind_left + self.ind_first - left.space - left.width / 2
        for idx, ln in enumerate(lines):
            deco = []
            y0 = 0.0
            y1 = ln.height
            if fill is not None:
                fx0 = x0 if left else self.ind_left + min(0.0, self.ind_first)
                fx1 = x1 if right else self.width - self.ind_right
                deco.append(RectOp(fx0, y0, fx1 - fx0, y1 - y0, fill=fill))
            if idx == 0 and top:
                deco_b = []
                _border_line(deco_b, top, x0, top.width / 2, x1, top.width / 2, True)
                deco.extend(deco_b)
            if idx == len(lines) - 1 and bottom:
                yb = ln.height - bottom.width / 2
                _border_line(deco, bottom, x0, yb, x1, yb, True)
            if left:
                _border_line(deco, left, x0, y0, x0, y1, False)
            if right:
                _border_line(deco, right, x1, y0, x1, y1, False)
            ln.ops = deco + ln.ops


def _clone(a: Atom, text: str) -> Atom:
    b = Atom(a.kind, text, a.face, a.size, a.style, 0.0, a.asc, a.desc, a.gap, a.shift, a.link,
             a.brk, a.payload, a.char_space, a.hscale)
    b.width = a.face.font.width(text, a.size) * a.hscale / 100 + a.char_space * len(text)
    return b


def _link(link: str, x, y, w, h):
    if link.startswith("#"):
        return LinkOp(x, y, w, h, anchor=link[1:])
    return LinkOp(x, y, w, h, url=link)


# --------------------------------------------------------------------------- shapes
def shape_height(shape: Shape, ctx: LayoutContext) -> float:
    """Height of a shape, growing auto-fit text boxes to their content."""
    if not (shape.autofit and shape.blocks):
        return shape.height
    l, t, r, b = shape.insets
    flows = layout_blocks(shape.blocks, max(1.0, shape.width - l - r), ctx)
    _, h = stack_flows(flows, ctx.sum_spacing)
    return h + t + b


def render_shape(shape: Shape, x: float, y: float, ctx: LayoutContext) -> list:
    ops: list = []
    if shape.geom == "group":
        for dx, dy, child in shape.children:
            if isinstance(child, Shape):
                ops.extend(render_shape(child, x + dx, y + dy, ctx))
            elif isinstance(child, Picture) and child.data:
                ops.append(ImageOp(x + dx, y + dy, child.width, child.height, child.data, child.crop))
        return ops
    if shape.autofit and shape.blocks:
        shape.height = shape_height(shape, ctx)
    geom = shape.geom
    if shape.fill is not None or shape.line is not None:
        if geom in ("ellipse",):
            ops.append(RectOp(x, y, shape.width, shape.height, fill=shape.fill, stroke=shape.line,
                              width=shape.line_width, ellipse=True))
        elif geom in ("line", "straightConnector1"):
            if shape.line is not None:
                x1, x2 = (x + shape.width, x) if shape.flip_x else (x, x + shape.width)
                y1, y2 = (y + shape.height, y) if shape.flip_y else (y, y + shape.height)
                ops.append(LineOp(x1, y1, x2, y2, shape.line_width, shape.line, shape.dash))
                if shape.end_arrow:
                    ops.append(_arrow(x1, y1, x2, y2, shape.line_width, shape.line, shape.end_arrow))
                if shape.start_arrow:
                    ops.append(_arrow(x2, y2, x1, y1, shape.line_width, shape.line, shape.start_arrow))
        elif geom == "path":
            pts = [(x + (shape.width - px if shape.flip_x else px), y + (shape.height - py if shape.flip_y else py))
                   for px, py in shape.points]
            ops.append(PolyOp(pts, shape.fill if shape.closed else None, shape.line, shape.line_width,
                              shape.closed, shape.dash))
        else:
            radius = min(shape.width, shape.height) * 0.1667 if geom == "roundRect" else 0.0
            ops.append(RectOp(x, y, shape.width, shape.height, fill=shape.fill, stroke=shape.line,
                              width=shape.line_width, radius=radius))
    if shape.data and geom not in ("line", "straightConnector1", "group"):
        ops.append(ImageOp(x, y, shape.width, shape.height, shape.data, None))
    if shape.blocks and shape.vert:
        # rotated text frame: lines run along the shape's height
        l, t, r, b = shape.insets
        frame_w = max(1.0, shape.height - l - r)
        frame_h = shape.width - t - b
        flows = layout_blocks(shape.blocks, frame_w, ctx)
        lines, height = stack_flows(flows, ctx.sum_spacing)
        off = 0.0
        if shape.v_anchor == "ctr":
            off = (frame_h - height) / 2
        elif shape.v_anchor == "b":
            off = frame_h - height
        inner = []
        for ly, ln in lines:
            inner.extend(move_ops(ln.ops, 0.0, off + ly))
        if shape.vert == "vert270":  # reads bottom to top
            m = (0.0, -1.0, 1.0, 0.0, x + t, y + shape.height - l)
        else:  # reads top to bottom
            m = (0.0, 1.0, -1.0, 0.0, x + shape.width - t, y + l)
        ops.append(TransformOp(m, inner))
    elif shape.blocks:
        l, t, r, b = shape.insets
        inner_w = max(1.0, shape.width - l - r)
        saved = ctx.dark_bg
        if shape.fill is not None:
            ctx.dark_bg = _is_dark(shape.fill)
        flows = layout_blocks(shape.blocks, inner_w, ctx)
        ctx.dark_bg = saved
        lines, height = stack_flows(flows, ctx.sum_spacing)
        off = 0.0
        avail = shape.height - t - b
        if shape.v_anchor == "ctr":
            off = (avail - height) / 2
        elif shape.v_anchor == "b":
            off = avail - height
        inner = []
        for ly, ln in lines:
            inner.extend(move_ops(ln.ops, x + l, y + t + off + ly))
        ops.extend(inner)
    return ops


def _longest_word(blocks, ctx) -> float:
    """Width of the widest unbreakable piece of text in some blocks."""
    best = 0.0
    for blk in blocks:
        if not isinstance(blk, Paragraph):
            continue
        pl = ParagraphLayout(blk, 100000.0, ctx)
        word = 0.0
        for a in pl.build_atoms():
            if a.kind in ("space", "tab", "break", "math"):
                word = 0.0  # equations can break at operators: not measured
                continue
            word += a.width
            best = max(best, word)  # Word measures the word itself, not the indents
            if a.brk:
                word = 0.0
    return best


def _autofit_grid(tbl, grid, cell_mar, avail, ctx):
    """Word's auto-fit columns: a column whose longest unbreakable word does not
    fit its preferred width grows to it; the other columns give up the excess in
    proportion to their room above their own minimum (large-document.docx 7.2)."""
    n = len(grid)
    if not n:
        return grid
    mins = [0.0] * n
    for row in tbl.rows:
        for cell in row.cells:
            if cell.span != 1 or not (0 <= cell.col < n):
                continue
            mar = dict(cell_mar)
            mar.update(cell.tcpr.get("tcMar") or {})
            need = _longest_word(cell.blocks, ctx) + mar.get("left", 0.0) + mar.get("right", 0.0)
            mins[cell.col] = max(mins[cell.col], need)
    if all(m <= g + 0.01 for m, g in zip(mins, grid)):
        return grid
    widths = [max(g, m) for g, m in zip(grid, mins)]
    excess = sum(widths) - avail
    if excess > 0.01:
        room = [w - m for w, m in zip(widths, mins)]
        total_room = sum(room)
        if total_room > 0:
            take = min(excess, total_room)
            widths = [w - take * r / total_room for w, r in zip(widths, room)]
    return widths


def _op_x(op) -> float:
    if isinstance(op, LineOp):
        return min(op.x1, op.x2)
    if isinstance(op, PolyOp):
        return min((p[0] for p in op.points), default=0.0)
    if isinstance(op, TransformOp):
        inner = min((_op_x(o) for o in op.ops), default=0.0)
        return op.matrix[0] * inner + op.matrix[4]
    return getattr(op, "x", 0.0)


def _split_math(box, cuts):
    """Split a laid-out equation at the (x, hang) positions in ``cuts``;
    returns [(piece, hang)]."""
    if not cuts:
        return [(box, 0.0)]
    from .math import Box
    hangs = [h for _, h in cuts] + [0.0]
    cuts = [x for x, _ in cuts]
    edges = [0.0] + list(cuts) + [box.w]
    pieces = [Box(w=edges[i + 1] - edges[i], asc=box.asc, desc=box.desc) for i in range(len(edges) - 1)]
    for op in box.ops:
        x = _op_x(op) + 0.01
        k = max(0, min(len(pieces) - 1, sum(1 for c in cuts if x >= c)))
        pieces[k].ops.append(op.moved(-edges[k], 0.0))
    return list(zip(pieces, hangs))


def _numeric_picture(n: int, pic: str) -> str:
    """Word's numeric picture switch for integers: '0' pads with zeros, '#'
    with spaces; other characters are literal."""
    slots = sum(1 for ch in pic if ch in "0#")
    if not slots:
        return str(n)
    digits = str(abs(n))
    lead, digits = digits[:-slots] if len(digits) > slots else "", digits[-slots:].rjust(slots, "_")
    out, k = [], 0
    for ch in pic:
        if ch in "0#":
            d = digits[k]
            k += 1
            out.append(d if d != "_" else ("0" if ch == "0" else " "))  # "_": no digit here
            if k == 1 and lead:
                out[-1] = lead + out[-1]
        else:
            out.append(ch)
    # blank '#' slots and the grouping commas before the first digit disappear
    return ("-" if n < 0 else "") + "".join(out).strip().lstrip(", ")


def _arrow(x1, y1, x2, y2, width, color, kind="block"):
    """A filled line end (arrowhead or dot) at (x2, y2) pointing along the line."""
    import math as _m
    ang = _m.atan2(y2 - y1, x2 - x1)
    size = max(3.0, width * 3.5)
    if kind == "oval":
        r = size * 0.45
        return PolyOp([(x2 + r * _m.cos(i * _m.pi / 12), y2 + r * _m.sin(i * _m.pi / 12)) for i in range(24)],
                      fill=color, closed=True)
    half = size * 0.45
    bx, by = x2 - size * _m.cos(ang), y2 - size * _m.sin(ang)
    px, py = -_m.sin(ang) * half, _m.cos(ang) * half
    return PolyOp([(x2, y2), (bx + px, by + py), (bx - px, by - py)], fill=color, closed=True)


def _is_dark(c) -> bool:
    if not isinstance(c, tuple):
        return False
    r, g, b = c
    return (0.299 * r + 0.587 * g + 0.114 * b) < 0.45


# --------------------------------------------------------------------------- tables
DEFAULT_CELL_MAR = {"left": 5.4, "right": 5.4, "top": 0.0, "bottom": 0.0}


_STYLE_WEIGHT = {"single": 1, "thick": 2, "double": 3, "dotted": 0, "dashed": 0}


def _heavier(a: Border | None, b: Border | None) -> Border | None:
    """Word's border conflict resolution: the wider border wins."""
    if a is None or not a.visible:
        return b if b is not None and b.visible else a
    if b is None or not b.visible:
        return a
    if b.width > a.width + 1e-6:
        return b
    if a.width > b.width + 1e-6:
        return a
    if _STYLE_WEIGHT.get(b.style, 1) > _STYLE_WEIGHT.get(a.style, 1):
        return b
    return a


def tb_left(tblpr) -> Border | None:
    return (tblpr.get("tblBorders") or {}).get("left")


def _bw(b: Border | None) -> float:
    return b.width if b is not None and b.visible else 0.0


def layout_table(tbl: Table, width: float, ctx: LayoutContext) -> TableFlow:
    tblpr = tbl.tblpr
    grid = list(tbl.grid)
    total = sum(grid)
    tw = tblpr.get("tblW")
    if tw and tw[1] == "pct" and tw[0] > 0 and total > 0:
        target = width * tw[0] / 100.0
        grid = [g * target / total for g in grid]
        total = target
    elif tw and tw[1] == "dxa" and tw[0] > 0 and total > 0 and abs(tw[0] - total) > 1 \
            and tblpr.get("layout") == "fixed":
        grid = [g * tw[0] / total for g in grid]
        total = tw[0]
    cell_mar = dict(DEFAULT_CELL_MAR)
    cell_mar.update(tblpr.get("cellMar") or {})
    if tblpr.get("layout") != "fixed" and not (tw and tw[1] == "pct"):
        # auto-fit: in legacy mode the table sticks out by the cell margins
        avail = width + ((cell_mar.get("left", 5.4) + cell_mar.get("right", 5.4)) if tbl.compat_legacy else 0.0)
        grid = _autofit_grid(tbl, grid, cell_mar, avail, ctx)
        total = sum(grid)
    ind = tblpr.get("tblInd", 0.0) or 0.0
    jc = tblpr.get("jc")
    if jc == "center":
        x0 = (width - total) / 2
    elif jc in ("right", "end"):
        x0 = width - total
    else:
        x0 = ind
        if tbl.compat_legacy:
            x0 -= cell_mar.get("left", 5.4)
        else:
            lb = (tbl.rows[0].cells[0].tcpr.get("tcBorders") or {}).get("left") \
                if tbl.rows and tbl.rows[0].cells else None
            lb = lb if lb is not None else tb_left(tblpr)
            if lb is not None and lb.visible:
                x0 += lb.width / 2
    col_x = [x0]
    for g in grid:
        col_x.append(col_x[-1] + g)
    ncols = len(grid)
    tb = tblpr.get("tblBorders") or {}
    tfill = tblpr.get("shd") if isinstance(tblpr.get("shd"), tuple) else None
    nrows = len(tbl.rows)

    # pass 1: own borders and content layout
    rows_cells: list[list[CellBox]] = []
    merge_fill: dict[int, object] = {}  # grid column -> shading of the vertical merge started there
    for ri, row in enumerate(tbl.rows):
        cells = []
        ex = row.trpr.get("tblPrEx") or {}
        row_mar = dict(cell_mar)
        row_mar.update(ex.get("cellMar") or {})
        row_tb = dict(tb)
        row_tb.update(ex.get("tblBorders") or {})
        for cell in row.cells:
            c0 = min(cell.col, ncols - 1) if ncols else 0
            c1 = min(cell.col + cell.span, ncols)
            x = col_x[c0]
            w = col_x[c1] - x if c1 > c0 else 0.0
            tcpr = cell.tcpr
            mar = dict(row_mar)
            mar.update(tcpr.get("tcMar") or {})
            cw = max(1.0, w - mar.get("left", 0) - mar.get("right", 0))
            shd = tcpr.get("shd")
            fill = shd if isinstance(shd, tuple) else (None if shd == "none" else tfill)
            is_cont = cell.vmerge == "continue"
            if is_cont:
                # a merged cell is one cell: it keeps the shading of its first row
                fill = merge_fill.get(c0, fill)
            else:
                merge_fill[c0] = fill
            lines, ch = [], 0.0
            if not is_cont:
                saved = ctx.dark_bg
                if fill is not None:
                    ctx.dark_bg = _is_dark(fill)
                blocks = cell.blocks
                if (len(blocks) >= 2 and isinstance(blocks[-2], Table) and isinstance(blocks[-1], Paragraph)
                        and not any(type(it) in (Text, Picture, Tab, Field) for it in blocks[-1].items)):
                    # the empty end-of-cell paragraph after a nested table takes no room
                    blocks = blocks[:-1]
                flows = layout_blocks(blocks, cw, ctx)
                ctx.dark_bg = saved
                lines, ch = stack_flows(flows, ctx.sum_spacing)
            cb = tcpr.get("tcBorders") or {}
            last_row_of_cell = ri + (cell.row_span if not is_cont else 1) - 1
            borders = {
                "top": cb.get("top", row_tb.get("top") if ri == 0 else row_tb.get("insideH")),
                "bottom": cb.get("bottom", row_tb.get("bottom") if last_row_of_cell >= nrows - 1 else row_tb.get("insideH")),
                "left": cb.get("left", row_tb.get("left") if c0 == 0 else row_tb.get("insideV")),
                "right": cb.get("right", row_tb.get("right") if c1 >= ncols else row_tb.get("insideV")),
            }
            cells.append(CellBox(x, w, mar, lines, ch, tcpr.get("vAlign", "top"), fill, borders,
                                 is_cont, cell.row_span if not is_cont else 1, c0=c0, c1=c1,
                                 own=dict(borders)))
        rows_cells.append(cells)

    # pass 2: resolve shared edges (the heavier border wins)
    for ri, cells in enumerate(rows_cells):
        # Word reserves a horizontal edge's width once, in the lower row when any
        # of its cells defines a top border, otherwise in the upper row
        lower_owns = any(c.borders["top"] is not None and c.borders["top"].visible
                         for c in cells if not c.is_continue)
        for k, c in enumerate(cells):
            if k > 0:
                prev = cells[k - 1]
                if prev.c1 == c.c0:
                    m = _heavier(prev.borders["right"], c.borders["left"])
                    prev.borders["right"] = m
                    c.borders["left"] = m
            if ri > 0:
                if c.is_continue:
                    c.borders["top"] = None
                else:
                    above = next((a for a in rows_cells[ri - 1] if a.c0 <= c.c0 < a.c1), None)
                    if above is not None:
                        own = c.borders["top"]
                        res = _heavier(above.borders.get("bottom"), own)
                        if lower_owns:
                            # the lower row reserves (and draws) the shared border
                            c.borders["top"] = res
                            above.borders["bottom"] = None
                        else:
                            # only the upper cell has it: the upper row reserves it,
                            # unless the lower cell's top margin can hold it
                            c.borders["top"] = None
                            above.borders["bottom"] = res
                            if _bw(res) > c.mar.get("top", 0):
                                above.bottom_band = True

    rows: list[RowFlow] = []
    pending_spans = []
    for ri, (row, cells) in enumerate(zip(tbl.rows, rows_cells)):
        trpr = row.trpr
        band_top = max((_bw(c.borders.get("top")) for c in cells), default=0.0)
        if ri == nrows - 1:
            band_bottom = max((_bw(c.borders.get("bottom")) for c in cells), default=0.0)
        else:
            band_bottom = max((_bw(c.borders.get("bottom")) for c in cells
                               if getattr(c, "bottom_band", False)), default=0.0)
        max_h = 0.0
        for c in cells:
            if c.is_continue:
                continue
            need = c.mar.get("top", 0) + c.content_h + c.mar.get("bottom", 0)
            if c.span_rows <= 1:
                max_h = max(max_h, need)
            else:
                pending_spans.append((ri, c.span_rows, c, need))
        h = band_top + max_h + band_bottom
        exact = False
        min_h = 0.0
        if "height" in trpr:
            live = [c for c in cells if not c.is_continue] or cells
            if trpr.get("hRule") == "exact" and trpr["height"] > 0:
                # Word: an exact height covers the top margin and border but
                # not the bottom cell margin
                h = trpr["height"] + max((c.mar.get("bottom", 0) for c in live), default=0.0)
                exact = True
            else:
                # an "at least" height is the minimum for the content area; cell
                # margins and border bands come on top of it
                margins = max((c.mar.get("top", 0) + c.mar.get("bottom", 0) for c in live), default=0.0)
                # (a bottom border band stays inside the minimum, a top one doesn't)
                min_h = trpr["height"] + margins + band_top
                h = max(h, min_h)
        rows.append(RowFlow(cells, h, header=trpr.get("header", False),
                            cant_split=trpr.get("cantSplit", False), exact=exact, table_x=x0,
                            first=ri == 0, last=ri == nrows - 1, band_top=band_top,
                            band_bottom=band_bottom, min_height=min_h))
    # vertical merges: grow the last spanned row if the merged content needs more room
    for ri, span, box, need in pending_spans:
        last = min(len(rows) - 1, ri + span - 1)
        have = sum(r.height for r in rows[ri:last + 1]) - rows[ri].band_top - rows[last].band_bottom
        if need > have + EPS and not rows[last].exact:
            rows[last].height += need - have
        box.span_height = sum(r.height for r in rows[ri:last + 1])
    for ri, row in enumerate(rows):
        for c in row.cells:
            if c.is_continue:
                nxt = rows[ri + 1] if ri + 1 < len(rows) else None
                c.span_last = not (nxt is not None and any(
                    abs(d.x - c.x) < 0.5 and d.is_continue for d in nxt.cells))
            else:
                c.span_last = c.span_rows <= 1
    header_rows = 0
    for r in rows:
        if r.header:
            header_rows += 1
        else:
            break
    return TableFlow(rows, header_rows, float=tbl.float, x0=x0, width=col_x[-1] - x0)


# --------------------------------------------------------------------------- blocks
def layout_blocks(blocks, width: float, ctx: LayoutContext) -> list:
    flows = []
    paras = [b for b in blocks]
    prev_para = None
    for idx, blk in enumerate(paras):
        if isinstance(blk, Paragraph):
            prev = paras[idx - 1] if idx > 0 else None
            nxt = paras[idx + 1] if idx + 1 < len(paras) else None
            bp = isinstance(prev, Paragraph) and _same_border(prev, blk)
            bn = isinstance(nxt, Paragraph) and _same_border(blk, nxt)
            pl = ParagraphLayout(blk, width, ctx, bp, bn)
            flow = pl.layout()
            flow.relayout = pl.layout
            if isinstance(prev_para, tuple):
                pflow, pblk = prev_para
                if pblk.style_id == blk.style_id:
                    if flow.contextual:
                        flow.space_before = 0.0
                    if pflow.contextual:
                        pflow.space_after = 0.0
                # HTML auto spacing: Word collapses to the larger value
                if pblk.ppr.get("sp_after_auto") and blk.ppr.get("sp_before_auto"):
                    m = max(pflow.space_after, flow.space_before)
                    pflow.space_after = 0.0
                    flow.space_before = m
            flows.append(flow)
            prev_para = (flow, blk)
        elif isinstance(blk, Table):
            flows.append(layout_table(blk, width, ctx))
            prev_para = None
    return flows


def _same_border(a: Paragraph, b: Paragraph) -> bool:
    ba = a.ppr.get("pBdr") or {}
    bb = b.ppr.get("pBdr") or {}
    if not ba or not bb:
        return False
    return ba == bb and a.ppr.get("ind_left", 0) == b.ppr.get("ind_left", 0) \
        and a.ppr.get("ind_right", 0) == b.ppr.get("ind_right", 0)


def gap(after: float, before: float, sum_spacing: bool) -> float:
    """Vertical space between two paragraphs."""
    return after + before if sum_spacing else max(after, before)


def stack_flows(flows, sum_spacing: bool = False) -> tuple[list, float]:
    """Stack flows vertically without pagination; returns ([(y, Line)], height)."""
    out = []
    y = 0.0
    prev_after = 0.0
    for f in flows:
        if isinstance(f, ParaFlow):
            y += gap(prev_after, f.space_before, sum_spacing)
            for k, ln in enumerate(f.lines):
                ln.para = (id(f), k, len(f.lines), f.space_after)
                out.append((y, ln))
                y += ln.height
            prev_after = f.space_after
        else:
            y += prev_after
            prev_after = 0.0
            for r in f.rows:
                out.append((y, r.as_line()))
                y += r.height
    y += prev_after
    return out, y
