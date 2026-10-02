"""Worksheet print layout: page setup, pagination and cell rendering like Excel."""
from __future__ import annotations

import math
import re
from dataclasses import dataclass

import openpyxl
from lxml import etree
from openpyxl.utils import column_index_from_string, range_boundaries

from ..draw import ClipOp, ImageOp, LineOp, RectOp, TextOp
from ..fonts import FontManager
from ..theme import Theme, apply_tint, hex_color
from .numfmt import format_value

PX = 0.75  # points per pixel at 96 dpi
EMU_PT = 12700.0

PAPER_SIZES = {  # paperSize code -> (width, height) in points
    1: (612, 792), 2: (612, 792), 3: (792, 1224), 4: (1224, 792), 5: (612, 1008),
    6: (396, 612), 7: (522, 756), 8: (842, 1191), 9: (595.3, 841.9), 10: (595.3, 841.9),
    11: (419.5, 595.3), 12: (729, 1032), 13: (516, 729), 14: (612, 936), 15: (609.4, 609.4),
    16: (720, 1008), 17: (792, 1224), 18: (612, 792), 19: (279, 639), 20: (297, 684),
    21: (324, 747), 22: (342, 792), 23: (360, 828), 24: (1224, 1584), 25: (1584, 2448),
    26: (2448, 3168), 27: (311.8, 623.6), 28: (459.2, 649.1), 29: (918.4, 1298.3),
    30: (649.1, 918.4), 31: (323.1, 459.2), 32: (323.1, 649.1), 33: (708.7, 1000.6),
    34: (498.9, 708.7), 35: (498.9, 354.3), 36: (311.8, 651.97), 37: (279, 540),
    38: (324, 648), 39: (1071, 792), 40: (612, 864), 41: (612, 936), 66: (1190.6, 1683.8),
    70: (297.6, 419.5),
}

INDEXED_COLORS = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
]

BORDER_WIDTHS = {  # style -> (line width pt, dash pattern)
    "hair": (0.25, (0.75, 0.75)), "thin": (0.5, None), "dotted": (0.5, (0.5, 1.0)),
    "dashed": (0.5, (2.25, 0.75)), "dashDot": (0.5, (2.25, 0.75, 0.75, 0.75)),
    "dashDotDot": (0.5, (2.25, 0.75, 0.75, 0.75, 0.75, 0.75)), "medium": (1.0, None),
    "mediumDashed": (1.0, (4.5, 1.5)), "mediumDashDot": (1.0, (4.5, 1.5, 1.5, 1.5)),
    "mediumDashDotDot": (1.0, (4.5, 1.5, 1.5, 1.5, 1.5, 1.5)), "slantDashDot": (1.0, (4.5, 1.5, 1.5, 1.5)),
    "thick": (1.5, None), "double": (0.5, None),
}


class WorkbookContext:
    """Workbook-wide metrics.

    Excel lays sheets out in screen pixels (96 dpi) but prints through the
    printer's device grid: column widths are recomputed with the printer's
    maximum digit width and row heights are scaled by the ratio of the
    printer's standard row height to the screen's.  ``dpi`` is the resolution
    of that virtual printer (600 matches "Microsoft Print to PDF" and most
    laser printers).
    """

    def __init__(self, wb, fonts: FontManager, theme: Theme, dpi: int = 600, locale=None):
        from .locale import get_locale
        self.wb = wb
        self.fonts = fonts
        self.theme = theme
        self.dpi = dpi
        self.locale = get_locale(locale)
        self.date1904 = bool(getattr(wb, "epoch", None) and wb.epoch.year == 1904)
        f0 = wb._fonts[0] if getattr(wb, "_fonts", None) else None
        self.default_font_name = self.font_name(f0) if f0 is not None else "Calibri"
        self.default_font_size = float(f0.sz) if f0 is not None and f0.sz else 11.0
        self._faces = {}
        face = fonts.face(self.default_font_name)
        f = face.font
        size = self.default_font_size
        digit = max(f.char_width(d, 1.0) for d in "0123456789")
        # maximum digit width in whole pixels (screen, hinted when the font says
        # so) and in device units (printer)
        ppem_s = int(round(size * 96 / 72))
        hinted = [f.gdi_advance(d, ppem_s) for d in "0123456789"]
        if all(h is not None for h in hinted):
            self.mdw = max(1, max(hinted))
        else:
            self.mdw = max(1, int(round(digit * size * 96 / 72)))
        ppem_p = int(round(size * dpi / 72))
        self.mdw_print = max(1, int(round(digit * ppem_p)))
        self.default_row_px = self.row_px_for(self.default_font_name, size, False, False)
        # the printer's standard row: GDI text height + external leading + 8 units
        asc, desc, ext = f.gdi_height(ppem_p)
        self.std_row_print = asc + desc + ext + int(round(8 * dpi / 600))
        self.row_scale = self.std_row_print / (self.default_row_px * PX)  # device units per point
        self.dev_pt = 72.0 / dpi
        self._dev = {}
        # text keeps 3 screen pixels from the left gridline and 2 from the right
        ratio = self.mdw_print / self.mdw
        self.pad_left = int(round(3 * ratio)) * self.dev_pt
        self.pad_right = (int(round(2 * ratio)) - 1) * self.dev_pt

    def font_name(self, font) -> str:
        """Font family for an openpyxl Font, honouring theme font schemes."""
        scheme = getattr(font, "scheme", None)
        if scheme == "minor":
            return self.theme.minor.get("latin") or font.name or "Calibri"
        if scheme == "major":
            return self.theme.major.get("latin") or font.name or "Calibri Light"
        return font.name or getattr(self, "default_font_name", "Calibri")

    def col_width_pt(self, chars: float) -> float:
        """Printed width of a column whose width is ``chars`` (as stored in the file)."""
        m = self.mdw_print
        dev = math.floor(((256 * chars + math.floor(128 / m)) / 256) * m)
        return dev * self.dev_pt

    def dev_font(self, name, size, bold=False, italic=False) -> "DevFont":
        key = (name, size, bold, italic)
        m = self._dev.get(key)
        if m is None:
            m = DevFont(self, self.face(name, bold, italic), size)
            self._dev[key] = m
        return m


    def face(self, name, bold, italic):
        key = (name, bold, italic)
        f = self._faces.get(key)
        if f is None:
            f = self.fonts.face(name, bold, italic)
            self._faces[key] = f
        return f

    def row_px_for(self, name, size, bold, italic) -> int:
        """Excel's auto-fit row height in screen pixels for a font."""
        f = self.fonts.face(name, bold, italic).font
        ppem = int(round(size * 96 / 72))
        asc, desc, ext = f.gdi_height(ppem)
        return asc + desc + ext + 2


    def color(self, c, default=None):
        """Resolve an openpyxl Color to an RGB tuple."""
        if c is None:
            return default
        try:
            t = c.type
        except AttributeError:
            return default
        rgb = None
        if t == "rgb" and isinstance(c.rgb, str):
            rgb = hex_color(c.rgb)
        elif t == "theme":
            rgb = self.theme.indexed_theme(int(c.theme))
        elif t == "indexed":
            idx = int(c.indexed)
            if idx == 64:  # system foreground
                return default if default is not None else (0, 0, 0)
            if idx == 65:
                return default
            if 0 <= idx < len(INDEXED_COLORS):
                rgb = hex_color(INDEXED_COLORS[idx])
        elif t == "auto":
            return default
        if rgb is None:
            return default
        tint = getattr(c, "tint", 0) or 0
        if tint:
            rgb = apply_tint(rgb, tint)
        return rgb


def load_workbook(source, fonts: FontManager, dpi: int = 600, locale=None):
    import os
    wb = openpyxl.load_workbook(source, data_only=True, rich_text=True)
    theme = Theme(etree.fromstring(wb.loaded_theme) if getattr(wb, "loaded_theme", None) else None)
    ctx = WorkbookContext(wb, fonts, theme, dpi, locale)
    ctx.file_name = os.path.basename(str(source)) if isinstance(source, (str, os.PathLike)) else "Book1.xlsx"
    return wb, ctx


# --------------------------------------------------------------------------- pages
@dataclass
class Page:
    width: float
    height: float
    ops: list
    header: tuple | None  # (left, center, right) raw strings
    footer: tuple | None
    margins: dict
    ctx: WorkbookContext
    sheet_name: str
    file_name: str
    scale: float = 1.0
    hf_scale: bool = True

    def render(self, number: int, total: int) -> list:
        ops = list(self.ops)
        for part, is_header in ((self.header, True), (self.footer, False)):
            if not part:
                continue
            for pos, txt in zip(("left", "center", "right"), part):
                if txt:
                    ops.extend(self._hf(txt, pos, is_header, number, total))
        return ops

    def _hf(self, raw: str, pos: str, is_header: bool, number: int, total: int):
        ctx = self.ctx
        runs = parse_hf(raw, ctx.default_font_name, ctx.default_font_size, number, total,
                        self.sheet_name, self.file_name)
        if not runs:
            return []
        lines = [[]]
        for r in runs:
            parts = r["text"].split("\n")
            for k, p in enumerate(parts):
                if k:
                    lines.append([])
                if p:
                    lines[-1].append(dict(r, text=p))
        s = self.scale if self.hf_scale else 1.0
        dev = ctx.dev_pt
        inset = round(8 * ctx.dpi / 600) * dev
        ops = []
        metrics = []
        for ln in lines:
            box = asc = 0.0
            w = 0.0
            for r in ln:
                m = ctx.dev_font(r["font"], r["size"] * s, r["bold"], r["italic"])
                r["_m"] = m
                r["_w"] = m.width(r["text"])
                w += r["_w"]
                box = max(box, m.box)
                asc = max(asc, m.lead + m.asc)
            if not ln:
                m = ctx.dev_font(ctx.default_font_name, ctx.default_font_size * s)
                box, asc = m.box, m.lead + m.asc
            metrics.append((box, asc, w))
        total_h = sum(bx for bx, _, _ in metrics)
        if is_header:
            y = self.margins["header"] * 72
        else:
            y = self.height - self.margins["footer"] * 72 - total_h
        ml = self.margins["left"] * 72
        mr = self.margins["right"] * 72
        for ln, (box, asc, w) in zip(lines, metrics):
            if pos == "left":
                x = ml + inset
            elif pos == "right":
                x = self.width - mr - inset - w
            else:
                x = (self.width - w) / 2
            base = y + asc
            for r in ln:
                m = r["_m"]
                ops.append(m.op(x, base, r["text"], r["color"]))
                f = m.face.font
                if r["underline"]:
                    uy = base - f.underline_pos * m.size
                    ops.append(LineOp(x, uy, x + r["_w"], uy, max(dev, f.underline_size * m.size), r["color"]))
                if r["strike"]:
                    sy = base - f.strike_pos * m.size
                    ops.append(LineOp(x, sy, x + r["_w"], sy, max(dev, f.strike_size * m.size), r["color"]))
                x += r["_w"]
            y += box
        return ops


def parse_hf(raw: str, font: str, size: float, page: int, total: int, sheet: str, fname: str):
    import datetime as _dt
    runs = []
    st = {"font": font, "size": size, "bold": False, "italic": False, "underline": False,
          "strike": False, "color": (0, 0, 0)}
    buf = []

    def flush():
        if buf:
            runs.append(dict(st, text="".join(buf)))
            buf.clear()
    i = 0
    n = len(raw)
    while i < n:
        c = raw[i]
        if c == "&" and i + 1 < n:
            d = raw[i + 1]
            i += 2
            if d == "&":
                buf.append("&")
            elif d == "P":
                m = re.match(r"([+-]\d+)", raw[i:])
                v = page
                if m:
                    v += int(m.group(1))
                    i += len(m.group(1))
                buf.append(str(v))
            elif d == "N":
                buf.append(str(total))
            elif d == "D":
                buf.append(_dt.date.today().strftime("%m/%d/%Y").lstrip("0").replace("/0", "/"))
            elif d == "T":
                buf.append(_dt.datetime.now().strftime("%I:%M %p").lstrip("0"))
            elif d == "F":
                buf.append(fname)
            elif d == "A":
                buf.append(sheet)
            elif d == "Z":
                buf.append("")
            elif d in "BIUSEXYGKH":
                flush()
                if d == "B":
                    st["bold"] = not st["bold"]
                elif d == "I":
                    st["italic"] = not st["italic"]
                elif d in "UE":
                    st["underline"] = not st["underline"]
                elif d == "S":
                    st["strike"] = not st["strike"]
                elif d == "K":
                    col = raw[i:i + 6]
                    i += 6
                    st["color"] = hex_color(col) or (0, 0, 0)
            elif d == '"':
                flush()
                j = raw.find('"', i)
                if j == -1:
                    j = n
                spec = raw[i:j]
                i = j + 1
                name, _, style = spec.partition(",")
                if name and name != "-":
                    st["font"] = name
                sl = style.lower()
                st["bold"] = "bold" in sl
                st["italic"] = "italic" in sl or "oblique" in sl
            elif d.isdigit():
                flush()
                m = re.match(r"\d+", raw[i - 1:])
                st["size"] = float(m.group(0))
                i = i - 1 + len(m.group(0))
            elif d in "LCR":
                pass
            else:
                buf.append(d)
        else:
            buf.append(c)
            i += 1
    flush()
    return runs


# --------------------------------------------------------------------------- printer
class SheetPrinter:
    def __init__(self, ws, ctx: WorkbookContext, ignore_print_areas=False):
        self.ws = ws
        self.ctx = ctx
        self.ignore_print_areas = ignore_print_areas
        self._col_pt: dict[int, float] = {}
        self._row_pt: dict[int, float] = {}
        fmt = ws.sheet_format
        # screen default row height (points) and default column width (chars)
        self.default_row_pt = (float(fmt.defaultRowHeight) if fmt.defaultRowHeight
                               else ctx.default_row_px * PX)
        # printed rows scale by printer standard row / this sheet's standard row
        self.row_scale = ctx.std_row_print / self.default_row_pt
        if fmt.defaultColWidth:
            self.default_col_chars = float(fmt.defaultColWidth)
        else:
            base = fmt.baseColWidth if fmt.baseColWidth is not None else 8
            px = int(math.ceil((base * ctx.mdw + 5) / 8.0) * 8)  # rounded up to 8 px
            self.default_col_chars = px / ctx.mdw
        self.default_col_print = round(self.default_col_chars * ctx.mdw_print) * ctx.dev_pt
        self.merged: dict[tuple[int, int], tuple[int, int]] = {}
        self.merged_inner: set[tuple[int, int]] = set()
        for rng in ws.merged_cells.ranges:
            c1, r1, c2, r2 = rng.min_col, rng.min_row, rng.max_col, rng.max_row
            self.merged[(r1, c1)] = (r2, c2)
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (r, c) != (r1, c1):
                        self.merged_inner.add((r, c))
        self._row_cells: dict[int, list] = {}
        for (r, c), cell in ws._cells.items():
            self._row_cells.setdefault(r, []).append(cell)
        self._col_dims = {}
        for key, dim in ws.column_dimensions.items():
            try:
                lo = dim.min or column_index_from_string(key)
                hi = dim.max or lo
            except Exception:
                continue
            for c in range(lo, hi + 1):
                self._col_dims[c] = dim

    # -- geometry ----------------------------------------------------------
    def col_pt(self, c: int) -> float:
        v = self._col_pt.get(c)
        if v is None:
            dim = self._col_dims.get(c)
            if dim is not None and dim.hidden:
                v = 0.0
            elif dim is not None and dim.width:
                v = self.ctx.col_width_pt(float(dim.width))
            else:
                v = self.default_col_print
            self._col_pt[c] = v
        return v

    def row_pt(self, r: int) -> float:
        v = self._row_pt.get(r)
        if v is None:
            dim = self.ws.row_dimensions.get(r) if r in self.ws.row_dimensions else None
            if dim is not None and dim.hidden:
                v = 0.0
            elif dim is not None and dim.ht is not None:
                v = self._print_row(float(dim.ht))
            else:
                v = self._print_row(self._autofit_row(r))
            self._row_pt[r] = v
        return v

    def _print_row(self, pt: float) -> float:
        """Printed height of a row that is ``pt`` points tall on screen."""
        return math.floor(pt * self.row_scale + 1e-6) * self.ctx.dev_pt

    def _autofit_row(self, r: int) -> float:
        best = self.default_row_pt
        for cell in self._row_cells.get(r, ()):
            if cell.value is None and not cell.has_style:
                continue
            f = cell.font
            name = self.ctx.font_name(f)
            size = float(f.sz or self.ctx.default_font_size)
            if cell.value is None:
                continue
            px = self.ctx.row_px_for(name, size, bool(f.b), bool(f.i))
            h = px * PX
            al = cell.alignment
            if al is not None and al.wrap_text and isinstance(cell.value, str):
                m = self.ctx.dev_font(name, size, bool(f.b), bool(f.i))
                width = self.col_pt(cell.column) - self.ctx.pad_left - self.ctx.pad_right
                nlines = len(self._wrap(cell.value, m, width))
                h = max(h, nlines * px * PX)
            if al is not None and al.text_rotation and al.text_rotation not in (0, 255):
                face = self.ctx.face(name, bool(f.b), bool(f.i))
                txt = str(cell.value)
                w = face.font.width(txt, size) * 4 / 3
                ang = math.radians(al.text_rotation if al.text_rotation <= 90 else al.text_rotation - 90)
                h = max(h, (w * math.sin(ang)) * PX + h * math.cos(ang))
            best = max(best, h)
        return best


    # -- ranges ------------------------------------------------------------
    def used_range(self):
        ws = self.ws
        min_r = min_c = None
        max_r = max_c = 0
        for (r, c), cell in ws._cells.items():
            visible = cell.value is not None
            if not visible and cell.has_style:
                fill = cell.fill
                if fill is not None and fill.fill_type not in (None, "none"):
                    visible = True
                b = cell.border
                if b is not None and any(getattr(b, s) is not None and getattr(b, s).style
                                         for s in ("left", "right", "top", "bottom")):
                    visible = True
            if not visible:
                continue
            min_r = r if min_r is None else min(min_r, r)
            min_c = c if min_c is None else min(min_c, c)
            max_r = max(max_r, r)
            max_c = max(max_c, c)
        for (r1, c1), (r2, c2) in self.merged.items():
            if min_r is None:
                min_r, min_c = r1, c1
            max_r = max(max_r, r2)
            max_c = max(max_c, c2)
        for img in getattr(ws, "_images", []):
            try:
                a = img.anchor._from
                r, c = a.row + 1, a.col + 1
                x1, y1, rr, cc = self._anchor_extent(img)
                min_r = r if min_r is None else min(min_r, r)
                min_c = c if min_c is None else min(min_c, c)
                max_r = max(max_r, rr)
                max_c = max(max_c, cc)
            except Exception:
                pass
        if min_r is None:
            return None
        max_c = max(max_c, self._overflow_extent(max_c))
        return (1, 1, max_r, max_c)

    def _overflow_extent(self, max_c: int) -> int:
        """Last column reached by left-aligned text spilling out of its cell."""
        ctx = self.ctx
        best = max_c
        for (r, c), cell in self.ws._cells.items():
            v = cell.value
            if not isinstance(v, str) or not v or (r, c) in self.merged:
                continue
            al = cell.alignment
            if al is not None and (al.wrap_text or al.horizontal not in (None, "general", "left")):
                continue
            name, size, bold, italic, _ = self._font(cell.font)
            need = ctx.dev_font(name, size, bold, italic).width(v) + ctx.pad_left
            cc = c
            room = self.col_pt(c)
            while room < need and self._empty(r, cc + 1) and cc < 16384:
                cc += 1
                room += self.col_pt(cc)
            best = max(best, cc)
        return best

    def print_areas(self):
        ws = self.ws
        if not self.ignore_print_areas and ws.print_area:
            areas = ws.print_area
            if isinstance(areas, str):
                areas = [a for a in areas.split(",") if a]
            out = []
            for a in areas:
                ref = a.split("!")[-1].replace("$", "")
                try:
                    c1, r1, c2, r2 = range_boundaries(ref)
                except Exception:
                    continue
                out.append((r1, c1, r2, c2))
            if out:
                return out
        u = self.used_range()
        return [u] if u else []

    def _titles(self):
        rows = cols = None
        ws = self.ws
        t = ws.print_title_rows
        if t:
            m = re.match(r"\$?(\d+):\$?(\d+)", t.split("!")[-1])
            if m:
                rows = (int(m.group(1)), int(m.group(2)))
        t = ws.print_title_cols
        if t:
            m = re.match(r"\$?([A-Z]+):\$?([A-Z]+)", t.split("!")[-1])
            if m:
                cols = (column_index_from_string(m.group(1)), column_index_from_string(m.group(2)))
        return rows, cols

    # -- pagination ----------------------------------------------------------
    def pages(self) -> list[Page]:
        ws = self.ws
        ps = ws.page_setup
        paper = PAPER_SIZES.get(int(ps.paperSize) if ps.paperSize else 1, (612, 792))
        pw, ph = paper
        if ps.orientation == "landscape":
            pw, ph = max(pw, ph), min(pw, ph)
        else:
            pw, ph = min(pw, ph), max(pw, ph)
        pm = ws.page_margins
        margins = {"left": pm.left if pm.left is not None else 0.7,
                   "right": pm.right if pm.right is not None else 0.7,
                   "top": pm.top if pm.top is not None else 0.75,
                   "bottom": pm.bottom if pm.bottom is not None else 0.75,
                   "header": pm.header if pm.header is not None else 0.3,
                   "footer": pm.footer if pm.footer is not None else 0.3}
        avail_w = pw - (margins["left"] + margins["right"]) * 72
        avail_h = ph - (margins["top"] + margins["bottom"]) * 72
        title_rows, title_cols = self._titles()
        po = ws.print_options
        grid = bool(po.gridLines) if po is not None else False
        headings = bool(po.headings) if po is not None else False
        hcenter = bool(po.horizontalCentered) if po is not None else False
        vcenter = bool(po.verticalCentered) if po is not None else False
        fit = bool(ws.sheet_properties.pageSetUpPr and ws.sheet_properties.pageSetUpPr.fitToPage)
        row_breaks = {b.id for b in (ws.row_breaks.brk if ws.row_breaks else [])}
        col_breaks = {b.id for b in (ws.col_breaks.brk if ws.col_breaks else [])}
        hf = ws.HeaderFooter if hasattr(ws, "HeaderFooter") else None
        pages: list[Page] = []
        for area in self.print_areas():
            r1, c1, r2, c2 = area
            cols = [c for c in range(c1, c2 + 1)]
            rows = [r for r in range(r1, r2 + 1)]
            tr = [r for r in range(title_rows[0], title_rows[1] + 1)] if title_rows else []
            tc = [c for c in range(title_cols[0], title_cols[1] + 1)] if title_cols else []

            def bands(items, size_fn, avail, breaks, titles):
                out = []
                cur = []
                used = 0.0
                tsize = sum(size_fn(t) for t in titles)
                for it in items:
                    s = size_fn(it)
                    # repeated titles take room on pages that do not already start with them
                    extra = tsize if (titles and cur and cur[0] > max(titles)) or \
                        (titles and not cur and it > max(titles)) else 0.0
                    if cur and used + s + extra > avail + 0.01:
                        out.append(cur)
                        cur = []
                        used = 0.0
                    cur.append(it)
                    used += s
                    if it in breaks:
                        out.append(cur)
                        cur = []
                        used = 0.0
                if cur:
                    out.append(cur)
                return out

            scale = (ps.scale or 100) / 100.0
            if fit:
                fw = ps.fitToWidth if ps.fitToWidth is not None else 1
                fh = ps.fitToHeight if ps.fitToHeight is not None else 1
                scale = 1.0
                for pct in range(100, 9, -1):
                    s = pct / 100.0
                    cb = bands(cols, self.col_pt, avail_w / s, col_breaks, tc)
                    rb = bands(rows, self.row_pt, avail_h / s, row_breaks, tr)
                    ok_w = fw == 0 or len(cb) <= fw
                    ok_h = fh == 0 or len(rb) <= fh
                    if ok_w and ok_h:
                        scale = s
                        break
                    scale = s
            col_bands = bands(cols, self.col_pt, avail_w / scale, col_breaks, tc)
            row_bands = bands(rows, self.row_pt, avail_h / scale, row_breaks, tr)
            order = []
            if ps.pageOrder == "overThenDown":
                for rb in row_bands:
                    for cb in col_bands:
                        order.append((rb, cb))
            else:
                for cb in col_bands:
                    for rb in row_bands:
                        order.append((rb, cb))
            for rb, cb in order:
                prow = ([r for r in tr if r < rb[0]] if tr else []) + rb
                pcol = ([c for c in tc if c < cb[0]] if tc else []) + cb
                if not any(self.row_pt(r) for r in prow) or not any(self.col_pt(c) for c in pcol):
                    continue
                ops, has_content = self.render_block(prow, pcol, grid, headings)
                if not has_content:
                    continue  # Excel does not print blank pages
                content_w = sum(self.col_pt(c) for c in pcol) * scale
                content_h = sum(self.row_pt(r) for r in prow) * scale
                x0 = margins["left"] * 72
                y0 = margins["top"] * 72
                if hcenter:
                    x0 += max(0.0, (avail_w - content_w) / 2)
                if vcenter:
                    y0 += max(0.0, (avail_h - content_h) / 2)
                ops = _transform(ops, x0, y0, scale)
                header = footer = None
                page_idx = len(pages)
                if hf is not None:
                    header = _hf_parts(hf, "header", page_idx)
                    footer = _hf_parts(hf, "footer", page_idx)
                pages.append(Page(pw, ph, ops, header, footer, margins, self.ctx, ws.title,
                                  getattr(self.ctx, "file_name", "Book1.xlsx"), scale,
                                  hf.scaleWithDoc if hf is not None and hf.scaleWithDoc is not None else True))
        return pages

    # -- rendering -----------------------------------------------------------
    def render_block(self, rows: list[int], cols: list[int], grid: bool, headings: bool) -> list:
        ws = self.ws
        xs = {}
        x = 0.0
        for c in cols:
            xs[c] = x
            x += self.col_pt(c)
        total_w = x
        ys = {}
        y = 0.0
        for r in rows:
            ys[r] = y
            y += self.row_pt(r)
        total_h = y
        col_set = set(cols)
        row_set = set(rows)
        fills: list = []
        borders: list = []
        texts: list = []
        gridlines: list = []
        # occupied cells prevent overflow
        cells = ws._cells

        def cell_rect(r, c):
            x0 = xs[c]
            y0 = ys[r]
            w = self.col_pt(c)
            h = self.row_pt(r)
            if (r, c) in self.merged:
                r2, c2 = self.merged[(r, c)]
                w = sum(self.col_pt(cc) for cc in range(c, c2 + 1) if cc in col_set) or w
                h = sum(self.row_pt(rr) for rr in range(r, r2 + 1) if rr in row_set) or h
            return x0, y0, w, h

        if grid:
            gc = (0.0, 0.0, 0.0)
            for c in cols + [None]:
                gx = xs[c] if c is not None else total_w
                gridlines.append(LineOp(gx, 0, gx, total_h, 0.25, gc))
            for r in rows + [None]:
                gy = ys[r] if r is not None else total_h
                gridlines.append(LineOp(0, gy, total_w, gy, 0.25, gc))

        # columns to the left of this page, for text that spills onto it
        left_cols = []
        lx = 0.0
        c = cols[0] - 1
        while c >= 1 and lx > -1500 and len(left_cols) < 64:
            lx -= self.col_pt(c)
            xs[c] = lx
            left_cols.insert(0, c)
            c -= 1
        all_cols = left_cols + cols
        for r in rows:
            if not self.row_pt(r):
                continue
            for c in cols:
                if not self.col_pt(c):
                    continue
                cell = cells.get((r, c))
                if cell is None:
                    continue
                inner = (r, c) in self.merged_inner
                x0, y0, w, h = cell_rect(r, c)
                if cell.has_style and not inner:
                    fill = cell.fill
                    fc = self._fill_color(fill)
                    if fc is not None:
                        fills.append(RectOp(x0, y0, w, h, fill=fc))
                if cell.has_style:
                    self._borders(cell, xs[c], ys[r], self.col_pt(c), self.row_pt(r), borders)
                if inner or cell.value is None:
                    continue
                texts.extend(self._cell_text(cell, r, c, x0, y0, w, h, xs, all_cols, col_set))
            # left-aligned text from earlier columns overflowing into this page
            for c in reversed(left_cols):
                cell = cells.get((r, c))
                if cell is None or cell.value is None:
                    continue
                if not isinstance(cell.value, str) or (r, c) in self.merged or (r, c) in self.merged_inner:
                    break
                al = cell.alignment
                if al is not None and (al.wrap_text or al.horizontal not in (None, "general", "left")):
                    break
                x0, y0 = xs[c], ys[r]
                ops = self._cell_text(cell, r, c, x0, y0, self.col_pt(c), self.row_pt(r), xs, all_cols, col_set)
                if ops and ops[0].x + ops[0].w > 0.01:
                    texts.extend(ops)
                break
        ops = fills + gridlines + borders + texts
        images = self._images(xs, ys, rows, cols)
        ops.extend(images)
        has_content = bool(fills or borders or texts or images)
        return [ClipOp(0, 0, total_w, total_h, ops)], has_content

    def _fill_color(self, fill):
        if fill is None:
            return None
        ft = getattr(fill, "fill_type", None) or getattr(fill, "patternType", None)
        if ft in (None, "none"):
            if fill.__class__.__name__ == "GradientFill" and fill.stop:
                cols = [self.ctx.color(s.color) for s in fill.stop]
                cols = [c for c in cols if c]
                if cols:
                    return tuple(sum(c[i] for c in cols) / len(cols) for i in range(3))
            return None
        fg = self.ctx.color(fill.fgColor, None)
        bg = self.ctx.color(fill.bgColor, (1, 1, 1))
        if ft == "solid":
            return fg if fg is not None else (0, 0, 0)
        cov = {"gray125": 0.125, "gray0625": 0.0625, "lightGray": 0.25, "mediumGray": 0.5,
               "darkGray": 0.75}.get(ft, 0.5)
        fg = fg or (0, 0, 0)
        return tuple(bg[i] * (1 - cov) + fg[i] * cov for i in range(3))

    def _borders(self, cell, x0, y0, w, h, out):
        b = cell.border
        if b is None:
            return
        for side, coords in (("top", (x0, y0, x0 + w, y0)), ("bottom", (x0, y0 + h, x0 + w, y0 + h)),
                             ("left", (x0, y0, x0, y0 + h)), ("right", (x0 + w, y0, x0 + w, y0 + h))):
            s = getattr(b, side)
            if s is None or not s.style:
                continue
            width, dash = BORDER_WIDTHS.get(s.style, (0.5, None))
            color = self.ctx.color(s.color, (0, 0, 0))
            if s.style == "double":
                gap = 0.75
                if side in ("top", "bottom"):
                    out.append(LineOp(coords[0], coords[1] - gap, coords[2], coords[3] - gap, 0.5, color))
                    out.append(LineOp(coords[0], coords[1] + gap, coords[2], coords[3] + gap, 0.5, color))
                else:
                    out.append(LineOp(coords[0] - gap, coords[1], coords[2] - gap, coords[3], 0.5, color))
                    out.append(LineOp(coords[0] + gap, coords[1], coords[2] + gap, coords[3], 0.5, color))
            else:
                out.append(LineOp(*coords, width, color, dash, cap=2 if not dash else 0))
        if b.diagonal is not None and b.diagonal.style:
            width, dash = BORDER_WIDTHS.get(b.diagonal.style, (0.5, None))
            color = self.ctx.color(b.diagonal.color, (0, 0, 0))
            if b.diagonalDown:
                out.append(LineOp(x0, y0, x0 + w, y0 + h, width, color, dash))
            if b.diagonalUp:
                out.append(LineOp(x0, y0 + h, x0 + w, y0, width, color, dash))

    # -- text ----------------------------------------------------------------
    def _font(self, f):
        name = self.ctx.font_name(f)
        size = float(f.sz) if f.sz else self.ctx.default_font_size
        bold = bool(f.b)
        italic = bool(f.i)
        color = self.ctx.color(f.color, (0, 0, 0))
        return name, size, bold, italic, color

    def _wrap(self, text: str, m: "DevFont", width: float) -> list[str]:
        out = []
        for para in text.split("\n"):
            words = re.split(r"( +)", para)
            line = ""
            for w in words:
                if not w:
                    continue
                cand = line + w
                if m.width(cand.rstrip()) <= width + 1e-6 or not line.strip():
                    if m.width(cand.rstrip()) > width and not line.strip() and w.strip():
                        buf = line
                        for ch in w:
                            if m.width((buf + ch).rstrip()) > width and buf.strip():
                                out.append(buf.rstrip())
                                buf = ch
                            else:
                                buf += ch
                        line = buf
                    else:
                        line = cand
                else:
                    out.append(line.rstrip())
                    line = "" if w.isspace() else w
            out.append(line.rstrip())
        return out

    def _cell_text(self, cell, r, c, x0, y0, w, h, xs, cols, col_set):
        ctx = self.ctx
        dev = ctx.dev_pt
        value = cell.value
        name, size, bold, italic, color = self._font(cell.font)
        al = cell.alignment
        horiz = (al.horizontal if al is not None else None) or "general"
        vert = (al.vertical if al is not None else None) or "bottom"
        wrap = bool(al.wrap_text) if al is not None else False
        shrink = bool(al.shrink_to_fit) if al is not None else False
        indent = float(al.indent or 0) if al is not None else 0.0
        rot = int(al.text_rotation or 0) if al is not None else 0
        rich = None
        if value.__class__.__name__ == "CellRichText":
            rich = value
            value = str(value)
        if isinstance(value, str) and cell.data_type == "f":
            value = ""  # formula without a cached result
        code = self._number_format(cell)
        fmt = format_value(value, code, ctx.date1904, loc=ctx.locale)
        if fmt.color is not None:
            color = fmt.color
        is_num = fmt.is_number
        if horiz == "general":
            if isinstance(value, bool) or (isinstance(value, str) and value.upper() in _ERRORS):
                horiz = "center"
            elif is_num:
                horiz = "right"
            else:
                horiz = "left"
        m = ctx.dev_font(name, size, bold, italic)
        lpad, rpad = ctx.pad_left, ctx.pad_right
        ind = indent * 3 * m.char(" ") if indent else 0.0
        avail = w - lpad - rpad - ind
        text = fmt.plain()
        if is_num and not wrap:
            if getattr(fmt, "general", False):
                chars = 11
                while chars > 1 and m.width(text) > avail + 1e-6:
                    chars -= 1
                    text = format_value(value, code, ctx.date1904, general_chars=chars,
                                        loc=ctx.locale).plain()
                if m.width(text) > avail + 1e-6:
                    text = "#" * max(1, int(avail / max(dev, m.char("#"))))
                fmt.parts = [("text", text)]
            elif _parts_width(fmt.parts, m) > avail + 1e-6 and not shrink:
                text = "#" * max(1, int(avail / max(dev, m.char("#"))))
                fmt.parts = [("text", text)]
        if shrink and m.width(text) > avail > 0:
            size = size * avail / m.width(text)
            m = ctx.dev_font(name, size, bold, italic)
        underline = cell.font.u if cell.font.u not in (None, "none") else None
        strike = bool(cell.font.strike)
        box = m.box  # Excel's text line box: ascent + descent + 8 device units

        def first_baseline(n_lines):
            block = n_lines * box
            if vert == "top":
                top = y0
            elif vert in ("center", "justify", "distributed"):
                top = y0 + (h - block) / 2
            else:
                top = y0 + h - block
            return top + m.lead + m.asc

        ops = []
        if rot == 255:
            yb = first_baseline(len(text))
            for ch in text:
                cw = m.width(ch)
                ops.append(m.op(x0 + (w - cw) / 2, yb, ch, color))
                yb += box
            return [ClipOp(x0, y0, w, h, ops)]
        if rot:
            angle = rot if rot <= 90 else -(rot - 90)
            tw = m.width(text)
            rad = math.radians(abs(angle))
            # the rotated line box is centred horizontally unless aligned left/right
            if angle > 0:
                if horiz == "left":
                    bx = x0 + lpad + m.asc * math.sin(rad)
                elif horiz == "right":
                    bx = x0 + w - rpad - tw * math.cos(rad) - m.desc * math.sin(rad)
                else:
                    bx = x0 + (w - tw * math.cos(rad)) / 2 + (m.asc - m.desc) / 2 * math.sin(rad)
                span = tw * math.sin(rad) + (m.asc + m.desc) * math.cos(rad)
                if vert == "top":
                    by = y0 + m.lead + span - m.desc * math.cos(rad)
                elif vert == "center":
                    by = y0 + (h + span) / 2 - m.desc * math.cos(rad)
                else:
                    by = y0 + h - m.lead - m.desc * math.cos(rad)
            else:
                if horiz == "left":
                    bx = x0 + lpad + m.desc * math.sin(rad)
                elif horiz == "right":
                    bx = x0 + w - rpad - m.asc * math.sin(rad)
                else:
                    bx = x0 + w / 2 - (m.asc - m.desc) / 2 * math.sin(rad)
                span = tw * math.sin(rad)
                if vert == "top":
                    by = y0 + m.lead
                elif vert == "center":
                    by = y0 + (h - span) / 2
                else:
                    by = y0 + h - m.lead - span
            op = m.op(bx, by, text, color)
            op.angle = angle
            ops.append(op)
            return [ClipOp(x0, y0, w, h, ops)]

        if wrap or (horiz in ("justify", "distributed") and isinstance(value, str)):
            lines = self._wrap(text, m, max(dev, avail))
            base = first_baseline(len(lines))
            for ln in lines:
                lw = m.width(ln)
                if horiz in ("center", "centerContinuous"):
                    lx = x0 + lpad + (w - lpad - rpad - lw) / 2
                elif horiz == "right":
                    lx = x0 + w - rpad - lw - ind
                else:
                    lx = x0 + lpad + ind
                ops.append(m.op(lx, base, ln, color))
                self._decorate(ops, lx, base, lw, m, color, underline, strike)
                base += box
            return [ClipOp(x0, y0, w, h, ops)]

        # single line
        parts = fmt.parts if not rich else [("text", text)]
        tw = _parts_width(parts, m)
        fill_idx = next((i for i, p in enumerate(parts) if p[0] == "fill"), None)
        if fill_idx is not None and tw < avail:
            ch = parts[fill_idx][1]
            cw = m.char(ch)
            if cw > 0:
                reps = int((avail - tw) / cw)
                parts = list(parts)
                parts[fill_idx] = ("text", ch * reps)
                parts.insert(fill_idx, ("space", avail - tw - reps * cw))
                tw = avail
        clip_x0, clip_x1 = x0, x0 + w
        if not is_num and (r, c) not in self.merged and tw > avail:
            idx = cols.index(c)
            if horiz in ("left", "general", "fill", "center", "centerContinuous"):
                need_right = tw + lpad + ind if horiz not in ("center", "centerContinuous") \
                    else (w + tw) / 2 + lpad
                right = x0 + w
                j = idx
                while right - x0 < need_right and j + 1 < len(cols):
                    nxt = cols[j + 1]
                    if not self._empty(r, nxt):
                        break
                    j += 1
                    right += self.col_pt(nxt)
                clip_x1 = right
            if horiz in ("right", "center", "centerContinuous"):
                need_left = tw + rpad if horiz == "right" else (w + tw) / 2 + rpad
                left = x0
                j = idx
                while (x0 + w) - left < need_left and j - 1 >= 0:
                    prv = cols[j - 1]
                    if not self._empty(r, prv):
                        break
                    j -= 1
                    left -= self.col_pt(prv)
                clip_x0 = left
        if horiz in ("center", "centerContinuous"):
            tx = x0 + lpad + (w - lpad - rpad - tw) / 2
        elif horiz == "right":
            tx = x0 + w - rpad - tw - ind
        else:
            tx = x0 + lpad + ind
        base = first_baseline(1)
        if rich:
            ops.extend(self._rich(rich, tx, base, cell.font, color))
        else:
            x = tx
            for kind, t in parts:
                if kind == "text":
                    if t:
                        ops.append(m.op(x, base, t, color))
                    x += m.width(t)
                elif kind == "pad":
                    x += m.char(t)
                elif kind == "space":
                    x += t
            self._decorate(ops, tx, base, tw, m, color, underline, strike)
        return [ClipOp(clip_x0, y0, clip_x1 - clip_x0, h, ops)]

    def _empty(self, r, c) -> bool:
        cell = self.ws._cells.get((r, c))
        if cell is not None and cell.value is not None and cell.value != "":
            return False
        return (r, c) not in self.merged and (r, c) not in self.merged_inner

    def _number_format(self, cell) -> str:
        """The format code Excel would use, including locale-specific built-ins."""
        code = cell.number_format
        try:
            fid = cell._style.numFmtId
        except AttributeError:
            return code
        if fid in LOCALE_BUILTIN_IDS and self.ctx.locale.name.split("-")[0] != "en":
            return _localized_builtin(fid, self.ctx.locale) or code
        return code

    def _rich(self, rich, x, base, cell_font, cell_color):
        ops = []
        for block in rich:
            if isinstance(block, str):
                text = block
                f = cell_font
                name, size, bold, italic, color = self._font(f)
                u = cell_font.u
            else:
                text = block.text
                f = block.font
                name = f.rFont or self.ctx.font_name(cell_font)
                size = float(f.sz) if f.sz else float(cell_font.sz or self.ctx.default_font_size)
                bold = bool(f.b) if f.b is not None else bool(cell_font.b)
                italic = bool(f.i) if f.i is not None else bool(cell_font.i)
                color = self.ctx.color(f.color, cell_color)
                u = getattr(f, "u", None)
            shift = 0.0
            va = getattr(f, "vertAlign", None)
            if va == "superscript":
                shift = -size * 0.33
                size *= 0.65
            elif va == "subscript":
                shift = size * 0.14
                size *= 0.65
            m = self.ctx.dev_font(name, size, bold, italic)
            ops.append(m.op(x, base + shift, text, color))
            w = m.width(text)
            if u and u != "none":
                self._decorate(ops, x, base, w, m, color, u, False)
            x += w
        return ops

    def _decorate(self, ops, x, base, w, m, color, underline, strike):
        f = m.face.font
        size = m.size
        if underline:
            th = max(m.dev, f.underline_size * size)
            y = base - f.underline_pos * size + th / 2
            if underline in ("double", "doubleAccounting"):
                ops.append(LineOp(x, y - th, x + w, y - th, th * 0.8, color))
                ops.append(LineOp(x, y + th, x + w, y + th, th * 0.8, color))
            else:
                ops.append(LineOp(x, y, x + w, y, th, color))
        if strike:
            y = base - f.strike_pos * size
            ops.append(LineOp(x, y, x + w, y, max(m.dev, f.strike_size * size), color))

    # -- images ----------------------------------------------------------------
    def _anchor_extent(self, img):
        a = img.anchor
        fr = a._from
        c0, r0 = fr.col + 1, fr.row + 1
        x = sum(self.col_pt(c) for c in range(1, c0)) + fr.colOff / EMU_PT
        y = sum(self.row_pt(r) for r in range(1, r0)) + fr.rowOff / EMU_PT
        to = getattr(a, "to", None)
        if to is not None:
            return x, y, to.row + 1, to.col + 1
        ext = getattr(a, "ext", None)
        w = (ext.width / EMU_PT) if ext is not None and ext.width else img.width * PX
        h = (ext.height / EMU_PT) if ext is not None and ext.height else img.height * PX
        # find the last row/col touched
        rr, yy = r0, y
        while yy + self.row_pt(rr) < y + h and rr < 1048576:
            yy += self.row_pt(rr)
            rr += 1
        cc, xx = c0, x
        while xx + self.col_pt(cc) < x + w and cc < 16384:
            xx += self.col_pt(cc)
            cc += 1
        return x, y, rr, cc

    def _images(self, xs, ys, rows, cols):
        ops = []
        for img in getattr(self.ws, "_images", []):
            try:
                a = img.anchor
                fr = a._from
                c0, r0 = fr.col + 1, fr.row + 1
                if c0 not in xs or r0 not in ys:
                    # starts on another page: position relative to this page's origin
                    x = sum(self.col_pt(c) for c in range(cols[0], c0)) if c0 >= cols[0] else \
                        -sum(self.col_pt(c) for c in range(c0, cols[0]))
                    y = sum(self.row_pt(r) for r in range(rows[0], r0)) if r0 >= rows[0] else \
                        -sum(self.row_pt(r) for r in range(r0, rows[0]))
                else:
                    x, y = xs[c0], ys[r0]
                x += fr.colOff / EMU_PT
                y += fr.rowOff / EMU_PT
                to = getattr(a, "to", None)
                if to is not None:
                    x2 = x - fr.colOff / EMU_PT + sum(self.col_pt(c) for c in range(c0, to.col + 1)) + to.colOff / EMU_PT
                    y2 = y - fr.rowOff / EMU_PT + sum(self.row_pt(r) for r in range(r0, to.row + 1)) + to.rowOff / EMU_PT
                    w, h = x2 - x, y2 - y
                else:
                    ext = getattr(a, "ext", None)
                    w = (ext.width / EMU_PT) if ext is not None and ext.width else img.width * PX
                    h = (ext.height / EMU_PT) if ext is not None and ext.height else img.height * PX
                data = img._data()
                ops.append(ImageOp(x, y, w, h, data))
            except Exception:
                continue
        return ops


class DevFont:
    """A font at the virtual printer's resolution.

    Excel prints text with integer device metrics: the font is realised at a
    whole number of pixels per em and every glyph advance is rounded to a
    device unit.  Widths and positions computed here reproduce that grid.
    """

    def __init__(self, ctx: "WorkbookContext", face, size: float):
        self.face = face
        self.dev = ctx.dev_pt
        k = ctx.dpi / 600.0
        self.ppem = max(1, int(round(size * ctx.dpi / 72)))
        self.size = self.ppem * self.dev
        f = face.font
        asc, desc, _ = f.gdi_height(self.ppem)
        self.asc = asc * self.dev
        self.desc = desc * self.dev
        self.lead = round(7 * k) * self.dev
        self.box = self.asc + self.desc + round(8 * k) * self.dev
        self._cw: dict[str, float] = {}

    def char(self, ch: str) -> float:
        w = self._cw.get(ch)
        if w is None:
            w = int(round(self.face.font.char_width(ch, self.ppem))) * self.dev
            self._cw[ch] = w
        return w

    def width(self, text: str) -> float:
        return sum(self.char(ch) for ch in text)

    def op(self, x, y, text, color):
        natural = self.face.font.width(text, self.size)
        cs = (self.width(text) - natural) / len(text) if text else 0.0
        return TextOp(x, y, text, self.face, self.size, color, cs)


_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!"}

# built-in number formats whose appearance depends on the regional settings
LOCALE_BUILTIN_IDS = {5, 6, 7, 8, 37, 38, 39, 40}


def _localized_builtin(fid: int, loc) -> str | None:
    cur = '"%s"' % loc.currency if loc.currency else ""
    pos = loc.currency_pos

    def money(num):
        if pos == 0:
            return cur + num
        if pos == 1:
            return num + cur
        if pos == 2:
            return cur + " " + num
        return num + " " + cur
    table = {
        37: "#,##0;-#,##0", 38: "#,##0;[Red]-#,##0",
        39: "#,##0.00;-#,##0.00", 40: "#,##0.00;[Red]-#,##0.00",
        5: money("#,##0") + ";-" + money("#,##0"), 6: money("#,##0") + ";[Red]-" + money("#,##0"),
        7: money("#,##0.00") + ";-" + money("#,##0.00"),
        8: money("#,##0.00") + ";[Red]-" + money("#,##0.00"),
    }
    return table.get(fid)


def _parts_width(parts, m: DevFont):
    w = 0.0
    for kind, t in parts:
        if kind == "text":
            w += m.width(t)
        elif kind == "pad":
            w += m.char(t)
        elif kind == "space":
            w += t
    return w


def _transform(ops, x0, y0, s):
    """Scale ops by s around the origin and move them to (x0, y0)."""
    out = []
    for op in ops:
        out.append(_scale_op(op, x0, y0, s))
    return out


def _scale_op(op, x0, y0, s):
    t = type(op)
    if t is TextOp:
        return TextOp(x0 + op.x * s, y0 + op.y * s, op.text, op.face, op.size * s, op.color,
                      op.char_space * s, op.hscale, op.angle)
    if t is LineOp:
        return LineOp(x0 + op.x1 * s, y0 + op.y1 * s, x0 + op.x2 * s, y0 + op.y2 * s,
                      op.width * s, op.color, tuple(d * s for d in op.dash) if op.dash else None, op.cap)
    if t is RectOp:
        return RectOp(x0 + op.x * s, y0 + op.y * s, op.w * s, op.h * s, op.fill, op.stroke,
                      op.width * s, op.radius * s, op.ellipse)
    if t is ImageOp:
        return ImageOp(x0 + op.x * s, y0 + op.y * s, op.w * s, op.h * s, op.data, op.crop, op.angle)
    if t is ClipOp:
        return ClipOp(x0 + op.x * s, y0 + op.y * s, op.w * s, op.h * s,
                      [_scale_op(o, x0, y0, s) for o in op.ops])
    return op


def _hf_parts(hf, which, page_idx):
    first = bool(getattr(hf, "differentFirst", False))
    odd_even = bool(getattr(hf, "differentOddEven", False))
    if first and page_idx == 0:
        item = getattr(hf, "first" + which.capitalize(), None)
    elif odd_even and page_idx % 2 == 1:
        item = getattr(hf, "even" + which.capitalize(), None)
    else:
        item = getattr(hf, "odd" + which.capitalize(), None)
    if item is None:
        return None
    out = []
    for pos in ("left", "center", "right"):
        part = getattr(item, pos, None)
        if part is None or not part.text:
            out.append(None)
            continue
        prefix = ""
        if part.font:
            prefix += f'&"{part.font}"'
        if part.size:
            prefix += f"&{int(part.size)}"
        if part.color:
            prefix += f"&K{part.color}"
        out.append(prefix + part.text)
    return tuple(out)


