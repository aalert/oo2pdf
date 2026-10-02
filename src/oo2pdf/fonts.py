"""Font discovery, metrics and substitution.

Office measures text with the font's own advance widths (no kerning by default)
and derives line height from the Windows metrics of the font (usWinAscent +
usWinDescent plus GDI "external leading").  This module exposes exactly those
numbers so the layout engines can reproduce Office's vertical rhythm.
"""
from __future__ import annotations

import hashlib
import json
import logging
import os
import sys
import threading
import unicodedata
from dataclasses import dataclass
from pathlib import Path

from fontTools.ttLib import TTCollection
from fontTools.ttLib import TTFont as _FTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase import ttfonts as _rl_ttfonts
from reportlab.pdfbase.ttfonts import TTFont as _RLFont

logging.getLogger("fontTools").setLevel(logging.ERROR)


def _utf16_hex(cp: int) -> str:
    if cp > 0xFFFF:
        cp -= 0x10000
        return "%04X%04X" % (0xD800 + (cp >> 10), 0xDC00 + (cp & 0x3FF))
    return "%04X" % cp


def _to_unicode_cmap(fontname, subset):
    """reportlab's ToUnicode CMap writes code points above U+FFFF as 5+ hex
    digits, which PDF readers truncate (math italic letters became garbage when
    copied or searched); ToUnicode values must be UTF-16BE, i.e. surrogate pairs."""
    return "\n".join([
        "/CIDInit /ProcSet findresource begin", "12 dict begin", "begincmap", "/CIDSystemInfo",
        "<< /Registry (%s)" % fontname, "/Ordering (%s)" % fontname, "/Supplement 0", ">> def",
        "/CMapName /%s def" % fontname, "/CMapType 2 def", "1 begincodespacerange",
        "<00> <%02X>" % (len(subset) - 1), "endcodespacerange", "%d beginbfchar" % len(subset),
    ] + ["<%02X> <%s>" % (i, _utf16_hex(v)) for i, v in enumerate(subset)] + [
        "endbfchar", "endcmap", "CMapName currentdict /CMap defineresource pop", "end", "end",
    ])


_rl_ttfonts.makeToUnicodeCMap = _to_unicode_cmap

_EXTS = (".ttf", ".ttc", ".otf")
ZERO_WIDTH = frozenset("\u200b\u2060\ufeff")
_INDEX_VERSION = 3

# Metric-compatible or visually closest substitutes, tried in order when the
# requested family is not installed.
SUBSTITUTES: dict[str, list[str]] = {
    "aptos": ["Aptos", "Calibri", "Carlito", "Arial", "Liberation Sans"],
    "aptos display": ["Aptos Display", "Calibri Light", "Calibri", "Carlito", "Arial"],
    "aptos narrow": ["Aptos Narrow", "Arial Narrow", "Liberation Sans Narrow", "Calibri"],
    "calibri": ["Carlito", "Arial", "Liberation Sans", "DejaVu Sans"],
    "calibri light": ["Calibri", "Carlito", "Arial"],
    "cambria": ["Caladea", "Times New Roman", "Liberation Serif", "DejaVu Serif"],
    "cambria math": ["Cambria", "Caladea", "Times New Roman"],
    "arial": ["Liberation Sans", "Arimo", "Helvetica", "DejaVu Sans"],
    "helvetica": ["Arial", "Liberation Sans"],
    "arial narrow": ["Liberation Sans Narrow", "Arial"],
    "times new roman": ["Liberation Serif", "Tinos", "Times", "DejaVu Serif"],
    "times": ["Times New Roman", "Liberation Serif"],
    "courier new": ["Liberation Mono", "Cousine", "Courier", "DejaVu Sans Mono"],
    "courier": ["Courier New", "Liberation Mono"],
    "consolas": ["Courier New", "Liberation Mono", "DejaVu Sans Mono"],
    "segoe ui": ["Selawik", "Arial", "Liberation Sans"],
    "tahoma": ["DejaVu Sans", "Verdana", "Arial"],
    "verdana": ["DejaVu Sans", "Arial"],
    "georgia": ["Gelasio", "Times New Roman"],
    "garamond": ["EB Garamond", "Times New Roman"],
    "century gothic": ["URW Gothic", "Arial"],
    "book antiqua": ["Palatino Linotype", "TeX Gyre Pagella", "Times New Roman"],
    "palatino linotype": ["TeX Gyre Pagella", "Times New Roman"],
    "ms mincho": ["MS Mincho", "Yu Mincho", "SimSun", "Noto Serif CJK JP"],
    "ms gothic": ["MS Gothic", "Yu Gothic", "Noto Sans CJK JP"],
    "宋体": ["SimSun", "NSimSun", "Noto Serif CJK SC"],
    "simsun": ["NSimSun", "Microsoft YaHei", "Noto Serif CJK SC"],
}

# Tried per character when the selected font has no glyph for it.
GLYPH_FALLBACKS = [
    "Segoe UI Symbol", "Segoe UI", "Arial", "Cambria Math", "Segoe UI Emoji",
    "Yu Gothic", "MS Gothic", "Microsoft YaHei", "SimSun", "Malgun Gothic",
    "Nirmala UI", "Leelawadee UI", "Ebrima", "Gadugi", "Segoe UI Historic",
    "Arial Unicode MS", "DejaVu Sans", "Noto Sans", "Noto Sans CJK SC",
    "Noto Sans Symbols", "Noto Sans Symbols2", "Liberation Sans",
]

LAST_RESORT = ["Calibri", "Arial", "Liberation Sans", "DejaVu Sans", "Segoe UI", "Times New Roman"]

# Symbol-encoded fonts (Word stores bullets as U+F0xx in these).
_SYMBOL_FONTS = {"symbol", "wingdings", "wingdings 2", "wingdings 3", "webdings"}

_SYMBOL_TO_UNICODE = {
    0x22: "∀", 0x24: "∃", 0x27: "∋", 0x2A: "∗", 0x2D: "−", 0x40: "≅", 0x5C: "∴", 0x5E: "⊥",
    0x60: "‾", 0x7E: "∼", 0xA1: "ϒ", 0xA2: "′", 0xA3: "≤", 0xA4: "⁄", 0xA5: "∞",
    0xA6: "ƒ", 0xA7: "♣", 0xA8: "♦", 0xA9: "♥", 0xAA: "♠", 0xAB: "↔", 0xAC: "←",
    0xAD: "↑", 0xAE: "→", 0xAF: "↓", 0xB0: "°", 0xB1: "±", 0xB2: "″", 0xB3: "≥",
    0xB4: "×", 0xB5: "∝", 0xB6: "∂", 0xB7: "•", 0xB8: "÷", 0xB9: "≠", 0xBA: "≡",
    0xBB: "≈", 0xBC: "…", 0xC0: "ℵ", 0xC1: "ℑ", 0xC2: "ℜ", 0xC3: "℘", 0xC4: "⊗",
    0xC5: "⊕", 0xC6: "∅", 0xC7: "∩", 0xC8: "∪", 0xC9: "⊃", 0xCA: "⊇", 0xCB: "⊄",
    0xCC: "⊂", 0xCD: "⊆", 0xCE: "∈", 0xCF: "∉", 0xD0: "∠", 0xD1: "∇", 0xD2: "®",
    0xD3: "©", 0xD4: "™", 0xD5: "∏", 0xD6: "√", 0xD7: "⋅", 0xD8: "¬", 0xD9: "∧",
    0xDA: "∨", 0xDB: "⇔", 0xDC: "⇐", 0xDD: "⇑", 0xDE: "⇒", 0xDF: "⇓", 0xE0: "◊",
    0xE1: "〈", 0xE5: "∑", 0xF1: "〉", 0xF2: "∫",
}
for _i, _c in enumerate("ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ"):
    _SYMBOL_TO_UNICODE[0x41 + _i] = _c
for _i, _c in enumerate("αβχδεφγηιϕκλμνοπθρστυϖωξψζ"):
    _SYMBOL_TO_UNICODE[0x61 + _i] = _c

_WINGDINGS_TO_UNICODE = {
    0x4A: "☺", 0x4B: "😐", 0x4C: "☹", 0x6C: "●", 0x6D: "❍", 0x6E: "■", 0x6F: "□",
    0x70: "◻", 0x71: "❑", 0x72: "❒", 0x73: "⬧", 0x74: "⧫", 0x75: "◆", 0x76: "❖",
    0x77: "⬥", 0x9F: "•", 0xA1: "○", 0xA2: "⭕", 0xA7: "▪", 0xA8: "◻", 0xD8: "➢",
    0xE0: "→", 0xE8: "➔", 0xEF: "⇦", 0xF0: "⇨", 0xFB: "✗", 0xFC: "✓", 0xFD: "☒",
    0xFE: "☑", 0x71 + 0x100: "❑",
}


def default_font_dirs() -> list[str]:
    dirs: list[str] = []
    env = os.environ.get("OO2PDF_FONT_DIRS")
    if env:
        dirs.extend(p for p in env.split(os.pathsep) if p)
    if sys.platform == "win32":
        windir = os.environ.get("WINDIR", r"C:\Windows")
        dirs.append(os.path.join(windir, "Fonts"))
        la = os.environ.get("LOCALAPPDATA")
        if la:
            dirs.append(os.path.join(la, "Microsoft", "Windows", "Fonts"))
            dirs.append(os.path.join(la, "Microsoft", "FontCache", "4", "CloudFonts"))
    elif sys.platform == "darwin":
        dirs += ["/System/Library/Fonts", "/Library/Fonts",
                 os.path.expanduser("~/Library/Fonts"),
                 "/Applications/Microsoft Word.app/Contents/Resources/DFonts"]
    else:
        dirs += ["/usr/share/fonts", "/usr/local/share/fonts",
                 os.path.expanduser("~/.fonts"), os.path.expanduser("~/.local/share/fonts")]
    return dirs


def _cache_file() -> Path:
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_CACHE_HOME") \
        or os.path.join(Path.home(), ".cache")
    return Path(base) / "oo2pdf" / f"fontindex-v{_INDEX_VERSION}.json"


def _scan_file(path: str) -> list[dict]:
    out = []
    try:
        if path.lower().endswith(".ttc"):
            fonts = list(TTCollection(path, lazy=True).fonts)
        else:
            fonts = [_FTFont(path, lazy=True)]
    except Exception:
        return out
    for idx, f in enumerate(fonts):
        try:
            if "glyf" not in f:  # CFF outlines cannot be embedded by reportlab
                continue
            name = f["name"]
            fams = set()
            primary = set()
            for rec in name.names:
                if rec.nameID == 1:
                    try:
                        primary.add(rec.toUnicode().strip().lower())
                    except Exception:
                        pass
                if rec.nameID in (1, 16, 21):
                    try:
                        s = rec.toUnicode().strip()
                    except Exception:
                        continue
                    if s:
                        fams.add(s)
            os2 = f["OS/2"] if "OS/2" in f else None
            weight = int(os2.usWeightClass) if os2 else 400
            fs = int(os2.fsSelection) if os2 else 0
            mac = int(f["head"].macStyle)
            italic = bool(fs & 1) or bool(mac & 2)
            if weight < 100:  # some old fonts use 1..9
                weight *= 100
            if (fs & 32 or mac & 1) and weight < 700:
                weight = 700
            out.append({"path": path, "index": idx, "families": sorted(fams),
                        "primary": sorted(primary), "weight": weight, "italic": italic,
                        "width": int(os2.usWidthClass) if os2 else 5})
        except Exception:
            continue
    return out


class FontIndex:
    def __init__(self, dirs: list[str] | None = None):
        self.dirs = dirs or default_font_dirs()
        self.entries: list[dict] = []
        self.by_family: dict[str, list[dict]] = {}
        self._build()

    def _build(self):
        cache_path = _cache_file()
        cache: dict = {}
        try:
            cache = json.loads(cache_path.read_text("utf-8"))
        except Exception:
            cache = {}
        files: dict = cache.get("files", {})
        new_files: dict = {}
        changed = False
        for d in self.dirs:
            if not os.path.isdir(d):
                continue
            for root, _dirs, names in os.walk(d):
                for n in names:
                    if not n.lower().endswith(_EXTS):
                        continue
                    p = os.path.join(root, n)
                    try:
                        st = os.stat(p)
                    except OSError:
                        continue
                    key = f"{st.st_mtime_ns}:{st.st_size}"
                    old = files.get(p)
                    if old and old.get("key") == key:
                        new_files[p] = old
                    else:
                        new_files[p] = {"key": key, "faces": _scan_file(p)}
                        changed = True
        if changed or set(new_files) != set(files):
            try:
                cache_path.parent.mkdir(parents=True, exist_ok=True)
                cache_path.write_text(json.dumps({"files": new_files}), "utf-8")
            except Exception:
                pass
        for p in sorted(new_files):
            for face in new_files[p]["faces"]:
                self.entries.append(face)
                for fam in face["families"]:
                    self.by_family.setdefault(fam.lower(), []).append(face)

    def match(self, family: str, bold: bool, italic: bool):
        cands = self.by_family.get(family.lower())
        if not cands:
            return None
        target = 700 if bold else 400
        fam = family.lower()

        def score(e):
            s = 0
            # a face whose legacy family name is the request (e.g. "Arial") beats
            # one only grouped under it by typographic family ("Arial Narrow")
            if fam not in e.get("primary", ()):
                s += 300
            s += abs(e.get("width", 5) - 5) * 400
            if e["italic"] != italic:
                s += 1000
            w = e["weight"]
            d = abs(w - target)
            # CSS-like preference: for regular prefer 400/500 before lighter;
            # for bold prefer heavier before lighter.
            if bold and w < target:
                d += 150
            if not bold and w < 400:
                d += 50
            return s + d
        best = min(cands, key=score)
        return best


@dataclass(frozen=True)
class FontFace:
    font: "Font"
    synth_bold: bool = False
    synth_italic: bool = False


class Font:
    """A loaded TrueType font with Office-relevant metrics."""

    def __init__(self, path: str, index: int = 0):
        self.path = path
        self.index = index
        h = hashlib.md5(f"{path}#{index}".encode()).hexdigest()[:10]
        self.name = f"F{h}"
        rl = _RLFont(self.name, path, subfontIndex=index)
        pdfmetrics.registerFont(rl)
        face = rl.face
        self._widths = face.charWidths  # code point -> width in 1/1000 em
        self._default_width = face.defaultWidth
        self._cmap = face.charToGlyph

        f = _FTFont(path, fontNumber=index, lazy=True) if path.lower().endswith(".ttc") \
            else _FTFont(path, lazy=True)
        upm = f["head"].unitsPerEm
        hhea = f["hhea"]
        os2 = f["OS/2"] if "OS/2" in f else None
        post = f["post"] if "post" in f else None
        self.family = f["name"].getDebugName(1) or ""
        if os2 is not None and (os2.usWinAscent or os2.usWinDescent):
            win_asc, win_desc = os2.usWinAscent, os2.usWinDescent
        else:
            win_asc, win_desc = hhea.ascent, -hhea.descent
        # GDI tmExternalLeading
        ext = max(0, hhea.lineGap - ((win_asc + win_desc) - (hhea.ascent - hhea.descent)))
        self.ascent = win_asc / upm
        self.descent = win_desc / upm
        self.ext_leading = ext / upm
        self.hhea_ascent = hhea.ascent / upm
        self.hhea_descent = -hhea.descent / upm
        self.hhea_gap = hhea.lineGap / upm
        if os2 is not None:
            self.typo_ascent = os2.sTypoAscender / upm
            self.typo_descent = -os2.sTypoDescender / upm
            self.typo_gap = os2.sTypoLineGap / upm
            self.strike_pos = (os2.yStrikeoutPosition or upm * 0.25) / upm
            self.strike_size = (os2.yStrikeoutSize or upm * 0.05) / upm
            self.x_height = (getattr(os2, "sxHeight", 0) or upm * 0.5) / upm
            self.cap_height = (getattr(os2, "sCapHeight", 0) or upm * 0.7) / upm
            self.use_typo = bool(os2.fsSelection & 128)
        else:
            self.typo_ascent, self.typo_descent, self.typo_gap = self.hhea_ascent, self.hhea_descent, self.hhea_gap
            self.strike_pos, self.strike_size = 0.25, 0.05
            self.x_height, self.cap_height = 0.5, 0.7
            self.use_typo = False
        # GDI (Excel) always uses the Windows metrics; Word honours the
        # USE_TYPO_METRICS flag and then uses the typographic metrics.
        self.win_ascent = self.ascent
        self.win_descent = self.descent
        self.win_ext_leading = self.ext_leading
        self._gdi = None
        if self.use_typo:
            self.ascent = self.typo_ascent
            self.descent = self.typo_descent
            self.ext_leading = max(0.0, self.typo_gap)
        if post is not None and post.underlineThickness:
            self.underline_pos = post.underlinePosition / upm
            self.underline_size = post.underlineThickness / upm
        else:
            self.underline_pos, self.underline_size = -0.1, 0.05
        self.is_symbol = self.family.lower() in _SYMBOL_FONTS
        self.upm = upm
        self._kern_tables = None
        self._kern_cache: dict[tuple[str, str], float] = {}
        self._glyph_of: dict[int, str] | None = None
        f.close()

    # --- GDI metrics (used to mimic Excel) -------------------------------------
    def _gdi_tables(self):
        if getattr(self, "_gdi", None) is None:
            vdmx, hdmx, glyph_of = {}, {}, {}
            try:
                f = _FTFont(self.path, fontNumber=self.index, lazy=True) if self.path.lower().endswith(".ttc") \
                    else _FTFont(self.path, lazy=True)
                if "VDMX" in f:
                    v = f["VDMX"]
                    group = None
                    for r in v.ratRanges:  # prefer the 1:1 aspect ratio group
                        if r["xRatio"] == r["yStartRatio"] == r["yEndRatio"] or r["xRatio"] == 0:
                            group = r["groupIndex"]
                            break
                    if group is not None and group < len(v.groups):
                        vdmx = dict(v.groups[group])
                if "hdmx" in f:
                    hdmx = f["hdmx"].hdmx
                    glyph_of = f.getBestCmap() or {}
                f.close()
            except Exception:
                pass
            self._gdi = (vdmx, hdmx, glyph_of)
        return self._gdi

    def gdi_height(self, ppem: int) -> tuple[int, int, int]:
        """(tmAscent, tmDescent, tmExternalLeading) in pixels at ``ppem``."""
        vdmx, _, _ = self._gdi_tables()
        if ppem in vdmx:
            ymax, ymin = vdmx[ppem]
            asc, desc = int(ymax), int(-ymin)
        else:
            asc = int(round(self.win_ascent * ppem))
            desc = int(round(self.win_descent * ppem))
        return asc, desc, int(round(self.win_ext_leading * ppem))

    def gdi_advance(self, ch: str, ppem: int) -> int | None:
        """Hinted advance width in pixels from the hdmx table, when present."""
        _, hdmx, glyph_of = self._gdi_tables()
        if not hdmx or ppem not in hdmx:
            return None
        g = glyph_of.get(ord(ch))
        if g is None:
            return None
        return hdmx[ppem].get(g)

    # --- kerning -------------------------------------------------------------
    def _load_kerning(self):
        """Collect pair-positioning subtables of the GPOS 'kern' feature (or the
        legacy 'kern' table) for pair lookups."""
        tables = []
        try:
            f = _FTFont(self.path, fontNumber=self.index, lazy=True) if self.path.lower().endswith(".ttc") \
                else _FTFont(self.path, lazy=True)
            self._glyph_of = f.getBestCmap() or {}
            if "GPOS" in f:
                gpos = f["GPOS"].table
                lookups = set()
                if gpos.FeatureList:
                    for fr in gpos.FeatureList.FeatureRecord:
                        if fr.FeatureTag == "kern":
                            lookups.update(fr.Feature.LookupListIndex)
                for li in sorted(lookups):
                    lk = gpos.LookupList.Lookup[li]
                    for st in lk.SubTable:
                        kind = lk.LookupType
                        if kind == 9:  # extension: the real type is inside
                            kind = st.ExtensionLookupType
                            st = st.ExtSubTable
                        if kind != 2 or not hasattr(st, "Format"):
                            continue  # only pair positioning carries kerning
                        cov = set(st.Coverage.glyphs)
                        if st.Format == 1:
                            pairs = {}
                            for g1, ps in zip(st.Coverage.glyphs, st.PairSet):
                                d = {}
                                for rec in ps.PairValueRecord:
                                    v = getattr(rec.Value1, "XAdvance", 0) if rec.Value1 else 0
                                    if v:
                                        d[rec.SecondGlyph] = v
                                if d:
                                    pairs[g1] = d
                            tables.append(("p", pairs))
                        elif st.Format == 2:
                            c1 = st.ClassDef1.classDefs if st.ClassDef1 else {}
                            c2 = st.ClassDef2.classDefs if st.ClassDef2 else {}
                            matrix = [[getattr(r.Value1, "XAdvance", 0) if r.Value1 else 0
                                       for r in rec.Class2Record] for rec in st.Class1Record]
                            tables.append(("c", cov, c1, c2, matrix))
            elif "kern" in f:
                pairs = {}
                for sub in f["kern"].kernTables:
                    for (a, b), v in getattr(sub, "kernTable", {}).items():
                        pairs.setdefault(a, {})[b] = v
                tables.append(("p", pairs))
            f.close()
        except Exception:
            tables = []
        self._kern_tables = tables

    def kern(self, a: str, b: str) -> float:
        """Pair kerning adjustment in em units (negative tightens)."""
        key = (a, b)
        v = self._kern_cache.get(key)
        if v is not None:
            return v
        if self._kern_tables is None:
            self._load_kerning()
        v = 0.0
        if self._kern_tables:
            g1 = self._glyph_of.get(ord(a))
            g2 = self._glyph_of.get(ord(b))
            if g1 and g2:
                for t in self._kern_tables:
                    if t[0] == "p":
                        d = t[1].get(g1)
                        if d is not None and g2 in d:
                            v = d[g2] / self.upm
                            break
                    else:
                        _, cov, c1, c2, matrix = t
                        if g1 in cov:
                            k1 = c1.get(g1, 0)
                            k2 = c2.get(g2, 0)
                            if k1 < len(matrix) and k2 < len(matrix[k1]):
                                x = matrix[k1][k2]
                                if x:
                                    v = x / self.upm
                                    break
                            # a pair in a class subtable stops the search only when non-zero
        self._kern_cache[key] = v
        return v

    # --- metrics -----------------------------------------------------------

    def has(self, ch: str) -> bool:
        return ord(ch) in self._cmap

    def char_width(self, ch: str, size: float) -> float:
        return self._widths.get(ord(ch), self._default_width) * size / 1000.0

    def width(self, text: str, size: float) -> float:
        w = self._widths
        dw = self._default_width
        total = 0
        for ch in text:
            if ch in ZERO_WIDTH:
                continue
            total += w.get(ord(ch), dw)
        return total * size / 1000.0

    def glyph_bounds(self, ch: str) -> tuple[float, float]:
        """(yMin, yMax) of a glyph's outline in em (ink extent)."""
        cache = getattr(self, "_bounds", None)
        if cache is None:
            cache = self._bounds = {}
        b = cache.get(ch)
        if b is None:
            b = (-self.descent, self.ascent)
            try:
                if getattr(self, "_glyf_font", None) is None:
                    self._glyf_font = _FTFont(self.path, fontNumber=self.index, lazy=True)                         if self.path.lower().endswith(".ttc") else _FTFont(self.path, lazy=True)
                    self._glyf_cmap = self._glyf_font.getBestCmap() or {}
                g = self._glyf_cmap.get(ord(ch))
                if g is not None:
                    glyf = self._glyf_font["glyf"]
                    gl = glyf[g]
                    if gl.numberOfContours:
                        gl.recalcBounds(glyf)
                        upm = self.upm
                        b = (gl.yMin / upm, gl.yMax / upm)
                    else:
                        b = (0.0, 0.0)
            except Exception:
                pass
            cache[ch] = b
        return b

    def italic_correction(self, ch: str) -> float:
        """MATH table italics correction (em) of a glyph, 0 when none."""
        table = getattr(self, "_italic_corr", None)
        if table is None:
            table = {}
            try:
                f = _FTFont(self.path, fontNumber=self.index, lazy=True) if self.path.lower().endswith(".ttc")                     else _FTFont(self.path, lazy=True)
                if "MATH" in f:
                    info = f["MATH"].table.MathGlyphInfo.MathItalicsCorrectionInfo
                    if info is not None:
                        names = info.Coverage.glyphs
                        by_name = {n: v.Value for n, v in zip(names, info.ItalicsCorrection)}
                        upm = f["head"].unitsPerEm
                        for cp, g in (f.getBestCmap() or {}).items():
                            if g in by_name:
                                table[cp] = by_name[g] / upm
                f.close()
            except Exception:
                table = {}
            self._italic_corr = table
        return table.get(ord(ch), 0.0)

    def script_advance(self, ch: str, level: int) -> float | None:
        """Advance (em) of the glyph a math font uses at script level 1/2
        (OpenType 'ssty' substitution), or None when there is no variant."""
        if getattr(self, "_ssty", None) is None:
            self._ssty = {}
            try:
                f = _FTFont(self.path, fontNumber=self.index, lazy=True) if self.path.lower().endswith(".ttc") \
                    else _FTFont(self.path, lazy=True)
                if "GSUB" in f and "MATH" in f:
                    gsub = f["GSUB"].table
                    cmap = f.getBestCmap() or {}
                    hmtx = f["hmtx"]
                    upm = f["head"].unitsPerEm
                    for fr in gsub.FeatureList.FeatureRecord:
                        if fr.FeatureTag != "ssty":
                            continue
                        for li in fr.Feature.LookupListIndex:
                            lk = gsub.LookupList.Lookup[li]
                            for st in lk.SubTable:
                                kind = lk.LookupType
                                if kind == 7:
                                    kind, st = st.ExtensionLookupType, st.ExtSubTable
                                if kind == 1:  # single substitution: level 1
                                    for cp, g in cmap.items():
                                        if g in st.mapping:
                                            self._ssty.setdefault((cp, 1), hmtx[st.mapping[g]][0] / upm)
                                elif kind == 3:  # alternates: [level 1, level 2]
                                    for cp, g in cmap.items():
                                        alts = st.alternates.get(g)
                                        if alts:
                                            for lv, alt in enumerate(alts[:2], 1):
                                                self._ssty.setdefault((cp, lv), hmtx[alt][0] / upm)
                        break
                f.close()
            except Exception:
                self._ssty = {}
        return self._ssty.get((ord(ch), min(level, 2)))


class FontManager:
    """Resolves (family, bold, italic) requests to concrete fonts, with caching."""

    _shared_index: FontIndex | None = None
    _lock = threading.Lock()

    def __init__(self, font_dirs: list[str] | None = None, substitutions: dict[str, str] | None = None):
        if font_dirs:
            self.index = FontIndex(font_dirs + default_font_dirs())
        else:
            with FontManager._lock:
                if FontManager._shared_index is None:
                    FontManager._shared_index = FontIndex()
            self.index = FontManager._shared_index
        self.user_subs = {k.lower(): v for k, v in (substitutions or {}).items()}
        self._fonts: dict[tuple[str, int], Font | None] = {}
        self._faces: dict[tuple[str, bool, bool], FontFace] = {}
        self._fallback_cache: dict[tuple[str, bool, bool], FontFace | None] = {}

    def _load(self, entry) -> Font | None:
        key = (entry["path"], entry["index"])
        if key not in self._fonts:
            try:
                self._fonts[key] = Font(entry["path"], entry["index"])
            except Exception:
                self._fonts[key] = None
        return self._fonts[key]

    def _try_family(self, family: str, bold: bool, italic: bool) -> FontFace | None:
        entry = self.index.match(family, bold, italic)
        if entry is None:
            return None
        font = self._load(entry)
        if font is None:
            # try other faces of the same family (e.g. a non-embeddable bold)
            for e in self.index.by_family.get(family.lower(), []):
                font = self._load(e)
                if font is not None:
                    entry = e
                    break
            if font is None:
                return None
        return FontFace(font, synth_bold=bold and entry["weight"] < 600,
                        synth_italic=italic and not entry["italic"])

    def face(self, family: str | None, bold: bool = False, italic: bool = False) -> FontFace:
        family = (family or "Calibri").strip()
        key = (family.lower(), bold, italic)
        f = self._faces.get(key)
        if f is not None:
            return f
        chain = [family]
        if family.lower() in self.user_subs:
            chain.insert(0, self.user_subs[family.lower()])
        chain += SUBSTITUTES.get(family.lower(), [])
        chain += LAST_RESORT
        for fam in chain:
            f = self._try_family(fam, bold, italic)
            if f is not None:
                break
        if f is None:
            if not self.index.entries:
                raise RuntimeError("No usable TrueType fonts found; set OO2PDF_FONT_DIRS")
            for e in self.index.entries:
                font = self._load(e)
                if font:
                    f = FontFace(font, bold, italic)
                    break
        self._faces[key] = f
        return f

    def fallback_for(self, ch: str, bold: bool, italic: bool) -> FontFace | None:
        for fam in GLYPH_FALLBACKS:
            key = (fam, bold, italic)
            if key not in self._fallback_cache:
                self._fallback_cache[key] = self._try_family(fam, bold, italic)
            f = self._fallback_cache[key]
            if f is not None and f.font.has(ch):
                return f
        return None

    def segments(self, text: str, face: FontFace, bold: bool, italic: bool):
        """Split text into runs that the given face (or a fallback) can render.

        Yields (text, FontFace) tuples.  Symbol-encoded fonts get Word's U+F0xx
        mapping; characters missing from the font are rendered with a fallback.
        """
        font = face.font
        out: list[tuple[str, FontFace]] = []
        cur_text: list[str] = []
        cur_face = face
        for ch in text:
            f = face
            c = ch
            if font.is_symbol:
                o = ord(ch)
                if 0x20 <= o <= 0xFF and not font.has(ch) and font.has(chr(0xF000 + o)):
                    c = chr(0xF000 + o)
                elif not font.has(ch):
                    mapped = map_symbol_char(font.family, ch)
                    if mapped:
                        c = mapped
                        f = self.fallback_for(mapped, bold, italic) or face
            elif not font.has(ch) and not _is_invisible(ch):
                o = ord(ch)
                if 0xF020 <= o <= 0xF0FF:  # symbol-area char in a text font
                    mapped = _SYMBOL_TO_UNICODE.get(o - 0xF000) or chr(o - 0xF000)
                    c = mapped
                    if not font.has(c):
                        f = self.fallback_for(c, bold, italic) or face
                else:
                    f = self.fallback_for(ch, bold, italic) or face
            if f is not cur_face and cur_text:
                out.append(("".join(cur_text), cur_face))
                cur_text = []
            cur_face = f
            cur_text.append(c)
        if cur_text:
            out.append(("".join(cur_text), cur_face))
        return out


def _is_invisible(ch: str) -> bool:
    return ch in "\u200b\u200c\u200d\ufeff\u00ad" or unicodedata.category(ch) in ("Cc", "Cf")


def map_symbol_char(family: str, ch: str) -> str | None:
    o = ord(ch)
    if o >= 0xF000:
        o -= 0xF000
    fam = family.lower()
    if fam == "symbol":
        return _SYMBOL_TO_UNICODE.get(o) or (chr(o) if 0x20 <= o < 0x7F else None)
    if fam.startswith("wingdings") or fam == "webdings":
        return _WINGDINGS_TO_UNICODE.get(o, "•")
    return None


