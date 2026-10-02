"""Office Math (OMML) layout.

Equations are laid out as nested boxes (width, ascent, descent, drawing ops
relative to the box's baseline origin) using the constants of the math font's
OpenType MATH table - the same data Word uses: script scaling and shifts,
fraction gaps, the math axis, and script glyph variants ('ssty').
"""
from __future__ import annotations

from dataclasses import dataclass, field

from ..draw import LineOp, TextOp, TransformOp, move_ops
from ..opc import M, W
from ..theme import hex_color

MATH_FONT = "Cambria Math"

_BINARY = set("+−±∓×÷·∗∘∙∩∪⊕⊗⊖⊘⊙⊚⊛⊞⊟⊠⊡∧∨∖∔∸⊓⊔⊎⋄⋅⋆⋇⋉⋊⋋⋌⋎⋏⋒⋓⨯⨿")
# relations get thick spacing; arrows (U+2190-21FF, U+27F0-27FF, U+2900-297F)
# are relations in Office Math too (e.g. U+27FC "long maps to")
_RELATION = (set("=<>≤≥≠≈≡∼≃∝∈∉∊∋∌∍⊂⊃⊆⊇⊄⊅⊈⊉⊊⊋≪≫≺≻≼≽≅≆≇≉≊≋≍≎≏≐≑≒≓≔≕≖≗≘≙≚≛≜≝≞≟"
                 "≡≢≣≤≥≦≧≨≩≮≯≰≱≲≳≴≵≶≷≸≹⊏⊐⊑⊒⊢⊣⊤⊥⊦⊧⊨⊩⊪⊫∣∤∥∦∷∺∻∽∾≀⋈⋍⋐⋑⋖⋗⋘⋙⋚⋛⋜⋝⋞⋟⋠⋡⋢⋣⋤⋥⋦⋧⋨⋩")
             | {chr(c) for c in range(0x2190, 0x2200)} | {chr(c) for c in range(0x27F0, 0x2800)}
             | {chr(c) for c in range(0x2900, 0x2980)})
_PUNCT = set(",;.")
_OPEN = set("([{⌈⌊⟨")

# defaults (em) used when the font has no MATH table
_DEFAULTS = {
    "ScriptPercentScaleDown": 73, "ScriptScriptPercentScaleDown": 60, "AxisHeight": 585,
    "SuperscriptShiftUp": 750, "SubscriptShiftDown": 418, "SpaceAfterScript": 85,
    "SuperscriptBaselineDropMax": 460, "SubscriptBaselineDropMin": 320, "SubscriptTopMax": 760,
    "SuperscriptBottomMin": 239, "SubSuperscriptGapMin": 300,
    "FractionNumeratorShiftUp": 1200, "FractionNumeratorDisplayStyleShiftUp": 1550,
    "FractionDenominatorShiftDown": 1030, "FractionDenominatorDisplayStyleShiftDown": 1370,
    "FractionNumeratorGapMin": 133, "FractionNumDisplayStyleGapMin": 260,
    "FractionDenominatorGapMin": 133, "FractionDenomDisplayStyleGapMin": 260,
    "FractionRuleThickness": 133, "MathLeading": 300, "RadicalVerticalGap": 166,
    "RadicalRuleThickness": 133, "UpperLimitGapMin": 133, "LowerLimitGapMin": 133,
    "OverbarVerticalGap": 345, "OverbarRuleThickness": 133, "DelimitedSubFormulaMinHeight": 3000,
}


@dataclass
class Box:
    w: float = 0.0
    asc: float = 0.0
    desc: float = 0.0
    ops: list = field(default_factory=list)
    last: str = ""  # last character, for operator spacing
    breaks: list = field(default_factory=list)  # (x, hang): where an inline equation may wrap
    italic_end: bool = False


def _hcat(boxes) -> Box:
    if len(boxes) == 1:
        return boxes[0]
    out = Box()
    for b in boxes:
        out.breaks.extend((out.w + x, hang) for x, hang in b.breaks)
        out.ops.extend(move_ops(b.ops, out.w, 0))
        out.w += b.w
        out.asc = max(out.asc, b.asc)
        out.desc = max(out.desc, b.desc)
        if b.last:
            out.last = b.last
    return out


_ALPHA = {"i": (0x1D44E, 0x1D434), "bi": (0x1D482, 0x1D468), "b": (0x1D41A, 0x1D400)}


def _ink(ops) -> tuple[float, float]:
    """Ink extent (above, below) the baseline of drawing ops."""
    asc = desc = 0.0
    for o in ops:
        if isinstance(o, TextOp):
            f = o.face.font
            for ch in o.text:
                lo, hi = f.glyph_bounds(ch)
                asc = max(asc, hi * o.size - o.y)
                desc = max(desc, o.y - lo * o.size)
        elif isinstance(o, LineOp):
            asc = max(asc, -min(o.y1, o.y2) + o.width / 2)
            desc = max(desc, max(o.y1, o.y2) + o.width / 2)
        elif isinstance(o, TransformOp):
            a_, d_ = _ink(o.ops)
            sy, fy = o.matrix[3], o.matrix[5]
            asc = max(asc, sy * a_ - fy)
            desc = max(desc, sy * d_ + fy)
    return asc, desc


# lowercase Greek (α..ω) per style; Word leaves capital Greek upright by default
_GREEK = {"i": 0x1D6FC, "bi": 0x1D736, "b": 0x1D6C2}


def math_italic(text: str, sty: str = "i") -> str:
    greek = _GREEK[sty]
    text = "".join(chr(greek + ord(c) - 0x3B1) if "α" <= c <= "ω" else c for c in text)
    if sty != "i":
        lo, up = _ALPHA[sty]
        return "".join(chr(lo + ord(c) - 97) if "a" <= c <= "z" else chr(up + ord(c) - 65) if "A" <= c <= "Z"
                       else c for c in text)
    out = []
    for ch in text:
        if "a" <= ch <= "z":
            out.append("ℎ" if ch == "h" else chr(0x1D44E + ord(ch) - ord("a")))
        elif "A" <= ch <= "Z":
            out.append(chr(0x1D434 + ord(ch) - ord("A")))
        else:
            out.append(ch)
    return "".join(out)


class MathLayout:
    def __init__(self, ctx, color=(0, 0, 0)):
        self.ctx = ctx
        self.color = color
        self.depth = 0  # nesting of structures; line breaks only at the top level
        self.face = ctx.face(MATH_FONT, False, False)
        self.font = self.face.font
        self.k = dict(_DEFAULTS)
        self._load_constants()

    def _load_constants(self):
        cache = getattr(self.font, "_math_constants", None)
        if cache is None:
            cache = {}
            try:
                from fontTools.ttLib import TTFont
                f = self.font
                t = TTFont(f.path, fontNumber=f.index, lazy=True) if f.path.lower().endswith(".ttc") \
                    else TTFont(f.path, lazy=True)
                if "MATH" in t:
                    mc = t["MATH"].table.MathConstants
                    for name in _DEFAULTS:
                        v = getattr(mc, name, None)
                        if v is not None:
                            cache[name] = getattr(v, "Value", v)
                t.close()
            except Exception:
                pass
            self.font._math_constants = cache
        self.k.update(cache)

    def em(self, name, size) -> float:
        v = self.k[name]
        if name.endswith("PercentScaleDown"):
            return v / 100.0
        return v / 2048.0 * size

    def script_size(self, size, level):
        if level <= 0:
            return size
        return size * (self.k["ScriptPercentScaleDown"] if level == 1 else self.k["ScriptScriptPercentScaleDown"]) / 100.0

    # --------------------------------------------------------------- text
    def text(self, s: str, size: float, level: int, prev: str = "") -> Box:
        f = self.font
        if level == 0:
            asc = f.ascent * size
            desc = f.descent * size
        else:
            # scripts are measured by their ink, like Word does
            asc = desc = 0.0
            for ch in s:
                lo, hi = f.glyph_bounds(ch)
                asc = max(asc, hi * size)
                desc = max(desc, -lo * size)
        box = Box(asc=asc, desc=desc)
        x = 0.0
        chars = list(s)
        for i, ch in enumerate(chars):
            before = after = 0.0
            nxt = chars[i + 1] if i + 1 < len(chars) else ""
            if level == 0:
                unary = prev == "" or prev in _OPEN or prev in _BINARY or prev in _RELATION
                if ch in _BINARY and not unary:
                    before = after = size * 4 / 18
                elif ch in _RELATION:
                    before = after = size * 5 / 18
                elif ch in _PUNCT and not (ch == "." and nxt.isdigit() and prev.isdigit()):
                    after = size * 3 / 18
            if self.depth == 0 and level == 0 and i and before and ch in _BINARY:
                box.breaks.append((x, 0.0))  # Word's default brkBin="before": "+" starts the new line
            x += before
            adv = f.char_width(ch, size)
            if level:
                alt = f.script_advance(ch, level)
                if alt is not None:
                    adv = alt * size
            natural = f.char_width(ch, size)
            if nxt and 0x1D400 <= ord(ch) <= 0x1D7FF and not (0x1D400 <= ord(nxt) <= 0x1D7FF):
                adv += f.italic_correction(ch) * size  # italic glyph before an upright one
            box.ops.append(TextOp(x, 0.0, ch, self.face, size, self.color, adv - natural))
            x += adv + after
            if self.depth == 0 and level == 0 and ch in _RELATION:
                # relations stay at the end of the line; their right spacing hangs
                box.breaks.append((x, after))
            prev = ch
        box.w = x
        box.last = prev
        box.simple = True
        return box

    # --------------------------------------------------------------- nodes
    def layout(self, el, size: float, level: int, display: bool) -> Box:
        box = self._seq(list(el), size, level, display)
        if box.last and 0x1D400 <= ord(box.last[-1]) <= 0x1D7FF and box.ops:
            # an equation ending in an italic letter keeps its italic correction
            # before the following (upright) text
            box.w += self.font.italic_correction(box.last[-1]) * size
        return box

    def _seq(self, children, size, level, display, prev="") -> Box:
        boxes = []
        for c in children:
            b = self._node(c, size, level, display, prev)
            if b is not None:
                boxes.append(b)
                prev = b.last or prev
        return _hcat(boxes)

    def _node(self, el, size, level, display, prev):
        from ..opc import local
        if local(el) == "r":
            return self._node_inner(el, size, level, display, prev)
        self.depth += 1
        try:
            return self._node_inner(el, size, level, display, prev)
        finally:
            self.depth -= 1

    def _node_inner(self, el, size, level, display, prev):
        from ..opc import local
        t = local(el)
        sz = self.script_size(size, level)
        if t == "r":
            text = "".join(x.text or "" for x in el.iter(M + "t"))
            if not text:
                return None
            mrpr = el.find(M + "rPr")
            sty_el = mrpr.find(M + "sty") if mrpr is not None else None
            sty = sty_el.get(M + "val") if sty_el is not None else "i"
            if not (mrpr is not None and mrpr.find(M + "nor") is not None) and sty in _ALPHA:
                text = math_italic(text, sty)
            text = text.replace("-", "−").replace("*", "∗").replace("'", "′")
            # invisible operators (function application, invisible times/separator/plus)
            text = "".join(c for c in text if not "⁡" <= c <= "⁤")
            if not text:
                return None
            col = el.find(W + "rPr/" + W + "color")
            saved = self.color
            if col is not None and col.get(W + "val") not in (None, "auto"):
                self.color = hex_color(col.get(W + "val")) or saved
            try:
                return self.text(text, sz, level, prev)
            finally:
                self.color = saved
        ch = {local(c): c for c in el}
        if t in ("sSub", "sSup", "sSubSup", "sPre"):
            return self._scripts(t, ch, size, level, display, prev)
        if t == "f":
            return self._frac(ch, size, level, display)
        if t == "d":
            return self._delim(el, ch, size, level, display)
        if t == "m":
            return self._matrix(el, size, level, display)
        if t == "eqArr":
            rows = [[e] for e in el if local(e) == "e"]
            return self._grid(rows, size, level, display, ["left"])
        if t == "func":
            name = self._seq(list(ch.get("fName", [])), size, level, display, prev)
            arg = self._seq(list(ch.get("e", [])), size, level, display)
            gap = Box(w=sz * 3 / 18)
            return _hcat([name, gap, arg])
        if t == "rad":
            return self._radical(ch, size, level, display)
        if t == "nary":
            return self._nary(ch, size, level, display)
        if t in ("acc", "bar", "groupChr", "limLow", "limUpp", "box", "borderBox", "phant"):
            base = self._seq(list(ch.get("e", [])), size, level, display, prev)
            if t in ("limLow", "limUpp"):
                lim = self._seq(list(ch.get("lim", [])), size, level + 1, display)
                return self._stack(base, lim, above=t == "limUpp", size=sz)
            if t == "bar":
                pos = ch.get("barPr")
                top = pos is None or pos.find(M + "pos") is None or pos.find(M + "pos").get(M + "val") != "bot"
                th = self.em("OverbarRuleThickness", sz)
                y = -(base.asc + self.em("OverbarVerticalGap", sz) / 2) if top else base.desc + th
                base.ops.append(LineOp(0, y, base.w, y, th, self.color))
                if top:
                    base.asc += self.em("OverbarVerticalGap", sz) / 2 + th
                else:
                    base.desc += th * 2
                return base
            if t == "acc":
                pr = ch.get("accPr")
                acc = "̂"
                if pr is not None and pr.find(M + "chr") is not None:
                    acc = pr.find(M + "chr").get(M + "val") or acc
                mark = self.text(acc, sz, level)
                mark.ops = move_ops(mark.ops, (base.w - mark.w) / 2, -base.asc * 0.1)
                base.ops.extend(mark.ops)
                base.asc += sz * 0.15
                return base
            if t == "borderBox":
                from ..draw import RectOp
                base.ops.append(RectOp(0, -base.asc, base.w, base.asc + base.desc, stroke=self.color, width=0.5))
            return base
        if t in ("rPr", "ctrlPr", "sSupPr", "sSubPr", "sSubSupPr", "fPr", "dPr", "radPr", "naryPr",
                 "oMathParaPr", "argPr", "mPr", "funcPr", "eqArrPr", "accPr", "barPr", "limLowPr",
                 "limUppPr", "boxPr", "borderBoxPr", "groupChrPr", "phantPr", "sPrePr"):
            return None
        return self._seq(list(el), size, level, display, prev)

    def _scripts(self, t, ch, size, level, display, prev):
        sz = self.script_size(size, level)
        base = self._seq(list(ch.get("e", [])), size, level, display, prev) if "e" in ch else Box()
        sub = self._seq(list(ch["sub"]), size, level + 1, display) if "sub" in ch else None
        sup = self._seq(list(ch["sup"]), size, level + 1, display) if "sup" in ch else None
        out = Box()
        parts = []
        if t == "sPre":
            parts.append(None)
        sup_shift = sub_shift = 0.0
        # baseline drops only apply to composite bases, not single characters
        simple = getattr(base, "simple", False)
        if sup is not None:
            sup_shift = max(self.em("SuperscriptShiftUp", sz), sup.desc + self.em("SuperscriptBottomMin", sz))
            if not simple:
                sup_shift = max(sup_shift, base.asc - self.em("SuperscriptBaselineDropMax", sz))
        if sub is not None:
            sub_shift = max(self.em("SubscriptShiftDown", sz), sub.asc - self.em("SubscriptTopMax", sz))
            if not simple:
                sub_shift = max(sub_shift, base.desc + self.em("SubscriptBaselineDropMin", sz))
        if sup is not None and sub is not None:
            gap = (sup_shift - sup.desc) - (sub.asc - sub_shift)
            need = self.em("SubSuperscriptGapMin", sz)
            if gap < need:
                sub_shift += need - gap
        script_w = max(sup.w if sup else 0.0, sub.w if sub else 0.0)
        x = 0.0
        if t == "sPre":
            if sup is not None:
                out.ops.extend(move_ops(sup.ops, script_w - sup.w, -sup_shift))
            if sub is not None:
                out.ops.extend(move_ops(sub.ops, script_w - sub.w, sub_shift))
            x = script_w
        out.ops.extend(move_ops(base.ops, x, 0))
        x += base.w
        if t != "sPre":
            if sup is not None:
                out.ops.extend(move_ops(sup.ops, x, -sup_shift))
            if sub is not None:
                out.ops.extend(move_ops(sub.ops, x, sub_shift))
            x += script_w + self.em("SpaceAfterScript", sz)
        out.w = x
        out.asc = max(base.asc, (sup_shift + sup.asc) if sup else 0.0)
        out.desc = max(base.desc, (sub_shift + sub.desc) if sub else 0.0)
        out.last = "x"
        return out

    def _frac(self, ch, size, level, display):
        sz = self.script_size(size, level)
        pr = ch.get("fPr")
        ftype = "bar"
        if pr is not None and pr.find(M + "type") is not None:
            ftype = pr.find(M + "type").get(M + "val") or "bar"
        inner = level if display and level == 0 else level + 1
        num = self._seq(list(ch.get("num", [])), size, inner, display)
        den = self._seq(list(ch.get("den", [])), size, inner, display)
        if ftype == "lin":
            slash = self.text("/", sz, level)
            return _hcat([num, slash, den])
        disp = display and level == 0
        axis = self.em("AxisHeight", sz)
        th = self.em("FractionRuleThickness", sz) if ftype != "noBar" else 0.0
        up = self.em("FractionNumeratorDisplayStyleShiftUp" if disp else "FractionNumeratorShiftUp", sz)
        down = self.em("FractionDenominatorDisplayStyleShiftDown" if disp else "FractionDenominatorShiftDown", sz)
        gnum = self.em("FractionNumDisplayStyleGapMin" if disp else "FractionNumeratorGapMin", sz)
        gden = self.em("FractionDenomDisplayStyleGapMin" if disp else "FractionDenominatorGapMin", sz)
        up = max(up, axis + th / 2 + gnum + num.desc)
        down = max(down, den.asc + gnum * 0 + gden - axis + th / 2)
        pad = sz * 0.1
        w = max(num.w, den.w) + 2 * pad
        out = Box(w=w, last="x")
        out.ops.extend(move_ops(num.ops, (w - num.w) / 2, -up))
        out.ops.extend(move_ops(den.ops, (w - den.w) / 2, down))
        if th:
            out.ops.append(LineOp(pad * 0.5, -axis, w - pad * 0.5, -axis, th, self.color))
        out.asc = up + num.asc
        out.desc = down + den.desc
        return out

    def _delimited(self, content: Box, beg: str, end: str, sz: float, level: int, grow=False) -> Box:
        axis = self.em("AxisHeight", sz)
        f = self.font
        ink_asc, ink_desc = _ink(content.ops)
        half = max(ink_asc - axis, ink_desc + axis)
        need = 2 * half
        out = Box(last=end or content.last)
        x = 0.0

        def glyph(ch):
            nonlocal x
            if not ch:
                return
            b = self.text(ch, sz, level)
            lo, hi = f.glyph_bounds(ch)
            normal = (hi - lo) * sz
            if grow and need > normal * 1.001:
                # content ink taller than the delimiter glyph: stretch it around the axis
                sy = need / normal
                # scale around the math axis
                out.ops.append(TransformOp((1.0, 0.0, 0.0, sy, x, axis * (sy - 1)), b.ops))
                out.asc = max(out.asc, axis + half)
                out.desc = max(out.desc, half - axis)
            else:
                out.ops.extend(move_ops(b.ops, x, 0))
                out.asc = max(out.asc, b.asc)
                out.desc = max(out.desc, b.desc)
            x += b.w
        glyph(beg)
        out.ops.extend(move_ops(content.ops, x, 0))
        x += content.w
        glyph(end)
        out.w = x
        out.asc = max(out.asc, content.asc)
        out.desc = max(out.desc, content.desc)
        return out

    def _delim(self, el, ch, size, level, display):
        from ..opc import local
        sz = self.script_size(size, level)
        beg, end, sep = "(", ")", "|"
        # Word stretches delimiters around multi-row content (arrays, matrices);
        # other content keeps text-size delimiters unless m:grow says otherwise
        grow = True
        pr = ch.get("dPr")
        if pr is not None:
            for c in pr:
                v = c.get(M + "val")
                if local(c) == "begChr":
                    beg = v or ""
                elif local(c) == "endChr":
                    end = v or ""
                elif local(c) == "sepChr":
                    sep = v or ""
                elif local(c) == "grow":
                    grow = v is None or v.lower() in ("1", "on", "true")
        parts = []
        for i, e in enumerate([c for c in el if local(c) == "e"]):
            if i:
                parts.append(self.text(sep, sz, level))
            parts.append(self._seq(list(e), size, level, display))
        return self._delimited(_hcat(parts), beg, end, sz, level, grow)

    def _grid(self, rows, size, level, display, col_jc):
        sz = self.script_size(size, level)
        cells = [[self._seq(list(c), size, level, display) for c in r] for r in rows]
        ncols = max((len(r) for r in cells), default=0)
        colw = [max((r[i].w for r in cells if i < len(r)), default=0.0) for i in range(ncols)]
        gap_x = sz * 0.5
        lead = self.em("MathLeading", sz)
        # row baselines
        ys = []
        y = 0.0
        prev_desc = None
        for r in cells:
            asc = max((c.asc for c in r), default=sz * 0.7)
            desc = max((c.desc for c in r), default=sz * 0.2)
            y = y + (prev_desc + lead + asc if prev_desc is not None else 0.0)
            ys.append((y, asc, desc))
            prev_desc = desc
        total_top = ys[0][1] if ys else 0.0
        total_bot = ys[-1][0] + ys[-1][2] if ys else 0.0
        height = total_top + total_bot
        axis = self.em("AxisHeight", sz)
        shift = -(height / 2) - axis + total_top  # baseline of row 0 relative to box baseline
        out = Box(last="x")
        for r, (ry, asc, desc) in zip(cells, ys):
            x = 0.0
            for i, c in enumerate(r):
                jc = col_jc[i] if i < len(col_jc) else "center"
                off = 0.0 if jc == "left" else (colw[i] - c.w if jc == "right" else (colw[i] - c.w) / 2)
                out.ops.extend(move_ops(c.ops, x + off, ry + shift))
                x += colw[i] + gap_x
        out.w = sum(colw) + gap_x * max(0, ncols - 1)
        out.asc = -shift + total_top
        out.desc = total_bot + shift
        return out

    def _matrix(self, el, size, level, display):
        from ..opc import local
        jcs = []
        pr = el.find(M + "mPr")
        if pr is not None:
            for mc in pr.iter(M + "mc"):
                mcpr = mc.find(M + "mcPr")
                jc = "center"
                count = 1
                if mcpr is not None:
                    if mcpr.find(M + "mcJc") is not None:
                        jc = mcpr.find(M + "mcJc").get(M + "val") or "center"
                    if mcpr.find(M + "count") is not None:
                        count = int(mcpr.find(M + "count").get(M + "val") or 1)
                jcs.extend([jc] * count)
        rows = [[c for c in mr if local(c) == "e"] for mr in el if local(mr) == "mr"]
        return self._grid(rows, size, level, display, jcs)

    def _stack(self, base: Box, other: Box, above: bool, size: float) -> Box:
        gap = self.em("UpperLimitGapMin" if above else "LowerLimitGapMin", size)
        w = max(base.w, other.w)
        out = Box(w=w, last=base.last)
        out.ops.extend(move_ops(base.ops, (w - base.w) / 2, 0))
        if above:
            dy = -(base.asc + gap + other.desc)
            out.asc = base.asc + gap + other.asc + other.desc
            out.desc = base.desc
        else:
            dy = base.desc + gap + other.asc
            out.asc = base.asc
            out.desc = base.desc + gap + other.asc + other.desc
        out.ops.extend(move_ops(other.ops, (w - other.w) / 2, dy))
        return out

    def _radical(self, ch, size, level, display):
        sz = self.script_size(size, level)
        body = self._seq(list(ch.get("e", [])), size, level, display)
        sign = self.text("√", sz, level)
        th = self.em("RadicalRuleThickness", sz)
        gap = self.em("RadicalVerticalGap", sz)
        top = body.asc + gap + th
        need = top + body.desc
        normal = sign.asc + sign.desc
        out = Box(last="x")
        if need > normal * 1.05:
            sy = need / normal
            out.ops.append(TransformOp((1.0, 0.0, 0.0, sy, 0.0, -(top - sign.asc * sy) + 0 * sy), sign.ops))
        else:
            out.ops.extend(sign.ops)
        out.ops.extend(move_ops(body.ops, sign.w, 0))
        out.ops.append(LineOp(sign.w, -top + th / 2, sign.w + body.w, -top + th / 2, th, self.color))
        out.w = sign.w + body.w
        out.asc = top
        out.desc = max(body.desc, sign.desc)
        return out

    def _nary(self, ch, size, level, display):
        sz = self.script_size(size, level)
        pr = ch.get("naryPr")
        op = "∫"
        lim_loc = None
        if pr is not None:
            if pr.find(M + "chr") is not None:
                op = pr.find(M + "chr").get(M + "val") or op
            if pr.find(M + "limLoc") is not None:
                lim_loc = pr.find(M + "limLoc").get(M + "val")
        big = sz * (1.4 if display and level == 0 else 1.0)
        opbox = self.text(op, big, level)
        axis = self.em("AxisHeight", sz)
        # centre the operator on the math axis
        mid = (opbox.asc - opbox.desc) / 2
        shift = mid - axis
        opbox.ops = move_ops(opbox.ops, 0, shift)
        opbox.asc -= shift
        opbox.desc += shift
        sub = self._seq(list(ch["sub"]), size, level + 1, display) if "sub" in ch and len(ch["sub"]) else None
        sup = self._seq(list(ch["sup"]), size, level + 1, display) if "sup" in ch and len(ch["sup"]) else None
        if lim_loc is None:
            lim_loc = "subSup" if op in ("∫", "∬", "∭", "∮") else "undOvr"
        if lim_loc == "undOvr":
            b = opbox
            if sup is not None:
                b = self._stack(b, sup, True, sz)
            if sub is not None:
                b = self._stack(b, sub, False, sz)
        else:
            parts = [opbox]
            b = _hcat(parts)
            w = max(sup.w if sup else 0, sub.w if sub else 0)
            if sup is not None:
                b.ops.extend(move_ops(sup.ops, opbox.w, -(opbox.asc - sup.asc * 0.5)))
                b.asc = max(b.asc, opbox.asc + sup.asc * 0.5)
            if sub is not None:
                b.ops.extend(move_ops(sub.ops, opbox.w, opbox.desc - sub.desc * 0.5))
                b.desc = max(b.desc, opbox.desc + sub.desc * 0.5)
            b.w += w
        body = self._seq(list(ch.get("e", [])), size, level, display)
        gap = Box(w=sz * 3 / 18)
        return _hcat([b, gap, body])
