"""List numbering (numbering.xml) and label generation."""
from __future__ import annotations

from dataclasses import dataclass, field

from ..opc import R, V, W
from .props import attr, num, parse_ppr, parse_rpr


@dataclass
class Level:
    start: int = 1
    fmt: str = "decimal"
    text: str = "%1."
    jc: str = "left"
    suff: str = "tab"
    ppr: dict = field(default_factory=dict)
    rpr: dict = field(default_factory=dict)
    restart: int | None = None
    is_lgl: bool = False
    pic_id: str | None = None  # lvlPicBulletId: a picture instead of the bullet text


def _parse_level(el) -> Level:
    lv = Level()
    s = el.find(W + "start")
    if s is not None:
        lv.start = int(num(attr(s, "val"), 1))
    f = el.find(W + "numFmt")
    if f is not None:
        lv.fmt = attr(f, "val", "decimal")
    t = el.find(W + "lvlText")
    if t is not None:
        lv.text = attr(t, "val", "")
    j = el.find(W + "lvlJc")
    if j is not None:
        lv.jc = attr(j, "val", "left")
    sf = el.find(W + "suff")
    if sf is not None:
        lv.suff = attr(sf, "val", "tab")
    r = el.find(W + "lvlRestart")
    if r is not None:
        lv.restart = int(num(attr(r, "val"), 0))
    lv.is_lgl = el.find(W + "isLgl") is not None
    pb = el.find(W + "lvlPicBulletId")
    if pb is not None:
        lv.pic_id = attr(pb, "val")
    lv.ppr = parse_ppr(el.find(W + "pPr"))
    lv.rpr = parse_rpr(el.find(W + "rPr"))
    return lv


_ROMAN = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
          (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
_ONES = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten",
         "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen",
         "eighteen", "nineteen"]
_TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]
_ORD = {"one": "first", "two": "second", "three": "third", "five": "fifth", "eight": "eighth",
        "nine": "ninth", "twelve": "twelfth"}


def roman(n: int) -> str:
    if n <= 0:
        return str(n)
    out = []
    for v, s in _ROMAN:
        while n >= v:
            out.append(s)
            n -= v
    return "".join(out)


def cardinal(n: int) -> str:
    if n < 20:
        return _ONES[n] if n >= 0 else str(n)
    if n < 100:
        t, o = divmod(n, 10)
        return _TENS[t] + ("-" + _ONES[o] if o else "")
    if n < 1000:
        h, r = divmod(n, 100)
        return _ONES[h] + " hundred" + (" " + cardinal(r) if r else "")
    th, r = divmod(n, 1000)
    return cardinal(th) + " thousand" + (" " + cardinal(r) if r else "")


def ordinal_text(n: int) -> str:
    words = cardinal(n).split("-")
    last = words[-1].split(" ")
    w = last[-1]
    if w in _ORD:
        w = _ORD[w]
    elif w.endswith("y"):
        w = w[:-1] + "ieth"
    else:
        w += "th"
    last[-1] = w
    words[-1] = " ".join(last)
    return "-".join(words)


def format_number(n: int, fmt: str) -> str:
    if fmt in ("decimal", "decimalHalfWidth", "decimalFullWidth"):
        return str(n)
    if fmt == "decimalZero":
        return f"{n:02d}"
    if fmt == "lowerRoman":
        return roman(n)
    if fmt == "upperRoman":
        return roman(n).upper()
    if fmt in ("lowerLetter", "upperLetter"):
        if n <= 0:
            return str(n)
        ch = chr(ord("a") + (n - 1) % 26) * ((n - 1) // 26 + 1)
        return ch.upper() if fmt == "upperLetter" else ch
    if fmt == "ordinal":
        suf = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
        return f"{n}{suf}"
    if fmt == "cardinalText":
        s = cardinal(n)
        return s[:1].upper() + s[1:]
    if fmt == "ordinalText":
        s = ordinal_text(n)
        return s[:1].upper() + s[1:]
    if fmt == "none":
        return ""
    if fmt == "decimalEnclosedCircle" and 1 <= n <= 20:
        return chr(0x2460 + n - 1)
    if fmt == "decimalEnclosedParen" and 1 <= n <= 20:
        return chr(0x2474 + n - 1)
    if fmt == "chicago":
        syms = "*†‡§"
        return syms[(n - 1) % 4] * ((n - 1) // 4 + 1)
    return str(n)


class Numbering:
    def __init__(self, root, styles):
        self.abstract: dict[str, dict[int, Level]] = {}
        self.style_link: dict[str, str] = {}
        self.nums: dict[str, tuple[str, dict]] = {}
        self.styles = styles
        self.counters: dict[str, list[int | None]] = {}
        self.seen_nums: set[str] = set()
        # numPicBulletId -> (width, height, relationship id); pic_data is filled
        # by the reader with the image bytes
        self.pic_bullets: dict[str, tuple[float, float, str]] = {}
        self.pic_data: dict[str, bytes] = {}
        if root is None:
            return
        for pb in root.findall(W + "numPicBullet"):
            shape = next(pb.iter(V + "shape"), None)
            img = next(pb.iter(V + "imagedata"), None)
            if shape is None or img is None:
                continue
            st = dict(p.split(":", 1) for p in (shape.get("style") or "").split(";") if ":" in p)
            w = num((st.get("width") or "9pt").replace("pt", ""), 9)
            h = num((st.get("height") or "9pt").replace("pt", ""), 9)
            self.pic_bullets[attr(pb, "numPicBulletId")] = (w, h, img.get(R + "id"))
        for an in root.findall(W + "abstractNum"):
            aid = attr(an, "abstractNumId")
            levels = {}
            for lvl in an.findall(W + "lvl"):
                levels[int(num(attr(lvl, "ilvl"), 0))] = _parse_level(lvl)
            self.abstract[aid] = levels
            nsl = an.find(W + "numStyleLink")
            if nsl is not None:
                self.style_link[aid] = attr(nsl, "val")
        for n in root.findall(W + "num"):
            nid = attr(n, "numId")
            a = n.find(W + "abstractNumId")
            overrides = {}
            for o in n.findall(W + "lvlOverride"):
                il = int(num(attr(o, "ilvl"), 0))
                so = o.find(W + "startOverride")
                lv = o.find(W + "lvl")
                overrides[il] = (int(num(attr(so, "val"), 1)) if so is not None else None,
                                 _parse_level(lv) if lv is not None else None)
            self.nums[nid] = (attr(a, "val") if a is not None else None, overrides)

    def _abstract_id(self, num_id: str, depth=0):
        entry = self.nums.get(num_id)
        if not entry:
            return None
        aid = entry[0]
        if aid in self.style_link and depth < 4:
            st = self.styles.styles.get(self.style_link[aid])
            if st is not None and st.ppr.get("numId"):
                return self._abstract_id(st.ppr["numId"], depth + 1)
        return aid

    def level(self, num_id: str | None, ilvl: int) -> Level | None:
        if not num_id or num_id == "0" or num_id not in self.nums:
            return None
        aid = self._abstract_id(num_id)
        _, overrides = self.nums[num_id]
        if ilvl in overrides and overrides[ilvl][1] is not None:
            return overrides[ilvl][1]
        levels = self.abstract.get(aid) or {}
        return levels.get(ilvl)

    def next_label(self, num_id: str, ilvl: int) -> tuple[str, Level] | None:
        lv = self.level(num_id, ilvl)
        if lv is None:
            return None
        aid = self._abstract_id(num_id) or num_id
        counts = self.counters.setdefault(aid, [None] * 9)
        _, overrides = self.nums[num_id]
        if num_id not in self.seen_nums:
            self.seen_nums.add(num_id)
            for il, (start, _lv) in overrides.items():
                if start is not None and 0 <= il < 9:
                    counts[il] = start - 1
        ilvl = max(0, min(8, ilvl))
        if counts[ilvl] is None:
            counts[ilvl] = lv.start
        else:
            counts[ilvl] += 1
        for deeper in range(ilvl + 1, 9):
            dl = self.level(num_id, deeper)
            if dl is not None and dl.restart == 0:
                continue
            counts[deeper] = None
        text = lv.text
        if lv.fmt == "bullet":
            return text, lv
        out = []
        i = 0
        while i < len(text):
            ch = text[i]
            if ch == "%" and i + 1 < len(text) and text[i + 1].isdigit():
                ref = int(text[i + 1]) - 1
                rl = self.level(num_id, ref)
                value = counts[ref] if 0 <= ref < 9 and counts[ref] is not None else (rl.start if rl else 1)
                fmt = rl.fmt if rl else "decimal"
                if lv.is_lgl and fmt not in ("decimalZero",):
                    fmt = "decimal"
                out.append(format_number(value, fmt))
                i += 2
                continue
            out.append(ch)
            i += 1
        return "".join(out), lv
