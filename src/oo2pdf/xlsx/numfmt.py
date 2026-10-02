"""Excel number format codes.

Implements the formatting language documented for Excel custom formats:
sections (positive;negative;zero;text), [Color] and [condition] prefixes,
digit placeholders (0 # ?), thousands separators and scaling commas,
percent, scientific notation, fractions, literal/escaped text, ``_x`` padding,
``*x`` fill, text ``@`` and the full date/time vocabulary including elapsed
time ([h], [m], [s]) and fractional seconds.
"""
from __future__ import annotations

import datetime as _dt
import math
import re
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from fractions import Fraction

from .locale import Locale, get_locale

MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December"]
DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

COLORS = {
    "black": (0, 0, 0), "blue": (0, 0, 1), "cyan": (0, 1, 1), "green": (0, 1, 0),
    "magenta": (1, 0, 1), "red": (1, 0, 0), "white": (1, 1, 1), "yellow": (1, 1, 0),
}

# Excel displays these built-in formats using the (en-US) locale patterns
LOCALE_BUILTINS = {
    "mm-dd-yy": "m/d/yyyy",
    "m/d/yy h:mm": "m/d/yyyy h:mm",
    "d-mmm-yy": "d-mmm-yy",
}


@dataclass
class Formatted:
    parts: list = field(default_factory=list)  # (kind, text): text | pad | fill
    color: tuple | None = None
    is_number: bool = False
    overflow_hash: bool = False  # numbers never spill into neighbours

    @property
    def text(self) -> str:
        return "".join(t for k, t in self.parts if k == "text")

    def plain(self) -> str:
        out = []
        for k, t in self.parts:
            if k == "text":
                out.append(t)
            elif k == "pad":
                out.append(" ")
        return "".join(out)


# --------------------------------------------------------------------------- parsing
_TOKEN = re.compile(
    r'"(?P<q>[^"]*)"'
    r"|\\(?P<esc>.)"
    r"|_(?P<pad>.)"
    r"|\*(?P<fill>.)"
    r"|\[(?P<br>[^\]]*)\]"
    r"|(?P<gen>[Gg]eneral)"
    r"|(?P<ampm>AM/PM|am/pm|A/P|a/p)"
    r"|(?P<exp>[eE][+-])"
    r"|(?P<date>[yY]+|[mM]+|[dD]+|[hH]+|[sS]+|[eE]|[gG]+)"
    r"|(?P<ch>.)",
    re.S,
)


@dataclass
class Section:
    tokens: list
    color: tuple | None = None
    cond: tuple | None = None  # (op, value)
    is_date: bool = False
    is_text: bool = False
    elapsed: bool = False
    general: bool = False


def split_sections(code: str) -> list[str]:
    out, cur = [], []
    i = 0
    in_q = False
    while i < len(code):
        c = code[i]
        if c == '"':
            in_q = not in_q
            cur.append(c)
        elif c == "\\" and not in_q and i + 1 < len(code):
            cur.append(code[i:i + 2])
            i += 2
            continue
        elif c == "[" and not in_q:
            j = code.find("]", i)
            if j == -1:
                j = len(code) - 1
            cur.append(code[i:j + 1])
            i = j + 1
            continue
        elif c == ";" and not in_q:
            out.append("".join(cur))
            cur = []
        else:
            cur.append(c)
        i += 1
    out.append("".join(cur))
    return out


_COND = re.compile(r"^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:[eE][+-]?\d+)?)$")


def parse_section(src: str) -> Section:
    toks = []
    sec = Section(tokens=toks)
    for m in _TOKEN.finditer(src):
        g = m.lastgroup
        v = m.group(g)
        if g == "q" or g == "esc":
            toks.append(("lit", v))
        elif g == "pad":
            toks.append(("pad", v))
        elif g == "fill":
            toks.append(("fill", v))
        elif g == "br":
            low = v.lower()
            if low in COLORS:
                sec.color = COLORS[low]
            elif low.startswith("color"):
                try:
                    idx = int(low[5:])
                    sec.color = _indexed(idx)
                except ValueError:
                    pass
            elif _COND.match(v.strip()):
                mm = _COND.match(v.strip())
                sec.cond = (mm.group(1), float(mm.group(2)))
            elif v.startswith("$"):
                sym = v[1:].split("-")[0]
                if sym:
                    toks.append(("lit", sym))
            elif re.fullmatch(r"h+|m+|s+", low):
                toks.append(("elapsed", low))
                sec.is_date = True
                sec.elapsed = True
            # [DBNum], [NatNum] and locale codes are ignored
        elif g == "ampm":
            toks.append(("ampm", v))
            sec.is_date = True
        elif g == "date":
            low = v.lower()
            if low[0] in "ge":
                toks.append(("date", "yyyy" if low[0] == "e" else low))
            else:
                toks.append(("date", low))
            sec.is_date = True
        elif g == "exp":
            toks.append(("exp", v[1]))
        elif g == "gen":
            toks.append(("general", v))
            sec.general = True
        else:
            c = v
            if c in "0#?":
                toks.append(("digit", c))
            elif c == ".":
                toks.append(("point", c))
            elif c == ",":
                toks.append(("comma", c))
            elif c == "%":
                toks.append(("percent", c))
            elif c == "@":
                toks.append(("text", c))
                sec.is_text = True
            elif c == "/":
                toks.append(("slash", c))
            else:
                toks.append(("lit", c))
    if sec.is_date and any(t[0] == "digit" for t in toks):
        # digits after seconds are fractional seconds: keep them as date tokens
        new = []
        for i, t in enumerate(toks):
            new.append(t)
        sec.tokens = new
    return sec


_INDEXED = [
    (0, 0, 0), (1, 1, 1), (1, 0, 0), (0, 1, 0), (0, 0, 1), (1, 1, 0), (1, 0, 1), (0, 1, 1),
    (0.5, 0, 0), (0, 0.5, 0), (0, 0, 0.5), (0.5, 0.5, 0), (0.5, 0, 0.5), (0, 0.5, 0.5),
    (0.75, 0.75, 0.75), (0.5, 0.5, 0.5),
]


def _indexed(i: int):
    if 1 <= i <= len(_INDEXED):
        return _INDEXED[i - 1]
    return None


_cache: dict[str, list[Section]] = {}


def parse(code: str) -> list[Section]:
    secs = _cache.get(code)
    if secs is None:
        secs = [parse_section(s) for s in split_sections(code)]
        _cache[code] = secs
    return secs


# --------------------------------------------------------------------------- helpers
def round_half_up(v: float, nd: int) -> Decimal:
    d = Decimal(repr(v))
    q = Decimal(1).scaleb(-nd)
    return d.quantize(q, rounding=ROUND_HALF_UP)


def serial_to_datetime(serial: float, date1904=False):
    if date1904:
        base = _dt.datetime(1904, 1, 1)
        return base + _dt.timedelta(days=serial), False
    days = int(math.floor(serial))
    frac = serial - days
    if days == 60:
        return (_dt.datetime(1900, 2, 28) + _dt.timedelta(days=frac)), True  # fake Feb 29
    if days < 60:
        base = _dt.datetime(1899, 12, 31)
    else:
        base = _dt.datetime(1899, 12, 30)
    return base + _dt.timedelta(days=serial), False


def to_serial(v, date1904=False) -> float | None:
    if isinstance(v, _dt.datetime):
        base = _dt.datetime(1904, 1, 1) if date1904 else _dt.datetime(1899, 12, 30)
        delta = v - base
        s = delta.days + delta.seconds / 86400 + delta.microseconds / 86400e6
        if not date1904 and v < _dt.datetime(1900, 3, 1):
            s -= 1  # Excel's fictitious 1900-02-29
        return s
    if isinstance(v, _dt.date):
        return to_serial(_dt.datetime(v.year, v.month, v.day), date1904)
    if isinstance(v, _dt.time):
        return (v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6) / 86400
    if isinstance(v, _dt.timedelta):
        return v.total_seconds() / 86400
    return None


# --------------------------------------------------------------------------- General
def format_general(v: float, max_chars: int = 11, loc: Locale | None = None) -> str:
    dec = loc.decimal if loc is not None else "."
    if v == 0:
        return "0"
    neg = v < 0
    a = abs(v)
    avail = max_chars - (1 if neg else 0)
    sign = "-" if neg else ""
    if avail < 1:
        return "#"
    if 1e-9 <= a < 1e11 or (a >= 1 and len(str(int(a))) <= avail):
        if a >= 1:
            int_digits = len(str(int(round_half_up(a, 0)))) if a < 1e15 else 16
        else:
            int_digits = 1
        if int_digits <= avail:
            decimals = max(0, avail - int_digits - 1)
            s = format(round_half_up(a, decimals), "f")
            if "." in s:
                s = s.rstrip("0").rstrip(".")
            if len(s) > avail:
                s = s[:avail].rstrip(".")
            if s not in ("0", ""):
                if a < 1 and len(s.replace("0.", "").lstrip("0")) < 1:
                    pass
                else:
                    return sign + s.replace(".", dec)
    # scientific
    exp = int(math.floor(math.log10(a)))
    mant = a / 10 ** exp
    exp_str = f"E{'+' if exp >= 0 else '-'}{abs(exp):02d}"
    room = avail - len(exp_str) - 2
    nd = max(0, min(5, room))
    m = round_half_up(mant, nd)
    if m >= 10:
        exp += 1
        m = round_half_up(a / 10 ** exp, nd)
        exp_str = f"E{'+' if exp >= 0 else '-'}{abs(exp):02d}"
    ms = format(m, "f")
    if "." in ms:
        ms = ms.rstrip("0").rstrip(".")
    return sign + ms.replace(".", dec) + exp_str


# --------------------------------------------------------------------------- formatting
def _choose(secs: list[Section], v: float):
    """Return (section, value_to_format, drop_sign)."""
    numeric = [s for s in secs[:3]]
    if any(s.cond for s in numeric):
        for idx, s in enumerate(numeric):
            if s.cond and _test(s.cond, v):
                return s, v, (v < 0 and idx > 0) or (s.cond[0] in ("<", "<=") and s.cond[1] <= 0)
        rest = [s for s in numeric if not s.cond]
        if rest:
            return rest[0], v, False
        return None, v, False
    n = len(numeric)
    if n == 1:
        return numeric[0], v, False
    if n == 2:
        if v < 0:
            return numeric[1], v, True
        return numeric[0], v, False
    if v > 0:
        return numeric[0], v, False
    if v < 0:
        return numeric[1], v, True
    return numeric[2], v, False


def _test(cond, v):
    op, x = cond
    return {"<": v < x, "<=": v <= x, ">": v > x, ">=": v >= x, "=": v == x, "<>": v != x}[op]


def format_value(value, code: str | None, date1904: bool = False, general_chars: int = 11,
                 loc: Locale | None = None) -> Formatted:
    """Format a cell value the way Excel displays it."""
    loc = loc or get_locale(None)
    if code is None or code == "":
        code = "General"
    if code == "mm-dd-yy":
        code = loc.short_date
    elif code == "m/d/yy h:mm":
        code = loc.short_date + " h:mm"
    else:
        code = LOCALE_BUILTINS.get(code, code)
    if value is None:
        return Formatted([])
    if isinstance(value, bool):
        return Formatted([("text", "TRUE" if value else "FALSE")])
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time, _dt.timedelta)):
        value = to_serial(value, date1904)
    secs = parse(code)
    if isinstance(value, str):
        if value.startswith("#") and value.upper() in _ERRORS:
            return Formatted([("text", value)])
        if len(secs) >= 4:
            sec = secs[3]
        elif any(s.is_text for s in secs[:1]) and len(secs) == 1:
            sec = secs[0]
        else:
            return Formatted([("text", value)])
        return _render_text(sec, value)
    try:
        v = float(value)
    except (TypeError, ValueError):
        return Formatted([("text", str(value))])
    if math.isnan(v) or math.isinf(v):
        return Formatted([("text", "#NUM!")])
    sec, v2, drop = _choose(secs, v)
    if sec is None:
        return Formatted([("text", "#" * 10)], is_number=True, overflow_hash=True)
    if sec.is_text and not any(t[0] in ("digit", "date", "general") for t in sec.tokens):
        # a text-only section used for a number: show literals
        pass
    if drop:
        v2 = abs(v2)
    if sec.general or (len(sec.tokens) == 0):
        txt = format_general(v2, general_chars, loc)
        out = Formatted(_with_literals(sec, txt), sec.color, True)
        out.general = True
        return out
    if sec.is_date:
        if v2 < 0:
            return Formatted([("text", "#" * 10)], is_number=True, overflow_hash=True)
        return _render_date(sec, v2, date1904, loc)
    return _render_number(sec, v2, loc)


_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA",
           "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#BUSY!", "#UNKNOWN!"}


def _with_literals(sec: Section, txt: str):
    parts = []
    placed = False
    for kind, v in sec.tokens:
        if kind == "general":
            parts.append(("text", txt))
            placed = True
        elif kind == "lit":
            parts.append(("text", v))
        elif kind == "pad":
            parts.append(("pad", v))
        elif kind == "fill":
            parts.append(("fill", v))
    if not placed:
        parts.append(("text", txt))
    return parts


def _render_text(sec: Section, value: str) -> Formatted:
    parts = []
    for kind, v in sec.tokens:
        if kind == "text":
            parts.append(("text", value))
        elif kind == "lit":
            parts.append(("text", v))
        elif kind == "pad":
            parts.append(("pad", v))
        elif kind == "fill":
            parts.append(("fill", v))
        elif kind in ("digit", "point", "comma", "slash", "percent"):
            parts.append(("text", v))
    return Formatted(parts, sec.color)


def _render_number(sec: Section, v: float, loc: Locale) -> Formatted:
    toks = sec.tokens
    neg = v < 0
    a = abs(v)
    pct = sum(1 for t in toks if t[0] == "percent")
    a *= 100 ** pct
    has_exp = any(t[0] == "exp" for t in toks)
    has_slash = any(t[0] == "slash" for t in toks) and any(t[0] == "digit" for t in toks)
    if has_exp:
        parts = _scientific(toks, a, loc)
    elif has_slash:
        parts = _fraction(toks, a)
        if parts is None:
            parts = _fixed(toks, a, loc)
    else:
        parts = _fixed(toks, a, loc)
    if neg and not _is_zero_display(parts):
        # the minus sign precedes the first digit / literal
        parts.insert(0, ("text", "-"))
    return Formatted(parts, sec.color, True, overflow_hash=True)


def _is_zero_display(parts):
    """True when every digit shown is zero (Excel then omits the minus sign)."""
    txt = "".join(t for k, t in parts if k == "text")
    return any(c.isdigit() for c in txt) and not any(c.isdigit() and c != "0" for c in txt)


def _fixed(toks, a: float, loc: Locale | None = None):
    dec_sep = loc.decimal if loc is not None else "."
    grp_sep = loc.group if loc is not None else ","
    # locate decimal point among number tokens
    point_idx = next((i for i, t in enumerate(toks) if t[0] == "point"), None)
    digit_idx = [i for i, t in enumerate(toks) if t[0] == "digit"]
    int_end = point_idx if point_idx is not None else (digit_idx[-1] + 1 if digit_idx else len(toks))
    # scaling commas: commas directly after the last integer placeholder
    last_int_digit = max((i for i in digit_idx if i < int_end), default=None)
    scale = 0
    thousands = False
    if last_int_digit is not None:
        j = last_int_digit + 1
        while j < len(toks) and toks[j][0] == "comma":
            scale += 1
            j += 1
        first_digit = min(digit_idx)
        for i in range(first_digit, last_int_digit):
            if toks[i][0] == "comma":
                thousands = True
    elif point_idx is None and digit_idx:
        pass
    # trailing commas after the decimals also scale
    if digit_idx:
        j = digit_idx[-1] + 1
        while j < len(toks) and toks[j][0] == "comma" and (point_idx is not None and digit_idx[-1] > point_idx):
            scale += 1
            j += 1
    a = a / (1000 ** scale)
    dec_tokens = [t for t in toks[point_idx + 1:] if t[0] == "digit"] if point_idx is not None else []
    nd = len(dec_tokens)
    r = round_half_up(a, nd)
    s = format(r, "f")
    if "." in s:
        ip, fp = s.split(".")
    else:
        ip, fp = s, ""
    if ip == "0":
        ip = ""
    # integer placeholders
    int_tokens = [(i, t) for i, t in enumerate(toks[:int_end]) if t[0] == "digit"]
    n_int = len(int_tokens)
    digits = list(ip)
    # pad for 0 / ? placeholders
    out_int: dict[int, str] = {}
    k = len(digits) - 1
    for idx in range(n_int - 1, -1, -1):
        ti, (kind, ch) = int_tokens[idx]
        if k >= 0:
            out_int[ti] = digits[k]
            k -= 1
        else:
            out_int[ti] = "0" if ch == "0" else (" " if ch == "?" else "")
    if k >= 0 and int_tokens:
        first_ti = int_tokens[0][0]
        out_int[first_ti] = "".join(digits[:k + 1]) + out_int[first_ti]
    if thousands and int_tokens:
        # regroup the concatenated integer digits
        seq = "".join(out_int[ti] for ti, _ in int_tokens)
        lead = len(seq) - len(seq.lstrip(" "))
        core = seq.strip(" ")
        if core.isdigit():
            grouped = _group(core) if core else ""
            grouped = grouped.replace(",", grp_sep)
            first_ti = int_tokens[0][0]
            for ti, _ in int_tokens:
                out_int[ti] = ""
            out_int[first_ti] = " " * lead + grouped
    # decimal placeholders
    out_dec: dict[int, str] = {}
    if point_idx is not None:
        dec_positions = [i for i in range(point_idx + 1, len(toks)) if toks[i][0] == "digit"]
        fds = list(fp.ljust(nd, "0"))
        trailing = True
        for j in range(len(dec_positions) - 1, -1, -1):
            ti = dec_positions[j]
            ch = toks[ti][1]
            d = fds[j] if j < len(fds) else "0"
            if trailing and d == "0" and ch in "#?":
                out_dec[ti] = "" if ch == "#" else " "
            else:
                trailing = False
                out_dec[ti] = d
    parts = []
    for i, (kind, v) in enumerate(toks):
        if kind == "digit":
            val = out_int.get(i, out_dec.get(i, ""))
            if val:
                parts.append(("text", val))
        elif kind == "point":
            parts.append(("text", dec_sep))
        elif kind == "comma":
            if last_int_digit is not None and i < last_int_digit and i > min(digit_idx):
                continue  # thousands marker
            if scale and ((last_int_digit is not None and i > last_int_digit) or
                          (digit_idx and i > digit_idx[-1])):
                continue
            parts.append(("text", ","))
        elif kind == "percent":
            parts.append(("text", "%"))
        elif kind == "lit":
            parts.append(("text", v))
        elif kind == "pad":
            parts.append(("pad", v))
        elif kind == "fill":
            parts.append(("fill", v))
        elif kind == "text":
            pass
        elif kind == "slash":
            parts.append(("text", "/"))
        elif kind in ("date", "ampm", "exp", "general", "elapsed"):
            parts.append(("text", v))
    if not digit_idx:
        # format without placeholders: Excel shows only literals
        pass
    return parts


def _group(core: str) -> str:
    out = []
    for i, ch in enumerate(reversed(core)):
        if i and i % 3 == 0:
            out.append(",")
        out.append(ch)
    return "".join(reversed(out))


def _scientific(toks, a: float, loc: Locale | None = None):
    ei = next(i for i, t in enumerate(toks) if t[0] == "exp")
    mant_toks = toks[:ei]
    exp_toks = toks[ei + 1:]
    point_idx = next((i for i, t in enumerate(mant_toks) if t[0] == "point"), None)
    int_ph = [t for t in (mant_toks[:point_idx] if point_idx is not None else mant_toks) if t[0] == "digit"]
    dec_ph = [t for t in mant_toks[point_idx + 1:] if t[0] == "digit"] if point_idx is not None else []
    ni = max(1, len(int_ph))
    nd = len(dec_ph)
    if a == 0:
        exp = 0
        mant = 0.0
    else:
        exp = int(math.floor(math.log10(a)))
        if len(int_ph) > 1 and any(t[1] == "#" for t in int_ph):
            exp = int(math.floor(exp / ni) * ni)
        mant = a / 10 ** exp
        m = float(round_half_up(mant, nd))
        if m >= 10 ** (ni if len(int_ph) > 1 and any(t[1] == "#" for t in int_ph) else 1):
            exp += ni if len(int_ph) > 1 and any(t[1] == "#" for t in int_ph) else 1
            mant = a / 10 ** exp
    mparts = _fixed(mant_toks, mant, loc)
    sign = toks[ei][1]
    exp_digits = [t for t in exp_toks if t[0] == "digit"]
    ndig = sum(1 for t in exp_digits if t[1] == "0")
    es = str(abs(exp)).rjust(max(1, ndig), "0")
    esign = "-" if exp < 0 else ("+" if sign == "+" else "")
    parts = list(mparts) + [("text", "E" + esign + es)]
    for kind, v in exp_toks:
        if kind == "lit":
            parts.append(("text", v))
    return parts


def _fraction(toks, a: float):
    si = next(i for i, t in enumerate(toks) if t[0] == "slash")
    # denominator: digits (or a literal number) after the slash
    den_toks = []
    j = si + 1
    fixed_den = ""
    while j < len(toks) and toks[j][0] in ("digit", "lit") and (toks[j][0] == "digit" or toks[j][1].isdigit()):
        if toks[j][0] == "lit" or (toks[j][1].isdigit() and toks[j][1] != "0" and not den_toks):
            fixed_den += toks[j][1]
        den_toks.append(toks[j])
        j += 1
    # numerator: digits immediately before the slash
    k = si - 1
    num_toks = []
    while k >= 0 and toks[k][0] == "digit":
        num_toks.insert(0, toks[k])
        k -= 1
    int_toks = [t for t in toks[:k + 1] if t[0] == "digit"]
    whole = 0
    frac = a
    if int_toks:
        whole = int(math.floor(a))
        frac = a - whole
    if fixed_den and fixed_den.isdigit() and all(t[0] == "lit" or t[1].isdigit() for t in den_toks):
        den = int(fixed_den)
        num = int(round_half_up(frac * den, 0))
    else:
        maxd = 10 ** max(1, len(den_toks)) - 1
        f = Fraction(frac).limit_denominator(maxd)
        num, den = f.numerator, f.denominator
    if int_toks and num == den and den:
        whole += 1
        num = 0
    parts = []
    if int_toks:
        ws = str(whole) if whole or num == 0 else ""
        if not ws and any(t[1] == "0" for t in int_toks):
            ws = "0"
        parts.append(("text", ws))
        # literals between integer part and numerator
        for t in toks[len([x for x in toks[:k + 1]]) - 0:k + 1]:
            pass
        for t in toks[:k + 1]:
            if t[0] == "lit":
                parts.append(("text", t[1] if (ws or num) else " "))
    if num == 0 and int_toks:
        width = len(num_toks) + 1 + len(den_toks)
        parts.append(("text", " " * width))
        return parts
    ns = str(num).rjust(len(num_toks), " ") if any(t[1] == "?" for t in num_toks) else str(num)
    ds = str(den).ljust(len(den_toks), " ") if any(t[1] == "?" for t in den_toks) else str(den)
    parts.append(("text", ns + "/" + ds))
    for t in toks[j:]:
        if t[0] == "lit":
            parts.append(("text", t[1]))
    return parts


def _render_date(sec: Section, serial: float, date1904: bool, loc: Locale | None = None) -> Formatted:
    loc = loc or get_locale(None)
    MONTHS = loc.months
    DAYS = loc.days
    toks = sec.tokens
    has_ampm = any(t[0] == "ampm" for t in toks)
    # sub-second precision requested?
    frac_digits = 0
    for i, (k, v) in enumerate(toks):
        if k == "point" and i + 1 < len(toks) and toks[i + 1][0] == "digit":
            j = i + 1
            while j < len(toks) and toks[j][0] == "digit":
                frac_digits += 1
                j += 1
    unit = 86400 * 10 ** frac_digits
    total = float(round_half_up(serial * unit, 0)) / unit
    dt, fake29 = serial_to_datetime(total, date1904)
    day_serial = int(math.floor(total))
    # minute vs month disambiguation
    kinds = [t for t in toks]
    resolved = []
    for i, (k, v) in enumerate(kinds):
        if k == "date" and v[0] == "m" and len(v) <= 2:
            prev = next((kinds[j] for j in range(i - 1, -1, -1) if kinds[j][0] in ("date", "elapsed")), None)
            nxt = next((kinds[j] for j in range(i + 1, len(kinds)) if kinds[j][0] in ("date", "elapsed")), None)
            if (prev and prev[1][0] == "h") or (nxt and nxt[1][0] == "s"):
                resolved.append(("minute", v))
                continue
        resolved.append((k, v))
    parts = []
    year, month, day = (1900, 2, 29) if fake29 else (dt.year, dt.month, dt.day)
    hour, minute, second = dt.hour, dt.minute, dt.second
    micro = dt.microsecond
    if date1904:
        weekday = (day_serial + 5) % 7
    else:
        weekday = (day_serial + 6) % 7
    i = 0
    while i < len(resolved):
        k, v = resolved[i]
        if k == "date":
            c = v[0]
            n = len(v)
            if c == "y":
                parts.append(("text", f"{year % 100:02d}" if n <= 2 else f"{year:04d}"))
            elif c == "m":
                if n == 1:
                    parts.append(("text", str(month)))
                elif n == 2:
                    parts.append(("text", f"{month:02d}"))
                elif n == 3:
                    parts.append(("text", loc.months_abbr[month - 1]))
                elif n == 5:
                    parts.append(("text", MONTHS[month - 1][0]))
                else:
                    parts.append(("text", MONTHS[month - 1]))
            elif c == "d":
                if n == 1:
                    parts.append(("text", str(day)))
                elif n == 2:
                    parts.append(("text", f"{day:02d}"))
                elif n == 3:
                    parts.append(("text", loc.days_abbr[weekday]))
                else:
                    parts.append(("text", DAYS[weekday]))
            elif c == "h":
                h = hour
                if has_ampm:
                    h = hour % 12 or 12
                parts.append(("text", f"{h:02d}" if n >= 2 else str(h)))
            elif c == "s":
                parts.append(("text", f"{second:02d}" if n >= 2 else str(second)))
            elif c == "g":
                parts.append(("text", ""))
        elif k == "minute":
            parts.append(("text", f"{minute:02d}" if len(v) >= 2 else str(minute)))
        elif k == "elapsed":
            c = v[0]
            if c == "h":
                val = int(math.floor(total * 24 + 1e-9))
            elif c == "m":
                val = int(math.floor(total * 1440 + 1e-9))
            else:
                val = int(math.floor(total * 86400 + 1e-9))
            parts.append(("text", str(val).rjust(len(v), "0")))
        elif k == "ampm":
            pm = hour >= 12
            if v.upper() == "AM/PM":
                parts.append(("text", loc.pm if pm else loc.am))
            else:
                a, p = v.split("/")
                parts.append(("text", p if pm else a))
        elif k == "point":
            if i + 1 < len(resolved) and resolved[i + 1][0] == "digit":
                j = i + 1
                n = 0
                while j < len(resolved) and resolved[j][0] == "digit":
                    n += 1
                    j += 1
                frac = f"{micro / 1e6:.{n}f}"[1:].replace(".", loc.decimal)
                parts.append(("text", frac))
                i = j
                continue
            parts.append(("text", "."))
        elif k == "lit":
            # unquoted / and : are the regional date and time separators
            if v == ":":
                v = loc.time_sep
            parts.append(("text", v))
        elif k == "pad":
            parts.append(("pad", v))
        elif k == "fill":
            parts.append(("fill", v))
        elif k == "slash":
            parts.append(("text", loc.date_sep))
        elif k in ("comma", "percent", "digit"):
            parts.append(("text", v))
        i += 1
    return Formatted(parts, sec.color, True, overflow_hash=True)


