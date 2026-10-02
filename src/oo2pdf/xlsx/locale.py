"""Regional settings used by Excel when displaying numbers and dates.

Excel stores format codes in an invariant (en-US) form and renders them with
the operating system's regional settings: decimal/grouping separators, month
and day names, AM/PM designators and the short date pattern used by built-in
format 14.
"""
from __future__ import annotations

import sys
from dataclasses import dataclass


@dataclass(frozen=True)
class Locale:
    name: str = "en-US"
    decimal: str = "."
    group: str = ","
    short_date: str = "m/d/yyyy"  # Excel format code for built-in 14
    months: tuple = ("January", "February", "March", "April", "May", "June", "July",
                     "August", "September", "October", "November", "December")
    months_abbr: tuple = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep",
                          "Oct", "Nov", "Dec")
    days: tuple = ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    days_abbr: tuple = ("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    am: str = "AM"
    pm: str = "PM"
    list_sep: str = ","
    currency: str = "$"
    currency_pos: int = 0  # 0 $1, 1 1$, 2 $ 1, 3 1 $
    date_sep: str = "/"
    time_sep: str = ":"


_ES_MONTHS = ("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
              "septiembre", "octubre", "noviembre", "diciembre")
_ES_DAYS = ("domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado")
_DE_MONTHS = ("Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August",
              "September", "Oktober", "November", "Dezember")
_DE_DAYS = ("Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag")
_FR_MONTHS = ("janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
              "septembre", "octobre", "novembre", "décembre")
_FR_DAYS = ("dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi")
_PT_MONTHS = ("janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto",
              "setembro", "outubro", "novembro", "dezembro")
_PT_DAYS = ("domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira",
            "sexta-feira", "sábado")


def _abbr(names, n=3):
    return tuple(x[:n] for x in names)


LOCALES: dict[str, Locale] = {
    "en-US": Locale(),
    "en-GB": Locale("en-GB", short_date="dd/mm/yyyy", am="am", pm="pm"),
    "es-ES": Locale("es-ES", ",", ".", "dd/mm/yyyy", _ES_MONTHS, _abbr(_ES_MONTHS), _ES_DAYS,
                    _abbr(_ES_DAYS), "a. m.", "p. m.", ";", "€", 3),
    "es-CL": Locale("es-CL", ",", ".", "dd-mm-yyyy", _ES_MONTHS, _abbr(_ES_MONTHS), _ES_DAYS,
                    _abbr(_ES_DAYS), "a. m.", "p. m.", ";"),
    "es-MX": Locale("es-MX", ".", ",", "dd/mm/yyyy", _ES_MONTHS, _abbr(_ES_MONTHS), _ES_DAYS,
                    _abbr(_ES_DAYS), "a. m.", "p. m.", ","),
    "de-DE": Locale("de-DE", ",", ".", "dd.mm.yyyy", _DE_MONTHS, _abbr(_DE_MONTHS), _DE_DAYS,
                    _abbr(_DE_DAYS, 2), "AM", "PM", ";"),
    "fr-FR": Locale("fr-FR", ",", " ", "dd/mm/yyyy", _FR_MONTHS, _abbr(_FR_MONTHS, 4),
                    _FR_DAYS, _abbr(_FR_DAYS), "AM", "PM", ";"),
    "pt-BR": Locale("pt-BR", ",", ".", "dd/mm/yyyy", _PT_MONTHS, _abbr(_PT_MONTHS), _PT_DAYS,
                    _abbr(_PT_DAYS), "AM", "PM", ";", "R$", 2),
}


def _with_seps(loc: Locale) -> Locale:
    from dataclasses import replace
    return replace(loc, date_sep=_date_sep(loc.short_date))


def get_locale(spec) -> Locale:
    if isinstance(spec, Locale):
        return spec
    if spec is None:
        return LOCALES["en-US"]
    if spec == "system":
        return system_locale()
    if spec in LOCALES:
        return LOCALES[spec]
    lang = str(spec).split("-")[0].lower()
    for k, v in LOCALES.items():
        if k.lower().startswith(lang + "-"):
            return Locale(str(spec), v.decimal, v.group, v.short_date, v.months, v.months_abbr,
                          v.days, v.days_abbr, v.am, v.pm, v.list_sep, v.currency, v.currency_pos,
                          v.date_sep, v.time_sep)
    return LOCALES["en-US"]


def _excel_date_pattern(p: str) -> str:
    """Convert a Windows short date pattern (dd-MM-yyyy) into an Excel code."""
    out = []
    i = 0
    while i < len(p):
        c = p[i]
        j = i
        while j < len(p) and p[j] == c:
            j += 1
        run = p[i:j]
        if c == "M":
            out.append(run.lower().replace("m", "m"))
        elif c in "dy":
            out.append(run)
        elif c == "'":
            k = p.find("'", i + 1)
            if k == -1:
                k = len(p)
            out.append('"' + p[i + 1:k] + '"')
            j = k + 1
        else:
            out.append(run)
        i = j
    return "".join(out)


def system_locale() -> Locale:
    if sys.platform != "win32":
        import locale as _pylocale
        try:
            conv = _pylocale.localeconv()
            base = get_locale((_pylocale.getlocale()[0] or "en_US").replace("_", "-"))
            return Locale(base.name, conv.get("decimal_point") or base.decimal,
                          conv.get("thousands_sep") or base.group, base.short_date, base.months,
                          base.months_abbr, base.days, base.days_abbr, base.am, base.pm, base.list_sep)
        except Exception:
            return LOCALES["en-US"]
    import ctypes
    k32 = ctypes.windll.kernel32
    buf = ctypes.create_unicode_buffer(256)

    def info(code):
        n = k32.GetLocaleInfoEx(None, code, buf, 256)
        return buf.value if n else ""
    name = info(0x5C) or "en-US"  # LOCALE_SNAME
    months = tuple(info(0x38 + i) for i in range(12))
    months_abbr = tuple(info(0x44 + i) for i in range(12))
    # LOCALE_SDAYNAME1 is Monday; Excel's order starts on Sunday
    days = [info(0x2A + i) for i in range(7)]
    days_abbr = [info(0x31 + i) for i in range(7)]
    days = tuple([days[6]] + days[:6])
    days_abbr = tuple([days_abbr[6]] + days_abbr[:6])
    return Locale(
        name=name, decimal=info(0x0E) or ".", group=info(0x0F) or ",",
        short_date=_excel_date_pattern(info(0x1F) or "M/d/yyyy"),
        months=months, months_abbr=months_abbr, days=days, days_abbr=days_abbr,
        am=info(0x28) or "AM", pm=info(0x29) or "PM", list_sep=info(0x0C) or ",",
        currency=info(0x14) or "$", currency_pos=int(info(0x1B) or 0),
        date_sep=_date_sep(info(0x1F) or "M/d/yyyy"), time_sep=info(0x1E) or ":",
    )


def _date_sep(pattern: str) -> str:
    """The separator used in a Windows short date pattern (LOCALE_SDATE is deprecated)."""
    for ch in pattern:
        if not ch.isalpha() and ch != "'":
            return ch
    return "/"


for _k in list(LOCALES):
    LOCALES[_k] = _with_seps(LOCALES[_k])
