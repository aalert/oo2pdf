"""Style sheet resolution (docDefaults -> basedOn chains -> direct formatting)."""
from __future__ import annotations

from dataclasses import dataclass, field

from ..opc import W
from .props import attr, merge, parse_ppr, parse_rpr, parse_tblpr, parse_tcpr, parse_trpr


@dataclass
class Style:
    id: str
    type: str
    name: str = ""
    based_on: str | None = None
    ppr: dict = field(default_factory=dict)
    rpr: dict = field(default_factory=dict)
    tblpr: dict = field(default_factory=dict)
    trpr: dict = field(default_factory=dict)
    tcpr: dict = field(default_factory=dict)
    cond: dict = field(default_factory=dict)  # type -> dict(ppr, rpr, tblpr, trpr, tcpr)
    default: bool = False


class Styles:
    def __init__(self, root):
        self.default_rpr: dict = {}
        self.default_ppr: dict = {}
        self.styles: dict[str, Style] = {}
        self.default_para: str | None = None
        self.default_char: str | None = None
        self.default_table: str | None = None
        self._pcache: dict = {}
        self._ccache: dict = {}
        self._tcache: dict = {}
        if root is None:
            return
        dd = root.find(W + "docDefaults")
        if dd is not None:
            r = dd.find(f"{W}rPrDefault/{W}rPr")
            p = dd.find(f"{W}pPrDefault/{W}pPr")
            self.default_rpr = parse_rpr(r)
            self.default_ppr = parse_ppr(p)
        for s in root.findall(W + "style"):
            sid = attr(s, "styleId")
            if sid is None:
                continue
            st = Style(id=sid, type=attr(s, "type", "paragraph"))
            n = s.find(W + "name")
            st.name = attr(n, "val", sid) if n is not None else sid
            b = s.find(W + "basedOn")
            if b is not None:
                st.based_on = attr(b, "val")
            st.default = attr(s, "default") in ("1", "true", "on")
            st.ppr = parse_ppr(s.find(W + "pPr"))
            st.rpr = parse_rpr(s.find(W + "rPr"))
            st.tblpr = parse_tblpr(s.find(W + "tblPr"))
            st.trpr = parse_trpr(s.find(W + "trPr"))
            st.tcpr = parse_tcpr(s.find(W + "tcPr"))
            for c in s.findall(W + "tblStylePr"):
                st.cond[attr(c, "type")] = {
                    "ppr": parse_ppr(c.find(W + "pPr")),
                    "rpr": parse_rpr(c.find(W + "rPr")),
                    "tblpr": parse_tblpr(c.find(W + "tblPr")),
                    "trpr": parse_trpr(c.find(W + "trPr")),
                    "tcpr": parse_tcpr(c.find(W + "tcPr")),
                }
            self.styles[sid] = st
            if st.default:
                if st.type == "paragraph" and self.default_para is None:
                    self.default_para = sid
                elif st.type == "character" and self.default_char is None:
                    self.default_char = sid
                elif st.type == "table" and self.default_table is None:
                    self.default_table = sid

    def _chain(self, sid):
        chain = []
        seen = set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            st = self.styles[sid]
            chain.append(st)
            sid = st.based_on
        return list(reversed(chain))

    def para_style(self, sid: str | None):
        """Return (ppr, rpr) of a paragraph style including basedOn ancestry."""
        if sid is None or sid not in self.styles:
            sid = self.default_para
        if sid in self._pcache:
            return self._pcache[sid]
        ppr: dict = {}
        rpr: dict = {}
        for st in self._chain(sid):
            ppr = merge(ppr, st.ppr)
            rpr = merge(rpr, st.rpr)
        self._pcache[sid] = (ppr, rpr)
        return ppr, rpr

    def char_style(self, sid: str | None) -> dict:
        if sid is None:
            return {}
        if sid in self._ccache:
            return self._ccache[sid]
        rpr: dict = {}
        for st in self._chain(sid):
            rpr = merge(rpr, st.rpr)
        self._ccache[sid] = rpr
        return rpr

    def table_style(self, sid: str | None) -> Style:
        if sid is None or sid not in self.styles:
            sid = self.default_table
        if sid in self._tcache:
            return self._tcache[sid]
        out = Style(id=sid or "", type="table")
        for st in self._chain(sid):
            out.ppr = merge(out.ppr, st.ppr)
            out.rpr = merge(out.rpr, st.rpr)
            out.tblpr = merge(out.tblpr, st.tblpr)
            out.trpr = merge(out.trpr, st.trpr)
            out.tcpr = merge(out.tcpr, st.tcpr)
            for k, v in st.cond.items():
                cur = out.cond.get(k, {})
                out.cond[k] = {n: merge(cur.get(n, {}), v[n]) for n in v}
        self._tcache[sid] = out
        return out

