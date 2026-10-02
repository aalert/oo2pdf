"""Reads a .docx package into the resolved model in :mod:`model`."""
from __future__ import annotations

import re

from ..opc import A, M, PIC, R, V, W, WP, WPG, WPS, Package, local
from ..theme import Theme, drawingml_color, hex_color
from .model import (Bookmark, Break, Cell, Field, Floating, FootnoteRef, MathItem, Paragraph,
                    Picture, Row, RunStyle, Section, Separator, Shape, Table, Tab, Text)
from .numbering import Numbering, format_number
from .props import (attr, merge, num, onoff, parse_borders, parse_ppr, parse_rpr,
                    parse_tblpr, parse_tcpr, parse_trpr, twips)
from .styles import Styles

EMU = 12700.0

HIGHLIGHT = {
    "yellow": (1, 1, 0), "green": (0, 1, 0), "cyan": (0, 1, 1), "magenta": (1, 0, 1),
    "blue": (0, 0, 1), "red": (1, 0, 0), "darkBlue": (0, 0, 0.5), "darkCyan": (0, 0.5, 0.5),
    "darkGreen": (0, 0.5, 0), "darkMagenta": (0.5, 0, 0.5), "darkRed": (0.5, 0, 0),
    "darkYellow": (0.5, 0.5, 0), "darkGray": (0.5, 0.5, 0.5), "lightGray": (0.75, 0.75, 0.75),
    "black": (0, 0, 0), "white": (1, 1, 1),
}

PAGE_FIELDS = {"PAGE", "NUMPAGES", "SECTIONPAGES"}

_MC_OK = {"wps", "wpg", "wpc", "wp14", "w14", "w15", "w16", "a14", "v", "o", "w10", "wne", "mc"}


class DocxReader:
    def __init__(self, pkg: Package):
        self.pkg = pkg
        self.main = pkg.main_part() or "word/document.xml"
        rels = pkg.rels(self.main)

        def part(t):
            for rtype, target, ext in rels.values():
                if rtype == t and not ext:
                    return target
            return None
        tpart = part("theme")
        self.theme = Theme(pkg.xml(tpart) if tpart else None)
        spart = part("styles")
        self.styles = Styles(pkg.xml(spart) if spart else None)
        npart = part("numbering")
        self.numbering = Numbering(pkg.xml(npart) if npart else None, self.styles)
        self._numbering_part = npart
        self.default_tab = 36.0
        self.compat_mode = 15
        self.even_odd = False
        # Word collapses paragraph spacing to max(after, before) unless this
        # compatibility option is set, in which case the two are added.
        self.sum_spacing = False
        # Word 2013+: a Normal-style font size / justification other than
        # 11 pt / left beats the table style (otherwise the table style wins)
        self.normal_beats_tstyle = False
        setpart = part("settings")
        sroot = pkg.xml(setpart) if setpart else None
        if sroot is not None:
            dt = sroot.find(W + "defaultTabStop")
            if dt is not None:
                self.default_tab = twips(attr(dt, "val"), 720) or 36.0
            self.even_odd = sroot.find(W + "evenAndOddHeaders") is not None \
                and onoff(sroot.find(W + "evenAndOddHeaders"))
            compat = sroot.find(W + "compat")
            if compat is not None and compat.find(W + "doNotUseHTMLParagraphAutoSpacing") is not None:
                self.sum_spacing = onoff(compat.find(W + "doNotUseHTMLParagraphAutoSpacing"))
            for cs in sroot.iter(W + "compatSetting"):
                if attr(cs, "name") == "overrideTableStyleFontSizeAndJustification":
                    self.normal_beats_tstyle = attr(cs, "val") in ("1", "true", "on")
                if attr(cs, "name") == "compatibilityMode":
                    try:
                        self.compat_mode = int(attr(cs, "val", "15"))
                    except ValueError:
                        pass
            if not any(True for _ in sroot.iter(W + "compatSetting")):
                self.compat_mode = 12 if sroot.find(W + "compat") is not None else 15
        for pid, (_, _, rid) in self.numbering.pic_bullets.items():
            data = self._blob(self._numbering_part, rid)
            if data:
                self.numbering.pic_data[pid] = data
        self.footnote_part = part("footnotes")
        self.endnote_part = part("endnotes")
        self.footnotes: dict[str, list] = {}
        self.footnote_sep: list = []
        self.endnote_sep: list = []
        self._fn_counter = 0
        self._en_counter = 0
        self.endnote_order: list = []
        self._fields: list[dict] = []
        self._style_cache: dict = {}
        self._bookmark_ids: dict[str, str] = {}
        self.base_rpr = self.styles.default_rpr
        core = pkg.xml("docProps/core.xml")
        self.title = None
        self.author = None
        if core is not None:
            for el in core:
                if local(el) == "title" and el.text:
                    self.title = el.text
                elif local(el) == "creator" and el.text:
                    self.author = el.text

    # ------------------------------------------------------------------ styles
    def run_style(self, rpr: dict) -> RunStyle:
        key = tuple(sorted((k, _freeze(v)) for k, v in rpr.items()))
        rs = self._style_cache.get(key)
        if rs is not None:
            return rs
        th = self.theme

        def font(slot, default):
            t = rpr.get("theme_" + slot)
            if t:
                f = th.font(t)
                if f:
                    return f
            return rpr.get("font_" + slot) or default
        ascii_ = font("ascii", None)
        hansi = font("hAnsi", None)
        ascii_ = ascii_ or hansi or "Times New Roman"
        hansi = hansi or ascii_
        ea = font("eastAsia", None) or ascii_
        cs = font("cs", None) or ascii_
        size = rpr.get("sz", 10.0)
        hl = rpr.get("highlight")
        shd = rpr.get("shd")
        rs = RunStyle(
            font_ascii=ascii_, font_hansi=hansi, font_ea=ea, font_cs=cs,
            size=size, size_cs=rpr.get("szCs", size),
            bold=rpr.get("b", False), italic=rpr.get("i", False),
            bold_cs=rpr.get("bCs", rpr.get("b", False)), italic_cs=rpr.get("iCs", rpr.get("i", False)),
            underline=None if rpr.get("u") in (None, "none") else rpr.get("u"),
            u_color=rpr.get("u_color"),
            strike=rpr.get("strike", False), dstrike=rpr.get("dstrike", False),
            color=rpr.get("color"),
            highlight=HIGHLIGHT.get(hl) if hl and hl != "none" else None,
            shading=shd if isinstance(shd, tuple) else None,
            vert=rpr.get("vertAlign") if rpr.get("vertAlign") in ("superscript", "subscript") else None,
            caps=rpr.get("caps", False), small_caps=rpr.get("smallCaps", False),
            spacing=rpr.get("spacing", 0.0), position=rpr.get("position", 0.0),
            scale=rpr.get("scale", 100.0) or 100.0,
            hidden=rpr.get("vanish", False) and not rpr.get("specVanish", False),
            hint=rpr.get("hint"),
            border=rpr.get("bdr") if rpr.get("bdr") and rpr["bdr"].visible else None,
            kern=rpr.get("kern", 0.0) or 0.0,
            math_script=bool(rpr.get("math_script")),
        )
        self._style_cache[key] = rs
        return rs

    def para_props(self, ppr_el, tstyle=None):
        """Resolve paragraph properties; returns (ppr, base_rpr, style_id)."""
        direct = parse_ppr(ppr_el)
        sid = direct.get("pStyle")
        if sid not in self.styles.styles:
            sid = self.styles.default_para
        s_ppr, s_rpr = self.styles.para_style(sid)
        defaults_p = self.styles.default_ppr
        defaults_r = self.styles.default_rpr
        if tstyle is not None and sid == self.styles.default_para:
            # Normal style in a table: its paragraph properties beat the table
            # style's, but the table style's character formatting beats Normal's
            # (overrideTableStyleFontSizeAndJustification keeps a Normal font
            # size other than 11 pt)
            t_ppr, t_rpr = tstyle
            ppr = merge(merge(defaults_p, t_ppr), s_ppr)
            if self.normal_beats_tstyle:
                rpr = merge(merge(defaults_r, s_rpr), t_rpr)
                for key in ("sz", "szCs"):
                    if key in s_rpr and s_rpr[key] != 11.0:
                        rpr[key] = s_rpr[key]
            else:
                # without that option Normal's font size is not used in a table
                # (demo.docx: Normal 12 pt, table text at the 11 pt default)
                n_rpr = {k: v for k, v in s_rpr.items() if k not in ("sz", "szCs")}
                rpr = merge(merge(defaults_r, n_rpr), t_rpr)
        elif tstyle is not None:
            # table style formatting sits between the document defaults and any
            # other paragraph style: the paragraph style wins
            t_ppr, t_rpr = tstyle
            ppr = merge(merge(defaults_p, t_ppr), s_ppr)
            rpr = merge(merge(defaults_r, t_rpr), s_rpr)
        else:
            ppr = merge(defaults_p, s_ppr)
            rpr = merge(defaults_r, s_rpr)
        num_id = direct.get("numId", s_ppr.get("numId"))
        ilvl = direct.get("ilvl", s_ppr.get("ilvl", 0))
        level = self.numbering.level(num_id, ilvl) if num_id else None
        if level is not None:
            if "numId" in direct:
                ppr = merge(ppr, level.ppr)
            else:
                # numbering defined by the style: the style's own indents win
                ppr = merge(merge(merge(defaults_p, level.ppr), s_ppr), {})
        ppr = merge(ppr, direct)
        ppr["numId"] = num_id
        ppr["ilvl"] = ilvl
        ppr["pStyle"] = sid
        return ppr, rpr, sid

    # ------------------------------------------------------------------ body
    def read_document(self) -> list[Section]:
        root = self.pkg.xml(self.main)
        body = root.find(W + "body")
        self._load_notes()
        sections: list[Section] = []
        blocks: list = []
        prev: Section | None = None
        for blk in self._blocks(body, self.main, None):
            blocks.append(blk)
            if isinstance(blk, Paragraph) and blk.sect is not None:
                sect = blk.sect
                self._inherit_hf(sect, prev)
                sect.blocks = blocks
                sections.append(sect)
                prev = sect
                blocks = []
        final = body.find(W + "sectPr")
        sect = self.read_sectpr(final)
        self._inherit_hf(sect, prev)
        sect.blocks = blocks
        sections.append(sect)
        return sections

    def _inherit_hf(self, sect, prev):
        if prev is None:
            return
        for kind in ("default", "first", "even"):
            if kind not in sect.headers and kind in prev.headers:
                sect.headers[kind] = prev.headers[kind]
            if kind not in sect.footers and kind in prev.footers:
                sect.footers[kind] = prev.footers[kind]

    def read_sectpr(self, el) -> Section:
        s = Section()
        if el is None:
            return s
        pg = el.find(W + "pgSz")
        if pg is not None:
            s.page_w = twips(attr(pg, "w"), 12240)
            s.page_h = twips(attr(pg, "h"), 15840)
            if attr(pg, "orient") == "landscape" and s.page_w < s.page_h:
                s.page_w, s.page_h = s.page_h, s.page_w
        mar = el.find(W + "pgMar")
        if mar is not None:
            s.margin_top = twips(attr(mar, "top"), 1440)
            s.margin_bottom = twips(attr(mar, "bottom"), 1440)
            s.margin_left = twips(attr(mar, "left"), 1440)
            s.margin_right = twips(attr(mar, "right"), 1440)
            s.header_dist = twips(attr(mar, "header"), 720)
            s.footer_dist = twips(attr(mar, "footer"), 720)
            s.gutter = twips(attr(mar, "gutter"), 0)
        cols = el.find(W + "cols")
        if cols is not None:
            s.cols = max(1, int(num(attr(cols, "num"), 1)))
            s.col_space = twips(attr(cols, "space"), 720)
            s.col_sep = attr(cols, "sep") in ("1", "true", "on")
            if attr(cols, "equalWidth") in ("0", "false", "off"):
                widths = []
                for c in cols.findall(W + "col"):
                    widths.append((twips(attr(c, "w"), 0), twips(attr(c, "space"), 0)))
                if widths:
                    s.col_widths = widths
                    s.cols = len(widths)
        s.title_pg = el.find(W + "titlePg") is not None and onoff(el.find(W + "titlePg"))
        t = el.find(W + "type")
        if t is not None:
            s.start = attr(t, "val", "nextPage")
        pn = el.find(W + "pgNumType")
        if pn is not None:
            if attr(pn, "start") is not None:
                s.pg_start = int(num(attr(pn, "start"), 1))
            s.pg_fmt = attr(pn, "fmt", "decimal")
        va = el.find(W + "vAlign")
        if va is not None:
            s.v_align = attr(va, "val", "top")
        pb = el.find(W + "pgBorders")
        if pb is not None:
            s.borders = parse_borders(pb)
        rels = self.pkg.rels(self.main)
        for tag, target in (("headerReference", s.headers), ("footerReference", s.footers)):
            for ref in el.findall(W + tag):
                rid = ref.get(R + "id")
                if rid in rels:
                    part = rels[rid][1]
                    root = self.pkg.xml(part)
                    if root is not None:
                        saved = self._fields
                        self._fields = []
                        target[attr(ref, "type", "default")] = list(self._blocks(root, part, None))
                        self._fields = saved
        return s

    def _blocks(self, parent, part, tstyle):
        """Yield blocks; drop caps become a frame floating at the start of the
        next paragraph."""
        dropcap = None
        for blk in self._joined_blocks(parent, part, tstyle):
            if dropcap is not None:
                if isinstance(blk, Paragraph):
                    blk.items.insert(0, _dropcap_float(dropcap))
                else:
                    yield dropcap
                dropcap = None
            fp = blk.ppr.get("framePr") if isinstance(blk, Paragraph) else None
            if fp and fp.get("dropCap") in ("drop", "margin"):
                dropcap = blk
                continue
            yield blk
        if dropcap is not None:
            yield dropcap

    def _joined_blocks(self, parent, part, tstyle):
        """Yield blocks; a paragraph whose mark is hidden (a style separator)
        runs on into the next paragraph, as Word lays it out."""
        pending = None
        for blk in self._raw_blocks(parent, part, tstyle):
            if pending is not None:
                if isinstance(blk, Paragraph):
                    blk = _join_paragraphs(pending, blk)
                elif any(not isinstance(i, Bookmark) for i in pending.items):
                    yield pending
                # an empty paragraph with a hidden mark before a table is not displayed
                pending = None
            if isinstance(blk, Paragraph) and blk.mark_hidden and blk.sect is None:
                pending = blk
                continue
            yield blk
        if pending is not None:
            yield pending

    def _raw_blocks(self, parent, part, tstyle):
        for el in parent:
            t = local(el)
            if t == "p":
                yield self.read_paragraph(el, part, tstyle)
            elif t == "tbl":
                yield self.read_table(el, part)
            elif t in ("sdt",):
                content = el.find(W + "sdtContent")
                if content is not None:
                    yield from self._raw_blocks(content, part, tstyle)
            elif t in ("customXml", "ins", "moveTo", "smartTag"):
                yield from self._raw_blocks(el, part, tstyle)
            elif t == "AlternateContent":
                chosen = _choose_alt(el)
                if chosen is not None:
                    yield from self._raw_blocks(chosen, part, tstyle)

    # ------------------------------------------------------------------ paragraph
    def read_paragraph(self, p, part, tstyle) -> Paragraph:
        ppr_el = p.find(W + "pPr")
        ppr, base_rpr, sid = self.para_props(ppr_el, tstyle)
        mark_rpr = merge(base_rpr, ppr.get("mark_rpr") or {})
        mark = self.run_style(mark_rpr)
        para = Paragraph(ppr=ppr, items=[], mark=mark, style_id=sid, in_table=tstyle is not None)
        para.mark_hidden = bool(mark_rpr.get("vanish"))
        num_id = ppr.get("numId")
        if num_id and num_id != "0":
            res = self.numbering.next_label(num_id, ppr.get("ilvl", 0))
            if res is not None:
                text, lvl = res
                lrpr = merge(mark_rpr, lvl.rpr)
                # Word does not underline / highlight labels from the paragraph mark
                para.label = (text, self.run_style(lrpr), lvl)
        ctx = {"part": part, "base": base_rpr, "link": None}
        self._inlines(p, para.items, ctx)
        # Word ignores whitespace that only leads into an equation at the start
        # of a paragraph
        k = 0
        while k < len(para.items) and isinstance(para.items[k], Text) and not para.items[k].text.strip():
            k += 1
        if 0 < k < len(para.items) and isinstance(para.items[k], MathItem):
            del para.items[:k]
        if ctx.get("display_jc"):
            # display equations: the paragraph's lines are centred like Word's math zones
            jc = ctx["display_jc"]
            para.ppr = dict(ppr, jc={"left": "left", "right": "right"}.get(jc, "center"))
        if ppr_el is not None:
            sp = ppr_el.find(W + "sectPr")
            if sp is not None:
                para.sect = self.read_sectpr(sp)
        return para

    def _suppressed(self):
        for f in self._fields:
            if f["phase"] == "instr" or f.get("suppress"):
                return True
        return False

    def _current_link(self, ctx):
        for f in reversed(self._fields):
            if f["phase"] == "result" and f.get("link"):
                return f["link"]
        return ctx["link"]

    def _inlines(self, parent, items, ctx):
        for el in parent:
            t = local(el)
            if t == "r":
                self._run(el, items, ctx)
            elif t == "hyperlink":
                rid = el.get(R + "id")
                link = None
                if rid:
                    rel = self.pkg.rels(ctx["part"]).get(rid)
                    if rel:
                        link = rel[1]
                anchor = attr(el, "anchor")
                if anchor and not link:
                    link = "#" + anchor
                sub = dict(ctx, link=link or ctx["link"])
                self._inlines(el, items, sub)
            elif t == "fldSimple":
                instr = attr(el, "instr", "")
                kind = _field_kind(instr)
                if kind in PAGE_FIELDS:
                    r = el.find(f".//{W}r")
                    rpr = merge(ctx["base"], self._run_rpr(r))
                    items.append(Field(kind, self.run_style(rpr), _field_fmt(instr),
                                       _text_of(el) or "1", _field_picture(instr)))
                else:
                    link = _field_link(instr)
                    sub = dict(ctx, link=link or ctx["link"])
                    self._inlines(el, items, sub)
            elif t in ("ins", "moveTo", "smartTag", "customXml", "dir", "bdo"):
                self._inlines(el, items, ctx)
            elif t == "sdt":
                content = el.find(W + "sdtContent")
                if content is not None:
                    self._inlines(content, items, ctx)
            elif t == "bookmarkStart":
                name = attr(el, "name")
                if name:
                    items.append(Bookmark(name))
            elif t == "AlternateContent":
                chosen = _choose_alt(el)
                if chosen is not None:
                    self._inlines(chosen, items, ctx)
            elif t in ("oMath", "oMathPara"):
                self._math(el, items, ctx)

    def _math(self, el, items, ctx):
        """Office Math: inline equations and display paragraphs (m:oMathPara),
        which Word centres with each equation on its own line."""
        style = self.run_style(ctx["base"])
        if local(el) == "oMath":
            items.append(MathItem(el, style, False))
            return
        first = True
        for c in el:
            if local(c) == "oMath":
                if not first:
                    items.append(Break("line", style))
                items.append(MathItem(c, style, True))
                first = False
            elif local(c) == "oMathParaPr":
                jc = c.find(M + "jc")
                ctx["display_jc"] = jc.get(M + "val") if jc is not None else "centerGroup"
        ctx.setdefault("display_jc", "centerGroup")

    def _run_rpr(self, r) -> dict:
        if r is None:
            return {}
        direct = parse_rpr(r.find(W + "rPr"))
        cs = direct.get("rStyle")
        if cs and cs.lower() == "hyperlink" and any(
                f["phase"] == "result" and f["instr"].strip().upper().startswith("TOC") for f in self._fields):
            # Word does not show the Hyperlink character style on the entries of a
            # table of contents (they carry it for the \h links)
            cs = None
        if cs:
            return merge(self.styles.char_style(cs), direct)
        return direct

    def _run(self, r, items, ctx):
        rpr_direct = self._run_rpr(r)
        rpr = merge(ctx["base"], rpr_direct)
        style = self.run_style(rpr)
        for el in r:
            t = local(el)
            if t == "fldChar":
                ft = attr(el, "fldCharType")
                if ft == "begin":
                    self._fields.append({"phase": "instr", "instr": ""})
                elif ft == "separate" and self._fields:
                    f = self._fields[-1]
                    f["phase"] = "result"
                    kind = _field_kind(f["instr"])
                    f["link"] = _field_link(f["instr"])
                    if kind in PAGE_FIELDS:
                        outer = self._fields[:-1]
                        if not any(x["phase"] == "instr" or x.get("suppress") for x in outer):
                            items.append(Field(kind, style, _field_fmt(f["instr"]), picture=_field_picture(f["instr"])))
                        f["suppress"] = True
                elif ft == "end" and self._fields:
                    f = self._fields.pop()
                    if f["phase"] == "instr":
                        kind = _field_kind(f["instr"])
                        if kind in PAGE_FIELDS and not self._suppressed():
                            items.append(Field(kind, style, _field_fmt(f["instr"]), picture=_field_picture(f["instr"])))
                continue
            if t == "instrText":
                if self._fields and self._fields[-1]["phase"] == "instr":
                    self._fields[-1]["instr"] += el.text or ""
                continue
            if self._suppressed() or style.hidden:
                continue
            link = self._current_link(ctx)
            if t == "t":
                if el.text:
                    items.append(Text(el.text, style, link))
            elif t == "tab":
                items.append(Tab(style, link))
            elif t == "ptab":
                items.append(Tab(style, link))
            elif t in ("br", "cr"):
                kind = attr(el, "type", "textWrapping")
                items.append(Break({"page": "page", "column": "column"}.get(kind, "line"), style))
            elif t == "noBreakHyphen":
                items.append(Text("\u2011", style, link))
            elif t == "softHyphen":
                items.append(Text("\u00ad", style, link))
            elif t == "sym":
                fam = attr(el, "font")
                ch = attr(el, "char", "0020")
                try:
                    c = chr(int(ch, 16))
                except ValueError:
                    continue
                if fam:
                    srpr = merge(rpr, {"font_ascii": fam, "font_hAnsi": fam,
                                       "font_eastAsia": fam, "font_cs": fam})
                    items.append(Text(c, self.run_style(srpr), link))
                else:
                    items.append(Text(c, style, link))
            elif t == "drawing":
                self._drawing(el, items, ctx, style, link)
            elif t in ("pict", "object"):
                self._vml(el, items, ctx, style, link)
            elif t == "AlternateContent":
                chosen = _choose_alt(el)
                if chosen is not None:
                    self._run_children(chosen, items, ctx, style, link)
            elif t in ("footnoteReference", "endnoteReference"):
                nid = attr(el, "id")
                kind = "footnote" if t == "footnoteReference" else "endnote"
                if kind == "footnote":
                    self._fn_counter += 1
                    label = str(self._fn_counter)
                else:
                    self._en_counter += 1
                    label = format_number(self._en_counter, "lowerRoman")
                    self.endnote_order.append((label, nid))
                if attr(el, "customMarkFollows") in ("1", "true", "on"):
                    label = ""
                items.append(FootnoteRef(nid, label, style, kind))
            elif t in ("footnoteRef", "endnoteRef"):
                items.append(FootnoteRef("__self__", "", style))
            elif t in ("separator", "continuationSeparator"):
                items.append(Separator(style))
            elif t == "lastRenderedPageBreak":
                pass

    def _run_children(self, parent, items, ctx, style, link):
        for el in parent:
            t = local(el)
            if t == "drawing":
                self._drawing(el, items, ctx, style, link)
            elif t in ("pict", "object"):
                self._vml(el, items, ctx, style, link)

    # ------------------------------------------------------------------ drawings
    def _blob(self, part, rid):
        if not rid:
            return None
        rel = self.pkg.rels(part).get(rid)
        if rel is None or rel[2]:
            return None
        return self.pkg.read(rel[1])

    def _drawing(self, el, items, ctx, style, link):
        for obj in el:
            t = local(obj)
            if t not in ("inline", "anchor"):
                continue
            ext = obj.find(WP + "extent")
            w = num(attr_plain(ext, "cx"), 0) / EMU
            h = num(attr_plain(ext, "cy"), 0) / EMU
            ee = obj.find(WP + "effectExtent")
            extent = (0.0, 0.0, 0.0, 0.0)
            if ee is not None:
                extent = tuple(num(attr_plain(ee, k), 0) / EMU for k in ("l", "t", "r", "b"))
            gd = obj.find(f"{A}graphic/{A}graphicData")
            pic = shape = None
            if gd is not None:
                p = gd.find(PIC + "pic")
                wsp = gd.find(WPS + "wsp")
                if p is not None:
                    pic = self._pic(p, ctx["part"], w, h, style, link, extent)
                elif wsp is not None:
                    shape = self._wsp(wsp, ctx["part"], w, h)
                else:
                    grp = gd.find(WPG + "wgp")
                    if grp is not None:
                        shape = self._group(grp, ctx["part"], w, h)
                    elif gd.get("uri", "").endswith("/diagram"):
                        from .diagram import diagram_shape
                        shape = diagram_shape(self, ctx["part"], gd, w, h)
            if pic is None and shape is None:
                continue
            if t == "inline":
                if pic is None:
                    pic = Picture(None, w, h, style, link, extent=extent, shape=shape)
                items.append(pic)
            else:
                items.append(self._floating(obj, w, h, pic, shape))

    def _group(self, grp, part, w, h) -> Shape:
        """A DrawingML group: children are placed through the group's child
        coordinate space (chOff/chExt) scaled onto its extent."""
        g = Shape(w, h, geom="group")
        xfrm = grp.find(f"{WPG}grpSpPr/{A}xfrm")
        if xfrm is None:
            xfrm = grp.find(f"{A}grpSpPr/{A}xfrm")
        cox = coy = 0.0
        sx = sy = 1.0
        if xfrm is not None:
            ext = xfrm.find(A + "ext")
            cho = xfrm.find(A + "chOff")
            che = xfrm.find(A + "chExt")
            if cho is not None:
                cox, coy = num(cho.get("x"), 0), num(cho.get("y"), 0)
            if ext is not None and che is not None:
                cw, ch = num(che.get("cx"), 0), num(che.get("cy"), 0)
                if cw:
                    sx = num(ext.get("cx"), 0) / cw
                if ch:
                    sy = num(ext.get("cy"), 0) / ch

        def place(sp_pr):
            x = sp_pr.find(A + "xfrm") if sp_pr is not None else None
            if x is None:
                return 0.0, 0.0, 0.0, 0.0
            off, ext = x.find(A + "off"), x.find(A + "ext")
            ox = (num(off.get("x"), 0) - cox) * sx / EMU if off is not None else 0.0
            oy = (num(off.get("y"), 0) - coy) * sy / EMU if off is not None else 0.0
            cw = num(ext.get("cx"), 0) * sx / EMU if ext is not None else 0.0
            ch = num(ext.get("cy"), 0) * sy / EMU if ext is not None else 0.0
            return ox, oy, cw, ch
        for c in grp:
            n = local(c)
            if n == "wsp":
                ox, oy, cw, ch = place(c.find(WPS + "spPr"))
                g.children.append((ox, oy, self._wsp(c, part, cw, ch)))
            elif n == "pic":
                ox, oy, cw, ch = place(c.find(PIC + "spPr"))
                pic = self._pic(c, part, cw, ch, None, None, (0.0, 0.0, 0.0, 0.0))
                g.children.append((ox, oy, pic))
            elif n == "grpSp":
                gx = c.find(f"{WPG}grpSpPr/{A}xfrm")
                ox = oy = cw = ch = 0.0
                if gx is not None:
                    off, ext = gx.find(A + "off"), gx.find(A + "ext")
                    ox = (num(off.get("x"), 0) - cox) * sx / EMU
                    oy = (num(off.get("y"), 0) - coy) * sy / EMU
                    cw = num(ext.get("cx"), 0) * sx / EMU
                    ch = num(ext.get("cy"), 0) * sy / EMU
                g.children.append((ox, oy, self._group(c, part, cw, ch)))
        return g

    def _pic(self, p, part, w, h, style, link, extent):
        blip = p.find(f".//{A}blip")
        rid = blip.get(R + "embed") if blip is not None else None
        data = self._blob(part, rid)
        crop = None
        src = p.find(f".//{A}srcRect")
        if src is not None:
            crop = tuple(num(src.get(k), 0) / 100000.0 for k in ("l", "t", "r", "b"))
        return Picture(data, w, h, style, link, crop, extent)

    def _wsp(self, wsp, part, w, h) -> Shape:
        sp = wsp.find(WPS + "spPr")
        shape = Shape(w, h)
        if sp is not None and sp.find(A + "xfrm") is not None and not (w and h):
            ext = sp.find(f"{A}xfrm/{A}ext")
            if ext is not None:
                shape.width = num(ext.get("cx"), 0) / EMU
                shape.height = num(ext.get("cy"), 0) / EMU
        if sp is not None:
            geom = sp.find(A + "prstGeom")
            if geom is not None:
                shape.geom = geom.get("prst", "rect")
            if sp.find(A + "noFill") is not None:
                shape.fill = None
            else:
                shape.fill = drawingml_color(sp.find(A + "solidFill"), self.theme)
            blip = sp.find(f"{A}blipFill/{A}blip")
            if blip is not None:
                # picture fill, stretched over the shape
                shape.data = self._blob(part, blip.get(R + "embed") or blip.get(R + "link"))
            ln = sp.find(A + "ln")
            if ln is not None:
                if ln.find(A + "noFill") is not None:
                    shape.line = None
                else:
                    shape.line = drawingml_color(ln.find(A + "solidFill"), self.theme)
                if ln.get("w"):
                    shape.line_width = num(ln.get("w"), 9525) / EMU
            xfrm = sp.find(A + "xfrm")
            if xfrm is not None and xfrm.get("rot"):
                shape.rotation = num(xfrm.get("rot"), 0) / 60000.0
        style_el = wsp.find(WPS + "style")
        if style_el is not None and sp is not None:
            if sp.find(A + "solidFill") is None and sp.find(A + "noFill") is None \
                    and sp.find(A + "blipFill") is None:
                fr = style_el.find(A + "fillRef")
                if fr is not None and fr.get("idx") not in ("0", None):
                    shape.fill = drawingml_color(fr, self.theme)
            ln = sp.find(A + "ln")
            if ln is None or (ln.find(A + "solidFill") is None and ln.find(A + "noFill") is None):
                lr = style_el.find(A + "lnRef")
                if lr is not None and lr.get("idx") not in ("0", None):
                    shape.line = drawingml_color(lr, self.theme)
        body = wsp.find(WPS + "bodyPr")
        if body is not None and body.find(A + "spAutoFit") is not None:
            shape.autofit = True
        if body is not None:
            shape.insets = tuple(num(body.get(k), d) / EMU for k, d in
                                 (("lIns", 91440), ("tIns", 45720), ("rIns", 91440), ("bIns", 45720)))
            shape.v_anchor = body.get("anchor", "t")
            if body.get("vert") in ("vert", "vert270", "eaVert", "wordArtVert"):
                shape.vert = "vert270" if body.get("vert") == "vert270" else "vert"
        tx = wsp.find(f"{WPS}txbx/{W}txbxContent")
        if tx is not None:
            saved = self._fields
            self._fields = []
            shape.blocks = list(self._blocks(tx, part, None))
            self._fields = saved
        return shape

    def _floating(self, obj, w, h, pic, shape) -> Floating:
        # positions/sizes may sit inside mc:AlternateContent (wp14 percentages)
        flat = []
        for c in obj:
            if local(c) == "AlternateContent":
                chosen = _choose_alt(c)
                if chosen is not None:
                    flat.extend(chosen)
            else:
                flat.append(c)

        def child(name):
            for c in flat:
                if local(c) == name:
                    return c
            return None

        pct = {}

        def pos(tag):
            el = child(tag)
            if el is None:
                return ("column" if tag == "positionH" else "paragraph", None, 0.0)
            rel = el.get("relativeFrom", "column")
            al = el.find(WP + "align")
            off = el.find(WP + "posOffset")
            for c in el:
                if local(c) in ("pctPosHOffset", "pctPosVOffset") and c.text:
                    pct[tag] = num(c.text, 0) / 100000.0
            return (rel, al.text.strip() if al is not None and al.text else None,
                    num(off.text if off is not None else "0", 0) / EMU)
        hr, ha, ho = pos("positionH")
        vr, va, vo = pos("positionV")
        size_pct = {}
        for name, key in (("sizeRelH", "w"), ("sizeRelV", "h")):
            el = child(name)
            if el is not None:
                for c in el:
                    if local(c) in ("pctWidth", "pctHeight") and c.text and num(c.text, 0) > 0:
                        size_pct[key] = (el.get("relativeFrom", "margin"), num(c.text, 0) / 100000.0)
        if obj.get("simplePos") in ("1", "true"):
            sp = obj.find(WP + "simplePos")
            if sp is not None:
                hr, ha, ho = "page", None, num(sp.get("x"), 0) / EMU
                vr, va, vo = "page", None, num(sp.get("y"), 0) / EMU
        wrap = "none"
        for c in obj:
            n = local(c)
            if n.startswith("wrap"):
                wrap = {"wrapNone": "none", "wrapSquare": "square", "wrapTight": "tight",
                        "wrapThrough": "through", "wrapTopAndBottom": "topAndBottom"}.get(n, "none")
        dist = tuple(num(obj.get(k), 0) / EMU for k in ("distT", "distB", "distL", "distR"))
        fl = Floating(w, h, hr, ha, ho, vr, va, vo, wrap,
                      obj.get("behindDoc") in ("1", "true"), dist, pic, shape,
                      int(num(obj.get("relativeHeight"), 0)))
        fl.h_pct = pct.get("positionH")
        fl.v_pct = pct.get("positionV")
        fl.size_pct = size_pct
        return fl

    def _text_blocks(self, content, part):
        """Blocks of a text box body, read with a fresh field state."""
        saved = self._fields
        self._fields = []
        blocks = list(self._blocks(content, part, None))
        self._fields = saved
        return blocks

    def _vml(self, el, items, ctx, style, link):
        from .vml import VmlReader, css, length
        group = el.find(V + "group")
        if group is not None:
            st = css(group.get("style", ""))
            w, h = length(st.get("width")), length(st.get("height"))
            if w and h:
                shape = VmlReader(self, ctx["part"]).group(group, w, h)
                if st.get("position") == "absolute":
                    items.append(Floating(w, h, "column", None, length(st.get("margin-left")),
                                          "paragraph", None, length(st.get("margin-top")),
                                          "square", False, (0, 0, 0, 0), None, shape))
                else:
                    items.append(Picture(None, w, h, style, link, shape=shape))
                return
        for shape in el.iter(V + "shape", V + "rect", V + "roundrect"):
            css = _parse_css(shape.get("style", ""))
            w = _css_len(css.get("width"))
            h = _css_len(css.get("height"))
            img = shape.find(V + "imagedata")
            if img is not None:
                rid = img.get(R + "id") or img.get("{urn:schemas-microsoft-com:office:office}relid")
                data = self._blob(ctx["part"], rid)
                if data and w and h:
                    items.append(Picture(data, w, h, style, link))
                    return
            tb = shape.find(f"{V}textbox/{W}txbxContent")
            if tb is not None and w and h:
                saved = self._fields
                self._fields = []
                s = Shape(w, h, fill=None, line=(0, 0, 0))
                fill = shape.get("fillcolor")
                if fill:
                    s.fill = hex_color(fill.split()[0].lstrip("#")) if fill.startswith("#") else None
                s.blocks = list(self._blocks(tb, ctx["part"], None))
                self._fields = saved
                if css.get("position") == "absolute":
                    items.append(Floating(w, h, "column", None, _css_len(css.get("margin-left")) or 0,
                                          "paragraph", None, _css_len(css.get("margin-top")) or 0,
                                          "square", False, (0, 0, 0, 0), None, s))
                else:
                    items.append(Picture(None, w, h, style, link, shape=s))
                return

    # ------------------------------------------------------------------ tables
    def read_table(self, tbl, part) -> Table:
        direct = parse_tblpr(tbl.find(W + "tblPr"))
        tstyle = self.styles.table_style(direct.get("tblStyle"))
        tblpr = merge(tstyle.tblpr, direct)
        look = tblpr.get("look") or {}
        grid = [twips(attr(g, "w"), 0) for g in tbl.findall(f"{W}tblGrid/{W}gridCol")]
        raw_rows = []
        for tr in _iter_children(tbl, "tr"):
            trpr = merge(tstyle.trpr, parse_trpr(tr.find(W + "trPr")))
            # table property exceptions for this row (cell margins, borders...)
            ex = parse_tblpr(tr.find(W + "tblPrEx"))
            if ex:
                trpr = dict(trpr, tblPrEx=ex)
            raw_rows.append((tr, trpr, list(_iter_children(tr, "tc"))))
        nrows = len(raw_rows)
        ncols = len(grid) or max((sum(int(num(attr(tc.find(f"{W}tcPr/{W}gridSpan"), "val"), 1) or 1)
                                      for tc in cells) for _, _, cells in raw_rows), default=1)
        rows = []
        row_band = max(1, tblpr.get("rowBand", 1))
        col_band = max(1, tblpr.get("colBand", 1))
        for ri, (tr, trpr, tcs) in enumerate(raw_rows):
            col = trpr.get("gridBefore", 0)
            cells = []
            for tc in tcs:
                tcpr_direct = parse_tcpr(tc.find(W + "tcPr"))
                span = max(1, tcpr_direct.get("gridSpan", 1))
                conds = ["wholeTable"]
                is_first_row = look.get("firstRow") and ri == 0
                is_last_row = look.get("lastRow") and ri == nrows - 1
                if not look.get("noHBand") and not is_first_row and not is_last_row:
                    rb = (ri - (1 if look.get("firstRow") else 0)) // row_band
                    conds.append("band1Horz" if rb % 2 == 0 else "band2Horz")
                is_first_col = look.get("firstColumn") and col == 0
                is_last_col = look.get("lastColumn") and col + span >= ncols
                if not look.get("noVBand") and not is_first_col and not is_last_col:
                    cb = (col - (1 if look.get("firstColumn") else 0)) // col_band
                    conds.append("band1Vert" if cb % 2 == 0 else "band2Vert")
                if is_first_col:
                    conds.append("firstCol")
                if is_last_col:
                    conds.append("lastCol")
                if is_first_row:
                    conds.append("firstRow")
                if is_last_row:
                    conds.append("lastRow")
                if is_first_row and is_first_col:
                    conds.append("nwCell")
                if is_first_row and is_last_col:
                    conds.append("neCell")
                if is_last_row and is_first_col:
                    conds.append("swCell")
                if is_last_row and is_last_col:
                    conds.append("seCell")
                c_ppr, c_rpr, c_tcpr = tstyle.ppr, tstyle.rpr, tstyle.tcpr
                for cname in conds:
                    cd = tstyle.cond.get(cname)
                    if cd:
                        c_ppr = merge(c_ppr, cd["ppr"])
                        c_rpr = merge(c_rpr, cd["rpr"])
                        c_tcpr = merge(c_tcpr, _cond_edges(cname, cd["tcpr"], ri, nrows, col, span, ncols))
                tcpr = merge(c_tcpr, tcpr_direct)
                blocks = list(self._blocks(tc, part, (c_ppr, c_rpr)))
                vm = tcpr_direct.get("vMerge")
                cells.append(Cell(tcpr, blocks, col, span, vm if vm else None))
                col += span
            rows.append(Row(trpr, cells))
        # compute vertical spans
        for ri, row in enumerate(rows):
            for cell in row.cells:
                if cell.vmerge == "restart":
                    n = 1
                    for rj in range(ri + 1, len(rows)):
                        nxt = next((c for c in rows[rj].cells if c.col == cell.col), None)
                        if nxt is not None and nxt.vmerge == "continue":
                            n += 1
                        else:
                            break
                    cell.row_span = n
        if not grid:
            grid = [72.0] * ncols
        elif len(grid) < ncols:
            grid = grid + [grid[-1]] * (ncols - len(grid))
        return Table(tblpr, grid, rows, compat_legacy=self.compat_mode < 15,
                     float=direct.get("float"))

    # ------------------------------------------------------------------ notes
    def _load_notes(self):
        for part, kind in ((self.footnote_part, "footnote"), (self.endnote_part, "endnote")):
            if not part:
                continue
            root = self.pkg.xml(part)
            if root is None:
                continue
            for fn in root:
                if local(fn) not in ("footnote", "endnote"):
                    continue
                fid = attr(fn, "id")
                ftype = attr(fn, "type", "normal")
                saved = self._fields
                self._fields = []
                blocks = list(self._blocks(fn, part, None))
                self._fields = saved
                if ftype == "separator" and kind == "footnote":
                    self.footnote_sep = blocks
                elif ftype == "separator" and kind == "endnote":
                    self.endnote_sep = blocks
                elif ftype in (None, "normal"):
                    self.footnotes[(kind, fid)] = blocks


def _join_paragraphs(first: Paragraph, second: Paragraph) -> Paragraph:
    """Merge a paragraph with a hidden mark into the following one: the line
    keeps the first paragraph's formatting and continues with the second's runs."""
    first.items = first.items + second.items
    first.mark = second.mark
    first.sect = second.sect
    first.mark_hidden = second.mark_hidden
    return first


def attr_plain(el, name, default=None):
    if el is None:
        return default
    v = el.get(name)
    return default if v is None else v


def _freeze(v):
    if isinstance(v, dict):
        return tuple(sorted((k, _freeze(x)) for k, x in v.items()))
    if isinstance(v, list):
        return tuple(_freeze(x) for x in v)
    return v


def _iter_children(el, name):
    for c in el:
        t = local(c)
        if t == name:
            yield c
        elif t in ("sdt",):
            content = c.find(W + "sdtContent")
            if content is not None:
                yield from _iter_children(content, name)
        elif t in ("customXml", "ins", "moveTo"):
            yield from _iter_children(c, name)


_ROW_CONDS = {"firstRow", "lastRow", "band1Horz", "band2Horz"}
_COL_CONDS = {"firstCol", "lastCol", "band1Vert", "band2Vert"}


def _cond_edges(cname, tcpr, ri, nrows, col, span, ncols):
    """A table style condition's cell borders describe its whole region: for a
    row region (header row, bands) left/right are the row's outer edges and
    insideV the edges between its cells; for a column region top/bottom and
    insideH likewise (table 4 of sample-files table-document: a header row
    without dividers)."""
    tb = tcpr.get("tcBorders")
    if not tb:
        return tcpr
    b = dict(tb)

    def inner(side, inside):
        # an interior edge takes the region's inside border (nil = none); when the
        # region has none, the table's own border for that edge applies
        if inside in tb:
            b[side] = tb[inside]
        else:
            b.pop(side, None)
    if cname in _ROW_CONDS:
        if col > 0:
            inner("left", "insideV")
        if col + span < ncols:
            inner("right", "insideV")
    elif cname in _COL_CONDS:
        if ri > 0:
            inner("top", "insideH")
        if ri < nrows - 1:
            inner("bottom", "insideH")
    else:
        return tcpr
    return dict(tcpr, tcBorders=b)


def _dropcap_float(para: Paragraph) -> Floating:
    """A drop cap paragraph as a text frame at the start of the next paragraph;
    its size is measured at layout time (Floating.dropcap)."""
    fp = para.ppr.get("framePr") or {}
    # the frame holds just the letters; the paragraph's indent places the frame
    indent = (para.ppr.get("ind_left") or 0.0) + (para.ppr.get("ind_first") or 0.0)
    para.ppr = {k: v for k, v in para.ppr.items() if k != "framePr"}
    para.ppr.update(ind_left=0.0, ind_right=0.0, ind_first=0.0, sp_before=0.0, sp_after=0.0)
    shape = Shape(0.0, 0.0, fill=None, line=None)
    shape.blocks = [para]
    shape.insets = (0.0, 0.0, 0.0, 0.0)
    h_space = num(fp.get("hSpace"), 0) / 20.0
    fl = Floating(0.0, 0.0, "column", None, max(0.0, indent), "paragraph", None, 0.0, "square", False,
                  (0.0, 0.0, 0.0, h_space), None, shape)
    fl.dropcap = max(1, int(num(fp.get("lines"), 3)))
    fl.h_align = "outside_margin" if fp.get("dropCap") == "margin" else None
    return fl


def _choose_alt(el):
    for c in el:
        if local(c) == "Choice":
            req = set((c.get("Requires") or "").split())
            if req <= _MC_OK:
                return c
    for c in el:
        if local(c) == "Fallback":
            return c
    return None


def _field_kind(instr: str) -> str:
    parts = instr.strip().split()
    return parts[0].upper() if parts else ""


def _field_fmt(instr: str):
    m = re.search(r"\\\*\s*(\w+)", instr)
    return m.group(1) if m else None


def _field_picture(instr: str):
    m = re.search(r'\#\s*(?:"([^"]*)"|(\S+))', instr)
    return (m.group(1) if m.group(1) is not None else m.group(2)) if m else None


def _field_link(instr: str):
    parts = instr.strip().split(None, 1)
    if not parts:
        return None
    kind = parts[0].upper()
    if kind == "HYPERLINK" and len(parts) > 1:
        rest = parts[1]
        m = re.search(r'\\l\s+"([^"]*)"', rest)
        if m:
            return "#" + m.group(1)
        m = re.match(r'\s*"([^"]*)"', rest) or re.match(r"\s*(\S+)", rest)
        return m.group(1) if m else None
    if kind in ("REF", "PAGEREF") and len(parts) > 1 and "\\h" in parts[1]:
        return "#" + parts[1].split()[0]
    return None


def _text_of(el):
    return "".join(t.text or "" for t in el.iter(W + "t"))


def _parse_css(s: str) -> dict:
    out = {}
    for part in s.split(";"):
        if ":" in part:
            k, v = part.split(":", 1)
            out[k.strip().lower()] = v.strip()
    return out


def _css_len(v):
    if not v:
        return None
    m = re.match(r"(-?[\d.]+)\s*(pt|in|cm|mm|px|emu)?", v)
    if not m:
        return None
    n = float(m.group(1))
    u = m.group(2) or "emu"
    return n * {"pt": 1, "in": 72, "cm": 28.3465, "mm": 2.83465, "px": 0.75, "emu": 1 / EMU}[u]
