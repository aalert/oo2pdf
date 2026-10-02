"""Pagination of laid-out flows into pages, Word style."""
from __future__ import annotations

from dataclasses import dataclass, field, replace

from ..draw import ImageOp, LineOp, move_ops
from .layout import (EPS, Line, LayoutContext, ParaFlow, RowFlow, TableFlow, _border_line,
                     gap, layout_blocks, render_shape, stack_flows)
from .model import Floating, Section
from .numbering import format_number


@dataclass
class Page:
    section: Section
    width: float
    height: float
    number: int  # displayed page number
    index: int  # physical index
    section_first: bool
    body_ops: list = field(default_factory=list)
    back_ops: list = field(default_factory=list)
    front_ops: list = field(default_factory=list)
    note_ops: list = field(default_factory=list)
    hf_ops: list = field(default_factory=list)
    section_index: int = 0
    content_bottom: float = 0.0
    exclusions: list = field(default_factory=list)  # (x0, y0, x1, y1)


FOOTER_OVERLAP = 0.25


class _Overflow(Exception):
    """Raised while balancing columns when a trial layout needs another page."""


class Paginator:
    def __init__(self, ctx: LayoutContext, reader):
        self.ctx = ctx
        self.reader = reader
        self.pages: list[Page] = []
        # wrapping floats: Word lays out the whole page around them, so even lines
        # above their anchor paragraph (or above a paragraph-anchored picture
        # whose wrap distance reaches back) flow around them. One pass records
        # where each float lands (anchor_pages: id -> (page, float, x, y)); the
        # next places those floats as soon as their page starts
        # (prefloats: page index -> [(float, x, y)]).
        self.anchor_pages: dict[int, tuple[int, Floating, float, float]] = {}
        self.prefloats: dict[int, list] = {}
        self._preplaced: set[int] = set()
        self.page: Page | None = None
        self.sect: Section | None = None
        self.sect_index = 0
        self.cols: list[tuple[float, float]] = []
        self.col = 0
        self.y = 0.0
        self.top = 0.0
        self.bottom = 0.0
        self.col_top = 0.0
        self.at_top = True  # nothing placed in the current column yet
        # why the current column started: natural | break | pbb | first | section
        self.top_mode = "first"
        self.page_no = 0
        self.notes_h = 0.0
        self.notes: list = []
        self.first_page_of_doc = True
        self.carry_after = 0.0
        self.balance: tuple[int, float] | None = None  # (page index, column bottom)
        self.trial_pages: int | None = None
        self.section_page_count: dict[int, int] = {}

    # ----------------------------------------------------------------- geometry
    def _header_blocks(self, sect: Section, section_first: bool, number: int, kind_map):
        if section_first and sect.title_pg:
            return kind_map.get("first")
        if self.reader.even_odd and number % 2 == 0:
            return kind_map.get("even")
        return kind_map.get("default")

    def _hf_height(self, blocks, width):
        if not blocks:
            return 0.0, []
        flows = layout_blocks(blocks, width, self.ctx)
        lines, h = stack_flows(flows, self.ctx.sum_spacing)
        return h, lines

    def _columns(self, sect: Section):
        body_w = sect.page_w - sect.margin_left - sect.margin_right - sect.gutter
        x0 = sect.margin_left + sect.gutter
        if sect.col_widths:
            cols = []
            x = x0
            for w, space in sect.col_widths:
                cols.append((x, w))
                x += w + space
            return cols
        n = max(1, sect.cols)
        w = (body_w - (n - 1) * sect.col_space) / n
        return [(x0 + i * (w + sect.col_space), w) for i in range(n)]

    def new_page(self, mode="natural"):
        if self.trial_pages is not None and len(self.pages) >= self.trial_pages:
            raise _Overflow()
        sect = self.sect
        if self.page is not None:
            self._finish_page()
        self.page_no += 1
        first_in_section = self._section_first
        if first_in_section and sect.pg_start is not None:
            self.page_no = sect.pg_start
        self._section_first = False
        page = Page(sect, sect.page_w, sect.page_h, self.page_no, len(self.pages), first_in_section,
                    section_index=self.sect_index)
        self.pages.append(page)
        self.page = page
        self.ctx.fields = {"PAGE": format_number(self.page_no, sect.pg_fmt), "NUMPAGES": "1",
                           "SECTIONPAGES": "1"}
        hw = sect.page_w - sect.margin_left - sect.margin_right
        hb = self._header_blocks(sect, first_in_section, self.page_no, sect.headers)
        fb = self._header_blocks(sect, first_in_section, self.page_no, sect.footers)
        hh, hlines = self._hf_height(hb, hw)
        fh, flines = self._hf_height(fb, hw)
        top = abs(sect.margin_top)
        if sect.margin_top >= 0:
            top = max(top, sect.header_dist + hh)
        bottom = sect.page_h - abs(sect.margin_bottom)
        if sect.margin_bottom >= 0:
            # a footer taller than the bottom margin pushes the body up; Word lets
            # the body reach a little into it (StudentReport.docx: 0.1 pt)
            bottom = min(bottom, sect.page_h - sect.footer_dist - fh + FOOTER_OVERLAP)
        self.top = top
        self.bottom = bottom
        self.cols = self._columns(sect)
        self.col = 0
        self.col_top = top
        self.y = top
        self.at_top = True
        self.top_mode = "first" if self.first_page_of_doc else mode
        self.first_page_of_doc = False
        self.notes = []
        self.notes_h = 0.0
        for fl, fx, fy in self.prefloats.get(page.index, ()):
            self._place_float(fl, sect.margin_left, top, top, pre=(fx, fy))
        # body text also wraps around objects anchored in the header/footer
        for lines, y0 in ((hlines, sect.header_dist), (flines, sect.page_h - sect.footer_dist - fh)):
            for ly, ln in lines:
                for fl, fx in ln.floats:
                    if fl.wrap in ("square", "tight", "through", "topAndBottom"):
                        x, y = _hf_float_pos(fl, sect.margin_left + fx, y0 + ly, sect, self.ctx)
                        dt, db, dl, dr = fl.dist
                        page.exclusions.append((x - dl, y - dt, x + fl.width + dr, y + fl.height + db,
                                                "jump" if fl.wrap == "topAndBottom" else "wrap"))

    def next_column(self, mode="natural"):
        if self.col + 1 < len(self.cols):
            self.col += 1
            self.y = self.col_top
            self.at_top = True
            self.top_mode = mode
        else:
            self.new_page(mode)

    @property
    def col_x(self):
        return self.cols[self.col][0]

    @property
    def col_w(self):
        return self.cols[self.col][1]

    @property
    def limit(self):
        lim = self.bottom - self.notes_h
        if self.page is not None:
            x0, x1 = self.col_x, self.col_x + self.col_w
            for ex in self.page.exclusions:
                if ex[3] >= lim - 1 and ex[1] >= self.y and ex[0] <= x0 + 1 and ex[2] >= x1 - 1:
                    lim = min(lim, ex[1])
        if self.balance is not None and self.page is not None and self.page.index == self.balance[0]:
            return min(lim, self.balance[1])
        return lim

    # ----------------------------------------------------------------- sections
    def run(self, sections: list[Section], endnotes: list | None = None):
        self._section_first = True
        for si, sect in enumerate(sections):
            self.sect_index = si
            prev = self.sect
            self.sect = sect
            self._section_first = True
            cols = self._columns(sect)
            flows = layout_blocks(sect.blocks, cols[0][1], self.ctx)
            if si == len(sections) - 1 and endnotes:
                # the endnotes follow the text after the endnote separator line
                if self.reader.endnote_sep:
                    flows += layout_blocks(self.reader.endnote_sep, cols[0][1], self.ctx)
                for label, blocks in endnotes:
                    self.ctx.note_label = label
                    flows += layout_blocks(blocks, cols[0][1], self.ctx)
                self.ctx.note_label = ""
            continuous = (sect.start == "continuous" and self.page is not None and prev is not None
                          and abs(prev.page_w - sect.page_w) < 1 and abs(prev.page_h - sect.page_h) < 1)
            if self.page is None:
                self.new_page("first")
            elif continuous:
                # new column set starts below the content on the current page
                self._section_first = False
                # headers/footers stay those of the section that started the page
                self.cols = cols
                self.col = 0
                self.col_top = max(self.y, self.page.content_bottom)
                # the spacing between the sections applies to every column
                first = flows[0] if flows else None
                if isinstance(first, ParaFlow) and first.lines:
                    self.col_top += gap(self.carry_after, first.space_before, self.ctx.sum_spacing)
                    self.carry_after = 0.0
                    self.at_top = True
                    self.top_mode = "natural"
                self.y = self.col_top
            else:
                self.new_page("section")
                if sect.start in ("evenPage", "oddPage"):
                    want_even = sect.start == "evenPage"
                    if (self.page_no % 2 == 0) != want_even:
                        self.new_page("section")
            nxt = sections[si + 1] if si + 1 < len(sections) else None
            if len(cols) > 1 and nxt is not None and nxt.start == "continuous":
                self._flow_balanced(flows)
            else:
                self._flow(flows)
        if self.page is not None:
            self._finish_page()
        return self.pages

    # ----------------------------------------------------------------- balancing
    _SCALARS = ("y", "col", "cols", "col_top", "top", "bottom", "at_top", "top_mode", "notes_h",
                "carry_after", "page_no", "_section_first", "first_page_of_doc", "sect_index", "sect")

    def _snapshot(self):
        p = self.page
        return {
            "npages": len(self.pages), "page": p,
            "lens": (len(p.body_ops), len(p.back_ops), len(p.front_ops), len(p.note_ops),
                     len(p.exclusions)),
            "content_bottom": p.content_bottom, "section": p.section,
            "scalars": {k: getattr(self, k) for k in self._SCALARS},
            "notes": list(self.notes), "fields": dict(self.ctx.fields),
        }

    def _restore(self, s):
        del self.pages[s["npages"]:]
        p = self.page = s["page"]
        b, bk, fr, nt, ex = s["lens"]
        del p.body_ops[b:]
        del p.back_ops[bk:]
        del p.front_ops[fr:]
        del p.note_ops[nt:]
        del p.exclusions[ex:]
        p.content_bottom = s["content_bottom"]
        p.section = s["section"]
        for k, v in s["scalars"].items():
            setattr(self, k, v)
        self.notes = list(s["notes"])
        self.ctx.fields = dict(s["fields"])

    def _flow_balanced(self, flows):
        """Lay out a multi-column section ending in a continuous break with its
        last page's columns balanced, as Word does."""
        start = self._snapshot()
        self._flow(flows)
        last = len(self.pages) - 1
        top = self.col_top
        lo, hi = 0.0, self.limit - top
        if hi <= 1 or len(self.cols) < 2:
            return
        self.trial_pages = last + 1
        try:
            while hi - lo > 0.25:
                mid = (lo + hi) / 2
                self._restore(start)
                self.balance = (last, top + mid)
                try:
                    self._flow(flows)
                    ok = len(self.pages) == last + 1
                except _Overflow:
                    ok = False
                if ok:
                    hi = mid
                else:
                    lo = mid
        finally:
            self.trial_pages = None
        self._restore(start)
        self.balance = (last, top + hi)
        self._flow(flows)
        self.balance = None

    def _flow(self, flows):
        i = 0
        n = len(flows)
        prev_after = self.carry_after
        while i < n:
            f = flows[i]
            if isinstance(f, ParaFlow):
                prev_after = self._place_para(flows, i, prev_after)
            elif f.float:
                self._place_float_table(f)
            else:
                self.y += prev_after
                prev_after = 0.0
                self._place_table(f)
            i += 1
        self.carry_after = prev_after
        self.page.content_bottom = max(self.page.content_bottom, self.y)

    # ----------------------------------------------------------------- paragraphs
    def _space_before(self, f: ParaFlow, prev_after):
        """Space above a paragraph's first line, following Word's page-top rules."""
        if self.at_top:
            mode = self.top_mode
            if mode in ("natural", "break"):
                return 0.0  # Word drops space before at the top of a new page
            if mode in ("pbb", "section"):
                # page-break-before keeps what is left after the previous paragraph's
                # space after (which stayed on the previous page)
                if self.ctx.sum_spacing:
                    return f.space_before
                return max(0.0, f.space_before - prev_after)
            return f.space_before
        return gap(prev_after, f.space_before, self.ctx.sum_spacing)

    def _para_height(self, f: ParaFlow, nlines=None):
        lines = f.lines if nlines is None else f.lines[:nlines]
        return sum(ln.height for ln in lines)

    def _keep_chain_height(self, flows, i):
        """Height needed to honour keep-with-next starting at flows[i]."""
        total = 0.0
        j = i
        prev_after = 0.0
        while j < len(flows):
            f = flows[j]
            if not isinstance(f, ParaFlow):
                if isinstance(f, TableFlow) and f.rows:
                    total += prev_after + f.rows[0].height
                break
            sb = gap(prev_after, f.space_before, self.ctx.sum_spacing) if j > i else 0.0
            if f.keep_next and j + 1 < len(flows):
                total += sb + self._para_height(f)
                prev_after = f.space_after
                j += 1
                continue
            # last paragraph of the chain: only its first lines must come along; with
            # widow control a paragraph of up to three lines cannot be split at all
            n = len(f.lines)
            if f.keep_lines:
                k = n
            elif f.widow:
                k = n if n <= 3 else 2
            else:
                k = min(1, n)
            total += sb + self._para_height(f, k)
            break
        return total

    def _place_para(self, flows, i, prev_after) -> float:
        f: ParaFlow = flows[i]
        if f.page_break_before and not (self.at_top and self.col == 0 and self.page is not None
                                          and not self.page.body_ops):
            self.new_page("pbb")
        # keep with next
        if f.keep_next and i + 1 < len(flows) and not self.at_top:
            need = self._space_before(f, prev_after) + self._keep_chain_height(flows, i)
            full = self.limit - self.col_top
            if self.y + need > self.limit + EPS and need <= full + EPS:
                self.next_column()
                prev_after = 0.0
        lines = f.lines
        floats_placed = False
        if lines and self.page is not None:
            sb0 = self._space_before(f, prev_after)
            para_top = self._jump(self.y + sb0, lines[0].height, self.col_x)
            anchored = [(fl, fx) for ln in lines for fl, fx in ln.floats]
            if anchored:
                for fl, fx in anchored:
                    self._place_float(fl, self.col_x + fx, para_top, para_top)
                floats_placed = True
            if f.relayout is not None:
                wrapped = self._wrap_lines(f, para_top)
                if wrapped is not None:
                    lines = wrapped
        n = len(lines)
        idx = 0
        first = True
        while idx < n:
            sb = self._space_before(f, prev_after) if first else 0.0
            avail = self.limit - self.y - sb
            k = 0
            used = 0.0
            while idx + k < n:
                ln = lines[idx + k]
                extra = self._notes_height(ln)
                # the extra leading of "multiple" spacing below an empty last line
                # may run past the bottom of the page
                if used + ln.height - ln.trailing + extra > avail - self.notes_extra_pending + EPS:
                    break
                used += ln.height
                k += 1
                if ln.break_after:
                    break
            remaining = n - idx
            if k < remaining and not (k > 0 and lines[idx + k - 1].break_after):
                if f.keep_lines and first and not self.at_top:
                    k = 0
                elif f.widow and remaining >= 2:
                    if first and k == 1 and not self.at_top:
                        k = 0  # orphan: move whole start to next page
                    elif remaining - k == 1 and k >= 1:
                        k -= 1  # widow: carry one more line over
                        if k == 1 and first and not self.at_top:
                            k = 0
                        elif k == 0 and self.at_top:
                            k = 1
            if k == 0:
                if self.at_top:
                    k = 1  # nothing fits on an empty page: place it anyway
                else:
                    self.next_column()
                    prev_after = 0.0
                    first_now = first
                    first = first_now
                    continue
            self.y += sb
            para_top = self.y
            for q in range(idx, idx + k):
                self._place_line(lines[q], para_top, floats_placed)
            idx += k
            first = False
            last = lines[idx - 1]
            if last.break_after:
                if last.break_after == "page":
                    self.new_page("break")
                else:
                    self.next_column("break")
                continue
            if idx < n:
                self.next_column()
        return f.space_after

    notes_extra_pending = 0.0

    MIN_WRAP_WIDTH = 18.0

    def _wrap_lines(self, f: ParaFlow, para_top: float):
        """Re-break a paragraph around square/tight floating objects."""
        x_col = self.col_x
        w_col = self.col_w
        height = sum(ln.height for ln in f.lines)
        wraps = [ex for ex in self.page.exclusions
                 if ex[4] == "wrap" and ex[0] < x_col + w_col and ex[2] > x_col
                 and ex[3] > para_top and ex[1] < para_top + height + 200]
        if not wraps:
            return None

        def bounds(off, h):
            top = para_top + off
            bottom = top + h
            blocked = [(ex[0] - x_col, ex[2] - x_col) for ex in wraps if ex[1] < bottom and ex[3] > top]
            if not blocked:
                return None
            # free intervals of the column; text goes to the widest one
            free = []
            cur = 0.0
            for b0, b1 in sorted(blocked):
                if b0 > cur:
                    free.append((cur, min(b0, w_col)))
                cur = max(cur, b1)
            if cur < w_col:
                free.append((cur, w_col))
            free = [iv for iv in free if iv[1] - iv[0] >= self.MIN_WRAP_WIDTH]
            if not free:
                # skip to the bottom of the first object in the way
                return min(ex[3] for ex in wraps if ex[1] < bottom and ex[3] > top) - top
            return max(free, key=lambda iv: iv[1] - iv[0])

        lines = f.relayout(bounds).lines
        # Word's wrap test for a paragraph's last line covers the space after the
        # paragraph too (a picture just below narrows the line above it)
        extend_from = None
        for _ in range(3):
            if not lines or not f.space_after:
                break
            off = sum(ln.height for ln in lines[:-1])
            last_h = lines[-1].height
            if bounds(off, last_h + f.space_after) == bounds(off, last_h):
                break
            if extend_from is not None and off <= extend_from + 0.01:
                break
            extend_from = off

            def extended(o, h, _from=extend_from):
                return bounds(o, h + f.space_after if o >= _from - 0.01 else h)
            lines = f.relayout(extended).lines
        return lines

    # ----------------------------------------------------------------- lines
    def _jump(self, y, h, x, wraps=False):
        """Move a line below top-and-bottom objects it would overlap (and below
        wrapping objects too when ``wraps``, used for table rows)."""
        for x0, y0, x1, y1, kind in sorted(self.page.exclusions, key=lambda e: e[1]):
            if (kind == "jump" or wraps) and y < y1 and y + h > y0 and x0 < x + self.col_w and x1 > x:
                y = y1
        return y

    def _place_line(self, ln: Line, para_top: float, floats_placed=False):
        x = self.col_x
        y = self._jump(self.y, ln.height, x)
        self.y = y
        if not floats_placed:
            for fl, fx in ln.floats:
                self._place_float(fl, x + fx, para_top, y)
        # floats placed at this line may push it down
        y = self._jump(y, ln.height, x)
        self.y = y
        self.page.body_ops.extend(move_ops(ln.ops, x, y))
        if ln.notes:
            self._add_notes(ln.notes)
        self.y += ln.height
        self.at_top = False
        self.page.content_bottom = max(self.page.content_bottom, self.y)

    def _place_float(self, fl: Floating, char_x, para_top, line_top, pre=None):
        if pre is not None:
            self._preplaced.add(id(fl))
            x, y = pre
        else:
            x, y = self._float_pos(fl, char_x, para_top, line_top)
            if page_anchored(fl):
                self.anchor_pages[id(fl)] = (self.page.index, fl, x, y)
            if id(fl) in self._preplaced:
                return  # already placed when its page started
        ops = []
        if fl.shape is not None:
            ops = render_shape(fl.shape, x, y, self.ctx)
        elif fl.picture is not None and fl.picture.data:
            p = fl.picture
            ops = [ImageOp(x, y, fl.width, fl.height, p.data, p.crop)]
        if fl.behind:
            self.page.back_ops.extend(ops)
        else:
            self.page.front_ops.extend(ops)
        if fl.wrap in ("topAndBottom", "square", "tight", "through"):
            dt, db, dl, dr = fl.dist
            if fl.wrap in ("tight", "through"):
                dt = db = 0.0  # the wrap polygon replaces the top/bottom distances
            kind = "jump" if fl.wrap == "topAndBottom" else "wrap"
            self.page.exclusions.append((x - dl, y - dt, x + fl.width + dr, y + fl.height + db, kind))

    def _float_pos(self, fl: Floating, char_x, para_top, line_top):
        sect = self.sect
        pw, ph = sect.page_w, sect.page_h
        ml, mr, mt, mb = sect.margin_left, sect.margin_right, abs(sect.margin_top), abs(sect.margin_bottom)
        _apply_size_pct(fl, sect, self.ctx)
        # horizontal reference
        hr = fl.h_rel
        if hr == "page":
            rx, rw = 0.0, pw
        elif hr in ("margin", "insideMargin", "outsideMargin"):
            rx, rw = ml, pw - ml - mr
        elif hr == "leftMargin":
            rx, rw = 0.0, ml
        elif hr == "rightMargin":
            rx, rw = pw - mr, mr
        elif hr == "character":
            rx, rw = char_x, 0.0
        else:  # column
            rx, rw = self.col_x, self.col_w
        if fl.h_align:
            a = fl.h_align
            if a in ("left", "inside"):
                x = rx
            elif a == "center":
                x = rx + (rw - fl.width) / 2
            else:
                x = rx + rw - fl.width
        elif fl.h_pct is not None:
            x = rx + fl.h_pct * rw
        else:
            x = rx + fl.h_off
        vr = fl.v_rel
        if vr == "page":
            ry, rh = 0.0, ph
        elif vr == "margin":
            ry, rh = mt, ph - mt - mb
        elif vr == "topMargin":
            ry, rh = 0.0, mt
        elif vr == "bottomMargin":
            ry, rh = ph - mb, mb
        elif vr == "line":
            ry, rh = line_top, 0.0
        else:  # paragraph
            ry, rh = para_top, 0.0
        if fl.v_align:
            a = fl.v_align
            if a in ("top", "inside"):
                y = ry
            elif a == "center":
                y = ry + (rh - fl.height) / 2
            else:
                y = ry + rh - fl.height
        elif fl.v_pct is not None:
            y = ry + fl.v_pct * rh
        else:
            y = ry + fl.v_off
        return x, y

    # ----------------------------------------------------------------- footnotes
    def _note_lines(self, note):
        key = (note.kind, note.id)
        blocks = self.reader.footnotes.get(key)
        if not blocks:
            return [], 0.0
        self.ctx.note_label = note.text
        flows = layout_blocks(blocks, self.col_w, self.ctx)
        self.ctx.note_label = ""
        return stack_flows(flows, self.ctx.sum_spacing)

    def _sep_height(self):
        if not hasattr(self, "_sep_cache"):
            if self.reader.footnote_sep:
                flows = layout_blocks(self.reader.footnote_sep, self.col_w, self.ctx)
                self._sep_cache = stack_flows(flows, self.ctx.sum_spacing)
            else:
                self._sep_cache = ([], 12.0)
        return self._sep_cache

    def _notes_height(self, ln: Line) -> float:
        if not ln.notes:
            return 0.0
        h = 0.0
        for n in ln.notes:
            if n.kind != "footnote":
                continue
            _, nh = self._note_lines(n)
            h += nh
        if h and not self.notes:
            h += self._sep_height()[1]
        return h

    def _add_notes(self, notes):
        for n in notes:
            if n.kind != "footnote":
                continue
            lines, h = self._note_lines(n)
            if not lines:
                continue
            if not self.notes:
                self.notes_h += self._sep_height()[1]
            self.notes.append((lines, h))
            self.notes_h += h

    def _render_notes(self):
        if not self.notes:
            return
        page = self.page
        y = self.bottom - self.notes_h
        sep_lines, sep_h = self._sep_height()
        x = self.cols[0][0]
        if sep_lines:
            for ly, ln in sep_lines:
                page.note_ops.extend(move_ops(ln.ops, x, y + ly))
        else:
            page.note_ops.append(LineOp(x, y + sep_h / 2, x + 144, y + sep_h / 2, 0.5, (0, 0, 0)))
        y += sep_h
        for lines, h in self.notes:
            for ly, ln in lines:
                page.note_ops.extend(move_ops(ln.ops, x, y + ly))
            y += h

    # ----------------------------------------------------------------- tables
    def _place_table(self, t: TableFlow):
        headers = t.rows[:t.header_rows]
        rows = list(t.rows)
        i = 0
        placed_on_page = 0
        self._rows_at_top = False
        while i < len(rows):
            row = rows[i]
            if self._rows_at_top:
                row = rows[i] = self._row_at_page_top(row)
                self._rows_at_top = False
            avail = self.limit - self.y
            nxt = rows[i + 1] if i + 1 < len(rows) else None
            close_w, close_ops = self._closing_border(row, nxt)
            if row.height <= avail + EPS and (
                    nxt is None or not close_w or self._row_starts(nxt, avail - row.height, i + 1 < t.header_rows)
                    or row.height + close_w <= avail + EPS):
                self._place_row(row)
                if nxt is not None and close_w and not self._row_starts(nxt, avail - row.height,
                                                                        i + 1 < t.header_rows):
                    # the table breaks after this row: Word closes it with the
                    # border it shares with the next row (large-document p. 9)
                    self._place_line(Line(close_w, close_ops), self.y)
                placed_on_page += 1
                i += 1
                continue
            if row.height <= avail + EPS:
                avail -= close_w  # no room left for the closing border
            is_header = i < t.header_rows
            # Word does not split a row across pages when the room left is less
            # than the row's minimum height: the whole row moves to the next page
            # (the minimum includes the row's border bands)
            if not row.cant_split and not is_header and (
                    not row.min_height or row.min_height + row.band_bottom <= avail + EPS):
                if not close_w:
                    # a split last row closes with its own bottom border
                    close_w, close_ops = self._closing_border(row, row)
                parts = row.split(avail - close_w)
                if parts is not None and (avail > 12 or self.at_top):
                    first, rest = parts
                    self._place_row(first)
                    if close_w:
                        self._place_line(Line(close_w, close_ops), self.y)
                    rows[i] = rest
                    self._table_break(headers, t)
                    placed_on_page = 0
                    continue
            if self.at_top:
                # the row does not fit on an empty page: place it clipped
                self._place_row(row)
                i += 1
                continue
            self._table_break(headers, t)
            placed_on_page = 0

    def _row_starts(self, row: RowFlow, room: float, is_header=False) -> bool:
        """Would ``row`` (or the first part of it) go on this page in ``room``?"""
        if row.height <= room + EPS:
            return True
        if row.cant_split or is_header or room <= 12:
            return False
        if row.min_height and row.min_height + row.band_bottom > room + EPS:
            return False
        return row.split(room) is not None

    @staticmethod
    def _closing_border(row: RowFlow, nxt):
        """(width, ops) of the border drawn under ``row`` when the table breaks
        right after it: the edge it shares with the next row, which that row
        otherwise draws as its top border."""
        if nxt is None or (row.band_bottom and nxt is not row):
            return 0.0, []
        ops: list = []
        width = 0.0
        # each row keeps its own border at a page break: the upper one closes
        # with its bottom border (large-document: a 2.25 pt header bottom above
        # a 1 pt inside border)
        for c in row.cells:
            b = (c.own or {}).get("bottom") if c.own is not None else None
            if b is None:
                b = next((d.borders.get("top") for d in nxt.cells if abs(d.x - c.x) < 0.5), None)
            if b is None or not b.visible:
                continue
            width = max(width, b.width)
            lb, rb = c.borders.get("left"), c.borders.get("right")
            lw = lb.width / 2 if lb is not None and lb.visible else 0.0
            rw = rb.width / 2 if rb is not None and rb.visible else 0.0
            _border_line(ops, b, c.x - lw, b.width / 2, c.x + c.w + rw, b.width / 2, horizontal=True, cap=0)
        return width, ops

    @staticmethod
    def _row_at_page_top(row: RowFlow) -> RowFlow:
        """A row that starts a page after the table broke before it opens with
        its own top border (not the shared edge merged with the row above)."""
        if row.first:
            return row
        cells = []
        band = 0.0
        changed = False
        for c in row.cells:
            own_top = (c.own or {}).get("top") if c.own is not None else c.borders.get("top")
            if own_top is not c.borders.get("top"):
                changed = True
            b = dict(c.borders)
            b["top"] = own_top
            if own_top is not None and own_top.visible and not c.is_continue:
                band = max(band, own_top.width)
            cells.append(replace(c, borders=b))
        if not changed:
            return row
        return replace(row, cells=cells, height=row.height - row.band_top + band, band_top=band)

    def _place_float_table(self, t: TableFlow):
        """Position a floating table (w:tblpPr) like Word and let text avoid it."""
        fl = t.float
        sect = self.sect
        pw, ph = sect.page_w, sect.page_h
        ml, mr = sect.margin_left, sect.margin_right
        mt, mb = abs(sect.margin_top), abs(sect.margin_bottom)
        h = sum(r.height for r in t.rows)
        w = t.width
        ha = fl.get("horzAnchor", "text")
        if ha == "page":
            rx, rw = 0.0, pw
        elif ha == "margin":
            rx, rw = ml, pw - ml - mr
        else:
            rx, rw = self.col_x, self.col_w
        xs = fl.get("tblpXSpec")
        if xs in ("center",):
            x = rx + (rw - w) / 2
        elif xs in ("right", "outside"):
            x = rx + rw - w
        elif xs in ("left", "inside"):
            x = rx
        else:
            # tblpX places the table like an indent: the same edge offset as an
            # inline table (legacy compat: the first cell's text at the position)
            x = rx + fl.get("tblpX", 0.0) + t.x0
        va = fl.get("vertAnchor", "text")
        ys = fl.get("tblpYSpec")
        if va == "page":
            ry, rh = 0.0, ph
        elif va == "margin" or (va == "text" and ys not in (None, "inline")):
            ry, rh = mt, ph - mt - mb
        else:
            ry, rh = self.y, 0.0
        if ys == "top":
            y = ry
        elif ys == "center":
            y = ry + (rh - h) / 2
        elif ys == "bottom":
            y = ry + rh - h
        else:
            y = ry + fl.get("tblpY", 0.0)
        yy = y
        for row in t.rows:
            self.page.front_ops.extend(move_ops(row.render(), x - t.x0, yy))
            yy += row.height
        self.page.exclusions.append((x - fl.get("leftFromText", 0.0), y - fl.get("topFromText", 0.0),
                                     x + w + fl.get("rightFromText", 0.0),
                                     y + h + fl.get("bottomFromText", 0.0), "wrap"))

    def _table_break(self, headers, t):
        self.next_column()
        self._rows_at_top = True
        if headers and len(headers) < len(t.rows):
            for h in headers:
                self._place_row(h)
            self.at_top = True  # a lone repeated header counts as an empty page

    def _place_row(self, row: RowFlow):
        self.y = self._jump(self.y, row.height, self.col_x, wraps=True)
        ops = row.render()
        ln = Line(row.height, ops, row.floats(), row.notes())
        self._place_line(ln, self.y)

    # ----------------------------------------------------------------- finishing
    def _finish_page(self):
        page = self.page
        self._render_notes()
        sect = page.section
        if sect.v_align in ("center", "bottom", "both") and page.body_ops:
            slack = self.bottom - page.content_bottom
            if slack > 0:
                dy = slack / 2 if sect.v_align == "center" else slack if sect.v_align == "bottom" else 0
                page.body_ops = move_ops(page.body_ops, 0, dy)
        page._top = self.top
        page._bottom = self.bottom


def render_headers(pages: list[Page], ctx: LayoutContext, reader):
    total = len(pages)
    by_section: dict[int, int] = {}
    for p in pages:
        by_section[p.section_index] = by_section.get(p.section_index, 0) + 1
    for p in pages:
        sect = p.section
        ctx.fields = {"PAGE": format_number(p.number, sect.pg_fmt), "NUMPAGES": str(total),
                      "SECTIONPAGES": str(by_section.get(p.section_index, total))}
        width = sect.page_w - sect.margin_left - sect.margin_right
        x = sect.margin_left
        for kind_map, is_header in ((sect.headers, True), (sect.footers, False)):
            if p.section_first and sect.title_pg:
                blocks = kind_map.get("first")
            elif reader.even_odd and p.number % 2 == 0:
                blocks = kind_map.get("even")
            else:
                blocks = kind_map.get("default")
            if not blocks:
                continue
            flows = layout_blocks(blocks, width, ctx)
            lines, h = stack_flows(flows, ctx.sum_spacing)
            y = sect.header_dist if is_header else sect.page_h - sect.footer_dist - h
            for ly, ln in lines:
                p.hf_ops.extend(move_ops(ln.ops, x, y + ly))
                for fl, fx in ln.floats:
                    _hf_float(p, fl, x + fx, y + ly, sect, ctx)
        if sect.borders:
            _page_borders(p, sect)


def page_anchored(fl: Floating) -> bool:
    """A float the text of its whole page wraps around (placed when the page
    starts, from the position found in the previous layout pass)."""
    if fl.h_rel == "character":
        return False
    if fl.wrap in ("square", "tight", "through"):
        return True
    # top-and-bottom objects only matter above their anchor when positioned
    # on the page itself
    return fl.wrap == "topAndBottom" and fl.v_rel in ("page", "margin", "topMargin", "bottomMargin")


def _apply_size_pct(fl: Floating, sect: Section, ctx):
    """Resolve relative sizes (wp14:sizeRelH/V) and auto-fit text boxes."""
    pw, ph = sect.page_w, sect.page_h
    ml, mr, mt, mb = sect.margin_left, sect.margin_right, abs(sect.margin_top), abs(sect.margin_bottom)
    refs_w = {"page": pw, "margin": pw - ml - mr, "leftMargin": ml, "rightMargin": mr,
              "insideMargin": ml, "outsideMargin": mr}
    refs_h = {"page": ph, "margin": ph - mt - mb, "topMargin": mt, "bottomMargin": mb,
              "insideMargin": mt, "outsideMargin": mb}
    if "w" in fl.size_pct:
        rel, f = fl.size_pct["w"]
        fl.width = f * refs_w.get(rel, pw - ml - mr)
        if fl.shape is not None:
            fl.shape.width = fl.width
    if "h" in fl.size_pct and not (fl.shape is not None and fl.shape.autofit):
        rel, f = fl.size_pct["h"]
        fl.height = f * refs_h.get(rel, ph - mt - mb)
        if fl.shape is not None:
            fl.shape.height = fl.height
    if fl.shape is not None and fl.shape.autofit:
        from .layout import shape_height
        fl.height = fl.shape.height = shape_height(fl.shape, ctx)


def _hf_float(page: Page, fl: Floating, char_x, line_top, sect, ctx):
    x, y = _hf_float_pos(fl, char_x, line_top, sect, ctx)
    if fl.shape is not None:
        ops = render_shape(fl.shape, x, y, ctx)
    elif fl.picture is not None and fl.picture.data:
        ops = [ImageOp(x, y, fl.width, fl.height, fl.picture.data, fl.picture.crop)]
    else:
        ops = []
    if fl.behind:
        page.back_ops = ops + page.back_ops
    else:
        page.hf_ops.extend(ops)


def _hf_float_pos(fl: Floating, char_x, line_top, sect, ctx):
    _apply_size_pct(fl, sect, ctx)
    pw, ph = sect.page_w, sect.page_h
    ml, mr, mt, mb = sect.margin_left, sect.margin_right, abs(sect.margin_top), abs(sect.margin_bottom)
    refs_h = {"page": (0, pw), "margin": (ml, pw - ml - mr), "leftMargin": (0, ml),
              "rightMargin": (pw - mr, mr), "character": (char_x, 0), "column": (ml, pw - ml - mr)}
    rx, rw = refs_h.get(fl.h_rel, (ml, pw - ml - mr))
    if fl.h_align:
        x = rx if fl.h_align == "left" else rx + (rw - fl.width) / 2 if fl.h_align == "center" else rx + rw - fl.width
    elif fl.h_pct is not None:
        x = rx + fl.h_pct * rw
    else:
        x = rx + fl.h_off
    refs_v = {"page": (0, ph), "margin": (mt, ph - mt - mb), "topMargin": (0, mt),
              "bottomMargin": (ph - mb, mb)}
    ry, rh = refs_v.get(fl.v_rel, (line_top, 0))
    if fl.v_align:
        y = ry if fl.v_align == "top" else ry + (rh - fl.height) / 2 if fl.v_align == "center" else ry + rh - fl.height
    elif fl.v_pct is not None:
        y = ry + fl.v_pct * rh
    else:
        y = ry + fl.v_off
    return x, y


def _page_borders(page: Page, sect: Section):
    b = sect.borders
    # offsets measured from the text (default) - approximate as from the margin
    x0 = sect.margin_left - (b["left"].space if b.get("left") else 0)
    x1 = sect.page_w - sect.margin_right + (b["right"].space if b.get("right") else 0)
    y0 = abs(sect.margin_top) - (b["top"].space if b.get("top") else 0)
    y1 = sect.page_h - abs(sect.margin_bottom) + (b["bottom"].space if b.get("bottom") else 0)
    ops = []
    _border_line(ops, b.get("top"), x0, y0, x1, y0, True)
    _border_line(ops, b.get("bottom"), x0, y1, x1, y1, True)
    _border_line(ops, b.get("left"), x0, y0, x0, y1, False)
    _border_line(ops, b.get("right"), x1, y0, x1, y1, False)
    page.hf_ops.extend(ops)
