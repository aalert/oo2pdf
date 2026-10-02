"""DOCX -> PDF conversion."""
from __future__ import annotations

from ..draw import PdfWriter
from ..fonts import FontManager
from ..opc import Package
from .layout import LayoutContext
from .paginate import Paginator, render_headers
from .reader import DocxReader


def convert_docx(source, target, fonts: FontManager | None = None) -> int:
    """Convert a .docx file (path or binary stream) to PDF. Returns the page count."""
    pkg = Package(source)
    fonts = fonts or FontManager()
    reader = DocxReader(pkg)
    sections = reader.read_document()
    ctx = LayoutContext(fonts, reader)
    endnotes = [(label, reader.footnotes.get(("endnote", nid)) or [])
                for label, nid in reader.endnote_order]
    pag = Paginator(ctx, reader)
    pages = pag.run(sections, endnotes)
    history: dict[int, list] = {}
    frozen: dict[int, tuple] = {}
    for _ in range(8):
        # wrapping floats also push away the text above their anchor: lay out
        # again with them placed when their page starts, until every float
        # lands where it was placed. Pages settle in document order: only the
        # first page whose floats moved is examined for oscillation.
        # positions are kept exact (a drop cap ends exactly on a line top);
        # comparisons are rounded
        latest = {id(fl): (fl, (pg, x, y)) for pg, fl, x, y in pag.anchor_pages.values()}
        placed = {id(f): (k, x, y) for k, v in pag.prefloats.items() for f, x, y in v}

        def same(a, b):
            return a is not None and b is not None and a[0] == b[0]                 and abs(a[1] - b[1]) < 0.01 and abs(a[2] - b[2]) < 0.01
        changed = [pos[0] for fid, (fl, pos) in latest.items()
                   if fid not in frozen and not same(placed.get(fid), pos)]
        if not changed:
            break
        first = min(changed)
        want: dict[int, list] = {}
        for fid, (fl, pos) in latest.items():
            if fid in frozen:
                pos = frozen[fid]
            elif pos[0] == first:
                seen = history.setdefault(fid, [])
                earlier = [q for q in seen[:-1] if same(q, pos)]
                if earlier:
                    # the float and the text it pushes away oscillate (wrapping
                    # the text above moves the anchor, which moves the float,
                    # which unwraps the text): Word keeps the float where it
                    # first was (sample-files image document, page 2)
                    pos = frozen[fid] = earlier[0]
                seen.append(pos)
            elif pos[0] > first:
                history.pop(fid, None)  # laid out after a page that is still moving
            want.setdefault(pos[0], []).append((fl, pos[1], pos[2]))
        pag = Paginator(ctx, reader)
        pag.prefloats = want
        pages = pag.run(sections, endnotes)
    render_headers(pages, ctx, reader)
    writer = PdfWriter(target, title=reader.title, author=reader.author)
    for p in pages:
        ops = p.back_ops + p.hf_ops + p.body_ops + p.note_ops + p.front_ops
        writer.page(p.width, p.height, ops)
    writer.save()
    return len(pages)
