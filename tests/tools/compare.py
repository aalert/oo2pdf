"""Compare oo2pdf output against an Office-generated reference PDF (dev only).

Prints per-page word position deltas and writes overlay PNGs:
our output in red, the reference in blue (overlapping = dark).

    python tests/tools/compare.py tests/fixtures/basic.docx
"""
from __future__ import annotations

import difflib
import statistics
import sys
from pathlib import Path

import pymupdf

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src"))


def words(doc):
    out = []
    for pno, page in enumerate(doc):
        for w in page.get_text("words"):
            out.append((w[4], pno, w[0], w[3]))  # text, page, x0, baseline-ish (y1)
    return out


def overlay(ours, ref, out_dir: Path, stem: str, dpi=60):
    from PIL import Image, ImageChops
    n = max(len(ours), len(ref))
    paths = []
    for i in range(n):
        imgs = []
        for doc in (ours, ref):
            if i < len(doc):
                pix = doc[i].get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY)
                imgs.append(Image.frombytes("L", (pix.width, pix.height), pix.samples))
            else:
                imgs.append(None)
        base = imgs[0] or imgs[1]
        a = imgs[0] or Image.new("L", base.size, 255)
        b = imgs[1] or Image.new("L", base.size, 255)
        if a.size != b.size:
            b = b.resize(a.size)
        # red channel = reference ink, blue = our ink
        rgb = Image.merge("RGB", (b, ImageChops.multiply(a, b), a))
        p = out_dir / f"{stem}-p{i + 1}.png"
        rgb.save(p)
        paths.append(p)
    return paths


def main(argv):
    import oo2pdf
    src = Path(argv[1])
    suffix = "_xlsx" if src.suffix.lower() == ".xlsx" else ""
    ref_pdf = src.parent / "reference" / (src.stem + suffix + ".pdf")
    out_dir = src.parent / "out"
    out_dir.mkdir(exist_ok=True)
    ours_pdf = out_dir / (src.stem + suffix + ".pdf")
    opts = {"locale": "system"} if src.suffix.lower() == ".xlsx" else {}
    oo2pdf.convert(src, ours_pdf, **opts)
    ours = pymupdf.open(ours_pdf)
    ref = pymupdf.open(ref_pdf)
    print(f"pages: ours={len(ours)} ref={len(ref)}")
    for i in range(min(len(ours), len(ref))):
        a, b = ours[i].rect, ref[i].rect
        if abs(a.width - b.width) > 1 or abs(a.height - b.height) > 1:
            print(f"  page {i+1} size ours={a.width:.1f}x{a.height:.1f} ref={b.width:.1f}x{b.height:.1f}")
    wa, wb = words(ours), words(ref)
    sm = difflib.SequenceMatcher(a=[w[0] for w in wa], b=[w[0] for w in wb], autojunk=False)
    per_page: dict[int, list] = {}
    page_mismatch = 0
    matched = 0
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag != "equal":
            continue
        for k in range(i2 - i1):
            x, y = wa[i1 + k], wb[j1 + k]
            matched += 1
            if x[1] != y[1]:
                page_mismatch += 1
                continue
            per_page.setdefault(y[1], []).append((x[2] - y[2], x[3] - y[3]))
    print(f"words: ours={len(wa)} ref={len(wb)} matched={matched} on-different-page={page_mismatch}")
    for p in sorted(per_page):
        d = per_page[p]
        dx = [abs(v[0]) for v in d]
        dy = [abs(v[1]) for v in d]
        print(f"  page {p+1}: n={len(d):4d}  |dx| med={statistics.median(dx):6.2f} max={max(dx):7.2f}"
              f"   |dy| med={statistics.median(dy):6.2f} max={max(dy):7.2f}")
    if "-v" in argv:
        print("  first large deltas:")
        shown = 0
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag != "equal":
                continue
            for k in range(i2 - i1):
                x, y = wa[i1 + k], wb[j1 + k]
                if x[1] != y[1] or abs(x[2] - y[2]) > 3 or abs(x[3] - y[3]) > 3:
                    print(f"    {x[0]!r:20} ours p{x[1]+1} ({x[2]:.1f},{x[3]:.1f})  ref p{y[1]+1} ({y[2]:.1f},{y[3]:.1f})")
                    shown += 1
                    if shown > 25:
                        break
            if shown > 25:
                break
    paths = overlay(ours, ref, out_dir, src.stem + suffix)
    print("overlays:", ", ".join(str(p) for p in paths))


if __name__ == "__main__":
    main(sys.argv)
