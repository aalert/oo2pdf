"""Render reference and generated pages side by side as PNGs (dev only).

    python tests/tools/sidebyside.py my_tests demo            # all pages
    python tests/tools/sidebyside.py my_tests demo 3 --dpi 110 # one page

Images go to <folder>/out/compare/<name>-pNN.png (left: Office, right: ours).
"""
from __future__ import annotations

import sys
from pathlib import Path

import pymupdf
from PIL import Image, ImageDraw


def render(pdf, pno, dpi):
    d = pymupdf.open(pdf)
    if pno >= len(d):
        return None
    pix = d[pno].get_pixmap(dpi=dpi)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dpi = 80
    if "--dpi" in sys.argv:
        dpi = int(sys.argv[sys.argv.index("--dpi") + 1])
        args = [a for a in args if a != str(dpi)]
    folder, stem = Path(args[0]), args[1]
    pages = [int(p) for p in args[2:]]
    ref = folder / "reference" / f"{stem}.pdf"
    ours = folder / "out" / f"{stem}.pdf"
    n = max(len(pymupdf.open(ref)), len(pymupdf.open(ours)))
    out_dir = folder / "out" / "compare"
    out_dir.mkdir(parents=True, exist_ok=True)
    for pno in (p - 1 for p in pages) if pages else range(n):
        a, b = render(ref, pno, dpi), render(ours, pno, dpi)
        w = max(i.width for i in (a, b) if i)
        h = max(i.height for i in (a, b) if i)
        img = Image.new("RGB", (w * 2 + 12, h), (120, 120, 120))
        for k, im in enumerate((a, b)):
            if im:
                img.paste(im, (k * (w + 12), 0))
        ImageDraw.Draw(img).text((4, 4), "Office", fill=(255, 0, 0))
        ImageDraw.Draw(img).text((w + 16, 4), "oo2pdf", fill=(255, 0, 0))
        path = out_dir / f"{stem}-p{pno + 1:02d}.png"
        img.save(path)
        print(path)


if __name__ == "__main__":
    main()
