"""Tabulate how far Word lets justified lines overflow before squeezing (dev only).

    python tests/tools/squeeze.py ref.pdf left right [pages]
Lines are reconstructed from the reference PDF; 'ACC' rows are lines Word
squeezed (natural width > available), 'REJ' rows are the overflow the next
word would have caused.
"""
import sys

import pymupdf

sys.path.insert(0, "src")
from oo2pdf.fonts import FontManager  # noqa: E402

fm = FontManager()


def face(fontname):
    base = fontname.split("+")[-1]
    fam = base.split("-")[0].replace("MT", "").replace("PS", "")
    fam = {"TimesNewRoman": "Times New Roman", "CourierNew": "Courier New"}.get(fam, fam)
    return fm.face(fam, "Bold" in base, "Italic" in base or "Oblique" in base).font


def main(pdf, left, right, pages=None):
    doc = pymupdf.open(pdf)
    out = []
    for pno, page in enumerate(doc):
        if pages and pno + 1 not in pages:
            continue
        rows = {}
        for b in page.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                for s in l["spans"]:
                    for c in s["chars"]:
                        rows.setdefault(round(s["origin"][1]), []).append((c["origin"][0], c["c"], s["font"], round(s["size"] * 2) / 2))
        ys = sorted(rows)
        for i, y in enumerate(ys[:-1]):
            cur = sorted(rows[y])
            nxt = sorted(rows[ys[i + 1]])
            if abs(cur[0][0] - left) > 40 or ys[i + 1] - y > 20:
                continue
            while cur and cur[-1][1] == " ":
                cur.pop()
            text = "".join(c[1] for c in cur)
            nt = "".join(c[1] for c in nxt).strip()
            if not nt or " " not in text:
                continue
            nat = sum(face(c[2]).char_width(c[1], c[3]) for c in cur)
            nw = nt.split(" ")[0]
            f0, s0 = face(nxt[0][2]), nxt[0][3]
            sp = face(cur[-1][2]).char_width(" ", cur[-1][3])
            size = cur[0][3]
            line_left = cur[0][0]
            av = right - line_left
            acc = nat - av
            rej = nat + sp + f0.width(nw, s0) - av
            if acc > 0:
                out.append(("ACC", pno + 1, y, round(acc, 2), round(acc / size, 3), text[-25:]))
            out.append(("REJ", pno + 1, y, round(rej, 2), round(rej / size, 3), nw))
    for o in sorted(out, key=lambda o: o[3]):
        if o[3] > -2:
            print(o)


if __name__ == "__main__":
    pages = [int(x) for x in sys.argv[4].split(",")] if len(sys.argv) > 4 else None
    main(sys.argv[1], float(sys.argv[2]), float(sys.argv[3]), pages)
