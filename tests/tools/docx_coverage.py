"""Check that the content of a .docx appears in the PDF we generated (dev only).

Compares the words of every story in the package (body, headers, footers,
footnotes, endnotes, text boxes, equations) with the PDF's words, and the
number of pictures referenced by the document with the images drawn.

    python tests/tools/docx_coverage.py my_tests            # every docx with out/<name>.pdf
    python tests/tools/docx_coverage.py doc.docx out.pdf -v # list the missing words
"""
from __future__ import annotations

import re
import sys
import unicodedata
import zipfile
from collections import Counter
from pathlib import Path

import pymupdf
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
V = "{urn:schemas-microsoft-com:vml}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

STORIES = re.compile(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")


def norm(w: str) -> str:
    w = unicodedata.normalize("NFKC", w)  # math italic letters -> ASCII
    w = w.replace("−", "-").replace("’", "'").replace("‘", "'")
    return w.strip(".,;:()[]{}\"'«»“”¡!¿?")


def _hidden(r) -> bool:
    rpr = r.find(W + "rPr")
    if rpr is None:
        return False
    v = rpr.find(W + "vanish")
    return v is not None and v.get(W + "val") not in ("0", "false")


def story_text(root) -> str:
    """Visible text of a story: skips mc:Fallback copies, deleted text and hidden runs."""
    for fb in list(root.iter(MC + "Fallback")):
        fb.getparent().remove(fb)
    parts = []
    for p in root.iter(W + "p"):
        # paragraphs inside text boxes are visited on their own; skip nested ones here
        buf = []
        for el in p.iter(W + "t", M + "t", W + "tab", W + "br", W + "delText", W + "instrText"):
            owner = el.getparent()
            while owner is not None and owner.tag != W + "p":
                owner = owner.getparent()
            if owner is not p:
                continue
            if el.tag in (W + "delText", W + "instrText"):
                continue
            if el.tag in (W + "tab", W + "br"):
                buf.append(" ")
                continue
            run = el.getparent()
            if run.tag == W + "r" and _hidden(run):
                continue
            buf.append(el.text or "")
        parts.append("".join(buf))
    return "\n".join(parts)


def docx_words(path) -> tuple[Counter, dict]:
    z = zipfile.ZipFile(path)
    words = Counter()
    for name in z.namelist():
        m = STORIES.match(name)
        if not m:
            continue
        root = etree.fromstring(z.read(name))
        if m.group(1).startswith(("header", "footer")) and not _referenced(z, name):
            continue
        for w in story_text(root).split():
            w = norm(w)
            if w:
                words[w] += 1
    root = etree.fromstring(z.read("word/document.xml"))
    for fb in list(root.iter(MC + "Fallback")):
        fb.getparent().remove(fb)
    pics = sum(1 for _ in root.iter(A + "blip")) + sum(1 for _ in root.iter(V + "imagedata"))
    return words, {"pictures": pics}


def _referenced(z, name) -> bool:
    doc = z.read("word/document.xml").decode("utf-8", "replace")
    rels = z.read("word/_rels/document.xml.rels").decode("utf-8", "replace")
    target = name.split("/", 1)[1]
    m = re.search(r'Id="([^"]+)"[^>]*Target="%s"|Target="%s"[^>]*Id="([^"]+)"' % (target, target), rels)
    if not m:
        return False
    rid = m.group(1) or m.group(2)
    return f'r:id="{rid}"' in doc


def pdf_words(path) -> tuple[Counter, dict]:
    words = Counter()
    images = 0
    raw = []
    for page in pymupdf.open(path):
        raw.append(unicodedata.normalize("NFKC", page.get_text()))
        for w in page.get_text("words"):
            w = norm(w[4])
            if w:
                words[w] += 1
        images += len(page.get_images(full=True))
    return words, {"images": images, "blob": re.sub(r"\s+", "", "".join(raw))}


def _loose(s):
    """Letters/digits only: equations are drawn glyph by glyph with operators spaced."""
    return re.sub(r"[^\w]", "", s).casefold()  # caps/smallCaps styles change the case


def compare(docx, pdf, verbose=False):
    dw, dinfo = docx_words(docx)
    pw, pinfo = pdf_words(pdf)
    # words may be split/merged differently (hyphenation, math): match on joined text too
    blob = pinfo["blob"]
    missing = Counter()
    for w, n in dw.items():
        have = pw.get(w, 0)
        if have < n and w not in blob and _loose(w) not in _loose(blob):
            missing[w] = n - have
    total = sum(dw.values())
    miss = sum(missing.values())
    print(f"{Path(docx).name:40} words {total:6}  missing {miss:4} ({100 * miss / max(total, 1):4.1f}%)  "
          f"pictures {dinfo['pictures']:3} images drawn {pinfo['images']:3}".encode("ascii", "replace").decode())
    if verbose and missing:
        print("   ", " ".join(f"{w}x{n}" if n > 1 else w for w, n in missing.most_common(80))
              .encode("ascii", "replace").decode())
    return missing


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    verbose = "-v" in sys.argv
    if len(args) == 2 and args[0].endswith(".docx"):
        compare(args[0], args[1], verbose)
    else:
        folder = Path(args[0])
        for d in sorted(folder.glob("*.docx")):
            out = folder / "out" / (d.stem + ".pdf")
            if out.exists():
                compare(d, out, verbose)
