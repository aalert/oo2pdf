"""Score oo2pdf against Office reference PDFs for every file in a corpus (dev only).

    python tests/tools/corpus_score.py tests/corpus [--locale system] [name-filter]

For each document prints: pages (ours/ref), share of reference words found
in our output, share of matched words on the same page, and how many of those
are within 1pt of Office's position.
"""
from __future__ import annotations

import difflib
import statistics
import sys
import time
import traceback
from pathlib import Path

import pymupdf

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src"))


def words(doc):
    out = []
    for pno, page in enumerate(doc):
        for w in page.get_text("words"):
            out.append((w[4], pno, w[0], w[3]))
    return out


def score(src: Path, ref_pdf: Path, out_pdf: Path, opts):
    import oo2pdf
    t0 = time.perf_counter()
    oo2pdf.convert(src, out_pdf, **opts)
    dt = time.perf_counter() - t0
    ours = pymupdf.open(out_pdf)
    ref = pymupdf.open(ref_pdf)
    wa, wb = words(ours), words(ref)
    sm = difflib.SequenceMatcher(a=[w[0] for w in wa], b=[w[0] for w in wb], autojunk=False)
    matched = same_page = close = 0
    dys = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag != "equal":
            continue
        for k in range(i2 - i1):
            x, y = wa[i1 + k], wb[j1 + k]
            matched += 1
            if x[1] == y[1]:
                same_page += 1
                d = max(abs(x[2] - y[2]), abs(x[3] - y[3]))
                dys.append(abs(x[3] - y[3]))
                if d <= 1.0:
                    close += 1
    n = max(1, len(wb))
    return {
        "pages": (len(ours), len(ref)), "words": len(wb), "found": matched / n,
        "same_page": same_page / max(1, matched), "close": close / max(1, matched),
        "dy_med": statistics.median(dys) if dys else 0.0, "time": dt,
    }


def main(argv):
    corpus = Path(argv[1]) if len(argv) > 1 else ROOT / "tests" / "corpus"
    flt = [a for a in argv[2:] if not a.startswith("-")]
    locale = "system" if "--locale" in argv else "en-US"
    out_dir = corpus / "out"
    out_dir.mkdir(exist_ok=True)
    rows = []
    for src in sorted(list(corpus.glob("*.docx")) + list(corpus.glob("*.xlsx"))):
        if flt and not any(f.lower() in src.name.lower() for f in flt):
            continue
        suffix = "_xlsx" if src.suffix == ".xlsx" else ""
        ref = corpus / "reference" / (src.stem + suffix + ".pdf")
        if not ref.exists():
            continue
        opts = {"locale": locale} if src.suffix == ".xlsx" else {}
        try:
            r = score(src, ref, out_dir / (src.stem + suffix + ".pdf"), opts)
            rows.append((src.name, r))
            print(f"{src.name:34} pages {r['pages'][0]:>2}/{r['pages'][1]:<2}  words {r['words']:5d}"
                  f"  found {r['found']:6.1%}  same-page {r['same_page']:6.1%}  <=1pt {r['close']:6.1%}"
                  f"  dy~{r['dy_med']:5.2f}  {r['time']:.2f}s", flush=True)
        except Exception as exc:
            print(f"{src.name:34} ERROR {exc!r}")
            if "-v" in argv:
                traceback.print_exc()
    if rows:
        tot = sum(r["words"] for _, r in rows)
        close = sum(r["close"] * r["found"] * r["words"] for _, r in rows) / max(1, tot)
        pages_ok = sum(1 for _, r in rows if r["pages"][0] == r["pages"][1])
        print(f"\n{len(rows)} files, page count match {pages_ok}/{len(rows)}, "
              f"reference words within 1pt: {close:.1%}")


if __name__ == "__main__":
    main(sys.argv)
