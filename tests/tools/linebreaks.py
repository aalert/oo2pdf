"""Compare where lines break: ours vs Word's PDF (dev only).

For every line of text the first word is taken; the two sequences of line
starts are aligned and differences listed. A different line start is what
eventually drifts a document by a line and then by a page.

    python tests/tools/linebreaks.py my_tests            # summary per document
    python tests/tools/linebreaks.py ref.pdf ours.pdf -v # list the differences
"""
from __future__ import annotations

import difflib
import sys
from pathlib import Path

import pymupdf


def line_starts(path):
    out = []
    for pno, page in enumerate(pymupdf.open(path)):
        words = sorted(page.get_text("words"), key=lambda w: (round(w[3]), w[0]))
        lines: list[list] = []
        for w in words:
            if lines and abs(lines[-1][-1][3] - w[3]) < 2.5 and w[0] >= lines[-1][-1][0] - 1:
                lines[-1].append(w)
            else:
                lines.append([w])
        for ln in lines:
            ln.sort(key=lambda w: w[0])
            # key: first two words, so repeated line starts ("the") stay distinguishable
            out.append((" ".join(w[4] for w in ln[:2]), pno + 1, " ".join(w[4] for w in ln)))
    return out


def compare(ref, ours, verbose=False, name=""):
    R, O = line_starts(ref), line_starts(ours)
    sm = difflib.SequenceMatcher(a=[r[0] for r in R], b=[o[0] for o in O], autojunk=False)
    diffs = [op for op in sm.get_opcodes() if op[0] != "equal"]
    bad = sum(max(i2 - i1, j2 - j1) for _, i1, i2, j1, j2 in diffs)
    print(f"{name:40} lines {len(R):5}  different breaks {bad:4} ({100 * bad / max(len(R), 1):4.1f}%)"
          .encode("ascii", "replace").decode())
    if verbose:
        for tag, i1, i2, j1, j2 in diffs:
            for r in R[i1:i2]:
                print(f"   ref  p{r[1]:<3} {r[2][:110]}".encode("ascii", "replace").decode())
            for o in O[j1:j2]:
                print(f"   ours p{o[1]:<3} {o[2][:110]}".encode("ascii", "replace").decode())
            print("   --")
    return bad


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    verbose = "-v" in sys.argv
    if len(args) == 2 and args[0].endswith(".pdf"):
        compare(args[0], args[1], verbose, Path(args[1]).name)
    else:
        folder = Path(args[0])
        only = args[1:]
        for d in sorted(folder.glob("*.docx")):
            if only and not any(o.lower() in d.name.lower() for o in only):
                continue
            ref, ours = folder / "reference" / (d.stem + ".pdf"), folder / "out" / (d.stem + ".pdf")
            if ref.exists() and ours.exists():
                compare(ref, ours, verbose, d.name)
