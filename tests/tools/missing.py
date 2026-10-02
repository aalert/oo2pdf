"""Show reference words missing from our output and vice versa (dev only)."""
import difflib
import sys

import pymupdf


def words(path):
    out = []
    for pno, page in enumerate(pymupdf.open(path)):
        for w in page.get_text("words"):
            out.append((w[4], pno, round(w[0], 1), round(w[3], 1)))
    return out


if __name__ == "__main__":
    ref, ours = words(sys.argv[1]), words(sys.argv[2])
    sm = difflib.SequenceMatcher(a=[w[0] for w in ours], b=[w[0] for w in ref], autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        a = " ".join(w[0] for w in ours[i1:i2])[:100]
        b = " ".join(w[0] for w in ref[j1:j2])[:100]
        loc = ref[j1] if j1 < len(ref) else (ours[i1] if i1 < len(ours) else None)
        print(f"{tag:8} ours: {a!r}\n         ref:  {b!r}  @ {loc[1:] if loc else ''}".encode("ascii", "replace").decode())
