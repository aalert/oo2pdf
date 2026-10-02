"""Print line baselines of a page for ours vs reference (dev only)."""
import sys
import pymupdf


def lines(f, pno):
    d = pymupdf.open(f)
    p = d[pno]
    out = []
    for b in p.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            s0 = l["spans"][0]
            out.append((round(s0["origin"][1], 2), round(s0["origin"][0], 2), t[:38]))
    return sorted(out)


if __name__ == "__main__":
    ref, ours, pno = sys.argv[1], sys.argv[2], int(sys.argv[3]) - 1
    R, O = lines(ref, pno), lines(ours, pno)
    for i in range(max(len(R), len(O))):
        r = R[i] if i < len(R) else (0, 0, "")
        o = O[i] if i < len(O) else (0, 0, "")
        line = f"{r[0]:8.2f} {r[1]:7.2f} {r[2]:<38} | {o[0]:8.2f} {o[1]:7.2f} {o[2]:<38}"
        print(line.encode("ascii", "replace").decode())
