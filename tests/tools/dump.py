"""Dump the body structure of a docx (dev only): python tests/tools/dump.py file.docx [max]"""
import sys
import zipfile

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def short(el, depth=0, maxlen=160):
    tag = etree.QName(el).localname
    txt = "".join(t.text or "" for t in el.iter(W + "t"))[:60]
    extra = []
    ppr = el.find(W + "pPr")
    if ppr is not None:
        ps = ppr.find(W + "pStyle")
        if ps is not None:
            extra.append("style=" + ps.get(W + "val"))
        if ppr.find(".//" + W + "sectPr") is not None:
            extra.append("SECT")
        sp = ppr.find(W + "spacing")
        if sp is not None:
            extra.append("sp=" + ",".join(f"{etree.QName(k).localname}={v}" for k, v in sp.attrib.items()))
    kinds = set()
    for d in el.iter():
        n = etree.QName(d).localname if isinstance(d.tag, str) else ""
        if n in ("anchor", "inline", "txbxContent", "fldChar", "fldSimple", "sdt", "tbl", "pict", "br", "tab", "ptab"):
            kinds.add(n)
    return f"{'  ' * depth}{tag} {' '.join(extra)} {sorted(kinds)} {txt!r}"


if __name__ == "__main__":
    z = zipfile.ZipFile(sys.argv[1])
    root = etree.fromstring(z.read("word/document.xml"))
    body = root.find(W + "body")
    limit = int(sys.argv[2]) if len(sys.argv) > 2 else 40
    for i, el in enumerate(body):
        if i >= limit:
            break
        print(i, short(el).encode("ascii", "replace").decode())
        if etree.QName(el).localname == "sdt":
            for j, c in enumerate(el.find(W + "sdtContent")):
                print("   ", j, short(c).encode("ascii", "replace").decode())
