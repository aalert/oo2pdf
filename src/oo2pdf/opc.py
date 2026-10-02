"""Open Packaging Conventions: reading parts and relationships from an OOXML zip."""
from __future__ import annotations

import posixpath
import zipfile
from typing import BinaryIO

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
V_NS = "urn:schemas-microsoft-com:vml"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

W = "{%s}" % W_NS
R = "{%s}" % R_NS
A = "{%s}" % A_NS
WP = "{%s}" % WP_NS
PIC = "{%s}" % PIC_NS
WPS = "{%s}" % WPS_NS
WPG = "{%s}" % WPG_NS
MC = "{%s}" % MC_NS
V = "{%s}" % V_NS
M = "{%s}" % M_NS

REL_OFFICE_DOC = "officeDocument"

_PARSER = etree.XMLParser(resolve_entities=False, huge_tree=True, remove_blank_text=False)

# ISO/IEC 29500 "strict" documents use different namespace URIs; map them to the
# transitional ones so the rest of the engine only deals with one vocabulary.
_STRICT_MAP = {
    b"http://purl.oclc.org/ooxml/wordprocessingml/main": W_NS.encode(),
    b"http://purl.oclc.org/ooxml/officeDocument/relationships": R_NS.encode(),
    b"http://purl.oclc.org/ooxml/drawingml/main": A_NS.encode(),
    b"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing": WP_NS.encode(),
    b"http://purl.oclc.org/ooxml/drawingml/picture": PIC_NS.encode(),
    b"http://purl.oclc.org/ooxml/spreadsheetml/main":
        b"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}


def local(el) -> str:
    tag = el.tag
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


class Package:
    """A read-only view of an OOXML package."""

    def __init__(self, source: str | BinaryIO):
        self.zip = zipfile.ZipFile(source)
        self._names = {n.lstrip("/"): n for n in self.zip.namelist()}
        self._lower = {n.lower(): n for n in self._names}
        self._xml: dict[str, etree._Element | None] = {}
        self._rels: dict[str, dict[str, tuple[str, str, bool]]] = {}

    def _real(self, name: str) -> str | None:
        name = name.lstrip("/")
        if name in self._names:
            return self._names[name]
        real = self._lower.get(name.lower())
        return self._names[real] if real else None

    def has(self, name: str) -> bool:
        return self._real(name) is not None

    def read(self, name: str) -> bytes | None:
        real = self._real(name)
        if real is None:
            return None
        return self.zip.read(real)

    def xml(self, name: str):
        name = name.lstrip("/")
        if name not in self._xml:
            data = self.read(name)
            if data is None:
                self._xml[name] = None
            else:
                for strict, trans in _STRICT_MAP.items():
                    if strict in data:
                        data = data.replace(strict, trans)
                self._xml[name] = etree.fromstring(data, _PARSER)
        return self._xml[name]

    def rels(self, part: str) -> dict[str, tuple[str, str, bool]]:
        """rId -> (type suffix, absolute target, is_external)."""
        part = part.lstrip("/")
        if part in self._rels:
            return self._rels[part]
        d, base = posixpath.split(part)
        rel_name = posixpath.join(d, "_rels", base + ".rels")
        out: dict[str, tuple[str, str, bool]] = {}
        root = self.xml(rel_name)
        if root is not None:
            for rel in root:
                rid = rel.get("Id")
                rtype = (rel.get("Type") or "").rsplit("/", 1)[-1]
                target = rel.get("Target") or ""
                external = rel.get("TargetMode") == "External"
                if not external:
                    if target.startswith("/"):
                        target = target.lstrip("/")
                    else:
                        target = posixpath.normpath(posixpath.join(d, target))
                out[rid] = (rtype, target, external)
        self._rels[part] = out
        return out

    def rel_target(self, part: str, rtype: str) -> str | None:
        for t, target, ext in self.rels(part).values():
            if t == rtype and not ext:
                return target
        return None

    def main_part(self) -> str | None:
        return self.rel_target("", REL_OFFICE_DOC)
