"""oo2pdf - pure-Python DOCX/XLSX to PDF conversion that mimics Microsoft Office.

Usage::

    import oo2pdf
    oo2pdf.convert("report.docx", "report.pdf")
    oo2pdf.convert("book.xlsx", "book.pdf")
"""
from __future__ import annotations

import io
import os
from pathlib import Path

from .fonts import FontManager

__all__ = ["convert", "convert_docx", "convert_xlsx", "FontManager"]
__version__ = "0.1.0"


def _kind(source, kind):
    if kind:
        return kind.lower().lstrip(".")
    if isinstance(source, (str, os.PathLike)):
        ext = Path(source).suffix.lower().lstrip(".")
        if ext in ("docx", "docm", "dotx", "dotm"):
            return "docx"
        if ext in ("xlsx", "xlsm", "xltx", "xltm"):
            return "xlsx"
    # sniff the package
    import zipfile
    data = source
    if not isinstance(source, (str, os.PathLike)):
        pos = source.tell()
    with zipfile.ZipFile(data) as z:
        names = z.namelist()
    if not isinstance(source, (str, os.PathLike)):
        source.seek(pos)
    if any(n.startswith("word/") for n in names):
        return "docx"
    if any(n.startswith("xl/") for n in names):
        return "xlsx"
    raise ValueError("Unrecognised document type")


def convert(source, target=None, *, kind: str | None = None, fonts: FontManager | None = None,
            **options):
    """Convert a DOCX or XLSX document to PDF.

    ``source`` is a path or binary file object. ``target`` is a path or binary
    stream; when omitted the PDF is returned as ``bytes`` (or written next to a
    path source when ``target`` is ``True``).

    Extra keyword options are passed to the format-specific converter (see
    :func:`convert_xlsx`).
    """
    k = _kind(source, kind)
    return_bytes = target is None
    if target is True:
        target = str(Path(source).with_suffix(".pdf"))
    out = io.BytesIO() if return_bytes else target
    if isinstance(out, os.PathLike):
        out = str(out)
    if k == "docx":
        from .docx import convert_docx
        convert_docx(source, out, fonts=fonts)
    elif k == "xlsx":
        from .xlsx import convert_xlsx
        convert_xlsx(source, out, fonts=fonts, **options)
    else:
        raise ValueError(f"Unsupported kind: {k}")
    if return_bytes:
        return out.getvalue()
    return target


def convert_docx(source, target, fonts: FontManager | None = None):
    from .docx import convert_docx as _c
    return _c(source, target, fonts=fonts)


def convert_xlsx(source, target, fonts: FontManager | None = None, **options):
    from .xlsx import convert_xlsx as _c
    return _c(source, target, fonts=fonts, **options)
