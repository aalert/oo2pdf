"""XLSX -> PDF conversion following Excel's print layout."""
from __future__ import annotations

from ..draw import PdfWriter
from ..fonts import FontManager
from .sheet import SheetPrinter, load_workbook


def convert_xlsx(source, target, fonts: FontManager | None = None, sheets=None,
                 ignore_print_areas: bool = False, active_sheet_only: bool = False,
                 locale="en-US", printer_dpi: int = 600) -> int:
    """Convert a workbook to PDF.

    ``sheets``: iterable of sheet names (or indexes) to print; by default every
    visible sheet is printed, like Excel's "Print Entire Workbook".
    ``locale``: regional settings used to display numbers and dates ("en-US",
    "es-CL", ..., or "system" to use the operating system's settings like Excel).
    ``printer_dpi``: resolution of the virtual printer Excel would print through.
    """
    fonts = fonts or FontManager()
    wb, ctx = load_workbook(source, fonts, printer_dpi, locale)
    if active_sheet_only:
        selected = [wb.active]
    elif sheets is not None:
        selected = []
        for s in sheets:
            selected.append(wb.worksheets[s] if isinstance(s, int) else wb[s])
    else:
        selected = [ws for ws in wb.worksheets if ws.sheet_state == "visible"
                    and ws.__class__.__name__ == "Worksheet"]
    pages = []
    for ws in selected:
        printer = SheetPrinter(ws, ctx, ignore_print_areas=ignore_print_areas)
        pages.extend(printer.pages())
    title = wb.properties.title if wb.properties else None
    author = wb.properties.creator if wb.properties else None
    writer = PdfWriter(target, title=title, author=author)
    total = len(pages)
    for i, page in enumerate(pages):
        ops = page.render(i + 1, total)
        writer.page(page.width, page.height, ops)
    if not pages:
        writer.page(612, 792, [])
    writer.save()
    return max(1, total)
