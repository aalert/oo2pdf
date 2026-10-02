# Excel printing rules

Excel prints through the printer driver, so its PDF depends on a device model,
not on the screen layout. oo2pdf models a virtual printer (default 600 dpi,
`WorkbookContext`, `xlsx/sheet.py`).

## Device model
- Column widths use the **printer's** maximum digit width (mdw) of the
  workbook's default font; the screen mdw (96 dpi GDI) defines the character
  unit. Default column width rounds up to 8 px.
- Row heights are scaled by `printer standard row height / screen standard row
  height`, with the sheet's `defaultRowHeight` as the standard.
- Glyphs use integer device advances at an integer ppem (GDI hdmx/VDMX):
  Calibri 11 prints at 11.04 pt. `DevFont`.
- Cell text box = ascent + descent + 8 device units, used for vertical
  alignment; left/right padding.

## Pagination
- Print areas, print titles, manual breaks, fit-to-page/scale, page order,
  centring, blank pages skipped, text overflow widens the printed range and
  continues across page boundaries.
- Headers/footers: `&P &N &D &T &F &A` and font codes (`parse_hf`).

## Cells
- Fills (solid/pattern, gradient approximation), all border styles, General
  alignment rules, indent = 3 spaces per level, wrap, shrink-to-fit, rotation,
  vertical text, overflow into empty neighbours, `####` for numbers that do not
  fit, General numbers reducing precision to fit.

## Number formats (`xlsx/numfmt.py`, `xlsx/locale.py`)
- Full format language: sections, conditions, colours, `_x`, `*x`, fractions,
  scientific, elapsed time.
- Locale-dependent built-ins (14, 37–40, 5–8) and date/time separators;
  `locale="system"` reads Windows settings via `GetLocaleInfoEx` exactly as
  Excel does.

## Not implemented yet (see status.md)
Conditional formatting, Excel table styles (ListObjects), charts, print areas
defined by formulas (e.g. `OFFSET(...)`), images in sheets beyond basics.
