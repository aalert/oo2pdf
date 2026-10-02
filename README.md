# oo2pdf

> **Disclaimer — AI-generated code.** This project was written entirely by an AI
> (Anthropic's Claude) and validated by a human, who compared its output page by
> page against PDFs exported by Microsoft Word and Excel. It comes with no
> warranty of any kind. It is free software: anyone may use it, fork it, modify
> it and redistribute it (see [LICENSE](LICENSE)).

Pure-Python DOCX/XLSX → PDF conversion that reproduces Microsoft Office's
layout rules. It needs no Word, Excel or LibreOffice at runtime: the engine
parses OOXML, lays text out with the same metrics Office uses, and writes the
PDF with reportlab, embedding subsetted TrueType fonts.

The name stands for **O**ffice **O**pen XML (OOXML, the ECMA-376 / ISO/IEC 29500
standard behind .docx and .xlsx files) **to PDF**.

## Installation

```bash
git clone <this repository>
cd py_office_pdf_engine
pip install .
```

Requires Python ≥ 3.10. Dependencies (installed automatically): `lxml`,
`openpyxl`, `reportlab`, `fonttools`, `pillow`. The fonts used by the documents
must be available on the machine (see *Known limitations*).

## Usage

### As a library

```python
import oo2pdf

# Word document -> PDF file
oo2pdf.convert("report.docx", "report.pdf")

# Excel workbook -> PDF, with Excel options
oo2pdf.convert(
    "book.xlsx", "book.pdf",
    locale="system",        # number/date formats as on this PC (or "en-US", "es-CL", ...)
    sheets=["Summary"],     # only these sheets
)

# In memory: bytes in, bytes out (the format comes from `kind`)
with open("letter.docx", "rb") as f:
    pdf_bytes = oo2pdf.convert(f, kind="docx")

# Many files: share one FontManager so the font index is loaded once
fonts = oo2pdf.FontManager()
for name in ("a.docx", "b.docx"):
    oo2pdf.convert(name, name.replace(".docx", ".pdf"), fonts=fonts)
```

### From the command line

```bash
oo2pdf report.docx                      # writes report.pdf next to it
oo2pdf report.docx book.xlsx -o out/    # several files into a folder
oo2pdf book.xlsx --locale es-CL --sheet Summary -o summary.pdf
python -m oo2pdf --help                 # all options
```

(`python -m oo2pdf ...` works the same when pip's scripts folder is not on
your `PATH`.)

## How close is it to Office?

The development tooling in `tests/tools` compares our output with PDFs that
Word/Excel export from the same files, matching every word and measuring its
position. On the fixtures and on Office's own shipped templates:

| document | pages | words within 1 pt of Office |
|---|---|---|
| `long.docx` (8 pages of mixed spacing, headings) | 8/8 | 100 % (max error 0.29 pt) |
| `features.docx` (Word-authored: footnotes, lists, tables with repeated headers, columns, sections, text box) | 6/6 | 100 % (max 0.36 pt) |
| `basic.xlsx` (styles, merges, number formats, print titles) | 7/7 | ~100 % (median 0.05 pt) |
| 20 Word templates shipped with Office (letters, resumes, reports, newsletter) | 18/20 same page count | 11 templates ≥ 97 %, 3 more ≥ 78 % |
| 7 Excel templates shipped with Office | 4/7 same page count | TimeCard 82 %, LoanAmortization 87 %; see limitations |

### Office rules the engine implements (each verified against Office)

**Word**

- Line height from the font's Windows metrics (usWinAscent + usWinDescent + GDI
  external leading), or typographic metrics when `USE_TYPO_METRICS` is set
  (Aptos). "Multiple" spacing adds its extra below the text; with exact
  spacing the baseline sits at 80 % of the line box.
- Paragraph spacing collapses to `max(after, before)` unless the document sets
  `doNotUseHTMLParagraphAutoSpacing`; contextual spacing; auto spacing.
- The paragraph mark's font only sizes empty/picture-only lines; inline
  pictures aren't scaled by line spacing and picture-only lines have no descent.
- Page top rules: space before is dropped after natural and manual page
  breaks, kept as `before − previous after` after page-break-before and
  section breaks; a trailing page break doesn't carry an empty line over.
- Widow/orphan control (a paragraph of ≤ 3 lines can't split), keep with next,
  keep lines together, page break before.
- Greedy line breaking with Word's break opportunities, left/center/right/
  decimal tabs with leaders, default tab stops, hanging-indent tab stops,
  justification (including lines ending in manual breaks), pair kerning
  (GPOS/`kern`) when `w:kern` is on.
- Styles: docDefaults → table style → paragraph style → numbering → direct
  formatting, with Word's table-style/Normal precedence quirk; theme fonts and
  colours; list numbering (all common formats, legal numbering, restarts).
- Tables: grid widths, spans, vertical merges, cell margins, table-style
  conditional formatting (header row, banding, first column...), border
  conflict resolution (heavier border wins) with border widths reserving row
  height (an edge defined only by the upper cell fits inside the lower cell's
  top margin when that is wide enough), compat-15 border alignment, auto-fit
  columns widened to their longest word, Word's border joins at cell corners,
  vertically merged cells shaded as one cell, repeated header rows, rows
  splitting across pages (only when every cell can start on the page, keeping
  each row's own border at the break), floating tables (`tblpPr`), nested
  tables (the empty end-of-cell paragraph after a nested table takes no room).
- Sections: page size/orientation, margins, columns with Word's column
  balancing at continuous breaks, different first/odd/even headers and footers,
  PAGE/NUMPAGES/SECTIONPAGES fields, page-number formats, page borders.
- Footnotes (placed at the bottom of the referencing page with the document's
  separator) and endnotes.
- Pictures (inline and anchored), DrawingML shapes, text boxes (auto-fit,
  rotated `vert270`), groups, wp14 percentage positions/sizes, VML fallbacks
  and legacy VML drawing canvases (`v:group` with rectangles, text boxes,
  connectors with arrowheads/dashes and paths), hyperlinks and bookmarks as PDF
  links.
- Drop caps (framePr dropCap) as a frame the paragraph wraps around; SmartArt
  diagrams drawn from the cached drawing Word saves with them; picture bullets;
  character borders; floats anchored to the page/margin also push away the text
  above their anchor (as Word lays out the whole page around them).
- Equations (OMML): fractions, scripts, delimiters that stretch around taller
  content, matrices/equation arrays, n-ary operators, radicals, accents, using
  Cambria Math's MATH table; display equations are centred.

**Excel**

- Print layout on a virtual printer (600 dpi by default): column widths use the
  printer's maximum digit width, row heights are scaled by the ratio of the
  printer's to the screen's standard row height, glyphs use integer device
  advances at an integer ppem (e.g. Calibri 11 prints at 11.04 pt).
- Default column width rounded up to 8 px, auto-fit row heights.
- Pagination: print areas, print titles, manual breaks, fit-to-page and scale,
  page order, centring, headers/footers with `&P &N &D &T &F &A` and font
  codes; blank pages are skipped; text overflow widens the printed range and
  continues across page boundaries.
- Cells: fills (solid/pattern/gradient approximation), all border styles,
  alignment (general alignment rules, indent = 3 spaces per level, wrap,
  shrink-to-fit, rotation, vertical text), Excel's text line box (ascent +
  descent + 8 device units) for vertical alignment, left/right padding, text
  overflow into empty neighbours, `####` for numbers that don't fit, General
  numbers shrinking their precision to fit.
- Number formats: the full format language (sections, conditions, colours,
  `_x` padding, `*x` fill, fractions, scientific, elapsed time, locale
  built-ins such as format 14 and 37–40) with regional settings (`locale=`
  `"en-US"`, `"es-CL"`, `"de-DE"`, ..., or `"system"` to read Windows' settings
  exactly as Excel does). Theme fonts (`scheme="minor"`) and theme/indexed
  colours with tints.

### Known limitations

- Text wraps around square/tight/through objects using their bounding box on
  the wider side (Word can also flow text on both sides or follow a tight
  polygon). Complex-script shaping (Arabic, Indic) and vertical East Asian text
  are not implemented.
- Equations: stretched delimiters are scaled rather than using the font's size
  variants, and array row spacing can differ from Word by up to ~1.5 pt per row.
- Charts, WordArt effects, shape gradients/shadows, EMF/WMF pictures on
  non-Windows systems are not rendered (SmartArt is drawn from the copy Word
  caches in the file). In XLSX, conditional formatting,
  Excel table styles (ListObjects), charts and print areas defined by formulas
  (e.g. `OFFSET(...)`) are not supported yet.
- Excel's printer row scaling is exact for most fonts; a few (e.g. Gill Sans MT,
  Segoe UI) print rows up to ~0.5 pt taller in Excel than modelled here.
- Formulas are not evaluated: XLSX files must contain cached values (anything
  saved by Excel does).
- Fields other than page numbers show their cached result (as Word prints
  them when not updated).
- Output depends on the fonts installed. Missing fonts fall back to metric-compatible
  substitutes (Carlito, Liberation, ...). Office cloud fonts such as Aptos are
  picked up from Office's font cache when present. Extra directories can be
  added with `OO2PDF_FONT_DIRS` or `FontManager(font_dirs=[...])`.

## API

`oo2pdf.convert(source, target=None, *, kind=None, fonts=None, **options)`
- `source`: path or binary file object; `kind` (`"docx"`/`"xlsx"`) is inferred
  from the extension or the package content.
- `target`: path or binary stream; when omitted the PDF is returned as bytes.
- XLSX options: `sheets=[...]`, `ignore_print_areas=True`,
  `active_sheet_only=True`, `locale="en-US"|"system"|...`, `printer_dpi=600`.
- `fonts`: a shared `oo2pdf.FontManager` (reuse it across conversions for
  speed; the font index is cached on disk).

## Development

Maintainers (people or AI agents): start with [AGENTS.md](AGENTS.md) and the
[maintenance guide](docs/maintenance/README.md) — architecture, every Office
layout rule with its evidence, the investigation workflow and regression gate,
and the current status/backlog.

```bash
pip install -e .[dev]
python -m pytest
```

Or with [uv](https://docs.astral.sh/uv/), using the pinned versions in `uv.lock`:

```bash
uv sync --extra dev
uv run pytest
```

The unit tests need neither Office nor any external document: they build small
documents with python-docx/openpyxl and check the layout rules listed above.

Fidelity testing compares against PDFs exported by Office itself. Those test
documents are not part of the repository (Office's templates are Microsoft's
content and personal test files stay local); the tooling below recreates them:

Fidelity tooling (needs Word/Excel on Windows, only for development):

- `tests/tools/make_fixtures.py` builds fixtures; `office_reference.ps1`
  exports Office reference PDFs; `word_features.ps1` authors a feature document
  in Word itself; `corpus_reference.ps1` builds a local corpus from Office's
  bundled templates.
- `tests/tools/compare.py file.docx` prints per-page word position deltas and
  writes overlay images; `corpus_score.py` scores a whole corpus;
  `experiment.py` runs targeted layout experiments; `lines.py`, `missing.py`,
  `dump.py` help diagnose differences.
- `tests/tools/docx_coverage.py folder -v` checks, without Office, that every
  word of each .docx (body, tables, text boxes, headers/footers, notes,
  equations) and every picture made it into `folder/out/<name>.pdf`.
- `tests/tools/sidebyside.py folder name [pages] [--dpi N]` renders Office's
  and our pages side by side into `folder/out/compare/` for visual review.
- `tests/tools/linebreaks.py folder -v` lists the lines that start with a
  different word than in Office's PDF (the first sign of a line/page drift).

## License

MIT — see [LICENSE](LICENSE). Use it, fork it, change it, ship it.
