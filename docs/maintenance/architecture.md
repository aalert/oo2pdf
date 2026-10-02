# Architecture

~10 000 lines of Python in `src/oo2pdf`. Dependencies: lxml (XML), openpyxl
(XLSX model), reportlab (PDF writing), fontTools (font metrics, GPOS kerning,
MATH table), Pillow (images).

```
convert(source, target)                    src/oo2pdf/__init__.py
 ├─ .docx → docx/__init__.py: convert_docx
 │    Package (opc.py)            unzip + parse parts and relationships
 │    DocxReader (docx/reader.py) XML → document model (docx/model.py)
 │       ├ styles.py / props.py / numbering.py / theme.py   property resolution
 │       ├ math.py      (OMML equations, laid out at layout time)
 │       ├ vml.py       (legacy VML drawings/canvases)
 │       └ diagram.py   (SmartArt, from Word's cached drawing)
 │    LayoutContext + layout.py   paragraphs → lines (ParaFlow), tables → rows (TableFlow)
 │    Paginator (docx/paginate.py) flows → pages, floats, footnotes, headers/footers
 │    PdfWriter (draw.py)         drawing ops → PDF (reportlab)
 └─ .xlsx → xlsx/__init__.py: convert_xlsx
      load_workbook / SheetPrinter (xlsx/sheet.py)  print layout on a virtual printer
      numfmt.py / locale.py       number formats and regional settings
      PdfWriter (draw.py)
```

## Shared modules

- **opc.py** — ZIP package access: `Package.xml(part)` (cached lxml trees),
  `Package.rels(part)` → `{rId: (type suffix, absolute target, external)}`.
  Namespace constants (`W`, `A`, `M`, `V`, `WP`, `WPS`, `WPG`, …).
- **fonts.py** — `FontIndex` (scans system font dirs, cached on disk),
  `Font` (metrics exactly as Office uses them: Windows ascent/descent, GDI
  external leading, `USE_TYPO_METRICS`, glyph bounds, pair kerning from GPOS,
  `ssty` script variants, italic correction), `FontManager.face(family, bold,
  italic)` with fallback/substitution. Also patches reportlab's ToUnicode CMap
  to write UTF-16 surrogate pairs (astral math letters must stay searchable).
- **draw.py** — drawing ops (`TextOp`, `LineOp`, `RectOp`, `ImageOp`,
  `PolyOp`, `ClipOp`, `TransformOp`, `LinkOp`, `AnchorOp`); `move_ops` shifts
  op lists; `PdfWriter` renders them. Coordinates are points, origin top-left.
- **theme.py** — theme colours/fonts, `drawingml_color` (scheme colours with
  lumMod/lumOff/tint/shade).

## DOCX pipeline in detail

### Reader (`docx/reader.py`, ~1100 lines)
- `read_document()` → `[Section]`, each with `blocks` (Paragraph | Table) and
  page setup. Headers/footers/footnotes/endnotes are read into side tables.
- `para_props()` resolves paragraph/run properties in Word's order: docDefaults
  → table style → paragraph style → numbering → direct (with the
  table-style-vs-Normal quirks, see word-rules.md).
- `read_paragraph()` turns runs into model items: `Text`, `Tab`, `Break`,
  `Field` (page numbers), `Picture`, `Floating`, `MathItem`, `FootnoteRef`,
  `Bookmark`. Complex fields are tracked in `self._fields`.
- `_blocks()` post-processes the block stream: hidden-mark paragraphs join the
  next one (style separators), empty hidden paragraphs before tables vanish,
  drop-cap paragraphs become a `Floating` frame on the next paragraph.
- `read_table()` builds `Table/Row/Cell`, applying table-style conditional
  formatting per cell (`_cond_edges` maps region borders to cell edges).
- `_drawing()` (DrawingML pictures/shapes/groups/SmartArt) and `_vml()` (VML)
  produce `Picture` (inline) or `Floating` (anchored) items.

### Layout (`docx/layout.py`, ~1900 lines)
- `AtomBuilder` turns items into `Atom`s (text runs split at break
  opportunities, spaces, tabs, boxes, math pieces, floats, anchors).
- `ParagraphLayout.layout(bounds)` breaks atoms into lines (`_LineBuilder`),
  justifies (Word 2013 "smart" justification), and `_finish()` computes line
  height/baseline and `_emit()` produces drawing ops. `bounds` narrows lines
  around floating objects (text wrapping).
- `layout_table()` computes grid widths (incl. auto-fit), cell content, border
  conflict resolution and border bands, row heights → `TableFlow` of
  `RowFlow`s. `RowFlow.render()` draws fills, content and borders;
  `RowFlow.split()` splits a row across pages.
- `render_shape()` draws shapes/text boxes/groups.

### Pagination (`docx/paginate.py`, ~1000 lines)
- `Paginator.run(sections)` places flows on pages: paragraph placement with
  page-top spacing modes, widow/orphan/keep rules (`_place_para`), tables with
  row splitting and page-break borders (`_place_table`), floats and text
  wrapping (`_place_float`, `_wrap_lines`), footnotes, columns with balancing
  (`_flow_balanced` uses snapshot/restore and bisection).
- **Multi-pass floats**: `convert_docx()` runs the paginator repeatedly. Wrapping
  floats found in pass *n* (`anchor_pages`) are pre-placed at their page's start
  in pass *n+1* (`prefloats`), because Word lays out a whole page around them
  (text above the anchor wraps too). Pages settle in document order; a float
  that oscillates is frozen at its first position (what Word does).
- `render_headers()` lays out headers/footers per page afterwards (they need
  NUMPAGES).

## XLSX pipeline (`xlsx/sheet.py`, ~1300 lines)
- `WorkbookContext` holds the virtual printer model (default 600 dpi): column
  widths from the printer's maximum digit width, row heights scaled by
  printer/screen standard row height, `DevFont` with integer device advances.
- `SheetPrinter` computes the print range (print areas, used range, overflow),
  page breaks (manual, fit-to-page, scale, order), and renders cells (fills,
  borders, alignment, overflow, `####`), headers/footers (`parse_hf`).
- `numfmt.py` implements Excel's number format language; `locale.py` provides
  regional settings (`"system"` reads Windows via `GetLocaleInfoEx`).

## Tests and tools
- `tests/test_*.py` — pytest; build documents with python-docx/openpyxl and
  check invariants. No Office needed.
- `tests/fixtures/` — committed fixtures (generated by `tests/tools/make_fixtures.py`,
  `features.docx` authored in Word by `word_features.ps1`) plus Office reference PDFs.
- `tests/tools/` — fidelity tooling, see [workflow.md](workflow.md).
