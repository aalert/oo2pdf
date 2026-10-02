# Word layout rules (and why they are the way they are)

Each rule below was observed in PDFs exported by Word (Microsoft 365, Windows)
and, where marked **[exp]**, confirmed with a controlled experiment
(`tests/tools/experiment.py` or a one-off script building variants and exporting
them through Word COM). "Where" gives the function holding the rule. Documents
named here are the ones that revealed it.

## Fonts and line height
- Line height = Windows metrics: usWinAscent + usWinDescent + GDI external
  leading (hhea lineGap logic); typo metrics when `USE_TYPO_METRICS` is set
  (Aptos). Where: `fonts.Font`, `AtomBuilder`.
- "Multiple" line spacing adds `(factor − 1) × text height` **below** the text;
  exact spacing puts the baseline at 80 % of the line box. Where:
  `ParagraphLayout._finish`.
- Inline pictures are never scaled by line spacing; picture-only lines have no
  descent; the paragraph-mark font only sizes empty/picture-only lines.
- **List labels (bullets/numbers) never deepen a line, and only their ascent
  above the text (incl. its external leading) adds height, unscaled by line
  spacing.** [exp: Calibri/Cambria × single/1.15 × Symbol/Wingdings/Courier
  bullets; matched to 0.02 pt]. Where: `_finish` (`label_top`). Found in
  demo.docx (calibre).
- Pair kerning only when `w:kern` is on and the size ≥ its threshold; GPOS
  PairPos only (contextual/extension subtables threw and silently disabled
  kerning for Calibri once).

## Paragraph spacing and page tops
- Spacing between paragraphs = `max(after, before)` unless the compat option
  `doNotUseHTMLParagraphAutoSpacing` (then summed). `gap()` in layout.py.
- Page top: space before is dropped after natural/manual page breaks; kept as
  `before − previous after` after page-break-before and section breaks.
  `Paginator._space_before`.
- Widow/orphan: a paragraph of ≤ 3 lines is not split; keep-with-next chains.
- **An empty paragraph at the bottom of a page may let its "multiple" spacing
  extra hang past the bottom** (only the text must fit). Lines with text may
  not. `Line.trailing`, `_place_para`. Found in sample-files image-document
  (page 2 top margin).
- **When the footer is taller than the bottom margin, the body may reach
  0.25 pt into it** (`FOOTER_OVERLAP`). Found in StudentReport.dotx. A general
  tolerance broke InterpretingDeepPixels, so it is footer-only.

## Line breaking and justification
- Greedy breaking at Word's break opportunities; Word 2013+ "smart"
  justification may squeeze spaces to fit one more word (compat ≥ 15).
- Inline equations break **after relation operators** (the operator stays at
  the end of the line) and **before binary operators** (Word's default
  `brkBin="before"`); only at the equation's top level. A relation's right
  spacing hangs past the margin like a trailing space (`Atom.hang`).
- **At an equation's edges, only a real text space is a break opportunity.** A
  space typed inside the equation (`and<m:t> </m:t>S…`) glues the preceding
  word to the equation. InterpretingDeepPixels "front / and Sᵢ(ZBack)".
- An equation ending in an italic letter keeps its italic correction before
  the following text.

## Equations (OMML, `docx/math.py`)
- Metrics from Cambria Math's MATH table (script scale 73/60 %, shifts, axis,
  fraction gaps); `ssty` variants for script glyph widths; italic correction.
- Operator spacing: binary 4/18 em, relation 5/18 em, punctuation 3/18 em.
  Relations include all Unicode arrows (U+2190–21FF, 27F0–27FF, 2900–297F) —
  U+27FC "⟼" missing once cost a page break.
- Letters are math italic by default; **lowercase Greek too** (U+1D6FC…),
  capital Greek stays upright.
- Delimiters grow when the content's **ink** is taller than the delimiter
  glyph's ink (default `m:grow` on). We scale the glyph; Word picks size
  variants (see status.md).
- Display equations (`m:oMathPara`) are centred; inline equations size the line
  by base metrics, display ones by their box.
- Invisible operators U+2061–2064 are not drawn.

## Styles and table styles
- Order: docDefaults → table style → paragraph style → numbering → direct.
- **Normal style inside a table with a table style:** Normal's *paragraph*
  properties beat the table style's; the table style's *character* properties
  beat Normal's; and without the compat setting
  `overrideTableStyleFontSizeAndJustification`, Normal's font size is not used
  in the table at all (demo.docx: Normal 12 pt, table text 11 pt default). With
  that setting (all modern documents) a Normal size other than 11 pt wins.
  Where: `DocxReader.para_props`. Got this wrong twice — see rejected
  hypotheses.
- Table-style conditional borders describe a **region**: for row regions
  (header row, bands) `left/right` are the row's outer edges and `insideV`
  the edges between its cells; for column regions `top/bottom` vs `insideH`.
  `_cond_edges` (sample-files table-document, header without dividers).

## Tables
- Border conflicts: the heavier border wins. A horizontal edge reserves its
  width once (lower row when any cell defines its own top border, else upper);
  an edge defined only by the upper cell fits into the lower cell's top margin
  when that margin is ≥ the border width [exp]. Where: `layout_table` pass 2.
- Legacy compat (< 15): a table at indent 0 is shifted left by the cell margin
  so text aligns with the margin; floating tables too (`tblpX` + `t.x0`).
- **Auto-fit tables** (`tblLayout` not fixed, width not pct): a column whose
  longest unbreakable word (+ cell margins) exceeds its preferred width grows
  to it; the others give up the excess in proportion to their room above their
  own minimum. Available width in legacy compat = text width + left/right cell
  margins. Indents and equations are not measured. `_autofit_grid`
  (sample-files large-document 7.2, matched within 0.5 pt).
- Row splitting across pages:
  - a row whose `trHeight` minimum (+ its border bands) exceeds the room left
    is moved whole [exp: 15/35/54 pt minimums];
  - a row splits only if **every** cell can put its first line on the page;
  - a line ending its paragraph in the first part needs the paragraph's space
    after too (a middle line does not);
  - the continuation keeps the rest of the cell's height incl. space after.
  Where: `RowFlow.split`, `Paginator._place_table`, `_row_starts`.
- **At a page break between rows each row keeps its own border**: the upper
  row is closed with its own bottom border (drawn and reserved — one row fewer
  may fit), the lower row opens the next page with its own top border.
  `_closing_border`, `_row_at_page_top`, `CellBox.own`.
- Border joins: horizontal borders run to the outer edges of the vertical
  borders they meet; vertical borders stop at horizontal ones; table borders use
  flat line caps (square caps overshoot). `RowFlow.render`, `_border_line(cap=)`.
- A vertically merged cell keeps its first row's shading (continuation pieces
  overlap by 0.5 pt so viewers show no seam).
- Floating tables (`tblpPr`): distances from text; wrapped text keeps the
  paragraph's first-line indent relative to the table edge.

## Floating objects and wrapping
- Wrapping text treats the object's edge as a new margin: indents count from
  the edge. `ParagraphLayout.layout` → `new_line`.
- tight/through wrapping ignores `distT/distB` (the wrap polygon replaces them);
  square wrapping uses them. `Paginator._place_float`.
- A paragraph's **last line** is tested against floats including the
  paragraph's space after (a picture just below narrows the line above it).
  `_wrap_lines`.
- Word lays out the whole page around wrapping floats → multi-pass placement
  (architecture.md). When wrapping the text above moves the anchor, Word keeps
  the float at its first position (sample-files image-document page 2).
- Top-and-bottom floats anchored to a paragraph are placed when reached (not
  pre-placed): pre-placing broke demo.docx page 7.

## Other features
- Drop caps (`framePr dropCap`): a frame `lines` × the paragraph's line height
  tall, as wide as the letters, positioned at the drop-cap paragraph's indent;
  the frame paragraph itself has no indents. `_dropcap_float`, `_size_dropcap`.
- TOC entries carry the Hyperlink character style but Word does not show it
  inside a TOC field result.
- `PAGE \# 0#` numeric pictures (`_numeric_picture`).
- Picture bullets: scaled to 0.83 × the text size, bottom on the baseline.
- Endnotes follow the endnote separator after the text.
- SmartArt: drawn from `word/diagrams/drawingN.xml` (Word's cached rendering),
  linked via `dsp:dataModelExt` in the data part.
- VML paths: empty parameters mean 0 (`m,1008`), curves are flattened.

## Rejected hypotheses (don't re-try these without new evidence)
- "Word never breaks before an inline equation" — wrong; it breaks at a real
  space before one (demo page 6). The real rule is the glue rule above.
- "Table style overrides Normal for everything" / "Normal overrides table style
  for everything" — both broke documents; it is split by property kind.
- "Widow/orphan control applies inside table cells" — Word splits a 2-line cell
  paragraph (large-document page 15/16); the row-move behaviour is the
  every-cell-must-start rule.
- "A flat bottom tolerance for the last line" — broke InterpretingDeepPixels;
  only the footer overlap is real.
- "Keep a float's first position always" — stale positions from early passes;
  only freeze on oscillation, settling pages in order.
- Space-after required for every line ending a split first part — only for a
  line that ends its paragraph.
