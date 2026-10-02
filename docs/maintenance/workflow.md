# Workflow: investigating differences and checking changes

Fidelity work needs Windows with Microsoft Office (Word/Excel) to export
reference PDFs. The library itself never needs Office.

## Folder conventions
- `<folder>/<name>.docx|xlsx` — documents under test
- `<folder>/reference/<name>.pdf` — the PDF **Office** exported
- `<folder>/out/<name>.pdf` — our output (written by `corpus_score.py`)

Used folders: `tests/fixtures` (committed), `tests/corpus` (Office's own
templates, built locally by `corpus_reference.ps1`, git-ignored), `my_tests`
(any local documents, git-ignored — the maintainer's personal set).

Reference PDFs: `powershell -File tests/tools/office_reference.ps1 <folder>`,
or export by hand in Word (File → Save As → PDF) / Excel (Export → PDF, "Entire
workbook").

## Tools (`tests/tools`, all dev-only)
| tool | what it does |
|---|---|
| `corpus_score.py <folder> [filter...]` | converts every document, prints pages (ours/ref), words found, same-page %, **≤ 1 pt %**; the main metric |
| `linebreaks.py <folder> [filter] -v` | lines starting with a different word than Office — the first sign of drift |
| `sidebyside.py <folder> <name> [pages] --dpi N` | Office page and ours side by side as PNG in `<folder>/out/compare/` — look at them |
| `lines.py ref.pdf ours.pdf <page>` | baselines + x of every line, side by side |
| `docx_coverage.py <folder> -v` | every word/picture of the docx present in our PDF (no Office needed) |
| `compare.py file` | per-page deltas + overlay PNGs |
| `missing.py`, `dump.py`, `squeeze.py` | word diff, XML structure dump, justification squeeze analysis |
| `experiment.py <name>` | build a test document, export with Word/Excel and with oo2pdf, print both line lists |
| `make_fixtures.py`, `word_features.ps1` | regenerate `tests/fixtures` |
| `corpus_reference.ps1` | copy Office's bundled templates into `tests/corpus` and export references |

## Investigation loop (what worked)
1. `corpus_score.py` → pick the worst document; `linebreaks.py -v` and
   `sidebyside.py` → find the **first** place it differs (later differences are
   often consequences).
2. Measure: `lines.py`, or a few lines of pymupdf (`page.get_text("words")`,
   `get_text("rawdict")` for glyph positions, `get_drawings()` for borders and
   fills, `get_image_rects`). Get exact numbers for both PDFs.
3. Read the XML of that spot (unzip and print the relevant `<w:p>`/`<w:tbl>`),
   including styles, numbering, settings (`w:compat`!).
4. Form a hypothesis that explains the numbers. If it is not obvious, write a
   small experiment: generate variants with python-docx, export them through
   Word COM, measure (see `experiment.py` for the pattern). This settled the row
   splitting, border band, and bullet line-height rules.
5. Implement with a comment naming the rule and the document; add a unit test
   in `tests/test_layout_rules.py` or `tests/test_units.py`.
6. Run the regression gate.

## Regression gate
Before and after a change, save and diff the per-document scores:

```bash
python -m pytest -q
python tests/tools/corpus_score.py tests/fixtures | awk '{print $1,$3,$(NF-3)}' > /tmp/f1.txt
python tests/tools/corpus_score.py tests/corpus   | awk '{print $1,$3,$(NF-3)}' > /tmp/c1.txt
python tests/tools/corpus_score.py my_tests       | awk '{print $1,$3,$(NF-3)}' > /tmp/m1.txt
# ... change ...  then the same into f2/c2/m2 and:
diff /tmp/c1.txt /tmp/c2.txt; diff /tmp/m1.txt /tmp/m2.txt
```

A change is acceptable only if no document gets worse (page counts and ≤ 1 pt
%), unless the reason is understood and documented. When a fix helps one
document and hurts another, the rule is incomplete — find the distinguishing
condition (compat mode, wrap type, footer vs margin, paragraph end…) instead of
tuning numbers.

## Gotchas
- **Word COM from PowerShell hangs** if paths are PSObject-wrapped strings
  (from `Join-Path`/`Resolve-Path`): cast with `[string]`. Kill stray
  `WINWORD.EXE`/`EXCEL.EXE` processes after an interrupted export.
- Export references with the **same fonts installed** as the machine running
  oo2pdf; missing fonts change everything.
- Compat mode matters: documents without `w:compatSetting` but with `<w:compat/>`
  are Word 2007 mode (`compat_mode` 12) — table indents, Normal-in-table sizes
  and justification differ from compat 15.
- pymupdf text extraction merges words when a space is not a real character;
  oo2pdf writes stretched spaces as characters for this reason.
- Positions in Word's PDF are quantized (~0.12 pt); differences below that are
  noise. Word-width differences of ~0.1 pt per word are measurement noise too
  (tested: no rounding model fits better than exact widths).
- Several PDFs show "?" for astral characters in a cp1252 console; set
  `PYTHONIOENCODING=utf-8` when printing extracted text.
