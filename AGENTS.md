# Notes for AI agents (and humans) maintaining oo2pdf

Start here: **[docs/maintenance/README.md](docs/maintenance/README.md)**.

The short version:

- oo2pdf converts DOCX/XLSX to PDF in pure Python, reproducing Microsoft
  Office's layout. "Correct" means **matching the PDF Word/Excel export**, not
  following the ECMA-376 spec literally; where they disagree, Office wins.
- Every layout rule in the engine was found by comparing against real Office
  output. Before changing a rule, read why it exists
  ([word-rules.md](docs/maintenance/word-rules.md)) and re-run the regression
  gate ([workflow.md](docs/maintenance/workflow.md)). Most "obvious fixes"
  break another document.
- Never commit test documents: `my_tests/` (personal files) and `tests/corpus/`
  (Microsoft's templates) are git-ignored on purpose.
- Run `python -m pytest` before and after any change (no Office needed).
