# Maintenance guide

This folder is written for whoever maintains oo2pdf next — another AI session or
a person. It records what the code does, *why* (the Office behaviour each rule
reproduces and how it was verified), how to test a change, and what is still
open.

| file | read it when |
|---|---|
| [architecture.md](architecture.md) | you need to find where something happens |
| [word-rules.md](word-rules.md) | you are about to change Word layout behaviour |
| [excel-rules.md](excel-rules.md) | you are about to change Excel printing behaviour |
| [workflow.md](workflow.md) | you investigate a difference with Office, or need the regression gate |
| [status.md](status.md) | you want to know how good it is today and what to work on next |

## Golden rules

1. **Office's PDF is the ground truth.** The goal is that our PDF matches the one
   Word/Excel export ("File → Save as PDF" / COM `ExportAsFixedFormat`) word by
   word, within 1 pt. The ECMA-376 spec is a guide; where Office differs, follow
   Office.
2. **Measure, don't guess.** When a document differs, find the exact numbers
   (positions, widths, row heights) in both PDFs, form a hypothesis that explains
   them, and when possible confirm it with a small Word/Excel experiment
   (`tests/tools/experiment.py`) before coding it. Several plausible rules in the
   history turned out wrong on the next document (see "Rejected hypotheses" in
   [word-rules.md](word-rules.md)).
3. **Every change passes the regression gate** ([workflow.md](workflow.md#regression-gate)):
   unit tests + the fixture/corpus scores must not get worse anywhere.
4. **Keep test documents out of git.** Office templates are Microsoft content;
   personal documents are private. Only the generated fixtures in
   `tests/fixtures` (and their Office reference PDFs) are committed. Scrub
   author metadata (`docProps/core.xml`) before committing any new document.
5. **Comment the why.** Rules in the code carry a short comment naming the
   behaviour and, when it came from a specific document, which one. Keep that
   habit: it is what makes the rules maintainable.

## Ten-minute orientation

```bash
pip install -e .[dev]          # editable install with test tools
python -m pytest               # ~60 tests, a few seconds, no Office needed
python -m oo2pdf tests/fixtures/basic.docx -o /tmp/out/   # try it
```

Then read [architecture.md](architecture.md) and skim the section headers of
[word-rules.md](word-rules.md).
