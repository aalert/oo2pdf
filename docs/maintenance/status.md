# Status and backlog (as of 2026-10-01)

Metric: share of Office's words that oo2pdf places on the same page within
1 pt of Office's position (`corpus_score.py`).

## Word (DOCX)
| set | result |
|---|---|
| fixtures (basic, long, features) | all pages match, 100 % |
| sample-files.com: large (49 p), table, lists, image, template | all pages match, 100 % each |
| calibre demo.docx (8 p: tables, calendar, drop cap, floats, lists) | 8/8, 93 % |
| personal set (CVs, forms, a paper) | 98–100 % |
| InterpretingDeepPixels / TechnicalIntroduction (OpenEXR, equation-heavy) | 17/17, 21/21 pages; 68–69 % (equation spacing) |
| Office's 20 Word templates | 18/20 page counts, 68 % overall; most letters/resumes/reports 92–100 % |

## Excel (XLSX) — needs the most work
| Office template | pages (ours/Excel) | ≤ 1 pt |
|---|---|---|
| ExpenseReport | 1/1 | 78 % |
| TimeCard | 1/1 | 80 % |
| LoanAmortization | **11/1** | 87 % |
| PersonalMonthlyBudget | 1/1 | 60 % |
| SalesReport | **10/13** | 68 % |
| BillingStatement | 1/1 | 9 % |
| BloodPressureTracker | **1/2** | 3 % |

## Backlog (highest value first)
1. **Excel templates** above: LoanAmortization's print area is a formula
   (`OFFSET(Impresión_completa,0,0,Última_fila)`), not evaluated → prints the
   whole used range. SalesReport page breaks; BillingStatement /
   BloodPressureTracker very low — likely Excel table styles (ListObjects) or
   conditional formatting, both unimplemented.
2. Excel features: conditional formatting, table styles, charts.
3. Real-world Excel files were never tested (candidates and download links were
   collected: Microsoft "Conditional Formatting examples.xlsx", Contextures
   samples, file-examples.com).
4. Equations: stretched delimiters should use the font's size variants (MATH
   `MathVariants`) instead of scaling; equation-array row pitch can be ~1 pt
   off; Word adds ~0.2 pt more space after some subscripts (α_i, c_i but not
   S_i) — unexplained.
5. Word templates: AdjacencyReport (3/2 pages) and ApothecaryNewsletter (2/3)
   still have page-count differences; EssentialResume/OriginResume ~70–80 %.
6. sample-files template.docx content controls look fine now; RedAndBlackReport
   positions.
7. Footnote/endnote separator line is thinner/lighter than Word's.
8. Image-heavy PDFs are large (photos embedded at full resolution; Word
   recompresses).
9. `compare.py` always writes overlay PNGs; consider an `--images` flag.

## Not supported (documented in the README)
Complex-script shaping (Arabic, Indic), vertical East Asian text, charts,
WordArt effects, gradients/shadows, EMF/WMF off Windows, tracked-change markup,
comments, formula evaluation (XLSX must have cached values), fields other than
page numbers (cached results are shown).
