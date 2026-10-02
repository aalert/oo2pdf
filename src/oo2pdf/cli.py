"""Command line interface: ``python -m oo2pdf input.docx [-o output.pdf]``."""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

from . import FontManager, __version__, convert


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="oo2pdf",
                                 description="Convert DOCX/XLSX files to PDF like Microsoft Office.")
    ap.add_argument("inputs", nargs="+", help="input .docx/.xlsx files or directories")
    ap.add_argument("-o", "--output", help="output file (single input) or directory")
    ap.add_argument("--font-dir", action="append", default=[], help="extra font directory (repeatable)")
    ap.add_argument("--locale", default="en-US",
                    help='XLSX regional settings, e.g. en-US, es-CL, de-DE or "system"')
    ap.add_argument("--sheet", action="append", help="XLSX: sheet name to print (repeatable)")
    ap.add_argument("--ignore-print-areas", action="store_true", help="XLSX: print the used range")
    ap.add_argument("--dpi", type=int, default=600, help="XLSX: virtual printer resolution")
    ap.add_argument("--version", action="version", version=f"oo2pdf {__version__}")
    args = ap.parse_args(argv)

    files: list[Path] = []
    for item in args.inputs:
        p = Path(item)
        if p.is_dir():
            files += sorted(x for x in p.iterdir() if x.suffix.lower() in (".docx", ".xlsx", ".docm", ".xlsm")
                            and not x.name.startswith("~$"))
        else:
            files.append(p)
    if not files:
        ap.error("no input files")
    out = Path(args.output) if args.output else None
    fonts = FontManager(font_dirs=args.font_dir or None)
    failures = 0
    # inputs sharing a name (report.docx + report.xlsx) keep their extension
    stems = [f.stem.lower() for f in files]
    for f in files:
        name = f.stem + ".pdf" if stems.count(f.stem.lower()) == 1 else f"{f.stem}_{f.suffix[1:]}.pdf"
        if out is None:
            target = f.with_name(name)
        elif len(files) == 1 and out.suffix.lower() == ".pdf":
            target = out
        else:
            out.mkdir(parents=True, exist_ok=True)
            target = out / name
        opts = {}
        if f.suffix.lower() in (".xlsx", ".xlsm"):
            opts = {"locale": args.locale, "sheets": args.sheet,
                    "ignore_print_areas": args.ignore_print_areas, "printer_dpi": args.dpi}
        t0 = time.perf_counter()
        try:
            convert(f, target, fonts=fonts, **opts)
        except Exception as exc:  # keep going with the other files
            failures += 1
            print(f"FAILED {f}: {exc}", file=sys.stderr)
            continue
        print(f"{f} -> {target} ({time.perf_counter() - t0:.2f}s)")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
