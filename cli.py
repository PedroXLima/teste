"""Command-line conversion: python cli.py input.docx [output.pdf]"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from app.converter import docx_to_html, docx_to_pdf_bytes


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert a DOCX into a styled institutional PDF."
    )
    parser.add_argument("input", help="Path to input .docx file")
    parser.add_argument("output", nargs="?", help="Path to output .pdf file (default: same name, .pdf)")
    parser.add_argument("--html", action="store_true", help="Write the intermediate HTML next to the PDF")
    args = parser.parse_args()

    in_path = Path(args.input)
    if not in_path.exists():
        print(f"error: {in_path} not found", file=sys.stderr)
        return 1

    out_path = Path(args.output) if args.output else in_path.with_suffix(".pdf")

    pdf_bytes = docx_to_pdf_bytes(in_path)
    out_path.write_bytes(pdf_bytes)
    print(f"wrote {out_path} ({len(pdf_bytes):,} bytes)")

    if args.html:
        html_path = out_path.with_suffix(".html")
        html_path.write_text(docx_to_html(in_path), encoding="utf-8")
        print(f"wrote {html_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
