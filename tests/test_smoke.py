"""End-to-end smoke test.

Generates a representative DOCX (cover, credits, copyright, presentation,
TOC, body sections, table, figure, callouts, references, annexes), runs
the full conversion pipeline, and checks that the resulting PDF is valid
and contains the expected sections.

Run with::

    python tests/test_smoke.py
"""

from __future__ import annotations

import io
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from app.converter import docx_to_pdf_bytes  # noqa: E402
from tests.build_sample_docx import build  # noqa: E402


def main() -> int:
    sample = ROOT / "tests" / "sample.docx"
    build(sample)

    pdf = docx_to_pdf_bytes(sample)
    assert pdf[:5] == b"%PDF-", "output is not a PDF"
    assert len(pdf) > 20_000, "PDF unexpectedly small"
    print(f"ok — pdf {len(pdf):,} bytes")

    try:
        from pypdf import PdfReader
    except ImportError:
        print("(pypdf not installed; skipping content checks)")
        return 0

    reader = PdfReader(io.BytesIO(pdf))
    n = len(reader.pages)
    assert n >= 12, f"expected at least 12 pages, got {n}"
    print(f"ok — {n} pages")

    text = "\n".join((p.extract_text() or "") for p in reader.pages)
    expected_markers = [
        "PANORAMA",         # cover title
        "Ficha Técnica",    # credits page
        "Apresentação",     # presentation page
        "Sumário",          # TOC page
        "INTRODUÇÃO",       # section divider
        "Metodologia",      # body section
        "Tabela 1",         # institutional table caption
        "Figura 1",         # figure caption
        "Referências",      # references section
        "ANEXOS",           # annexes title
    ]
    missing = [m for m in expected_markers if m not in text]
    assert not missing, f"missing markers in PDF: {missing}"
    print("ok — all expected sections present:", ", ".join(expected_markers))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
