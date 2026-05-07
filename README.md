# DOCX → PDF Institucional

Web app that converts a `.docx` file into a professionally designed,
print-ready PDF following the **Observatório da Indústria / Sistema FIEA**
inspired institutional design system (corporate blue palette, modern
sans-serif typography, A4 portrait, cover, section dividers, blue table
headers, callouts, references, annexes and back cover).

The pipeline preserves the original textual content (paragraphs,
headings, tables, figures, captions, lists, references, annexes) and
applies the design system on top of it.

## Architecture

```
app/
├── docx_parser.py   # python-docx → typed block model (text, tables, images)
├── structurer.py    # block list → logical sections (cover, credits, TOC, …)
├── renderer.py      # sections → semantic HTML
├── static/report.css# institutional design system (CSS Paged Media)
├── templates/
│   └── index.html   # upload UI
├── converter.py     # docx bytes → pdf bytes (via WeasyPrint)
└── web.py           # Flask app (POST /convert, POST /preview)
```

The HTML is rendered to PDF with [WeasyPrint](https://weasyprint.org/),
which fully supports CSS Paged Media (named pages, page numbers,
running headers/footers, page breaks).

## Requirements

System libraries (already present on most Linux distributions):
`libcairo2`, `libpango-1.0-0`, `libpangoft2-1.0-0`, `libgdk-pixbuf-2.0-0`.

Python ≥ 3.10. Install dependencies:

```bash
pip install -r requirements.txt
```

## Running the web app

```bash
python run.py
# → http://localhost:5000
```

Open the page, drop a `.docx` file in the upload area and click
**Gerar PDF Institucional**. The PDF is downloaded automatically.
There is also a **Pré-visualizar HTML** button that opens the
intermediate HTML in a new tab — useful to debug styling.

Environment variables:

- `PORT` (default `5000`)
- `HOST` (default `0.0.0.0`)
- `DEBUG` (`1` to enable Flask debug)

## Command-line usage

```bash
python cli.py path/to/input.docx                # writes path/to/input.pdf
python cli.py input.docx output.pdf             # custom output path
python cli.py input.docx output.pdf --html      # also write HTML next to it
```

## Design system

| Token            | Hex     | Use                                       |
|------------------|---------|-------------------------------------------|
| Primary Blue     | #005A9C | Section titles, accents, blue bars         |
| Dark Navy        | #003B6F | Main headings, table headers, cover bg    |
| Medium Blue      | #0B63B6 | Sub-accents, gradients                    |
| Light Blue       | #D9EAF7 | Alternate table rows                      |
| Very Light Blue  | #EEF6FC | Callout backgrounds, drop area            |
| Main Text        | #2B2B2B | Body text                                 |
| Secondary Text   | #6B7280 | Captions, sources, metadata               |
| Border / Divider | #D1D5DB | Table borders, separator lines            |

Page anatomy:

- **Cover** — full-bleed blue gradient with subtle network grid, large
  uppercase title, blue translucent bottom panel with metadata
  (autor, revisão, versão, data) and a date/version badge.
- **Section dividers** — full-bleed page placed before each top-level
  section with a large white uppercase title and blue accent bar.
- **Body pages** — white background, blue header strip with section
  name and a thin blue rule, page numbers in the bottom right.
- **Tables** — dark-navy header row with white bold text, alternating
  white / light-blue rows, thin gray borders.
- **Callouts** — light-blue background, dark-blue left border,
  rendered from `Quote`/`Citação` paragraph styles.
- **Back cover** — full-bleed blue gradient with title, subtitle and
  contact block in the bottom corners.

## Section detection

`structurer.py` classifies sections from heading text using
Portuguese-aware regular expressions:

| Detected kind   | Trigger headings (regex, case-insensitive)                                |
|-----------------|---------------------------------------------------------------------------|
| `credits`       | Ficha Técnica, Créditos, Equipe, Realização, Execução, Coordenação, Autores, Revisão |
| `copyright`     | Catalogação, CDU, CDD, Copyright, ISBN                                    |
| `presentation`  | Apresentação, Prefácio, Carta do Presidente, Mensagem                     |
| `toc`           | Sumário, Índice, Conteúdo                                                 |
| `list_figures`  | Lista de Figuras                                                          |
| `list_tables`   | Lista de Tabelas                                                          |
| `abbreviations` | Lista de Abreviações, Siglas, Glossário                                   |
| `references`    | Referências, Bibliografia                                                 |
| `annexes`       | Anexos, Apêndices                                                         |
| `back_cover`    | Contracapa, Contato                                                       |

Anything else becomes a body section preceded by a section divider.

## Limitations

- A generated table of contents is shown on the Sumário page; per-item
  page numbers are not yet pre-computed (PDF anchors are reserved).
- Document-level numbering (Heading 1 → "1.", Heading 2 → "1.1") is
  not auto-injected. Numbering already present in the DOCX is preserved.
- Charts embedded as Word objects (not rasterized images) are not
  re-rendered.
