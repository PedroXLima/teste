# DOCX → PDF Institucional

Aplicação web em Flask para converter arquivos DOCX em PDFs com diagramação institucional, corporativa e analítica.

## Recursos

- Upload de arquivo `.docx` via interface web (com botão de seleção).
- Preservação estrutural do conteúdo (títulos, parágrafos, tabelas e imagens).
- Geração de PDF em formato A4 com:
  - Capa institucional em azul.
  - Página de créditos e metadados (quando identificados).
  - Sumário automático.
  - Estilo editorial consistente para corpo do relatório.
  - Tabelas no padrão azul institucional.
  - Divisórias de seções principais.
  - Contracapa institucional.

## Requisitos

- Python 3.10+

## Como executar

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Abra `http://localhost:5000` no navegador.
