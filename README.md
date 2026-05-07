# DOCX para PDF Institucional

Aplicação web para converter arquivos `.docx` em PDFs institucionais com identidade editorial corporativa, voltada para relatórios técnicos, econômicos e de inteligência.

## Funcionalidades

- Upload de arquivo DOCX pela interface web.
- Geração automática de PDF em formato A4.
- Sistema visual institucional (azuis, tipografia sóbria, cabeçalho, hierarquia e linhas divisórias).
- Estrutura de publicação com:
  - Capa
  - Página de créditos (quando identificada no DOCX)
  - Página catalográfica/contatos (quando identificada no DOCX)
  - Sumário automático
  - Páginas de conteúdo com cabeçalho e paginação
  - Estilo institucional de tabelas
  - Captions e fontes
  - Caixas de destaque (callouts)
  - Divisórias de seção para blocos principais
  - Contracapa

## Requisitos

- Python 3.10+

## Instalação

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Execução

```bash
python app.py
```

A aplicação ficará disponível em:

`http://localhost:8000`

## Uso

1. Abra a interface web.
2. Clique em **Selecionar arquivo DOCX**.
3. Envie o documento.
4. Clique em **Gerar PDF**.
5. O download do arquivo PDF será iniciado automaticamente.
