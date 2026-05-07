# DOCX → PDF Institucional

Aplicativo web que transforma arquivos **DOCX** em relatórios PDF profissionais no estilo do **Observatório da Indústria / Sistema FIEA**.

## Funcionalidades

- Upload de arquivo `.docx` via interface drag-and-drop ou clique
- Geração automática de PDF com design sistema institucional:
  - Capa com gradiente azul-marinho e padrão de rede geométrica
  - Cabeçalhos e rodapés automáticos com número de página
  - Tipografia hierárquica (Montserrat + Inter)
  - Tabelas estilizadas (cabeçalho azul-escuro, linhas alternadas)
  - Callouts, citações e listas formatadas
  - Contracapa institucional
- Preservação total do conteúdo original (textos, tabelas, imagens, referências)
- Download direto do PDF gerado

## Pré-requisitos

- Python 3.10+
- Sistema operacional Linux (Ubuntu recomendado)

## Instalação

```bash
pip install -r requirements.txt
```

## Uso

```bash
python3 app.py
```

Acesse [http://localhost:5000](http://localhost:5000) no navegador.

## Tecnologias

| Componente        | Biblioteca         |
|-------------------|--------------------|
| Backend           | Flask              |
| Parsing DOCX      | python-docx        |
| Geração de PDF    | WeasyPrint         |
| Imagens           | Pillow             |
| Frontend          | HTML/CSS/JS puro   |

## Paleta de Cores

| Token           | Hex       |
|-----------------|-----------|
| Primary Blue    | `#005A9C` |
| Dark Navy       | `#003B6F` |
| Medium Blue     | `#0B63B6` |
| Light Blue      | `#D9EAF7` |
| Very Light Blue | `#EEF6FC` |

## Estrutura do Projeto

```
app.py                  # Aplicação Flask
requirements.txt        # Dependências Python
utils/
  docx_parser.py        # Parsing e extração de conteúdo DOCX
  html_builder.py       # Construção do HTML com design system
  pdf_generator.py      # Geração do PDF via WeasyPrint
templates/
  index.html            # Interface web
```
