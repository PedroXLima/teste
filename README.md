# DOCX para PDF institucional

Aplicacao web em Flask para receber arquivos DOCX e gerar PDFs editoriais em estilo institucional, com foco em relatorios tecnicos, inteligencia industrial e analise economica.

## O que a aplicacao faz

- recebe um arquivo `.docx` por upload;
- interpreta titulo, subtitulo, creditos, secoes, paragrafos, tabelas e imagens;
- gera um PDF A4 com capa, sumario, paginas internas com cabecalho institucional e contracapa;
- aplica tabelas em azul institucional, caixas de destaque, legendas e notas de fonte.

## Requisitos

- Python 3.12+

## Instalar dependencias

```bash
python3 -m pip install -r requirements.txt
```

## Executar a aplicacao

```bash
python3 -m app.main
```

Depois acesse `http://localhost:8000`.

## Rodar o teste de smoke

```bash
python3 -m unittest tests.test_pipeline
```

## Estrutura

- `app/main.py`: interface web e endpoint de upload
- `app/docx_parser.py`: leitura e classificacao do DOCX
- `app/pdf_generator.py`: composicao editorial do PDF final
- `app/templates/index.html`: tela principal
- `app/static/styles.css`: identidade visual da interface
- `tests/test_pipeline.py`: verificacao basica da pipeline de conversao
