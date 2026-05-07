# DOCX PDF Designer

Web app para enviar arquivos `.docx` e gerar um PDF institucional com layout editorial, capa,
sumario, divisorias de secao, tabelas estilizadas e numeracao de paginas.

## O que o app faz

- recebe um arquivo DOCX por upload
- converte o conteudo preservando texto, titulos, tabelas e imagens
- reorganiza o documento para um estilo institucional/analitico
- gera capa e contracapa em azul corporativo
- cria paginas internas com cabecalho, rodape e paginação
- estiliza tabelas, figuras, legendas, fontes e callouts
- entrega o resultado final como arquivo PDF

## Stack

- Node.js + Express
- Mammoth para converter DOCX em HTML
- Playwright para renderizar PDF em A4
- pdf-lib para mesclar capa, miolo e contracapa

## Como executar

```bash
npm install
npm run install:browsers
npm start
```

Depois abra:

```text
http://localhost:3000
```

## Fluxo de uso

1. Clique em **Gerar PDF institucional**
2. Escolha um arquivo `.docx`
3. Aguarde a renderizacao
4. O download do PDF comecara automaticamente

## Observacoes de qualidade

- estilos de titulo e subtitulo no DOCX melhoram a deteccao de capa e sumario
- tabelas e imagens do DOCX sao preservadas sempre que possivel
- o PDF e gerado com foco em leitura, impressao e identidade institucional
