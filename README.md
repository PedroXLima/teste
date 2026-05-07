# DOCX para PDF institucional

Aplicacao web para receber arquivos `.docx` e gerar PDFs A4 com identidade
corporativa, institucional e analitica inspirada em relatorios tecnicos do
Observatorio da Industria / Sistema FIEA.

## Recursos

- Botao de upload e suporte a arrastar/soltar arquivos DOCX.
- Conversao de DOCX para HTML preservando texto, titulos, tabelas, imagens,
  legendas, fontes, referencias e anexos sempre que o arquivo original permitir.
- Geracao de PDF com:
  - capa institucional azul;
  - pagina de creditos quando houver metadados no DOCX;
  - sumario automatico por titulos;
  - divisorias de secao;
  - cabecalho, rodape e numeracao;
  - tabelas com cabecalho azul, linhas alternadas e bordas discretas;
  - figuras com cartoes visuais, legendas e fontes;
  - caixas de destaque para notas como `Insight:`, `Recomendacao:`,
    `Risco:` ou `Oportunidade:`;
  - contracapa institucional.

## Como executar

```bash
npm install
npm run install:browsers
npm start
```

Acesse `http://localhost:3000` e envie um arquivo `.docx`.

## Desenvolvimento

```bash
npm run dev
npm test
```

## Observacoes

- O limite de upload e 30 MB por arquivo.
- A conversao usa Mammoth para extrair o conteudo do DOCX e Playwright/Chromium
  para renderizar o HTML editorial em PDF print-ready.
- O app nao reescreve ou resume o texto original; ele aplica estrutura visual,
  hierarquia editorial e identidade grafica ao conteudo fornecido.
