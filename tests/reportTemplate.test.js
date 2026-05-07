import { describe, expect, it } from "vitest";
import { analyzeDocument } from "../src/docxReport.js";
import { buildReportHtml, decorateBodyHtml } from "../src/reportTemplate.js";

describe("report generation", () => {
  it("infers report metadata and preserves structured content", () => {
    const report = analyzeDocument(`
      <h1>Analise Industrial 2026</h1>
      <p class="subtitle">Panorama estrategico para cadeias produtivas</p>
      <p>Versao: Maio 2026</p>
      <h2>Introducao</h2>
      <p>Texto original preservado.</p>
      <p>Tabela 1 - Indicadores</p>
      <table><tr><td>Indicador</td><td>Valor</td></tr><tr><td>PIB</td><td>2%</td></tr></table>
      <p>Fonte: Elaboracao propria.</p>
    `);

    expect(report.title).toBe("Analise Industrial 2026");
    expect(report.subtitle).toBe("Panorama estrategico para cadeias produtivas");
    expect(report.metadata).toContainEqual({ label: "Versao", value: "Maio 2026" });
    expect(report.bodyHtml).toContain("institutional-table");
    expect(report.headings.map((heading) => heading.text)).toEqual([
      "Analise Industrial 2026",
      "Introducao"
    ]);
  });

  it("decorates major sections, tables, images, sources, and callouts", () => {
    const html = decorateBodyHtml(`
      <h1 id="metodologia">Metodologia</h1>
      <p>Insight: A cadeia produtiva exige monitoramento continuo.</p>
      <p><img src="data:image/png;base64,abc" /></p>
      <p>Figura 1 - Mapa industrial</p>
      <table class="institutional-table"><tr><th>A</th></tr><tr><td>B</td></tr></table>
    `, "Outro titulo");

    expect(html).toContain("section-divider");
    expect(html).toContain("callout");
    expect(html).toContain("image-card");
    expect(html).toContain("table-shell");
  });

  it("builds a complete institutional PDF HTML document", () => {
    const report = analyzeDocument(`
      <h1>Relatorio Setorial</h1>
      <p>Autor: Observatorio da Industria</p>
      <h2>Resultados</h2>
      <p>Conteudo analitico.</p>
    `);

    const html = buildReportHtml(report);

    expect(html).toContain("cover__panel");
    expect(html).toContain("Creditos institucionais");
    expect(html).toContain("Sumario");
    expect(html).toContain("back-cover");
    expect(html).toContain("Conteudo analitico.");
  });
});
