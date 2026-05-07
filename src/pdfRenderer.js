import { chromium } from "playwright";

export async function renderPdf(html, { title = "Relatorio tecnico" } = {}) {
  let browser;

  try {
    browser = await chromium.launch({
      headless: true,
      args: ["--no-sandbox", "--disable-dev-shm-usage"]
    });

    const page = await browser.newPage({
      viewport: { width: 1240, height: 1754 },
      deviceScaleFactor: 1
    });

    await page.setContent(html, {
      waitUntil: "networkidle",
      timeout: 45_000
    });

    await page.emulateMedia({ media: "print" });

    return await page.pdf({
      format: "A4",
      printBackground: true,
      preferCSSPageSize: true,
      displayHeaderFooter: true,
      margin: {
        top: "20mm",
        right: "16mm",
        bottom: "19mm",
        left: "16mm"
      },
      headerTemplate: renderHeaderTemplate(title),
      footerTemplate: renderFooterTemplate()
    });
  } catch (error) {
    if (isMissingBrowserError(error)) {
      throw new Error("Playwright Chromium is not installed. Run `npm run install:browsers` and try again.");
    }

    throw error;
  } finally {
    if (browser) {
      await browser.close();
    }
  }
}

function renderHeaderTemplate(title) {
  return `
    <div style="width:100%; padding:0 16mm; font-family:Inter,Arial,sans-serif; font-size:7.4pt; color:#6B7280;">
      <div style="display:flex; align-items:center; justify-content:space-between; gap:8mm; padding-top:6mm;">
        <span style="max-width:145mm; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${escapeHtml(title)}</span>
        <span style="color:#005A9C; font-weight:700;">Pagina <span class="pageNumber"></span></span>
      </div>
      <div style="height:1px; margin-top:2mm; background:#005A9C;"></div>
    </div>`;
}

function renderFooterTemplate() {
  return `
    <div style="width:100%; padding:0 16mm; font-family:Inter,Arial,sans-serif; font-size:7.2pt; color:#6B7280;">
      <div style="height:1px; margin-bottom:2mm; background:#D1D5DB;"></div>
      <div style="display:flex; align-items:center; justify-content:space-between;">
        <span>Documento institucional gerado automaticamente</span>
        <span><span class="pageNumber"></span>/<span class="totalPages"></span></span>
      </div>
    </div>`;
}

function isMissingBrowserError(error) {
  return /Executable doesn't exist|browserType.launch|install/i.test(String(error?.message ?? error));
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
