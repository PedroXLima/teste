import path from "node:path";
import { fileURLToPath } from "node:url";
import express from "express";
import multer from "multer";
import { convertDocxBuffer } from "./docxReport.js";
import { buildReportHtml } from "./reportTemplate.js";
import { renderPdf } from "./pdfRenderer.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.resolve(__dirname, "../public");
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 30 * 1024 * 1024,
    files: 1
  },
  fileFilter: (_request, file, callback) => {
    const extension = path.extname(file.originalname).toLowerCase();
    const allowedMimeTypes = new Set([
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "application/octet-stream"
    ]);

    if (extension === ".docx" && allowedMimeTypes.has(file.mimetype)) {
      callback(null, true);
      return;
    }

    callback(new Error("Envie um arquivo DOCX valido."));
  }
});

export function createApp() {
  const app = express();

  app.disable("x-powered-by");
  app.use(express.static(publicDir, { extensions: ["html"] }));

  app.get("/health", (_request, response) => {
    response.json({ ok: true });
  });

  app.post("/api/convert", upload.single("docx"), async (request, response, next) => {
    try {
      if (!request.file) {
        response.status(400).json({ error: "Selecione um arquivo DOCX para converter." });
        return;
      }

      const report = await convertDocxBuffer(request.file.buffer, request.file.originalname);
      const html = buildReportHtml(report);
      const pdf = await renderPdf(html, report);
      const filename = `${safeFilename(report.filenameBase || "relatorio")}.pdf`;

      response.setHeader("Content-Type", "application/pdf");
      response.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
      response.setHeader("Cache-Control", "no-store");
      response.send(pdf);
    } catch (error) {
      next(error);
    }
  });

  app.use((error, _request, response, _next) => {
    const status = error instanceof multer.MulterError ? 400 : 500;
    const message = error instanceof multer.MulterError
      ? formatMulterError(error)
      : error.message || "Nao foi possivel gerar o PDF.";

    response.status(status).json({ error: message });
  });

  return app;
}

function formatMulterError(error) {
  if (error.code === "LIMIT_FILE_SIZE") {
    return "O arquivo excede o limite de 30 MB.";
  }

  return "Nao foi possivel processar o upload.";
}

function safeFilename(value) {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9._-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 120) || "relatorio";
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const port = Number(process.env.PORT || 3000);

  createApp().listen(port, () => {
    console.log(`DOCX PDF app listening on http://localhost:${port}`);
  });
}
