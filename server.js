const path = require("path");

const express = require("express");
const multer = require("multer");

const { convertDocxToDesignedPdf } = require("./src/docx-to-pdf");

const app = express();
const port = process.env.PORT || 3000;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 25 * 1024 * 1024,
  },
});

app.use(express.static(path.join(__dirname, "public")));

app.get("/health", (_req, res) => {
  res.json({ ok: true });
});

app.post("/api/generate", upload.single("docx"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "Selecione um arquivo DOCX para continuar." });
    }

    const filename = req.file.originalname || "documento.docx";

    if (!filename.toLowerCase().endsWith(".docx")) {
      return res.status(400).json({ error: "O arquivo enviado deve estar no formato .docx." });
    }

    const { pdfBuffer, warnings, outputFilename } = await convertDocxToDesignedPdf({
      buffer: req.file.buffer,
      originalName: filename,
    });

    if (warnings.length > 0) {
      res.setHeader("X-Generation-Warnings", warnings.join(" | ").slice(0, 1024));
    }

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${outputFilename}"`);
    res.send(Buffer.from(pdfBuffer));
  } catch (error) {
    console.error("PDF generation failed:", error);

    res.status(500).json({
      error:
        "Nao foi possivel gerar o PDF. Verifique se o navegador do Playwright foi instalado com 'npm run install:browsers' e tente novamente.",
      details: error.message,
    });
  }
});

app.listen(port, () => {
  console.log(`DOCX PDF Designer running at http://localhost:${port}`);
});
