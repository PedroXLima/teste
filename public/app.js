const form = document.getElementById("upload-form");
const input = document.getElementById("docx-input");
const dropzone = document.getElementById("dropzone");
const selectedFile = document.getElementById("selected-file");
const statusBox = document.getElementById("status-box");
const submitButton = document.getElementById("submit-button");

function setStatus(message, tone = "") {
  statusBox.textContent = message;
  statusBox.classList.remove("is-error", "is-success");
  if (tone) {
    statusBox.classList.add(tone);
  }
}

function updateSelectedFile(file) {
  selectedFile.textContent = file ? file.name : "Nenhum arquivo selecionado";
}

function ensureDocx(file) {
  return file && file.name.toLowerCase().endsWith(".docx");
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

input.addEventListener("change", () => {
  const [file] = input.files;
  updateSelectedFile(file);

  if (file && !ensureDocx(file)) {
    setStatus("Selecione um arquivo com extensao .docx.", "is-error");
    input.value = "";
    updateSelectedFile(null);
    return;
  }

  setStatus(file ? "Arquivo pronto. Clique em 'Gerar PDF institucional'." : "Selecione um documento para iniciar a geracao.");
});

["dragenter", "dragover"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.add("is-dragover");
  });
});

["dragleave", "dragend", "drop"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.remove("is-dragover");
  });
});

dropzone.addEventListener("drop", (event) => {
  const [file] = event.dataTransfer.files;
  if (!file) {
    return;
  }

  if (!ensureDocx(file)) {
    setStatus("Apenas arquivos .docx sao aceitos.", "is-error");
    return;
  }

  const dataTransfer = new DataTransfer();
  dataTransfer.items.add(file);
  input.files = dataTransfer.files;
  updateSelectedFile(file);
  setStatus("Arquivo carregado. Inicie a geracao quando quiser.");
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const [file] = input.files;
  if (!file) {
    setStatus("Selecione um arquivo DOCX antes de continuar.", "is-error");
    return;
  }

  if (!ensureDocx(file)) {
    setStatus("O arquivo enviado deve terminar com .docx.", "is-error");
    return;
  }

  submitButton.disabled = true;
  setStatus("Gerando o PDF institucional. Isso pode levar alguns segundos, especialmente em documentos longos.");

  try {
    const formData = new FormData();
    formData.append("docx", file);

    const response = await fetch("/api/generate", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      const payload = await response.json().catch(() => ({}));
      throw new Error(payload.error || "Falha ao gerar o PDF.");
    }

    const outputFilenameHeader = response.headers.get("content-disposition") || "";
    const filenameMatch = outputFilenameHeader.match(/filename="(.+?)"/i);
    const outputFilename = filenameMatch ? filenameMatch[1] : `${file.name.replace(/\.docx$/i, "")}.pdf`;

    const warningHeader = response.headers.get("x-generation-warnings");
    const pdfBlob = await response.blob();

    downloadBlob(pdfBlob, outputFilename);
    setStatus(
      warningHeader
        ? `PDF gerado com avisos de conversao: ${warningHeader}`
        : "PDF gerado com sucesso. O download foi iniciado automaticamente.",
      "is-success",
    );
  } catch (error) {
    setStatus(error.message || "Nao foi possivel concluir a geracao do PDF.", "is-error");
  } finally {
    submitButton.disabled = false;
  }
});
