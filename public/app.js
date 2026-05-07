const form = document.querySelector("#upload-form");
const input = document.querySelector("#docx-input");
const dropZone = document.querySelector("#drop-zone");
const fileName = document.querySelector("#file-name");
const button = document.querySelector("#submit-button");
const statusBox = document.querySelector("#status");

input.addEventListener("change", () => {
  updateSelectedFile(input.files?.[0]);
});

for (const eventName of ["dragenter", "dragover"]) {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add("is-dragging");
  });
}

for (const eventName of ["dragleave", "drop"]) {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove("is-dragging");
  });
}

dropZone.addEventListener("drop", (event) => {
  const file = event.dataTransfer.files?.[0];

  if (!file) {
    return;
  }

  const transfer = new DataTransfer();
  transfer.items.add(file);
  input.files = transfer.files;
  updateSelectedFile(file);
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const file = input.files?.[0];

  if (!file) {
    setStatus("Selecione um arquivo DOCX antes de gerar o PDF.", "error");
    return;
  }

  if (!file.name.toLowerCase().endsWith(".docx")) {
    setStatus("O arquivo precisa estar no formato .docx.", "error");
    return;
  }

  const formData = new FormData();
  formData.append("docx", file);
  setLoading(true);
  setStatus("Convertendo DOCX e aplicando o design institucional...", "");

  try {
    const response = await fetch("/api/convert", {
      method: "POST",
      body: formData
    });

    if (!response.ok) {
      const payload = await response.json().catch(() => ({}));
      throw new Error(payload.error || "Nao foi possivel gerar o PDF.");
    }

    const blob = await response.blob();
    const downloadName = buildDownloadName(file.name);
    downloadBlob(blob, downloadName);
    setStatus(`PDF gerado com sucesso: ${downloadName}`, "success");
  } catch (error) {
    setStatus(error.message || "Erro inesperado ao gerar o PDF.", "error");
  } finally {
    setLoading(false);
  }
});

function updateSelectedFile(file) {
  fileName.textContent = file ? file.name : "Nenhum arquivo selecionado";
  statusBox.textContent = "";
  statusBox.className = "status";
}

function setLoading(isLoading) {
  button.disabled = isLoading;
  button.querySelector("span").textContent = isLoading ? "Gerando..." : "Gerar PDF";
}

function setStatus(message, type) {
  statusBox.textContent = message;
  statusBox.className = `status${type ? ` is-${type}` : ""}`;
}

function buildDownloadName(sourceName) {
  return `${sourceName.replace(/\.[^.]+$/, "").replace(/[^a-z0-9._-]+/gi, "-") || "relatorio"}.pdf`;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}
