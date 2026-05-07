(() => {
    "use strict";

    const $ = (sel) => document.querySelector(sel);
    const dropZone = $("#drop-zone");
    const fileInput = $("#file-input");
    const browseBtn = $("#browse-btn");
    const fileInfo = $("#file-info");
    const fileName = $("#file-name");
    const fileSize = $("#file-size");
    const removeBtn = $("#remove-btn");
    const metaSection = $("#metadata-section");
    const actionBar = $("#action-bar");
    const convertBtn = $("#convert-btn");
    const uploadSection = $("#upload-section");
    const progressSection = $("#progress-section");
    const progressBar = $("#progress-bar");
    const progressTitle = $("#progress-title");
    const progressDetail = $("#progress-detail");
    const resultSection = $("#result-section");
    const downloadBtn = $("#download-btn");
    const newUploadBtn = $("#new-upload-btn");
    const errorSection = $("#error-section");
    const errorMessage = $("#error-message");
    const retryBtn = $("#retry-btn");

    let selectedFile = null;

    function formatSize(bytes) {
        if (bytes < 1024) return bytes + " B";
        if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
        return (bytes / 1048576).toFixed(1) + " MB";
    }

    function showSection(section) {
        [uploadSection, progressSection, resultSection, errorSection].forEach(
            (s) => (s.style.display = "none")
        );
        section.style.display = "";
    }

    function setFile(file) {
        if (!file) return;
        if (
            !file.name.toLowerCase().endsWith(".docx") &&
            file.type !==
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ) {
            alert("Please select a valid .docx file.");
            return;
        }
        selectedFile = file;
        fileName.textContent = file.name;
        fileSize.textContent = formatSize(file.size);
        dropZone.style.display = "none";
        fileInfo.style.display = "";
        metaSection.style.display = "";
        actionBar.style.display = "";
    }

    function resetUI() {
        selectedFile = null;
        fileInput.value = "";
        dropZone.style.display = "";
        fileInfo.style.display = "none";
        metaSection.style.display = "none";
        actionBar.style.display = "none";
        showSection(uploadSection);
        progressBar.style.width = "0%";
        ["meta-title", "meta-subtitle", "meta-authors", "meta-institution", "meta-version", "meta-date"]
            .forEach((id) => ($("#" + id).value = ""));
    }

    // Drag & drop
    ["dragenter", "dragover"].forEach((evt) =>
        dropZone.addEventListener(evt, (e) => {
            e.preventDefault();
            dropZone.classList.add("drag-over");
        })
    );
    ["dragleave", "drop"].forEach((evt) =>
        dropZone.addEventListener(evt, (e) => {
            e.preventDefault();
            dropZone.classList.remove("drag-over");
        })
    );
    dropZone.addEventListener("drop", (e) => {
        const files = e.dataTransfer.files;
        if (files.length) setFile(files[0]);
    });
    dropZone.addEventListener("click", () => fileInput.click());
    browseBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        fileInput.click();
    });
    fileInput.addEventListener("change", () => {
        if (fileInput.files.length) setFile(fileInput.files[0]);
    });
    removeBtn.addEventListener("click", resetUI);
    newUploadBtn.addEventListener("click", resetUI);
    retryBtn.addEventListener("click", resetUI);

    // Convert
    convertBtn.addEventListener("click", async () => {
        if (!selectedFile) return;
        convertBtn.disabled = true;

        showSection(progressSection);
        progressBar.style.width = "10%";
        progressTitle.textContent = "Uploading document...";
        progressDetail.textContent = "Sending your DOCX file to the server.";

        const formData = new FormData();
        formData.append("file", selectedFile);

        const metaFields = {
            title: $("#meta-title").value.trim(),
            subtitle: $("#meta-subtitle").value.trim(),
            authors: $("#meta-authors").value.trim(),
            institution: $("#meta-institution").value.trim(),
            version: $("#meta-version").value.trim(),
            date: $("#meta-date").value.trim(),
        };
        formData.append("metadata", JSON.stringify(metaFields));

        try {
            progressBar.style.width = "30%";
            progressTitle.textContent = "Processing document...";
            progressDetail.textContent =
                "Parsing DOCX structure and generating PDF layout.";

            const resp = await fetch("/api/convert", {
                method: "POST",
                body: formData,
            });

            progressBar.style.width = "80%";

            if (!resp.ok) {
                const err = await resp.json().catch(() => ({}));
                throw new Error(err.error || `Server error (${resp.status})`);
            }

            progressBar.style.width = "100%";
            progressTitle.textContent = "Finalizing...";

            const blob = await resp.blob();
            const url = URL.createObjectURL(blob);

            const disposition = resp.headers.get("Content-Disposition") || "";
            let outName = "output.pdf";
            const match = disposition.match(/filename="?(.+?)"?$/);
            if (match) outName = match[1];

            downloadBtn.href = url;
            downloadBtn.download = outName;

            setTimeout(() => {
                showSection(resultSection);
                convertBtn.disabled = false;
            }, 400);
        } catch (err) {
            errorMessage.textContent = err.message;
            showSection(errorSection);
            convertBtn.disabled = false;
        }
    });
})();
