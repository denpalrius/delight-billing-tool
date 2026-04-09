const form = document.querySelector("#upload-form");
const input = document.querySelector("#files");
const status = document.querySelector("#status");
const selectedFiles = document.querySelector("#selected-files");
const submitButton = document.querySelector("#submit-button");
const progressPanel = document.querySelector("#progress-panel");
const progressBar = document.querySelector("#progress-bar");
const progressValue = document.querySelector("#progress-value");

let pendingDownload = null;
let progressTimer = null;
let currentProgress = 0;
let selectedUploads = [];

input.addEventListener("change", () => {
  clearDownloadState();
  if (!input.files || input.files.length === 0) {
    syncSelectedFiles();
    return;
  }

  const incomingFiles = Array.from(input.files);
  for (const file of incomingFiles) {
    const fileKey = createFileKey(file);
    if (!selectedUploads.some((entry) => createFileKey(entry) === fileKey)) {
      selectedUploads.push(file);
    }
  }

  syncSelectedFiles();
  input.value = "";
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (selectedUploads.length === 0) {
    showProgress();
    setStatus("Select at least one Excel file to continue.", "error");
    return;
  }

  setBusy(true);
  clearDownloadState();
  showProgress();
  startProgress();
  setStatus("Checking files and preparing secure upload.", "pending");

  try {
    const formData = new FormData();
    for (const file of selectedUploads) {
      formData.append("files", file, file.name);
    }

    const response = await fetch("/api/process", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      const payload = await readError(response);
      throw new Error(payload);
    }

    const blob = await response.blob();
    const filename = readFilename(response.headers) ?? "daily_summary.xlsx";
    pendingDownload = {
      url: URL.createObjectURL(blob),
      filename,
    };
    stopProgress(100);
    triggerDownload();
    setStatus(
      "Processing complete. The workbook was downloaded automatically.",
      "success",
    );
  } catch (error) {
    clearInterval(progressTimer);
    setStatus(error instanceof Error ? error.message : "Processing failed.", "error");
  } finally {
    setBusy(false);
  }
});

selectedFiles.addEventListener("click", (event) => {
  const removeButton = event.target.closest("[data-remove-index]");
  if (!removeButton) {
    return;
  }

  const index = Number(removeButton.dataset.removeIndex);
  if (Number.isNaN(index)) {
    return;
  }

  selectedUploads.splice(index, 1);
  clearDownloadState();
  hideProgress();
  syncSelectedFiles();
});

function setBusy(isBusy) {
  submitButton.disabled = isBusy;
  input.disabled = isBusy;
  selectedFiles.classList.toggle("is-busy", isBusy);
}

function setStatus(message, tone) {
  status.textContent = message;
  status.dataset.tone = tone;
  progressPanel.dataset.tone = tone;
}

function showProgress() {
  progressPanel.classList.remove("is-hidden");
}

function hideProgress() {
  clearInterval(progressTimer);
  progressPanel.classList.add("is-hidden");
}

function clearDownloadState() {
  if (pendingDownload) {
    URL.revokeObjectURL(pendingDownload.url);
  }
  pendingDownload = null;
}

function triggerDownload() {
  if (!pendingDownload) {
    return;
  }

  const anchor = document.createElement("a");
  anchor.href = pendingDownload.url;
  anchor.download = pendingDownload.filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
}

function startProgress() {
  clearInterval(progressTimer);
  currentProgress = 18;
  renderProgress();

  const checkpoints = [
    [32, "Uploading files to the billing processor."],
    [58, "Reading worksheet data and checking entries."],
    [82, "Building the workbook and preparing the download."],
  ];

  progressTimer = window.setInterval(() => {
    if (currentProgress >= 88) {
      return;
    }

    currentProgress += currentProgress < 58 ? 8 : 4;
    const activeCheckpoint = checkpoints.find(([limit]) => currentProgress <= limit + 6);
    if (activeCheckpoint) {
      setStatus(activeCheckpoint[1], "pending");
    }
    renderProgress();
  }, 650);
}

function stopProgress(value) {
  clearInterval(progressTimer);
  currentProgress = value;
  renderProgress();
}

function renderProgress() {
  progressBar.style.width = `${currentProgress}%`;
  progressValue.textContent = `${Math.min(currentProgress, 100)}%`;
}

async function readError(response) {
  const contentType = response.headers.get("content-type") ?? "";
  if (contentType.includes("application/json")) {
    const payload = await response.json();
    return payload.error ?? "Processing failed.";
  }

  return (await response.text()) || "Processing failed.";
}

function readFilename(headers) {
  const disposition = headers.get("content-disposition");
  if (!disposition) {
    return null;
  }

  const match = disposition.match(/filename="([^"]+)"/i);
  return match ? match[1] : null;
}

function formatFileSize(bytes) {
  if (bytes < 1024 * 1024) {
    return `${Math.round(bytes / 1024)} KB`;
  }

  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function syncSelectedFiles() {
  const transfer = new DataTransfer();
  for (const file of selectedUploads) {
    transfer.items.add(file);
  }
  input.files = transfer.files;

  if (selectedUploads.length === 0) {
    selectedFiles.textContent = "No files selected yet.";
    submitButton.disabled = true;
    return;
  }

  selectedFiles.innerHTML = selectedUploads
    .map(
      (file, index) => `
        <span class="file-chip">
          <span class="file-chip-name">${file.name}</span>
          <span>${formatFileSize(file.size)}</span>
          <button
            type="button"
            class="file-chip-remove"
            data-remove-index="${index}"
            aria-label="Remove ${file.name}"
          >
            ×
          </button>
        </span>
      `,
    )
    .join("");
  submitButton.disabled = false;
}

function createFileKey(file) {
  return `${file.name}:${file.size}:${file.lastModified}`;
}
