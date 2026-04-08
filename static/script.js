const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");
const form = document.getElementById("uploadForm");
const loading = document.getElementById("loading");

// Click to open file
dropZone.addEventListener("click", () => fileInput.click());

// Drag events
dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("dragover");
});

dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    fileInput.files = e.dataTransfer.files;
});

// Show loading
form.addEventListener("submit", () => {
    loading.classList.remove("hidden");
});