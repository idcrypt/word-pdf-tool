const { jsPDF } = window.jspdf;

// Helpers
function updateProgress(barId, percent) {
  document.getElementById(barId).style.width = percent + "%";
}
function triggerDownload(blob, filename) {
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
}

// === Word ➝ PDF ===
let wordFile = null;
const wordUpload = document.getElementById("wordUpload");
const wordDrop = document.getElementById("wordDrop");
const wordBtn = document.getElementById("convertWord");

wordUpload.addEventListener("change", (e) => {
  wordFile = e.target.files[0];
  if (wordFile) wordBtn.disabled = false;
});
wordDrop.addEventListener("click", () => wordUpload.click());
wordDrop.addEventListener("dragover", (e) => { e.preventDefault(); wordDrop.classList.add("hover"); });
wordDrop.addEventListener("dragleave", () => wordDrop.classList.remove("hover"));
wordDrop.addEventListener("drop", (e) => {
  e.preventDefault();
  wordFile = e.dataTransfer.files[0];
  if (wordFile) { wordBtn.disabled = false; }
});

wordBtn.addEventListener("click", async () => {
  if (!wordFile) return;
  updateProgress("wordProgress", 10);

  // parse Word → HTML
  const arrayBuffer = await wordFile.arrayBuffer();
  const result = await window.mammoth.convertToHtml({ arrayBuffer });
  const container = document.createElement("div");
  container.style.width = "800px";
  container.innerHTML = result.value;

  updateProgress("wordProgress", 30);

  // screenshot full
  const canvas = await html2canvas(container, { scale: 2 });
  updateProgress("wordProgress", 60);

  // PDF multi-page
  const pdf = new jsPDF("p", "pt", "a4");
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();

  let imgHeight = (canvas.height * pageWidth) / canvas.width;
  let y = 0;

  while (y < imgHeight) {
    const sliceCanvas = document.createElement("canvas");
    sliceCanvas.width = canvas.width;
    sliceCanvas.height = (pageHeight * canvas.width) / pageWidth;

    const ctx = sliceCanvas.getContext("2d");
    ctx.drawImage(
      canvas,
      0, y * (canvas.width / pageWidth),
      canvas.width, sliceCanvas.height,
      0, 0,
      canvas.width, sliceCanvas.height
    );

    const sliceData = sliceCanvas.toDataURL("image/png");
    if (y > 0) pdf.addPage();
    pdf.addImage(sliceData, "PNG", 0, 0, pageWidth, pageHeight);

    y += pageHeight;
  }

  updateProgress("wordProgress", 100);
  pdf.save("converted.pdf");
});

// === PDF ➝ Word ===
let pdfFile = null;
const pdfUpload = document.getElementById("pdfUpload");
const pdfDrop = document.getElementById("pdfDrop");
const pdfBtn = document.getElementById("convertPdf");

pdfUpload.addEventListener("change", (e) => {
  pdfFile = e.target.files[0];
  if (pdfFile) pdfBtn.disabled = false;
});
pdfDrop.addEventListener("click", () => pdfUpload.click());
pdfDrop.addEventListener("dragover", (e) => { e.preventDefault(); pdfDrop.classList.add("hover"); });
pdfDrop.addEventListener("dragleave", () => pdfDrop.classList.remove("hover"));
pdfDrop.addEventListener("drop", (e) => {
  e.preventDefault();
  pdfFile = e.dataTransfer.files[0];
  if (pdfFile) { pdfBtn.disabled = false; }
});

pdfBtn.addEventListener("click", async () => {
  if (!pdfFile) return;
  updateProgress("pdfProgress", 10);

  const arrayBuffer = await pdfFile.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  updateProgress("pdfProgress", 30);

  let textContent = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const txt = await page.getTextContent();
    textContent += txt.items.map((item) => item.str).join(" ") + "\n\n";
    updateProgress("pdfProgress", Math.min(90, (i / pdf.numPages) * 80));
  }

  updateProgress("pdfProgress", 100);

  // save as simple .doc
  const blob = new Blob([textContent], { type: "application/msword" });
  triggerDownload(blob, "converted.doc");
});
