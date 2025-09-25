const { jsPDF } = window.jspdf;

// helpers
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
document.getElementById("wordUpload").addEventListener("change", (e) => {
  wordFile = e.target.files[0];
  if (wordFile) document.getElementById("convertWord").disabled = false;
});
document.getElementById("wordDrop").addEventListener("click", () =>
  document.getElementById("wordUpload").click()
);

document.getElementById("convertWord").addEventListener("click", async () => {
  if (!wordFile) return;
  updateProgress("wordProgress", 10);

  // render docx → HTML
  const container = document.createElement("div");
  container.style.width = "800px"; // approx A4 width
  await new window.docxPreview.DocxPreview(wordFile, container);

  updateProgress("wordProgress", 40);

  // screenshot
  const canvas = await html2canvas(container, { scale: 2 });
  updateProgress("wordProgress", 70);

  // save as PDF
  const pdf = new jsPDF("p", "pt", "a4");
  const imgData = canvas.toDataURL("image/png");
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = (canvas.height * pageWidth) / canvas.width;
  pdf.addImage(imgData, "PNG", 0, 0, pageWidth, pageHeight);

  updateProgress("wordProgress", 100);

  pdf.save("converted.pdf");
});

// === PDF ➝ Word ===
let pdfFile = null;
document.getElementById("pdfUpload").addEventListener("change", (e) => {
  pdfFile = e.target.files[0];
  if (pdfFile) document.getElementById("convertPdf").disabled = false;
});
document.getElementById("pdfDrop").addEventListener("click", () =>
  document.getElementById("pdfUpload").click()
);

document.getElementById("convertPdf").addEventListener("click", async () => {
  if (!pdfFile) return;
  updateProgress("pdfProgress", 10);

  const arrayBuffer = await pdfFile.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  updateProgress("pdfProgress", 40);

  let textContent = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const txt = await page.getTextContent();
    textContent +=
      txt.items.map((item) => item.str).join(" ") + "\n\n";
    updateProgress("pdfProgress", Math.min(90, (i / pdf.numPages) * 80));
  }

  updateProgress("pdfProgress", 95);

  // create docx via Mammoth (simple text only)
  const blob = new Blob([textContent], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });

  updateProgress("pdfProgress", 100);
  triggerDownload(blob, "converted.docx");
});
