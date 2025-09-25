// === Helpers ===
function showStatus(box, msg, type="loading") {
  box.style.display = "block";
  box.className = type;
  box.textContent = msg;
}
function setProgress(wrapper, bar, percent) {
  wrapper.style.display = "block";
  bar.style.width = percent + "%";
}
function setupDropArea(area, input) {
  area.addEventListener("click", () => input.click());
  area.addEventListener("dragover", e => {
    e.preventDefault();
    area.classList.add("dragover");
  });
  area.addEventListener("dragleave", () => area.classList.remove("dragover"));
  area.addEventListener("drop", e => {
    e.preventDefault();
    area.classList.remove("dragover");
    if (e.dataTransfer.files.length) {
      input.files = e.dataTransfer.files;
      input.dispatchEvent(new Event("change"));
    }
  });
}

// === DOCX ➝ PDF ===
const docxPreview = document.getElementById("docxPreview");
const statusBoxDocx = document.getElementById("statusBoxDocx");
const progressDocxBar = document.getElementById("progressDocx");

setupDropArea(document.getElementById("dropDocx"), document.getElementById("docxUpload"));

document.getElementById("docxUpload").addEventListener("change", async e => {
  const file = e.target.files[0];
  if (!file) return;
  showStatus(statusBoxDocx, "📂 Word file uploaded successfully.", "success");
  setProgress(statusBoxDocx.nextElementSibling, progressDocxBar, 20);

  const buffer = await file.arrayBuffer();
  docxPreview.innerHTML = "";
  window.docx.renderAsync(buffer, docxPreview);
  setProgress(statusBoxDocx.nextElementSibling, progressDocxBar, 40);
});

document.getElementById("docxToPdf").addEventListener("click", () => {
  if (!docxPreview.innerHTML.trim()) {
    showStatus(statusBoxDocx, "❌ Please upload a Word file first.", "error");
    return;
  }
  showStatus(statusBoxDocx, "⚡ Converting to PDF...", "loading");
  setProgress(statusBoxDocx.nextElementSibling, progressDocxBar, 60);

  html2pdf().from(docxPreview).set({
    margin: 10,
    filename: "converted.pdf",
    html2canvas: { scale: 2 },
    jsPDF: { unit: "mm", format: "a4", orientation: "portrait" }
  }).save().then(() => {
    showStatus(statusBoxDocx, "✅ Conversion complete!", "success");
    setProgress(statusBoxDocx.nextElementSibling, progressDocxBar, 100);
  });
});

// === PDF ➝ DOCX ===
const statusBoxPdf = document.getElementById("statusBoxPdf");
const progressPdfBar = document.getElementById("progressPdf");
const pdfResult = document.getElementById("pdfResult");

setupDropArea(document.getElementById("dropPdf"), document.getElementById("pdfUpload"));

document.getElementById("pdfToDocx").addEventListener("click", async () => {
  const input = document.getElementById("pdfUpload");
  if (!input.files.length) {
    showStatus(statusBoxPdf, "❌ Please upload a PDF first.", "error");
    return;
  }

  const file = input.files[0];
  showStatus(statusBoxPdf, "📂 PDF uploaded successfully.", "success");
  setProgress(statusBoxPdf.nextElementSibling, progressPdfBar, 20);

  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  let text = "";

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map(it => it.str).join(" ") + "\n\n";
    setProgress(statusBoxPdf.nextElementSibling, progressPdfBar, 20 + (i/pdf.numPages*60));
  }

  const doc = new docx.Document({
    sections: [{
      children: text.split("\n").map(line =>
        new docx.Paragraph(line)
      )
    }]
  });

  const blob = await docx.Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);

  pdfResult.innerHTML = `<a href="${url}" download="${file.name.replace(".pdf", "")}.docx">⬇ Download Word File</a>`;
  showStatus(statusBoxPdf, "✅ Conversion complete!", "success");
  setProgress(statusBoxPdf.nextElementSibling, progressPdfBar, 100);
});
