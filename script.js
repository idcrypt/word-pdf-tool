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

// === DOCX ➝ PDF ===
const docxPreview = document.getElementById("docxPreview");
const statusBoxDocx = document.getElementById("statusBoxDocx");
const progressDocxWrapper = statusBoxDocx.nextElementSibling;
const progressDocxBar = document.getElementById("progressDocx");

document.getElementById("docxUpload").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  showStatus(statusBoxDocx, "📂 Word file uploaded successfully.", "success");
  setProgress(progressDocxWrapper, progressDocxBar, 20);

  const arrayBuffer = await file.arrayBuffer();
  const docx = new window.DOCXJS.DocxPreview(docxPreview);
  docx.render(arrayBuffer);

  setProgress(progressDocxWrapper, progressDocxBar, 40);
});

document.getElementById("docxToPdf").addEventListener("click", () => {
  if (!docxPreview.innerHTML.trim()) {
    showStatus(statusBoxDocx, "❌ Please upload a Word file first.", "error");
    return;
  }

  showStatus(statusBoxDocx, "⚡ Converting to PDF... please wait.", "loading");
  setProgress(progressDocxWrapper, progressDocxBar, 60);

  html2pdf().from(docxPreview).set({
    margin: 10,
    filename: "converted.pdf",
    html2canvas: { scale: 2, useCORS: true },
    jsPDF: { unit: "mm", format: "a4", orientation: "portrait" }
  }).save()
    .then(() => {
      setProgress(progressDocxWrapper, progressDocxBar, 100);
      showStatus(statusBoxDocx, "✅ Conversion complete! PDF is ready for download.", "success");
      setTimeout(() => {
        progressDocxWrapper.style.display = "none";
        progressDocxBar.style.width = "0%";
      }, 2000);
    })
    .catch(err => {
      showStatus(statusBoxDocx, "❌ Error: " + err.message, "error");
      setProgress(progressDocxWrapper, progressDocxBar, 0);
    });
});

// === PDF ➝ DOCX ===
const statusBoxPdf = document.getElementById("statusBoxPdf");
const progressPdfWrapper = statusBoxPdf.nextElementSibling;
const progressPdfBar = document.getElementById("progressPdf");
const pdfResult = document.getElementById("pdfResult");

document.getElementById("pdfToDocx").addEventListener("click", async () => {
  const fileInput = document.getElementById("pdfUpload");
  pdfResult.innerHTML = "";

  if (!fileInput.files.length) {
    showStatus(statusBoxPdf, "❌ Please upload a PDF file first.", "error");
    return;
  }

  const file = fileInput.files[0];
  showStatus(statusBoxPdf, "📂 PDF file uploaded successfully.", "success");
  setProgress(progressPdfWrapper, progressPdfBar, 20);

  const arrayBuffer = await file.arrayBuffer();

  try {
    showStatus(statusBoxPdf, "⚡ Extracting text from PDF...", "loading");
    setProgress(progressPdfWrapper, progressPdfBar, 40);

    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let fullText = "";

    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const strings = content.items.map(item => item.str);
      fullText += strings.join(" ") + "\n\n";
      setProgress(progressPdfWrapper, progressPdfBar, 40 + Math.floor((i/pdf.numPages) * 40));
    }

    const doc = new docx.Document({
      sections: [{
        properties: {},
        children: fullText.split("\n").map(line =>
          new docx.Paragraph({ children: [new docx.TextRun(line)] })
        )
      }]
    });

    showStatus(statusBoxPdf, "⚡ Generating Word file...", "loading");
    setProgress(progressPdfWrapper, progressPdfBar, 90);

    const blob = await docx.Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);

    pdfResult.innerHTML = `<p>✅ Converted to Word!</p>
                           <a href="${url}" download="${file.name.replace(".pdf", "")}.docx">Download Word</a>`;

    setProgress(progressPdfWrapper, progressPdfBar, 100);
    showStatus(statusBoxPdf, "✅ Conversion complete! DOCX is ready for download.", "success");

    setTimeout(() => {
      progressPdfWrapper.style.display = "none";
      progressPdfBar.style.width = "0%";
    }, 2000);

  } catch (err) {
    showStatus(statusBoxPdf, "❌ Error: " + err.message, "error");
    setProgress(progressPdfWrapper, progressPdfBar, 0);
  }
});
