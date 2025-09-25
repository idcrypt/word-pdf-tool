// === DOCX ➝ PDF (Full Render) ===
const previewDiv = document.getElementById("docxPreview");

document.getElementById("docxUpload").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const arrayBuffer = await file.arrayBuffer();
  // render docx ke HTML
  const docx = new window.DOCXJS.DocxPreview(previewDiv);
  docx.render(arrayBuffer);
});

document.getElementById("docxToPdf").addEventListener("click", () => {
  if (!previewDiv.innerHTML.trim()) {
    alert("Please upload a DOCX file first.");
    return;
  }

  // convert preview ke PDF
  html2pdf().from(previewDiv).set({
    margin: 10,
    filename: "converted.pdf",
    html2canvas: { scale: 2, useCORS: true },
    jsPDF: { unit: "mm", format: "a4", orientation: "portrait" }
  }).save();
});


// === PDF ➝ DOCX (Text Extract) ===
document.getElementById("pdfToDocx").addEventListener("click", async () => {
  const fileInput = document.getElementById("pdfUpload");
  const resultDiv = document.getElementById("pdfResult");
  resultDiv.innerHTML = "";

  if (!fileInput.files.length) {
    alert("Please upload a PDF file first.");
    return;
  }

  const file = fileInput.files[0];
  const arrayBuffer = await file.arrayBuffer();

  try {
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let fullText = "";

    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const strings = content.items.map(item => item.str);
      fullText += strings.join(" ") + "\n\n";
    }

    // buat DOCX dari teks
    const doc = new docx.Document({
      sections: [{
        properties: {},
        children: fullText.split("\n").map(line =>
          new docx.Paragraph({ children: [new docx.TextRun(line)] })
        )
      }]
    });

    const blob = await docx.Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);

    resultDiv.innerHTML = `<p>✅ Converted to Word!</p>
                           <a href="${url}" download="${file.name.replace(".pdf", "")}.docx">Download Word</a>`;
  } catch (err) {
    resultDiv.innerHTML = `<p style="color:red;">❌ Error: ${err.message}</p>`;
  }
});
