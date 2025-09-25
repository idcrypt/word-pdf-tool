// === DOCX ➝ PDF ===
document.getElementById("docxToPdf").addEventListener("click", async () => {
  const fileInput = document.getElementById("docxUpload");
  const resultDiv = document.getElementById("docxResult");
  resultDiv.innerHTML = "";

  if (!fileInput.files.length) {
    alert("Please upload a DOCX file first.");
    return;
  }

  const file = fileInput.files[0];
  const arrayBuffer = await file.arrayBuffer();

  try {
    const result = await window.mammoth.extractRawText({ arrayBuffer });
    const text = result.value || "(No text found)";

    // Create PDF
    const { PDFDocument, StandardFonts, rgb } = PDFLib;
    const pdfDoc = await PDFDocument.create();
    let page = pdfDoc.addPage([595, 842]);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontSize = 12;

    const lines = text.split("\n");
    let y = 800;

    lines.forEach(line => {
      if (y < 50) {
        page = pdfDoc.addPage([595, 842]);
        y = 800;
      }
      page.drawText(line, { x: 50, y, size: fontSize, font, color: rgb(1, 1, 1) });
      y -= 20;
    });

    const pdfBytes = await pdfDoc.save();
    const blob = new Blob([pdfBytes], { type: "application/pdf" });

    // Show preview
    const url = URL.createObjectURL(blob);
    resultDiv.innerHTML = `<p>✅ Converted to PDF!</p>
                           <iframe src="${url}" width="100%" height="200"></iframe>`;
    // Auto download
    const link = document.createElement("a");
    link.href = url;
    link.download = file.name.replace(".docx", "") + ".pdf";
    link.click();
  } catch (err) {
    resultDiv.innerHTML = `<p style="color:red;">❌ Error: ${err.message}</p>`;
  }
});


// === PDF ➝ DOCX ===
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

    // Create DOCX with extracted text
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
                           <iframe src="https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(url)}" 
                                   width="100%" height="200"></iframe>`;

    // Auto download
    const link = document.createElement("a");
    link.href = url;
    link.download = file.name.replace(".pdf", "") + ".docx";
    link.click();
  } catch (err) {
    resultDiv.innerHTML = `<p style="color:red;">❌ Error: ${err.message}</p>`;
  }
});
