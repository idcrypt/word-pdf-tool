// === DOCX ➝ PDF ===
document.getElementById("docxToPdf").addEventListener("click", async () => {
  const fileInput = document.getElementById("docxUpload");
  if (!fileInput.files.length) {
    alert("Please upload a DOCX file first.");
    return;
  }

  const file = fileInput.files[0];
  const arrayBuffer = await file.arrayBuffer();

  // Extract text from DOCX
  const result = await window.mammoth.extractRawText({ arrayBuffer });
  const text = result.value || "(No text found)";

  // Create PDF
  const { PDFDocument, StandardFonts, rgb } = PDFLib;
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([595, 842]); // A4 size
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const fontSize = 12;

  const lines = text.split("\n");
  let y = 800;
  lines.forEach(line => {
    if (y < 50) {
      y = 800;
      pdfDoc.addPage([595, 842]);
    }
    const currentPage = pdfDoc.getPages()[pdfDoc.getPageCount() - 1];
    currentPage.drawText(line, { x: 50, y, size: fontSize, font, color: rgb(1, 1, 1) });
    y -= 20;
  });

  const pdfBytes = await pdfDoc.save();

  // Download PDF
  const blob = new Blob([pdfBytes], { type: "application/pdf" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = file.name.replace(".docx", "") + ".pdf";
  link.click();

  document.getElementById("docxResult").innerText = "✅ Converted to PDF!";
});


// === PDF ➝ DOCX ===
document.getElementById("pdfToDocx").addEventListener("click", async () => {
  const fileInput = document.getElementById("pdfUpload");
  if (!fileInput.files.length) {
    alert("Please upload a PDF file first.");
    return;
  }

  const file = fileInput.files[0];
  const arrayBuffer = await file.arrayBuffer();

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
      children: [
        new docx.Paragraph({
          children: [new docx.TextRun(fullText)]
        })
      ]
    }]
  });

  const blob = await docx.Packer.toBlob(doc);
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = file.name.replace(".pdf", "") + ".docx";
  link.click();

  document.getElementById("pdfResult").innerText = "✅ Converted to Word!";
});
