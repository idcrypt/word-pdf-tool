// Setup drag-drop area
function setupDropArea(label, input) {
  label.addEventListener("click", () => input.click());
  label.addEventListener("dragover", e => { e.preventDefault(); label.style.background="rgba(0,224,255,0.2)"; });
  label.addEventListener("dragleave", () => label.style.background="transparent");
  label.addEventListener("drop", e => {
    e.preventDefault();
    label.style.background="transparent";
    input.files = e.dataTransfer.files;
  });
}
setupDropArea(document.querySelector('label[for="wordUpload"]'), document.getElementById("wordUpload"));
setupDropArea(document.querySelector('label[for="pdfUpload"]'), document.getElementById("pdfUpload"));

// Word -> PDF
document.getElementById("convertWord").addEventListener("click", async () => {
  const file = document.getElementById("wordUpload").files[0];
  if (!file) return alert("Upload a Word file first!");
  const prog = document.getElementById("wordProgress");
  prog.hidden = false; prog.value = 10;

  const reader = new FileReader();
  reader.onload = async () => {
    prog.value = 40;
    // Render DOCX into preview
    const preview = document.getElementById("preview");
    preview.innerHTML = "";
    const docx = new window.DocxPreview();
    await docx.renderAsync(reader.result, preview);
    prog.value = 60;

    // Screenshot preview
    const canvas = await html2canvas(preview, {scale:2});
    const imgData = canvas.toDataURL("image/png");
    prog.value = 80;

    // Save to PDF
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF("p", "pt", "a4");
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const imgProps = pdf.getImageProperties(imgData);
    const imgHeight = (imgProps.height * pageWidth) / imgProps.width;
    pdf.addImage(imgData, "PNG", 0, 0, pageWidth, imgHeight);
    prog.value = 100;
    pdf.save(file.name.replace(/\.docx$/i,"") + ".pdf");
  };
  reader.readAsArrayBuffer(file);
});

// PDF -> Word
document.getElementById("convertPdf").addEventListener("click", async () => {
  const file = document.getElementById("pdfUpload").files[0];
  if (!file) return alert("Upload a PDF file first!");
  const prog = document.getElementById("pdfProgress");
  prog.hidden = false; prog.value = 10;

  const reader = new FileReader();
  reader.onload = async () => {
    const typedarray = new Uint8Array(reader.result);
    const pdf = await pdfjsLib.getDocument(typedarray).promise;
    let fullText = "";

    for (let i=1; i<=pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      fullText += content.items.map(it => it.str).join(" ") + "\n\n";
      prog.value = 10 + (i/pdf.numPages * 80);
    }

    prog.value = 95;
    // Save as .doc (simple text-based)
    const blob = new Blob([fullText], {type: "application/msword"});
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = file.name.replace(/\.pdf$/i,"") + ".doc";
    link.click();
    prog.value = 100;
  };
  reader.readAsArrayBuffer(file);
});
