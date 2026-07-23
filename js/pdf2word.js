'use strict';

(() => {

let pdfFile = null;

const input = document.getElementById('pdf2word-input');
const dropzone = document.getElementById('pdf2word-dropzone');
const actions = document.getElementById('pdf2word-actions');
const convertBtn = document.getElementById('pdf2word-btn');
const clearBtn = document.getElementById('pdf2word-clear');
const status = document.getElementById('pdf2word-status');

if (!input) return;

dropzone.querySelector('.link-btn')
.addEventListener('click', (e) => {
  e.stopPropagation();
  input.click();
});

dropzone.addEventListener('click', () => {
  input.click();
});

input.addEventListener('change', () => {

  const file = input.files[0];

  if (!file) return;

  if (
    file.type !== 'application/pdf' &&
    !file.name.toLowerCase().endsWith('.pdf')
  ) {
    status.textContent = 'Please choose PDF file.';
    return;
  }

  pdfFile = file;

  actions.classList.remove('hidden');

  status.textContent =
    `Selected: ${file.name}`;
});

dropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
});

dropzone.addEventListener('drop', (e) => {

  e.preventDefault();

  const file =
    e.dataTransfer.files[0];

  if (!file) return;

  pdfFile = file;

  actions.classList.remove('hidden');

  status.textContent =
    `Selected: ${file.name}`;
});

clearBtn.addEventListener('click', () => {

  pdfFile = null;

  input.value = '';

  actions.classList.add('hidden');

  status.textContent = '';
});

convertBtn.addEventListener('click', convertPdf);

async function convertPdf() {

try {

  if (!pdfFile) return;

  convertBtn.disabled = true;

  status.textContent =
    'Loading PDF...';

  const buffer =
    await pdfFile.arrayBuffer();

  const pdf =
    await pdfjsLib.getDocument({
      data: buffer
    }).promise;

  const children = [];

  for (
    let pageNum = 1;
    pageNum <= pdf.numPages;
    pageNum++
  ) {

    status.textContent =
      `Processing page ${pageNum}/${pdf.numPages}`;

    const page =
      await pdf.getPage(pageNum);

    const viewport =
      page.getViewport({
        scale: 2
      });

    const canvas =
      document.createElement('canvas');

    const ctx =
      canvas.getContext('2d');

    canvas.width =
      viewport.width;

    canvas.height =
      viewport.height;

    await page.render({
      canvasContext: ctx,
      viewport
    }).promise;

    const imageData =
      canvas.toDataURL(
        'image/jpeg',
        0.95
      );

    const imageBytes =
      dataURLToUint8Array(
        imageData
      );

    const image =
      new docx.ImageRun({
        data: imageBytes,
        transformation: {
          width: 520,
          height:
            Math.round(
              520 *
              canvas.height /
              canvas.width
            )
        }
      });

    children.push(
      new docx.Paragraph({
        children: [image]
      })
    );

    let extractedText = '';

    try {

      const textContent =
        await page.getTextContent();

      extractedText =
        textContent.items
        .map(
          item => item.str
        )
        .join(' ');

      if (
        extractedText.trim().length <
        20
      ) {

        status.textContent =
          `OCR page ${pageNum}`;

        const result =
          await Tesseract.recognize(
            canvas,
            'eng+ind'
          );

        extractedText =
          result.data.text;
      }

    } catch (err) {

      console.error(err);
    }

    if (
      extractedText.trim()
    ) {

      children.push(
        new docx.Paragraph({
          text: extractedText
        })
      );
    }

    if (
      pageNum <
      pdf.numPages
    ) {

      children.push(
        new docx.Paragraph({
          pageBreakBefore: true
        })
      );
    }
  }

  status.textContent =
    'Generating DOCX...';

  const doc =
    new docx.Document({
      sections: [
        {
          children
        }
      ]
    });

  const blob =
    await docx.Packer.toBlob(
      doc
    );

  saveAs(
    blob,
    pdfFile.name.replace(
      /\.pdf$/i,
      '.docx'
    )
  );

  status.textContent =
    'Done.';

}
catch(error) {

  console.error(error);

  status.textContent =
    'Conversion failed.';
}
finally {

  convertBtn.disabled = false;
}

}

function dataURLToUint8Array(
  dataURL
) {

  const binary =
    atob(
      dataURL.split(',')[1]
    );

  const bytes =
    new Uint8Array(
      binary.length
    );

  for (
    let i = 0;
    i < binary.length;
    i++
  ) {
    bytes[i] =
      binary.charCodeAt(i);
  }

  return bytes;
}

})();
