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

function setStatus(text) {
    status.textContent = text;
}

dropzone.addEventListener('click', () => {
    input.click();
});

dropzone.querySelector('.link-btn')
.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    input.click();
});

input.addEventListener('change', () => {

    const file = input.files[0];

    if (!file) return;

    pdfFile = file;

    actions.classList.remove('hidden');

    setStatus(file.name);
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

    setStatus(file.name);
});

clearBtn.addEventListener('click', () => {

    pdfFile = null;

    input.value = '';

    actions.classList.add('hidden');

    setStatus('');
});

convertBtn.addEventListener('click', convertPdfToWord);

async function convertPdfToWord() {

    if (!pdfFile) return;

    try {

        convertBtn.disabled = true;

        setStatus('Loading PDF...');

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

            setStatus(
                `Rendering page ${pageNum}/${pdf.numPages}`
            );

            const page =
                await pdf.getPage(pageNum);

            const viewport =
                page.getViewport({
                    scale: 2.5
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

            const blob =
                await new Promise(resolve =>
                    canvas.toBlob(
                        resolve,
                        'image/png'
                    )
                );

            const imageBuffer =
                await blob.arrayBuffer();

            const image = new docx.ImageRun({
                data: imageBuffer,
                transformation: {
                    width: 550,
                    height: Math.round(
                        550 *
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

            if (pageNum < pdf.numPages) {

                children.push(
                    new docx.Paragraph({
                        pageBreakBefore: true
                    })
                );

            }
        }

        setStatus('Building DOCX...');

        const doc =
            new docx.Document({
                sections: [{
                    children
                }]
            });

        const docBlob =
            await docx.Packer.toBlob(doc);

        saveAs(
            docBlob,
            pdfFile.name.replace(
                /\.pdf$/i,
                '.docx'
            )
        );

        setStatus(
            'Finished.'
        );

    }
    catch (err) {

        console.error(err);

        setStatus(
            'Conversion failed.'
        );
    }
    finally {

        convertBtn.disabled = false;
    }
}

})();
