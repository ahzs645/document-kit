// The layout API in a real browser: a report written with the `docx` library
// in the page, laid out to a PDF and to page pictures with the bundled fonts.

import {
  Document,
  Footer,
  HeadingLevel,
  Packer,
  PageNumber,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
} from 'docx';
import { createFontProvider, docxToHtml, docxToPages, docxToPdf } from 'document-kit/layout';

function check(condition, message) {
  if (!condition) throw new Error(message);
}

function report() {
  const cell = (text, bold = false) =>
    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text, bold })] })] });
  const header = new TableRow({
    tableHeader: true,
    children: [cell('Sample', true), cell('Measured value', true), cell('Note', true)],
  });
  const rows = Array.from(
    { length: 70 },
    (_, i) =>
      new TableRow({
        children: [
          cell(`S-${String(i + 1).padStart(3, '0')}`),
          cell(`${(i * 1.7).toFixed(1)} units`),
          cell('Within range'),
        ],
      }),
  );
  return new Document({
    sections: [
      {
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                // One run holding the whole field, as the docx library writes it.
                children: [
                  new TextRun({
                    children: ['Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES],
                  }),
                ],
              }),
            ],
          }),
        },
        children: [
          new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('Example results')] }),
          new Paragraph({
            children: [new TextRun('A made-up table of results, long enough to need two pages.')],
          }),
          new Table({ rows: [header, ...rows] }),
        ],
      },
    ],
  });
}

export default async function layoutCase(ctx) {
  const docx = new Uint8Array(await Packer.toArrayBuffer(report()));
  const asked = new Set();
  const fonts = createFontProvider((file) => {
    asked.add(file.file);
    return ctx.loadFont(file.file);
  });

  const started = performance.now();
  const pdf = await docxToPdf(docx, { fonts });
  const pdfMs = Math.round(performance.now() - started);
  const head = new TextDecoder('latin1').decode(pdf.bytes.subarray(0, 5));
  check(head === '%PDF-', `the PDF starts with ${JSON.stringify(head)}`);
  check(pdf.pageCount >= 2, `the PDF has ${pdf.pageCount} page(s), expected at least 2`);

  const { pages, pageSizes, pageText, losses } = await docxToPages(docx, { fonts });
  check(pages.length === pdf.pageCount, `${pages.length} SVG pages for a ${pdf.pageCount}-page PDF`);
  check(pageSizes.length === pages.length, 'one size per page');
  pages.forEach((svg, i) => check(svg.startsWith('<svg'), `page ${i + 1} is not an SVG`));
  pageText.forEach((text, i) =>
    check(
      text.includes(`Page ${i + 1} of ${pages.length}`),
      `page ${i + 1} does not print "Page ${i + 1} of ${pages.length}"`,
    ),
  );
  check(pageText[1].includes('Measured value'), 'the header row is not repeated on page 2');

  // The SVG must show in this browser: parse it, and load it as a picture.
  const doc = new DOMParser().parseFromString(pages[0], 'image/svg+xml');
  check(!doc.querySelector('parsererror'), 'page 1 SVG does not parse');
  const url = URL.createObjectURL(new Blob([pages[0]], { type: 'image/svg+xml' }));
  const img = new Image();
  await new Promise((resolve, reject) => {
    img.onload = resolve;
    img.onerror = () => reject(new Error('page 1 SVG does not load as an image'));
    img.src = url;
  });
  URL.revokeObjectURL(url);
  check(Math.round(img.naturalWidth) === Math.round(pageSizes[0].widthPt), 'page 1 image has the wrong width');

  const { html } = await docxToHtml(docx);
  check(html.includes('Example results'), 'the HTML lacks the heading');

  return {
    pdfBytes: pdf.bytes.length,
    pdfMs,
    pages: pages.length,
    pageSize: [Math.round(pageSizes[0].widthPt), Math.round(pageSizes[0].heightPt)],
    footer: pageText.map((t) => t.split('\n').at(-1)),
    fontsLoaded: [...asked].sort(),
    losses: losses.length,
  };
}
