// The whole chain, each piece from its own entry: a document built with the
// block model, written as a Word file, laid out to a PDF and to page pictures,
// with only the bundled fonts and no network. And the template path: a Word
// file the kit wrote, carrying tags, filled from data and laid out.

import { mkdirSync, writeFileSync } from 'node:fs';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { document, heading, paragraph, section, table, headerRow, row, pageNumber, pageCount, text } from '../../src';
import { writeDocx } from '../../src/docx';
import { createFontProvider, docxToPages, docxToPdf } from '../../src/layout';
import { fillTemplate } from '../../src/template';
import { diskLoader } from '../layout/sample';
import { laboratoryReport, reviewReport, reviewRegistry } from '../docx/fixtures';

const OUT = process.env.DOCUMENT_KIT_SAMPLES;
const keep = (name: string, bytes: Uint8Array): void => {
  if (!OUT) return;
  mkdirSync(OUT, { recursive: true });
  writeFileSync(`${OUT}/${name}`, bytes);
};

const fetchSpy = vi.fn(() => {
  throw new Error('fetch was called');
});
beforeEach(() => vi.stubGlobal('fetch', fetchSpy));
afterEach(() => vi.unstubAllGlobals());

describe('model → Word → PDF and pages', () => {
  it('a laboratory-style report', async () => {
    const docx = await writeDocx(laboratoryReport());
    keep('laboratory.docx', docx);
    const fonts = createFontProvider(diskLoader);
    const pdf = await docxToPdf(docx, { fonts });
    keep('laboratory.pdf', pdf.bytes);
    expect(new TextDecoder().decode(pdf.bytes.slice(0, 5))).toBe('%PDF-');
    expect(pdf.pageCount).toBeGreaterThanOrEqual(2);
    const pages = await docxToPages(docx, { fonts });
    expect(pages.pages).toHaveLength(pdf.pageCount);
    // The footer's fields are filled on every page.
    pages.pageText.forEach((t, i) => expect(t).toContain(`Page ${i + 1} of ${pdf.pageCount}`));
    expect(fetchSpy).not.toHaveBeenCalled();
  }, 60_000);

  it('a review-style report with landscape and 11 × 17 sections', async () => {
    const docx = await writeDocx(reviewReport(), { registry: reviewRegistry });
    keep('review.docx', docx);
    const fonts = createFontProvider(diskLoader);
    const pdf = await docxToPdf(docx, { fonts });
    keep('review.pdf', pdf.bytes);
    const { pageSizes } = await docxToPages(docx, { fonts });
    const landscape = pageSizes.filter((p) => p.widthPt > p.heightPt);
    expect(landscape.length).toBeGreaterThanOrEqual(2);
    // 11 × 17 inches on its side.
    expect(pageSizes.some((p) => Math.round(p.widthPt) === 1224 && Math.round(p.heightPt) === 792)).toBe(true);
    expect(fetchSpy).not.toHaveBeenCalled();
  }, 60_000);
});

describe('template written by the kit → filled → laid out', () => {
  it('fills tags in a Word file the model wrote', async () => {
    const template = await writeDocx(
      document([
        section(
          [
            heading(1, 'Results for {{ client }}'),
            paragraph('{{ if interim }}Interim: {{ pending }} to follow.{{ endif }}'),
            table([
              headerRow(['Item', 'Value']),
              row(['{{ for r in rows }}{{ r.name }}', '{{ r.value | format "0.00" }}{{ endfor }}']),
            ]),
          ],
          { footer: [paragraph([text('{{ reference }}  Page '), pageNumber(), text(' of '), pageCount()])] },
        ),
      ]),
    );
    const rows = Array.from({ length: 70 }, (_, i) => ({ name: `Item ${i + 1}`, value: (i + 1) * 1.5 }));
    const filled = await fillTemplate(template, {
      data: { reference: 'REF-1', client: 'A. Sample', interim: true, pending: 'two items', rows },
      onMissing: 'error',
    });
    keep('template-filled.docx', filled);
    const pages = await docxToPages(filled, { fonts: createFontProvider(diskLoader) });
    const all = pages.pageText.join('\n');
    expect(all).toContain('Results for A. Sample');
    expect(all).toContain('Interim: two items to follow.');
    expect(all).toContain('Item 70');
    expect(all).toContain('105.00');
    expect(all).not.toContain('{{');
    expect(pages.pages.length).toBeGreaterThanOrEqual(2);
    pages.pageText.forEach((t, i) => expect(t).toContain(`REF-1  Page ${i + 1} of ${pages.pages.length}`));
    expect(fetchSpy).not.toHaveBeenCalled();
  }, 60_000);
});
