import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import { describe, expect, it } from 'vitest';

import type { StandInFile } from '../../src/layout';
import {
  STAND_IN_FILES,
  createFontProvider,
  docxToHtml,
  docxToPages,
  docxToPdf,
  standInFor,
} from '../../src/layout';
import { diskLoader, fontFile, sampleDocx } from './sample';

const fonts = createFontProvider(diskLoader);
const latin1 = (bytes: Uint8Array): string => new TextDecoder('latin1').decode(bytes);

describe('docxToPdf', () => {
  it('writes a PDF of the whole report', async () => {
    const { bytes, pageCount, losses } = await docxToPdf(await sampleDocx(), { fonts });
    expect(latin1(bytes.subarray(0, 5))).toBe('%PDF-');
    expect(pageCount).toBeGreaterThanOrEqual(2);
    // One /Type /Page object per page.
    expect(latin1(bytes).match(/\/Type\s*\/Page\b/g)?.length).toBe(pageCount);
    expect(losses.filter((l) => l.severity === 'dropped')).toEqual([]);
  });

  it('lays a named family over its stand-in and says so', async () => {
    const { losses } = await docxToPdf(await sampleDocx({ font: 'Calibri', rows: 3 }), { fonts });
    expect(losses.map((l) => l.detail)).toContain('"calibri" drawn with its stand-in Carlito');
  });

  it('reports a stand-in nobody supplied instead of fetching it', async () => {
    const onlyArimo = createFontProvider((f) => (f.family === 'arimo' ? fontFile(f.file) : undefined));
    const { bytes, losses } = await docxToPdf(await sampleDocx({ font: 'Calibri', rows: 3 }), {
      fonts: onlyArimo,
    });
    expect(latin1(bytes.subarray(0, 5))).toBe('%PDF-');
    const loss = losses.find((l) => l.detail.startsWith('"calibri" drawn with Arimo'));
    expect(loss?.severity).toBe('substituted');
    expect(loss?.detail).toContain('Carlito Regular was not available (not supplied)');
  });

  it('reports a missing bold face as degraded', async () => {
    const noBold = createFontProvider((f) => (f.style === 'bold' ? undefined : fontFile(f.file)));
    const { losses } = await docxToPdf(await sampleDocx({ rows: 3 }), { fonts: noBold });
    expect(losses.find((l) => l.detail.startsWith('Arimo Bold was not available'))?.severity).toBe(
      'degraded',
    );
  });

  it('keeps the message of a loader that fails', async () => {
    const failing = createFontProvider((f) => {
      if (f.family === 'carlito') throw new Error('disk on fire');
      return fontFile(f.file);
    });
    const { losses } = await docxToPdf(await sampleDocx({ font: 'Calibri', rows: 3 }), {
      fonts: failing,
    });
    expect(losses.some((l) => l.detail.includes('the loader failed: disk on fire'))).toBe(true);
  });

  it('asks a failed loader again on the next conversion', async () => {
    let calls = 0;
    const flaky = createFontProvider((f) => {
      if (f.family === 'carlito' && f.style === 'regular' && calls++ === 0) throw new Error('blip');
      return fontFile(f.file);
    });
    const docx = await sampleDocx({ font: 'Calibri', rows: 3 });
    const first = await docxToPdf(docx, { fonts: flaky });
    expect(first.losses.some((l) => l.detail.includes('the loader failed: blip'))).toBe(true);
    const second = await docxToPdf(docx, { fonts: flaky });
    expect(second.losses.map((l) => l.detail)).toContain('"calibri" drawn with its stand-in Carlito');
  });

  it('asks the loader only for the families the document uses', async () => {
    const asked: Array<string> = [];
    const spy = createFontProvider((f: StandInFile) => {
      asked.push(f.file);
      return fontFile(f.file);
    });
    await docxToPdf(await sampleDocx({ font: 'Times New Roman', rows: 3 }), { fonts: spy });
    expect(new Set(asked.map((f) => f.split('-')[0]))).toEqual(new Set(['Arimo', 'Tinos']));
  });

  it('names no symbol loss for symbols the supplied faces draw', async () => {
    const docx = await sampleDocx({ rows: 3, extra: [{ text: 'Trend ▼ ► ■ ↑ made up' }] });
    const { losses } = await docxToPdf(docx, { fonts });
    expect(losses.filter((l) => /symbol/i.test(l.detail))).toEqual([]);
    expect(losses.filter((l) => l.severity === 'dropped')).toEqual([]);
  });

  it('still names symbols no supplied face has, by code point', async () => {
    // A tick box is drawn from the kit's face of boxes (src/layout/tickBoxes.ts); a telephone sign is not.
    const docx = await sampleDocx({ rows: 3, extra: [{ text: 'Done ☐ ☎ ▼' }] });
    const { losses } = await docxToPdf(docx, { fonts });
    const loss = losses.find((l) => /symbol text/.test(l.detail));
    expect(loss?.severity).toBe('degraded');
    expect(loss?.detail).toContain('U+260E');
    expect(loss?.detail).not.toContain('U+2610');
    expect(loss?.detail).not.toContain('U+25BC');
  });

  it('refuses a document it has no fonts for', async () => {
    await expect(docxToPdf(await sampleDocx({ rows: 3 }))).rejects.toThrow(/No fonts/);
  });

  it('refuses a file that is not a Word file', async () => {
    await expect(docxToPdf(new Uint8Array([1, 2, 3]), { fonts })).rejects.toThrow();
  });
});

describe('docxToPages', () => {
  it('draws one SVG per page, with the page numbers in the footer', async () => {
    const pdf = await docxToPdf(await sampleDocx(), { fonts });
    const { pages, pageSizes, pageText } = await docxToPages(await sampleDocx(), { fonts });
    expect(pages.length).toBe(pdf.pageCount);
    expect(pageSizes.length).toBe(pages.length);
    for (const svg of pages) {
      expect(svg.startsWith('<svg xmlns="http://www.w3.org/2000/svg"')).toBe(true);
      expect(svg.trimEnd().endsWith('</svg>')).toBe(true);
    }
    // The docx library writes A4 by default: 595.3 x 841.9 pt.
    expect(pageSizes[0]!.widthPt).toBeCloseTo(595.3, 0);
    expect(pageSizes[0]!.heightPt).toBeCloseTo(841.9, 0);
    const total = pages.length;
    pageText.forEach((text, i) => expect(text).toContain(`Page ${i + 1} of ${total}`));
  });

  it('gives each page its own ids, so pages can sit inline together', async () => {
    const { pages } = await docxToPages(await sampleDocx(), { fonts });
    const idsOf = (svg: string): Array<string> => [...svg.matchAll(/\sid="([^"]+)"/g)].map((m) => m[1]!);
    const first = idsOf(pages[0]!);
    expect(first.length).toBeGreaterThan(0);
    expect(first.every((id) => id.startsWith('dk-p1-'))).toBe(true);
    expect(idsOf(pages[1]!).every((id) => id.startsWith('dk-p2-'))).toBe(true);
    // Every reference points at an id the same page defines.
    for (const svg of pages) {
      const defined = new Set(idsOf(svg));
      for (const m of svg.matchAll(/(?:href="#|url\(#)([^")]+)/g)) expect(defined.has(m[1]!)).toBe(true);
    }
  });

  it('repeats the table header row on the next page', async () => {
    const { pageText } = await docxToPages(await sampleDocx(), { fonts });
    expect(pageText.length).toBeGreaterThanOrEqual(2);
    expect(pageText[1]).toContain('Measured value');
    expect(pageText[1]).not.toContain('Example results');
  });
});

describe('docxToHtml', () => {
  it('writes flowed HTML without any fonts', async () => {
    const { html, losses } = await docxToHtml(await sampleDocx({ rows: 3 }));
    expect(html).toContain('Example results');
    expect(html).toContain('S-003');
    expect(losses.some((l) => /header|footer/i.test(`${l.feature} ${l.detail}`))).toBe(true);
  });
});

describe('fonts', () => {
  it('maps the common Office faces by the engine table', () => {
    expect(standInFor('Calibri')).toBe('carlito');
    expect(standInFor('Calibri Light')).toBe('carlito');
    expect(standInFor('Cambria')).toBe('caladea');
    expect(standInFor('Times New Roman')).toBe('tinos');
    expect(standInFor('Georgia')).toBe('tinos');
    expect(standInFor('Arial')).toBe('arimo');
    expect(standInFor('Helvetica')).toBe('arimo');
    expect(standInFor('Verdana')).toBe('arimo');
    expect(standInFor('Courier New')).toBe('cousine');
    expect(standInFor('Consolas')).toBe('cousine');
  });

  it('names twenty files, all present in fonts/', () => {
    expect(STAND_IN_FILES.length).toBe(20);
    for (const f of STAND_IN_FILES) {
      expect(existsSync(fileURLToPath(new URL(`../../fonts/${f.file}`, import.meta.url)))).toBe(true);
    }
  });
});
