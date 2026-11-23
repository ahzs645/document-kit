// One layout for both outputs (layoutDocx), and page pictures that share their
// glyph outlines (docs/layout.md).

import { describe, expect, it, vi } from 'vitest';

import { createFontProvider, docxToPages, docxToPdf, layoutDocx } from '../../src/layout';
import { layoutStyledDocument } from '../../src/ream/layout/styled-layout';
import { diskLoader, sampleDocx } from './sample';

vi.mock('../../src/ream/layout/styled-layout', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../src/ream/layout/styled-layout')>();
  return { ...actual, layoutStyledDocument: vi.fn(actual.layoutStyledDocument) };
});

const fonts = createFontProvider(diskLoader);
const layoutCalls = () => vi.mocked(layoutStyledDocument).mock.calls.length;

describe('layoutDocx', () => {
  it('lays the document out once for the PDF and the pages', async () => {
    const docx = await sampleDocx();
    const before = layoutCalls();
    const laid = await layoutDocx(docx, { fonts });
    const pdf = laid.toPdf();
    const pages = laid.toPages();
    laid.toPdf();
    expect(layoutCalls() - before).toBe(1);
    expect(pdf.pageCount).toBe(laid.pageCount);
    expect(pdf.pageSizes).toEqual(pages.pageSizes);
    expect(pages.pages.length).toBe(laid.pageCount);
    expect(laid.pageSizes.length).toBe(laid.pageCount);
    expect(laid.pageText.length).toBe(laid.pageCount);
  });

  it('writes what docxToPdf and docxToPages write', async () => {
    const docx = await sampleDocx({ font: 'Calibri' });
    const laid = await layoutDocx(docx, { fonts });
    const info = { title: 'A made-up report' };
    const separatePdf = await docxToPdf(docx, { fonts, info });
    const separatePages = await docxToPages(docx, { fonts });
    expect(laid.toPdf({ info }).bytes).toEqual(separatePdf.bytes);
    expect(laid.toPdf({ info }).losses).toEqual(separatePdf.losses);
    expect(laid.toPages()).toEqual(separatePages);
    expect(laid.pageText).toEqual(separatePages.pageText);
    expect(laid.losses).toEqual(separatePdf.losses);
  });

  it('lays out once for each of docxToPdf and docxToPages', async () => {
    const docx = await sampleDocx({ rows: 3 });
    const before = layoutCalls();
    await docxToPdf(docx, { fonts });
    await docxToPages(docx, { fonts });
    expect(layoutCalls() - before).toBe(2);
  });
});

describe('page pictures with shared glyph outlines', () => {
  it('puts each outline once in glyphDefs and has every page refer to it', async () => {
    const docx = await sampleDocx();
    const own = (await layoutDocx(docx, { fonts })).toPages();
    const shared = (await layoutDocx(docx, { fonts })).toPages({ sharedGlyphs: true });
    expect(own.glyphDefs).toBeUndefined();
    expect(shared.pages.length).toBe(own.pages.length);
    const defs = shared.glyphDefs!;
    expect(defs.startsWith('<svg xmlns="http://www.w3.org/2000/svg"')).toBe(true);
    const defined = new Set([...defs.matchAll(/<path id="(dk-g\d+)"/g)].map((m) => m[1]!));
    const paths = [...defs.matchAll(/<path id="dk-g\d+" d="([^"]*)"/g)].map((m) => m[1]!);
    expect(new Set(paths).size).toBe(paths.length); // each outline once
    for (const svg of shared.pages) {
      expect(svg).not.toMatch(/<path id="dk-p\d+-g/);
      for (const m of svg.matchAll(/href="#(dk-g\d+)"/g)) expect(defined.has(m[1]!)).toBe(true);
      // The page's own ids (clips and the like) still resolve on the page.
      const ownIds = new Set([...svg.matchAll(/\sid="([^"]+)"/g)].map((m) => m[1]!));
      for (const m of svg.matchAll(/(?:href="#|url\(#)(dk-p[^")]+)/g)) expect(ownIds.has(m[1]!)).toBe(true);
    }
    // Outlines are a part of each page, not most of it (one <use> per glyph drawn
    // is the rest): on this two-page sample sharing saves about a sixth.
    const size = (list: ReadonlyArray<string>) => list.reduce((n, s) => n + s.length, 0);
    expect(size(shared.pages) + defs.length).toBeLessThan(size(own.pages) * 0.9);
    const ownOutlines = own.pages.reduce((n, svg) => n + (svg.match(/<path id="dk-p\d+-g/g)?.length ?? 0), 0);
    expect(defined.size).toBeLessThan(ownOutlines);
  });

  it('keeps each page a standalone picture by default', async () => {
    const { pages } = await docxToPages(await sampleDocx({ rows: 3 }), { fonts });
    expect(pages[0]).toMatch(/<path id="dk-p1-g\d+-\d+"/);
    expect(pages[0]).not.toContain('href="#dk-g');
  });
});
