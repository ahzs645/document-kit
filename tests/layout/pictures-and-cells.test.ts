// A band of a picture (crop), a frame round it, and turned text in a table cell,
// written by the kit's Word writer and read back by the layout.

import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import type { BorderItem, ImageItem, TextLineItem } from '../../src/ream/layout/page-doc';
import { document, figure, image, row, section, table } from '../../src/model';
import { writeDocx } from '../../src/docx';
import { makePng } from '../docx/helpers';
import { layOut } from './engine'
import { buildDocxFromBody } from '../ream/fixtures/build-docx';

const band = () =>
  writeDocx(
    document([
      section([
        figure(image(makePng(200, 100)), { crop: { top: 0.25, bottom: 0.5 }, border: { width: 2, color: '#c00' } }),
        table([row([{ content: 'Upward', textDirection: 'up' }, { content: 'Down', textDirection: 'down' }, 'flat'])]),
      ]),
    ]),
  );

describe('a band of a picture', () => {
  it('writes Word’s crop and fits the part shown', async () => {
    const xml = strFromU8(unzipSync(await band())['word/document.xml']!);
    expect(xml).toContain('<a:srcRect l="0" t="25000" r="0" b="50000"/>');
    // 200 × 25 pixels shown: 150 × 18.75 points.
    expect(xml).toContain('<wp:extent cx="1905000" cy="238125"/>');
  });

  it('is laid out as that band', async () => {
    const page = layOut(await band()).pages[0]!;
    const picture = page.commands.find((c): c is ImageItem => c.type === 'image')!;
    expect(picture.crop).toEqual({ left: 0, top: 0.25, right: 0, bottom: 0.5 });
    expect(picture.width).toBeCloseTo(150, 3);
    expect(picture.height).toBeCloseTo(18.75, 3);
  });

  it('refuses a crop that leaves nothing, naming it', async () => {
    const doc = document([section([figure(image(makePng(20, 10)), { crop: { left: 0.6, right: 0.4 } })])]);
    await expect(writeDocx(doc)).rejects.toThrow('crop {"left":0.6,"right":0.4}');
  });
});

describe('a frame round a picture', () => {
  it('is written as the picture’s outline and drawn on its four edges', async () => {
    const xml = strFromU8(unzipSync(await band())['word/document.xml']!);
    expect(xml).toMatch(/<a:ln w="25400"[^>]*>.*?<a:srgbClr val="CC0000"\/>/s);
    const frame = layOut(await band()).pages[0]!.commands.filter(
      (c): c is BorderItem => c.type === 'border' && c.borderColorHex === 'CC0000',
    );
    expect(frame.map((b) => b.side).sort()).toEqual(['bottom', 'left', 'right', 'top']);
    expect(frame.every((b) => b.borderSizePt === 2)).toBe(true);
  });
});

describe('turned text in a table cell', () => {
  it('is written as Word’s text direction and laid out a quarter turned', async () => {
    const xml = strFromU8(unzipSync(await band())['word/document.xml']!);
    expect(xml).toContain('<w:textDirection w:val="btLr"/>');
    expect(xml).toContain('<w:textDirection w:val="tbRl"/>');
    const lines = layOut(await band()).pages[0]!.commands.filter((c): c is TextLineItem => c.type === 'line');
    const text = (l: TextLineItem) => l.line.tokens.map((t) => (t.kind === 'text' ? t.text : '')).join('');
    expect(lines.find((l) => text(l) === 'Upward')?.rotationDeg).toBe(90);
    expect(lines.find((l) => text(l) === 'Down')?.rotationDeg).toBe(-90);
    expect(lines.find((l) => text(l) === 'flat')?.rotationDeg).toBeUndefined();
  });
});

describe('a table width written as a percentage with its sign', () => {
  // As the `docx` library wrote a placeholder box up to its 9.7 releases: the width as "86%" and a
  // one-column grid of 100 twips that only the percentage overrides.
  const box = (width: string) =>
    buildDocxFromBody(
      `<w:tbl><w:tblPr><w:tblW w:type="pct" w:w="${width}"/><w:jc w:val="center"/></w:tblPr>` +
        '<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:type="pct" w:w="100%"/></w:tcPr>' +
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Picture to come: a made-up placeholder caption</w:t></w:r></w:p>' +
        '</w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>',
    );
  const linesIn = (docx: Uint8Array) =>
    layOut(docx).pages[0]!.commands.filter((c): c is TextLineItem => c.type === 'line');

  it('reads "86%" as 86 per cent of the text width, as it reads 4300 fiftieths', () => {
    const signed = linesIn(box('86%'));
    const fiftieths = linesIn(box('4300'));
    // One line, not a word a line in a box five points wide.
    expect(signed).toHaveLength(1);
    expect(signed[0]!.line.availableWidthPt).toBeCloseTo(fiftieths[0]!.line.availableWidthPt, 1);
    expect(signed[0]!.line.availableWidthPt).toBeGreaterThan(0.8 * 468);
  });
});

describe('a table whose grid is a stand-in', () => {
  // The `docx` library, given no column widths, writes 100-twip grid columns beside the table's width and the cells'
  // own widths: here a 6.5-inch table of three columns asking for a half and two quarters.
  const table = (grid: string) =>
    buildDocxFromBody(
      '<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="9360"/></w:tblPr>' +
        `<w:tblGrid>${grid}</w:tblGrid><w:tr>` +
        ['4680', '2340', '2340']
          .map((w, i) => `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="${w}"/></w:tcPr><w:p><w:r><w:t>Cell ${i + 1}</w:t></w:r></w:p></w:tc>`)
          .join('') +
        '</w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>',
    );
  const starts = (docx: Uint8Array) =>
    layOut(docx)
      .pages[0]!.commands.filter((c): c is TextLineItem => c.type === 'line')
      .map((line) => Math.round(line.originX));

  it('takes the cells’ widths, as Word does, not equal columns from the grid', () => {
    const [first, second, third] = starts(table('<w:gridCol w:w="100"/>'.repeat(3)))
    // A half and two quarters of 468pt: the columns start 234pt and 117pt apart.
    expect(second! - first!).toBe(234);
    expect(third! - second!).toBe(117);
  });

  it('keeps a grid that is the table’s own', () => {
    const [first, second, third] = starts(table('<w:gridCol w:w="3120"/>'.repeat(3)))
    expect(second! - first!).toBe(156);
    expect(third! - second!).toBe(156);
  });
});
