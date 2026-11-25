// document-kit/layout writes compact PDFs: each page's drawing deflated and a palette picture kept as a palette
// (an /Indexed image). The engine's own default is unchanged (see tests/ream/byte-gate.test.ts); the kit's front
// turns `compact` on. These read the PDF's objects back and check that nothing on the pages changed.

import { unzlibSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { document, figure, heading, image, paragraph, section } from '../../src/model';
import { writeDocx } from '../../src/docx';
import { createFontProvider, docxToPdf } from '../../src/layout';
import { buildIndexedPng } from '../ream/fixtures/build-png';
import { diskLoader } from './sample';

const fonts = createFontProvider(diskLoader);
const latin1 = new TextDecoder('latin1');

interface PdfObject { num: number; dict: string; data: Uint8Array | null }

/** Each object of a PDF, its stream inflated when deflated. */
function objects(pdf: Uint8Array): PdfObject[] {
  const text = latin1.decode(pdf);
  const out: PdfObject[] = [];
  for (const m of text.matchAll(/(?:^|\n)(\d+) 0 obj\s*(<<[\s\S]*?>>)\s*(stream\r?\n|endobj)/g)) {
    const num = Number(m[1]);
    const dict = m[2]!;
    if (!m[3]!.startsWith('stream')) { out.push({ num, dict, data: null }); continue; }
    const from = m.index! + m[0].length;
    const length = Number(/\/Length (\d+)/.exec(dict)![1]);
    const raw = pdf.subarray(from, from + length);
    out.push({ num, dict, data: /\/FlateDecode/.test(dict) ? unzlibSync(raw) : raw });
  }
  return out;
}

/** Each page's content stream, found through the page's /Contents. */
const pageContents = (pdf: Uint8Array) => {
  const all = objects(pdf);
  const refs = all.filter((o) => /\/Type \/Page\b/.test(o.dict)).map((o) => Number(/\/Contents (\d+) 0 R/.exec(o.dict)![1]));
  return refs.map((n) => all.find((o) => o.num === n)!);
};

const report = () =>
  writeDocx(document([section([heading(1, 'Compact'), ...Array.from({ length: 60 }, (_, i) => paragraph(`Line ${i + 1} of a long report, drawn the same either way.`))])]));

describe('a compact PDF', () => {
  it('deflates every page’s drawing', async () => {
    const { bytes, pageCount } = await docxToPdf(await report(), { fonts });
    const contents = pageContents(bytes);
    expect(contents).toHaveLength(pageCount);
    for (const c of contents) {
      expect(c.dict).toContain('/FlateDecode');
      // Still the page's drawing: text objects with glyphs shown.
      expect(latin1.decode(c.data!)).toMatch(/BT[\s\S]*Tj|TJ[\s\S]*ET/);
    }
  });

  it('keeps a palette picture as a palette, pixel for pixel', async () => {
    const BLUE = [28, 92, 171] as const, RED = [239, 159, 158] as const, WHITE = [255, 255, 255] as const;
    const w = 12, h = 4;
    const indices = Array.from({ length: w * h }, (_, i) => (i % w < 4 ? 0 : i % w < 8 ? 1 : 2));
    const png = buildIndexedPng(w, h, 8, [BLUE, RED, WHITE], indices);
    const docx = await writeDocx(document([section([figure(image(png))])]));
    const pdf = (await docxToPdf(docx, { fonts })).bytes;
    const picture = objects(pdf).find((o) => o.dict.includes('/Subtype /Image'))!;
    expect(picture.dict).toMatch(/\/ColorSpace \[\/Indexed \/DeviceRGB 2 <1c5cabef9f9effffff>\]/i);
    expect([...picture.data!]).toEqual(indices);
    expect(picture.dict).not.toContain('/SMask');
  });
});
