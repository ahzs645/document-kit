// A band of a picture (crop), a frame round it, and turned text in a table cell,
// written by the kit's Word writer and read back by the layout.

import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import type { BorderItem, ImageItem, TextLineItem } from '../../src/ream/layout/page-doc';
import { document, figure, image, row, section, table } from '../../src/model';
import { writeDocx } from '../../src/docx';
import { makePng } from '../docx/helpers';
import { layOut } from './engine';

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
