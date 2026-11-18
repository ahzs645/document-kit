// Keeping table rows together: a header row with the row after it, and a whole
// table on one page (docs/model.md, docs/ream-upstream.md change 9).

import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import type { TableRow } from '../../src/model';
import { document, headerRow, paragraph, row, section, spacer, table } from '../../src/model';
import { writeDocx } from '../../src/docx';
import { layOut, textOf } from './engine';

const page = { paper: 'letter' as const, margins: { top: 72, bottom: 72, left: 72, right: 72 } };

/** A page filled to `fill` points, then a table: a header row and body rows. */
async function tableAfter(fill: number, header: TableRow, bodyRows = 4, options: { keepTogether?: boolean } = {}) {
  const rows = [header, ...Array.from({ length: bodyRows }, (_, i) => row([`Body ${i + 1}`, `${i}.0`]))];
  return writeDocx(document([section([spacer(fill), table(rows, options)], { page })]));
}

const pageOf = (laid: ReturnType<typeof layOut>, text: string): number =>
  laid.pages.findIndex((p) => textOf(p).split('\n').some((line) => line.startsWith(text)));

/** The spacer heights, 2 points apart, that leave between 0 and 60 points under them. */
const FILLS = Array.from({ length: 31 }, (_, i) => 586 + i * 2);

describe('keeping a header row with the next row', () => {
  it('writes keep-with-next on a header row’s paragraphs, not on the last row', async () => {
    const xml = strFromU8(unzipSync(await tableAfter(0, headerRow(['Item', 'Value']), 2))['word/document.xml']!);
    const rows = xml.split('<w:tr>').slice(1);
    expect(rows[0]).toContain('<w:keepNext/>');
    expect(rows[0]).toContain('<w:tblHeader/>');
    expect(rows[1]).not.toContain('<w:keepNext/>');
    expect(rows[2]).not.toContain('<w:keepNext/>');
  });

  it('never leaves the header row at the foot of a page with its body on the next', async () => {
    for (const fill of FILLS) {
      const laid = layOut(await tableAfter(fill, headerRow(['Item', 'Value'])));
      expect({ fill, header: pageOf(laid, 'Item') }).toEqual({ fill, header: pageOf(laid, 'Body 1') });
    }
  });

  it('lets a header row part from its body when asked (the control)', async () => {
    let parted = 0;
    for (const fill of FILLS) {
      const laid = layOut(await tableAfter(fill, headerRow(['Item', 'Value'], { keepWithNext: false })));
      if (pageOf(laid, 'Item') !== pageOf(laid, 'Body 1')) parted += 1;
    }
    expect(parted).toBeGreaterThan(0);
  });

  it('still repeats the header row on the next page', async () => {
    const laid = layOut(await tableAfter(560, headerRow(['Item', 'Value']), 6));
    const header = pageOf(laid, 'Item');
    const last = pageOf(laid, 'Body 6');
    expect(last).toBeGreaterThan(header);
    expect(textOf(laid.pages[last]!).split('\n')[0]).toBe('Item');
  });
});

describe('keeping a table on one page', () => {
  it('writes every row but the last kept with the next, and every row whole', async () => {
    const xml = strFromU8(
      unzipSync(await tableAfter(0, row(['A', 'B']), 3, { keepTogether: true }))['word/document.xml']!,
    );
    const rows = xml.split('<w:tr>').slice(1);
    expect(rows.map((r) => r.includes('<w:keepNext/>'))).toEqual([true, true, true, false]);
    expect(rows.every((r) => r.includes('<w:cantSplit/>'))).toBe(true);
  });

  it('moves a small table whole to the next page rather than splitting it', async () => {
    for (const fill of [500, 540, 580, 620]) {
      const laid = layOut(await tableAfter(fill, row(['First', 'row']), 4, { keepTogether: true }));
      const first = pageOf(laid, 'First');
      expect({ fill, last: pageOf(laid, 'Body 4') }).toEqual({ fill, last: first });
    }
  });

  it('splits the same table where it is not kept together (the control)', async () => {
    const pages = [];
    for (const fill of [500, 540, 580, 620]) {
      const laid = layOut(await tableAfter(fill, row(['First', 'row']), 4));
      pages.push(pageOf(laid, 'Body 4') - pageOf(laid, 'First'));
    }
    expect(pages.some((d) => d > 0)).toBe(true);
  });
});
