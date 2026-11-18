// §17.4.66 — the border shown where two table cells share an edge, and the
// order rules are painted in (docs/ream-upstream.md, change 8).

import { describe, expect, it } from 'vitest';

import type { BorderItem } from '../../src/ream/layout/page-doc';
import { document, paragraph, row, section, table } from '../../src/model';
import { writeDocx } from '../../src/docx';
import { edgeWinner } from '../../src/ream/core/border-conflict';
import { pt } from '../../src/ream/core/ir';
import { paintPlan } from '../../src/ream/layout/page-doc';
import { layOut } from './engine';

const red = { width: 3, color: 'FF0000' };
const box = { top: red, bottom: red, left: red, right: red };

/** A 3 × 3 table, the middle cell boxed in thick red, the rest on the table's defaults. */
async function redMiddle(): Promise<ReturnType<typeof layOut>> {
  return layOut(
    await writeDocx(
      document([
        section([
          paragraph('Above.'),
          table([row(['a', 'b', 'c']), row(['d', { content: 'red', borders: box }, 'f']), row(['g', 'h', 'i'])]),
        ]),
      ]),
    ),
  );
}

const isRed = (b: BorderItem) => b.borderColorHex.toUpperCase() === 'FF0000';

describe('a thick red cell border beside default-bordered neighbours', () => {
  it('shows red on all four edges of the cell', async () => {
    const borders = (await redMiddle()).pages[0]!.commands.filter((c): c is BorderItem => c.type === 'border');
    const reds = borders.filter(isRed);
    // The middle cell's own top and left, its right as the next cell's left,
    // its bottom as the cell below's top.
    expect(reds.map((b) => b.side).sort()).toEqual(['left', 'left', 'top', 'top']);
    expect(reds.every((b) => b.borderSizePt === 3)).toBe(true);
  });

  it('paints every red rule after the thinner rules that run into its ends', async () => {
    const plan = paintPlan((await redMiddle()).pages[0]!.commands);
    const firstRed = plan.borders.findIndex(isRed);
    expect(firstRed).toBeGreaterThan(0);
    // Every rule painted after the first red one is red: no default rule's
    // projecting end is laid over a corner of the box.
    expect(plan.borders.slice(firstRed).every(isRed)).toBe(true);
  });
});

describe('§17.4.66 conflict rules', () => {
  const thin = { style: 'single' as const, width: pt(0.5), colorHex: '000000' };
  const thick = { style: 'single' as const, width: pt(3), colorHex: 'CBD2D9' };

  it('shows a cell’s own border before the table’s, however light', () => {
    expect(edgeWinner({ border: thin, fromCell: true }, { border: thick, fromCell: false })).toBe(thin);
    expect(edgeWinner({ border: thick, fromCell: false }, { border: thin, fromCell: true })).toBe(thin);
  });

  it('lets a side with none give way, even a cell’s own', () => {
    const none = { style: 'none' as const };
    expect(edgeWinner({ border: none, fromCell: true }, { border: thick, fromCell: false })).toBe(thick);
  });

  it('weighs width in eighths times the style’s number', () => {
    const double = { style: 'double' as const, width: pt(0.5) }; // 4 × 3 = 12
    const single = { style: 'single' as const, width: pt(1) }; // 8 × 1 = 8
    expect(edgeWinner({ border: single, fromCell: true }, { border: double, fromCell: true })).toBe(double);
  });

  it('breaks an equal weight by the earlier style, then by the darker colour', () => {
    const single = { style: 'single' as const, width: pt(2) }; // 16 × 1
    const thickStyle = { style: 'thick' as const, width: pt(1) }; // 8 × 2
    expect(edgeWinner({ border: thickStyle, fromCell: true }, { border: single, fromCell: true })).toBe(single);
    const grey = { style: 'single' as const, width: pt(1), colorHex: 'CCCCCC' };
    const darkRed = { style: 'single' as const, width: pt(1), colorHex: 'C00000' };
    expect(edgeWinner({ border: grey, fromCell: true }, { border: darkRed, fromCell: true })).toBe(darkRed);
  });

  it('lays a thin cell border over a heavier table rule on the page', async () => {
    const laid = layOut(
      await writeDocx(
        document([
          section([
            table([row(['a', { content: 'b', borders: { left: { width: 0.5, color: '000000' } } }])], {
              borders: { insideVertical: { width: 3, color: 'CBD2D9' } },
            }),
          ]),
        ]),
      ),
    );
    const borders = laid.pages[0]!.commands.filter((c): c is BorderItem => c.type === 'border' && c.side === 'left');
    // The shared edge is cell b's left: its own thin black rule, not the table's 3pt grey.
    const shared = borders.reduce((a, b) => (b.x > a.x ? b : a));
    expect(shared.borderColorHex).toBe('000000');
    expect(shared.borderSizePt).toBe(0.5);
  });
});
