// §17.6.23 `w:vAlign`: a section's text set in the middle or at the foot of each
// of its pages (docs/ream-upstream.md, change 6).

import { describe, expect, it } from 'vitest';

import { document, paragraph, section } from '../../src/model';
import { writeDocx } from '../../src/docx';
import { layOut, linesOf } from './engine';

// Letter, one-inch margins: the text area runs from 72pt to 720pt down the page.
const TOP = 72;
const BOTTOM = 792 - 72;

async function cover(align?: 'top' | 'center' | 'bottom') {
  return writeDocx(
    document([
      section([paragraph('A made-up cover', { style: 'title' }), paragraph('Prepared for nobody')], {
        page: { paper: 'letter', margins: { top: 72, bottom: 72, left: 72, right: 72 }, ...(align ? { verticalAlign: align } : {}) },
      }),
      section([paragraph('The body starts here.'), paragraph('And goes on.')]),
    ]),
  );
}

describe('a section’s vertical alignment', () => {
  it('sets a centred cover page in the middle of the room between the margins', async () => {
    const laid = layOut(await cover('center'));
    const lines = linesOf(laid.pages[0]!);
    const first = lines[0]!.baselineY;
    const last = lines[lines.length - 1]!.baselineY;
    // The block's middle is within a line of the text area's middle.
    expect(Math.abs((first + last) / 2 - (TOP + BOTTOM) / 2)).toBeLessThan(20);
    expect(first).toBeGreaterThan(300);
  });

  it('sets a bottom-aligned page at the foot of the text area', async () => {
    const laid = layOut(await cover('bottom'));
    const lines = linesOf(laid.pages[0]!);
    expect(lines[lines.length - 1]!.baselineY).toBeGreaterThan(BOTTOM - 30);
    expect(lines[lines.length - 1]!.baselineY).toBeLessThanOrEqual(BOTTOM);
  });

  it('leaves a top-aligned page, and the next section, at the top', async () => {
    const top = layOut(await cover());
    expect(linesOf(top.pages[0]!)[0]!.baselineY).toBeLessThan(TOP + 40);
    const centred = layOut(await cover('center'));
    // The second section states no alignment: its page starts at the top.
    expect(linesOf(centred.pages[1]!)[0]!.baselineY).toBeLessThan(TOP + 30);
    expect(linesOf(centred.pages[1]!)[0]!.baselineY).toBe(linesOf(top.pages[1]!)[0]!.baselineY);
  });
});
