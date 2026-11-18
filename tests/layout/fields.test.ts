// Page-number fields in the body: a contents list's PAGEREFs, and PAGE and
// NUMPAGES in body text (docs/ream-upstream.md, change 7).

import { describe, expect, it } from 'vitest';

import { document, heading, pageBreak, pageCount, pageNumber, paragraph, section, toc } from '../../src/model';
import { writeDocx } from '../../src/docx';
import { MAX_FIELD_PASSES } from '../../src/ream/layout/styled-layout';
import { layOut, linesOf, textOf } from './engine';

/** A made-up report: a roman-numbered front section with the contents, then three chapters. */
async function report() {
  return writeDocx(
    document([
      section([paragraph('Front matter'), toc({ title: 'Contents', levels: 2 }), paragraph(['This front page is ', pageNumber(), '.'])], {
        pageNumbers: { start: 1, format: 'lowerRoman' },
      }),
      section(
        [
          heading(1, 'Alpha chapter'),
          paragraph('First words.'),
          pageBreak(),
          paragraph('Filler.'),
          pageBreak(),
          heading(2, 'Beta part'),
          paragraph(['This is page ', pageNumber(), ' of ', pageCount(), '.']),
          pageBreak(),
          heading(1, 'Gamma chapter'),
        ],
        { pageNumbers: { start: 1 } },
      ),
    ]),
  );
}

/** The page (0-based) whose first line is exactly `text`. */
const pageStartingWith = (pages: ReturnType<typeof layOut>['pages'], text: string): number =>
  pages.findIndex((p) => linesOf(p)[0]?.text === text);

describe('page-number fields in the body', () => {
  it('fills each contents entry with the page its heading lands on, in that page’s numbering', async () => {
    const laid = layOut(await report());
    const alpha = pageStartingWith(laid.pages, 'Alpha chapter');
    const beta = pageStartingWith(laid.pages, 'Beta part');
    const gamma = pageStartingWith(laid.pages, 'Gamma chapter');
    expect([alpha, beta, gamma]).toEqual([1, 3, 4]);
    const contents = linesOf(laid.pages[0]!).map((l) => l.text);
    // The body section restarts at 1, so the headings print 1, 3 and 4.
    expect(contents.find((t) => t.startsWith('Alpha chapter'))).toMatch(/\t?\D*1$/u);
    expect(contents.find((t) => t.startsWith('Beta part'))).toMatch(/\D3$/u);
    expect(contents.find((t) => t.startsWith('Gamma chapter'))).toMatch(/\D4$/u);
    // Gamma (level 1) is listed; nothing deeper than level 2 is.
    expect(laid.losses.filter((l) => l.detail.includes('did not settle'))).toEqual([]);
  });

  it('prints PAGE and NUMPAGES in the body as the real numbers, PAGE in its section’s format', async () => {
    const laid = layOut(await report());
    expect(textOf(laid.pages[0]!)).toContain('This front page is i.');
    expect(textOf(laid.pages[3]!)).toContain('This is page 3 of 5.');
  });

  it('keeps its own bookmarks out of the link destinations', async () => {
    const laid = layOut(await report());
    expect([...laid.pdf.bookmarks.keys()].every((name) => name.startsWith('_Toc'))).toBe(true);
    expect(laid.pdf.bookmarks.size).toBe(3);
  });

  it('lays a document out at most a few times', () => {
    expect(MAX_FIELD_PASSES).toBeGreaterThanOrEqual(2);
    expect(MAX_FIELD_PASSES).toBeLessThanOrEqual(5);
  });
});
