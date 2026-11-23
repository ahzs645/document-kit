// What the layout makes of the report features the Word writer writes: a block
// pinned to the foot of the page, tick boxes drawn as boxes, two short fields
// on one line. Made-up text throughout.

import { describe, expect, it } from 'vitest';

import {
  checkbox,
  choices,
  document,
  field,
  fields,
  paragraph,
  pinned,
  section,
  table,
  row,
  cell,
  type Block,
} from '../../src/model';
import { writeDocx } from '../../src/docx';
import { createFontProvider, docxToPages, layoutDocx } from '../../src/layout';
import { layOut, linesOf } from './engine';
import { diskLoader } from './sample';
import { unzipSync } from 'fflate';
import { PdfFile } from '../../src/ream/pdf-reader/document';
import { extractPageText } from '../../src/ream/pdf-reader/text';

const fonts = createFontProvider(diskLoader);
// Letter, one-inch margins: the text area ends 720pt down the page.
const BOTTOM = 792 - 72;
const page = { paper: 'letter' as const, margins: { top: 72, bottom: 72, left: 72, right: 72 } };

const filler = (count: number, word = 'Body'): Block[] =>
  Array.from({ length: count }, (_, i) => paragraph(`${word} line ${i + 1} of a made-up report.`));

describe('a pinned block', () => {
  it('sits at the foot of the page it falls on, the flow carrying on above it', async () => {
    const docx = await writeDocx(
      document([
        section(
          [
            paragraph('Opening line.'),
            pinned([paragraph('Made-up disclaimer, pinned to the foot.'), paragraph('Its second line.')]),
            paragraph('After the pin.'),
          ],
          { page },
        ),
      ]),
    );
    const lines = linesOf(layOut(docx).pages[0]!);
    const at = (text: string) => lines.find((l) => l.text.includes(text))!;
    expect(at('Its second line.').baselineY).toBeGreaterThan(BOTTOM - 20);
    expect(at('Its second line.').baselineY).toBeLessThanOrEqual(BOTTOM);
    expect(at('Made-up disclaimer').baselineY).toBeLessThan(at('Its second line.').baselineY);
    // The flow does not wait for it: the next paragraph follows the opening line.
    expect(at('After the pin.').baselineY).toBeLessThan(200);
  });

  it('pins a block on the last page when it closes the document, its room kept under it', async () => {
    const docx = await writeDocx(
      document([
        section(
          [...filler(40), pinned([paragraph('Made-up contact block.')], { spaceBelow: 24 })],
          { page },
        ),
      ]),
    );
    const laid = layOut(docx);
    expect(laid.pages.length).toBe(2);
    const last = linesOf(laid.pages[1]!);
    const contact = last.find((l) => l.text.includes('Made-up contact block.'))!;
    // 24pt kept under it inside the frame lifts it clear of the margin.
    expect(contact.baselineY).toBeGreaterThan(BOTTOM - 50);
    expect(contact.baselineY).toBeLessThan(BOTTOM - 20);
    expect(linesOf(laid.pages[0]!).some((l) => l.text.includes('contact'))).toBe(false);
  });

  it('pushes text that would run into it on to the next page', async () => {
    // Enough body text to reach the foot of page 1 with the pin anchored at its top.
    const docx = await writeDocx(
      document([
        section([pinned([paragraph('Pinned A.'), paragraph('Pinned B.'), paragraph('Pinned C.')]), ...filler(52)], { page }),
      ]),
    );
    const laid = layOut(docx);
    const first = linesOf(laid.pages[0]!);
    const pin = first.find((l) => l.text.includes('Pinned A.'))!;
    const body = first.filter((l) => l.text.startsWith('Body'));
    // No body line overlaps the pinned box: every one ends above it.
    for (const line of body) expect(line.baselineY).toBeLessThan(pin.baselineY - 8);
    expect(laid.pages.length).toBe(2);
  });

  it('goes to the foot of the next page when the text above it leaves no room for it', async () => {
    // 29 lines fill page 1 to about 680pt; a three-line pin needs the 66pt above the margin.
    const docx = await writeDocx(
      document([section([...filler(29), pinned([paragraph('Pinned A.'), paragraph('Pinned B.'), paragraph('Pinned C.')])], { page })]),
    );
    const laid = layOut(docx);
    expect(laid.pages.length).toBe(2);
    expect(linesOf(laid.pages[0]!).some((l) => l.text.startsWith('Pinned'))).toBe(false);
    const second = linesOf(laid.pages[1]!);
    expect(second.map((l) => l.text)).toEqual(['Pinned A.', 'Pinned B.', 'Pinned C.'])
    expect(second[2]!.baselineY).toBeGreaterThan(BOTTOM - 20);
  });

  it('refuses a table inside a pin', async () => {
    await expect(
      writeDocx(document([section([pinned([{ type: 'table', rows: [{ cells: ['x'] }] }])])])),
    ).rejects.toThrow(/pinned block cannot hold a table/);
  });
});

describe('tick boxes', () => {
  it('are drawn as boxes in the PDF and the page pictures, their state read as ☐ ☒', async () => {
    const docx = await writeDocx(
      document([
        section([
          paragraph(['Agreed ', checkbox(true), ' Not agreed ', checkbox(false)]),
          choices(['Low', 'Moderate', 'High'], 'Moderate', { label: 'Made-up rating:' }),
        ]),
      ]),
    );
    const pages = await docxToPages(docx, { fonts });
    expect(pages.pageText[0]).toContain('Agreed ☒ Not agreed ☐');
    expect(pages.pageText[0]).toMatch(/Made-up rating:\s+☐\s?Low\s+☒\s?Moderate\s+☐\s?High/);
    // No box is lost: the face of boxes draws them.
    expect(pages.losses.filter((l) => /2610|2612|symbol/.test(l.detail))).toEqual([]);
    const laid = await layoutDocx(docx, { fonts });
    const pdf = laid.toPdf();
    expect(new TextDecoder('latin1').decode(pdf.bytes)).toContain('Tick');
    // The picture draws a glyph for every box: three more glyph uses than letters alone would make.
    expect((pages.pages[0]!.match(/<use /g) ?? []).length).toBeGreaterThan(40);
  });
});

describe('circled numbers', () => {
  it('prints and embeds the original characters in a bold table header and plain text', async () => {
    const docx = await writeDocx(document([
      section([
        table([row([
          cell({ type: 'text', text: 'Run ①', bold: true }),
          cell({ type: 'text', text: 'Run ②', bold: true }),
        ], { role: 'header' })]),
        paragraph('⓪ through ⑳'),
      ], { page }),
    ]));
    const laid = await layoutDocx(docx, { fonts });
    const pdf = laid.toPdf();
    const file = PdfFile.parse(pdf.bytes);
    const text = extractPageText(file, file.pages()[0]!).map((run) => run.text).join(' ');
    for (const marker of ['⓪', '①', '②', '⑳']) {
      expect(text).toContain(marker);
      expect(laid.toPages().pageText[0]).toContain(marker);
    }
    expect(pdf.losses.filter((loss) => loss.severity === 'dropped' || loss.severity === 'degraded')).toEqual([]);
    expect(laid.toPages().losses.filter((loss) => loss.severity === 'dropped' || loss.severity === 'degraded')).toEqual([]);
  });
});

describe('two fields on one line', () => {
  const contact = (email: string) =>
    document([
      section(
        [
          fields([
            field('E-mail', email),
            field('Phone', '555-010-0199', { sameLine: true }),
            field('Date', '12 March'),
          ]),
        ],
        { page },
      ),
    ]);

  it('shares a line when both fit', async () => {
    const pages = await docxToPages(await writeDocx(contact('someone@example.org')), { fonts });
    const lines = pages.pageText[0]!.split('\n');
    expect(lines[0]).toMatch(/^E-mail: someone@example\.org\s+Phone:\s555-010-0199$/);
    expect(lines[1]).toBe('Date: 12 March');
  });

  it('reads and prints the unbreakable hyphens as the hyphen a person types', async () => {
    const docx = await writeDocx(contact('someone@example.org'))
    // Word gets its own no-break hyphen, so the number moves whole there too.
    expect(new TextDecoder().decode(unzipSync(docx)['word/document.xml']!)).toContain('<w:noBreakHyphen/>');
    const laid = await layoutDocx(docx, { fonts });
    expect(laid.pageText[0]).not.toContain('\u2011');
    expect(laid.pageLines[0]![0]!.text).toMatch(/555-010-0199$/);
    // The PDF's text layer gives U+002D back, not U+2011, and the hyphen is Tinos's own (one face draws the line).
    const file = PdfFile.parse(laid.toPdf().bytes);
    const text = extractPageText(file, file.pages()[0]!).map((run) => run.text).join('');
    expect(text).toContain('555-010-0199');
    expect(text).not.toContain('\u2011');
    const pages = laid.toPages();
    expect(pages.losses.filter((loss) => /U\+2011/.test(loss.detail))).toEqual([]);
  });

  it('gives each line its place on the page', async () => {
    const laid = await layoutDocx(await writeDocx(contact('someone@example.org')), { fonts });
    const [first, second] = laid.pageLines[0]!;
    expect(first!.x).toBeCloseTo(72, 0);
    expect(second!.text).toBe('Date: 12 March');
    expect(second!.y).toBeGreaterThan(first!.y);
    expect(laid.toPages().pageLines).toEqual(laid.pageLines);
  });

  it('moves the second field whole to a line of its own when it does not fit', async () => {
    const long = Array.from({ length: 4 }, (_, i) => `person${i}@a-long-made-up-domain.example.org`).join(', ');
    const pages = await docxToPages(await writeDocx(contact(long)), { fonts });
    const lines = pages.pageText[0]!.split('\n');
    // The addresses wrap at their own spaces; the phone is never split inside the number or from its label.
    expect(lines.length).toBeGreaterThan(2);
    const phone = lines.find((l) => l.includes('Phone'))!;
    expect(phone).toMatch(/Phone:\s555-010-0199$/);
    expect(lines.some((l) => /Phone:$|555-$|^010/.test(l.trim()))).toBe(false);
  });
});
