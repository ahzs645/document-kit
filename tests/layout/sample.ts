// Made-up Word files for the layout tests, written with the `docx` library the
// way the applications write theirs.

import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import {
  Document,
  Footer,
  HeadingLevel,
  Packer,
  PageNumber,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
} from 'docx';

import type { StandInFile } from '../../src/layout';

const FONTS = fileURLToPath(new URL('../../fonts/', import.meta.url));

/** The bytes of a file in `fonts/`. */
export function fontFile(file: string): Uint8Array {
  return new Uint8Array(readFileSync(resolve(FONTS, file)));
}

/** A loader that reads every bundled font from disk. */
export const diskLoader = (f: StandInFile): Uint8Array => fontFile(f.file);

export interface SampleOptions {
  /** Rows in the table (enough to break across pages by default). */
  readonly rows?: number;
  /** The font every run names (none by default). */
  readonly font?: string;
  /** Extra paragraphs, each its own run. */
  readonly extra?: ReadonlyArray<{ readonly text: string; readonly font?: string }>;
}

/** A report: heading, paragraph, a long table with a repeated header row, a page-numbered footer. */
export async function sampleDocx(options: SampleOptions = {}): Promise<Uint8Array> {
  const rows = options.rows ?? 70;
  const font = options.font;
  const run = (text: string, bold = false): TextRun =>
    new TextRun({ text, bold, ...(font ? { font } : {}) });
  const cell = (text: string, bold = false): TableCell =>
    new TableCell({ children: [new Paragraph({ children: [run(text, bold)] })] });
  const header = new TableRow({
    tableHeader: true,
    children: [cell('Sample', true), cell('Measured value', true), cell('Note', true)],
  });
  const body = Array.from(
    { length: rows },
    (_, i) =>
      new TableRow({
        children: [
          cell(`S-${String(i + 1).padStart(3, '0')}`),
          cell(`${(i * 1.7).toFixed(1)} units`),
          cell('Within range'),
        ],
      }),
  );
  const doc = new Document({
    sections: [
      {
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    children: ['Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES],
                    ...(font ? { font } : {}),
                  }),
                ],
              }),
            ],
          }),
        },
        children: [
          new Paragraph({ heading: HeadingLevel.HEADING_1, children: [run('Example results')] }),
          new Paragraph({ children: [run('A made-up table of results, long enough to need two pages.')] }),
          new Table({ rows: [header, ...body] }),
          ...(options.extra ?? []).map(
            (e) =>
              new Paragraph({
                children: [new TextRun({ text: e.text, ...(e.font ? { font: e.font } : {}) })],
              }),
          ),
        ],
      },
    ],
  });
  return new Uint8Array(await Packer.toArrayBuffer(doc));
}
