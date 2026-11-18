// Lay a Word file out with the engine directly, Arimo for every face, for tests
// that look at where things land on the page rather than at a writer's output.

import type { LaidOutPage } from '../../src/ream/layout/page-doc';
import type { LaidOutPdfDocument } from '../../src/ream/layout/styled-layout';
import type { Loss } from '../../src/ream/core/ir';

import { flowRenderOptions } from '../../src/ream/core/converter/project';
import { FontRegistry } from '../../src/ream/core/font';
import { layoutStyledDocument } from '../../src/ream/layout/styled-layout';
import { docxReader } from '../../src/ream/word/docx-reader';
import { fontFile } from './sample';

let arimo: FontRegistry | undefined;

/** Arimo's four faces as the engine's registry (read once). */
function registry(): FontRegistry {
  arimo ??= FontRegistry.fromBytes({
    regular: fontFile('Arimo-Regular.ttf'),
    bold: fontFile('Arimo-Bold.ttf'),
    italic: fontFile('Arimo-Italic.ttf'),
    boldItalic: fontFile('Arimo-BoldItalic.ttf'),
  });
  return arimo;
}

/** The engine's layout of a Word file, and the losses it recorded. */
export function layOut(docx: Uint8Array): LaidOutPdfDocument & { readonly losses: Array<Loss> } {
  const { doc } = docxReader.read(docx);
  const losses: Array<Loss> = [];
  const laid = layoutStyledDocument(doc.body, {
    ...flowRenderOptions(doc),
    registry: registry(),
    losses,
  });
  return { ...laid, losses };
}

/** The text of each line on a page, with its baseline (pt from the page's top). */
export function linesOf(page: LaidOutPage): Array<{ text: string; baselineY: number; x: number }> {
  const out: Array<{ text: string; baselineY: number; x: number }> = [];
  for (const item of page.commands) {
    if (item.type !== 'line') continue;
    let text = '';
    for (const token of item.line.tokens) if (token.kind === 'text') text += token.text;
    out.push({ text, baselineY: item.baselineY, x: item.originX });
  }
  return out;
}

/** The text of a page's lines joined by `\n`. */
export const textOf = (page: LaidOutPage): string =>
  linesOf(page)
    .map((l) => l.text)
    .join('\n');
