// Reading a PDF the layout wrote back: its objects (streams inflated) and each page's drawing as text.

import { unzlibSync } from 'fflate';

const latin1 = new TextDecoder('latin1');

export interface PdfObject {
  readonly num: number;
  readonly dict: string;
  readonly data: Uint8Array | null;
}

/** Each object of a PDF, its stream inflated when deflated. */
export function pdfObjects(pdf: Uint8Array): PdfObject[] {
  const text = latin1.decode(pdf);
  const out: PdfObject[] = [];
  for (const m of text.matchAll(/(?:^|\n)(\d+) 0 obj\s*(<<[\s\S]*?>>)\s*(stream\r?\n|endobj)/g)) {
    const num = Number(m[1]);
    const dict = m[2]!;
    if (!m[3]!.startsWith('stream')) {
      out.push({ num, dict, data: null });
      continue;
    }
    const from = m.index! + m[0].length;
    const length = Number(/\/Length (\d+)/.exec(dict)![1]);
    const raw = pdf.subarray(from, from + length);
    out.push({ num, dict, data: /\/FlateDecode/.test(dict) ? unzlibSync(raw) : raw });
  }
  return out;
}

/** Each page's drawing (its content stream) as text, in page order. */
export function pageDrawings(pdf: Uint8Array): string[] {
  const all = pdfObjects(pdf);
  return all
    .filter((o) => /\/Type \/Page\b/.test(o.dict))
    .map((o) => Number(/\/Contents (\d+) 0 R/.exec(o.dict)![1]))
    .map((n) => latin1.decode(all.find((o) => o.num === n)!.data!));
}

/**
 * Where each picture is painted on a page: the `cm` matrix before each `Do`, as `{ name, x, y, width, height }` in
 * PDF points (y up from the foot of the page), in paint order. Matrices of the form `w 0 0 h x y cm`.
 */
export function paintedImages(drawing: string): Array<{ name: string; x: number; y: number; width: number; height: number; at: number }> {
  const out: Array<{ name: string; x: number; y: number; width: number; height: number; at: number }> = [];
  for (const m of drawing.matchAll(/([-\d.]+) 0 0 ([-\d.]+) ([-\d.]+) ([-\d.]+) cm\s*\/(\w+) Do/g)) {
    out.push({ name: m[5]!, width: Number(m[1]), height: Number(m[2]), x: Number(m[3]), y: Number(m[4]), at: m.index! });
  }
  return out;
}
