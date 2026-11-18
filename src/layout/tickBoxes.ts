// Tick boxes drawn as boxes. A Word check box (the block model's `checkbox`
// run, a `w14:checkbox` content control) prints its state as BALLOT BOX ☐
// (U+2610), BALLOT BOX WITH CHECK ☑ (U+2611) or BALLOT BOX WITH X ☒ (U+2612)
// in a symbol face. None of the bundled stand-ins carries those characters, so
// they came out as empty boxes with a loss. The layout is handed a small face
// built here from outlines — a square, a tick, a cross — through the engine's
// own TrueType builder, and the engine's per-character fallback draws the
// characters from it in the PDF and the page pictures alike. The characters
// stay in the text, so `pageText` reads a box's state as ☐ or ☒.

import type { GlyphSeg } from '../ream/core/font/ttf-build';

import { FontRegistry } from '../ream/core/font';
import { buildTrueType } from '../ream/core/font/ttf-build';

/** The characters the face draws. */
export const TICK_BOX_CHARACTERS: ReadonlyArray<number> = [0x2610, 0x2611, 0x2612];

/** The box's edges, in thousandths of an em: it stands on the baseline and is as tall as a capital. */
const LEFT = 90;
const RIGHT = 760;
const BOTTOM = -10;
const TOP = 660;
const STROKE = 56;
const ADVANCE = 850;

const rect = (x0: number, y0: number, x1: number, y1: number): Array<GlyphSeg> => [
  { op: 'move', x: x0, y: y0 },
  { op: 'line', x: x0, y: y1 },
  { op: 'line', x: x1, y: y1 },
  { op: 'line', x: x1, y: y0 },
  { op: 'close' },
];

/** A bar of width `w` from one point to another, as a four-sided contour. */
function bar(ax: number, ay: number, bx: number, by: number, w: number): Array<GlyphSeg> {
  const length = Math.hypot(bx - ax, by - ay);
  const nx = (-(by - ay) / length) * (w / 2);
  const ny = ((bx - ax) / length) * (w / 2);
  return [
    { op: 'move', x: ax + nx, y: ay + ny },
    { op: 'line', x: bx + nx, y: by + ny },
    { op: 'line', x: bx - nx, y: by - ny },
    { op: 'line', x: ax - nx, y: ay - ny },
    { op: 'close' },
  ];
}

/** The square, as four bars that do not overlap. */
const square = (): Array<GlyphSeg> => [
  ...rect(LEFT, BOTTOM, RIGHT, BOTTOM + STROKE),
  ...rect(LEFT, TOP - STROKE, RIGHT, TOP),
  ...rect(LEFT, BOTTOM + STROKE, LEFT + STROKE, TOP - STROKE),
  ...rect(RIGHT - STROKE, BOTTOM + STROKE, RIGHT, TOP - STROKE),
];

const INSET = STROKE + 70;

function cross(): Array<GlyphSeg> {
  const x0 = LEFT + INSET;
  const x1 = RIGHT - INSET;
  const y0 = BOTTOM + INSET;
  const y1 = TOP - INSET;
  return [...bar(x0, y0, x1, y1, STROKE * 1.3), ...bar(x0, y1, x1, y0, STROKE * 1.3)];
}

function tick(): Array<GlyphSeg> {
  const x0 = LEFT + INSET;
  const x1 = RIGHT - INSET;
  const y0 = BOTTOM + INSET;
  const y1 = TOP - INSET;
  const midX = x0 + (x1 - x0) * 0.38;
  return [...bar(x0, y0 + (y1 - y0) * 0.45, midX, y0, STROKE * 1.3), ...bar(midX, y0, x1, y1, STROKE * 1.3)];
}

let built: FontRegistry | undefined;

/** The face that draws ☐ ☑ ☒, built once. */
export function tickBoxRegistry(): FontRegistry {
  built ??= FontRegistry.fromBytes({
    regular: buildTrueType({
      family: 'document-kit Tick Boxes',
      bold: false,
      italic: false,
      postScriptName: 'DocumentKitTickBoxes',
      ascent: 900,
      descent: -250,
      capHeight: 660,
      xHeight: 480,
      fsType: 0,
      glyphs: [
        { codePoints: [0x2610], outline: square(), advance: ADVANCE },
        { codePoints: [0x2611], outline: [...square(), ...tick()], advance: ADVANCE },
        { codePoints: [0x2612], outline: [...square(), ...cross()], advance: ADVANCE },
      ],
    }),
  });
  return built;
}
