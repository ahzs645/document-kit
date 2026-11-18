// document-kit: ECMA-376 Part 1 §17.4.66 (`w:tcBorders`) — which border shows
// where two table cells share an edge and no cell spacing separates them.
//
// The standard's rules, in order:
//   1. A cell's own border is shown in preference to the table's.
//   2. A side that is nil or none gives way: the other border is shown.
//   3. The heavier border is shown: its width in eighths of a point times the
//      number of its style in the list below.
//   4. Equal weights: the style earlier in the list is shown.
//   5. Equal styles: the darker colour is shown — the lower R + B + 2G, then
//      the lower B + 2G, then the lower G.
// An explicit none on a cell is taken as rule 2 before rule 1: it never hides
// the border the neighbour draws.
//
// Also the order borders are PAINTED in (`borderPaintRank`): where two rules
// meet at a corner, each is stroked with a projecting cap, and the one painted
// last covers the other's end. Painting the winners of this order last keeps a
// heavy border whole where lighter rules run into it.

import type { Border, BorderStyle } from './document-model';

/** One side's claim on a shared edge: its border, and whether the cell itself states it. */
export interface EdgeClaim {
  readonly border: Border | undefined;
  /** The cell's own `w:tcBorders`, as against the table's `w:tblBorders` it inherits. */
  readonly fromCell: boolean;
}

/**
 * §17.4.66 — each style's number in the standard's list (ST_Border without
 * nil and none). The engine reads a subset; `dashDot` and `dashDotDot` are
 * the list's `dotDash` and `dotDotDash`.
 */
const BORDER_NUMBER: Readonly<Record<Exclude<BorderStyle, 'none'>, number>> = {
  single: 1,
  thick: 2,
  double: 3,
  dotted: 4,
  dashed: 5,
  dashDot: 6,
  dashDotDot: 7,
  dashSmallGap: 20,
};

const DEFAULT_WIDTH_PT = 0.5;

const isDrawn = (b: Border | undefined): b is Border => b !== undefined && b.style !== 'none';

/** §17.4.66 rule 3 — width in eighths of a point times the style's number. */
export function borderConflictWeight(b: Border): number {
  if (b.style === 'none') return 0;
  return Math.round((b.width ?? DEFAULT_WIDTH_PT) * 8) * BORDER_NUMBER[b.style];
}

/** §17.4.66 rule 5 — the three brightness measures, lower is darker. Unset colour is black. */
function brightness(b: Border): readonly [number, number, number] {
  const hex = /^[0-9a-f]{6}$/iu.test(b.colorHex ?? '') ? b.colorHex! : '000000';
  const r = Number.parseInt(hex.slice(0, 2), 16);
  const g = Number.parseInt(hex.slice(2, 4), 16);
  const bl = Number.parseInt(hex.slice(4, 6), 16);
  return [r + bl + 2 * g, bl + 2 * g, g];
}

/**
 * Compare two drawn borders by §17.4.66 rules 3–5.
 *
 * @returns Positive when `a` is shown over `b`, negative when `b` is, 0 when
 *          the standard does not tell them apart.
 */
export function compareBorders(a: Border, b: Border): number {
  const wa = borderConflictWeight(a);
  const wb = borderConflictWeight(b);
  if (wa !== wb) return wa - wb;
  const na = a.style === 'none' ? 0 : BORDER_NUMBER[a.style];
  const nb = b.style === 'none' ? 0 : BORDER_NUMBER[b.style];
  if (na !== nb) return nb - na;
  const [a1, a2, a3] = brightness(a);
  const [b1, b2, b3] = brightness(b);
  if (a1 !== b1) return b1 - a1;
  if (a2 !== b2) return b2 - a2;
  return b3 - a3;
}

/**
 * The border shown on an edge two cells share.
 *
 * @param a The claim of the cell that draws the edge (the lower or right one);
 *          it keeps the edge when the rules cannot tell the two apart.
 * @param b The facing claim of its neighbour.
 * @returns The border shown, or undefined when neither side draws one.
 */
export function edgeWinner(a: EdgeClaim, b: EdgeClaim): Border | undefined {
  const da = isDrawn(a.border);
  const db = isDrawn(b.border);
  if (!da && !db) return undefined;
  if (!db) return a.border;
  if (!da) return b.border;
  if (a.fromCell !== b.fromCell) return a.fromCell ? a.border : b.border;
  return compareBorders(a.border!, b.border!) >= 0 ? a.border : b.border;
}

/**
 * Where a laid-out rule falls in the paint order: lighter rules first, so the
 * border §17.4.66 shows over them is painted over their ends. Rules of the same
 * rank keep the order they were laid out in.
 */
export function borderPaintRank(b: {
  readonly borderSizePt: number;
  readonly borderStyle?: BorderStyle;
}): number {
  const style = b.borderStyle ?? 'single';
  return style === 'none' ? 0 : Math.round(b.borderSizePt * 8) * BORDER_NUMBER[style];
}
