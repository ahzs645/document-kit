/**
 * Caption numbering: tables and figures numbered in reading order, each kind on its own count ("Table 1", "Table 2",
 * "Figure 1"), across every section's body. Pure, so an application can ask what number a table will print with.
 */
import { walkDocument } from './walk'
import type { Block, Caption, Document, Theme } from './types'

export type CaptionKind = 'table' | 'figure'

export type CaptionedBlock = Extract<Block, { type: 'table' | 'spreadsheet' | 'figure' | 'figureGrid' }>

/**
 * Which count a block's caption takes: a table's (tables and spreadsheets share one), or a figure's (figures and figure
 * grids share one).
 */
export function captionKind(block: Block): CaptionKind | null {
  if (block.type === 'table' || block.type === 'spreadsheet') return 'table'
  if (block.type === 'figure' || block.type === 'figureGrid') return 'figure'
  return null
}

/** True when the caption takes a number of its kind (printed or not). */
export const takesNumber = (caption: Caption | undefined): boolean =>
  caption !== undefined && caption.numbered !== false

/**
 * Each captioned table and figure's number, by block, in reading order. Custom blocks are not looked into: number the
 * expanded document ({@link expandDocument}) to count what they become, as the writers do.
 */
export function numberCaptions(source: Document | readonly Block[]): Map<CaptionedBlock, number> {
  const numbers = new Map<CaptionedBlock, number>()
  const counts: Record<CaptionKind, number> = { table: 0, figure: 0 }
  const document: Document = Array.isArray(source) ? { sections: [{ body: source as readonly Block[] }] } : (source as Document)
  for (const { block } of walkDocument(document, { headers: false })) {
    const kind = captionKind(block)
    if (!kind) continue
    const captioned = block as CaptionedBlock
    if (!takesNumber(captioned.caption)) continue
    counts[kind] += 1
    numbers.set(captioned, counts[kind])
  }
  return numbers
}

/** "Table 3", in the theme's words. */
export function captionLabel(kind: CaptionKind, number: number, theme: Pick<Theme, 'captions'>): string {
  return `${kind === 'table' ? theme.captions.table : theme.captions.figure} ${number}`
}

/**
 * What a caption prints before its own text: `"Table 3: "` for a numbered caption, nothing for `numbered: false` or
 * `'silent'`. A caption with no text of its own prints the label alone.
 */
export function captionPrefix(
  kind: CaptionKind,
  caption: Caption,
  number: number | undefined,
  theme: Pick<Theme, 'captions'>,
  hasText = true,
): string {
  if (caption.numbered === false || caption.numbered === 'silent' || number === undefined) return ''
  return captionLabel(kind, number, theme) + (hasText ? theme.captions.separator : '')
}
