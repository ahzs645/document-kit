/** Visiting every block in reading order, into table cells. */
import { toCell } from './table'
import type { Block, Document } from './types'

export interface BlockVisit {
  block: Block
  /** Index of the section the block is in. */
  section: number
  /** Where it is: `body`, `header`, `footer`, `firstHeader`, `firstFooter`. */
  part: 'body' | 'header' | 'footer' | 'firstHeader' | 'firstFooter'
  /** The blocks it sits inside (a table, pinned block or frame holding it), outermost first. */
  parents: readonly Block[]
}

/** Every block of a list in reading order, a table's cell blocks (and a pinned block's blocks) after the block itself. */
export function* walkBlocks(
  blocks: readonly Block[],
  context: Omit<BlockVisit, 'block'>,
): Generator<BlockVisit> {
  for (const block of blocks) {
    yield { ...context, block }
    if (block.type === 'pinned' || block.type === 'frame') {
      yield* walkBlocks(block.blocks, { ...context, parents: [...context.parents, block] })
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const input of row.cells) {
          const cell = toCell(input)
          if (cell.blocks?.length) {
            yield* walkBlocks(cell.blocks, { ...context, parents: [...context.parents, block] })
          }
        }
      }
    }
  }
}

/** Every block in the document: each section's body first, then its headers and footers. */
export function* walkDocument(document: Document, options: { headers?: boolean } = {}): Generator<BlockVisit> {
  for (const [section, sec] of document.sections.entries()) {
    yield* walkBlocks(sec.body, { section, part: 'body', parents: [] })
    if (options.headers === false) continue
    if (sec.header) yield* walkBlocks(sec.header, { section, part: 'header', parents: [] })
    if (sec.footer) yield* walkBlocks(sec.footer, { section, part: 'footer', parents: [] })
    if (sec.firstPage?.header) yield* walkBlocks(sec.firstPage.header, { section, part: 'firstHeader', parents: [] })
    if (sec.firstPage?.footer) yield* walkBlocks(sec.firstPage.footer, { section, part: 'firstFooter', parents: [] })
  }
}
