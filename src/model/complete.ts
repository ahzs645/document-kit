/**
 * The list of things still to complete in a document: every value nobody gave (a null field, cell or signature line,
 * a blank rule), every sentence left out for a missing slot, every missing picture, and every item a block says it
 * is waiting on. Pure: an application prints the list where it likes (the end of a review draft, a checklist).
 */
import { expandDocument, type CustomRegistry } from './custom'
import { fillSlots, inlines, isRichText, plainText } from './inline'
import { tableGrid, toCell } from './table'
import { walkDocument, type BlockVisit } from './walk'
import type { Block, Document, RichText } from './types'

export type ToCompleteKind = 'blank' | 'field' | 'cell' | 'sentence' | 'figure' | 'signature' | 'pending'

export interface ToCompleteItem {
  kind: ToCompleteKind
  /** What is missing, in words: a field's label, a sentence's label or template, a figure's caption. */
  what: string
  /** For a sentence: the slots with no value. For a pending item: what it waits on. */
  missing?: readonly string[]
  /** The source tag the value would carry, when the model says. */
  source?: string
  section: number
  part: BlockVisit['part']
  /** The text of the nearest heading before it, in its section part. */
  heading?: string
}

export interface ToCompleteOptions {
  /**
   * Expand custom blocks through this registry first, so what they become is walked too. Without it a custom block
   * contributes only its own `pending` items.
   */
  registry?: CustomRegistry
}

export function toComplete(document: Document, options: ToCompleteOptions = {}): ToCompleteItem[] {
  const doc = options.registry ? expandDocument(document, options.registry) : document
  const items: ToCompleteItem[] = []
  let heading: string | undefined
  let place = ''
  for (const visit of walkDocument(doc)) {
    const where = `${visit.section}:${visit.part}`
    if (where !== place) {
      place = where
      heading = undefined
    }
    const { block } = visit
    const push = (item: Omit<ToCompleteItem, 'section' | 'part' | 'heading'>) => {
      const full: ToCompleteItem = { ...item, section: visit.section, part: visit.part }
      if (heading !== undefined) full.heading = heading
      items.push(full)
    }
    for (const pending of block.pending ?? []) {
      push({
        kind: 'pending',
        what: pending.what,
        ...(pending.missing ? { missing: pending.missing } : {}),
        ...(pending.source ? { source: pending.source } : {}),
      })
    }
    scanBlock(block, push)
    if (block.type === 'heading') heading = plainText(block.content)
  }
  return items
}

type Push = (item: Omit<ToCompleteItem, 'section' | 'part' | 'heading'>) => void

function scanText(text: RichText | null | undefined, push: Push, about?: string) {
  if (text === null || text === undefined) return
  for (const inline of inlines(text)) {
    if (typeof inline === 'string') continue
    if (inline.type === 'blank') {
      push({
        kind: 'blank',
        what: inline.what ?? about ?? 'Blank',
        ...(inline.source ? { source: inline.source } : {}),
      })
    } else if (inline.type === 'sentence') {
      const { missing } = fillSlots(inline.template, inline.values)
      if (missing.length) {
        push({
          kind: 'sentence',
          what: inline.label ?? inline.template,
          missing,
          ...(inline.source ? { source: inline.source } : {}),
        })
      }
    }
  }
}

const withSource = (source: string | undefined) => (source ? { source } : {})

function scanBlock(block: Block, push: Push) {
  switch (block.type) {
    case 'heading':
    case 'paragraph':
      scanText(block.content, push)
      return
    case 'list': {
      const visit = (items: typeof block.items) => {
        for (const item of items) {
          if (isRichText(item)) scanText(item, push)
          else {
            const listItem = item
            scanText(listItem.content, push)
            if (listItem.items) visit(listItem.items)
          }
        }
      }
      visit(block.items)
      return
    }
    case 'fields':
      scanText(block.title, push)
      for (const row of block.rows) {
        const label = plainText(row.label).replace(/:\s*$/, '')
        scanText(row.label, push)
        if (row.value === null) push({ kind: 'field', what: row.what ?? label, ...withSource(row.source) })
        else scanText(row.value, push, label)
      }
      return
    case 'table': {
      if (block.caption) scanText(block.caption.text, push)
      const title = block.caption ? plainText(block.caption.text) : ''
      const grid = gridOf(block)
      const names = columnNames(block, grid)
      block.rows.forEach((row, r) => {
        let column = 0
        const covered = coveredColumns(grid, r)
        row.cells.forEach((input, index) => {
          while (covered.has(column)) column += 1
          const cell = toCell(input)
          const at = column
          column += Math.max(1, Math.floor(cell.colSpan ?? 1))
          if (cell.blocks?.length) return // walked as blocks of their own
          if (cell.content === null) {
            const rowLabel = index === 0 ? `row ${r + 1}` : plainText(toCell(row.cells[0] ?? '').content ?? '') || `row ${r + 1}`
            const what = cell.what ?? [title, rowLabel, names[at] || `column ${at + 1}`].filter(Boolean).join(', ')
            push({ kind: 'cell', what, ...withSource(cell.source) })
          } else scanText(cell.content, push)
        })
      })
      scanText(block.note, push)
      return
    }
    case 'figure':
      if (block.caption) scanText(block.caption.text, push)
      if (block.image === null) {
        push({
          kind: 'figure',
          what: plainText(block.caption?.text) || plainText(block.placeholder) || block.alt || 'Figure',
          ...withSource(block.source),
        })
      }
      return
    case 'figureGrid':
      if (block.caption) scanText(block.caption.text, push)
      for (const item of block.items) {
        scanText(item.label, push)
        if (item.image === null) {
          push({
            kind: 'figure',
            what: plainText(item.label) || plainText(item.placeholder) || plainText(block.caption?.text) || 'Figure',
            ...withSource(block.source),
          })
        }
      }
      return
    case 'signature':
      scanText(block.label, push)
      if (block.name === null) push({ kind: 'signature', what: plainText(block.label) || 'Name of the signatory' })
      else scanText(block.name, push)
      for (const row of block.lines ?? []) {
        const label = plainText(row.label).replace(/:\s*$/, '')
        if (row.value === null) push({ kind: 'signature', what: row.what ?? label, ...withSource(row.source) })
        else scanText(row.value, push, label)
      }
      return
    case 'toc':
      scanText(block.title, push)
      return
    case 'pageBreak':
    case 'spacer':
    case 'custom':
      return
  }
}

type Grid = ReturnType<typeof tableGrid> | null

/** The table's grid, or null for a table whose spans cannot be placed (it is then named by cell position alone). */
function gridOf(table: Extract<Block, { type: 'table' }>): Grid {
  try {
    return tableGrid(table)
  } catch {
    return null
  }
}

/** Each grid column's name: the text of the lowest header row's cell over it. */
function columnNames(table: Extract<Block, { type: 'table' }>, grid: Grid): string[] {
  const names: string[] = []
  if (!grid) return names
  table.rows.forEach((row, r) => {
    if (row.role !== 'header') return
    for (const entry of grid.rows[r] ?? []) {
      if (entry.kind !== 'cell') continue
      const text = plainText(entry.placed.cell.content ?? '')
      if (!text) continue
      for (let c = entry.placed.column; c < entry.placed.column + entry.placed.colSpan; c += 1) names[c] = text
    }
  })
  return names
}

/** The grid columns of row `r` that a cell from a row above covers. */
function coveredColumns(grid: Grid, r: number): Set<number> {
  const out = new Set<number>()
  for (const entry of grid?.rows[r] ?? []) {
    if (entry.kind !== 'covered') continue
    for (let c = entry.by.column; c < entry.by.column + entry.by.colSpan; c += 1) out.add(c)
  }
  return out
}
