/**
 * Hand edits of a rendered part's printed values. A part fills its blocks from the data; an edit then replaces one
 * printed value (a field's value, a signature's name, a body cell, a paragraph, a list item) with a person's own words,
 * marked with the edit source. An edit is keyed by the part instance and its place in the part — the heading it sits
 * under (or the block's own `id`), its kind and its label — never by the recipe, so two recipes sharing a part share
 * its edits, and never by a loop's index, so removing a row does not move an edit onto the next one. Pure.
 */
import { plainText, resolveSentences, toCell } from '../model'
import type { Block, CellInput, FieldRow, ListItem, RichText, TableBlock } from '../model'

/** Edits by key. An empty string blanks the value (printed blank and listed to complete). */
export type RecipeEdits = Readonly<Record<string, string>>

export type EditKind = 'field' | 'name' | 'cell' | 'paragraph' | 'item'

/** One printed value a person can change by hand. */
export interface EditTarget {
  key: string
  /** The part instance's id and title. */
  part: string
  partTitle: string
  kind: EditKind
  /** Where it is, in words: the heading or block id, then its label. */
  label: string
  /** What the part printed, before any edit; null for a blank. */
  original: string | null
  /** The edit applied, when there is one. */
  edited?: string
}

const SEP = ' › '

/** A value's printed text: sentences resolved, null when it prints nothing. */
function printed(text: RichText | null | undefined): string | null {
  if (text === null || text === undefined) return null
  const value = plainText(resolveSentences(text)).trim()
  return value === '' ? null : value
}

const isListItem = (item: ListItem | RichText): item is ListItem =>
  typeof item === 'object' && item !== null && !Array.isArray(item) && !('type' in item)

/** The part's blocks with the edits applied, and every editable value. With no edit for the part, the same objects. */
export function applyEdits(
  part: { id: string; title: string },
  blocks: readonly Block[],
  edits: RecipeEdits,
  source: string,
): { blocks: Block[]; targets: EditTarget[] } {
  const targets: EditTarget[] = []
  const seen = new Map<string, number>()
  const leaf = (scope: string, kind: EditKind, label: string, original: string | null) => {
    const base = `${part.id}|${scope}|${kind}|${label}`
    const n = (seen.get(base) ?? 0) + 1
    seen.set(base, n)
    const key = n > 1 ? `${base}#${n}` : base
    const has = Object.prototype.hasOwnProperty.call(edits, key) && typeof edits[key] === 'string'
    targets.push({ key, part: part.id, partTitle: part.title, kind, label: [scope, label].filter(Boolean).join(SEP), original, ...(has ? { edited: edits[key] } : {}) })
    if (!has) return null
    const text = edits[key] as string
    return { value: text === '' ? null : ({ type: 'text', text, source } as const) }
  }
  const editRow = (row: FieldRow, scope: string, prefix: string, kind: EditKind = 'field'): FieldRow => {
    const label = [prefix, plainText(row.label).replace(/:\s*$/, '')].filter(Boolean).join(SEP)
    const out = leaf(scope, kind, label, printed(row.value))
    return out ? { ...row, value: out.value, source } : row
  }
  const walk = (list: readonly Block[], scopeIn: string): Block[] => {
    let scope = scopeIn
    return list.map((block): Block => {
      const own = block.id ? `#${block.id}` : scope
      switch (block.type) {
        case 'heading':
          scope = printed(block.content) ?? scope
          return block
        case 'fields': {
          const title = printed(block.title) ?? ''
          const rows = block.rows.map((row) => editRow(row, own, title))
          return rows.some((row, i) => row !== block.rows[i]) ? { ...block, rows } : block
        }
        case 'signature': {
          const label = printed(block.label) ?? ''
          const name = leaf(own, 'name', label || 'Name', block.name === undefined ? null : printed(block.name))
          const lines = block.lines?.map((row) => editRow(row, own, label))
          const changed = name || lines?.some((row, i) => row !== block.lines?.[i])
          return changed ? { ...block, ...(name ? { name: name.value } : {}), ...(lines ? { lines } : {}) } : block
        }
        case 'paragraph': {
          const out = leaf(own, 'paragraph', '', printed(block.content))
          return out ? { ...block, content: out.value ? [out.value] : '', source } : block
        }
        case 'list': {
          let changed = false
          const items = (entries: readonly (ListItem | RichText)[]): (ListItem | RichText)[] =>
            entries.map((entry) => {
              const item = isListItem(entry) ? entry : null
              const content = item ? item.content : (entry as RichText)
              const out = leaf(own, 'item', '', printed(content))
              const nested = item?.items ? items(item.items) : undefined
              if (!out && !(nested && nested.some((n, i) => n !== item?.items?.[i]))) return entry
              changed = true
              const next = out ? (out.value ?? '') : content
              return item ? { ...item, content: next, ...(nested ? { items: nested } : {}) } : next
            })
          const next = items(block.items)
          return changed ? { ...block, items: next } : block
        }
        case 'table':
          return editTable(block, own)
        default:
          return block
      }
    })
  }
  const editTable = (table: TableBlock, scope: string): TableBlock => {
    const title = printed(table.caption?.text) ?? ''
    const names: string[] = []
    for (const row of table.rows) {
      if (row.role !== 'header') continue
      let column = 0
      for (const input of row.cells) {
        const cell = toCell(input)
        const span = Math.max(1, Math.floor(cell.colSpan ?? 1))
        const text = printed(cell.content)
        if (text) for (let c = column; c < column + span; c += 1) names[c] = text
        column += span
      }
    }
    let changed = false
    const rows = table.rows.map((row, r) => {
      if (row.role !== undefined && row.role !== 'body') return row
      let column = 0
      const first = printed(toCell(row.cells[0] ?? null).content)
      let rowChanged = false
      const cells = row.cells.map((input): CellInput => {
        const cell = toCell(input)
        const at = column
        column += Math.max(1, Math.floor(cell.colSpan ?? 1))
        if (cell.blocks?.length) {
          const inner = walk(cell.blocks, scope)
          if (inner.every((b, i) => b === cell.blocks?.[i])) return input
          rowChanged = true
          return { ...cell, blocks: inner }
        }
        // A row is named by its first cell as printed (never its place), so a removed row moves no edit onto the next.
        const rowLabel = first ?? `row ${r + 1}`
        const label = [title, rowLabel, names[at] ?? `column ${at + 1}`].filter(Boolean).join(SEP)
        const out = leaf(scope, 'cell', label, printed(cell.content))
        if (!out) return input
        rowChanged = true
        return { ...cell, content: out.value, source }
      })
      if (!rowChanged) return row
      changed = true
      return { ...row, cells }
    })
    return changed ? { ...table, rows } : table
  }
  const out = walk(blocks, '')
  return { blocks: out, targets }
}
