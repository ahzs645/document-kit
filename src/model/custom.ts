/**
 * Custom blocks: the blocks an application has and the kit does not know. The application registers, per kind, a
 * function that expands the block's data into the kit's own blocks; the writers expand before they write. A kind with
 * no expander is an error naming the kind, never a block silently left out.
 */
import { pageBox, type PageBox } from './page'
import { resolveTheme } from './theme'
import { toCell } from './table'
import type { Block, CellInput, CustomBlock, Document, Section, TableBlock, Theme } from './types'

export interface ExpandContext {
  theme: Theme
  /** The page of the section the block is in: its text width is what a wide block has to fit. */
  page: PageBox
  /** Index of that section. */
  section: number
}

export type CustomExpander<Data = unknown> = (
  data: Data,
  context: ExpandContext,
  block: CustomBlock<Data>,
) => readonly Block[]

export interface CustomRegistry {
  /** The expander for a kind, or undefined. */
  get(kind: string): CustomExpander | undefined
  has(kind: string): boolean
  readonly kinds: readonly string[]
  /** A new registry with one more kind (or one replaced). */
  with<Data>(kind: string, expander: CustomExpander<Data>): CustomRegistry
}

export class UnknownCustomBlockError extends Error {
  readonly kind: string
  constructor(kind: string, known: readonly string[]) {
    super(
      `document-kit: no expander is registered for custom block kind "${kind}"` +
        (known.length ? ` (registered: ${known.join(', ')})` : ' (no kinds are registered)'),
    )
    this.name = 'UnknownCustomBlockError'
    this.kind = kind
  }
}

/** A registry of custom block kinds. */
export function defineCustomBlocks<Kinds extends Record<string, unknown> = Record<string, unknown>>(
  expanders: { [K in keyof Kinds]: CustomExpander<Kinds[K]> } = {} as { [K in keyof Kinds]: CustomExpander<Kinds[K]> },
): CustomRegistry {
  return makeRegistry(new Map(Object.entries(expanders) as Array<[string, CustomExpander]>))
}

function makeRegistry(map: ReadonlyMap<string, CustomExpander>): CustomRegistry {
  return {
    get: (kind) => map.get(kind),
    has: (kind) => map.has(kind),
    kinds: [...map.keys()],
    with: (kind, expander) => {
      if (!kind) throw new Error('document-kit: a custom block kind needs a name')
      const next = new Map(map)
      next.set(kind, expander as CustomExpander)
      return makeRegistry(next)
    },
  }
}

/** How deep custom blocks may expand into further custom blocks before the expansion is taken to be a loop. */
export const MAX_CUSTOM_DEPTH = 16

/**
 * The blocks with every custom block expanded, into table cells too. An expansion may itself hold custom blocks; they
 * are expanded in turn. Throws {@link UnknownCustomBlockError} for a kind the registry does not have.
 */
export function expandBlocks(
  blocks: readonly Block[],
  registry: CustomRegistry | undefined,
  context: ExpandContext,
  depth = 0,
): Block[] {
  if (depth > MAX_CUSTOM_DEPTH) {
    throw new Error(`document-kit: custom blocks expanded more than ${MAX_CUSTOM_DEPTH} levels deep; is one a loop?`)
  }
  const out: Block[] = []
  for (const block of blocks) {
    if (block.type === 'custom') {
      const expander = registry?.get(block.kind)
      if (!expander) throw new UnknownCustomBlockError(block.kind, registry?.kinds ?? [])
      const expanded = expandBlocks(expander(block.data, context, block), registry, context, depth + 1)
      // The custom block's own page break and pending items carry onto what it became.
      if (block.pageBreakBefore && expanded[0]) expanded[0] = { ...expanded[0], pageBreakBefore: true }
      if (block.pending?.length) {
        if (expanded[0]) {
          const first = expanded[0]
          expanded[0] = { ...first, pending: [...(first.pending ?? []), ...block.pending] }
        } else {
          out.push({ type: 'spacer', height: 0, pending: block.pending })
        }
      }
      out.push(...expanded)
    } else if (block.type === 'table') {
      out.push(expandTable(block, registry, context, depth))
    } else if (block.type === 'pinned' && block.blocks.some(holdsCustom)) {
      out.push({ ...block, blocks: expandBlocks(block.blocks, registry, context, depth) })
    } else {
      out.push(block)
    }
  }
  return out
}

function expandTable(table: TableBlock, registry: CustomRegistry | undefined, context: ExpandContext, depth: number) {
  let changed = false
  const rows = table.rows.map((row) => {
    let rowChanged = false
    const cells = row.cells.map((input): CellInput => {
      const cell = toCell(input)
      if (!cell.blocks?.some(holdsCustom)) return input
      rowChanged = true
      return { ...cell, blocks: expandBlocks(cell.blocks, registry, context, depth) }
    })
    if (!rowChanged) return row
    changed = true
    return { ...row, cells }
  })
  return changed ? { ...table, rows } : table
}

function holdsCustom(block: Block): boolean {
  if (block.type === 'custom') return true
  if (block.type === 'pinned') return block.blocks.some(holdsCustom)
  if (block.type !== 'table') return false
  return block.rows.some((row) => row.cells.some((input) => toCell(input).blocks?.some(holdsCustom) ?? false))
}

/** True when a document holds a custom block anywhere. */
export function hasCustomBlocks(document: Document): boolean {
  return document.sections.some((section) => sectionParts(section).some((blocks) => blocks.some(holdsCustom)))
}

function sectionParts(section: Section): Array<readonly Block[]> {
  return [
    section.body,
    section.header ?? [],
    section.footer ?? [],
    section.firstPage?.header ?? [],
    section.firstPage?.footer ?? [],
  ]
}

/**
 * The document with every custom block expanded, in its body, headers and footers. A document without custom blocks
 * comes back as it is (the same object).
 */
export function expandDocument(document: Document, registry?: CustomRegistry): Document {
  if (!hasCustomBlocks(document)) return document
  const theme = resolveTheme(document.theme)
  const sections = document.sections.map((section, index): Section => {
    const context: ExpandContext = { theme, page: pageBox(section.page, theme), section: index }
    const expand = (blocks: readonly Block[] | undefined) =>
      blocks === undefined ? undefined : expandBlocks(blocks, registry, context)
    const next: Section = { ...section, body: expandBlocks(section.body, registry, context) }
    const header = expand(section.header)
    const footer = expand(section.footer)
    if (header) next.header = header
    if (footer) next.footer = footer
    if (section.firstPage) {
      const firstPage: NonNullable<Section['firstPage']> = {}
      const firstHeader = expand(section.firstPage.header)
      const firstFooter = expand(section.firstPage.footer)
      if (firstHeader) firstPage.header = firstHeader
      if (firstFooter) firstPage.footer = firstFooter
      next.firstPage = firstPage
    }
    return next
  })
  return { ...document, sections }
}
