import { describe, expect, it } from 'vitest'
import {
  cell,
  custom,
  defineCustomBlocks,
  document,
  expandDocument,
  hasCustomBlocks,
  paragraph,
  row,
  section,
  table,
  UnknownCustomBlockError,
  type Block,
} from '../../src/index'

describe('custom blocks', () => {
  it('expands a registered kind into the kit’s blocks, with the section’s page in the context', () => {
    const registry = defineCustomBlocks({
      banner: (data: { words: string }, context) => [paragraph(`${data.words} @ ${Math.round(context.page.contentWidth)}pt`)],
    })
    const doc = document([section([custom('banner', { words: 'Hello' })], { page: { orientation: 'landscape' } })])
    const expanded = expandDocument(doc, registry)
    expect(expanded.sections[0]?.body).toEqual([{ type: 'paragraph', content: 'Hello @ 648pt' }])
    expect(hasCustomBlocks(expanded)).toBe(false)
  })

  it('fails loudly, naming the kind, when nothing is registered for it', () => {
    const doc = document([section([custom('scaleBar', {})])])
    expect(() => expandDocument(doc)).toThrow(UnknownCustomBlockError)
    expect(() => expandDocument(doc)).toThrow(/"scaleBar"/)
    const registry = defineCustomBlocks({ matrix: () => [] })
    try {
      expandDocument(doc, registry)
      expect.unreachable()
    } catch (error) {
      expect(error).toBeInstanceOf(UnknownCustomBlockError)
      expect((error as UnknownCustomBlockError).kind).toBe('scaleBar')
      expect((error as Error).message).toContain('registered: matrix')
    }
  })

  it('expands inside table cells and expansions of expansions, and carries a page break', () => {
    const registry = defineCustomBlocks()
      .with('outer', () => [custom('inner', 'deep')])
      .with('inner', (data: string) => [paragraph(data)])
    const doc = document([
      section([
        table([row([cell(null, { blocks: [custom('outer', null)] })])]),
        custom('inner', 'after', { pageBreakBefore: true }),
      ]),
    ])
    const body = expandDocument(doc, registry).sections[0]!.body
    const first = body[0] as Extract<Block, { type: 'table' }>
    const inCell = first.rows[0]!.cells[0] as { blocks: Block[] }
    expect(inCell.blocks).toEqual([{ type: 'paragraph', content: 'deep' }])
    expect(body[1]).toEqual({ type: 'paragraph', content: 'after', pageBreakBefore: true })
  })

  it('stops an expansion that loops', () => {
    const registry = defineCustomBlocks().with('loop', () => [custom('loop', null)])
    expect(() => expandDocument(document([section([custom('loop', null)])]), registry)).toThrow(/loop/)
  })

  it('leaves a document without custom blocks as it is', () => {
    const doc = document([section([paragraph('a')])])
    expect(expandDocument(doc)).toBe(doc)
  })
})
