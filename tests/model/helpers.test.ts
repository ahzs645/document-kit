import { describe, expect, it } from 'vitest'
import {
  DEFAULT_THEME,
  fillSlots,
  fitImage,
  imageInfo,
  pageBox,
  plainText,
  resolveImage,
  resolveSentences,
  resolveTheme,
  sentence,
  tableGrid,
  type TableBlock,
} from '../../src/index'
import { makeJpegHeader, makePng } from '../docx/helpers'

describe('sentences', () => {
  it('fill slots, and leave the sentence out when one has no value', () => {
    expect(fillSlots('{n} trees on {area} ha', { n: 40, area: '2.5' })).toEqual({ text: '40 trees on 2.5 ha', missing: [] })
    expect(fillSlots('{a} and {b} and {a}', { a: null })).toEqual({ text: null, missing: ['a', 'b'] })
    expect(fillSlots('{a}', { a: 0 }).text).toBe('0')
  })

  it('drop a left-out sentence with its separating space', () => {
    const a = sentence('One {x}.', { x: 1 })
    const b = sentence('Two {y}.', { y: null })
    const c = sentence('Three.', {})
    expect(plainText(resolveSentences([a, ' ', b, ' ', c]))).toBe('One 1. Three.')
    expect(plainText(resolveSentences([a, ' ', b]))).toBe('One 1.')
    expect(plainText(resolveSentences([b, ' ', c]))).toBe('Three.')
    expect(resolveSentences([b])).toEqual([])
  })
})

describe('tableGrid', () => {
  const rows = (cells: TableBlock['rows']) => ({ rows: cells })
  it('places cells around row and column spans', () => {
    const grid = tableGrid(
      rows([
        { cells: [{ content: 'P', rowSpan: 2 }, { content: 'Samples', colSpan: 2 }] },
        { cells: ['S1', 'S2'] },
        { cells: ['a', 'b', 'c'] },
      ]),
    )
    expect(grid.columns).toBe(3)
    expect(grid.rows[1]?.map((entry) => entry.kind)).toEqual(['covered', 'cell', 'cell'])
    const second = grid.rows[1]?.[1]
    expect(second?.kind === 'cell' && second.placed.column).toBe(1)
  })

  it('spans a lone note cell across the table and pads nothing else', () => {
    const grid = tableGrid(rows([{ cells: ['a', 'b', 'c'] }, { role: 'note', cells: ['note'] }]))
    const note = grid.rows[1]?.[0]
    expect(note?.kind === 'cell' && note.placed.colSpan).toBe(3)
  })

  it('refuses rows wider than the columns and spans past the end', () => {
    expect(() => tableGrid({ rows: [{ cells: ['a', 'b', 'c'] }], columnWidths: [1, 1] })).toThrow(/needs 3 columns/)
    expect(() => tableGrid(rows([{ cells: [{ content: 'a', rowSpan: 3 }] }, { cells: [] }]))).toThrow(/past the table/)
  })
})

describe('pictures', () => {
  it('reads PNG and JPEG sizes from their headers', () => {
    expect(imageInfo(makePng(30, 20))).toEqual({ format: 'png', width: 30, height: 20 })
    expect(imageInfo(makeJpegHeader(640, 480))).toEqual({ format: 'jpeg', width: 640, height: 480 })
    expect(imageInfo(new Uint8Array([1, 2, 3]))).toBeNull()
    expect(() => resolveImage({ bytes: new Uint8Array([1, 2, 3]) })).toThrow(/PNG or JPEG/)
    expect(resolveImage({ bytes: new Uint8Array(0), format: 'png', width: 5, height: 6 })).toEqual({
      format: 'png',
      width: 5,
      height: 6,
    })
  })

  it('fits inside a box keeping proportions, enlarging only when asked', () => {
    expect(fitImage({ width: 800, height: 400 }, { width: 300, height: 300 })).toEqual({ width: 300, height: 150 })
    expect(fitImage({ width: 80, height: 40 }, { width: 300, height: 300 })).toEqual({ width: 60, height: 30 })
    expect(fitImage({ width: 80, height: 40 }, { width: 300, height: 300 }, true)).toEqual({ width: 300, height: 150 })
  })
})

describe('pages and themes', () => {
  it('turns paper for landscape and reports the text box', () => {
    const box = pageBox({ paper: 'tabloid', orientation: 'landscape', margins: { left: 36, right: 36, top: 36, bottom: 36 } }, DEFAULT_THEME)
    expect([box.width, box.height]).toEqual([1224, 792])
    expect([box.contentWidth, box.contentHeight]).toEqual([1152, 720])
    expect(pageBox({ paper: 'legal' }, DEFAULT_THEME).height).toBe(1008)
    expect(() => pageBox({ paper: 'quarto' as never }, DEFAULT_THEME)).toThrow(/unknown paper/)
  })

  it('merges a partial theme over the default, heading by heading', () => {
    const theme = resolveTheme({ body: { font: 'Times New Roman' }, headings: [{ color: '035642' }], table: { headerFill: 'cfe6d8' } })
    expect(theme.body).toMatchObject({ font: 'Times New Roman', size: 11 })
    expect(theme.headings[0]).toMatchObject({ color: '035642', size: 16, bold: true })
    expect(theme.headings[1]).toEqual(DEFAULT_THEME.headings[1])
    expect(theme.table.headerFill).toBe('cfe6d8')
    expect(theme.table.outerBorder).toEqual(DEFAULT_THEME.table.outerBorder)
  })
})
