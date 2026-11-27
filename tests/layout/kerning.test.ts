// Kerned text in the layout: a run, a style or the theme asks for kerning (Word's w:kern), and the layout closes the
// face's own pairs up by what its GPOS (or legacy kern) table says, from the size asked and only there.
import { describe, expect, it } from 'vitest'

import { document, paragraph, section, text, type Block, type Document } from '../../src/model'
import { writeDocx } from '../../src/docx'
import { createFontProvider, layoutDocx } from '../../src/layout'
import { parseTtf } from '../../src/ream/core/font/ttf-parser'
import { parseKernTable } from '../../src/ream/core/font/opentype-layout'
import { fontFile, diskLoader } from './sample'
import { pageDrawings } from './pdf'

const fonts = createFontProvider(diskLoader)
const page = { paper: 'letter' as const, margins: { top: 72, bottom: 72, left: 72, right: 72, header: 36, footer: 36 } }
const docxOf = (blocks: Block[], extra: Partial<Document> = {}) => writeDocx(document([section(blocks, { page })], extra))

// A right-aligned line ends at the margin, so it starts exactly its width left of it: two layouts' starts differ by
// what kerning took out.
const startOf = async (blocks: Block[], extra: Partial<Document> = {}) =>
  (await layoutDocx(await docxOf(blocks, extra), { fonts })).pageLines[0]![0]!.x

/**
 * What a face's pair kerning takes out of `text` at `size`, points (negative: closed up). Pairs with a space are left
 * out: the layout measures words apart (a known gap, docs/layout.md).
 */
function kernedBy(file: string, text: string, size: number): number {
  const face = parseTtf(fontFile(file))
  let units = 0
  const chars = [...text]
  for (let i = 1; i < chars.length; i++) {
    if (chars[i - 1] === ' ' || chars[i] === ' ') continue
    const left = face.glyphForCodepoint(chars[i - 1]!.codePointAt(0)!)
    const right = face.glyphForCodepoint(chars[i]!.codePointAt(0)!)
    units += face.kerning.get(`${left},${right}`) ?? 0
  }
  return (units * size) / face.unitsPerEm
}

describe('kerning', () => {
  const PAIRS = 'AV To AV To'

  it('closes up "AV" and "To" by the face\'s own pairs when a run asks (Carlito, whose pairs sit in an extension lookup)', async () => {
    const expected = kernedBy('Carlito-Regular.ttf', PAIRS, 24)
    // Both pairs kern in Carlito: AV and To close up.
    expect(kernedBy('Carlito-Regular.ttf', 'AV', 24)).toBeLessThan(-0.5)
    expect(kernedBy('Carlito-Regular.ttf', 'To', 24)).toBeLessThan(-0.5)
    const plain = await startOf([paragraph([text(PAIRS, { size: 24 })], { align: 'right' })])
    const kerned = await startOf([paragraph([text(PAIRS, { size: 24, kern: true })], { align: 'right' })])
    // Right-aligned: the kerned line is narrower, so it starts further right, by exactly the pairs' kerning.
    expect(kerned - plain).toBeCloseTo(-expected, 1)
  })

  it('kerns Arial text (Arimo) by its pairs too', async () => {
    const expected = kernedBy('Arimo-Regular.ttf', PAIRS, 24)
    expect(expected).toBeLessThan(-2)
    const plain = await startOf([paragraph([text(PAIRS, { size: 24, font: 'Arial' })], { align: 'right' })])
    const kerned = await startOf([paragraph([text(PAIRS, { size: 24, font: 'Arial', kern: true })], { align: 'right' })])
    expect(kerned - plain).toBeCloseTo(-expected, 1)
  })

  it('kerns nothing a run does not ask for, as Word', async () => {
    const plain = await startOf([paragraph([text(PAIRS, { size: 24, font: 'Arial' })], { align: 'right' })])
    // Set at the glyphs' own widths: the line is as wide as the sum of its advances.
    const face = parseTtf(fontFile('Arimo-Regular.ttf'))
    const width = [...PAIRS].reduce((sum, c) => sum + face.advanceWidths[face.glyphForCodepoint(c.codePointAt(0)!)]!, 0)
    expect(540 - plain).toBeCloseTo((width * 24) / face.unitsPerEm, 1)
  })

  it('kerns from the size asked and up, and not below it', async () => {
    const at = (size: number, kern: number | boolean) =>
      startOf([paragraph([text(PAIRS, { size, kern })], { align: 'right' })])
    // Kerning from 14 pt: 12 pt text is set plain, 16 pt kerned.
    expect(await at(12, 14)).toBeCloseTo(await at(12, false), 2)
    expect((await at(16, 14)) - (await at(16, false))).toBeCloseTo(-kernedBy('Carlito-Regular.ttf', PAIRS, 16), 1)
  })

  it('takes kerning from the theme\'s body for the whole document, and a run can turn it off', async () => {
    const theme = { body: { kern: true } }
    const plain = await startOf([paragraph([text(PAIRS, { size: 24 })], { align: 'right' })])
    const themed = await startOf([paragraph([text(PAIRS, { size: 24 })], { align: 'right' })], { theme })
    const off = await startOf([paragraph([text(PAIRS, { size: 24, kern: false })], { align: 'right' })], { theme })
    expect(themed - plain).toBeCloseTo(-kernedBy('Carlito-Regular.ttf', PAIRS, 24), 1)
    expect(off).toBeCloseTo(plain, 2)
  })

  it('takes a named style\'s kerning', async () => {
    const styles: Document['styles'] = { paragraph: [{ id: 'Display', text: { size: 24, kern: 10, align: 'right' } }] }
    const plain = await startOf([paragraph(PAIRS, { style: 'Display' })], {
      styles: { paragraph: [{ id: 'Display', text: { size: 24, align: 'right' } }] },
    })
    const kerned = await startOf([paragraph(PAIRS, { style: 'Display' })], { styles })
    expect(kerned - plain).toBeCloseTo(-kernedBy('Carlito-Regular.ttf', PAIRS, 24), 1)
  })

  it('draws a kerned run with the kerning in the PDF (TJ adjustments), a plain run without', async () => {
    const pdf = async (kern: boolean) =>
      pageDrawings((await layoutDocx(await docxOf([paragraph([text('AVATAR', { size: 24, kern })])]), { fonts })).toPdf().bytes)[0]!
    expect(await pdf(true)).toMatch(/\[<[0-9A-F]+> [\d.]+ <[0-9A-F]+>[^\]]*\] TJ/)
    expect(await pdf(false)).not.toMatch(/\] TJ/)
  })
})

describe('the legacy kern table', () => {
  it('is read for a face whose GPOS kerns nothing', () => {
    // A 'kern' table, version 0, one horizontal format 0 subtable of two pairs: (36,57) −80 and (55,82) −40.
    const bytes = new Uint8Array(4 + 14 + 2 * 6)
    const view = new DataView(bytes.buffer)
    view.setUint16(0, 0) // version
    view.setUint16(2, 1) // nTables
    view.setUint16(4, 0) // subtable version
    view.setUint16(6, 14 + 12) // length
    view.setUint16(8, 0x0001) // coverage: horizontal, format 0
    view.setUint16(10, 2) // nPairs
    const pairs = [[36, 57, -80], [55, 82, -40]] as const
    pairs.forEach(([l, r, v], i) => {
      view.setUint16(18 + i * 6, l)
      view.setUint16(20 + i * 6, r)
      view.setInt16(22 + i * 6, v)
    })
    const kerning = parseKernTable(bytes, { offset: 0, length: bytes.length })
    expect([...kerning]).toEqual([['36,57', -80], ['55,82', -40]])
  })
})
