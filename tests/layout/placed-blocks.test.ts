// What the layout makes of frames placed on the page and pictures behind the text, as the kit writes them: the PDF
// and the page pictures set them where Word does. Made-up text throughout.
import { describe, expect, it } from 'vitest'

import { document, floatingImage, frame, paragraph, section, type Block, type Document } from '../../src/model'
import { writeDocx } from '../../src/docx'
import { createFontProvider, layoutDocx } from '../../src/layout'
import { layOut, linesOf } from './engine'
import { diskLoader } from './sample'
import { makePng } from '../docx/helpers'

const fonts = createFontProvider(diskLoader)
const page = { paper: 'letter' as const, margins: { top: 72, bottom: 72, left: 72, right: 72, header: 36, footer: 36 } }
const docxOf = (blocks: Block[], extra: Partial<Document> = {}, options: Parameters<typeof writeDocx>[1] = {}) =>
  writeDocx(document([section(blocks, { page })], extra), options)

describe('frames placed on the page', () => {
  it('sets the frame\'s lines at its x and below its y, the flow untouched', async () => {
    const docx = await docxOf([
      paragraph('Flowing first line.'),
      frame([paragraph('Framed line.')], { x: 300, y: 400, width: 200 }, { wrap: 'through' }),
      paragraph('Flowing second line.'),
    ])
    const laid = await layoutDocx(docx, { fonts })
    const lines = laid.pageLines[0]!
    const at = (t: string) => lines.find((l) => l.text === t)!
    expect(at('Framed line.').x).toBe(300)
    expect(at('Framed line.').y).toBeGreaterThan(400)
    expect(at('Framed line.').y).toBeLessThan(420)
    // The frame takes nothing from the flow: the second line follows the first.
    expect(at('Flowing second line.').y - at('Flowing first line.').y).toBeLessThan(25)
  })

  it('sets the flow\'s lines beside a frame that wraps around', async () => {
    const filler = Array.from({ length: 14 }, (_, i) => paragraph(`Line ${i + 1} of made-up body text, long enough to run across the text width.`))
    const docx = await docxOf([...filler.slice(0, 2), frame([paragraph('Boxed.')], { x: 72, y: 150, width: 150, height: 100 }), ...filler.slice(2)])
    const lines = (await layoutDocx(docx, { fonts })).pageLines[0]!
    const beside = lines.filter((l) => l.y > 150 && l.y < 250 && l.text.startsWith('Line'))
    expect(beside.length).toBeGreaterThan(0)
    for (const line of beside) expect(line.x).toBeGreaterThanOrEqual(222)
    expect(lines.find((l) => l.text.startsWith('Line 14'))!.x).toBe(72)
  })

  it('measures from the margins when asked', async () => {
    const docx = await docxOf([paragraph('Flow.'), frame([paragraph('In the margin box.')], { x: 36, y: 100, width: 200 }, { relativeTo: 'margin', wrap: 'through' })])
    const line = (await layoutDocx(docx, { fonts })).pageLines[0]!.find((l) => l.text === 'In the margin box.')!
    expect(line.x).toBe(72 + 36)
    expect(line.y).toBeGreaterThan(172)
    expect(line.y).toBeLessThan(192)
  })

  it('places a frame from the first page\'s header on that page', async () => {
    const docx = await writeDocx(
      document([
        section([paragraph('Body.')], {
          page,
          firstPage: { header: [frame([paragraph('Header-borne label.')], { x: 400, y: 700, width: 150 }, { wrap: 'through' })] },
        }),
      ]),
    )
    const line = (await layoutDocx(docx, { fonts })).pageLines[0]!.find((l) => l.text === 'Header-borne label.')!
    expect(line.x).toBe(400)
    expect(line.y).toBeGreaterThan(700)
  })
})

describe('a picture behind the text', () => {
  it('is drawn first, over the whole page, and the text keeps its place', async () => {
    const art = { bytes: makePng(51, 66, [20, 50, 90]) }
    const plain = await docxOf([paragraph('Over the artwork.')])
    const backed = await docxOf([floatingImage(art, 'page', { alt: 'Artwork' }), paragraph('Over the artwork.')])
    const laid = layOut(backed)
    const commands = laid.pages[0]!.commands
    const image = commands.find((c) => c.type === 'image') as { x: number; y: number; width: number; height: number }
    expect(image).toMatchObject({ x: 0, y: 0, width: 612, height: 792 })
    expect(commands.findIndex((c) => c.type === 'image')).toBeLessThan(commands.findIndex((c) => c.type === 'line'))
    const y = (docx: Uint8Array) => linesOf(layOut(docx).pages[0]!).find((l) => l.text === 'Over the artwork.')!.baselineY
    expect(Math.abs(y(backed) - y(plain))).toBeLessThan(0.5)
  })

  it('is drawn on every page from a section\'s header', async () => {
    const art = { bytes: makePng(10, 10) }
    const docx = await writeDocx(
      document([
        section([paragraph('One.'), { type: 'pageBreak' }, paragraph('Two.')], { page, header: [floatingImage(art, 'page')] }),
      ]),
    )
    const laid = layOut(docx)
    expect(laid.pages).toHaveLength(2)
    for (const p of laid.pages) expect(p.commands.some((c) => c.type === 'image')).toBe(true)
  })
})
