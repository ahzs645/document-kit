// What the layout makes of frames placed on the page and pictures behind the text, as the kit writes them: the PDF
// and the page pictures set them where Word does. Made-up text throughout.
import { describe, expect, it } from 'vitest'

import { document, figure, floatingImage, frame, inlineImage, paragraph, section, table, type Block, type Document } from '../../src/model'
import { writeDocx } from '../../src/docx'
import { createFontProvider, layoutDocx } from '../../src/layout'
import { layOut, linesOf } from './engine'
import { diskLoader } from './sample'
import { makePng, readPackage } from '../docx/helpers'
import { pageDrawings, paintedImages } from './pdf'

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

  it('sets a picture in a frame inside the frame\'s box, in the PDF too', async () => {
    const logo = { bytes: makePng(12, 6, [200, 30, 30]) }
    const docx = await docxOf([
      paragraph('Flowing first line.'),
      frame(
        [paragraph('Label above.'), paragraph([inlineImage(logo, 60, 30)]), paragraph('Caption under.')],
        { x: 300, y: 400, width: 200 },
        { wrap: 'through' },
      ),
      frame([figure(logo, { fit: { width: 80, height: 40, upscale: true }, align: 'left' })], { x: 100, y: 600, width: 200 }, { wrap: 'through' }),
    ])
    const laid = await layoutDocx(docx, { fonts })
    const lines = laid.pageLines[0]!
    const at = (t: string) => lines.find((l) => l.text === t)!
    // One frame: its three paragraphs one under the other, the picture between the two lines.
    expect(at('Label above.').x).toBe(300)
    expect(at('Caption under.').x).toBe(300)
    expect(at('Caption under.').y - at('Label above.').y).toBeGreaterThan(30)
    const painted = paintedImages(pageDrawings(laid.toPdf().bytes)[0]!)
    expect(painted).toHaveLength(2)
    const [inFrame, figured] = painted
    // PDF y is up from the foot of the 792 pt page: the picture's top is under the label and above the caption.
    const top = 792 - inFrame!.y - inFrame!.height
    expect(inFrame).toMatchObject({ x: 300, width: 60, height: 30 })
    expect(top).toBeGreaterThan(at('Label above.').y)
    expect(top + 30).toBeLessThan(at('Caption under.').y)
    // A figure in a frame (a picture in a paragraph of its own) is placed at the frame's corner.
    expect(figured!.x).toBe(100)
    expect(figured!.width).toBeCloseTo(80, 0)
    expect(792 - figured!.y - figured!.height).toBeCloseTo(600, 0)
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

describe('a page ground behind shading', () => {
  // A header's full-page picture behind the text, and a body of filled boxes over it: a shaded cell (a table header
  // row, a callout) and a shaded paragraph. Word paints behind-text pictures under the body's text and fills.
  const ground = { bytes: makePng(10, 10, [255, 238, 170]) }
  const shaded = () =>
    writeDocx(
      document([
        section(
          [
            table([{ cells: [{ content: 'Shaded cell', shading: '0B2B4A', text: { color: 'FFFFFF' } }] }]),
            paragraph('Shaded paragraph', { shading: '0B2B4A', text: { color: 'FFFFFF' } }),
          ],
          { page, header: [floatingImage(ground, 'page', { alt: 'Ground' })] },
        ),
      ]),
    )
  const FILL = /0\.043 0\.169 0\.29 rg/

  it('paints the picture before the cell and paragraph fills in the PDF', async () => {
    const drawing = pageDrawings((await layoutDocx(await shaded(), { fonts })).toPdf().bytes)[0]!
    const [picture] = paintedImages(drawing)
    expect(picture).toMatchObject({ x: 0, y: 0, width: 612, height: 792 })
    const fill = FILL.exec(drawing)
    expect(fill).not.toBeNull()
    expect(picture!.at).toBeLessThan(fill!.index)
    // Both fills come after the picture: nothing it covers.
    expect(drawing.slice(picture!.at).match(/ re\s+f/g)!.length).toBeGreaterThanOrEqual(2)
  })

  it('marks the header\'s picture as behind the page in the layout, and in the page pictures', async () => {
    const commands = layOut(await shaded()).pages[0]!.commands
    const image = commands.find((c) => c.type === 'image')!
    expect(image.behind).toBe(true)
    const svg = (await layoutDocx(await shaded(), { fonts })).toPages().pages[0]!
    const fillAt = svg.search(/fill="#0B2B4A"/i)
    expect(fillAt).toBeGreaterThan(0)
    expect(svg.indexOf('<image')).toBeLessThan(fillAt)
  })

  it('stacks pictures in one layer in document order (relativeHeight), whatever their size', async () => {
    const big = floatingImage({ bytes: makePng(10, 10, [10, 10, 10]) }, 'page')
    const mark = floatingImage({ bytes: makePng(4, 4, [200, 30, 30]) }, { x: 400, y: 40, width: 120, height: 40 })
    const docx = await writeDocx(document([section([paragraph('Body.')], { page, header: [big, mark] })]))
    const pkg = readPackage(docx)
    const header = pkg.text(pkg.names.find((name) => /word\/header\d*\.xml$/.test(name))!)
    const heights = [...header.matchAll(/relativeHeight="(\d+)"/g)].map((m) => Number(m[1]))
    expect(heights).toEqual([1, 2])
    const drawing = pageDrawings((await layoutDocx(docx, { fonts })).toPdf().bytes)[0]!
    const painted = paintedImages(drawing)
    expect(painted.map((p) => p.width)).toEqual([612, 120])
  })
})
