// Line heights in points in the layout and the PDF it writes: an exact height sets every baseline that many points
// below the last whatever the face, as Word sets it; a multiple is measured on the face's own line and does not.
// Made-up text throughout.
import { describe, expect, it } from 'vitest'

import { document, frame, lineBreak, paragraph, section, text, type Block, type Document, type LineHeight } from '../../src/model'
import { writeDocx } from '../../src/docx'
import { createFontProvider, layoutDocx } from '../../src/layout'
import { diskLoader } from './sample'
import { PdfFile } from '../../src/ream/pdf-reader/document'
import { extractPageText } from '../../src/ream/pdf-reader/text'

const fonts = createFontProvider(diskLoader)
const page = { paper: 'letter' as const, margins: { top: 72, bottom: 72, left: 72, right: 72, header: 36, footer: 36 } }
const PAGE_HEIGHT = 792

/** Five lines, broken by hand so every face sets the same words on each. */
const fiveLines = (font: string, size = 11) =>
  ['First made-up line', 'Second made-up line', 'Third made-up line', 'Fourth made-up line', 'Fifth made-up line'].flatMap(
    (words, i) => [...(i ? [lineBreak()] : []), text(words, { font, size })],
  )

/** The baselines the PDF draws text on, top down, in points from the top of the first page; those left of `fromX` dropped. */
async function pdfBaselines(blocks: Block[], extra: Partial<Document> = {}, fromX = 0): Promise<number[]> {
  const laid = await layoutDocx(await writeDocx(document([section(blocks, { page })], extra)), { fonts })
  const file = PdfFile.parse(laid.toPdf().bytes)
  const runs = extractPageText(file, file.pages()[0]!).filter((run) => run.text.trim() && run.x >= fromX)
  const ys = [...new Set(runs.map((run) => Math.round((PAGE_HEIGHT - run.y) * 1000) / 1000))].sort((a, b) => a - b)
  // The layout's own account of the page says the same.
  expect(ys).toEqual(laid.pageLines[0]!.filter((l) => l.text.trim() && l.x >= fromX).map((l) => expect.closeTo(l.y, 1)))
  return ys
}

const steps = (ys: number[]) => ys.slice(1).map((y, i) => y - ys[i]!)

/** Every step between baselines `pitch` points, to a thousandth. */
function expectPitch(ys: number[], pitch: number) {
  for (const step of steps(ys)) expect(step).toBeCloseTo(pitch, 3)
}

// Carlito (Calibri's stand-in) and Arimo (Arial's) have different natural lines: 1.22 and 1.15 of the size.
const FACES = ['Calibri', 'Arial'] as const

describe('an exact line height in the PDF', () => {
  const lines = (font: string, line: LineHeight) =>
    pdfBaselines([
      paragraph(fiveLines(font), { spacing: { line, after: 0 } }),
      paragraph([text('After it.', { font, size: 11 })], { spacing: { line, after: 0 } }),
    ])

  it('sets the baselines exactly that many points apart, in two faces with different natural lines', async () => {
    const [calibri, arial] = await Promise.all(FACES.map((font) => lines(font, { exact: 14 })))
    expect(calibri).toHaveLength(6)
    expect(arial).toHaveLength(6)
    expectPitch(calibri!, 14)
    expectPitch(arial!, 14)
    // The faces' lines stand in the same boxes: their baselines differ only by the faces' descents, the same on every
    // line, so the two set line for line together.
    const offset = calibri![0]! - arial![0]!
    expect(Math.abs(offset)).toBeLessThan(1)
    calibri!.forEach((y, i) => expect(y - arial![i]!).toBeCloseTo(offset, 3))
  })

  it('where a multiple drifts apart between the same two faces', async () => {
    const [calibri, arial] = await Promise.all(FACES.map((font) => lines(font, 1)))
    const [c, a] = [steps(calibri!)[0]!, steps(arial!)[0]!]
    expect(Math.abs(c - a)).toBeGreaterThan(0.5)
    // Five lines on, the faces are more than three points apart.
    expect(Math.abs(calibri![5]! - arial![5]! - (calibri![0]! - arial![0]!))).toBeGreaterThan(3)
  })

  it('holds a line of larger text to the height too, where at least lets it grow', async () => {
    const big = (line: LineHeight) =>
      pdfBaselines([
        paragraph([text('Small', { size: 11 }), lineBreak(), text('Large', { size: 24 }), lineBreak(), text('Small', { size: 11 })], {
          spacing: { line, after: 0 },
        }),
      ])
    // Each line's box is 14pt whatever its text; a baseline sits the line's descent above its box's foot, so the line
    // after the large one is back on the pitch.
    const exact = await big({ exact: 14 })
    expect(exact[2]! - exact[0]!).toBeCloseTo(28, 3)
    const least = await big({ atLeast: 14 })
    // The large line is taller than 14pt; the small one after it is not held under 14pt.
    expect(steps(least)[0]).toBeGreaterThan(20)
    expect(steps(least)[1]).toBeGreaterThanOrEqual(14 - 1e-3)
    // A least height over 11pt text is the height itself.
    expectPitch(await pdfBaselines([paragraph(fiveLines('Arial'), { spacing: { line: { atLeast: 15 }, after: 0 } })]), 15)
  })

  it('takes the height from the theme\'s body and from a named style', async () => {
    for (const font of FACES) {
      const themed = await pdfBaselines([paragraph(fiveLines(font), { spacing: { after: 0 } })], {
        theme: { body: { font, line: { exact: 13 } } },
      })
      expectPitch(themed, 13)
      const styled = await pdfBaselines([paragraph(fiveLines(font), { style: 'CoverLine' })], {
        styles: { paragraph: [{ id: 'CoverLine', text: { font, after: 0, line: { exact: 12.5 } } }] },
      })
      expectPitch(styled, 12.5)
    }
  })

  it('spaces paragraphs on the same pitch, a paragraph\'s space after added to it', async () => {
    const ys = await pdfBaselines(
      ['One', 'Two', 'Three'].map((words) => paragraph([text(words, { font: 'Arial' })], { spacing: { line: { exact: 14 }, after: 6 } })),
    )
    expectPitch(ys, 20)
  })
})

describe('an exact line height in a frame placed on the page', () => {
  const framed = (font: string) =>
    pdfBaselines(
      [
        paragraph('Flow.'),
        frame([paragraph(fiveLines(font), { spacing: { line: { exact: 14 }, after: 0 } })], { x: 300, y: 400, width: 220 }, { wrap: 'through' }),
      ],
      {},
      300,
    )

  it('sets the frame\'s baselines exactly that many points apart, under the frame\'s top, in both faces', async () => {
    const [calibri, arial] = await Promise.all(FACES.map(framed))
    for (const ys of [calibri!, arial!]) {
      expect(ys).toHaveLength(5)
      expectPitch(ys, 14)
      // The first line's box is the frame's top 14pt; its baseline sits above the box's foot by the face's descent.
      expect(ys[0]).toBeGreaterThan(400 + 14 - 4)
      expect(ys[0]).toBeLessThanOrEqual(400 + 14)
    }
    const offset = calibri![0]! - arial![0]!
    calibri!.forEach((y, i) => expect(y - arial![i]!).toBeCloseTo(offset, 3))
  })

  it('in a frame measured from the margins, with the height from a named style', async () => {
    const ys = await pdfBaselines(
      [
        paragraph('Flow.'),
        frame([paragraph(fiveLines('Arial'), { style: 'CoverLine' })], { x: 228, y: 300, width: 220 }, { relativeTo: 'margin', wrap: 'through' }),
      ],
      { styles: { paragraph: [{ id: 'CoverLine', text: { after: 0, line: { exact: 16 } } }] } },
      300,
    )
    expect(ys).toHaveLength(5)
    expectPitch(ys, 16)
    expect(ys[0]).toBeGreaterThan(72 + 300 + 16 - 4)
    expect(ys[0]).toBeLessThanOrEqual(72 + 300 + 16)
  })
})
