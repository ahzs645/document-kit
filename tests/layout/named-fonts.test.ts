// The faces a document embeds and the faces a caller names, TrueType and CFF, in the layout, the PDF and the page
// pictures. Made-up text throughout.
import { describe, expect, it } from 'vitest'
import { strFromU8, unzlibSync } from 'fflate'

import { document, heading, paragraph, section, text, type Block, type Document } from '../../src/model'
import { writeDocx } from '../../src/docx'
import { createFontProvider, layoutDocx } from '../../src/layout'
import { layOut } from './engine'
import { diskLoader } from './sample'
import { cidKeyedCff, otf, ttf } from '../fonts'
import { PdfFile } from '../../src/ream/pdf-reader/document'
import { extractPageText } from '../../src/ream/pdf-reader/text'

const fonts = createFontProvider(diskLoader)
const page = { paper: 'letter' as const, margins: { top: 72, bottom: 72, left: 72, right: 72, header: 36, footer: 36 } }
const docxOf = (blocks: Block[], extra: Partial<Document> = {}, options: Parameters<typeof writeDocx>[1] = {}) =>
  writeDocx(document([section(blocks, { page })], extra), options)

/** Every stream of a PDF, inflated where it is deflated, as text. */
function pdfStreams(bytes: Uint8Array): string {
  const raw = strFromU8(bytes, true)
  const out: string[] = []
  for (const m of raw.matchAll(/stream\r?\n/g)) {
    const start = m.index! + m[0].length
    const end = raw.indexOf('endstream', start)
    const data = bytes.subarray(start, end)
    try {
      out.push(strFromU8(unzlibSync(data), true))
    } catch {
      out.push(strFromU8(data, true))
    }
  }
  return `${raw}\n${out.join('\n')}`
}

describe('fonts by family name', () => {
  const brand = [
    { family: 'Brand Sans', bytes: ttf() },
    { family: 'Brand Sans', weight: 700, bytes: ttf('Bold') },
  ]

  it('draws a family the document embeds with its own faces, with no substitution', async () => {
    const docx = await docxOf([paragraph([text('Embedded face', { font: 'Brand Sans' })])], {}, { embedFonts: brand })
    const laid = await layoutDocx(docx, { fonts })
    expect(laid.losses.filter((l) => /brand sans/i.test(l.detail))).toEqual([])
    expect(pdfStreams(laid.toPdf().bytes)).toMatch(/\+Roboto-Regular/)
  })

  it('thickens bold text in an embedded family that brought no bold face', async () => {
    const docx = await docxOf([paragraph([text('Bold here', { font: 'Brand Sans', bold: true })])], {}, { embedFonts: [brand[0]!] })
    const line = layOut(docx).pages[0]!.commands.find((c) => c.type === 'line') as unknown as { line: { tokens: Array<{ kind: string; synthetic?: { bold?: boolean } }> } }
    expect(line.line.tokens.some((t) => t.kind === 'text' && t.synthetic?.bold === true)).toBe(true)
  })

  it('draws a caller\'s TrueType and CFF faces by name, over the stand-ins', async () => {
    const docx = await docxOf([
      paragraph([text('Brand face', { font: 'Brand Sans' })]),
      paragraph([text('Code face', { font: 'Code Mono' })]),
    ])
    const laid = await layoutDocx(docx, { fonts, namedFonts: [...brand, { family: 'code mono', bytes: otf() }] })
    expect(laid.losses.filter((l) => /brand sans|code mono/i.test(l.detail))).toEqual([])
    const pdf = laid.toPdf().bytes
    const streams = pdfStreams(pdf)
    expect(streams).toMatch(/\/FontFile3/)
    expect(streams).toMatch(/\/Subtype\s*\/OpenType/)
    expect(streams).toMatch(/\/CIDFontType0\b/)
    expect(streams).toMatch(/\/BaseFont\s*\/Inconsolata/)
    expect(streams).toMatch(/\+Roboto-Regular/)
    const file = PdfFile.parse(pdf)
    const words = extractPageText(file, file.pages()[0]!).map((run) => run.text).join('')
    expect(words.replace(/\s+/g, '')).toContain('Codeface')
    // The page pictures draw both from their outlines.
    const svg = laid.toPages().pages[0]!
    expect((svg.match(/<use /g) ?? []).length).toBeGreaterThanOrEqual('BrandfaceCodeface'.length)
  })

  it('puts a caller\'s face over the face the document embeds', async () => {
    const docx = await docxOf([paragraph([text('Which face', { font: 'Brand Sans' })])], {}, { embedFonts: brand })
    const laid = await layoutDocx(docx, { fonts, namedFonts: [{ family: 'Brand Sans', bytes: otf() }] })
    const streams = pdfStreams(laid.toPdf().bytes)
    expect(streams).toMatch(/\/BaseFont\s*\/Inconsolata/)
    expect(streams).not.toMatch(/Roboto/)
  })

  it('names a face it cannot draw with and falls back', async () => {
    const docx = await docxOf([paragraph([text('Fallback', { font: 'CID Face' })])])
    const laid = await layoutDocx(docx, { fonts, namedFonts: [{ family: 'CID Face', bytes: cidKeyedCff() }] })
    expect(laid.losses.some((l) => /CID-keyed CFF/.test(l.detail) && /"CID Face"/.test(l.detail))).toBe(true)
    expect(laid.losses.some((l) => /"cid face" drawn with its stand-in/.test(l.detail))).toBe(true)
    expect(laid.pageText[0]).toContain('Fallback')
  })

  it('lays out with named fonts alone, no provider', async () => {
    const docx = await writeDocx(
      document([section([heading(1, 'Only named'), paragraph('Body.')])], { theme: { body: { font: 'Brand Sans' } } }),
    )
    const laid = await layoutDocx(docx, { namedFonts: brand })
    expect(laid.pageText[0]).toContain('Only named')
  })
})
