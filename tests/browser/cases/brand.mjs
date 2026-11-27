/**
 * Browser case: a designed document (artwork behind a cover, text in frames, tracked text in named styles, custom
 * properties, an embedded face) written as Word, laid out with a face named by the caller, and written as HTML, all
 * from the built package in the page. Made-up content only. Throws on failure; returns a JSON-serialisable summary.
 */
import { document, floatingImage, fontEmbeddingReport, frame, heading, headerRow, paragraph, row, section, table, text } from 'document-kit'
import { writeDocx } from 'document-kit/docx'
import { writeHtml } from 'document-kit/html'
import { createFontProvider, layoutDocx } from 'document-kit/layout'

/** A 2 × 1 PNG (two grey pixels). */
const PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAC0lEQVR4nGNoAAMACocDAcTGH2YAAAAASUVORK5CYII='

function check(condition, message) {
  if (!condition) throw new Error(`brand case: ${message}`)
}

export default async function brandCase(ctx) {
  const art = { bytes: Uint8Array.from(atob(PNG_BASE64), (c) => c.charCodeAt(0)) }
  const brand = [
    { family: 'Brand Serif', bytes: ctx.loadFont('Caladea-Regular.ttf') },
    { family: 'Brand Serif', weight: 700, bytes: ctx.loadFont('Caladea-Bold.ttf') },
  ]
  const doc = document(
    [
      section(
        [
          frame([heading(1, 'Made-up cover')], { x: 72, y: 144, width: 400 }, { wrap: 'through' }),
          frame([paragraph('Prepared for nobody in particular', { style: 'CoverLine' })], { x: 72, y: 600, width: 400 }, { wrap: 'through' }),
        ],
        { page: { paper: 'letter', margins: { top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0 } }, header: [floatingImage(art, 'page', { alt: 'Artwork' })] },
      ),
      section([paragraph([text('Tracked', { characterStyle: 'Wide' })], { style: 'CoverLine' }), table([headerRow(['Quoted face']), row(['Body'])])], { header: [] }),
    ],
    {
      metadata: { title: 'Made-up brand document', custom: { 'app.id': 'case-1' } },
      theme: { body: { font: 'Brand Serif' }, table: { headerText: { font: 'Brand "Quoted" Face', color: 'aa3311', bold: true } } },
      styles: {
        paragraph: [{ id: 'CoverLine', text: { size: 14, characterSpacing: 1 } }],
        character: [{ id: 'Wide', text: { characterSpacing: 3 } }],
      },
    },
  )
  const report = fontEmbeddingReport(brand, 'docx')
  check(report.every((face) => face.embedded), `faces refused: ${JSON.stringify(report)}`)
  const docx = await writeDocx(doc, { embedFonts: brand })
  const fonts = createFontProvider((file) => ctx.loadFont(file.file))
  const laid = await layoutDocx(docx, { fonts })
  check(laid.pageCount === 2, `laid out ${laid.pageCount} pages`)
  const title = laid.pageLines[0].find((line) => line.text === 'Made-up cover')
  check(title && title.x === 72 && title.y > 144, `the framed title is at ${JSON.stringify(title)}`)
  check(!laid.losses.some((loss) => /brand serif/i.test(loss.detail)), 'the embedded face was substituted')
  const named = await layoutDocx(docx, { fonts, namedFonts: [{ family: 'Brand Serif', bytes: ctx.loadFont('Tinos-Regular.ttf') }] })
  const pdf = named.toPdf().bytes
  check(new TextDecoder('latin1').decode(pdf).includes('Tinos'), 'the named face is not in the PDF')
  const html = writeHtml(doc, { fonts: brand })
  check(html.includes('@font-face{font-family:"Brand Serif";font-weight:700'), 'the HTML does not carry the bold face')
  check(html.includes('letter-spacing:3pt'), 'the HTML lost the character style tracking')
  // The page itself, put in a frame of this document, shows its sheets at their paper size.
  const frameEl = document_.createElement('iframe')
  frameEl.style.width = '900px'
  frameEl.style.height = '1200px'
  frameEl.srcdoc = html
  document_.body.append(frameEl)
  await new Promise((resolve) => frameEl.addEventListener('load', resolve, { once: true }))
  const sheets = frameEl.contentDocument.querySelectorAll('.dk-sheet')
  check(sheets.length === 2, `the HTML has ${sheets.length} sheets`)
  const width = sheets[0].getBoundingClientRect().width
  check(Math.abs(width - 816) < 1, `the first sheet is ${width}px wide, not 816`)
  const placed = frameEl.contentDocument.querySelector('.dk-frame').getBoundingClientRect()
  const sheetBox = sheets[0].getBoundingClientRect()
  check(Math.abs(placed.left - sheetBox.left - 96) < 1 && Math.abs(placed.top - sheetBox.top - 192) < 1, 'the frame is not where it was put')
  const view = frameEl.contentWindow
  // A quoted face name in the header cell's style attribute does not cut off what follows it.
  const headerText = frameEl.contentDocument.querySelector('th p')
  const headerStyle = view.getComputedStyle(headerText)
  check(headerStyle.color === 'rgb(170, 51, 17)', `the header text is ${headerStyle.color}`)
  check(/Quoted.*Face.*sans-serif/.test(headerStyle.fontFamily), `the header face is ${headerStyle.fontFamily}`)
  frameEl.remove()
  return { pages: laid.pageCount, htmlBytes: html.length, sheets: sheets.length }
}

const document_ = globalThis.document
