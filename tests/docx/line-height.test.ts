// Line heights in points (exact and at-least) beside the multiple, as the Word writer writes them on paragraphs, named
// styles, the theme's styles and tables. Made-up text throughout.
import { describe, expect, it } from 'vitest'

import { document, paragraph, section, table, type Block, type Document } from '../../src/model'
import { writeDocx } from '../../src/docx'
import { readPackage } from './helpers'

const write = async (blocks: Block[], extra: Partial<Document> = {}) => readPackage(await writeDocx(document([section(blocks)], extra)))
const styleXml = (styles: string, id: string) => new RegExp(`<w:style [^>]*w:styleId="${id}"[\\s\\S]*?</w:style>`).exec(styles)?.[0] ?? ''
/** The `w:spacing` element of each paragraph's own properties in the body. */
const paragraphSpacings = (xml: string): string[] =>
  [...xml.matchAll(/<w:p>([\s\S]*?)<\/w:p>|<w:p [^>]*>([\s\S]*?)<\/w:p>/g)].map((m) => /<w:pPr>[\s\S]*?(<w:spacing [^>]*\/>)/.exec(m[1] ?? m[2] ?? '')?.[1] ?? '')

describe('line heights in points', () => {
  it('writes a paragraph\'s exact and least heights in twips with their rule, and a multiple as before', async () => {
    const { document: xml } = await write([
      paragraph('Exact.', { spacing: { line: { exact: 14 } } }),
      paragraph('At least.', { spacing: { line: { atLeast: 13.5 } } }),
      paragraph('Multiple.', { spacing: { line: 1.5 } }),
    ])
    const spacings = paragraphSpacings(xml).filter(Boolean)
    expect(spacings[0]).toMatch(/w:line="280"/)
    expect(spacings[0]).toMatch(/w:lineRule="exact"/)
    expect(spacings[1]).toMatch(/w:line="270"/)
    expect(spacings[1]).toMatch(/w:lineRule="atLeast"/)
    // A multiple over the default body (itself a multiple) keeps the kit's old bytes: no rule.
    expect(spacings[2]).toMatch(/w:line="360"/)
    expect(spacings[2]).not.toMatch(/lineRule/)
  })

  it('writes the theme\'s text roles with their rule: the body in the document defaults, the others as styles', async () => {
    const pkg = await write([paragraph('Body.')], {
      theme: {
        body: { line: { exact: 13 } },
        headings: [{ line: { atLeast: 20 } }],
        caption: { line: { exact: 11 } },
        table: { line: { exact: 10 } },
      },
    })
    const styles = pkg.text('word/styles.xml')
    const defaults = /<w:docDefaults>[\s\S]*?<\/w:docDefaults>/.exec(styles)![0]
    expect(defaults).toMatch(/<w:spacing [^>]*w:line="260" w:lineRule="exact"/)
    expect(styleXml(styles, 'Heading1')).toMatch(/w:line="400" w:lineRule="atLeast"/)
    expect(styleXml(styles, 'Caption')).toMatch(/w:line="220" w:lineRule="exact"/)
    expect(styleXml(styles, 'KitTable')).toMatch(/w:line="200" w:lineRule="exact"/)
  })

  it('names the rule of a multiple written over a height in points, so no reader takes the style\'s rule', async () => {
    const pkg = await write([paragraph('Single over exact.', { spacing: { line: 1 } })], { theme: { body: { line: { exact: 13 } } } })
    const styles = pkg.text('word/styles.xml')
    // Header, Footer and table text are single-spaced over Normal's exact 13pt.
    for (const id of ['Header', 'Footer', 'KitTable']) expect(styleXml(styles, id)).toMatch(/w:line="240" w:lineRule="auto"/)
    expect(paragraphSpacings(pkg.document).filter(Boolean)[0]).toMatch(/w:line="240" w:lineRule="auto"/)
  })

  it('writes a named style\'s height in points, and a style based on it that goes back to a multiple says so', async () => {
    const pkg = await write([paragraph('Label', { style: 'CoverLabel' }), paragraph('Loose', { style: 'CoverLoose' })], {
      styles: {
        paragraph: [
          { id: 'CoverLabel', text: { size: 9, line: { exact: 12 } } },
          { id: 'CoverLoose', basedOn: 'CoverLabel', text: { line: 1.2 } },
          { id: 'CoverFloor', basedOn: 'CoverLabel', text: { line: { atLeast: 16 } } },
        ],
      },
    })
    const styles = pkg.text('word/styles.xml')
    expect(styleXml(styles, 'CoverLabel')).toMatch(/w:line="240" w:lineRule="exact"/)
    expect(styleXml(styles, 'CoverLoose')).toMatch(/w:line="288" w:lineRule="auto"/)
    expect(styleXml(styles, 'CoverFloor')).toMatch(/w:line="320" w:lineRule="atLeast"/)
  })

  it('writes a table\'s own height in points on its cells\' paragraphs', async () => {
    const { document: xml } = await write([table([{ cells: ['A', 'B'] }], { line: { exact: 11 } })])
    const cells = paragraphSpacings(xml).filter((s) => s.includes('w:line'))
    expect(cells).toHaveLength(2)
    for (const s of cells) expect(s).toMatch(/w:line="220" w:lineRule="exact"/)
  })

  it('refuses a line height that is not one', async () => {
    const bad = [{ exact: 0 }, { atLeast: -2 }, { exact: 12, atLeast: 14 }, { exactly: 12 }, Number.NaN, 0, { exact: 2000 }]
    for (const line of bad) {
      await expect(write([paragraph('x', { spacing: { line: line as never } })])).rejects.toThrow(/line height/)
    }
  })
})
