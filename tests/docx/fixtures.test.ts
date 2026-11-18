import { XMLValidator } from 'fast-xml-parser'
import { describe, expect, it } from 'vitest'
import { pageBox, resolveTheme, toComplete } from '../../src/index'
import { writeDocx } from '../../src/docx/index'
import { laboratoryReport, reviewRegistry, reviewReport } from './fixtures'
import { all, paragraphTexts, readPackage } from './helpers'

describe('a laboratory-style report', () => {
  it('writes every part the laboratory application prints', async () => {
    const doc = laboratoryReport()
    const pkg = readPackage(await writeDocx(doc))
    const xml = pkg.document
    const texts = paragraphTexts(xml)

    // Header: letterhead picture and the code; footer: Page X of Y with the reference at the left.
    const header = pkg.names.filter((n) => /^word\/header\d+\.xml$/.test(n)).map((n) => pkg.text(n)).join('')
    expect(header).toContain('<w:drawing>')
    expect(header).toContain('LAB-0001 R0')
    const footer = pkg.names.filter((n) => /^word\/footer\d+\.xml$/.test(n)).map((n) => pkg.text(n)).join('')
    expect(footer).toContain('LAB-0001')
    expect(footer).toMatch(/>PAGE<\/w:instrText>[\s\S]*>NUMPAGES<\/w:instrText>/)

    expect(texts[0]).toBe('Certificate of Analysis')
    expect(texts).toContain('Re: Water samples, made-up project')
    expect(texts.find((t) => t.startsWith('Phone: '))).toMatch(/^Phone:  +$/)
    expect(pkg.text('word/_rels/document.xml.rels')).toContain('Target="mailto:client@example.com" TargetMode="External"')

    // The results table: two repeating header rows with merged cells, shaded group rows, 60 body rows, a note row.
    const tables = all(xml, /<w:tbl>([\s\S]*?)<\/w:tbl>/)
    const results = tables[0] ?? ''
    const rows = all(results, /<w:tr>([\s\S]*?)<\/w:tr>/)
    expect(rows.length).toBe(2 + 2 + 60 + 1)
    expect(rows.filter((r) => r.includes('<w:tblHeader/>')).length).toBe(2)
    expect(rows[0]).toContain('<w:gridSpan w:val="3"/>')
    expect(all(rows[0] ?? '', /(<w:vMerge w:val="restart"\/>)/).length).toBe(2)
    expect(rows.filter((r) => r.includes('w:fill="D5E8DC"')).length).toBe(2)
    expect(rows.filter((r) => r.includes('w:fill="E8F1EB"')).length).toBe(2)
    expect(rows[rows.length - 1]).toContain('<w:gridSpan w:val="5"/>')
    expect(results).toContain('<w:tr2bl ')
    // Footnote markers and chemistry are real super- and subscripts.
    expect(all(results, /(<w:vertAlign w:val="superscript"\/>)/).length).toBe(7)
    expect(all(results, /(<w:vertAlign w:val="subscript"\/>)/).length).toBeGreaterThan(10)

    expect(texts).toContain('Table 1: Results for made-up samples')
    expect(texts).toContain('Table 2: Methods')
    expect(texts).toContain('Figure 1: Made-up chart of the results')
    expect(texts).toContain('Report authorized by:')
    expect(texts).toContain('B. Analyst, Laboratory Lead')
    expect(xml).toMatch(/<w:pageBreakBefore\/>[\s\S]{0,200}Sample and analysis notes/)

    // Times New Roman body and the green title come from the theme.
    const styles = pkg.text('word/styles.xml')
    expect(styles).toMatch(/<w:docDefaults>[\s\S]*w:ascii="Times New Roman"[\s\S]*<w:sz w:val="21"\/>/)
    expect(/<w:style [^>]*w:styleId="Title"[\s\S]*?<\/w:style>/.exec(styles)?.[0]).toContain('<w:color w:val="145A46"/>')

    expect(toComplete(doc).map((item) => item.what)).toEqual(['Phone'])
  })
})

describe('a review-style report', () => {
  it('writes the cover, contents, custom blocks, landscape and 11 × 17 sections and the closing list', async () => {
    const doc = reviewReport()
    const pkg = readPackage(await writeDocx(doc, { registry: reviewRegistry }))
    const xml = pkg.document
    const texts = paragraphTexts(xml)

    const sections = all(xml, /(<w:sectPr>[\s\S]*?<\/w:sectPr>)/)
    expect(sections.length).toBe(5)
    expect(sections.map((s) => /<w:pgSz [^>]*\/>/.exec(s)?.[0])).toEqual([
      '<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>',
      '<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>',
      '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>',
      '<w:pgSz w:w="24480" w:h="15840" w:orient="landscape"/>',
      '<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>',
    ])
    expect(sections[0]).toContain('<w:vAlign w:val="center"/>')

    // Contents over the headings.
    expect(xml).toContain('<w:docPartGallery w:val="Table of Contents"/>')
    expect(all(xml, /<w:bookmarkStart w:name="(_Toc\d+)"/).length).toBeGreaterThanOrEqual(4)

    // The unfilled sentence is left out; the rest of the paragraph reads on.
    expect(texts).toContain('The proposal alters 4.2% of the landform seen from the viewpoint. The figures below are made up.')

    // The matrix became a table with one marked cell.
    expect(texts).toContain('Table 1. Made-up rating matrix')
    expect(all(xml, /(w:fill="F4C7C3")/).length).toBe(1)
    expect(texts).toContain('Objective met? ☐ Yes   ☐ No   ☐ Inconclusive')

    // The wide table is on the landscape page: the third section's body holds it at the landscape text width.
    expect(texts).toContain('Table 2. Made-up inventory across nine measures')
    expect(xml).toContain('<w:tblW w:type="dxa" w:w="12960"/>')

    // The analysis sheet fills the 11 × 17 page's text box, under its caption.
    const box = pageBox({ paper: 'tabloid', orientation: 'landscape', margins: { top: 36, bottom: 36, left: 36, right: 36 } }, resolveTheme(doc.theme))
    const extents = all(xml, /<wp:extent cx="(\d+)" cy="\d+"\/>/).map(Number)
    const sheetWidth = Math.max(...extents) / 12700
    expect(sheetWidth).toBeLessThanOrEqual(box.contentWidth)
    const heights = all(xml, /<wp:extent cx="\d+" cy="(\d+)"\/>/).map(Number)
    expect(Math.max(...heights) / 12700).toBeLessThan(box.contentHeight)
    expect(Math.max(...heights) / 12700).toBeGreaterThan(box.contentHeight * 0.85)
    expect(texts).toContain('Figure 1. Made-up analysis sheet')

    // The closing list names what is still to give, under the heading it belongs to.
    const closing = texts.slice(texts.indexOf('To complete before signing') + 1)
    expect(closing).toEqual(
      expect.arrayContaining([
        'Prepared by',
        'Summary: Visibility statement (needs visibility)',
        'Assessment: Visibility term',
        'Assessment: Objective met?',
      ]),
    )
    // The cover has empty header and footer of its own; the body sections have the running ones.
    expect(sections[0]).toContain('headerReference')
    const headers = pkg.names.filter((n) => /^word\/header\d+\.xml$/.test(n)).map((n) => pkg.text(n))
    expect(headers.some((h) => h.includes('Made-up Visual Review'))).toBe(true)
  })

  it('refuses to write the review report without its registry', async () => {
    await expect(writeDocx(reviewReport())).rejects.toThrow(/"matrix"/)
  })
})

describe('both reports', () => {
  it('write well-formed XML in every part', async () => {
    for (const bytes of [await writeDocx(laboratoryReport()), await writeDocx(reviewReport(), { registry: reviewRegistry })]) {
      const pkg = readPackage(bytes)
      for (const name of pkg.names.filter((n) => /\.(xml|rels)$/.test(n))) {
        expect([name, XMLValidator.validate(pkg.text(name))]).toEqual([name, true])
      }
    }
  })
})
