import { describe, expect, it } from 'vitest'
import {
  blank,
  bold,
  caption,
  cell,
  custom,
  defineCustomBlocks,
  document,
  field,
  fields,
  figure,
  figureGrid,
  groupRow,
  headerRow,
  heading,
  image,
  inlineImage,
  item,
  lineBreak,
  link,
  list,
  noteRow,
  orderedList,
  pageBreak,
  pageCount,
  pageNumber,
  paragraph,
  row,
  section,
  sectionPageCount,
  sentence,
  sentences,
  signature,
  spacer,
  spacerRow,
  sub,
  sup,
  tab,
  table,
  text,
  toc,
  UnknownCustomBlockError,
  type Block,
  type Document,
  type Section,
} from '../../src/index'
import { writeDocx, type WriteDocxOptions } from '../../src/docx/index'
import { all, makePng, paragraphTexts, readPackage } from './helpers'
import { hex } from '../../src/docx/units'

const write = async (body: readonly Block[], extra: Omit<Section, 'body'> = {}, options?: WriteDocxOptions) =>
  readPackage(await writeDocx(document([section(body, extra)]), options))
const writeDoc = async (doc: Document, options?: WriteDocxOptions) => readPackage(await writeDocx(doc, options))
const body = (xml: string) => xml.slice(xml.indexOf('<w:body>'), xml.indexOf('</w:body>'))

describe('inline text', () => {
  it('writes real run properties', async () => {
    const pkg = await write([
      paragraph([
        bold('B'),
        text('I', { italic: true }),
        text('U', { underline: true }),
        text('S', { strike: true }),
        'H',
        sub('2'),
        'O',
        sup('a'),
        text('C', { color: '#035642', size: 13, font: 'Georgia', highlight: 'yellow', smallCaps: true }),
      ]),
    ])
    const xml = body(pkg.document)
    expect(xml).toContain('<w:b/>')
    expect(xml).toContain('<w:i/>')
    expect(xml).toContain('<w:u w:val="single"/>')
    expect(xml).toContain('<w:strike/>')
    expect(xml).toContain('<w:vertAlign w:val="subscript"/>')
    expect(xml).toContain('<w:vertAlign w:val="superscript"/>')
    expect(xml).toContain('<w:color w:val="035642"/>')
    expect(xml).toContain('<w:sz w:val="26"/>')
    expect(xml).toMatch(/<w:rFonts w:ascii="Georgia"/)
    expect(xml).toContain('<w:highlight w:val="yellow"/>')
    expect(xml).toContain('<w:smallCaps/>')
  })

  it('writes line breaks and tabs, from runs and from the characters', async () => {
    const xml = body((await write([paragraph(['a', lineBreak(), 'b', tab(), 'c\nd\te'])])).document)
    expect(all(xml, /(<w:br\/>)/).length).toBe(2)
    expect(all(xml, /(<w:tab\/>)/).length).toBe(2)
    expect(paragraphTexts(xml)).toEqual(['abcde'])
  })

  it('writes a link as a hyperlink with an external relationship', async () => {
    const pkg = await write([paragraph(['See ', link('the site', 'https://example.org/a?b=1&c=2')])])
    const id = /<w:hyperlink [^>]*r:id="([^"]+)"/.exec(pkg.document)?.[1]
    expect(id).toBeDefined()
    const rels = pkg.text('word/_rels/document.xml.rels')
    expect(rels).toContain(`Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/a?b=1&amp;c=2" TargetMode="External"`)
  })

  it('prints a blank as an underlined rule and leaves out an unfilled sentence', async () => {
    const xml = body(
      (
        await write([
          paragraph(['Signed ', blank('initials')]),
          paragraph(sentences(sentence('Kept {a}.', { a: 1 }), sentence('Gone {b}.', { b: null }), 'End.')),
          paragraph([sentence('Only {gone}.', {})]),
        ])
      ).document,
    )
    expect(xml).toMatch(/<w:u w:val="single"\/><\/w:rPr><w:t xml:space="preserve">( ){4,}<\/w:t>/)
    expect(paragraphTexts(xml).filter(Boolean).slice(1)).toEqual(['Kept 1. End.'])
    expect(xml).not.toContain('Only')
  })

  it('marks text by its source when the writer is told how', async () => {
    const xml = body(
      (
        await write(
          [paragraph([text('given', { source: 'reviewer' }), ' plain']), fields([field('Area', '3 ha', { source: 'reviewer' })])],
          {},
          { sourceStyles: { reviewer: { highlight: 'green' } } },
        )
      ).document,
    )
    expect(all(xml, /(<w:highlight w:val="green"\/>)/).length).toBe(2)
  })
})

describe('paragraphs, headings and lists', () => {
  it('writes Word heading styles and themed heading definitions', async () => {
    const doc = document([section([heading(1, 'One'), heading(3, 'Three')])], {
      theme: { headings: [{ color: '035642', size: 18 }] },
    })
    const pkg = await writeDoc(doc)
    expect(all(pkg.document, /<w:pStyle w:val="(Heading\d)"\/>/)).toEqual(['Heading1', 'Heading3'])
    const styles = pkg.text('word/styles.xml')
    const h1 = /<w:style [^>]*w:styleId="Heading1"[\s\S]*?<\/w:style>/.exec(styles)?.[0] ?? ''
    expect(h1).toContain('<w:color w:val="035642"/>')
    expect(h1).toContain('<w:sz w:val="36"/>')
    expect(h1).toContain('<w:outlineLvl w:val="0"/>')
  })

  it('writes alignment, spacing, indent, keep-with-next and a page break before', async () => {
    const xml = body(
      (
        await write([
          paragraph('x', {
            align: 'justify',
            spacing: { before: 6, after: 3, line: 1.5 },
            indent: { left: 36, hanging: 18 },
            keepWithNext: true,
            pageBreakBefore: true,
          }),
        ])
      ).document,
    )
    expect(xml).toContain('<w:jc w:val="both"/>')
    expect(xml).toContain('<w:keepNext/>')
    expect(xml).toContain('<w:pageBreakBefore/>')
    expect(xml).toMatch(/<w:spacing w:after="60" w:before="120" w:line="360"\/>/)
    expect(xml).toMatch(/<w:ind w:left="720" w:hanging="360"\/>/)
  })

  it('moves the block after a page break onto a new page', async () => {
    const xml = body((await write([paragraph('a'), pageBreak(), paragraph('b')])).document)
    expect(xml).toMatch(/<w:pageBreakBefore\/><\/w:pPr><w:r><w:t xml:space="preserve">b<\/w:t>/)
  })

  it('writes real numbered and bulleted lists, nested, each ordered list counting from its own start', async () => {
    const pkg = await write([
      list(['dot', item('dot2', ['inner'])]),
      orderedList(['one', 'two']),
      orderedList(['again'], { start: 5 }),
      orderedList([item('outer', [item('lettered')], false)]),
    ])
    const numIds = all(pkg.document, /<w:numId w:val="(\d+)"\/>/)
    const levels = all(pkg.document, /<w:ilvl w:val="(\d+)"\/>/)
    expect(levels).toEqual(['0', '0', '1', '0', '0', '0', '0', '1'])
    // "one" and "two" share a list; "again" and the last list are lists of their own.
    expect(numIds[3]).toBe(numIds[4])
    expect(new Set([numIds[3], numIds[5], numIds[6]]).size).toBe(3)
    const numbering = pkg.text('word/numbering.xml')
    expect(numbering).toContain('<w:numFmt w:val="bullet"/>')
    expect(numbering).toContain('<w:numFmt w:val="decimal"/>')
    expect(numbering).toContain('<w:start w:val="5"/>')
  })

  it('writes a spacer as an exact-height empty paragraph', async () => {
    expect(body((await write([spacer(18)])).document)).toContain('<w:spacing w:after="0" w:before="0" w:line="360" w:lineRule="exact"/>')
  })
})

describe('fields', () => {
  it('writes label and value lines, a null value as a blank rule', async () => {
    const xml = body((await write([fields([field('Date', '1 June'), field('Inspector', null)])])).document)
    expect(paragraphTexts(xml)[0]).toBe('Date: 1 June')
    expect(paragraphTexts(xml)[1]).toMatch(/^Inspector:  +$/)
    expect(xml).toMatch(/<w:keepNext\/>/)
  })

  it('writes a two-column table layout', async () => {
    const xml = body((await write([fields([field('Date', '1 June'), field('Inspector', null)], { layout: 'table', labelWidth: 25 })])).document)
    expect(all(xml, /<w:gridCol w:w="(\d+)"\/>/)).toEqual(['2340', '7020'])
    expect(xml).toContain('<w:shd w:fill="F5F7FA"')
  })
})

describe('tables', () => {
  const results = () =>
    table(
      [
        headerRow([cell('Parameter', { rowSpan: 2 }), cell('Samples', { colSpan: 2 })]),
        headerRow(['S-1', 'S-2']),
        groupRow([cell('Group A', { colSpan: 3 })]),
        row(['pH', '7.1', cell('', { diagonal: 'up' })]),
        row([['Nitrate', sup('a')], '0.4', null]),
        noteRow('a Made-up footnote.'),
        spacerRow(),
      ],
      { columnWidths: [2, 1, 1], caption: caption('Made-up results'), note: 'All values invented.', cantSplit: true },
    )

  it('repeats the leading header rows and shades the group row', async () => {
    const xml = body((await write([results()])).document)
    const rows = all(xml, /<w:tr>([\s\S]*?)<\/w:tr>/)
    expect(rows.map((r) => r.includes('<w:tblHeader/>'))).toEqual([true, true, false, false, false, false, false])
    expect(rows[2]).toContain('<w:shd w:fill="F0F2F5"')
    expect(rows.every((r) => r.includes('<w:cantSplit/>'))).toBe(true)
  })

  it('merges across columns and rows', async () => {
    const rows = all(body((await write([results()])).document), /<w:tr>([\s\S]*?)<\/w:tr>/)
    expect(rows[0]).toContain('<w:vMerge w:val="restart"/>')
    expect(rows[0]).toContain('<w:gridSpan w:val="2"/>')
    expect(rows[1]).toMatch(/<w:vMerge(?: w:val="continue")?\/>/)
    expect(all(rows[1] ?? '', /(<w:tc>)/).length).toBe(3)
    expect(rows[5]).toContain('<w:gridSpan w:val="3"/>')
  })

  it('sizes columns from their shares of the text width', async () => {
    const xml = body((await write([results()])).document)
    expect(all(xml, /<w:gridCol w:w="(\d+)"\/>/)).toEqual(['4680', '2340', '2340'])
    expect(xml).toContain('<w:tblW w:type="dxa" w:w="9360"/>')
  })

  it('writes the struck cell, the footnote marker, the spacer, the caption and the note', async () => {
    const xml = body((await write([results()])).document)
    expect(xml).toMatch(/<w:tr2bl w:val="single"/)
    expect(xml).toMatch(/<w:vertAlign w:val="superscript"\/><\/w:rPr><w:t xml:space="preserve">a<\/w:t>/)
    expect(xml).toMatch(/<w:trHeight w:val="120" w:hRule="exact"\/>/)
    expect(xml).toContain('<w:fldSimple w:instr="SEQ Table \\* ARABIC"><w:r><w:t xml:space="preserve">1</w:t>')
    expect(paragraphTexts(xml)[0]).toBe('Table 1: Made-up results')
    expect(xml).toMatch(/<w:pStyle w:val="Caption"\/><w:keepNext\/>/)
    expect(xml).toMatch(/<w:pStyle w:val="KitNote"\/><\/w:pPr><w:r><w:t xml:space="preserve">All values invented./)
  })

  it('prints a null cell empty, or as a rule when asked', async () => {
    const t = table([row(['a', null])])
    const empty = body((await write([t])).document)
    expect(empty).not.toContain('<w:u w:val="single"/>')
    const rule = body((await write([t], {}, { nullCells: 'rule' })).document)
    expect(rule).toContain('<w:u w:val="single"/>')
  })

  it('holds blocks in a cell, nested tables included', async () => {
    const xml = body(
      (await write([table([row([cell(null, { blocks: [paragraph('inside'), table([row(['deep'])])] }), 'b'])])])).document,
    )
    expect(all(xml, /(<w:tbl>)/).length).toBe(2)
    expect(paragraphTexts(xml)).toContain('deep')
  })
})

describe('figures', () => {
  const picture = image(makePng(400, 200))

  it('embeds the picture at its fitted size with a numbered caption', async () => {
    const pkg = await write([figure(picture, { fit: { width: 150 }, caption: caption('A made-up chart') })])
    // 150 pt wide, 75 pt tall: 200 by 100 pixels at 96 to the inch, 9525 EMU each.
    expect(pkg.document).toContain('<wp:extent cx="1905000" cy="952500"/>')
    expect(pkg.names.some((name) => /^word\/media\/.+\.png$/.test(name))).toBe(true)
    expect(paragraphTexts(body(pkg.document))).toContain('Figure 1: A made-up chart')
    expect(pkg.document).toContain('w:instr="SEQ Figure \\* ARABIC"')
  })

  it('keeps a small picture its own size unless told to enlarge', async () => {
    const small = image(makePng(40, 20))
    expect((await write([figure(small, { fit: { width: 300, height: 300 } })])).document).toContain('cx="381000" cy="190500"')
    const big = await write([figure(small, { fit: { width: 300, height: 300, upscale: true } })])
    expect(big.document).toContain('cx="3810000" cy="1905000"')
  })

  it('prints a placeholder box when the picture is missing', async () => {
    const xml = body((await write([figure(null, { placeholder: 'Photograph to come', caption: caption('Site view') })])).document)
    expect(xml).toContain('<w:tbl>')
    expect(xml).toMatch(/<w:top w:val="dashed"/)
    expect(paragraphTexts(xml)).toEqual(['Photograph to come', 'Figure 1: Site view'])
  })

  it('lays a grid of pictures in columns with labels and one caption', async () => {
    const xml = body(
      (
        await write([
          figureGrid(
            [
              { image: picture, label: 'Before' },
              { image: picture, label: 'After' },
              { image: null, label: 'Later', placeholder: 'Not yet taken' },
            ],
            2,
            { caption: caption('Made-up comparison') },
          ),
        ])
      ).document,
    )
    expect(all(xml, /(<w:drawing>)/).length).toBe(2)
    expect(all(xml, /<w:gridCol w:w="(\d+)"\/>/).slice(0, 2)).toEqual(['4680', '4680'])
    const texts = paragraphTexts(xml)
    expect(texts).toEqual(expect.arrayContaining(['Before', 'After', 'Later', 'Not yet taken', 'Figure 1: Made-up comparison']))
  })

  it('reads a picture’s size from a JPEG header', async () => {
    const jpeg = Uint8Array.from(
      atob(
        '/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////wgALCAABAAEBAREA/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPxA=',
      ),
      (c) => c.charCodeAt(0),
    )
    const pkg = await write([figure(image(jpeg), { fit: { width: 72, height: 72, upscale: true } })])
    expect(pkg.names.some((name) => /^word\/media\/.+\.jpe?g$/.test(name))).toBe(true)
    expect(pkg.document).toContain('cx="914400" cy="914400"')
  })
})

describe('signature', () => {
  it('writes the label, room to sign, the rule with the name, and blank lines', async () => {
    const xml = body(
      (await write([signature({ label: 'Approved by:', name: 'A. Person, Analyst', lines: [field('Date', null)], align: 'center' })])).document,
    )
    expect(xml).toMatch(/<w:pBdr><w:top w:val="single"/)
    expect(xml).toContain('<w:spacing w:after="0" w:before="0" w:line="720" w:lineRule="exact"/>')
    const texts = paragraphTexts(xml)
    expect(texts).toContain('Approved by:')
    expect(texts).toContain('A. Person, Analyst')
    expect(texts.find((t) => t.startsWith('Date: '))).toMatch(/^Date:  +$/)
  })

  it('embeds a signature picture when given', async () => {
    const xml = body((await write([signature({ image: image(makePng(300, 100)), name: null })])).document)
    expect(xml).toContain('<w:drawing>')
  })
})

describe('sections, headers and footers', () => {
  it('gives each section its own paper and orientation', async () => {
    const pkg = await writeDoc(
      document([
        section([paragraph('portrait letter')]),
        section([paragraph('landscape legal')], { page: { paper: 'legal', orientation: 'landscape' } }),
        section([paragraph('11 by 17')], { page: { paper: 'tabloid', orientation: 'landscape', margins: { left: 36, right: 36 } } }),
        section([paragraph('a4')], { page: { paper: 'a4' } }),
      ]),
    )
    const sizes = all(pkg.document, /(<w:pgSz [^>]*\/>)/)
    expect(sizes).toEqual([
      '<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>',
      '<w:pgSz w:w="20160" w:h="12240" w:orient="landscape"/>',
      '<w:pgSz w:w="24480" w:h="15840" w:orient="landscape"/>',
      '<w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/>',
    ])
    expect(pkg.document).toMatch(/<w:pgMar w:top="1440" w:right="720" w:bottom="1440" w:left="720"/)
  })

  it('writes PAGE and NUMPAGES fields, a header picture and a first-page header', async () => {
    const logo = image(makePng(120, 40))
    const pkg = await write([paragraph('body')], {
      header: [paragraph([inlineImage(logo, 90, 30, 'Made-up logo')], { align: 'center' })],
      footer: [
        paragraph(['REF-000', tab(), 'Page ', pageNumber(), ' of ', pageCount(), ' (', sectionPageCount(), ')'], {
          tabs: [{ at: 'end', align: 'right' }],
        }),
      ],
      firstPage: { header: [paragraph('Cover header')] },
      page: { verticalAlign: 'center' },
      pageNumbers: { start: 3, format: 'lowerRoman' },
    })
    const footerName = pkg.names.find((name) => /^word\/footer\d+\.xml$/.test(name) && pkg.text(name).includes('PAGE'))
    expect(footerName).toBeDefined()
    const footer = pkg.text(footerName!)
    expect(footer).toMatch(/<w:instrText xml:space="preserve">PAGE<\/w:instrText>/)
    expect(footer).toMatch(/<w:instrText xml:space="preserve">NUMPAGES<\/w:instrText>/)
    expect(footer).toMatch(/<w:instrText xml:space="preserve">SECTIONPAGES<\/w:instrText>/)
    expect(footer).toContain('<w:tab w:val="right" w:pos="9360"/>')
    const headers = pkg.names.filter((name) => /^word\/header\d+\.xml$/.test(name)).map((name) => pkg.text(name))
    expect(headers.some((h) => h.includes('<w:drawing>') && h.includes('descr="Made-up logo"'))).toBe(true)
    expect(headers.some((h) => h.includes('Cover header'))).toBe(true)
    expect(pkg.document).toContain('<w:headerReference w:type="first"')
    expect(pkg.document).toContain('<w:titlePg/>')
    expect(pkg.document).toContain('<w:vAlign w:val="center"/>')
    expect(pkg.document).toContain('<w:pgNumType w:start="3" w:fmt="lowerRoman"/>')
  })

  it('leaves a header out to inherit, and writes an empty list as no header', async () => {
    const pkg = await writeDoc(
      document([
        section([paragraph('a')], { header: [paragraph('Running head')] }),
        section([paragraph('b')]),
        section([paragraph('c')], { header: [] }),
      ]),
    )
    const sects = all(pkg.document, /(<w:sectPr>[\s\S]*?<\/w:sectPr>)/)
    expect(sects.map((s) => s.includes('headerReference'))).toEqual([true, false, true])
  })

  it('writes the metadata', async () => {
    const pkg = await writeDoc(
      document([section([paragraph('x')])], {
        metadata: { title: 'Made-up report', author: 'Test Writer', subject: 'Testing', keywords: ['a', 'b'], created: '2026-03-04T05:06:07Z' },
      }),
    )
    const core = pkg.text('docProps/core.xml')
    expect(core).toContain('<dc:title>Made-up report</dc:title>')
    expect(core).toContain('<dc:creator>Test Writer</dc:creator>')
    expect(core).toContain('<cp:keywords>a, b</cp:keywords>')
    expect(core).toContain('>2026-03-04T05:06:07Z</dcterms:created>')
    expect(core).toContain('>2026-03-04T05:06:07Z</dcterms:modified>')
  })
})

describe('table of contents', () => {
  it('writes a TOC field over the heading levels, filled from the headings', async () => {
    const pkg = await write([toc({ title: 'Contents', levels: 2 }), heading(1, 'Alpha'), heading(2, 'Beta'), heading(3, 'Gamma')])
    expect(pkg.document).toContain('<w:docPartGallery w:val="Table of Contents"/>')
    expect(pkg.document).toContain('TOC \\h \\o &quot;1-2&quot;')
    expect(pkg.document).toMatch(/<w:pStyle w:val="KitTocTitle"\/>/)
    expect(all(pkg.document, /<w:bookmarkStart w:name="(_Toc\d+)"/)).toEqual(['_Toc1', '_Toc2'])
  })
})

describe('custom blocks in the writer', () => {
  it('fails with the kind’s name when it is not registered', async () => {
    const doc = document([section([custom('ratingMatrix', {})])])
    await expect(writeDocx(doc)).rejects.toBeInstanceOf(UnknownCustomBlockError)
    await expect(writeDocx(doc)).rejects.toThrow(/"ratingMatrix"/)
  })

  it('writes what a registered kind expands to', async () => {
    const registry = defineCustomBlocks({ note: (data: string) => [paragraph(`Note: ${data}`)] })
    const pkg = await writeDoc(document([section([custom('note', 'made up')])]), { registry })
    expect(paragraphTexts(body(pkg.document))[0]).toBe('Note: made up')
  })
})

describe('determinism', () => {
  const sample = () =>
    document(
      [
        section(
          [
            toc(),
            heading(1, 'One'),
            paragraph(['A ', link('link', 'https://example.org'), '.']),
            figure(image(makePng(20, 10)), { caption: caption('Pic') }),
            heading(1, 'Two'),
          ],
          { footer: [paragraph(['see ', link('footer link', 'https://example.org/f')])] },
        ),
      ],
      { metadata: { created: '2026-01-02T03:04:05Z' } },
    )

  it('writes the same bytes for the same document, however many documents came before', async () => {
    const first = await writeDocx(sample())
    await writeDocx(sample())
    const third = await writeDocx(sample())
    expect(third).toEqual(first)
    const pkg = readPackage(first)
    expect(pkg.document).toContain('<wp:docPr id="1"')
    expect(pkg.document).toMatch(/r:id="rIdL1"/)
  })

  it('writes the same document.xml even without a creation time', async () => {
    const doc = () => document([section([heading(1, 'H'), paragraph(link('x', 'https://example.org'))])])
    const a = readPackage(await writeDocx(doc())).document
    const b = readPackage(await writeDocx(doc())).document
    expect(b).toBe(a)
  })
})

describe('colours', () => {
  it('takes six or three hex digits, with or without #, in any case', () => {
    expect(hex('#035642')).toBe('035642')
    expect(hex('035642')).toBe('035642')
    expect(hex('#aBcDeF')).toBe('ABCDEF')
    expect(hex('#0a3')).toBe('00AA33')
    expect(hex('fC0')).toBe('FFCC00')
    expect(hex('  #abc ')).toBe('AABBCC')
  })

  it('throws on anything else, naming the value', () => {
    for (const bad of ['red', '#12345', '#1234567', '#ggg', '', 'rgb(1,2,3)', '##abc']) {
      expect(() => hex(bad)).toThrow(`document-kit: ${JSON.stringify(bad)} is not a hex colour`)
    }
  })

  it('writes a short colour in a document as Word’s six digits', async () => {
    const xml = body((await write([paragraph([{ type: 'text', text: 'x', color: '#c00' }], { shading: 'eee' })])).document)
    expect(xml).toContain('<w:color w:val="CC0000"/>')
    expect(xml).toContain('w:fill="EEEEEE"')
  })
})

describe('a rule set off from its text, and a cell with its own padding', () => {
  it('writes a paragraph border’s distance from the text', async () => {
    const pkg = await write([
      paragraph('Made-up letterhead line', { borders: { bottom: { style: 'single', width: 0.75, color: '#0c5a36', space: 6 } } }),
    ])
    expect(body(pkg.document)).toMatch(/<w:bottom w:val="single" w:color="0C5A36" w:sz="6" w:space="6"\/>/)
  })

  it('writes a cell’s own padding over the table’s, and only for that cell', async () => {
    const pkg = await write([
      table([row(['Plain', { content: 'Padded note', padding: { top: 3.5, bottom: 3.5 } }])], { cellPadding: { top: 5, bottom: 5, left: 5.5, right: 5.5 } }),
    ])
    const cells = body(pkg.document).match(/<w:tc>[\s\S]*?<\/w:tc>/g)!
    expect(cells[0]).not.toContain('<w:tcMar>')
    expect(cells[1]).toContain('<w:tcMar><w:top w:type="dxa" w:w="70"/><w:left w:type="dxa" w:w="110"/><w:bottom w:type="dxa" w:w="70"/><w:right w:type="dxa" w:w="110"/></w:tcMar>')
  })
})
