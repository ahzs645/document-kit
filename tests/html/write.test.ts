// The block model as a print-styled HTML page: the same document, theme and named styles as the Word file, pages as
// sheets of the section's paper. Made-up text throughout.
import { describe, expect, it } from 'vitest'

import {
  caption,
  custom,
  defineCustomBlocks,
  document,
  field,
  fields,
  floatingImage,
  frame,
  headerRow,
  heading,
  list,
  pageCount,
  pageNumber,
  paragraph,
  pinned,
  row,
  section,
  table,
  text,
  toc,
  type Document,
} from '../../src/model'
import { writeHtml } from '../../src/html'
import { makePng } from '../docx/helpers'
import { otf, ttf, withFsType } from '../fonts'

const sheetsOf = (html: string) => [...html.matchAll(/<section class="dk-sheet[^"]*"[\s\S]*?<\/section>/g)].map((m) => m[0])

describe('writeHtml', () => {
  it('writes one sheet a page, each section on its own paper, with its header, footer and page numbers', () => {
    const html = writeHtml(
      document(
        [
          section([paragraph('Cover.')], { page: { paper: 'a4' } }),
          section([paragraph('One.'), { type: 'pageBreak' }, paragraph('Two.', { pageBreakBefore: true })], {
            page: { paper: 'letter', orientation: 'landscape' },
            header: [paragraph('Running head')],
            footer: [paragraph(['Page ', pageNumber(), ' of ', pageCount()], { tabs: [{ at: 'end', align: 'right' }] })],
            pageNumbers: { start: 1, format: 'lowerRoman' },
          }),
        ],
        { metadata: { title: 'Made-up <title>', author: 'An Author', custom: { 'app.id': 'pub-7', revision: 3 } } },
      ),
    )
    expect(html.startsWith('<!doctype html>')).toBe(true)
    expect(html).toContain('<title>Made-up &lt;title&gt;</title>')
    expect(html).toContain('<meta name="author" content="An Author">')
    expect(html).toContain('<meta name="app.id" content="pub-7">')
    expect(html).toContain('<meta name="revision" content="3">')
    expect(html).toMatch(/@page dk-s0\{size:595.3pt 841.9pt;margin:0\}/)
    expect(html).toMatch(/@page dk-s1\{size:792pt 612pt;margin:0\}/)
    const sheets = sheetsOf(html)
    expect(sheets).toHaveLength(3)
    expect(sheets[0]).not.toContain('Running head')
    expect(sheets[1]).toContain('Running head')
    expect(sheets.map((s) => /data-page="([^"]+)"/.exec(s)?.[1])).toEqual(['1', 'i', 'ii'])
    expect(sheets[2]).toMatch(/data-field="page">ii</)
    expect(sheets[2]).toMatch(/data-field="pages">3</)
    // Nothing in it reaches out.
    expect(html).not.toMatch(/\b(src|href)="(?!data:|#)/)
  })

  it('writes the theme and the document\'s named styles as classes, tracking included', () => {
    const html = writeHtml(
      document([section([paragraph([text('Unit', { characterStyle: 'Accent' })], { style: 'AffiliationDepartment' }), heading(1, 'Title')])], {
        theme: { body: { font: 'Brand Sans', size: 10, color: '222222' }, headings: [{ characterSpacing: 1 }] },
        styles: {
          paragraph: [{ id: 'AffiliationDepartment', text: { size: 12, bold: true, characterSpacing: 0.6, after: 2 } }],
          character: [{ id: 'Accent', text: { color: 'b45309', characterSpacing: 1.5 } }],
        },
      }),
    )
    const rule = (cls: string) => new RegExp(`\\.${cls}\\{([^}]*)\\}`).exec(html)?.[1] ?? ''
    expect(rule('dk-ps-AffiliationDepartment')).toContain('letter-spacing:0.6pt')
    expect(rule('dk-ps-AffiliationDepartment')).toContain('font-size:12pt')
    expect(rule('dk-ps-AffiliationDepartment')).toContain('font-weight:700')
    expect(rule('dk-ps-AffiliationDepartment')).toContain('font-family:"Brand Sans", sans-serif')
    expect(rule('dk-ps-Heading1')).toContain('letter-spacing:1pt')
    expect(rule('dk-cs-Accent')).toContain('letter-spacing:1.5pt')
    expect(html).toMatch(/<p class="dk-p dk-ps-AffiliationDepartment"><span class="dk-cs-Accent">Unit<\/span><\/p>/)
    expect(() => writeHtml(document([section([paragraph('x', { style: 'Nope' })])]))).toThrow(/"Nope"/)
  })

  it('carries the fonts the licence allows as @font-face, TrueType and OpenType', () => {
    const html = writeHtml(document([section([paragraph('x')])]), {
      fonts: [
        { family: 'Brand Sans', bytes: ttf() },
        { family: 'Brand Sans', weight: 700, bytes: ttf('Bold') },
        { family: 'Code', bytes: otf() },
        { family: 'Locked', bytes: withFsType(ttf(), 2) },
      ],
    })
    const faces = [...html.matchAll(/@font-face\{font-family:"([^"]+)";font-weight:(\d+);[^}]*?data:font\/(\w+);base64/g)].map((m) => [m[1], m[2], m[3]])
    expect(faces).toEqual([
      ['Brand Sans', '400', 'ttf'],
      ['Brand Sans', '700', 'ttf'],
      ['Code', '400', 'otf'],
    ])
  })

  it('places frames and floating pictures on the sheet, behind or in front', () => {
    const art = { bytes: makePng(4, 4) }
    const sheet = sheetsOf(
      writeHtml(
        document([
          section(
            [
              floatingImage(art, 'page', { alt: 'Artwork' }),
              floatingImage(art, { x: 10, y: 20, width: 30, height: 40 }, { relativeTo: 'margin', behind: false }),
              frame([paragraph('Framed')], { x: 54, y: 120, width: 300 }),
              frame([paragraph('In margin')], { x: 0, y: 0, width: 100, height: 50 }, { relativeTo: 'margin', heightRule: 'exact' }),
            ],
            { page: { paper: 'letter', margins: { top: 72, left: 72, right: 72, bottom: 72 } } },
          ),
        ]),
      ),
    )[0]!
    expect(sheet).toMatch(/<img class="dk-float" src="data:image\/png;base64,[^"]+" alt="Artwork" style="left:0pt;top:0pt;width:612pt;height:792pt;z-index:-1">/)
    expect(sheet).toMatch(/style="left:82pt;top:92pt;width:30pt;height:40pt;z-index:2"/)
    expect(sheet).toMatch(/<div class="dk-frame" style="left:54pt;top:120pt;width:300pt;z-index:1"><p[^>]*>Framed<\/p><\/div>/)
    expect(sheet).toMatch(/left:72pt;top:72pt;width:100pt;z-index:1;height:50pt;overflow:hidden/)
  })

  it('reads the document as the Word writer does: custom blocks, numbered captions, notes, contents, pinned blocks', () => {
    const registry = defineCustomBlocks({ note: (data: { text: string }) => [paragraph(data.text, { style: 'note' })] })
    const doc: Document = document(
      [
        section([
          toc({ title: 'Contents' }),
          heading(1, 'Results'),
          paragraph(['A finding', { type: 'footnote', id: 'src' }, ' at www.example.org.']),
          table([headerRow(['Item', 'Value']), row([{ content: 'Merged', colSpan: 2 }])], { caption: caption('Readings') }),
          list(['First', { content: 'Second', items: ['Inner'] }], { ordered: true }),
          fields([field('Checked by', null)]),
          custom('note', { text: 'Expanded note.' }),
          pinned([paragraph('Pinned at the foot.')]),
        ]),
      ],
      { footnotes: { src: ['A made-up source.'] }, textRules: [{ builtIn: 'links' }] },
    )
    const html = writeHtml(doc, { registry })
    expect(html).toContain('Table 1: Readings')
    expect(html).toMatch(/<thead><tr[^>]*><th scope="col"/)
    expect(html).toMatch(/<td colspan="2"/)
    expect(html).toMatch(/<sup class="dk-noteref"><a href="#dk-fn-1">1<\/a><\/sup>/)
    expect(html).toMatch(/<p id="dk-fn-1" class="dk-p dk-ps-KitNote"[^>]*><sup>1<\/sup> A made-up source.<\/p>/)
    expect(html).toMatch(/<a href="https:\/\/www.example.org">www.example.org<\/a>/)
    expect(html).toMatch(/<a href="#dk-h\d+"[^>]*>Results<\/a>.*?<span>1<\/span>/)
    expect(html).toMatch(/<ol class="dk-list" style="list-style-type:decimal">.*<ol class="dk-list" style="list-style-type:lower-alpha">/)
    expect(html).toContain('Expanded note.')
    expect(html).toMatch(/<div class="dk-foot">.*Pinned at the foot\./)
    expect(html).toMatch(/class="dk-blank"/)
    expect(() => writeHtml(document([section([custom('unknown', {})])]))).toThrow(/unknown/)
  })

  it('runs with no Buffer in the global scope', () => {
    const saved = globalThis.Buffer
    let html: string
    try {
      // @ts-expect-error removed to stand in for a browser
      delete globalThis.Buffer
      html = writeHtml(document([section([floatingImage({ bytes: makePng(2, 2) }, 'page')])]), { fonts: [{ family: 'F', bytes: ttf() }] })
    } finally {
      globalThis.Buffer = saved
    }
    expect(html).toContain('data:font/ttf;base64,AAEAAA')
    expect(html).toContain('data:image/png;base64,iVBORw0KGgo')
  })
})

/** Every start tag in the page, its attributes read as a browser reads them (double-quoted values, references decoded). */
function startTags(html: string): Array<{ name: string; attributes: Record<string, string>; whole: boolean }> {
  const decode = (value: string) =>
    value.replace(/&(quot|amp|lt|gt|#39);/g, (_, name: string) => ({ quot: '"', amp: '&', lt: '<', gt: '>', '#39': "'" })[name]!)
  return [...html.matchAll(/<([a-z][a-z0-9]*)((?:\s[^\s=>]+(?:="[^"]*")?)*)(\s*[^>]*)>/g)].map((m) => ({
    name: m[1]!,
    attributes: Object.fromEntries([...m[2]!.matchAll(/\s([^\s=>]+)(?:="([^"]*)")?/g)].map((a) => [a[1]!, decode(a[2] ?? '')])),
    // Nothing left over between the last attribute and the tag's end: no attribute ended early.
    whole: m[3]!.trim() === '',
  }))
}

describe('attribute values', () => {
  const family = 'Brand "Display" Sans'

  it('escapes a font stack\'s quoted names in a style attribute, so nothing after the font is lost', () => {
    const html = writeHtml(
      document([section([table([headerRow(['Header text']), row(['Body'])]), paragraph([text('Run in the face', { font: family, color: '#336699' })])])], {
        theme: { table: { headerText: { font: family, color: 'aa3311', bold: true } } },
      }),
    )
    const tags = startTags(html.slice(html.indexOf('<body')))
    expect(tags.every((tag) => tag.whole)).toBe(true)
    const header = tags.find((tag) => tag.name === 'p' && tags[tags.indexOf(tag) - 1]?.name === 'th')!
    expect(header.attributes.style).toContain(`font-family:"Brand \\22 Display\\22  Sans", sans-serif`)
    expect(header.attributes.style).toContain('color:#aa3311')
    expect(header.attributes.style).toContain('font-weight:700')
    const runs = tags.filter((tag) => tag.name === 'span' && tag.attributes.style?.includes('Display'))
    // The header cell's text and the paragraph's run, each with what follows its font.
    expect(runs.map((tag) => tag.attributes.style)).toEqual([
      'font-family:"Brand \\22 Display\\22  Sans", sans-serif;color:#aa3311;font-weight:700',
      'font-family:"Brand \\22 Display\\22  Sans", sans-serif;color:#336699',
    ])
  })

  it('keeps the style sheet whole whatever a font is named', () => {
    const html = writeHtml(document([section([paragraph('x')])], { theme: { body: { font: 'Odd</style><b>Face' } } }))
    expect(html.match(/<\/style>/g)).toHaveLength(1)
    expect(html).not.toContain('<b>')
  })
})

describe('line heights in points', () => {
  it('writes an exact height in points and a least height as the larger of it and a single line', () => {
    const html = writeHtml(
      document(
        [
          section([
            paragraph('Exact.', { spacing: { line: { exact: 14 } } }),
            paragraph('At least.', { spacing: { line: { atLeast: 15 } } }),
            paragraph('Multiple.', { spacing: { line: 1.5 } }),
            paragraph('Styled.', { style: 'CoverLine' }),
            table([row(['Cell'])], { line: { exact: 10 } }),
          ]),
        ],
        {
          theme: { body: { line: { exact: 13 } }, headings: [{ line: { atLeast: 20 } }] },
          styles: { paragraph: [{ id: 'CoverLine', text: { line: { exact: 12 } } }] },
        },
      ),
    )
    expect(html).toMatch(/<p class="dk-p dk-ps-Normal" style="line-height:14pt">Exact\./)
    expect(html).toMatch(/<p class="dk-p dk-ps-Normal" style="line-height:max\(15pt,1\.2em\)">At least\./)
    expect(html).toMatch(/<p class="dk-p dk-ps-Normal" style="line-height:1\.8">Multiple\./)
    expect(html).toMatch(/\.dk-sheet\{[^}]*line-height:13pt\}/)
    expect(html).toMatch(/\.dk-ps-Heading1\{[^}]*line-height:max\(20pt,1\.2em\)/)
    expect(html).toMatch(/\.dk-ps-CoverLine\{[^}]*line-height:12pt/)
    expect(html).toMatch(/<td[^>]*><p class="dk-p dk-ps-KitTable" style="line-height:10pt">Cell/)
  })
})

describe('kerning', () => {
  it('sets no kerning on the sheet, as Word kerns nothing unasked', () => {
    const html = writeHtml(document([section([paragraph('AVATAR')])]))
    expect(html).toMatch(/\.dk-sheet\{[^}]*font-kerning:none/)
    expect(html).not.toMatch(/font-kerning:normal/)
  })

  it('kerns a run, a style and the whole document where they ask (font-kerning: normal)', () => {
    const run = writeHtml(document([section([paragraph([text('AVATAR', { kern: true }), text('Plain', { kern: false })])])]))
    expect(run).toMatch(/<span style="font-kerning:normal">AVATAR<\/span>/)
    expect(run).toMatch(/<span style="font-kerning:none">Plain<\/span>/)
    const themed = writeHtml(document([section([heading(1, 'Kerned'), paragraph('Kerned')])], { theme: { body: { kern: true } } }))
    expect(themed).toMatch(/\.dk-ps-Normal\{[^}]*font-kerning:normal/)
    expect(themed).toMatch(/\.dk-ps-Heading1\{[^}]*font-kerning:normal/)
  })

  it('reads a size threshold against the style\'s own size', () => {
    const html = writeHtml(
      document([section([paragraph('x')])], {
        theme: { body: { kern: 14 }, headings: [{ size: 20 }] },
      }),
    )
    // Body text is 11 pt, under the threshold; a 20 pt heading is over it.
    expect(html).toMatch(/\.dk-ps-Normal\{[^}]*font-kerning:none/)
    expect(html).toMatch(/\.dk-ps-Heading1\{[^}]*font-kerning:normal/)
  })
})
