// The Word writer's side of the report features: pinned blocks as one text frame at the foot of the page, tick boxes
// as Word check boxes, placeholders, signature options, fields on one line, theme row heights and text rules.
import { describe, expect, it } from 'vitest'

import {
  checkbox,
  choices,
  document,
  field,
  fields,
  paragraph,
  pinned,
  placeholder,
  section,
  sentence,
  signature,
  table,
  headerRow,
  row,
  type Block,
  type Document,
} from '../../src/model'
import { writeDocx } from '../../src/docx'
import { all, makePng, paragraphTexts, readPackage } from './helpers'

const xmlOf = async (blocks: Block[], extra: Partial<Document> = {}) =>
  readPackage(await writeDocx(document([section(blocks)], extra))).document
const paragraphsOf = (xml: string) => all(xml, /(<w:p\b[^>]*>[\s\S]*?<\/w:p>)/)

describe('pinned blocks', () => {
  it('writes each of its paragraphs in one frame anchored to the foot of the margin', async () => {
    const xml = await xmlOf([paragraph('Flow.'), pinned([paragraph('Pinned one.'), paragraph('Pinned two.')], { spaceBelow: 12 }), paragraph('After.')])
    const paras = paragraphsOf(xml)
    const framed = paras.filter((p) => p.includes('<w:framePr'))
    expect(framed).toHaveLength(3)
    const frames = framed.map((p) => /<w:framePr[^>]*\/>/.exec(p)?.[0])
    expect(new Set(frames).size).toBe(1)
    expect(frames[0]).toMatch(/w:vAnchor="margin"/)
    expect(frames[0]).toMatch(/w:yAlign="bottom"/)
    expect(frames[0]).toMatch(/w:wrap="around"/)
    expect(frames[0]).toMatch(/w:w="9360"/)
    expect(paragraphTexts(xml)).toEqual(['Flow.', 'Pinned one.', 'Pinned two.', '', 'After.'])
  })
})

describe('tick boxes', () => {
  it('writes a check box content control, ticked or not', async () => {
    const xml = await xmlOf([paragraph(['Yes ', checkbox(true), ' No ', checkbox(false)])])
    const boxes = all(xml, /(<w14:checkbox>[\s\S]*?<\/w14:checkbox>)/)
    expect(boxes).toHaveLength(2)
    expect(boxes[0]).toMatch(/<w14:checked w14:val="1"\/>/)
    expect(boxes[1]).toMatch(/<w14:checked w14:val="0"\/>/)
    expect(xml).toMatch(/<w:sym w:char="2612" w:font="MS Gothic"\/>/)
    expect(xml).toMatch(/<w:sym w:char="2610" w:font="MS Gothic"\/>/)
  })

  it('writes a question boxed: a ruled table, the question left and the boxes right', async () => {
    const xml = await xmlOf([choices(['Yes', 'No'], 'No', { label: 'Made-up question?', answer: 'Because.', boxed: true })])
    expect(xml).toContain('<w:tbl>')
    expect(all(xml, /<w14:checked w14:val="(\d)"\/>/)).toEqual(['0', '1'])
    expect(paragraphTexts(xml)).toEqual(expect.arrayContaining(['Made-up question?', 'Because.', ' Yes', ' No']))
  })
})

describe('placeholders', () => {
  it('prints a missing value as its bracketed name, marked', async () => {
    const xml = await xmlOf([paragraph(['Run on ', placeholder('metals run date'), '.']), paragraph([sentence('It is {state}.', { state: null }, { missing: 'placeholder' })])], {
      theme: { placeholderMark: { highlight: 'cyan', bold: true } },
    })
    expect(paragraphTexts(xml)).toEqual(['Run on [metals run date].', 'It is [state].'])
    const marked = all(xml, /<w:r>(<w:rPr>(?:(?!<\/w:rPr>)[\s\S])*<\/w:rPr><w:t[^>]*>\[[^<]*<\/w:t>)/)
    expect(marked).toHaveLength(2)
    for (const run of marked) {
      expect(run).toMatch(/<w:highlight w:val="cyan"\/>/)
      expect(run).toMatch(/<w:b\/>/)
    }
  })
})

describe('signature options', () => {
  const sig = (options: Parameters<typeof signature>[0]) => xmlOf([signature({ label: 'Approved by:', name: 'A. Person', ...options })])

  it('puts the title on the name\'s line or the next, centres the rule, sets its weight, and keeps the block on one page', async () => {
    const same = await sig({ title: 'Made-up Officer', align: 'center', ruleWidth: 216, ruleThickness: 1, ruleColor: '112233', ruleGap: 5, signingSpace: 37, spaceAbove: 13, spaceBelow: 12 })
    expect(paragraphTexts(same)).toEqual(['Approved by:', '', 'A. Person, Made-up Officer'])
    const paras = paragraphsOf(same).slice(0, 3)
    // Kept together: keep-with-next on all but the last, keep-lines on every one.
    expect(paras.map((p) => p.includes('<w:keepNext/>'))).toEqual([true, true, false])
    expect(paras.every((p) => p.includes('<w:keepLines/>'))).toBe(true)
    expect(paras[0]).toMatch(/<w:spacing w:after="80" w:before="260"\/>/)
    expect(paras[1]).toMatch(/w:line="740" w:lineRule="exact"/)
    // Letter, one-inch margins: 468pt of text, a 216pt rule, 126pt each side.
    expect(paras[2]).toMatch(/<w:ind w:left="2520" w:right="2520"\/>/)
    expect(paras[2]).toMatch(/<w:top w:val="single" w:color="112233" w:sz="8" w:space="5"\/>/)
    expect(paras[2]).toMatch(/w:after="240"/)
    const next = await sig({ title: 'Made-up Officer', titleLayout: 'nextLine', labelPosition: 'below', keepTogether: false })
    expect(paragraphTexts(next)).toEqual(['', 'A. Person', 'Made-up Officer', 'Approved by:'])
    expect(paragraphsOf(next).some((p) => p.includes('<w:keepNext/>'))).toBe(false)
  })

  it('fits a picture in its box and leaves no signing room', async () => {
    const xml = await sig({ image: { bytes: makePng(400, 100) }, imageBox: { width: 120, height: 46 } })
    // 400 × 100 into 120 × 46: 120 wide, 30 high (EMU).
    expect(xml).toContain('cx="1524000" cy="381000"')
    expect(xml).not.toMatch(/w:lineRule="exact"/)
  })
})

describe('fields on one line', () => {
  it('writes a same-line row into the row before it, its spaces and hyphens unbreakable', async () => {
    const xml = await xmlOf([fields([field('E-mail', 'a@example.org'), field('Phone', '555-0100', { sameLine: true }), field('Date', 'Today')], { sameLineGap: 24, rowSpacing: 3.5 })])
    const texts = paragraphTexts(xml)
    // 24pt of gap at 11pt is nine spaces; the hyphen is Word's no-break hyphen, which has no text of its own.
    expect(texts[0]).toMatch(/^E-mail: a@example\.org {9}Phone:\u00a05550100$/)
    expect(paragraphsOf(xml)[0]).toContain('<w:noBreakHyphen/>')
    expect(texts[1]).toBe('Date: Today')
    expect(paragraphsOf(xml)[0]).toMatch(/w:after="70"/)
  })
})

describe('theme table density', () => {
  it('gives header and body rows their least heights and the table text its line height', async () => {
    const pkg = readPackage(
      await writeDocx(document([section([table([headerRow(['A']), row(['1']), row(['2'], { height: 30 })])])], { theme: { table: { headerRowHeight: 20, rowHeight: 14, line: 1.2 } } })),
    )
    expect(all(pkg.document, /<w:trHeight w:val="(\d+)" w:hRule="atLeast"\/>/)).toEqual(['400', '280', '600'])
    expect(pkg.text('word/styles.xml')).toMatch(/w:styleId="KitTable"[\s\S]*?w:line="288"/)
  })
})

describe('text rules in the Word file', () => {
  it('writes what the rules made as ordinary runs: a link, italics, a raised marker, a low digit', async () => {
    const xml = await xmlOf([table([row(['Total a', 'mg CaCO3/L'])]), paragraph('Made-up term at www.example.org.')], {
      textRules: [
        { builtIn: 'links' },
        { terms: ['Made-up term'], style: { italic: true } },
        { builtIn: 'trailingMarker', markers: 'a-h', in: ['cell'] },
        { builtIn: 'formulaDigits', formulas: ['CaCO3'] },
      ],
    })
    expect(xml).toMatch(/<w:hyperlink [^>]*>[\s\S]*?<w:t[^>]*>www\.example\.org<\/w:t>/)
    expect(xml).toMatch(/<w:i\/>[\s\S]*?<w:t[^>]*>Made-up term<\/w:t>/)
    expect(xml).toMatch(/<w:vertAlign w:val="superscript"\/>[\s\S]*?<w:t[^>]*>a<\/w:t>/)
    expect(xml).toMatch(/<w:vertAlign w:val="subscript"\/>[\s\S]*?<w:t[^>]*>3<\/w:t>/)
  })
})

describe('a table\'s own density', () => {
  it('writes autofit and a line height when the table asks, fixed and the style\'s line otherwise', async () => {
    const own = await xmlOf([table([row(['a'])], { autoFit: true, line: 0.85 })])
    expect(own).toContain('<w:tblLayout w:type="autofit"/>')
    expect(own).toMatch(/<w:pStyle w:val="KitTable"\/><w:spacing w:line="204"\/>/)
    const plain = await xmlOf([table([row(['a'])])])
    expect(plain).toContain('<w:tblLayout w:type="fixed"/>')
    expect(plain).not.toMatch(/<w:pStyle w:val="KitTable"\/><w:spacing/)
  })
})
