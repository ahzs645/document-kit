// Tracking on runs and styles, and the named paragraph and character styles a document defines. Made-up text.
import { describe, expect, it } from 'vitest'

import { document, heading, paragraph, section, text, type Block, type Document } from '../../src/model'
import { writeDocx } from '../../src/docx'
import { all, readPackage } from './helpers'

const write = async (blocks: Block[], extra: Partial<Document> = {}) => readPackage(await writeDocx(document([section(blocks)], extra)))
const styleXml = (styles: string, id: string) => new RegExp(`<w:style [^>]*w:styleId="${id}"[\\s\\S]*?</w:style>`).exec(styles)?.[0] ?? ''

describe('character spacing', () => {
  it('writes a run\'s tracking in twentieths of a point, an explicit 0 and a negative included', async () => {
    const { document: xml } = await write([
      paragraph([text('Wide', { characterSpacing: 1.5 }), text('Undone', { characterSpacing: 0 }), text('Tight', { characterSpacing: -0.25 })]),
    ])
    expect(all(xml, /<w:spacing w:val="(-?\d+)"\/>/)).toEqual(['30', '0', '-5'])
  })

  it('writes the theme\'s tracking on its styles', async () => {
    const pkg = await write([heading(1, 'Tracked')], { theme: { headings: [{ characterSpacing: 2 }] } })
    expect(styleXml(pkg.text('word/styles.xml'), 'Heading1')).toMatch(/<w:spacing w:val="40"\/>/)
  })
})

describe('named styles', () => {
  const styles: Document['styles'] = {
    paragraph: [
      { id: 'AffiliationDepartment', name: 'Affiliation Department', text: { size: 12, bold: true, characterSpacing: 0.5, after: 2 } },
      { id: 'AffiliationSubdepartment', basedOn: 'AffiliationDepartment', next: 'normal', text: { size: 10.5, bold: false }, indent: { left: 6 } },
    ],
    character: [{ id: 'Accent', name: 'Accent', text: { color: 'b45309', characterSpacing: 1 } }],
  }

  it('defines each style in styles.xml with its base, next style and own look', async () => {
    const pkg = await write([paragraph('Dept', { style: 'AffiliationDepartment' })], { styles })
    const xml = pkg.text('word/styles.xml')
    const dept = styleXml(xml, 'AffiliationDepartment')
    expect(dept).toMatch(/w:type="paragraph"/)
    expect(dept).toMatch(/<w:name w:val="Affiliation Department"\/>/)
    expect(dept).toMatch(/<w:basedOn w:val="Normal"\/>/)
    expect(dept).toMatch(/<w:spacing w:val="10"\/>/)
    expect(dept).toMatch(/<w:sz w:val="24"\/>/)
    expect(dept).toMatch(/<w:b\/>/)
    const sub = styleXml(xml, 'AffiliationSubdepartment')
    expect(sub).toMatch(/<w:basedOn w:val="AffiliationDepartment"\/>/)
    expect(sub).toMatch(/<w:next w:val="Normal"\/>/)
    expect(sub).toMatch(/<w:ind w:left="120"\/>/)
    expect(sub).toMatch(/<w:b w:val="false"\/>/)
    const accent = styleXml(xml, 'Accent')
    expect(accent).toMatch(/w:type="character"/)
    expect(accent).toMatch(/<w:color w:val="B45309"\/>/)
  })

  it('references them from paragraphs and runs', async () => {
    const { document: xml } = await write(
      [paragraph([text('Sub', { characterStyle: 'Accent' })], { style: 'AffiliationSubdepartment' })],
      { styles },
    )
    expect(xml).toMatch(/<w:pStyle w:val="AffiliationSubdepartment"\/>/)
    expect(xml).toMatch(/<w:rPr><w:rStyle w:val="Accent"\/>/)
  })

  it('stops, naming it, at a style the document does not define', async () => {
    await expect(write([paragraph('x', { style: 'Missing' })])).rejects.toThrow(/"Missing"/)
    await expect(write([paragraph([text('x', { characterStyle: 'Gone' })])])).rejects.toThrow(/"Gone"/)
  })
})
