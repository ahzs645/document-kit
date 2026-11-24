// Space after a table: the theme's table.spaceAfter, a table's own spaceAfter over it, none inside a table cell.
import { describe, expect, it } from 'vitest'

import { document, headerRow, paragraph, row, section, table, type Block, type Document } from '../../src/model'
import { writeDocx } from '../../src/docx'
import { readPackage } from './helpers'

const xmlOf = async (blocks: Block[], extra: Partial<Document> = {}) =>
  readPackage(await writeDocx(document([section(blocks)], extra))).document
const small = (options = {}) => table([headerRow(['Name', 'Role']), row(['A. Person', 'Made-up role'])], options)
// The body after the table, up to the next text paragraph.
const afterTable = (xml: string) => xml.slice(xml.lastIndexOf('</w:tbl>')).split('After.')[0]!
const gap = (twips: number) => new RegExp(`<w:spacing w:after="0" w:before="0" w:line="${twips}" w:lineRule="exact"/>`)

describe('space after a table', () => {
  it('leaves the theme default of 9pt before the next block', async () => {
    expect(afterTable(await xmlOf([small(), paragraph('After.')]))).toMatch(gap(180))
  })

  it('uses the table’s own spaceAfter, and none for 0', async () => {
    expect(afterTable(await xmlOf([small({ spaceAfter: 14 }), paragraph('After.')]))).toMatch(gap(280))
    expect(afterTable(await xmlOf([small({ spaceAfter: 0 }), paragraph('After.')]))).not.toMatch(/w:lineRule="exact"/)
  })

  it('follows the theme', async () => {
    const xml = await xmlOf([small(), paragraph('After.')], { theme: { table: { spaceAfter: 4 } } } as Partial<Document>)
    expect(afterTable(xml)).toMatch(gap(80))
  })

  it('adds nothing inside a table cell', async () => {
    const inner = small()
    const outer = table([row([{ blocks: [inner] }])])
    const xml = await xmlOf([outer, paragraph('After.')])
    const cell = xml.slice(xml.indexOf('<w:tc>'), xml.lastIndexOf('</w:tc>'))
    expect(cell).not.toMatch(gap(180))
    expect(afterTable(xml)).toMatch(gap(180))
  })
})
