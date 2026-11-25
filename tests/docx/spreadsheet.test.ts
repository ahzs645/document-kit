// The spreadsheet block: one sheet of an .xlsx printed as a native Word table, formulas evaluated where the file
// cached no value, on a landscape 17 x 11 in page with the header row repeating.
import { describe, expect, it } from 'vitest'

import { caption, document, plainText, section, spreadsheet, toCell, toComplete, type Block, type TableBlock } from '../../src/model'
import { spreadsheetTable, writeDocx } from '../../src/docx'
import { readPackage } from './helpers'
import { budgetWorkbook, makeXlsx } from './xlsx'

const TABLOID_LANDSCAPE = { paper: 'tabloid', orientation: 'landscape', margins: { left: 36, right: 36, top: 54, bottom: 54 } } as const
const xmlOf = async (blocks: Block[]) =>
  readPackage(await writeDocx(document([section(blocks, { page: TABLOID_LANDSCAPE })]))).document
const texts = (table: TableBlock) => table.rows.map((row) => row.cells.map((input) => plainText(toCell(input).content ?? '')))
const textOf = (table: TableBlock, row: number, column: number) => plainText(toCell(table.rows[row]!.cells[column]!).content ?? '')
// The texts of the table's rows in the written Word XML.
const rowTexts = (xml: string) =>
  [...xml.matchAll(/<w:tr>([\s\S]*?)<\/w:tr>/g)].map((m) => [...m[1]!.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((t) => t[1]).join('|'))

describe('spreadsheet block: values', () => {
  const book = budgetWorkbook()

  it('evaluates formulas the file cached no value for: products, a shared formula, SUM, arithmetic', () => {
    const table = spreadsheetTable(spreadsheet(book, { sheet: 'Budget' }), { width: 1152 })
    // Hours x rate, written once as a shared formula and shifted down the column.
    expect(texts(table).slice(1, 5).map((row) => row[4])).toEqual(['$1,500.00', '$3,000.00', '$2,250.00', '$2,250.00'])
    expect(textOf(table, 5, 0)).toBe('Total before tax')
    expect(texts(table)[5]).toContain('120')
    expect(texts(table)[5]).toContain('$9,000.00')
    expect(texts(table)[7]).toContain('$450.00')
    expect(texts(table)[9]).toContain('$9,450.00')
    expect(texts(table)[8]).toContain('No charge')
  })

  it('follows references to another sheet, and formats percentages', () => {
    const table = spreadsheetTable(spreadsheet(book, { sheet: 'Breakdown' }), { width: 1152 })
    expect(texts(table)[2]).toEqual(['Planning', '$1,500.00', '16.7%'])
    expect(texts(table)[3]).toEqual(['Design', '$3,000.00', '33.3%'])
  })

  it('prints a cached value as the file stores it, unless asked to evaluate every formula', () => {
    const cached = spreadsheetTable(spreadsheet(book, { sheet: 'Breakdown' }), { width: 1152 })
    expect(texts(cached)[4]).toEqual(['Cached', '$1,234.00', ''])
    const evaluated = spreadsheetTable(spreadsheet(book, { sheet: 'Breakdown', formulas: 'evaluate' }), { width: 1152 })
    expect(texts(evaluated)[4]).toEqual(['Cached', '$2,250.00', ''])
  })

  it('takes a sheet by position, the first sheet by default, and a range', () => {
    const first = spreadsheetTable(spreadsheet(book), { width: 1152 })
    expect(textOf(first, 0, 0)).toBe('TASK')
    const second = spreadsheetTable(spreadsheet(book, { sheet: 1 }), { width: 1152 })
    expect(textOf(second, 1, 0)).toBe('TASK')
    const part = spreadsheetTable(spreadsheet(book, { sheet: 'Budget', range: 'A1:C3' }), { width: 1152 })
    expect(texts(part)).toEqual([
      ['TASK', 'DESCRIPTION', 'HOURS'],
      ['Planning', 'Work plan, schedule and weekly meetings with the client team.', '20'],
      ['Design', 'Templates, components and a tested prototype.', '40'],
    ])
  })

  it('names the sheets the workbook has when the one asked for is not there', () => {
    expect(() => spreadsheetTable(spreadsheet(book, { sheet: 'Summary' }), { width: 1152 })).toThrow(/"Budget", "Breakdown"/)
  })

  it('marks a formula that refers to itself, and does not loop', () => {
    const loop = makeXlsx([{ name: 'Loop', rows: [{ cells: { A: { f: 'B1+1' }, B: { f: 'A1*2' } } }] }])
    const table = spreadsheetTable(spreadsheet(loop), { width: 600 })
    expect(texts(table)[0]).toEqual(['#REF!', '#REF!'])
  })
})

describe('spreadsheet block: the look of the sheet', () => {
  const book = budgetWorkbook()

  it('keeps the column widths in proportion, and fits a sheet set to fit its page width', () => {
    const table = spreadsheetTable(spreadsheet(book, { sheet: 'Budget' }), { width: 400 })
    const widths = table.columnWidths!
    expect(widths).toHaveLength(5)
    expect(widths[1]! / widths[0]!).toBeCloseTo(2, 1)
    // Wider than 400pt: fitted to the width, its text scaled down with it.
    expect(table.width).toBe(100)
    expect(toCell(table.rows[0]!.cells[0]!).text?.size).toBeLessThan(11)
    // On a wide page it prints at its own size, no wider than it is.
    const wide = spreadsheetTable(spreadsheet(book, { sheet: 'Budget' }), { width: 1152 })
    expect(wide.width).toBeLessThan(100)
    expect(toCell(wide.rows[0]!.cells[0]!).text?.size).toBe(11)
    // A scale of its own wins.
    const scaled = spreadsheetTable(spreadsheet(book, { sheet: 'Budget', scale: 50 }), { width: 1152 })
    expect(toCell(scaled.rows[0]!.cells[0]!).text?.size).toBe(5.5)
  })

  it('keeps fills, bold, sizes, font colours, borders, merges and wrapped text', () => {
    const table = spreadsheetTable(spreadsheet(book, { sheet: 'Budget' }), { width: 1152 })
    const header = toCell(table.rows[0]!.cells[0]!)
    expect(header.shading).toBe('E7E6E6')
    expect(header.text?.bold).toBe(true)
    expect(header.borders?.top).toMatchObject({ style: 'single', width: 0.5, color: '000000' })
    const band = toCell(table.rows[6]!.cells[0]!)
    expect(band).toMatchObject({ colSpan: 5, shading: 'FCE4D6', text: { bold: true } })
    const grand = table.rows[9]!.cells.map(toCell)
    expect(grand[0]).toMatchObject({ colSpan: 4, shading: 'FFFF00', text: { bold: true, size: 14 } })
    expect(grand[1]).toMatchObject({ align: 'right', shading: 'FFFF00' })
    const noCharge = toCell(table.rows[8]!.cells.at(-1)!)
    expect(noCharge).toMatchObject({ align: 'right', text: { bold: true, color: 'FF0000' } })
    // Wrapped text sits at the top of its cell; others at the bottom, as Excel sets them.
    expect(toCell(table.rows[1]!.cells[1]!).verticalAlign).toBe('top')
    expect(toCell(table.rows[5]!.cells[0]!).verticalAlign).toBe('bottom')
    // A custom row height is a least height.
    expect(table.rows[0]!.height).toBe(30)
  })

  it('lets unwrapped text run into the empty cells beside it, as Excel prints it', () => {
    const table = spreadsheetTable(spreadsheet(book, { sheet: 'Breakdown' }), { width: 1152 })
    const title = toCell(table.rows[0]!.cells[0]!)
    expect(plainText(title.content ?? '')).toBe('Budget breakdown for the made-up project')
    expect(title.colSpan).toBeGreaterThan(1)
  })

  it('repeats the frozen rows on each page, or as many rows as the block says', () => {
    const table = spreadsheetTable(spreadsheet(book, { sheet: 'Budget' }), { width: 1152 })
    expect(table.rows[0]).toMatchObject({ repeat: true, keepWithNext: true })
    expect(table.rows[1]!.repeat).toBeUndefined()
    const none = spreadsheetTable(spreadsheet(book, { sheet: 'Breakdown' }), { width: 1152 })
    expect(none.rows.some((row) => row.repeat)).toBe(false)
    const two = spreadsheetTable(spreadsheet(book, { sheet: 'Breakdown', headerRows: 2 }), { width: 1152 })
    expect(two.rows.slice(0, 3).map((row) => row.repeat === true)).toEqual([true, true, false])
  })
})

describe('spreadsheet block: the Word file', () => {
  const book = budgetWorkbook()

  it('writes both sheets as tables on landscape 17 x 11 in pages, the header row repeating', async () => {
    const xml = await xmlOf([
      spreadsheet(book, { sheet: 'Budget' }),
      { ...spreadsheet(book, { sheet: 'Breakdown' }), pageBreakBefore: true },
    ])
    expect(xml.match(/<w:tbl>/g)).toHaveLength(2)
    expect(xml).toMatch(/<w:pgSz w:w="24480" w:h="15840" w:orient="landscape"\/>/)
    // The header row repeats; its cells are grey and bold.
    const firstRow = xml.slice(xml.indexOf('<w:tr>'), xml.indexOf('</w:tr>'))
    expect(firstRow).toContain('<w:tblHeader/>')
    expect(firstRow).toContain('w:fill="E7E6E6"')
    expect(firstRow).toContain('<w:b/>')
    // The merged band and the grand total.
    expect(xml).toMatch(/<w:gridSpan w:val="5"\/>[\s\S]*?w:fill="FCE4D6"/)
    const rows = rowTexts(xml)
    expect(rows).toContain('TOTAL INCLUDING TAX|$9,450.00')
    expect(rows).toContain('Planning|$1,500.00|16.7%')
    // Line heights are stated as multiples, so LibreOffice sets a large font's lines apart too.
    expect(firstRow).toMatch(/<w:spacing w:line="240" w:lineRule="auto"\/>/)
    // No row breaks across a page.
    expect(xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>'))).not.toMatch(/<w:tr>(?:(?!<w:cantSplit\/>)[\s\S])*?<\/w:tr>/)
  })

  it('prints a custom paper size turned landscape', async () => {
    const bytes = await writeDocx(
      document([section([spreadsheet(book)], { page: { paper: { width: 792, height: 1224 }, orientation: 'landscape' } })]),
    )
    expect(readPackage(bytes).document).toMatch(/<w:pgSz w:w="24480" w:h="15840" w:orient="landscape"\/>/)
  })

  it('numbers its caption with the tables', async () => {
    const xml = await xmlOf([spreadsheet(book, { caption: caption('Project budget') })])
    expect(xml).toMatch(/<w:fldSimple w:instr="SEQ Table \\\* ARABIC"><w:r><w:t xml:space="preserve">1<\/w:t>/)
    expect(xml).toContain('Project budget')
  })

  it('prints a placeholder for a missing workbook and lists it to complete', async () => {
    const block = spreadsheet(null, { sheet: 'Budget' })
    const xml = await xmlOf([block])
    expect(xml).toContain('Spreadsheet: Budget')
    expect(toComplete(document([section([block])]))).toEqual([expect.objectContaining({ kind: 'spreadsheet', what: 'Budget' })])
  })
})
