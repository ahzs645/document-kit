// A spreadsheet block in a recipe: the workbook comes from the data as bytes, like a picture.
import { describe, expect, it } from 'vitest'

import { writeDocx } from '../../src/docx'
import { compileRecipe, renderRecipe, sharedLibrary } from '../../src/recipe'
import { readPackage } from '../docx/helpers'
import { budgetWorkbook } from '../docx/xlsx'
import { comp, lib, recipe } from './helpers'

const TABLOID_LANDSCAPE = { paper: 'tabloid', orientation: 'landscape' }
const rowTexts = (xml: string) =>
  [...xml.matchAll(/<w:tr>([\s\S]*?)<\/w:tr>/g)].map((m) => [...m[1]!.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((t) => t[1]).join('|'))

describe('spreadsheet block in a recipe', () => {
  it('takes the workbook from the data, like a picture', async () => {
    const compiled = compileRecipe(
      recipe('budget', [], { sections: [{ page: TABLOID_LANDSCAPE, parts: [{ use: 'sheet', id: 'part' }] }] }),
      lib([comp('sheet', [{ type: 'spreadsheet', workbook: '{{ budget }}', sheet: '{{ sheetName }}', range: 'A1:E10' }])]),
    )
    const rendered = renderRecipe(compiled, { budget: budgetWorkbook(), sheetName: 'Budget' })
    const block = rendered.document.sections[0]!.body[0]!
    expect(block).toMatchObject({ type: 'spreadsheet', sheet: 'Budget', range: 'A1:E10', workbook: { bytes: expect.any(Uint8Array) } })
    const xml = readPackage(await writeDocx(rendered.document)).document
    expect(rowTexts(xml)).toContain('TOTAL INCLUDING TAX|$9,450.00')
    expect(xml).toMatch(/w:orient="landscape"/)
  })

  it('rejects bytes written into the recipe itself', () => {
    expect(() =>
      compileRecipe(recipe('r', [{ use: 'c' }]), lib([comp('c', [{ type: 'spreadsheet', workbook: { bytes: [1, 2] } }])])),
    ).toThrow(/a file comes from the data/)
  })
})

describe('the shared spreadsheet component', () => {
  it('prints a sheet from the data, and a placeholder when the data has none', async () => {
    const parts = [{ use: 'spreadsheet', with: { workbook: '{{ budget }}', sheet: 'Breakdown', caption: 'Breakdown', headerRows: 2 } }]
    const compiled = compileRecipe(recipe('r', parts), sharedLibrary())
    const rendered = renderRecipe(compiled, { budget: budgetWorkbook() })
    expect(rendered.document.sections[0]!.body[0]).toMatchObject({
      type: 'spreadsheet',
      sheet: 'Breakdown',
      headerRows: 2,
      formulas: 'cached',
      caption: { text: 'Breakdown', numbered: true },
    })
    expect(rowTexts(readPackage(await writeDocx(rendered.document)).document)).toContain('Planning|$1,500.00|16.7%')
    const missing = renderRecipe(compiled, { budget: null })
    expect(missing.toComplete).toEqual([expect.objectContaining({ kind: 'spreadsheet', what: 'Breakdown' })])
  })
})
