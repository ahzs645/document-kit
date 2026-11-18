// Both made-up report families, from one component library, run the whole way: compile → render → Word → layout,
// with the bundled fonts and no network. Set RECIPE_SAMPLES to a folder to keep the compiled JSON and the PDFs.

import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { writeDocx } from '../../src/docx'
import { createFontProvider, layoutDocx } from '../../src/layout'
import { compileRecipe, renderRecipe, type CompiledRecipe } from '../../src/recipe'
import { makePng, readPackage } from '../docx/helpers'
import { diskLoader } from '../layout/sample'
import { code, fixtureData, fixtureLibrary, registry } from './fixtures/library'

const OUT = process.env.RECIPE_SAMPLES
const keep = (name: string, bytes: Uint8Array | string): void => {
  if (!OUT) return
  mkdirSync(OUT, { recursive: true })
  writeFileSync(join(OUT, name), bytes)
}

const fetchSpy = vi.fn(() => {
  throw new Error('fetch was called')
})
beforeEach(() => vi.stubGlobal('fetch', fetchSpy))
afterEach(() => vi.unstubAllGlobals())

const SOURCE_STYLES = { reviewer: { highlight: 'yellow' as const }, edited: { highlight: 'cyan' as const } }

async function run(recipe: string, data: Record<string, unknown>, options: Parameters<typeof renderRecipe>[2] = {}) {
  const compiled: CompiledRecipe = compileRecipe(recipe, fixtureLibrary(), { code })
  const rendered = renderRecipe(compiled, data, { code, registry, ...options })
  const docx = await writeDocx(rendered.document, { registry, sourceStyles: SOURCE_STYLES })
  const laid = await layoutDocx(docx, { fonts: createFontProvider(diskLoader) })
  return { compiled, rendered, docx, laid, text: laid.pageText.join('\n'), flat: laid.pageText.join(' ').replace(/\s+/g, ' ') }
}

describe('a laboratory-style report', () => {
  it('prints its letterhead, fields, notes, one table per set of results and a signature', async () => {
    const { compiled, rendered, docx, laid, text, flat } = await run('laboratory', fixtureData('laboratory'), {
      sources: { 'tables[].rows': 'instrument', client: 'office' },
    })
    keep('laboratory.compiled.json', `${JSON.stringify(compiled, null, 2)}\n`)
    keep('laboratory.pdf', laid.toPdf().bytes)
    expect(text).toContain('Made-up Testing Laboratory')
    expect(text).toContain('QF-07 rev 3')
    expect(text).toContain('Analysis of Water Samples')
    expect(text).toContain('Attention: A. Sample')
    expect(flat).toContain('3 samples were received on 2 March 2026 in good condition.')
    expect(flat).toContain('They were analysed between 3 March and 9 March 2026.')
    // The notes take the placeholder rule: the sentence whose slot the data lacks prints its slot's name, marked.
    expect(flat).toContain('Results are reported on [basis].')
    expect(text).toContain('E-mail: office@example.org')
    // The phone shares the e-mail's line; its spaces and hyphen do not break.
    expect(text).toMatch(/E-mail: office@example\.org\s+Phone:\s555-0100/)
    // The contact block is pinned to the foot of the last page, under everything else on it.
    const last = laid.pageText[laid.pageText.length - 1]!.split('\n')
    expect(last.slice(-4, -1).join(' ')).toContain('Phone 555-0199 | lab@example.org | www.example.org/lab')
    // The text rules reach the Word file as ordinary runs: links, a raised marker, a low digit, a term in italics.
    const xml = readPackage(docx).document
    expect(xml).toMatch(/<w:hyperlink [^>]*>[\s\S]*?<w:t[^>]*>lab@example\.org<\/w:t>/)
    expect(xml).toMatch(/<w:vertAlign w:val="superscript"\/>[\s\S]{0,200}<w:t[^>]*>a<\/w:t>/)
    expect(xml).toMatch(/<w:vertAlign w:val="subscript"\/>[\s\S]{0,200}<w:t[^>]*>3<\/w:t>/)
    expect(xml).toMatch(/<w:i\/>[\s\S]{0,200}<w:t[^>]*>M\. exemplum<\/w:t>/)
    // The signature takes its options: centred, the title on the name's line.
    expect(text).toContain('B. Analyst, Laboratory Supervisor')
    expect(text).toContain('General chemistry')
    expect(text).toContain('Metals')
    expect(text).toContain('Nutrients')
    expect(text).toContain('Reviewed and approved')
    expect(text).toContain('B. Analyst')
    laid.pageText.forEach((page, i) => expect(page).toContain(`Page ${i + 1} of ${laid.pageCount}`))
    expect(rendered.parts.find((p) => p.id === 'results')?.repeats).toBe(2)
    expect(rendered.toComplete.map((item) => [item.partId, item.kind, item.what])).toEqual([
      ['addressee', 'field', 'Client reference'],
      ['notes', 'sentence', 'Method'],
      ['results', 'cell', 'General chemistry, Nitrate as NO3 a, W-2'],
    ])
    expect(fetchSpy).not.toHaveBeenCalled()
  }, 60_000)

  it('the interim report is the base report with its status, a notice and the brief notes', async () => {
    const { compiled, text, flat } = await run('laboratory-interim', fixtureData('laboratory'))
    keep('laboratory-interim.compiled.json', `${JSON.stringify(compiled, null, 2)}\n`)
    expect(compiled.from.recipes.map((r) => r.id)).toEqual(['laboratory-interim', 'laboratory'])
    expect(text).toContain('Status: Interim')
    expect(flat).toContain('Interim report. 2 results are still to follow')
    expect(flat).toContain('3 samples were received on 2 March 2026.')
    expect(text).not.toContain('in good condition')
    expect(text).toContain('Metals')
  }, 60_000)
})

describe('a review-style report', () => {
  it('prints a cover, contents, blanks, a computed summary, sentences, a matrix, a landscape table, a figure page and the to-complete list', async () => {
    const data = fixtureData('review')
    const sheet = data.sheet as Record<string, unknown>
    const withPicture = { ...data, sheet: { ...sheet, image: { bytes: makePng(1224, 792, [205, 222, 210]) } } }
    const { compiled, rendered, laid, text, flat } = await run('review', withPicture)
    keep('review.compiled.json', `${JSON.stringify(compiled, null, 2)}\n`)
    keep('review.pdf', laid.toPdf().bytes)
    expect(compiled.code.map((c) => c.id)).toEqual(['review-summary'])
    expect(compiled.customKinds).toEqual(['matrix'])
    expect(text).toContain('Visual Review')
    expect(text).toContain('Contents')
    expect(text).toContain('Summary')
    expect(flat).toContain('2 of 2 viewpoints see the opening.')
    expect(text).toContain('Above the limit.')
    expect(flat).toContain('From Viewpoint 1, lake shore, the proposed opening is easy to see.')
    expect(flat).not.toContain('From Viewpoint 2, highway pull-out, the proposed opening')
    expect(flat).toContain('About 7% of its area is seen, 5.8 km away.')
    expect(text).toContain('Made-up rating matrix')
    expect(text).toContain('Made-up inventory units')
    expect(text).toContain('Analysis sheet')
    expect(text).toContain('To complete before signing')
    expect(text).toContain('Checked by')
    // The inventory is landscape; the sheet 11 × 17 on its side.
    expect(laid.pageSizes.some((p) => p.widthPt > p.heightPt && Math.round(p.widthPt) === 792)).toBe(true)
    expect(laid.pageSizes.some((p) => Math.round(p.widthPt) === 1224 && Math.round(p.heightPt) === 792)).toBe(true)
    // The to-complete list printed at the end names what the rest of the document left.
    const last = (laid.pageText[laid.pageText.length - 1] ?? '').replace(/\s+/g, ' ')
    expect(last).toContain('Inventory date')
    expect(last).toContain('Viewpoint 2, highway pull-out: Visibility (needs visibility)')
    expect(rendered.toComplete.map((item) => item.partId)).toEqual(['project', 'summary', 'discussion', 'objective-met', 'inventory-table'])
    // The tick-box question: boxes drawn, the chosen one crossed, in a ruled box.
    expect(flat).toMatch(/Is the made-up objective met\?.*☐\s?Yes ☒\s?Partly ☐\s?No/)
    expect(fetchSpy).not.toHaveBeenCalled()
  }, 60_000)
})
