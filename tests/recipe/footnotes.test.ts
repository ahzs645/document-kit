import { describe, expect, it } from 'vitest'
import { writeDocx } from '../../src/docx'
import { compileRecipe, renderRecipe } from '../../src/recipe'
import { readPackage } from '../docx/helpers'
import { comp, lib, para, recipe } from './helpers'

const footnotes = { capacity: ['Minimum per unit.', 'Measured separately.'], target: ['Annual average.'], unused: ['Omitted.'] }
const compiled = compileRecipe(recipe('r', [{ use: 'matrix' }], { document: { footnotes } }),
  lib([comp('matrix', [{ type: 'table', rows: '{{ rows }}' }])]))

async function exportRows(ids: string[]) {
  const rows = ids.map(id => ({ cells: [{ content: ['Value', { type: 'footnote', id }] }] }))
  return readPackage(await writeDocx(renderRecipe(compiled, { rows }).document))
}

describe('linked footnotes', () => {
  it.each([
    { ids: ['target', 'capacity', 'capacity'], marker: '2' },
    { ids: ['capacity', 'target', 'capacity'], marker: '1' },
    { ids: ['capacity', 'capacity'], marker: '1' },
  ])('numbers by first citation in $ids, independent of definition order', async ({ ids, marker }) => {
    const pkg = await exportRows(ids)
    const anchors = [...pkg.document.matchAll(/w:name="dkNote_(\w+)"[^>]*\/>.*?<w:footnoteReference w:id="(\d+)"/gs)]
    expect(anchors.map(m => [m[1], m[2]])).toEqual([...new Set(ids)].map((id, i) => [id, String(i + 1)]))
    expect(pkg.document).toMatch(new RegExp('NOTEREF dkNote_capacity[^>]*>.*?<w:t[^>]*>' + marker + '</w:t>', 's'))
    const notes = pkg.text('word/footnotes.xml')
    expect(notes.match(/Minimum per unit\./g)).toHaveLength(1)
    expect(notes).toContain('Measured separately.')
    expect(notes.includes('Annual average.')).toBe(ids.includes('target'))
    expect(notes).not.toContain('Omitted.')
  })

  it('rejects an unresolved citation or a running header citation', async () => {
    await expect(exportRows(['typo'])).rejects.toThrow('footnote "typo"')
    const header = [para({ type: 'footnote', id: 'target' })]
    const job = { ...recipe('r', []), document: { footnotes }, sections: [{ id: 's', header, parts: [{ use: 'p' }] }] }
    const { document } = renderRecipe(compileRecipe(job, lib([comp('p', [para('Body')])])), {})
    await expect(writeDocx(document)).rejects.toThrow('document body')
  })
})
