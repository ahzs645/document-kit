// The two made-up report families of tests/recipe/fixtures, from one library: what they compile to, what the data
// lacks, and the review's code component beside its JSON parts and its custom matrix block.

import { describe, expect, it } from 'vitest'

import { checkRecipe, compileRecipe, isRecipeError, readCompiledRecipe, renderRecipe } from '../../src/recipe'
import { code, fixtureData, fixtureLibrary, registry } from './fixtures/library'

describe('the fixture recipes', () => {
  it('compile to the same JSON text every time, and read back as they were written', () => {
    for (const id of ['laboratory', 'laboratory-interim', 'review']) {
      const one = JSON.stringify(compileRecipe(id, fixtureLibrary(), { code }))
      const two = JSON.stringify(compileRecipe(id, fixtureLibrary(), { code }))
      expect(two).toBe(one)
      expect(JSON.stringify(readCompiledRecipe(JSON.parse(one)))).toBe(one)
    }
  })

  it('share the kit\'s components between the two families', () => {
    const used = (id: string) => compileRecipe(id, fixtureLibrary(), { code }).from.components.map((c) => c.id)
    const shared = used('laboratory').filter((id) => used('review').includes(id))
    expect(shared).toEqual(['data-table', 'fields', 'page-footer', 'running-header', 'text-section'])
  })

  it('the interim recipe compiles from its base plus its changes', () => {
    const interim = compileRecipe('laboratory-interim', fixtureLibrary(), { code })
    expect(interim.sections[0]?.parts.map((p) => p.part.id)).toEqual(['title', 'interim-notice', 'addressee', 'notes', 'results', 'footnotes', 'sign-off', 'contact'])
    expect(interim.sections[0]?.parts[3]?.variant).toBe('brief')
  })

  it('check names what the data lacks, by data path', () => {
    const lab = compileRecipe('laboratory', fixtureLibrary(), { code })
    expect(checkRecipe(lab, fixtureData('laboratory')).map((p) => p.path)).toEqual(['basis', 'tables.1.note'])
    const review = compileRecipe('review', fixtureLibrary(), { code })
    expect(checkRecipe(review, fixtureData('review'), { code, registry }).map((p) => p.path)).toEqual(['project.checker', 'viewpoints.1.visibility', 'sheet.image'])
  })
})

describe('what the laboratory fixture says beside its blocks', () => {
  const lab = compileRecipe('laboratory', fixtureLibrary(), { code })

  it('carries its meta to the compiled recipe and the rendered parts', () => {
    expect(lab.recipe.meta).toEqual({ family: 'laboratory', reviewedBy: 'nobody: a made-up fixture' })
    const { parts } = renderRecipe(lab, fixtureData('laboratory'))
    expect(parts.find((p) => p.id === 'notes')?.meta).toEqual({ wording: 'standard notes' })
    expect(parts.find((p) => p.id === 'footnotes')?.componentMeta).toEqual({ source: 'made-up fixture', rules: ['footnote-markers'] })
  })

  it('pins its contact block, and states its text rules on the document', () => {
    const { document } = renderRecipe(lab, fixtureData('laboratory'))
    const body = document.sections[0]!.body
    expect(body[body.length - 1]).toMatchObject({ type: 'pinned', spaceBelow: 12 })
    expect(document.textRules?.map((rule) => rule.id)).toEqual(['links', 'species', 'cell-markers', 'footnote-markers', 'formulas'])
  })
})

describe('the review\'s tick-box question', () => {
  const review = compileRecipe('review', fixtureLibrary(), { code })

  it('ticks the data\'s answer, and takes a hand edit of the tick and of the answer', () => {
    const plain = renderRecipe(review, fixtureData('review'), { code, registry })
    const targets = plain.edits.targets.filter((t) => t.part === 'objective-met')
    expect(targets.map((t) => [t.key, t.kind, t.original])).toEqual([
      ['objective-met||choice|Is the made-up objective met?', 'choice', 'Partly'],
      ['objective-met||answer|Is the made-up objective met?', 'answer', null],
    ])
    expect(plain.toComplete.filter((t) => t.partId === 'objective-met').map((t) => t.kind)).toEqual(['answer'])
    const edited = renderRecipe(review, fixtureData('review'), {
      code,
      registry,
      edits: { [targets[0]!.key]: 'Yes', [targets[1]!.key]: 'Seen only from the lake.' },
    })
    const part = edited.parts.find((p) => p.id === 'objective-met')!
    expect(edited.document.sections[1]!.body[part.start]).toMatchObject({ type: 'choices', chosen: 'Yes', answer: [{ text: 'Seen only from the lake.', source: 'edited' }] })
    expect(edited.toComplete.some((t) => t.partId === 'objective-met')).toBe(false)
  })
})

describe('the review mixes a code component, JSON components and a custom block', () => {
  const compiled = compileRecipe('review', fixtureLibrary(), { code })

  it('lists the code component and what it declares it reads', () => {
    expect(compiled.code).toEqual([{ id: 'review-summary', title: 'Summary', version: '1', parts: ['summary'] }])
    expect(compiled.sections[1]?.parts[2]).toMatchObject({ code: 'review-summary', params: { assessor: '{{ project.preparer }}', limit: 0.15 } })
    expect(compiled.embedded && Object.keys(compiled.embedded)).toEqual(['notice'])
    const declared = compiled.needs.filter((need) => need.places.some((place) => place.declared)).map((need) => need.path)
    expect(declared).toEqual(['viewpoints[].distanceKm', 'viewpoints[].name', 'viewpoints[].share'])
  })

  it('fails loudly without the code component, naming it; check reports it', () => {
    try {
      renderRecipe(compiled, fixtureData('review'), { registry })
      throw new Error('rendered without its code component')
    } catch (error) {
      expect(isRecipeError(error, 'missing-code')).toBe(true)
      expect((error as Error).message).toContain('review-summary')
    }
    expect(checkRecipe(compiled, fixtureData('review'), { registry }).filter((p) => p.code === 'missing-code').map((p) => p.message)).toEqual([
      'the code component "review-summary" (Summary) was not given',
    ])
  })

  it('its blocks take hand edits, list what they leave and keep their sources like any part\'s', () => {
    const edits = { 'summary|Summary|field|Checked by': 'D. Checker', 'summary|Summary|paragraph|': 'Seen from both viewpoints.' }
    const { document, toComplete, edits: result, parts } = renderRecipe(compiled, fixtureData('review'), { code, registry, edits })
    expect(result.targets.filter((t) => t.part === 'summary').map((t) => [t.key, t.original])).toEqual([
      ['summary|Summary|paragraph|', '2 of 2 viewpoints see the opening. The most is seen from Viewpoint 1, lake shore (18%); the nearest is Viewpoint 1, lake shore, 3.2 km away.'],
      ['summary|Summary|paragraph|#2', 'Above the limit. The share seen from Viewpoint 1, lake shore passes 15%.'],
      ['summary|Summary|field|Viewpoints seeing it', '2'],
      ['summary|Summary|field|Assessed by', 'C. Reviewer'],
      ['summary|Summary|field|Checked by', null],
    ])
    expect(result.orphans).toEqual([])
    expect(toComplete.some((item) => item.partId === 'summary')).toBe(false)
    expect(parts.find((p) => p.id === 'summary')?.toComplete).toBe(0)
    const summary = parts.find((p) => p.id === 'summary')
    const blocks = document.sections[1]?.body.slice(summary?.start, summary?.end) ?? []
    expect(blocks[1]).toMatchObject({ content: [{ text: 'Seen from both viewpoints.', source: 'edited' }] })
    expect(blocks[3]).toMatchObject({ rows: [{ source: 'computed' }, { value: [{ text: 'C. Reviewer', source: 'office' }] }, { value: { text: 'D. Checker', source: 'edited' } }] })
  })
})
