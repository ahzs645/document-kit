import { describe, expect, it } from 'vitest'

import {
  compileRecipe,
  copyRecipe,
  defineCodeComponents,
  editRecipe,
  emptyOverrides,
  isRecipeError,
  layerLibraries,
  overridesFileText,
  problemsOf,
  readComponent,
  readLibrary,
  readOverridesFile,
  readRecipe,
  readRecipeFile,
  recipeFileText,
  recipePartIds,
  renderRecipe,
  sharedComponents,
  sharedLibrary,
  withOverride,
  withoutOverride,
  withRemoval,
} from '../../src/recipe'
import { comp, lib, para, recipe } from './helpers'

describe('reading JSON safely', () => {
  it('reports every problem with its path, unknown keys included, and refuses the whole', () => {
    const value = {
      kind: 'component',
      id: 'bad id',
      title: '',
      parms: {},
      params: { 'not a name': { default: 1, requird: true } },
      variants: { short: { title: 'Short' } },
      blocks: {},
    }
    expect(problemsOf(readComponent, value).map((p) => p.path)).toEqual([
      'parms',
      'id',
      'title',
      'params["not a name"]',
      'params["not a name"].requird',
      'variants.short',
      'blocks',
    ])
    try {
      readComponent(value)
      throw new Error('not refused')
    } catch (error) {
      expect(isRecipeError(error, 'invalid')).toBe(true)
      if (isRecipeError(error)) expect(error.problems).toHaveLength(7)
    }
  })

  it('checks recipes: parts, changes and a recipe with neither sections nor extends', () => {
    expect(problemsOf(readRecipe, { kind: 'recipe', id: 'r', name: 'R' }).map((p) => p.message)).toEqual(['A recipe that extends none gives its sections'])
    expect(
      problemsOf(readRecipe, {
        kind: 'recipe',
        id: 'r',
        name: 'R',
        extends: 'base',
        changes: [{ add: { use: 'x', colour: 'red' }, before: 'a', after: 'b' }, { rename: 'x' }],
      }).map((p) => p.path),
    ).toEqual(['changes[0].add.colour', 'changes[0]', 'changes[1]'])
  })

  it('refuses a file of a newer format version', () => {
    for (const read of [readComponent, readRecipe, readLibrary] as const) {
      const value = read === readLibrary ? { version: 2, components: [], recipes: [] } : read === readRecipe ? { ...recipe('r', []), version: 2 } : { ...comp('c', []), version: 2 }
      try {
        read(value)
        throw new Error('not refused')
      } catch (error) {
        expect(isRecipeError(error, 'version')).toBe(true)
      }
    }
  })

  it('refuses a library with two items of one id', () => {
    expect(problemsOf(readLibrary, lib([comp('a', []), comp('a', [])])).map((p) => p.path)).toEqual(['components[1].id'])
  })
})

describe('layering a user\'s overrides over a shipped library', () => {
  const shipped = lib([comp('a', [para('A')]), comp('b', [para('B')]), comp('c', [para('C')])], [recipe('r', [{ use: 'a' }, { use: 'b' }])])

  it('says per item whether it is shipped, modified, added or removed', () => {
    let overrides = emptyOverrides()
    overrides = withOverride(overrides, 'component', comp('a', [para('A, changed')]))
    overrides = withOverride(overrides, 'component', comp('b', [para('B')])) // the same as shipped
    overrides = withOverride(overrides, 'component', comp('d', [para('D')]))
    overrides = withRemoval(overrides, 'component', 'c')
    const layered = layerLibraries(shipped, overrides)
    expect(layered.status.components).toEqual({ a: 'modified', b: 'shipped', c: 'removed', d: 'added' })
    expect(layered.library.components.map((c) => c.id)).toEqual(['a', 'b', 'd'])
    expect(layered.status.recipes).toEqual({ r: 'shipped' })
    expect(layered.problems).toEqual([])
    // The shipped library is untouched.
    expect(shipped.components).toHaveLength(3)
    const rendered = renderRecipe(compileRecipe('r', layered.library), {})
    expect(rendered.document.sections[0]?.body).toEqual([para('A, changed'), para('B')])
    expect(layerLibraries(shipped, withoutOverride(overrides, 'component', 'a')).status.components.a).toBe('shipped')
  })

  it('names a recipe left using a removed component, and an override that takes a code component\'s id', () => {
    const removed = layerLibraries(shipped, withRemoval(emptyOverrides(), 'component', 'b'))
    expect(removed.problems).toEqual(['The recipe "r" uses "b", which is not in the library'])
    const code = defineCodeComponents([{ id: 'summary', title: 'Summary', build: () => [] }])
    const clash = layerLibraries(shipped, withOverride(emptyOverrides(), 'component', comp('summary', [para('mine')])), { code })
    expect(clash.problems).toEqual(['The JSON component "summary" has a code component\'s id; an override cannot replace a code component (give it another id)'])
    expect(() => compileRecipe(recipe('x', [{ use: 'summary' }]), clash.library, { code })).toThrow(/both a JSON component in the library and a code component/)
  })

  it('saves and loads overrides as a file, the same changes always writing the same text', () => {
    const overrides = withRemoval(withOverride(emptyOverrides(), 'recipe', recipe('mine', [{ use: 'a' }])), 'component', 'c')
    const text = overridesFileText(overrides)
    expect(readOverridesFile(text)).toEqual(JSON.parse(text))
    expect(overridesFileText(readOverridesFile(text))).toBe(text)
    expect(() => readOverridesFile('{"kind":"library-overrides","version":9}')).toThrow(/version 9/)
    expect(() => readOverridesFile('not json')).toThrow(/not JSON/)
  })
})

describe('own recipes', () => {
  const base = recipe('lab', [{ use: 'a' }, { use: 'b' }], { name: 'Lab report' })
  const taken = new Set(['lab'])
  const random = () => 0.5

  it('copies a recipe, or extends it, as one\'s own', () => {
    const copy = copyRecipe(base, taken, { random })
    expect(copy.id).toMatch(/^own-lab-report-your-copy-[0-9a-z]{4}$/)
    expect(copy.name).toBe('Lab report (your copy)')
    expect(copy.own).toEqual({ basedOn: 'lab' })
    expect(recipePartIds(copy)).toEqual(['a', 'b'])
    const extension = copyRecipe(base, taken, { as: 'extension', random })
    expect(extension).toMatchObject({ extends: 'lab', own: { basedOn: 'lab' } })
    expect(extension.sections).toBeUndefined()
  })

  it('edits one: changes apply to its own sections, or are kept for its base', () => {
    const copy = copyRecipe(base, taken, { random })
    const edited = editRecipe(copy, { name: 'Mine', changes: [{ remove: 'b' }, { add: { use: 'c' }, before: 'a' }] })
    expect(recipePartIds(edited)).toEqual(['c', 'a'])
    expect(edited.name).toBe('Mine')
    const extension = editRecipe(copyRecipe(base, taken, { as: 'extension', random }), { changes: [{ remove: 'b' }] })
    expect(extension.changes).toEqual([{ remove: 'b' }])
    const library = lib([comp('a', [para('A')]), comp('b', [para('B')]), comp('c', [para('C')])], [base, extension])
    expect(compileRecipe(extension.id, library).sections[0]?.parts.map((p) => p.part.id)).toEqual(['a'])
    expect(() => editRecipe(copy, { name: '' })).toThrow(/name/)
  })

  it('writes and reads a recipe file, never overwriting one already there', () => {
    const copy = copyRecipe(base, taken, { random })
    const text = recipeFileText(copy, new Date('2026-05-01T00:00:00Z'))
    expect(readRecipeFile(text, { ids: new Set(), names: new Set() })).toEqual(copy)
    const again = readRecipeFile(text, { ids: new Set([copy.id]), names: new Set([copy.name]) }, () => 0.25)
    expect(again.id).not.toBe(copy.id)
    expect(again.name).toBe('Lab report (your copy) (imported)')
    expect(() => readRecipeFile(text.replace('"version": 1', '"version": 3'), { ids: new Set(), names: new Set() })).toThrow(/version 3/)
    expect(() => readRecipeFile('{"kind":"other"}', { ids: new Set(), names: new Set() })).toThrow(/not a recipe file/)
  })
})

describe('the shared components', () => {
  it('every one reads, and a recipe using all of them compiles and renders', () => {
    const components = sharedComponents()
    expect(components.map((c) => c.id)).toEqual([
      'cover-page',
      'contents',
      'fields',
      'text-section',
      'notice',
      'data-table',
      'figure',
      'figure-page',
      'spreadsheet',
      'signature',
      'to-complete',
      'page-footer',
      'running-header',
    ])
    const parts = [
      { use: 'cover-page', with: { title: 'T', lines: [{ label: 'L', value: '{{ v }}' }] } },
      { use: 'contents' },
      { use: 'fields', with: { rows: [{ label: 'A', value: '{{ a }}' }] } },
      { use: 'text-section', with: { heading: 'H', paragraphs: ['One.', 'Two.'] } },
      { use: 'notice', with: { text: 'Note.' } },
      { use: 'data-table', with: { columns: ['X', 'Y'], rows: [[1, 2], { group: 'G' }, [3, null], { note: 'n' }] } },
      { use: 'figure', with: { image: '{{ picture }}' } },
      { use: 'figure-page', with: { title: 'Sheet', image: '{{ picture }}' } },
      { use: 'signature', with: { name: '{{ signer }}' } },
      { use: 'to-complete' },
    ]
    const withRunning = { ...recipe('all', parts), sections: [{ header: [{ use: 'running-header', with: { left: 'L', right: 'R' } }], footer: [{ use: 'page-footer', with: { text: 'Ref' } }], parts }] }
    const compiled = compileRecipe(withRunning, sharedLibrary())
    const { toComplete, document } = renderRecipe(compiled, { v: 'x', a: 'y' })
    expect(toComplete.map((item) => [item.kind, item.what])).toEqual([
      ['cell', '3, Y'],
      ['figure', 'Picture to come'],
      ['figure', 'Picture to come'],
      ['signature', 'Signed'],
      ['signature', 'Title'],
      ['signature', 'Date'],
    ])
    const body = document.sections[0]?.body ?? []
    expect(body[body.length - 1]).toEqual({
      type: 'list',
      items: ['H: 3, Y', 'H: Picture to come', 'Sheet: Picture to come', 'Sheet: Signed', 'Sheet: Title', 'Sheet: Date'],
    })
  })

  it('passes a data table a line height in points as well as a multiple', () => {
    const lineOf = (line: number | { exact: number } | undefined) => {
      const parts = [{ use: 'data-table', with: { columns: ['X'], rows: [[1]], ...(line === undefined ? {} : { line }) } }]
      const { document } = renderRecipe(compileRecipe(recipe('lines', parts), sharedLibrary()), {})
      const table = (document.sections[0]?.body ?? []).find((block) => block.type === 'table')
      return table?.type === 'table' ? table.line : 'no table'
    }
    expect(lineOf({ exact: 12 })).toEqual({ exact: 12 })
    expect(lineOf(1.1)).toBe(1.1)
    expect(lineOf(undefined)).toBeUndefined()
  })
})
