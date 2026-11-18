import { describe, expect, it } from 'vitest'

import {
  compileRecipe,
  defineCodeComponents,
  isRecipeError,
  type CompiledRecipe,
  type Json,
  type RecipeErrorCode,
} from '../../src/recipe'
import { comp, lib, para, recipe } from './helpers'

const fails = (code: RecipeErrorCode, run: () => unknown, path?: string | RegExp) => {
  try {
    run()
  } catch (error) {
    if (!isRecipeError(error, code)) throw error
    if (path instanceof RegExp) expect(error.message).toMatch(path)
    else if (path !== undefined) expect(error.path).toBe(path)
    return error
  }
  throw new Error(`expected a ${code} error`)
}

const greeting = comp('greeting', [para('Hello {{ params.name }}.')], {
  params: { name: { default: 'reader' }, tone: { default: 'plain' } },
  variants: {
    formal: { title: 'Formal', blocks: [para('Dear {{ params.name }},')] },
    friend: { title: 'To a friend', with: { name: 'friend' } },
  },
})

describe('compiling', () => {
  it('inlines every component into one object with no use left, marking each part', () => {
    const wrapper = comp('wrapper', [{ use: 'greeting', with: { name: '{{ params.who }}' } }], { params: { who: { required: true } } })
    const compiled = compileRecipe(recipe('r', [{ use: 'wrapper', id: 'hello', with: { who: '{{ person.name }}' } }]), lib([greeting, wrapper]))
    const hasUse = (value: unknown): boolean =>
      Array.isArray(value) ? value.some(hasUse) : !!value && typeof value === 'object' && ('use' in value || Object.values(value).some(hasUse))
    expect(hasUse(compiled.sections)).toBe(false)
    const part = compiled.sections[0]?.parts[0]
    expect(part?.part).toEqual({ id: 'hello', component: 'wrapper', title: 'Component wrapper' })
    expect(part?.params).toEqual({ who: '{{ person.name }}' })
    // The inner component, inlined, with its default filled in.
    expect(part?.blocks?.[0]).toEqual({ component: 'greeting', params: { name: '{{ params.who }}', tone: 'plain' }, blocks: [para('Hello {{ params.name }}.')] })
    expect(compiled.from.components.map((c) => c.id)).toEqual(['greeting', 'wrapper'])
    expect(compiled.from.hash).toMatch(/^sha256:[0-9a-f]{64}$/)
    expect(compiled.version).toBe(1)
  })

  it('applies a variant: another body, or other parameter values', () => {
    const formal = compileRecipe(recipe('r', [{ use: 'greeting', variant: 'formal' }]), lib([greeting]))
    expect(formal.sections[0]?.parts[0]?.blocks).toEqual([para('Dear {{ params.name }},')])
    expect(formal.sections[0]?.parts[0]?.variant).toBe('formal')
    const friend = compileRecipe(recipe('r', [{ use: 'greeting', variant: 'friend' }]), lib([greeting]))
    expect(friend.sections[0]?.parts[0]?.params.name).toBe('friend')
    // The caller's own value wins over the variant's.
    const named = compileRecipe(recipe('r', [{ use: 'greeting', variant: 'friend', with: { name: 'Ada' } }]), lib([greeting]))
    expect(named.sections[0]?.parts[0]?.params.name).toBe('Ada')
  })

  it('is deterministic: the same inputs give the same JSON text, whatever the key or library order', () => {
    const other = comp('other', [para('{{ x }}')])
    const one = JSON.stringify(compileRecipe(recipe('r', [{ use: 'greeting' }, { use: 'other' }]), lib([greeting, other])))
    const reordered = JSON.parse(JSON.stringify(greeting), (_key, value) =>
      value && typeof value === 'object' && !Array.isArray(value) ? Object.fromEntries(Object.entries(value).reverse()) : value,
    )
    const two = JSON.stringify(compileRecipe(recipe('r', [{ use: 'greeting' }, { use: 'other' }]), lib([other, reordered])))
    expect(two).toBe(one)
  })

  it('lists what the recipe reads, with [] for list items, where, and whether in a loop', () => {
    const table = comp('table', [{ type: 'table', rows: [{ for: 'row in params.rows', do: [{ cells: ['{{ row.name }}', '{{ row.value | format "0.0" }}'] }] }] }], {
      params: { rows: { required: true } },
    })
    const compiled = compileRecipe(
      recipe('r', [
        { use: 'table', id: 'per-set', for: 'set in sets', if: 'set.shown', with: { rows: '{{ set.rows }}' } },
        { use: 'greeting', with: { name: '{{ client.name | upper }}' } },
      ]),
      lib([table, greeting]),
    )
    const byPath = Object.fromEntries(compiled.needs.map((need) => [need.path, need]))
    expect(Object.keys(byPath)).toEqual(['client.name', 'sets', 'sets[].rows', 'sets[].rows[].name', 'sets[].rows[].value', 'sets[].shown'])
    expect(byPath['sets[].rows']?.uses).toEqual(['value', 'list'])
    expect(byPath.sets?.uses).toEqual(['list'])
    expect(byPath['sets[].shown']?.uses).toEqual(['condition'])
    expect(byPath['sets[].rows[].name']?.inLoop).toBe(true)
    expect(byPath['client.name']?.inLoop).toBe(false)
    expect(byPath['sets[].rows[].value']?.places[0]).toEqual({
      part: 'per-set',
      at: 'sections[0].parts[0].blocks[0].rows[0].do[0].cells[1]',
      use: 'value',
      inLoop: true,
    })
  })

  it('lists custom block kinds, filters other than the built-in ones, and code components', () => {
    const summary = { id: 'summary', title: 'Summary', version: '2', needs: ['scores[].value'], build: () => [] }
    const compiled = compileRecipe(
      recipe('r', [
        { use: 'matrix-part' },
        { use: 'summary', id: 'sum' },
      ]),
      lib([comp('matrix-part', [{ type: 'custom', kind: 'matrix', data: { cells: '{{ grid }}' } }, para('{{ who | initials }} {{ when | format "yyyy" }}')])]),
      { code: defineCodeComponents([summary]) },
    )
    expect(compiled.customKinds).toEqual(['matrix'])
    expect(compiled.filters).toEqual(['initials'])
    expect(compiled.code).toEqual([{ id: 'summary', title: 'Summary', version: '2', parts: ['sum'] }])
    expect(compiled.from.code).toEqual([{ id: 'summary', version: '2' }])
    expect(compiled.sections[0]?.parts[1]).toEqual({ part: { id: 'sum', component: 'summary', title: 'Summary' }, params: {}, code: 'summary' })
    expect(compiled.needs.find((need) => need.path === 'scores[].value')?.places[0]?.declared).toBe(true)
  })

  it('carries the JSON components a code component declares it uses', () => {
    const notice = comp('notice', [para('{{ params.text }}')], { params: { text: { required: true } } })
    const code = defineCodeComponents([{ id: 'builder', title: 'Builder', uses: ['notice'], build: () => [] }])
    const compiled = compileRecipe(recipe('r', [{ use: 'builder' }]), lib([notice]), { code })
    expect(Object.keys(compiled.embedded ?? {})).toEqual(['notice'])
    fails('unknown-component', () =>
      compileRecipe(recipe('r', [{ use: 'builder' }]), lib([]), { code: defineCodeComponents([{ id: 'builder', title: 'B', uses: ['gone'], build: () => [] }]) }),
    )
  })
})

describe('compile errors name the place', () => {
  it('an unknown component', () => {
    const error = fails('unknown-component', () => compileRecipe(recipe('r', [{ use: 'greeting' }, { use: 'nope' }]), lib([greeting])), 'recipe/r:sections[0].parts[1]')
    expect(error.message).toContain('the library has greeting')
  })

  it('an unknown variant', () => {
    fails('unknown-variant', () => compileRecipe(recipe('r', [{ use: 'greeting', variant: 'shouted' }]), lib([greeting])), 'recipe/r:sections[0].parts[0].variant')
  })

  it('a missing required parameter, and a parameter the component does not take', () => {
    const needy = comp('needy', [para('{{ params.x }}')], { params: { x: { required: true } } })
    fails('missing-param', () => compileRecipe(recipe('r', [{ use: 'needy' }]), lib([needy])), 'recipe/r:sections[0].parts[0]')
    fails('unknown-param', () => compileRecipe(recipe('r', [{ use: 'needy', with: { x: 1, y: 2 } }]), lib([needy])), 'recipe/r:sections[0].parts[0].with.y')
  })

  it('a cycle in use, and in extends', () => {
    const a = comp('a', [{ use: 'b' }])
    const b = comp('b', [para('x'), { use: 'a' }])
    const error = fails('cycle', () => compileRecipe(recipe('r', [{ use: 'a' }]), lib([a, b])), 'component/b:blocks[1]')
    expect(error.message).toContain('a → b → a')
    const one = recipe('one', [], { extends: 'two', sections: undefined } as never)
    const two = { kind: 'recipe' as const, id: 'two', name: 'Two', extends: 'one' }
    fails('cycle', () => compileRecipe('one', lib([], [{ ...one, sections: undefined } as never, two])))
    fails('unknown-recipe', () => compileRecipe('missing', lib([])))
    fails('unknown-recipe', () => compileRecipe({ kind: 'recipe', id: 'x', name: 'X', extends: 'nowhere' }, lib([])))
  })

  it('a tag that does not read, with the JSON path to it', () => {
    const bad = (blocks: Json[]) => () => compileRecipe(recipe('r', [{ use: 'bad' }]), lib([comp('bad', blocks)]))
    fails('bad-tag', bad([para('Hello {{ first name }}')]), 'component/bad:blocks[0].content')
    fails('bad-tag', bad([para('{{ if a }}open')]), /never closed/)
    fails('bad-tag', bad([para('{{ total | format "[Red]0" }}')]), /picture/)
    fails('bad-tag', bad([{ if: 'a === 1', then: [] }]), 'component/bad:blocks[0].if')
    fails('bad-tag', bad([{ for: 'x of list', do: [] }]), 'component/bad:blocks[0].for')
    fails('bad-tag', bad([para([{ type: 'sentence', template: 'It is {{ x }}.', values: {} }])]), 'component/bad:blocks[0].content[0].template')
  })

  it('a key the block model does not have, and a node that does not stand where it is', () => {
    const bad = (blocks: Json[]) => () => compileRecipe(recipe('r', [{ use: 'bad' }]), lib([comp('bad', blocks)]))
    fails('unknown-key', bad([{ type: 'heading', level: 1, conent: 'x' }]), 'component/bad:blocks[0].conent')
    fails('unknown-key', bad([{ type: 'table', rows: [{ cells: [{ contents: 'x' }] }] }]), 'component/bad:blocks[0].rows[0].cells[0].contents')
    fails('unknown-key', bad([{ if: 'a', then: [], otherwise: [] }]), 'component/bad:blocks[0].otherwise')
    fails('bad-node', bad([{ type: 'chart' }]), 'component/bad:blocks[0].type')
    fails('bad-node', bad([{ type: 'paragraph', content: 'x', spacing: { before: { if: 'a', then: [] } } }]), /only in a list/)
    fails('bad-node', bad([{ component: 'x', params: {}, blocks: [] }]), /written only by the compile/)
  })

  it('a part named twice, and a change naming a part the recipe does not have', () => {
    fails('duplicate-part', () => compileRecipe(recipe('r', [{ use: 'greeting' }, { use: 'greeting' }]), lib([greeting])), 'recipe/r:sections[0].parts[1]')
    const base = recipe('base', [{ use: 'greeting' }])
    fails('unknown-part', () => compileRecipe({ kind: 'recipe', id: 'd', name: 'D', extends: 'base', changes: [{ remove: 'nope' }] }, lib([greeting], [base])), 'recipe/d:changes[0]')
  })

  it('a JSON component with a code component\'s id', () => {
    fails('component-conflict', () =>
      compileRecipe(recipe('r', [{ use: 'greeting' }]), lib([greeting]), { code: defineCodeComponents([{ id: 'greeting', title: 'G', build: () => [] }]) }),
    )
  })

  it('a filter the application does not have, when it says which it has', () => {
    const library = lib([comp('f', [para('{{ name | initals }}')])])
    expect(() => compileRecipe(recipe('r', [{ use: 'f' }]), library)).not.toThrow()
    fails('unknown-filter', () => compileRecipe(recipe('r', [{ use: 'f' }]), library, { filters: ['initials'] }), 'component/f:blocks[0].content')
  })

  it('the limits: include depth, extends depth, template nodes', () => {
    const chain = [comp('c1', [{ use: 'c2' }]), comp('c2', [{ use: 'c3' }]), comp('c3', [para('end')])]
    expect(() => compileRecipe(recipe('r', [{ use: 'c1' }]), lib(chain))).not.toThrow()
    fails('limit', () => compileRecipe(recipe('r', [{ use: 'c1' }]), lib(chain), { limits: { includeDepth: 2 } }), /nest more than 2/)
    const recipes = [recipe('r0', [{ use: 'c3' }]), ...[1, 2, 3].map((n) => ({ kind: 'recipe' as const, id: `r${n}`, name: `R${n}`, extends: `r${n - 1}` }))]
    expect(() => compileRecipe('r3', lib(chain, recipes))).not.toThrow()
    fails('limit', () => compileRecipe('r3', lib(chain, recipes), { limits: { extendsDepth: 2 } }))
    fails('limit', () => compileRecipe(recipe('r', [{ use: 'c1' }]), lib(chain), { limits: { nodes: 3 } }))
  })
})

describe('extends', () => {
  const parts = ['a', 'b', 'c', 'd'].map((id) => comp(id, [para(id.toUpperCase())]))
  const base = recipe('base', [{ use: 'a' }, { use: 'b' }, { use: 'c' }], { options: { size: 1, tone: 'plain' }, document: { metadata: { title: 'Base', author: 'Lab' } } })
  const ids = (compiled: CompiledRecipe) => compiled.sections.map((section) => section.parts.map((part) => part.part.id))

  it('adds, removes, replaces, moves and sets parts, in order', () => {
    const derived = {
      kind: 'recipe' as const,
      id: 'derived',
      name: 'Derived',
      extends: 'base',
      options: { tone: 'warm' },
      document: { metadata: { title: 'Derived' } },
      changes: [
        { add: { use: 'd' }, after: 'a' },
        { remove: 'b' },
        { replace: 'c', by: { use: 'a', title: 'A again' } },
        { move: 'd', before: 'a' },
        { set: 'a', pageBreakBefore: true, title: 'First' },
      ],
    }
    const compiled = compileRecipe('derived', lib(parts, [base, derived]))
    expect(ids(compiled)).toEqual([['d', 'a', 'c']])
    // A replaced part keeps its id (and its hand edits) unless the replacement names its own.
    expect(compiled.sections[0]?.parts[2]?.part).toEqual({ id: 'c', component: 'a', title: 'A again' })
    expect(compiled.sections[0]?.parts[1]?.pageBreakBefore).toBe(true)
    expect(compiled.sections[0]?.parts[1]?.part.title).toBe('First')
    expect(compiled.options).toEqual({ size: 1, tone: 'warm' })
    expect(compiled.document.metadata).toEqual({ title: 'Derived', author: 'Lab' })
    expect(compiled.from.recipes.map((r) => r.id)).toEqual(['derived', 'base'])
  })

  it('sets a section and adds one', () => {
    const derived = {
      kind: 'recipe' as const,
      id: 'derived',
      name: 'Derived',
      extends: 'base',
      changes: [
        { setSection: 0, page: { orientation: 'landscape' } },
        { addSection: { id: 'appendix', parts: [{ use: 'd' }] } },
        { add: { use: 'b', id: 'b2' }, section: 'appendix' },
      ],
    }
    const compiled = compileRecipe('derived', lib(parts, [base, derived]))
    expect(compiled.sections[0]?.page).toEqual({ orientation: 'landscape' })
    expect(ids(compiled)).toEqual([['a', 'b', 'c'], ['d', 'b2']])
  })
})
