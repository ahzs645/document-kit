import { describe, expect, it } from 'vitest'

import { numberCaptions, plainText, resolveSentences, type Block, type Document, type Inline, type RichText } from '../../src'
import { compileRecipe, isRecipeError, readCompiledRecipe, renderRecipe, sharedLibrary, type Json, type RenderOptions } from '../../src/recipe'
import { comp, lib, para, recipe } from './helpers'

const body = (document: Document): readonly Block[] => document.sections[0]?.body ?? []
const text = (content: RichText | null | undefined) => (content === null || content === undefined ? null : plainText(resolveSentences(content)))

function render(blocks: Json[], data: Record<string, unknown>, options: RenderOptions & { extra?: Record<string, Json> } = {}) {
  const { extra, ...rest } = options
  const compiled = compileRecipe(recipe('r', [{ use: 'c', id: 'part' }], extra ?? {}), lib([comp('c', blocks)]))
  return renderRecipe(compiled, data, rest)
}

describe('meta', () => {
  it('carries a recipe\'s, a component\'s and a part\'s meta through the compile to the rendered parts, untouched', () => {
    const meta = { source: 'made-up', tags: ['a', 'b'], count: 3, nested: { id: 'x-1', on: true, none: null } }
    const compiled = compileRecipe(
      recipe('r', [{ use: 'c', id: 'one', meta: { order: 1, note: '{{ not a tag }}' } }], { meta }),
      lib([comp('c', [para('Hi')], { meta: { family: 'notes' } })]),
    )
    expect(compiled.recipe.meta).toEqual(meta)
    expect(compiled.sections[0]!.parts[0]!.part).toMatchObject({ meta: { order: 1, note: '{{ not a tag }}' }, componentMeta: { family: 'notes' } })
    expect(readCompiledRecipe(JSON.parse(JSON.stringify(compiled)))).toEqual(compiled)
    const { parts } = renderRecipe(compiled, {})
    expect(parts[0]).toMatchObject({ meta: { order: 1, note: '{{ not a tag }}' }, componentMeta: { family: 'notes' } })
  })

  it('refuses meta that is not a JSON object, and still refuses unknown keys elsewhere', () => {
    for (const bad of [
      () => compileRecipe(recipe('r', [{ use: 'c', meta: 'x' as never }]), lib([comp('c', [para('Hi')])])),
      () => compileRecipe(recipe('r', [{ use: 'c' }]), lib([comp('c', [para('Hi')], { meta: { n: Infinity } })])),
      () => compileRecipe(recipe('r', [{ use: 'c', metadata: {} } as never]), lib([comp('c', [para('Hi')])])),
    ]) {
      try {
        bad()
        throw new Error('expected a refusal')
      } catch (error) {
        expect(isRecipeError(error, 'invalid')).toBe(true)
      }
    }
  })
})

describe('the placeholder rule', () => {
  const blocks: Json[] = [
    { type: 'fields', rows: [{ label: 'Date', value: '{{ date }}' }, { label: 'Lot', value: '{{ lot }}' }] },
    para('Run on {{ run | missing "placeholder" "metals run date" }} by {{ who }}.'),
    para([{ type: 'sentence', label: 'Temperature', template: 'It arrived at {temp}.', values: { temp: '{{ temp }}' } }]),
    { type: 'table', rows: [{ role: 'header', cells: ['Item', 'Value'] }, { cells: ['pH', '{{ ph }}'] }] },
  ]
  const data = { lot: 'L-7' }

  it('keeps the two rules by default: blanks printed blank, a sentence left out; a value can ask for a placeholder', () => {
    const { document, toComplete } = render(blocks, data)
    const [fields, run, sentence, table] = body(document) as [Block, Block, Block, Block]
    expect(fields.type === 'fields' && fields.rows[0]!.value).toBe(null)
    expect(run.type === 'paragraph' && text(run.content)).toBe('Run on [metals run date] by ______.')
    expect(sentence.type === 'paragraph' && text(sentence.content)).toBe('')
    expect(table.type === 'table' && table.rows[1]!.cells[1]).toBe(null)
    expect(toComplete.map((item) => item.what)).toEqual(['Date', 'metals run date', 'who', 'Temperature', 'pH, Value'])
  })

  it('prints every value nobody gave as its bracketed name, marked, when the recipe says placeholder; still listed', () => {
    const { document, toComplete } = render(blocks, data, { extra: { missing: 'placeholder' } })
    const [fields, run, sentence, table] = body(document) as [Block, Block, Block, Block]
    expect(fields.type === 'fields' && text(fields.rows[0]!.value)).toBe('[Date]')
    expect(fields.type === 'fields' && text(fields.rows[1]!.value)).toBe('L-7')
    expect(run.type === 'paragraph' && text(run.content)).toBe('Run on [metals run date] by [who].')
    expect(sentence.type === 'paragraph' && text(sentence.content)).toBe('It arrived at [temp].')
    expect(table.type === 'table' && text(table.rows[1]!.cells[1] as RichText)).toBe('[ph]')
    const runs = run.type === 'paragraph' ? (run.content as Inline[]) : []
    expect(runs.filter((r) => typeof r !== 'string' && r.type === 'blank').every((r) => (r as { show?: string }).show === 'placeholder')).toBe(true)
    expect(toComplete.map((item) => item.what)).toEqual(['Date', 'metals run date', 'who', 'Temperature', 'ph'])
  })

  it('is chosen per part, and a value can ask for the blank rule inside a placeholder part', () => {
    const compiled = compileRecipe(
      recipe('r', [
        { use: 'c', id: 'plain' },
        { use: 'd', id: 'marked', missing: 'placeholder' },
      ]),
      lib([comp('c', [para('A {{ x }}.')]), comp('d', [para('B {{ x }} and {{ y | missing "blank" }}.')])]),
    )
    const out = body(renderRecipe(compiled, {}).document).map((b) => (b.type === 'paragraph' ? text(b.content) : ''))
    expect(out).toEqual(['A ______.', 'B [x] and ______.'])
  })
})

describe('pinned parts', () => {
  it('pins a part\'s blocks to the foot of the page, with room under them', () => {
    const compiled = compileRecipe(
      recipe('r', [
        { use: 'c', id: 'body' },
        { use: 'd', id: 'contact', pin: { spaceBelow: 24 } },
        { use: 'd', id: 'again', pin: 'bottom', pageBreakBefore: true },
      ]),
      lib([comp('c', [para('Body.')]), comp('d', [{ type: 'heading', level: 2, content: 'Contact' }, para('{{ phone }}')])]),
    )
    const { document, parts, toComplete } = renderRecipe(compiled, {})
    const blocks = body(document)
    expect(blocks.map((b) => b.type)).toEqual(['paragraph', 'pinned', 'pinned'])
    expect(blocks[1]).toMatchObject({ type: 'pinned', spaceBelow: 24, blocks: [{ type: 'heading' }, { type: 'paragraph' }] })
    expect(blocks[2]).toMatchObject({ pageBreakBefore: true })
    expect(parts.map((p) => [p.id, p.start, p.end])).toEqual([['body', 0, 1], ['contact', 1, 2], ['again', 2, 3]])
    // What is pinned is listed like anything else, under its heading.
    expect(toComplete.map((item) => [item.partId, item.heading, item.what])).toEqual([['contact', 'Contact', 'phone'], ['again', 'Contact', 'phone']])
  })

  it('pins a block a component writes as pinned, through the JSON', () => {
    const { document } = render([{ type: 'pinned', spaceBelow: 6, blocks: [para('Pinned {{ x }}')] }], { x: 'here' })
    expect(body(document)[0]).toEqual({ type: 'pinned', spaceBelow: 6, blocks: [{ type: 'paragraph', content: 'Pinned here' }] })
  })
})

describe('tick boxes and their hand edits', () => {
  const question: Json = { type: 'choices', label: 'Made-up rating:', options: ['Low', 'Moderate', 'High'], chosen: '{{ rating }}', answer: '{{ why }}', layout: 'inline' }

  it('ticks the option the data chose, and lists an unanswered question and answer', () => {
    const answered = render([question], { rating: 'Moderate', why: 'Because.' })
    expect(body(answered.document)[0]).toMatchObject({ type: 'choices', chosen: 'Moderate', answer: 'Because.' })
    expect(answered.toComplete).toEqual([])
    const open = render([question], {})
    expect(body(open.document)[0]).toMatchObject({ type: 'choices', answer: null })
    expect(open.toComplete.map((item) => [item.kind, item.what])).toEqual([['choice', 'Made-up rating'], ['answer', 'Made-up rating']])
  })

  it('lets a hand edit change the tick and the answer, keyed like the other edits; a tick that is no option is not applied', () => {
    const first = render([question], { rating: 'Low', why: 'Short.' })
    const [tick, answer] = first.edits.targets
    expect(tick).toMatchObject({ key: 'part||choice|Made-up rating', kind: 'choice', original: 'Low', options: ['Low', 'Moderate', 'High'] })
    expect(answer).toMatchObject({ key: 'part||answer|Made-up rating', kind: 'answer', original: 'Short.' })
    const edited = render([question], { rating: 'Low', why: 'Short.' }, { edits: { [tick!.key]: 'High', [answer!.key]: 'Reviewed again.' } })
    const block = body(edited.document)[0]!
    expect(block).toMatchObject({ type: 'choices', chosen: 'High', source: 'edited' })
    expect(block.type === 'choices' && text(block.answer)).toBe('Reviewed again.')
    const cleared = render([question], { rating: 'Low' }, { edits: { [tick!.key]: '' } })
    expect(body(cleared.document)[0]).toMatchObject({ chosen: null })
    expect(cleared.toComplete.map((item) => item.kind)).toEqual(['choice', 'answer'])
    const wrong = render([question], { rating: 'Low' }, { edits: { [tick!.key]: 'Extreme' } })
    expect(body(wrong.document)[0]).toMatchObject({ chosen: 'Low' })
    expect(wrong.edits.orphans).toEqual([tick!.key])
  })
})

describe('text rules in a recipe', () => {
  const rules: Json[] = [
    { builtIn: 'links' },
    { id: 'chem', builtIn: 'formulaDigits', formulas: ['NO3'], auto: false },
  ]

  it('carries the document\'s rules to the document, and applies named rules to one value through the rules filter', () => {
    const { document } = render([para('Nitrate as {{ unit | rules "chem" }}, not {{ unit }}.')], { unit: 'mg NO3/L' }, { extra: { document: { textRules: rules } } })
    expect(document.textRules).toEqual(rules)
    const runs = (body(document)[0] as unknown as { content: Inline[] }).content
    expect(runs.filter((r) => typeof r !== 'string' && r.type === 'text' && r.subscript).map((r) => (r as { text: string }).text)).toEqual(['3'])
    expect(text(runs)).toBe('Nitrate as mg NO3/L, not mg NO3/L.')
  })

  it('refuses a rule that does not read, naming its place, and a filter naming a rule the document lacks', () => {
    try {
      compileRecipe(recipe('r', [{ use: 'c' }], { document: { textRules: [{ pattern: '(a+)+', style: { bold: true } }] } }), lib([comp('c', [para('x')])]))
      throw new Error('expected a refusal')
    } catch (error) {
      expect(isRecipeError(error, 'invalid')).toBe(true)
      expect((error as { problems: Array<{ path: string }> }).problems[0]!.path).toBe('recipe/r:document.textRules[0].pattern')
    }
    try {
      compileRecipe(recipe('r', [{ use: 'c' }], { document: { textRules: rules } }), lib([comp('c', [para('{{ x | rules "nope" }}')])]))
      throw new Error('expected a refusal')
    } catch (error) {
      expect(isRecipeError(error, 'bad-tag')).toBe(true)
    }
  })
})

describe('the shared components\' new options', () => {
  const shared = (components: ReturnType<typeof comp>[], parts: Parameters<typeof recipe>[1]) => compileRecipe(recipe('r', parts), sharedLibrary({ components, recipes: [] }))

  it('prints an unnumbered table caption that takes no number', () => {
    const compiled = shared([], [
      { use: 'data-table', id: 'a', with: { caption: 'Plain', numbered: false, columns: ['X'], rows: [['1']] } },
      { use: 'data-table', id: 'b', with: { caption: 'Counted', columns: ['X'], rows: [['2']] } },
      { use: 'figure', id: 'f', with: { image: null, caption: 'View', numbered: false } },
      { use: 'data-table', id: 'c', with: { columns: ['X'], rows: [['3']] } },
    ])
    const { document } = renderRecipe(compiled, {})
    const [a, b, f, c] = body(document) as [Block, Block, Block, Block]
    expect(a).toMatchObject({ caption: { text: 'Plain', numbered: false } })
    expect(f).toMatchObject({ caption: { text: 'View', numbered: false } })
    expect(c.type === 'table' && c.caption).toBeUndefined()
    const numbers = numberCaptions(document)
    expect(numbers.get(a as never)).toBeUndefined()
    expect(numbers.get(b as never)).toBe(1)
  })

  it('takes header rows among the rows (a stacked block\'s own), a padding, a line height and no repeated header', () => {
    const compiled = shared([], [
      { use: 'data-table', with: { columns: ['Item', 'A'], rows: [['x', '1'], { header: ['Item', 'B'] }, ['x', '2']], padding: { top: 2, bottom: 2 }, line: 0.85, repeatHeader: false, autoFit: true } },
    ])
    const block = body(renderRecipe(compiled, {}).document)[0]!
    expect(block).toMatchObject({ type: 'table', cellPadding: { top: 2, bottom: 2 }, line: 0.85, repeatHeader: false, autoFit: true })
    expect(block.type === 'table' && block.rows.map((r) => r.role ?? 'body')).toEqual(['header', 'body', 'header', 'body'])
  })

  it('says a whole signature block through options', () => {
    const compiled = shared([], [
      {
        use: 'signature',
        with: {
          label: 'Approved by:',
          name: 'A. Person',
          title: 'Made-up Officer',
          titleLayout: 'sameLine',
          dateLine: false,
          align: 'center',
          ruleWidth: 216,
          inkGap: 37,
          imageWidth: 120,
          imageHeight: 46,
          spaceAbove: 13,
          spaceBelow: 12,
        },
      },
    ])
    expect(body(renderRecipe(compiled, {}).document)[0]).toEqual({
      type: 'signature',
      label: 'Approved by:',
      image: null,
      imageBox: { width: 120, height: 46 },
      signingSpace: 37,
      name: 'A. Person',
      title: 'Made-up Officer',
      titleLayout: 'sameLine',
      align: 'center',
      ruleWidth: 216,
      spaceAbove: 13,
      spaceBelow: 12,
      keepTogether: true,
      lines: [],
    })
  })

  it('puts two short fields on one line', () => {
    const compiled = shared([], [{ use: 'fields', with: { rows: [{ label: 'E-mail', value: 'a@example.org' }, { label: 'Phone', value: '555', sameLine: true }], rowSpacing: 3 } }])
    expect(body(renderRecipe(compiled, {}).document)[0]).toMatchObject({ type: 'fields', rowSpacing: 3, rows: [{ label: 'E-mail' }, { label: 'Phone', sameLine: true }] })
  })
})

describe('sentence filters in a recipe', () => {
  it('reads capitalize, plural, prose and ranges, and a recipe\'s own filter of the name wins', () => {
    const blocks: Json[] = [para('{{ count | plural "sample was" "samples were" | capitalize }} received: {{ ids | ranges | prose }}.')]
    const data = { count: 2, ids: ['S-1', 'S-2', 'S-3', 'S-9'] }
    expect(text((body(render(blocks, data).document)[0] as { content: RichText }).content)).toBe('Samples were received: S-1–S-3 and S-9.')
    const own = render(blocks, data, { filters: { capitalize: (v) => `<${String(v)}>` } })
    expect(text((body(own.document)[0] as { content: RichText }).content)).toBe('<samples were> received: S-1–S-3 and S-9.')
    expect(compileRecipe(recipe('r', [{ use: 'c' }]), lib([comp('c', blocks)])).filters).toEqual([])
  })
})
