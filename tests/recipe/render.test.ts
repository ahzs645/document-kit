import { describe, expect, it } from 'vitest'

import { defineCustomBlocks, plainText, resolveSentences, type Block, type Document, type RichText } from '../../src'
import {
  checkRecipe,
  compileRecipe,
  defineCodeComponents,
  isRecipeError,
  readCompiledRecipe,
  renderRecipe,
  type CodeComponent,
  type Json,
  type RecipeErrorCode,
  type RenderOptions,
} from '../../src/recipe'
import { comp, lib, para, recipe } from './helpers'

/** Compiles one component's blocks as a one-part recipe and renders it. */
function render(blocks: Json[], data: Record<string, unknown>, options: RenderOptions & { params?: Record<string, { default?: Json; required?: boolean }>; with?: Record<string, Json> } = {}) {
  const { params, with: given, ...rest } = options
  const compiled = compileRecipe(recipe('r', [{ use: 'c', id: 'part', ...(given ? { with: given } : {}) }]), lib([comp('c', blocks, params ? { params } : {})]), rest.code ? { code: rest.code } : {})
  return renderRecipe(compiled, data, rest)
}
const body = (document: Document): readonly Block[] => document.sections[0]?.body ?? []
const text = (content: RichText | null | undefined) => (content === null || content === undefined ? null : plainText(resolveSentences(content)))
const paragraphs = (document: Document) => body(document).map((block) => (block.type === 'paragraph' ? text(block.content) : `[${block.type}]`))

const fails = (code: RecipeErrorCode, run: () => unknown, match?: RegExp) => {
  try {
    run()
  } catch (error) {
    if (!isRecipeError(error, code)) throw error
    if (match) expect(error.message).toMatch(match)
    return error
  }
  throw new Error(`expected a ${code} error`)
}

describe('conditions and loops', () => {
  it('keeps a branch by a condition: comparisons, and, or, not, filters', () => {
    const { document } = render(
      [
        { if: 'status == "final" and total > 100', then: [para('final and large')], else: [para('other')] },
        { if: 'not draft or urgent', then: [para('shown')] },
        { if: 'missing', then: [para('never')], else: [para('missing is false')] },
        { if: 'name | upper == "ADA"', then: [para('filtered')] },
        { if: 'count <= 2', then: [para('few')] },
      ],
      { status: 'final', total: 250, draft: true, urgent: false, name: 'Ada', count: 2 },
    )
    expect(paragraphs(document)).toEqual(['final and large', 'missing is false', 'filtered', 'few'])
  })

  it('repeats per item, with loop.index, loop.number, loop.first and loop.last, and an empty branch', () => {
    const { document } = render(
      [
        { for: 'item in items', do: [para('{{ loop.number }}/{{ loop.length }} {{ item.name }}{{ if loop.first }} (first){{ endif }}{{ if loop.last }} (last){{ endif }} at {{ loop.index }}')] },
        { for: 'item in nothing', do: [para('never')], empty: [para('none')] },
      ],
      { items: [{ name: 'a' }, { name: 'b' }, { name: 'c' }], nothing: [] },
    )
    expect(paragraphs(document)).toEqual(['1/3 a (first) at 0', '2/3 b at 1', '3/3 c (last) at 2', 'none'])
  })

  it('reads inline if and for inside one string, as a Word template does', () => {
    const { document } = render([para('Tags: {{ for t in tags }}{{ t }}{{ if not loop.last }}, {{ endif }}{{ endfor }}. {{ if ok }}Approved{{ else }}Pending{{ endif }}.')], {
      tags: ['red', 'blue'],
      ok: false,
    })
    expect(paragraphs(document)).toEqual(['Tags: red, blue. Pending.'])
  })

  it('puts control nodes where rows, cells, list items and runs stand', () => {
    const { document } = render(
      [
        {
          type: 'table',
          rows: [
            { role: 'header', cells: [{ for: 'c in columns', do: ['{{ c }}'] }] },
            { for: 'r in rows', do: [{ cells: [{ for: 'v in r', do: ['{{ v }}'] }] }] },
            { if: 'footnote', then: [{ role: 'note', cells: ['{{ footnote }}'] }] },
          ],
        },
        { type: 'list', items: [{ for: 'x in list', do: ['{{ x }}'] }, { if: 'extra', then: ['extra'] }] },
        para(['Hello', { if: 'name', then: [' ', { type: 'text', text: '{{ name }}', bold: true }] }, '.']),
      ],
      { columns: ['A', 'B'], rows: [[1, 2], [3, 4]], footnote: 'n.', list: ['x', 'y'], extra: true, name: 'Ada' },
    )
    const [table, list, line] = body(document)
    expect(table?.type === 'table' && table.rows.map((row) => row.cells.map((cell) => (typeof cell === 'string' ? cell : JSON.stringify(cell))))).toEqual([
      ['A', 'B'],
      ['1', '2'],
      ['3', '4'],
      ['n.'],
    ])
    expect(list?.type === 'list' && list.items).toEqual(['x', 'y', 'extra'])
    expect(line?.type === 'paragraph' && line.content).toEqual(['Hello ', { type: 'text', text: 'Ada', bold: true }, '.'])
  })

  it('uses a component where rows stand', () => {
    const totals = comp('total-rows', [{ role: 'spacer', cells: [] }, { cells: ['Total', '{{ params.total }}'] }], { params: { total: { required: true } } })
    const table = comp('t', [{ type: 'table', rows: [{ cells: ['a', '1'] }, { use: 'total-rows', with: { total: '{{ sum }}' } }] }])
    const compiled = compileRecipe(recipe('r', [{ use: 't' }]), lib([totals, table]))
    const [block] = body(renderRecipe(compiled, { sum: 3 }).document)
    expect(block?.type === 'table' && block.rows).toEqual([{ cells: ['a', '1'] }, { role: 'spacer', cells: [] }, { cells: ['Total', '3'] }])
  })

  it('a string that is exactly one tag gives the raw value: a list, a number, rich text', () => {
    const { document } = render(
      [{ type: 'heading', level: '{{ level }}', content: '{{ title }}' }, { type: 'table', rows: '{{ rows }}', columnWidths: '{{ widths }}' }, para('{{ runs }}')],
      { level: 3, title: 'Results', rows: [{ cells: ['a', 'b'] }], widths: [2, 1], runs: [{ type: 'text', text: 'bold', bold: true }] },
    )
    expect(body(document)).toEqual([
      { type: 'heading', level: 3, content: 'Results' },
      { type: 'table', rows: [{ cells: ['a', 'b'] }], columnWidths: [2, 1] },
      { type: 'paragraph', content: [{ type: 'text', text: 'bold', bold: true }] },
    ])
  })

  it('formats with the Word engine\'s pictures and the application\'s filters', () => {
    const { document } = render([para('{{ total | format "#,##0.00" }} on {{ day | format "d MMMM yyyy" }}, {{ name | initials }}')], { total: 1234.5, day: '2026-03-05', name: 'Ada Lovelace' }, {
      filters: { initials: (value) => String(value).split(' ').map((word) => word[0]).join('') },
    })
    expect(paragraphs(document)).toEqual(['1,234.50 on 5 March 2026, AL'])
    fails('unknown-filter', () => render([para('{{ name | initials }}')], { name: 'x' }))
    fails('bad-value', () => render([para('{{ name | format "0.0" }}')], { name: 'not a number' }), /not a number/)
  })

  it('refuses a list or an object printed where text goes', () => {
    fails('not-text', () => render([para('Items: {{ items }}')], { items: [1, 2] }), /a list/)
  })
})

describe('missing values', () => {
  const fieldsAndTable: Json[] = [
    { type: 'heading', level: 1, content: 'Report' },
    { type: 'fields', rows: [{ label: 'Client', value: '{{ client }}' }, { label: 'Reference', value: '{{ reference }}' }] },
    { type: 'table', caption: 'Readings', rows: [{ role: 'header', cells: ['Item', 'Value'] }, { for: 'r in rows', do: [{ cells: ['{{ r.name }}', '{{ r.value }}'] }] }] },
    para('Received on {{ received }}.'),
    para([
      { type: 'sentence', label: 'Aspect', template: 'The site faces {aspect}.', values: { aspect: '{{ aspect }}' } },
      ' ',
      { type: 'sentence', template: 'It slopes {slope}.', values: { slope: '{{ slope }}' } },
    ]),
    { type: 'fields', rows: [{ label: 'Pages', value: 'x' }], size: '{{ size }}' },
  ]

  it('a field, a cell or text prints blank and is listed; a sentence with a missing slot is not printed and is listed', () => {
    const { document, toComplete, parts } = render(fieldsAndTable, { client: 'Example Board', reference: null, rows: [{ name: 'pH', value: 7.1 }, { name: 'Iron' }], slope: 'gently' })
    const [, fields, table, received, sentences] = body(document)
    expect(fields?.type === 'fields' && fields.rows.map((row) => row.value)).toEqual(['Example Board', null])
    expect(table?.type === 'table' && table.rows[2]?.cells).toEqual(['Iron', null])
    expect(received?.type === 'paragraph' && received.content).toEqual(['Received on ', { type: 'blank', what: 'received' }, '.'])
    expect(sentences?.type === 'paragraph' && text(sentences.content)).toBe('It slopes gently.')
    expect(toComplete.map((item) => [item.kind, item.what, item.missing ?? null, item.heading ?? null, item.partId])).toEqual([
      ['field', 'Reference', null, 'Report', 'part'],
      ['cell', 'Readings, Iron, Value', null, 'Report', 'part'],
      ['blank', 'received', null, 'Report', 'part'],
      ['sentence', 'Aspect', ['aspect'], 'Report', 'part'],
      ['value', '{{ size }}'.slice(3, -3), ['size'], null, 'part'],
    ])
    expect(parts[0]?.toComplete).toBe(5)
  })

  it('a condition over a missing value is false, and a missing list repeats nothing', () => {
    const { document, toComplete } = render([{ if: 'flag', then: [para('a')] }, { for: 'x in absent', do: [para('b')], empty: [para('none')] }], {})
    expect(paragraphs(document)).toEqual(['none'])
    expect(toComplete).toEqual([])
  })

  it('checkRecipe reports what the data lacks, by data path, without making the document', () => {
    const compiled = compileRecipe(recipe('r', [{ use: 'c' }]), lib([comp('c', fieldsAndTable)]))
    const problems = checkRecipe(compiled, { client: 'x', rows: [{ name: 'pH', value: 7 }, { name: 'Iron' }] })
    expect(problems.map((p) => [p.code, p.path, p.use])).toEqual([
      ['missing', 'reference', 'value'],
      ['missing', 'rows.1.value', 'value'],
      ['missing', 'received', 'value'],
      ['missing', 'aspect', 'value'],
      ['missing', 'slope', 'value'],
      ['missing', 'size', 'value'],
    ])
    expect(checkRecipe(compiled, { client: 'x', reference: 'r', rows: [], received: 'today', aspect: 'north', slope: 'gently', size: 10 })).toEqual([])
  })
})

describe('provenance', () => {
  it('a sources map by path (with [] and *) and wrapped values give each run its source', () => {
    const { document } = render(
      [
        { type: 'fields', rows: [{ label: 'Client', value: '{{ client.name }}' }, { label: 'Checked', value: '{{ checker }}' }, { label: 'Note', value: '{{ note }}' }] },
        { type: 'table', rows: [{ for: 'r in rows', do: [{ cells: ['{{ r.name }}', '{{ r.value }}'] }] }] },
        para('By {{ author }}.'),
      ],
      { client: { name: 'Example Board' }, rows: [{ name: 'pH', value: 7.1 }], author: { $value: 'A. Reviewer', $source: 'reviewer' }, note: 'plain' },
      { sources: { client: 'office', 'rows[].value': 'instrument', checker: 'reviewer' } },
    )
    const [fields, table, line] = body(document)
    expect(fields?.type === 'fields' && fields.rows).toEqual([
      { label: 'Client', value: [{ type: 'text', text: 'Example Board', source: 'office' }] },
      // A blank carries the source its value would have had.
      { label: 'Checked', value: null, source: 'reviewer' },
      { label: 'Note', value: 'plain' },
    ])
    expect(table?.type === 'table' && table.rows[0]?.cells).toEqual(['pH', [{ type: 'text', text: '7.1', source: 'instrument' }]])
    expect(line?.type === 'paragraph' && line.content).toEqual(['By ', { type: 'text', text: 'A. Reviewer', source: 'reviewer' }, '.'])
  })

  it('a source travels through a component parameter and a sentence', () => {
    const inner = comp('inner', [para([{ type: 'sentence', template: 'Seen from {where}.', values: { where: '{{ params.where }}' } }]), para('{{ params.where }}')], { params: { where: { required: true } } })
    const compiled = compileRecipe(recipe('r', [{ use: 'inner', with: { where: '{{ vp }}' } }]), lib([inner]))
    const { document } = renderRecipe(compiled, { vp: 'the shore' }, { sources: { vp: 'run' } })
    expect(body(document)[0]).toMatchObject({ content: [{ type: 'sentence', source: 'run' }] })
    expect(body(document)[1]).toEqual({ type: 'paragraph', content: [{ type: 'text', text: 'the shore', source: 'run' }] })
  })
})

describe('hand edits', () => {
  const blocks: Json[] = [
    { type: 'heading', level: 2, content: 'Results' },
    { type: 'fields', rows: [{ label: 'Client', value: '{{ client }}' }, { label: 'Reference', value: '{{ reference }}' }] },
    { type: 'table', caption: 'Readings', rows: [{ role: 'header', cells: ['Item', 'Value'] }, { for: 'r in rows', do: [{ cells: ['{{ r.name }}', '{{ r.value }}'] }] }] },
    para('Summary {{ client }}.'),
  ]
  const data = { client: 'Example Board', rows: [{ name: 'pH', value: 7.1 }, { name: 'Iron', value: 0.1 }] }

  it('lists every editable value with a stable key', () => {
    const { edits } = render(blocks, data)
    expect(edits.targets.map((t) => [t.key, t.original])).toEqual([
      ['part|Results|field|Client', 'Example Board'],
      ['part|Results|field|Reference', null],
      ['part|Results|cell|Readings › pH › Item', 'pH'],
      ['part|Results|cell|Readings › pH › Value', '7.1'],
      ['part|Results|cell|Readings › Iron › Item', 'Iron'],
      ['part|Results|cell|Readings › Iron › Value', '0.1'],
      ['part|Results|paragraph|', 'Summary Example Board.'],
    ])
  })

  it('applies edits after the part filled its blocks, marked, and keeps an edit whose place is gone as an orphan', () => {
    const { document, edits, toComplete } = render(blocks, data, {
      edits: {
        'part|Results|field|Reference': 'REF-7',
        'part|Results|cell|Readings › Iron › Value': '0.12',
        'part|Results|paragraph|': '',
        'part|Results|cell|Readings › Lead › Value': '3',
        'other-part|x|field|y': 'kept',
      },
      editSource: 'reviewer',
    })
    const [, fields, table, summary] = body(document)
    expect(fields?.type === 'fields' && fields.rows[1]).toEqual({ label: 'Reference', value: { type: 'text', text: 'REF-7', source: 'reviewer' }, source: 'reviewer' })
    expect(table?.type === 'table' && table.rows[2]?.cells[1]).toEqual({ content: { type: 'text', text: '0.12', source: 'reviewer' }, source: 'reviewer' })
    expect(summary?.type === 'paragraph' && summary.content).toBe('')
    expect(edits.targets.filter((t) => t.edited !== undefined).map((t) => t.key)).toEqual([
      'part|Results|field|Reference',
      'part|Results|cell|Readings › Iron › Value',
      'part|Results|paragraph|',
    ])
    expect(edits.orphans).toEqual(['part|Results|cell|Readings › Lead › Value'])
    expect(edits.otherParts).toEqual(['other-part|x|field|y'])
    // The blank the edit filled is no longer to complete.
    expect(toComplete).toEqual([])
  })

  it('keys a row by its label, so removing a row never moves an edit onto the next', () => {
    const edit = { 'part|Results|cell|Readings › Iron › Value': '0.12' }
    const { edits } = render(blocks, { ...data, rows: [{ name: 'Iron', value: 0.1 }] }, { edits: edit })
    expect(edits.targets.find((t) => t.edited)?.label).toBe('Results › Readings › Iron › Value')
    const without = render(blocks, { ...data, rows: [{ name: 'pH', value: 7 }] }, { edits: edit })
    expect(without.edits.orphans).toEqual(['part|Results|cell|Readings › Iron › Value'])
  })
})

describe('parts', () => {
  it('reports each part: blocks printed, values left, repeats, where it is', () => {
    const library = lib([comp('title', [para('{{ params.text }}')], { params: { text: { required: true } } }), comp('blank', [{ type: 'fields', rows: [{ label: 'X', value: '{{ x }}' }] }])])
    const compiled = compileRecipe(
      recipe('r', [
        { use: 'title', with: { text: 'Hello' }, pageBreakBefore: true },
        { use: 'title', id: 'per-item', for: 'item in items', with: { text: '{{ item }}' } },
        { use: 'blank', if: 'show' },
        { use: 'blank', id: 'blank-shown', if: 'not show' },
      ]),
      library,
    )
    const { parts, document } = renderRecipe(compiled, { items: ['a', 'b'], show: false })
    expect(parts.map((p) => [p.id, p.blocks, p.toComplete, p.repeats, p.start, p.end])).toEqual([
      ['title', 1, 0, 1, 0, 1],
      ['per-item', 2, 0, 2, 1, 3],
      ['blank', 0, 0, 0, 3, 3],
      ['blank-shown', 1, 1, 1, 3, 4],
    ])
    expect(body(document)[0]).toEqual({ type: 'paragraph', content: 'Hello', pageBreakBefore: true })
  })

  it('fills the to-complete component with everything else left, and the custom kinds through the registry', () => {
    const registry = defineCustomBlocks({ check: (data) => [{ type: 'fields', rows: [{ label: 'Checked', value: (data as { by: string | null }).by }] }] })
    const library = lib([
      comp('form', [{ type: 'heading', level: 1, content: 'Form' }, { type: 'custom', kind: 'check', data: { by: '{{ checker }}' } }, para('Date {{ date }}')]),
      comp('todo', [{ type: 'custom', kind: 'recipe.toComplete', data: { none: 'All done.' } }]),
    ])
    const compiled = compileRecipe(recipe('r', [{ use: 'form' }, { use: 'todo' }]), library)
    expect(compiled.customKinds).toEqual(['check'])
    const { document, toComplete } = renderRecipe(compiled, {}, { registry })
    expect(toComplete.map((item) => item.what)).toEqual(['date'])
    expect(body(document)[3]).toEqual({ type: 'list', items: ['Form: date'] })
    expect(body(renderRecipe(compiled, { date: 'today', checker: 'A' }, { registry }).document)[3]).toEqual({ type: 'paragraph', content: 'All done.' })
    fails('missing-custom-kind', () => renderRecipe(compiled, {}, { registry: defineCustomBlocks({}) }))
  })
})

describe('code components', () => {
  const tally: CodeComponent = {
    id: 'tally',
    title: 'Tally',
    params: { label: { default: 'Total' } },
    needs: ['scores[]'],
    uses: ['note'],
    build: ({ params, value }) => {
      const scores = (value('scores') as number[] | undefined) ?? []
      const total = scores.reduce((sum, n) => sum + n, 0)
      return [
        { type: 'fields', rows: [{ label: String(params.label), value: String(total), source: 'computed' }, { label: 'Checked by', value: '{{ checker }}' }] },
        ...(total > 10 ? [{ use: 'note', with: { text: 'High: {{ scores.length }} scores.' } }] : []),
      ]
    },
  }
  const note = comp('note', [para('{{ params.text }}')], { params: { text: { required: true } } })
  const code = defineCodeComponents([tally])
  const compiled = compileRecipe(recipe('r', [{ use: 'tally', id: 'sum', with: { label: 'Sum' } }]), lib([note]), { code })

  it('builds blocks that go through tags, the JSON components it declared, to-complete, provenance and edits', () => {
    const { document, toComplete, edits } = renderRecipe(compiled, { scores: [4, 9] }, { code, edits: { 'sum||field|Sum': '14' } })
    expect(body(document)).toEqual([
      { type: 'fields', rows: [{ label: 'Sum', value: { type: 'text', text: '14', source: 'edited' }, source: 'edited' }, { label: 'Checked by', value: null }] },
      { type: 'paragraph', content: 'High: 2 scores.' },
    ])
    expect(toComplete.map((item) => [item.partId, item.what])).toEqual([['sum', 'Checked by']])
    expect(edits.targets.map((t) => t.key)).toEqual(['sum||field|Sum', 'sum||field|Checked by', 'sum||paragraph|'])
  })

  it('fails loudly without it, and check names it', () => {
    fails('missing-code', () => renderRecipe(compiled, { scores: [] }), /"tally"/)
    expect(checkRecipe(compiled, { scores: [] }).map((p) => p.code)).toContain('missing-code')
  })

  it('refuses a code component that uses itself through what it returns', () => {
    const loop = defineCodeComponents([{ id: 'loop', title: 'Loop', build: () => [{ use: 'loop' }] }])
    fails('cycle', () => renderRecipe(compileRecipe(recipe('r', [{ use: 'loop' }]), lib([]), { code: loop }), {}, { code: loop }), /loop → loop/)
  })

  it('refuses blocks that use a component it did not declare, and a build that throws', () => {
    const sneaky = defineCodeComponents([{ id: 'sneaky', title: 'S', build: () => [{ use: 'note', with: { text: 'x' } }] }])
    const one = compileRecipe(recipe('r', [{ use: 'sneaky' }]), lib([note]), { code: sneaky })
    fails('undeclared-use', () => renderRecipe(one, {}, { code: sneaky }))
    const broken = defineCodeComponents([{ id: 'broken', title: 'B', build: () => { throw new Error('no data') } }])
    fails('code-failed', () => renderRecipe(compileRecipe(recipe('r', [{ use: 'broken' }]), lib([]), { code: broken }), {}, { code: broken }), /no data/)
  })
})

describe('limits and versions', () => {
  it('stops a render past its loop passes or its blocks', () => {
    const blocks: Json[] = [{ for: 'x in items', do: [para('{{ x }}')] }]
    const items = Array.from({ length: 20 }, (_, i) => i)
    expect(() => render(blocks, { items })).not.toThrow()
    fails('limit', () => render(blocks, { items }, { limits: { loopPasses: 10 } }), /more than 10/)
    fails('limit', () => render(blocks, { items }, { limits: { blocks: 10 } }), /more than 10 blocks/)
  })

  it('refuses a compiled recipe of a newer version', () => {
    const compiled = compileRecipe(recipe('r', [{ use: 'c' }]), lib([comp('c', [para('x')])]))
    fails('version', () => readCompiledRecipe({ ...compiled, version: 2 }))
    fails('version', () => renderRecipe({ ...compiled, version: 2 }, {}))
    expect(readCompiledRecipe(JSON.parse(JSON.stringify(compiled)))).toEqual(compiled)
  })
})
