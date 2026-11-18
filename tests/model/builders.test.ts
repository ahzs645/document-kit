import { describe, expect, it } from 'vitest'
import {
  blank,
  bold,
  caption,
  custom,
  field,
  fields,
  figure,
  groupRow,
  heading,
  item,
  link,
  list,
  noteRow,
  orderedList,
  pageCount,
  pageNumber,
  paragraph,
  sentence,
  sentences,
  spacerRow,
  sub,
  sup,
  table,
  text,
  title,
} from '../../src/index'

describe('builders', () => {
  it('build plain data', () => {
    expect(heading(2, 'Scope')).toEqual({ type: 'heading', level: 2, content: 'Scope' })
    expect(paragraph(['H', sub('2'), 'O ', sup('a')], { align: 'justify' })).toEqual({
      type: 'paragraph',
      align: 'justify',
      content: ['H', { type: 'text', text: '2', subscript: true }, 'O ', { type: 'text', text: 'a', superscript: true }],
    })
    expect(title('Report')).toMatchObject({ type: 'paragraph', style: 'title' })
    expect(bold('x')).toEqual({ type: 'text', text: 'x', bold: true })
    expect(link('site', 'https://example.org')).toEqual({ type: 'text', text: 'site', link: 'https://example.org' })
    expect(text('v', { source: 'run' })).toEqual({ type: 'text', text: 'v', source: 'run' })
    expect([pageNumber(), pageCount()]).toEqual([
      { type: 'field', field: 'page' },
      { type: 'field', field: 'pages' },
    ])
    expect(blank()).toEqual({ type: 'blank' })
    expect(orderedList([item('a', ['b'])])).toEqual({
      type: 'list',
      ordered: true,
      items: [{ content: 'a', items: ['b'] }],
    })
    expect(list(['x'])).toEqual({ type: 'list', items: ['x'] })
    expect(figure(null, { caption: caption('Map', false) })).toEqual({
      type: 'figure',
      image: null,
      caption: { text: 'Map', numbered: false },
    })
    expect(custom('matrix', { a: 1 })).toEqual({ type: 'custom', kind: 'matrix', data: { a: 1 } })
  })

  it('treat a missing or empty field value as a blank to complete', () => {
    expect(field('Date', undefined).value).toBeNull()
    expect(field('Date', '').value).toBeNull()
    expect(field('Date', '3 May', { source: 'lab' })).toEqual({ label: 'Date', value: '3 May', source: 'lab' })
    expect(fields([field('A', 'b')], { layout: 'table' })).toMatchObject({ type: 'fields', layout: 'table' })
  })

  it('give rows their roles', () => {
    const t = table([groupRow(['Metals']), noteRow('Below detection limit'), spacerRow()])
    expect(t.rows.map((r) => r.role)).toEqual(['group', 'note', 'spacer'])
    expect(t.rows[1]?.cells).toEqual(['Below detection limit'])
  })

  it('join sentences with spaces', () => {
    const s = sentence('A {x}.', { x: 1 })
    expect(sentences(s, 'B.')).toEqual([s, ' ', 'B.'])
  })
})
