import { describe, expect, it } from 'vitest'
import {
  blank,
  caption,
  custom,
  defineCustomBlocks,
  document,
  field,
  fields,
  figure,
  headerRow,
  heading,
  paragraph,
  row,
  section,
  sentence,
  sentences,
  signature,
  table,
  toComplete,
} from '../../src/index'

describe('toComplete', () => {
  it('lists every value nobody gave, with the heading it falls under', () => {
    const doc = document([
      section([
        heading(1, 'Site'),
        fields([field('Inspector', null, { source: 'reviewer' }), field('Area', '12 ha')]),
        paragraph(['Checked by ', blank('checker initials'), '.']),
        heading(1, 'Results'),
        table([headerRow(['Item', 'Value']), row(['Depth', null])], { caption: caption('Readings') }),
        paragraph(
          sentences(
            sentence('The plot is {slope}% slope.', { slope: 12 }),
            sentence('It faces {aspect} and sits at {elevation} m.', { aspect: null, elevation: '' }, { label: 'Aspect statement' }),
          ),
        ),
        figure(null, { caption: caption('Overview photograph') }),
        signature({ label: 'Approved by', name: null, lines: [field('Date', null)] }),
      ]),
    ])
    const items = toComplete(doc)
    expect(items.map((item) => [item.kind, item.what, item.heading])).toEqual([
      ['field', 'Inspector', 'Site'],
      ['blank', 'checker initials', 'Site'],
      ['cell', 'Readings, Depth, Value', 'Results'],
      ['sentence', 'Aspect statement', 'Results'],
      ['figure', 'Overview photograph', 'Results'],
      ['signature', 'Approved by', 'Results'],
      ['signature', 'Date', 'Results'],
    ])
    expect(items[0]?.source).toBe('reviewer')
    expect(items[3]?.missing).toEqual(['aspect', 'elevation'])
  })

  it('names a blank cell by the header over its grid column, past merged cells', () => {
    const doc = document([
      section([
        table([
          headerRow([{ content: 'Item', rowSpan: 2 }, { content: 'Values', colSpan: 2 }]),
          headerRow(['V1', 'V2']),
          row(['Item 1', '3', null]),
          row([{ content: 'Both', rowSpan: 2 }, null, '4']),
          row(['5', { content: null, what: 'Last value' }]),
        ]),
      ]),
    ])
    expect(toComplete(doc).map((item) => item.what)).toEqual(['Item 1, V2', 'Both, V1', 'Last value'])
  })

  it('finds blanks in headers and footers, by part', () => {
    const doc = document([section([paragraph('body')], { footer: [paragraph(['Ref. ', blank('reference number')])] })])
    expect(toComplete(doc)).toEqual([{ kind: 'blank', what: 'reference number', section: 0, part: 'footer' }])
  })

  it('takes a custom block’s own pending items, and what it expands to when given the registry', () => {
    const block = custom('rating', { value: null as number | null }, { pending: [{ what: 'Rating rationale' }] })
    const doc = document([section([block])])
    expect(toComplete(doc).map((item) => item.what)).toEqual(['Rating rationale'])
    const registry = defineCustomBlocks({
      rating: (data: { value: number | null }) => [fields([field('Rating', data.value === null ? null : String(data.value))])],
    })
    expect(toComplete(doc, { registry }).map((item) => [item.kind, item.what])).toEqual([
      ['pending', 'Rating rationale'],
      ['field', 'Rating'],
    ])
  })
})
