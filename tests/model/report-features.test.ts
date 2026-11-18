import { describe, expect, it } from 'vitest'

import {
  checkbox,
  choices,
  defineCustomBlocks,
  document,
  expandDocument,
  heading,
  numberCaptions,
  paragraph,
  pinned,
  placeholder,
  plainText,
  resolveSentences,
  section,
  sentence,
  signature,
  table,
  row,
  toComplete,
  type Block,
} from '../../src/model'

describe('the model\'s report features', () => {
  it('reads a tick box and a placeholder as text', () => {
    expect(plainText(['Yes ', checkbox(true), ' No ', checkbox(false)])).toBe('Yes ☒ No ☐')
    expect(plainText(['Run on ', placeholder('run date')])).toBe('Run on [run date]')
  })

  it('prints a sentence with a missing slot as placeholders when it asks to, and still lists it', () => {
    const run = sentence('Received {count} samples on {date}.', { count: 3, date: null }, { label: 'Receipt', missing: 'placeholder', bold: true })
    const printed = resolveSentences([run])
    expect(plainText(printed)).toBe('Received 3 samples on [date].')
    expect(printed[1]).toEqual({ type: 'blank', what: 'date', show: 'placeholder', bold: true })
    const items = toComplete(document([section([paragraph([run])])]))
    expect(items.map((i) => [i.kind, i.what, i.missing])).toEqual([['sentence', 'Receipt', ['date']]])
  })

  it('lists an unanswered tick-box question and a blank answer, and walks into pinned blocks', () => {
    const doc = document([
      section([
        heading(1, 'Review'),
        choices(['Yes', 'No'], null, { label: 'Met?', answer: null }),
        choices(['Yes', 'No'], 'Yes', { label: 'Seen?' }),
        pinned([heading(2, 'Contact'), paragraph(['Phone ', placeholder('phone')]), table([row([null])], { caption: { text: 'Inside' } })]),
        signature({ name: 'A. Person', title: null }),
      ]),
    ])
    expect(toComplete(doc).map((i) => [i.kind, i.heading, i.what])).toEqual([
      ['choice', 'Review', 'Met'],
      ['answer', 'Review', 'Met'],
      ['blank', 'Contact', 'phone'],
      ['cell', 'Contact', 'Inside, row 1, column 1'],
      ['signature', 'Contact', 'Title of the signatory'],
    ])
    const tables = numberCaptions(doc)
    expect([...tables.values()]).toEqual([1])
  })

  it('expands custom blocks inside a pinned block', () => {
    const registry = defineCustomBlocks({ note: (data: unknown) => [paragraph(String(data))] })
    const out = expandDocument(document([section([pinned([{ type: 'custom', kind: 'note', data: 'Expanded.' } as Block])])]), registry)
    expect(out.sections[0]!.body[0]).toEqual({ type: 'pinned', blocks: [{ type: 'paragraph', content: 'Expanded.' }] })
  })
})
