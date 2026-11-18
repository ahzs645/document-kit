import { describe, expect, it } from 'vitest'
import {
  caption,
  captionLabel,
  captionPrefix,
  cell,
  document,
  DEFAULT_THEME,
  figure,
  figureGrid,
  numberCaptions,
  paragraph,
  row,
  section,
  table,
} from '../../src/index'

describe('numberCaptions', () => {
  it('numbers tables and figures in reading order, each kind on its own count, across sections', () => {
    const t1 = table([row(['a'])], { caption: caption('First') })
    const f1 = figure(null, { caption: caption('Map') })
    const t2 = table([row(['b'])], { caption: caption('Second') })
    const g1 = figureGrid([{ image: null }], 2, { caption: caption('Pair') })
    const t3 = table([row(['c'])], { caption: caption('Third') })
    const doc = document([section([t1, paragraph('x'), f1, t2]), section([g1, t3])])
    const numbers = numberCaptions(doc)
    expect(numbers.get(t1)).toBe(1)
    expect(numbers.get(t2)).toBe(2)
    expect(numbers.get(t3)).toBe(3)
    expect(numbers.get(f1)).toBe(1)
    expect(numbers.get(g1)).toBe(2)
  })

  it('skips unnumbered captions and uncaptioned blocks, but a silent caption takes a number', () => {
    const plain = table([row(['a'])])
    const unnumbered = table([row(['a'])], { caption: caption('As written', false) })
    const silent = table([row(['a'])], { caption: caption('Table 1a states its own number', 'silent') })
    const next = table([row(['a'])], { caption: caption('Next') })
    const numbers = numberCaptions([plain, unnumbered, silent, next])
    expect(numbers.has(plain)).toBe(false)
    expect(numbers.has(unnumbered)).toBe(false)
    expect(numbers.get(silent)).toBe(1)
    expect(numbers.get(next)).toBe(2)
  })

  it('counts a figure inside a table cell where it is read', () => {
    const inner = figure(null, { caption: caption('Inside') })
    const outer = table([row([cell(null, { blocks: [inner] })])], { caption: caption('Holder') })
    const after = figure(null, { caption: caption('After') })
    const numbers = numberCaptions([outer, after])
    expect(numbers.get(outer)).toBe(1)
    expect(numbers.get(inner)).toBe(1)
    expect(numbers.get(after)).toBe(2)
  })

  it('writes labels and prefixes in the theme words', () => {
    expect(captionLabel('table', 3, DEFAULT_THEME)).toBe('Table 3')
    const theme = { captions: { table: 'Tableau', figure: 'Fig.', separator: '. ' } }
    expect(captionPrefix('figure', caption('A view'), 2, theme)).toBe('Fig. 2. ')
    expect(captionPrefix('table', caption('x', false), 2, theme)).toBe('')
    expect(captionPrefix('table', caption('x', 'silent'), 2, theme)).toBe('')
    expect(captionPrefix('table', caption(''), 4, theme, false)).toBe('Tableau 4')
  })
})
