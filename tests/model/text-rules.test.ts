import { describe, expect, it } from 'vitest'

import {
  applyTextRules,
  compileTextRules,
  ruleText,
  TextRuleError,
  type Block,
  type Document,
  type Inline,
  type TextRule,
} from '../../src/model'

const runs = (text: string, rules: TextRule[], place?: Parameters<typeof ruleText>[2]) => ruleText(text, compileTextRules(rules), place)
/** Runs as compact strings: text, then the touches in brackets. */
const show = (input: readonly Inline[] | string) =>
  (typeof input === 'string' ? [input] : input).map((inline) => {
    if (typeof inline === 'string') return inline
    if (inline.type !== 'text') return `<${inline.type}>`
    const marks = [
      inline.link ? `link=${inline.link}` : '',
      inline.italic ? 'i' : '',
      inline.bold ? 'b' : '',
      inline.superscript ? 'sup' : '',
      inline.subscript ? 'sub' : '',
      inline.color ? `color=${inline.color}` : '',
    ].filter(Boolean)
    return marks.length ? `${inline.text}[${marks.join(',')}]` : inline.text
  })

const refuses = (rule: TextRule, message: RegExp) => {
  try {
    compileTextRules([rule])
  } catch (error) {
    expect(error).toBeInstanceOf(TextRuleError)
    expect((error as TextRuleError).problems.map((p) => p.message).join(' / ')).toMatch(message)
    return
  }
  throw new Error('expected the rule to be refused')
}

describe('the links rule', () => {
  const links: TextRule = { builtIn: 'links' }

  it('links e-mail and web addresses, leaving trailing punctuation outside', () => {
    expect(show(runs('Write to someone@example.org, or see https://example.org/a_(b)).', [links]))).toEqual([
      'Write to ',
      'someone@example.org[link=mailto:someone@example.org]',
      ', or see ',
      'https://example.org/a_(b)[link=https://example.org/a_(b)]',
      ').',
    ])
    expect(show(runs('See www.example.org and example.ca/page; not file.pdf or e.g. this.', [links]))).toEqual([
      'See ',
      'www.example.org[link=https://www.example.org]',
      ' and ',
      'example.ca/page[link=https://example.ca/page]',
      '; not file.pdf or e.g. this.',
    ])
  })

  it('links only the schemes allowed, and bare domains only on the listed top-level domains', () => {
    expect(show(runs('ftp://example.org/x and someone@example.org', [{ builtIn: 'links', schemes: ['https'] }]))).toEqual([
      'ftp://example.org/x and someone@example.org',
    ])
    expect(show(runs('example.ca and example.zz', [{ builtIn: 'links', domains: ['zz'] }]))).toEqual([
      'example.ca and ',
      'example.zz[link=https://example.zz]',
    ])
    expect(show(runs('javascript://alert and https://ok.org', [links]))).toEqual(['javascript://alert and ', 'https://ok.org[link=https://ok.org]'])
  })
})

describe('terms, markers and formulas', () => {
  it('sets listed terms in italics, whole words only', () => {
    const rule: TextRule = { terms: ['Made-up bacterium', 'M. exemplum'], style: { italic: true } }
    expect(show(runs('M. exemplum and Made-up bacteria; not XM. exemplum2.', [rule]))).toEqual([
      'M. exemplum[i]',
      ' and Made-up bacteria; not XM. exemplum2.',
    ])
    expect(show(runs('m. EXEMPLUM', [{ ...rule, ignoreCase: true }]))).toEqual(['m. EXEMPLUM[i]'])
  })

  it('raises a trailing footnote marker after a space, and one or several letters', () => {
    const rule: TextRule = { builtIn: 'trailingMarker', markers: 'a-h' }
    expect(show(runs('Total a', [rule]))).toEqual(['Total ', 'a[sup]'])
    expect(show(runs('Lead a, b', [rule]))).toEqual(['Lead ', 'a, b[sup]'])
    expect(show(runs('Lead a,b', [rule]))).toEqual(['Lead ', 'a,b[sup]'])
    // Not a marker: no space before it, a letter outside the list, a lone letter, a word.
    for (const text of ['Totala', 'Total z', 'a', 'Total ab']) expect(show(runs(text, [rule]))).toEqual([text])
  })

  it('raises a leading marker before a space or a capital', () => {
    const rule: TextRule = { builtIn: 'leadingMarker', markers: 'a-h' }
    expect(show(runs('a Dried at 105 °C', [rule]))).toEqual(['a[sup]', ' Dried at 105 °C'])
    expect(show(runs('bAir dried', [rule]))).toEqual(['b[sup]', 'Air dried'])
    expect(show(runs('apple', [rule]))).toEqual(['apple'])
  })

  it('sets the digits of the listed formulas low, standing alone only', () => {
    const rule: TextRule = { builtIn: 'formulaDigits', formulas: ['CaCO3', 'NO3', 'H2SO4'] }
    expect(show(runs('mg CaCO3/L, NO3 and H2SO4; not XNO3 or NO35.', [rule]))).toEqual([
      'mg CaCO',
      '3[sub]',
      '/L, NO',
      '3[sub]',
      ' and H',
      '2[sub]',
      'SO',
      '4[sub]',
      '; not XNO3 or NO35.',
    ])
    refuses({ builtIn: 'formulaDigits' }, /lists no formulas/)
  })

  it('styles one group of a pattern, by number or name', () => {
    expect(show(runs('Run 12 of 30', [{ pattern: 'Run (\\d+)', group: 1, style: { bold: true } }]))).toEqual(['Run ', '12[b]', ' of 30'])
    expect(show(runs('Run 12 of 30', [{ pattern: 'of (?<total>\\d+)', group: 'total', style: { color: 'aa0000' } }]))).toEqual([
      'Run 12 of ',
      '30[color=aa0000]',
    ])
  })
})

describe('how rules compose', () => {
  it('applies rules in the order listed, never searching text inside a link again', () => {
    const rules: TextRule[] = [
      { builtIn: 'links' },
      { builtIn: 'formulaDigits', formulas: ['NO3'] },
      { terms: ['example'], style: { italic: true } },
    ]
    expect(show(runs('See https://example.org/NO3 for NO3 example.', rules))).toEqual([
      'See ',
      'https://example.org/NO3[link=https://example.org/NO3]',
      ' for NO',
      '3[sub]',
      ' ',
      'example[i]',
      '.',
    ])
  })

  it('matches across runs of different looks, and lays its style over each', () => {
    const list = ruleText([{ type: 'text', text: 'Made-', bold: true }, 'up term'], compileTextRules([{ terms: ['Made-up term'], style: { italic: true } }]))
    expect(show(list)).toEqual(['Made-[i,b]', 'up term[i]'])
  })

  it('stops a stretch at a field, a break or a blank', () => {
    const list = ruleText(['Total ', { type: 'break' }, 'a'], compileTextRules([{ builtIn: 'trailingMarker', markers: 'a' }]))
    expect(show(list)).toEqual(['Total ', '<break>', 'a'])
  })
})

describe('where rules apply', () => {
  const doc = (blocks: Block[], rules: TextRule[]): Document => ({ textRules: rules, sections: [{ body: blocks, footer: [{ type: 'paragraph', content: 'Total a' }] }] })
  const marker: TextRule = { builtIn: 'trailingMarker', markers: 'a', in: ['cell', 'label'] }

  it('limits a rule to its places', () => {
    const out = applyTextRules(
      doc(
        [
          { type: 'paragraph', content: 'Total a' },
          { type: 'table', rows: [{ cells: ['Total a', { content: 'Total a', shading: 'eeeeee' }] }] },
          { type: 'fields', rows: [{ label: 'Total a', value: 'Total a' }] },
        ],
        [marker],
      ),
    ).document
    const [paragraph, table, fields] = out.sections[0]!.body as [Block, Block, Block]
    expect(paragraph).toEqual({ type: 'paragraph', content: 'Total a' })
    expect(table.type === 'table' && show(table.rows[0]!.cells[0] as Inline[])).toEqual(['Total ', 'a[sup]'])
    // A cell given as an object stays one.
    expect(table.type === 'table' && (table.rows[0]!.cells[1] as { shading: string }).shading).toBe('eeeeee')
    expect(fields.type === 'fields' && show(fields.rows[0]!.label as Inline[])).toEqual(['Total ', 'a[sup]'])
    expect(fields.type === 'fields' && fields.rows[0]!.value).toBe('Total a')
    expect(out.sections[0]!.footer).toEqual([{ type: 'paragraph', content: 'Total a' }])
    expect(out.textRules).toBeUndefined()
  })

  it('applies a rule that is not automatic only where a block names it, and none where a block says false', () => {
    const lead: TextRule = { id: 'notes', builtIn: 'leadingMarker', markers: 'a', auto: false }
    const italic: TextRule = { terms: ['term'], style: { italic: true } }
    const out = applyTextRules(
      doc(
        [
          { type: 'paragraph', content: 'a term' },
          { type: 'paragraph', content: 'a term', textRules: ['notes'] },
          { type: 'paragraph', content: 'a term', textRules: false },
          { type: 'pinned', textRules: ['notes'], blocks: [{ type: 'paragraph', content: 'a term' }] },
        ],
        [lead, italic],
      ),
    ).document.sections[0]!.body
    const contents = out.map((block) => (block.type === 'paragraph' ? show(block.content as Inline[]) : show((block as unknown as { blocks: Array<{ content: Inline[] }> }).blocks[0]!.content)))
    expect(contents).toEqual([['a ', 'term[i]'], ['a[sup]', ' ', 'term[i]'], ['a term'], ['a[sup]', ' ', 'term[i]']])
  })

  it('leaves a document without rules as it is', () => {
    const document: Document = { sections: [{ body: [{ type: 'paragraph', content: 'x' }] }] }
    expect(applyTextRules(document).document).toBe(document)
  })
})

describe('guards on rules a person may edit', () => {
  it('refuses the constructs that cannot be bounded, naming each', () => {
    refuses({ pattern: '(a)\\1', style: { bold: true } }, /back-reference/)
    refuses({ pattern: 'a(?=b)', style: { bold: true } }, /look-around/)
    refuses({ pattern: '(?<=a)b', style: { bold: true } }, /look-around/)
    refuses({ pattern: '(a+)+$', style: { bold: true } }, /repeats a group that holds a repeat/)
    refuses({ pattern: '(a|aa)*$', style: { bold: true } }, /repeats a group that holds a repeat or an alternative/)
    refuses({ pattern: 'a{1,500}', style: { bold: true } }, /repeats up to 500 times/)
    refuses({ pattern: '\\w+\\s+\\w+\\d+', style: { bold: true } }, /repeats too much/)
    refuses({ pattern: 'x'.repeat(201), style: { bold: true } }, /longer than 200/)
    refuses({ pattern: '\\p{L}', style: { bold: true } }, /Unicode property/)
    refuses({ pattern: 'a{', style: { bold: true } }, /not a repeat/)
    refuses({ pattern: 'ab', flags: 'g', style: { bold: true } }, /only flag/)
    refuses({ pattern: 'ab' }, /neither style nor link/)
    refuses({ pattern: 'ab', style: { blink: true } as never }, /not a text style/)
    refuses({ terms: Array.from({ length: 1001 }, (_, i) => `t${i}`), style: { italic: true } }, /more than 1000 terms/)
    refuses({ builtIn: 'nope' as never }, /not one of/)
    refuses({ pattern: 'a', terms: ['a'], style: { bold: true } }, /one of pattern, terms or builtIn/)
    refuses({ pattern: 'a', style: { bold: true }, colour: 'x' } as never, /no key "colour"/)
  })

  it('bounds a repeat to 100 characters, so a long run of letters takes no longer than a short one', () => {
    const list = runs(`${'a'.repeat(150)}b`, [{ pattern: 'a+b', style: { bold: true } }])
    expect(show(list)).toEqual(['a'.repeat(50), `${'a'.repeat(100)}b[b]`])
    // The classic runaway pattern is refused rather than run; a worst case of what is allowed stays quick.
    const started = Date.now()
    runs('a'.repeat(9_999), [{ pattern: '[a]*[a]*c', style: { bold: true } }])
    expect(Date.now() - started).toBeLessThan(2_000)
  })

  it('leaves text longer than the limit unsearched by a pattern, and says so', () => {
    const long = `${'word '.repeat(3_000)}term`
    const out = applyTextRules({ textRules: [{ terms: ['term'], style: { italic: true } }, { builtIn: 'links' }], sections: [{ body: [{ type: 'paragraph', content: `${long} https://example.org` }] }] })
    expect(out.skipped).toHaveLength(1)
    expect(out.skipped[0]).toMatchObject({ rule: 0, start: 'word word word word word word word word ' })
    const content = (out.document.sections[0]!.body[0] as unknown as { content: Inline[] }).content
    // The built-in links rule still runs.
    expect(content.some((run) => typeof run !== 'string' && run.type === 'text' && run.link === 'https://example.org')).toBe(true)
  })

  it('stops with an error when the rules find more matches than the document allows', () => {
    const document: Document = { textRules: [{ pattern: 'x', style: { bold: true } }], sections: [{ body: [{ type: 'paragraph', content: 'x '.repeat(50) }] }] }
    expect(() => applyTextRules(document, { matches: 20 })).toThrow(/more than 20 matches/)
    expect(applyTextRules(document).document).not.toBe(document)
  })
})
