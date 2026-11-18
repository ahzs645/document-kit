/**
 * Filters for the wording of sentences, shared by Word templates and JSON recipes: the first word of a fill
 * capitalised, a word chosen by a count, a list joined in prose, and runs of consecutive values collapsed to ranges.
 * They stand beside `format`, `upper` and `lower`; a caller's own filter of the same name (`formatters` in a template,
 * `filters` in a recipe) is used instead. document-kit's own, not officeParser's (docs/template-upstream.md). Pure.
 */

/** The names of the sentence filters. */
export const SENTENCE_FILTER_NAMES: readonly string[] = ['capitalize', 'plural', 'prose', 'ranges']

type Filter = (value: unknown, ...args: Array<string | number>) => unknown

/** A filter given something it cannot read; the message names the filter and what it takes. */
export class SentenceFilterError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'SentenceFilterError'
  }
}

const textOf = (value: unknown, filter: string): string => {
  if (typeof value === 'string') return value
  if (typeof value === 'number' || typeof value === 'boolean' || typeof value === 'bigint') return String(value)
  throw new SentenceFilterError(`${filter} takes text, and was given ${Array.isArray(value) ? 'a list' : typeof value}`)
}

/**
 * `{{ value | capitalize }}`: a plain lower-case first word gets a capital ("three samples" → "Three samples"), for a
 * fill that opens a sentence. A word cased on purpose ("pH", "mL", "iPhone") and a lone letter (a footnote marker) are
 * left as they are (a value that is one letter alone), as is anything not starting with a letter.
 */
export function capitalize(value: unknown): unknown {
  if (value === null || value === undefined) return value
  const text = textOf(value, 'capitalize')
  const lead = text.length - text.trimStart().length
  const rest = text.slice(lead)
  if (/^[a-z]$/.test(rest.trim())) return text
  if (!/^[a-z]+(?![A-Za-z])/.test(rest)) return text
  return text.slice(0, lead) + rest.charAt(0).toUpperCase() + rest.slice(1)
}

/** A count from a number, a number written as text, or a list (its length). */
function countOf(value: unknown): number {
  if (Array.isArray(value)) return value.length
  if (typeof value === 'number' && Number.isFinite(value)) return value
  if (typeof value === 'string' && /^\s*-?\d+(\.\d+)?\s*$/.test(value)) return Number(value)
  throw new SentenceFilterError(`plural takes a count (a number or a list), and was given ${JSON.stringify(value)}`)
}

/**
 * `{{ n | plural "sample" "samples" }}`: the first word for a count of exactly one, the second otherwise (the first
 * with an `s` when the second is not given). A list counts as its length.
 */
export function plural(value: unknown, singular?: string | number, pluralForm?: string | number): unknown {
  if (singular === undefined) throw new SentenceFilterError('plural takes the singular word, and the plural when it is not the singular with an s')
  const one = String(singular)
  return countOf(value) === 1 ? one : pluralForm === undefined ? `${one}s` : String(pluralForm)
}

/**
 * `{{ names | prose }}`: a list joined as a sentence says it — "a", "a and b", "a, b and c". The first argument is the
 * last joining word (`prose "or"`), the second what goes between the others (", " by default). Empty items are left
 * out; a value that is not a list is printed as it is.
 */
export function prose(value: unknown, conjunction: string | number = 'and', separator: string | number = ', '): unknown {
  if (!Array.isArray(value)) return value
  const items = value.filter((item) => item !== null && item !== undefined && item !== '').map((item) => textOf(item, 'prose'))
  if (items.length <= 1) return items[0] ?? ''
  return `${items.slice(0, -1).join(String(separator))} ${String(conjunction)} ${items[items.length - 1]}`
}

/** A value read as a prefix and a whole number at its end: `12` → ('', 12), `S-7` → ('S-', 7). */
function numbered(value: unknown): { prefix: string; n: number; text: string } | null {
  if (typeof value === 'number') return Number.isInteger(value) ? { prefix: '', n: value, text: String(value) } : null
  if (typeof value !== 'string') return null
  const match = /^(.*?)(\d{1,15})$/.exec(value)
  if (!match) return null
  // A zero-padded number keeps its width: 007, 008 run on; 009, 010 too, while 09, 100 do not.
  return { prefix: match[1]!, n: Number(match[2]), text: value }
}

/**
 * `{{ ids | ranges }}`: a list with each run of consecutive values (numbers, or texts that differ only in a number at
 * their end, "S-1", "S-2", "S-3") of at least `min` items (3 by default) collapsed to its first and last joined by
 * `dash` ("–" by default): [1, 2, 3, 5] → ["1–3", "5"]. A list, in the order given, so it can go on to `prose`:
 * `{{ ids | ranges | prose }}` → "S-1–S-3 and S-5". A value that is not a list is printed as it is.
 */
export function ranges(value: unknown, dash: string | number = '–', min: string | number = 3): unknown {
  if (!Array.isArray(value)) return value
  const least = Math.max(2, Math.floor(Number(min)) || 3)
  const out: string[] = []
  let run: Array<{ prefix: string; n: number; text: string }> = []
  const flush = () => {
    if (run.length >= least) out.push(`${run[0]!.text}${String(dash)}${run[run.length - 1]!.text}`)
    else out.push(...run.map((item) => item.text))
    run = []
  }
  for (const item of value) {
    if (item === null || item === undefined || item === '') continue
    const read = numbered(item)
    const last = run[run.length - 1]
    if (!read) {
      flush()
      out.push(textOf(item, 'ranges'))
      continue
    }
    if (last && last.prefix === read.prefix && read.n === last.n + 1 && read.text.length - read.prefix.length >= String(read.n).length) run.push(read)
    else {
      flush()
      run = [read]
    }
  }
  flush()
  return out
}

const FILTERS: Readonly<Record<string, Filter>> = {
  capitalize: (value) => capitalize(value),
  plural: (value, singular, pluralForm) => plural(value, singular, pluralForm),
  prose: (value, conjunction, separator) => prose(value, conjunction, separator),
  ranges: (value, dash, min) => ranges(value, dash, min),
}

/** The sentence filter of a name, or undefined. */
export function sentenceFilter(name: string): Filter | undefined {
  return Object.prototype.hasOwnProperty.call(FILTERS, name) ? FILTERS[name] : undefined
}
