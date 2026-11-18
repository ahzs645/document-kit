/**
 * Text rules: automatic text touches stated as data (a document's `textRules`), applied to the text's ordinary runs so
 * every writer, the layout and the PDF follow them. Web and e-mail addresses become links, listed terms go into
 * italics, footnote markers are raised, the digits of listed formulas are set low — none of it hard-coded: an
 * application states its patterns and term lists as JSON, and the kit ships only generic built-ins.
 *
 * Patterns come from JSON a person may edit, so they are read through a safe subset of regular expressions and every
 * search is bounded (see {@link TEXT_RULE_LIMITS} and `docs/model.md`, "Text rules"). Pure: no DOM, no network.
 *
 * How rules compose: a block's rules are applied one after another in the order the document lists them, each to the
 * runs the rules before it made. A rule sees a stretch of consecutive text runs as one text, so a match may run across
 * two runs of different looks; a link, a field, a break, a tab, a blank, a picture or a tick box ends a stretch. Text
 * inside a link (a run with `link`, given or made by an earlier rule) is never searched again, so nothing is linked
 * twice and a web address's digits are never set low. A style is laid over the matched text's own.
 */
import { resolveSentences } from './inline'
import { toCell } from './table'
import type {
  Block,
  CellInput,
  Document,
  FieldRow,
  Inline,
  ListItem,
  RichText,
  Section,
  TableBlock,
  TextPlace,
  TextRule,
  TextRun,
  TextStyle,
} from './types'

/** Bounds on text rules; each may be lowered or raised (`compileTextRules(rules, limits)`, `applyTextRules`). */
export const TEXT_RULE_LIMITS = {
  /** Rules in a document. */
  rules: 100,
  /** Characters in one pattern. */
  patternLength: 200,
  /** Most times `*`, `+` or `{n,}` repeat; an explicit `{n,m}` may not ask for more. */
  repeat: 100,
  /** The product, over a pattern's repeats, of (most times + 1): a bound on how hard one search can try. */
  cost: 12_000,
  /** Terms in one rule's list, and characters in one term. */
  terms: 1_000,
  termLength: 100,
  /** Characters of one stretch of text a pattern or term list searches; a longer stretch is left as it is, and reported. */
  textLength: 10_000,
  /** Matches in one document, every rule together; past it, applying the rules stops with an error. */
  matches: 20_000,
}
export type TextRuleLimits = typeof TEXT_RULE_LIMITS

export interface TextRuleProblem {
  /** The rule's index in the list. */
  rule: number
  /** The key at fault (`pattern`, `terms[3]`, …), or '' for the rule as a whole. */
  key: string
  message: string
}

/** Rules that cannot be used, each problem named; or a document whose rules found more matches than allowed. */
export class TextRuleError extends Error {
  readonly problems: readonly TextRuleProblem[]
  constructor(message: string, problems: readonly TextRuleProblem[] = []) {
    super(`document-kit: ${message}`)
    this.name = 'TextRuleError'
    this.problems = problems
  }
}

/** A range of a stretch's text and what is done to it. */
interface Hit {
  start: number
  end: number
  style?: TextStyle
  href?: string
}

interface CompiledRule {
  index: number
  id?: string
  auto: boolean
  places: ReadonlySet<TextPlace> | null
  /** A pattern or term list: searched only in stretches up to `textLength`. */
  bounded: boolean
  find(text: string): Hit[]
}

/** Rules read once and checked: what {@link ruleText} and {@link applyTextRules} apply. */
export interface CompiledTextRules {
  readonly rules: readonly CompiledRule[]
  readonly limits: TextRuleLimits
}

/** A stretch of text a pattern or term rule did not search because it was too long. */
export interface SkippedText {
  rule: number
  id?: string
  length: number
  /** Its first characters. */
  start: string
}

const RULE_KEYS = new Set([
  'id',
  'pattern',
  'flags',
  'terms',
  'ignoreCase',
  'builtIn',
  'group',
  'style',
  'link',
  'schemes',
  'domains',
  'markers',
  'formulas',
  'in',
  'auto',
])
const BUILT_INS = new Set(['links', 'trailingMarker', 'leadingMarker', 'formulaDigits'])
const PLACES = new Set<TextPlace>([
  'paragraph',
  'heading',
  'list',
  'label',
  'value',
  'cell',
  'caption',
  'note',
  'signature',
  'choice',
  'header',
  'footer',
])
const STYLE_KEYS = new Set([
  'bold',
  'italic',
  'underline',
  'strike',
  'superscript',
  'subscript',
  'smallCaps',
  'allCaps',
  'color',
  'size',
  'font',
  'highlight',
])
const DEFAULT_SCHEMES = ['https', 'http', 'mailto']
/** Top-level domains a bare address (`example.org/page`) is linked on, unless a rule gives its own. */
const DEFAULT_DOMAINS = ['com', 'org', 'net', 'edu', 'gov', 'int', 'info', 'io', 'ca', 'us', 'uk', 'au', 'nz', 'ie', 'eu', 'de', 'fr']

/* ------------------------------------------------------------------ the safe pattern subset */

interface GroupState {
  repeats: boolean
  alternates: boolean
}

/**
 * A pattern checked against the safe subset and rewritten for the search: `*`, `+` and `{n,}` become bounded repeats.
 * Refused: back-references (`\1`, `\k<name>`), look-arounds (`(?=`, `(?!`, `(?<=`, `(?<!`), Unicode property escapes,
 * a repeat of a group that holds a repeat or an alternative (`(a+)+`, `(a|b)*`), a `{` that is not a repeat, an
 * explicit repeat past `repeat`, and patterns whose repeats together could try more than `cost` ways.
 */
function safePattern(pattern: string, limits: TextRuleLimits): { source: string } | { problem: string } {
  if (pattern.length === 0) return { problem: 'is empty' }
  if (pattern.length > limits.patternLength) return { problem: `is longer than ${limits.patternLength} characters` }
  let out = ''
  const root: GroupState = { repeats: false, alternates: false }
  const stack: GroupState[] = []
  const top = () => stack[stack.length - 1] ?? root
  let last: GroupState | 'atom' | null = null
  let cost = 1
  let i = 0
  while (i < pattern.length) {
    const c = pattern[i]!
    if (c === '\\') {
      const next = pattern[i + 1]
      if (next === undefined) return { problem: 'ends in a lone \\' }
      if (/[1-9]/.test(next) || next === 'k') return { problem: `uses a back-reference (\\${next}), which cannot be bounded` }
      if (next === 'p' || next === 'P') return { problem: `uses a Unicode property (\\${next}{…}), which this subset does not read` }
      if (next === 'c') return { problem: 'uses a control escape (\\c), which this subset does not read' }
      if (next === 'u') {
        if (!/^[0-9a-fA-F]{4}$/.test(pattern.slice(i + 2, i + 6))) return { problem: 'has a \\u that is not four hex digits' }
        out += pattern.slice(i, i + 6)
        i += 6
        last = 'atom'
        continue
      }
      if (next === 'x') {
        if (!/^[0-9a-fA-F]{2}$/.test(pattern.slice(i + 2, i + 4))) return { problem: 'has a \\x that is not two hex digits' }
        out += pattern.slice(i, i + 4)
        i += 4
        last = 'atom'
        continue
      }
      out += c + next
      i += 2
      last = next === 'b' || next === 'B' ? null : 'atom'
      continue
    }
    if (c === '[') {
      let j = i + 1
      if (pattern[j] === '^') j += 1
      if (pattern[j] === ']') j += 1
      while (j < pattern.length && pattern[j] !== ']') {
        if (pattern[j] === '\\') {
          const next = pattern[j + 1]
          if (next === 'p' || next === 'P') return { problem: `uses a Unicode property (\\${next}{…}), which this subset does not read` }
          j += 2
        } else j += 1
      }
      if (j >= pattern.length) return { problem: 'has a [ with no ]' }
      out += pattern.slice(i, j + 1)
      i = j + 1
      last = 'atom'
      continue
    }
    if (c === '(') {
      if (pattern.startsWith('(?=', i) || pattern.startsWith('(?!', i) || pattern.startsWith('(?<=', i) || pattern.startsWith('(?<!', i))
        return { problem: 'uses a look-around ((?= (?! (?<= (?<!), which this subset does not read' }
      let open = '('
      if (pattern.startsWith('(?:', i)) open = '(?:'
      else if (pattern.startsWith('(?<', i)) {
        const name = /^\(\?<([A-Za-z_][A-Za-z0-9_]*)>/.exec(pattern.slice(i))
        if (!name) return { problem: 'has a (?< that is not a named group (?<name>…)' }
        open = name[0]
      } else if (pattern.startsWith('(?', i)) return { problem: 'has a (? this subset does not read' }
      out += open
      i += open.length
      stack.push({ repeats: false, alternates: false })
      last = null
      continue
    }
    if (c === ')') {
      const group = stack.pop()
      if (!group) return { problem: 'has a ) with no (' }
      const parent = top()
      parent.repeats ||= group.repeats
      parent.alternates ||= group.alternates
      out += ')'
      i += 1
      last = group
      continue
    }
    if (c === '|') {
      top().alternates = true
      out += c
      i += 1
      last = null
      continue
    }
    if (c === '*' || c === '+' || c === '?' || c === '{') {
      let min: number
      let max: number
      let length = 1
      if (c === '{') {
        const repeat = /^\{(\d{1,4})(,(\d{0,4}))?\}/.exec(pattern.slice(i))
        if (!repeat) return { problem: 'has a { that is not a repeat ({2}, {1,3}); write \\{ for the character' }
        min = Number(repeat[1])
        max = repeat[2] === undefined ? min : repeat[3] === '' ? Infinity : Number(repeat[3])
        length = repeat[0].length
        if (max !== Infinity && max > limits.repeat) return { problem: `repeats up to ${max} times; ${limits.repeat} is the most` }
        if (min > max || min > limits.repeat) return { problem: `has a repeat ${repeat[0]} that cannot be met` }
      } else {
        min = c === '+' ? 1 : 0
        max = c === '?' ? 1 : Infinity
      }
      if (last === null) return { problem: `has a ${c} with nothing to repeat` }
      if (last !== 'atom' && (last.repeats || last.alternates))
        return { problem: 'repeats a group that holds a repeat or an alternative ((a+)+, (a|b)*), which cannot be bounded' }
      if (max === Infinity) max = Math.max(min, limits.repeat)
      top().repeats = true
      cost *= max + 1
      if (cost > limits.cost) return { problem: `repeats too much: its repeats together could try more than ${limits.cost} ways` }
      out += c === '?' ? '?' : `{${min},${max}}`
      i += length
      if (pattern[i] === '?') {
        out += '?'
        i += 1
      }
      last = null
      continue
    }
    if (c === '^' || c === '$') {
      out += c
      i += 1
      last = null
      continue
    }
    out += c
    i += 1
    last = 'atom'
  }
  if (stack.length) return { problem: 'has a ( with no )' }
  return { source: out }
}

/* ------------------------------------------------------------------ links */

const EMAIL = '[A-Za-z0-9._%+-]{1,64}@[A-Za-z0-9-]{1,63}(?:\\.[A-Za-z0-9-]{1,63}){1,8}'
const SCHEMED = '\\b[A-Za-z][A-Za-z0-9+.-]{0,30}:\\/\\/[^\\s<>"]{1,2000}|\\bmailto:[^\\s<>"]{1,500}'
const WWW = '\\bwww\\.[^\\s<>"]{1,2000}'

/** The address a text is a link to, under the allowed schemes; undefined when it is not one. */
export function linkHref(text: string, schemes: readonly string[]): string | undefined {
  const allowed = new Set(schemes.map((s) => s.toLowerCase()))
  const scheme = /^([A-Za-z][A-Za-z0-9+.-]*):/.exec(text)?.[1]?.toLowerCase()
  if (scheme && (text.slice(scheme.length + 1).startsWith('//') || scheme === 'mailto')) return allowed.has(scheme) ? text : undefined
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(text)) return allowed.has('mailto') ? `mailto:${text}` : undefined
  if (/^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+(\/\S*)?$/.test(text)) {
    if (allowed.has('https')) return `https://${text}`
    if (allowed.has('http')) return `http://${text}`
  }
  return undefined
}

/** Trailing sentence punctuation, and a closing bracket the address did not open, stay outside the link. */
function trimLink(text: string): string {
  let out = text
  for (;;) {
    const last = out[out.length - 1]
    if (last === undefined) return out
    if ('.,;:!?\'"’”'.includes(last)) out = out.slice(0, -1)
    else if (last === ')' && count(out, '(') < count(out, ')')) out = out.slice(0, -1)
    else if (last === ']' && count(out, '[') < count(out, ']')) out = out.slice(0, -1)
    else return out
  }
}
const count = (text: string, ch: string) => text.split(ch).length - 1

function linkFinder(rule: TextRule): (text: string) => Hit[] {
  const schemes = rule.schemes ?? DEFAULT_SCHEMES
  const domains = (rule.domains ?? DEFAULT_DOMAINS).map((d) => d.replace(/[^A-Za-z0-9-]/g, '')).filter(Boolean)
  const bare = domains.length ? `|\\b(?:[A-Za-z0-9-]{1,63}\\.){1,8}(?:${domains.join('|')})\\b(?:\\/[^\\s<>"]{0,2000})?` : ''
  const source = `${EMAIL}|${SCHEMED}|${WWW}${bare}`
  return (text) => {
    const out: Hit[] = []
    const regex = new RegExp(source, 'gi')
    for (const match of text.matchAll(regex)) {
      const start = match.index ?? 0
      // An address glued to a word or another address is not one ("x.org" inside "file@x.org" was taken whole).
      const before = text[start - 1]
      if (before !== undefined && /[A-Za-z0-9@./_-]/.test(before)) continue
      const link = trimLink(match[0])
      if (!link) continue
      const href = linkHref(link, schemes)
      if (!href) continue
      out.push({ start, end: start + link.length, href, ...(rule.style ? { style: rule.style } : {}) })
    }
    return out
  }
}

/* ------------------------------------------------------------------ markers and formulas */

/** A marker character list (`a-h`, `abc`, `a-cx`) as a set of characters; null when it does not read. */
function markerSet(spec: string): Set<string> | null {
  if (!spec || spec.length > 40) return null
  const set = new Set<string>()
  for (let i = 0; i < spec.length; i += 1) {
    const ch = spec[i]!
    if (!/[A-Za-z0-9*†‡§]/.test(ch)) return null
    if (spec[i + 1] === '-' && spec[i + 2] !== undefined) {
      const end = spec[i + 2]!
      const from = ch.charCodeAt(0)
      const to = end.charCodeAt(0)
      if (to < from || to - from > 60) return null
      for (let code = from; code <= to; code += 1) set.add(String.fromCharCode(code))
      i += 2
    } else set.add(ch)
  }
  return set
}

const isSpace = (ch: string | undefined) => ch !== undefined && /\s/.test(ch)

/** "Total a", "Lead a, b": the trailing marker letters, after a space, after some text. */
function trailingFinder(markers: Set<string>, style: TextStyle): (text: string) => Hit[] {
  return (text) => {
    let i = text.length - 1
    if (i < 0 || !markers.has(text[i]!)) return []
    let first = i
    i -= 1
    for (;;) {
      let j = i
      while (j >= 0 && text[j] === ' ') j -= 1
      if (text[j] !== ',') break
      j -= 1
      while (j >= 0 && text[j] === ' ') j -= 1
      if (j < 0 || !markers.has(text[j]!) || (j > 0 && !isSpace(text[j - 1]) && text[j - 1] !== ',')) break
      first = j
      i = j - 1
    }
    // The marker stands alone after a space, and there is text before it.
    if (!isSpace(text[first - 1])) return []
    let k = first - 1
    while (k >= 0 && isSpace(text[k])) k -= 1
    if (k < 0) return []
    return [{ start: first, end: text.length, style }]
  }
}

/** "a Dried at 105 °C", "bAir dried": a marker letter opening the text, then a space and more, or a capital. */
function leadingFinder(markers: Set<string>, style: TextStyle): (text: string) => Hit[] {
  return (text) => {
    const first = text[0]
    if (first === undefined || !markers.has(first)) return []
    const next = text[1]
    const spaced = isSpace(next) && /\S/.test(text.slice(1))
    const capital = next !== undefined && /[A-Z]/.test(next)
    return spaced || capital ? [{ start: 0, end: 1, style }] : []
  }
}

const escapeRegex = (text: string) => text.replace(/[.*+?^${}()|[\]\\/-]/g, '\\$&')

function formulaFinder(formulas: readonly string[], style: TextStyle): (text: string) => Hit[] {
  const sorted = [...formulas].sort((a, b) => b.length - a.length).map(escapeRegex)
  const source = `(?<![A-Za-z0-9])(?:${sorted.join('|')})(?![A-Za-z0-9])`
  return (text) => {
    const out: Hit[] = []
    for (const match of text.matchAll(new RegExp(source, 'g'))) {
      const start = match.index ?? 0
      for (const digits of match[0].matchAll(/\d+/g)) {
        const at = start + (digits.index ?? 0)
        out.push({ start: at, end: at + digits[0].length, style })
      }
    }
    return out
  }
}

/* ------------------------------------------------------------------ patterns and terms */

function regexFinder(regex: RegExp, rule: TextRule): (text: string) => Hit[] {
  const group = rule.group
  return (text) => {
    const out: Hit[] = []
    regex.lastIndex = 0
    for (;;) {
      const match = regex.exec(text) as (RegExpExecArray & { indices?: Array<[number, number] | undefined> & { groups?: Record<string, [number, number] | undefined> } }) | null
      if (!match) break
      if (match[0] === '') {
        regex.lastIndex += 1
        continue
      }
      const range =
        group === undefined
          ? match.indices?.[0]
          : typeof group === 'number'
            ? match.indices?.[group]
            : match.indices?.groups?.[group]
      if (!range || range[0] === range[1]) continue
      const hit: Hit = { start: range[0], end: range[1], ...(rule.style ? { style: rule.style } : {}) }
      if (rule.link) {
        const href = linkHref(text.slice(range[0], range[1]), rule.schemes ?? DEFAULT_SCHEMES)
        if (href) hit.href = href
      }
      if (hit.style || hit.href) out.push(hit)
    }
    return out
  }
}

/* ------------------------------------------------------------------ compiling */

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === 'object' && value !== null && !Array.isArray(value)

/**
 * Reads and checks a list of rules. Throws a {@link TextRuleError} naming every problem: an unknown key, a pattern
 * outside the safe subset or past a limit, a built-in without what it needs, a style key the model does not have.
 */
export function compileTextRules(input: readonly TextRule[] | undefined, limits: Partial<TextRuleLimits> = {}): CompiledTextRules {
  const bounds: TextRuleLimits = { ...TEXT_RULE_LIMITS, ...limits }
  const rules: CompiledRule[] = []
  const problems: TextRuleProblem[] = []
  const list = input ?? []
  if (!Array.isArray(list)) throw new TextRuleError('the text rules are not a list')
  if (list.length > bounds.rules) problems.push({ rule: bounds.rules, key: '', message: `more than ${bounds.rules} rules` })
  const ids = new Set<string>()
  list.slice(0, bounds.rules).forEach((raw: unknown, index) => {
    const problem = (key: string, message: string) => problems.push({ rule: index, key, message })
    if (!isRecord(raw)) return problem('', 'is not an object')
    const rule = raw as TextRule
    for (const key of Object.keys(raw)) if (!RULE_KEYS.has(key)) problem(key, `has no key "${key}"`)
    if (rule.id !== undefined) {
      if (typeof rule.id !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9_.-]{0,99}$/.test(rule.id)) problem('id', 'is not an id (letters, digits, - _ .)')
      else if (ids.has(rule.id)) problem('id', `repeats the id "${rule.id}"`)
      else ids.add(rule.id)
    }
    const kinds = ['pattern', 'terms', 'builtIn'].filter((key) => raw[key] !== undefined)
    if (kinds.length !== 1) problem('', 'gives one of pattern, terms or builtIn')
    if (rule.style !== undefined) {
      if (!isRecord(rule.style)) problem('style', 'is not an object')
      else for (const key of Object.keys(rule.style)) if (!STYLE_KEYS.has(key)) problem(`style.${key}`, `is not a text style (${[...STYLE_KEYS].join(', ')})`)
    }
    if (rule.link !== undefined && typeof rule.link !== 'boolean') problem('link', 'is not true or false')
    if (rule.auto !== undefined && typeof rule.auto !== 'boolean') problem('auto', 'is not true or false')
    if (rule.ignoreCase !== undefined && typeof rule.ignoreCase !== 'boolean') problem('ignoreCase', 'is not true or false')
    for (const key of ['schemes', 'domains', 'formulas', 'terms'] as const) {
      const value = rule[key]
      if (value !== undefined && (!Array.isArray(value) || value.some((item) => typeof item !== 'string' || !item)))
        problem(key, 'is not a list of words')
    }
    let places: Set<TextPlace> | null = null
    if (rule.in !== undefined) {
      if (!Array.isArray(rule.in) || rule.in.some((place) => !PLACES.has(place))) problem('in', `is not a list of ${[...PLACES].join(', ')}`)
      else places = new Set(rule.in)
    }
    let find: ((text: string) => Hit[]) | null = null
    let bounded = false
    if (rule.pattern !== undefined) {
      bounded = true
      if (typeof rule.pattern !== 'string') problem('pattern', 'is not text')
      else if (rule.flags !== undefined && (typeof rule.flags !== 'string' || !/^i?$/.test(rule.flags))) problem('flags', 'is not "i" (the only flag)')
      else if (rule.group !== undefined && !(typeof rule.group === 'string' || (Number.isInteger(rule.group) && rule.group >= 0)))
        problem('group', 'is not a group number or name')
      else if (!rule.style && !rule.link) problem('', 'says neither style nor link: it would do nothing')
      else {
        const safe = safePattern(rule.pattern, bounds)
        if ('problem' in safe) problem('pattern', safe.problem)
        else {
          try {
            const regex = new RegExp(safe.source, `gd${rule.flags ?? ''}`)
            if (typeof rule.group === 'number' && rule.group > groupCount(regex)) problem('group', `names group ${rule.group}; the pattern has ${groupCount(regex)}`)
            else if (typeof rule.group === 'string' && !safe.source.includes(`(?<${rule.group}>`)) problem('group', `names a group "${rule.group}" the pattern does not have`)
            else find = regexFinder(regex, rule)
          } catch (error) {
            problem('pattern', `does not read: ${error instanceof Error ? error.message : String(error)}`)
          }
        }
      }
    } else if (rule.terms !== undefined) {
      bounded = true
      const terms = Array.isArray(rule.terms) ? rule.terms : []
      if (terms.length === 0) problem('terms', 'is empty')
      else if (terms.length > bounds.terms) problem('terms', `holds more than ${bounds.terms} terms`)
      else if (terms.some((term) => typeof term === 'string' && term.length > bounds.termLength)) problem('terms', `holds a term longer than ${bounds.termLength} characters`)
      else if (!rule.style && !rule.link) problem('', 'says neither style nor link: it would do nothing')
      else if (terms.every((term) => typeof term === 'string' && term)) {
        const alternatives = [...terms]
          .sort((a, b) => b.length - a.length)
          .map((term) => `${/^\w/.test(term) ? '(?<![A-Za-z0-9_])' : ''}${escapeRegex(term)}${/\w$/.test(term) ? '(?![A-Za-z0-9_])' : ''}`)
        const { group: _group, ...whole } = rule
        find = regexFinder(new RegExp(alternatives.join('|'), `gd${rule.ignoreCase ? 'i' : ''}`), whole)
      }
    } else if (rule.builtIn !== undefined) {
      if (!BUILT_INS.has(rule.builtIn)) problem('builtIn', `is not one of ${[...BUILT_INS].join(', ')}`)
      else if (rule.builtIn === 'links') find = linkFinder(rule)
      else if (rule.builtIn === 'formulaDigits') {
        const formulas = Array.isArray(rule.formulas) ? rule.formulas : []
        if (!formulas.length) problem('formulas', 'lists no formulas (formulaDigits sets only the formulas listed)')
        else if (formulas.length > bounds.terms) problem('formulas', `holds more than ${bounds.terms} formulas`)
        else if (formulas.some((f) => typeof f !== 'string' || !/^[A-Za-z0-9()]{1,40}$/.test(f) || !/\d/.test(f)))
          problem('formulas', 'holds something that is not a formula with a digit (letters, digits and brackets)')
        else find = formulaFinder(formulas, rule.style ?? { subscript: true })
      } else {
        const markers = markerSet(rule.markers ?? 'a-z')
        if (!markers) problem('markers', 'is not a list of marker characters with ranges (a-h)')
        else {
          const style = rule.style ?? { superscript: true }
          find = rule.builtIn === 'trailingMarker' ? trailingFinder(markers, style) : leadingFinder(markers, style)
        }
      }
    }
    if (find) {
      rules.push({
        index,
        ...(typeof rule.id === 'string' ? { id: rule.id } : {}),
        auto: rule.auto !== false,
        places,
        bounded,
        find,
      })
    }
  })
  if (problems.length) {
    const first = problems[0]!
    const more = problems.length > 1 ? ` and ${problems.length - 1} more` : ''
    throw new TextRuleError(`text rule ${first.rule}${first.key ? ` (${first.key})` : ''} ${first.message}${more}`, problems)
  }
  return { rules, limits: bounds }
}

const groupCount = (regex: RegExp): number => (new RegExp(`${regex.source}|`).exec('')?.length ?? 1) - 1

/* ------------------------------------------------------------------ applying */

interface Budget {
  matches: number
  limit: number
  skipped: SkippedText[]
  textLength: number
}

type Plain = string | TextRun

const inStretch = (inline: Inline): inline is Plain =>
  typeof inline === 'string' || (inline.type === 'text' && !inline.link)

const textOf = (run: Plain) => (typeof run === 'string' ? run : run.text)

function applyRule(list: readonly Inline[], rule: CompiledRule, budget: Budget): Inline[] {
  const out: Inline[] = []
  let i = 0
  while (i < list.length) {
    const inline = list[i]!
    if (!inStretch(inline)) {
      out.push(inline)
      i += 1
      continue
    }
    let j = i
    while (j < list.length && inStretch(list[j]!)) j += 1
    out.push(...applyToStretch(list.slice(i, j) as Plain[], rule, budget))
    i = j
  }
  return out
}

function applyToStretch(runs: readonly Plain[], rule: CompiledRule, budget: Budget): Inline[] {
  const text = runs.map(textOf).join('')
  if (!text) return [...runs]
  if (rule.bounded && text.length > budget.textLength) {
    budget.skipped.push({ rule: rule.index, ...(rule.id ? { id: rule.id } : {}), length: text.length, start: text.slice(0, 40) })
    return [...runs]
  }
  // Non-overlapping, in order: a later hit that overlaps an earlier one is dropped.
  const hits: Hit[] = []
  for (const hit of rule.find(text).sort((a, b) => a.start - b.start)) {
    const last = hits[hits.length - 1]
    if (hit.end <= hit.start || (last && hit.start < last.end)) continue
    hits.push(hit)
  }
  if (!hits.length) return [...runs]
  budget.matches += hits.length
  if (budget.matches > budget.limit)
    throw new TextRuleError(`the text rules found more than ${budget.limit} matches in one document (rule ${rule.index}${rule.id ? `, "${rule.id}"` : ''} reached it)`)
  const out: Inline[] = []
  let offset = 0
  for (const run of runs) {
    const own = textOf(run)
    const start = offset
    const end = offset + own.length
    offset = end
    if (!own) continue
    // The places this run is cut: every hit edge inside it.
    const cuts = new Set<number>([start, end])
    for (const hit of hits) {
      if (hit.start > start && hit.start < end) cuts.add(hit.start)
      if (hit.end > start && hit.end < end) cuts.add(hit.end)
    }
    const points = [...cuts].sort((a, b) => a - b)
    for (let k = 0; k + 1 < points.length; k += 1) {
      const a = points[k]!
      const b = points[k + 1]!
      const piece = own.slice(a - start, b - start)
      const hit = hits.find((h) => h.start <= a && h.end >= b)
      if (!hit) {
        out.push(typeof run === 'string' ? piece : { ...run, text: piece })
        continue
      }
      const base: TextRun = typeof run === 'string' ? { type: 'text', text: piece } : { ...run, text: piece }
      out.push({ ...base, ...(hit.style ?? {}), ...(hit.href ? { link: hit.href } : {}) })
    }
  }
  return out
}

/** Which rules a block's text takes: the automatic ones, plus those it names, in the document's order; none for `false`. */
function rulesFor(compiled: CompiledTextRules, setting: Block['textRules'], place: TextPlace): CompiledRule[] {
  if (setting === false) return []
  const named = new Set(setting ?? [])
  return compiled.rules.filter((rule) => (rule.auto || (rule.id !== undefined && named.has(rule.id))) && (!rule.places || rule.places.has(place)))
}

/**
 * Rich text with rules applied. `use` names the rules to apply whatever their `auto` and `in` (a filter's list); without
 * it, the rules a block with setting `setting` takes at `place`. Sentences are filled first. Returns the runs.
 */
export function ruleText(
  text: RichText,
  compiled: CompiledTextRules,
  options: { place?: TextPlace; setting?: Block['textRules']; use?: readonly string[] } = {},
): Inline[] {
  const budget = newBudget(compiled.limits)
  const rules = options.use
    ? compiled.rules.filter((rule) => rule.id !== undefined && options.use!.includes(rule.id))
    : rulesFor(compiled, options.setting, options.place ?? 'paragraph')
  let list: Inline[] = resolveSentences(text)
  if (!list.length) return list
  for (const rule of rules) list = applyRule(list, rule, budget)
  return list
}

const newBudget = (limits: TextRuleLimits): Budget => ({ matches: 0, limit: limits.matches, skipped: [], textLength: limits.textLength })

export interface AppliedTextRules {
  document: Document
  /** Stretches a pattern or term rule did not search because they were longer than `textLength`. */
  skipped: SkippedText[]
}

/**
 * The document with its `textRules` applied to every block's text (bodies, headers and footers, table cells, pinned
 * blocks), each rule where its places and the block's setting allow. Throws a {@link TextRuleError} for rules that do
 * not read, and when the rules together find more than `matches` matches. A document without rules comes back as it is.
 */
export function applyTextRules(document: Document, limits: Partial<TextRuleLimits> = {}): AppliedTextRules {
  if (!document.textRules?.length) return { document, skipped: [] }
  const compiled = compileTextRules(document.textRules, limits)
  const budget = newBudget(compiled.limits)
  const ctx: Ctx = { compiled, budget, part: null }
  const sections = document.sections.map((section): Section => {
    const next: Section = { ...section, body: blocksWith(section.body, { ...ctx, part: null }, undefined) }
    if (section.header) next.header = blocksWith(section.header, { ...ctx, part: 'header' }, undefined)
    if (section.footer) next.footer = blocksWith(section.footer, { ...ctx, part: 'footer' }, undefined)
    if (section.firstPage) {
      const first: NonNullable<Section['firstPage']> = {}
      if (section.firstPage.header) first.header = blocksWith(section.firstPage.header, { ...ctx, part: 'header' }, undefined)
      if (section.firstPage.footer) first.footer = blocksWith(section.firstPage.footer, { ...ctx, part: 'footer' }, undefined)
      next.firstPage = first
    }
    return next
  })
  const { textRules: _rules, ...rest } = document
  return { document: { ...rest, sections }, skipped: budget.skipped }
}

interface Ctx {
  compiled: CompiledTextRules
  budget: Budget
  /** In a header or footer: every place there is that part. */
  part: 'header' | 'footer' | null
}

function blocksWith(blocks: readonly Block[], ctx: Ctx, inherited: Block['textRules']): Block[] {
  return blocks.map((block) => blockWith(block, ctx, block.textRules === undefined ? inherited : block.textRules))
}

function blockWith(block: Block, ctx: Ctx, setting: Block['textRules']): Block {
  const rich = (text: RichText, place: TextPlace): RichText => {
    const at = ctx.part ?? place
    const rules = rulesFor(ctx.compiled, setting, at)
    if (!rules.length) return text
    let list: Inline[] = resolveSentences(text)
    for (const rule of rules) list = applyRule(list, rule, ctx.budget)
    return list.length === 1 && typeof list[0] === 'string' ? list[0] : list
  }
  const opt = (text: RichText | undefined, place: TextPlace) => (text === undefined ? undefined : rich(text, place))
  const nul = (text: RichText | null | undefined, place: TextPlace) => (text === null || text === undefined ? text : rich(text, place))
  const row = (r: FieldRow, label: TextPlace, value: TextPlace): FieldRow => ({ ...r, label: rich(r.label, label), value: nul(r.value, value) ?? null })
  const set = <T extends object>(base: T, over: Record<string, unknown>): T => {
    const out = { ...base } as Record<string, unknown>
    for (const [key, value] of Object.entries(over)) if (value !== undefined) out[key] = value
    return out as T
  }
  switch (block.type) {
    case 'heading':
      return { ...block, content: rich(block.content, 'heading') }
    case 'paragraph': {
      const place: TextPlace = block.style === 'caption' ? 'caption' : block.style === 'note' || block.style === 'small' ? 'note' : 'paragraph'
      return { ...block, content: rich(block.content, place) }
    }
    case 'list': {
      const items = (entries: readonly (ListItem | RichText)[]): (ListItem | RichText)[] =>
        entries.map((entry) => {
          const isItem = typeof entry === 'object' && entry !== null && !Array.isArray(entry) && !('type' in entry)
          if (!isItem) return rich(entry as RichText, 'list')
          const item = entry as ListItem
          return { ...item, content: rich(item.content, 'list'), ...(item.items ? { items: items(item.items) } : {}) }
        })
      return { ...block, items: items(block.items) }
    }
    case 'fields':
      return set(block, { title: opt(block.title, 'label'), rows: block.rows.map((r) => row(r, 'label', 'value')) })
    case 'table':
      return tableWith(block, ctx, setting, rich)
    case 'figure':
      return set(block, {
        caption: block.caption ? { ...block.caption, text: rich(block.caption.text, 'caption') } : undefined,
        note: opt(block.note, 'note'),
        placeholder: opt(block.placeholder, 'note'),
      })
    case 'figureGrid':
      return set(block, {
        caption: block.caption ? { ...block.caption, text: rich(block.caption.text, 'caption') } : undefined,
        items: block.items.map((item) => set(item, { label: opt(item.label, 'caption'), placeholder: opt(item.placeholder, 'note') })),
      })
    case 'signature':
      return set(block, {
        label: opt(block.label, 'signature'),
        name: nul(block.name, 'signature'),
        title: nul(block.title, 'signature'),
        lines: block.lines?.map((r) => row(r, 'label', 'value')),
      })
    case 'choices':
      return set(block, { label: opt(block.label, 'choice'), answer: nul(block.answer, 'value') })
    case 'pinned':
      return { ...block, blocks: blocksWith(block.blocks, ctx, setting) }
    case 'toc':
      return set(block, { title: opt(block.title, 'heading') })
    default:
      return block
  }
}

function tableWith(table: TableBlock, ctx: Ctx, setting: Block['textRules'], rich: (text: RichText, place: TextPlace) => RichText): TableBlock {
  const rows = table.rows.map((r) => ({
    ...r,
    cells: r.cells.map((input): CellInput => {
      const cell = toCell(input)
      if (cell.blocks?.length) return { ...cell, blocks: blocksWith(cell.blocks, ctx, setting) }
      if (cell.content === null || cell.content === undefined) return input
      const content = rich(cell.content, 'cell')
      // A cell given as text alone stays text alone.
      return input === null || typeof input === 'string' || Array.isArray(input) || (typeof input === 'object' && 'type' in input) ? content : { ...cell, content }
    }),
  }))
  const out: TableBlock = { ...table, rows }
  if (table.caption) out.caption = { ...table.caption, text: rich(table.caption.text, 'caption') }
  if (table.note !== undefined) out.note = rich(table.note, 'note')
  return out
}

