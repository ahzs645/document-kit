/**
 * Rendering a compiled recipe against data: every tag read, control nodes expanded, code components built, each part's
 * blocks filled, hand edits applied, and the to-complete list made by the model's own rules. The result is a kit
 * `Document`, ready for `writeDocx`. Pure: no DOM, no network.
 *
 * Missing values follow the two rules both applications have: a missing value in a field, a cell or a line of text
 * prints blank (a blank rule, or a null the model prints blank) and is listed; a sentence whose slot is missing is not
 * printed and is listed. A missing value tested by a condition is false, and a missing list repeats nothing.
 */
import {
  compileTextRules,
  plainText,
  ruleText,
  toComplete as modelToComplete,
  type CompiledTextRules,
  type TextRule,
  type Block,
  type CustomRegistry,
  type Document,
  type Inline,
  type Section,
  type ToCompleteItem,
} from '../model'
import { evaluateCondition, MISSING, type ConditionReader, type ConditionValue } from '../template/templateTags'
import type { CodeComponents } from './code'
import { compileCodeBlocks, TO_COMPLETE_KIND } from './compile'
import { applyEdits, type EditTarget, type RecipeEdits } from './edits'
import { jsonPath, RecipeError } from './errors'
import { hasOwn, isRecord, sameJson } from './json'
import { readCompiledRecipe } from './read'
import { CONTROL_KEYS, controlKind, ITEM_OF, isListPos, keysOf, type Pos } from './schema'
import { missingOf, readCondition, readLoop, readString, type Condition, type FieldSeg, type Loop, type ParsedString, type Seg } from './tags'
import {
  applyFilters,
  formatting,
  isEmptyValue,
  isWrapped,
  mapSource,
  numberOf,
  sourceMap,
  unwrapDeep,
  valueText,
  type FormatOptions,
  type Formatting,
  type RecipeFilter,
  type SourceMap,
} from './values'
import type { CompiledPart, CompiledRecipe, Json, JsonObject, MissingRule } from './types'

/** Bounds on a render. */
export const RENDER_LIMITS = {
  /** Passes of every loop together (part loops, `for` nodes and inline `{{ for }}`). */
  loopPasses: 100_000,
  /** Blocks made, at every depth (cells' blocks included). */
  blocks: 50_000,
  /** How deep code components may nest through the blocks they return. */
  includeDepth: 16,
}
export type RenderLimits = typeof RENDER_LIMITS

export interface RenderOptions extends FormatOptions {
  /** Expands custom blocks for the to-complete list; given, every custom kind the recipe uses must be in it. */
  registry?: CustomRegistry
  /** The application's filters by name. */
  filters?: Readonly<Record<string, RecipeFilter>>
  /** Code components; every one the compiled recipe lists must be here. */
  code?: CodeComponents
  /** Hand edits by key (see `RenderResult.edits.targets`). */
  edits?: RecipeEdits
  /** The source an edited value carries. `edited` by default. */
  editSource?: string
  /** Where values came from, by data path: `"client.name": "lab"`, `"tables[].rows": "instrument"`. */
  sources?: Readonly<Record<string, string>>
  /** Options laid over the recipe's own. */
  options?: JsonObject
  limits?: Partial<RenderLimits>
}

/** A to-complete item, as the model lists it, with the part instance it falls in and a `value` kind for other gaps. */
export type RecipeToCompleteItem = Omit<ToCompleteItem, 'kind'> & {
  kind: ToCompleteItem['kind'] | 'value'
  /** The part instance id, or null for a header or footer. */
  partId: string | null
}

export interface RenderedPart {
  id: string
  component: string
  title: string
  section: number
  /** Its blocks in the section's body: `body.slice(start, end)`. */
  start: number
  end: number
  /** Top-level blocks it printed: zero when its condition was false or it printed nothing. */
  blocks: number
  /** Values it left to complete. */
  toComplete: number
  /** Times it was printed (its `for` passes that met its `if`). */
  repeats: number
  /** The part's own `meta` and its component's, as the recipe and component give them. */
  meta?: JsonObject
  componentMeta?: JsonObject
}

export interface RenderResult {
  document: Document
  parts: RenderedPart[]
  toComplete: RecipeToCompleteItem[]
  edits: {
    /** Every editable value, in reading order; an edited one carries its edit. */
    targets: EditTarget[]
    /** Edits for this recipe's parts whose place nothing printed: kept, never applied elsewhere. */
    orphans: string[]
    /** Edits for parts this recipe does not have (another recipe's): left alone. */
    otherParts: string[]
  }
}

export interface CheckProblem {
  code: 'missing' | 'missing-code' | 'missing-filter' | 'missing-custom-kind' | 'bad-value' | 'not-text' | 'code-failed'
  /** The data path read, with list indexes (`tables.1.rows.0.value`). */
  path?: string
  use?: 'value' | 'condition' | 'list'
  part: string | null
  /** The JSON path in the compiled recipe. */
  at: string
  message: string
}

/* ------------------------------------------------------------------ scope and lookup */

interface Bound {
  value: unknown
  /** The data path it was read from, or null (a literal, a parameter's own value). */
  path: string[] | null
  source?: string
  /** How specific the source is (the path length it was given at); -1 for none. */
  depth: number
}

interface LoopDetails {
  odd: boolean
  index: number
  number: number
  first: boolean
  last: boolean
  length: number
}

interface Scope {
  params: ReadonlyMap<string, Bound>
  vars: ReadonlyMap<string, Bound>
  loop: LoopDetails | null
  stack: readonly string[]
}

const OMIT = Symbol('omit')
type Out = ReadonlyArray<string | number>

interface State {
  compiled: CompiledRecipe
  data: Record<string, unknown>
  options: Record<string, unknown>
  sources: SourceMap
  filters: Readonly<Record<string, RecipeFilter>>
  code: CodeComponents | undefined
  format: Formatting
  limits: RenderLimits
  check: boolean
  problems: CheckProblem[]
  problemKeys: Set<string>
  gaps: RecipeToCompleteItem[]
  /** Set by a nullable place that came out null, so the row or cell holding it carries the source it would have had. */
  lastNullSource: string | undefined
  part: string | null
  section: number
  where: 'body' | 'header' | 'footer' | 'firstHeader' | 'firstFooter'
  loopPasses: number
  blocks: number
  strings: Map<string, ParsedString>
  conditions: Map<string, Condition>
  loops: Map<string, Loop>
  /** The missing-value rule of the part being rendered (the recipe's in headers and footers). */
  missing: MissingRule
  /** The recipe's text rules, read once, for the `rules` filter. */
  rules: CompiledTextRules | null
  /** Placeholders named after their tag, which a field row names after its label instead. */
  autoNamed: WeakSet<object>
}

const where = (out: Out) => jsonPath(out)

function problem(state: State, p: Omit<CheckProblem, 'part'>) {
  const key = `${p.code}|${p.path ?? ''}|${p.at}`
  if (state.problemKeys.has(key)) return
  state.problemKeys.add(key)
  state.problems.push({ ...p, part: state.part })
}

function lookup(state: State, path: readonly string[], scope: Scope, use: 'value' | 'condition' | 'list', out: Out): Bound | typeof MISSING {
  const [first, ...rest] = path
  let current: Bound
  let steps: readonly string[] = rest
  if (first === undefined) return MISSING
  const fromVar = scope.vars.get(first)
  if (fromVar) current = fromVar
  else if (first === 'loop' && scope.loop) current = { value: scope.loop, path: null, depth: -1 }
  else if (first === 'params') {
    const [name, ...after] = rest
    const param = name === undefined ? undefined : scope.params.get(name)
    if (!param) return missing(state, path, use, out, scope)
    current = param
    steps = after
  } else if (first === 'options') current = { value: state.options, path: null, depth: -1 }
  else {
    // A key holding the whole dotted text is read as typed first, as the Word engine does.
    const whole = path.join('.')
    if (path.length > 1 && hasOwn(state.data, whole)) {
      current = { value: state.data[whole], path: [whole], depth: -1 }
      steps = []
    } else {
      current = { value: state.data, path: [], depth: -1 }
      steps = path
    }
  }
  let { value, source, depth } = current
  const trail = current.path ? [...current.path] : null
  if (isWrapped(value)) {
    source = value.$source ?? source
    depth = trail ? trail.length : depth
    value = value.$value
  }
  for (const step of steps) {
    if (!(isRecord(value) || Array.isArray(value)) || !hasOwn(value, step)) return missing(state, path, use, out, scope)
    value = (value as Record<string, unknown>)[step]
    trail?.push(step)
    if (isWrapped(value)) {
      source = value.$source ?? source
      depth = trail ? trail.length : depth + 1
      value = value.$value
    }
  }
  if (trail) {
    const mapped = mapSource(state.sources, trail)
    if (mapped && mapped.depth > depth) {
      source = mapped.source
      depth = mapped.depth
    }
  }
  return { value, path: trail, ...(source !== undefined ? { source } : {}), depth }
}

function missing(state: State, path: readonly string[], use: 'value' | 'condition' | 'list', out: Out, scope: Scope): typeof MISSING {
  if (!state.check || use === 'condition') return MISSING
  // A parameter left unbound was reported where its binding read the data (or is optional); options are constants; a
  // condition reads a missing value as false, by design; a key a literal structure leaves out is not data.
  const [first, name] = path
  if (first === undefined || first === 'options' || first === 'loop') return MISSING
  const base = first === 'params' ? scope.params.get(name ?? '') : scope.vars.get(first)
  if (first === 'params' && !base) return MISSING
  if (base && base.path === null) return MISSING
  const data = dataPath(state, path, scope)
  problem(state, { code: 'missing', path: data, use, at: where(out), message: `the data has no ${data}` })
  return MISSING
}

/** The data path a scope-relative path names, with list indexes: for the to-complete list and checks. */
function dataPath(state: State, path: readonly string[], scope: Scope): string {
  const [first, ...rest] = path
  const base = first === undefined ? undefined : first === 'params' ? scope.params.get(rest[0] ?? '') : scope.vars.get(first)
  if (base?.path) return [...base.path, ...(first === 'params' ? rest.slice(1) : rest)].join('.')
  return path.join('.')
}

/* ------------------------------------------------------------------ tags */

function parsed(state: State, text: string, out: Out): ParsedString {
  let read = state.strings.get(text)
  if (!read) {
    try {
      read = readString(text)
    } catch (error) {
      throw new RecipeError('bad-tag', error instanceof Error ? error.message : String(error), where(out))
    }
    state.strings.set(text, read)
  }
  return read
}

function filtered(state: State, field: FieldSeg, scope: Scope, out: Out): Bound | typeof MISSING {
  const found = lookup(state, field.path, scope, 'value', out)
  if (found === MISSING || !field.filters.length) return found
  const value = applyFilters(found.value, field.filters, state.format, state.filters, field.text, where(out), state.check ? (name) => problem(state, { code: 'missing-filter', at: where(out), message: `no filter "${name}"` }) : undefined, rulesFilter(state, field.text, out))
  return { ...found, value }
}

/** The `rules` filter: a value's text with the named text rules applied, as runs. */
function rulesFilter(state: State, what: string, out: Out) {
  return (value: unknown, ids: readonly string[]): unknown => {
    if (!state.rules) throw new RecipeError('bad-value', `${what} names text rules, and the recipe's document has none`, where(out))
    const unknown = ids.find((id) => !state.rules?.rules.some((rule) => rule.id === id))
    if (unknown !== undefined) throw new RecipeError('bad-value', `${what} names the text rule "${unknown}", which the recipe's document does not have`, where(out))
    const plain = unwrapDeep(value)
    const rich = asRich(plain) ?? valueText(plain, what, where(out)) ?? ''
    return ruleText(rich, state.rules, { use: ids })
  }
}

function test(state: State, text: string, scope: Scope, out: Out): boolean {
  let condition = state.conditions.get(text)
  if (!condition) {
    try {
      condition = readCondition(text)
    } catch (error) {
      throw new RecipeError('bad-tag', error instanceof Error ? error.message : String(error), where(out))
    }
    state.conditions.set(text, condition)
  }
  return testCondition(state, condition, scope, out)
}

function testCondition(state: State, condition: Condition, scope: Scope, out: Out): boolean {
  const reader: ConditionReader = {
    lookup: (path) => {
      const found = lookup(state, path, scope, 'condition', out)
      return found === MISSING ? MISSING : (unwrapDeep(found.value) as ConditionValue)
    },
    incomparable: (why) => {
      throw new RecipeError('bad-value', `the condition "${condition.text}" compares ${why.left} with ${why.right}`, where(out))
    },
    filter: (chain, value) =>
      applyFilters(value, condition.chains[chain] ?? [], state.format, state.filters, condition.text, where(out), state.check ? () => undefined : undefined) as never,
    readNumber: (text) => numberOf(text, state.format.ways),
    compareTexts: (operator, left, right) => {
      if (state.format.numericText !== 'number') return undefined
      const a = numberOf(left, state.format.ways)
      const b = numberOf(right, state.format.ways)
      if (a === null || b === null) return undefined
      switch (operator) {
        case '==': return a === b
        case '!=': return a !== b
        case '<': return a < b
        case '<=': return a <= b
        case '>': return a > b
        case '>=': return a >= b
      }
    },
  }
  return evaluateCondition(condition.expression, reader) === true
}

function loopOf(state: State, text: string, out: Out): Loop {
  let loop = state.loops.get(text)
  if (!loop) {
    try {
      loop = readLoop(text)
    } catch (error) {
      throw new RecipeError('bad-tag', error instanceof Error ? error.message : String(error), where(out))
    }
    state.loops.set(text, loop)
  }
  return loop
}

/** The scopes of a loop's passes: one per item, each with the item and `loop`. */
function passes(state: State, loop: Loop, scope: Scope, out: Out): Scope[] {
  const found = lookup(state, loop.list, scope, 'list', out)
  if (found === MISSING || found.value === null || found.value === undefined) return []
  if (!Array.isArray(found.value)) throw new RecipeError('bad-value', `"${loop.listText}" is looped over but is not a list`, where(out))
  const list = found.value
  state.loopPasses += list.length
  if (state.loopPasses > state.limits.loopPasses) throw new RecipeError('limit', `loops passed more than ${state.limits.loopPasses} times in all`, where(out))
  return list.map((item, index) => {
    const vars = new Map(scope.vars)
    let value = item
    let source = found.source
    let depth = found.depth
    if (isWrapped(value)) {
      source = value.$source ?? source
      value = value.$value
      depth = (found.path?.length ?? 0) + 1
    }
    vars.set(loop.name, { value, path: found.path ? [...found.path, String(index)] : null, ...(source !== undefined ? { source } : {}), depth })
    return { ...scope, vars, loop: { index, odd: index % 2 === 0, number: index + 1, first: index === 0, last: index === list.length - 1, length: list.length } }
  })
}

/** A lone tag's value: whether it is empty (null, missing, '') and whether the data lacks it (absent). */
interface Lone {
  value: unknown
  empty: boolean
  /** The data does not have it (an optional parameter nobody gave is not absent: it is simply not given). */
  absent: boolean
  /** The data path, with list indexes. */
  path: string
  source?: string
}

function lone(state: State, field: FieldSeg, scope: Scope, out: Out): Lone {
  const found = filtered(state, field, scope, out)
  if (found === MISSING) {
    const [root, name] = field.path
    const path = dataPath(state, field.path, scope)
    const source = mapSource(state.sources, path.split('.'))?.source
    // Absent from the data: read from the data, or from an item or parameter that came from it. A key a literal
    // structure leaves out (an optional `what` in a list written in the recipe) is not data the data lacks.
    const base = root === undefined ? undefined : root === 'params' ? scope.params.get(name ?? '') : scope.vars.get(root)
    const absent = root !== 'options' && root !== 'loop' && (base ? base.path !== null : root !== 'params')
    return { value: undefined, empty: true, absent, path, ...(source ? { source } : {}) }
  }
  const path = found.path ? found.path.join('.') : dataPath(state, field.path, scope)
  return { value: found.value, empty: isEmptyValue(found.value), absent: found.value === undefined, path, ...(found.source ? { source: found.source } : {}) }
}

/** A piece of text: printed text (with its source), runs a value gave, or a value with none (and its missing rule). */
type Piece =
  | { text: string; source?: string }
  | { runs: Inline[]; source?: string }
  | { missing: string; path: string; absent: boolean; source?: string; rule?: { mode: MissingRule; name?: string } }

function pieces(state: State, segs: readonly Seg[], scope: Scope, out: Out, into: Piece[] = []): Piece[] {
  for (const seg of segs) {
    if (seg.t === 'text') into.push({ text: seg.text })
    else if (seg.t === 'field') {
      const value = lone(state, seg, scope, out)
      const rule = missingOf(seg.filters)
      if (value.empty) into.push({ missing: seg.text, path: value.path, absent: value.absent, ...(value.source ? { source: value.source } : {}), ...(rule ? { rule } : {}) })
      else {
        const plain = unwrapDeep(value.value)
        const rich = asRich(plain)
        if (rich) into.push({ runs: rich, ...(value.source ? { source: value.source } : {}) })
        else into.push({ text: valueText(plain, seg.text, where(out)) ?? '', ...(value.source ? { source: value.source } : {}) })
      }
    } else if (seg.t === 'if') {
      const branch = seg.branches.find((b) => testCondition(state, b.condition, scope, out))
      const body = branch ? branch.body : seg.otherwise
      if (body) pieces(state, body, scope, out, into)
    } else {
      for (const pass of passes(state, seg, scope, out)) pieces(state, seg.body, pass, out, into)
    }
  }
  return into
}

/** A value the data lacks at a place that prints no blank of its own (plain text, a number): listed as `value`. */
function gap(state: State, what: string, path: string, source?: string) {
  state.gaps.push({ kind: 'value', what, missing: [path], section: state.section, part: state.where, partId: state.part, ...(source ? { source } : {}) })
}

/* ------------------------------------------------------------------ values by place */

type Style = Record<string, unknown>

/**
 * Rich text from a string: text runs (with their sources) and a blank rule for each value with none. A lone tag with
 * no value gives a blank rule (`blank`), null (`null`: the model prints it blank and lists it), or nothing (`omit`).
 */
function richFromString(state: State, text: string, scope: Scope, out: Out, style: Style, empty: 'blank' | 'null' | 'omit'): Inline[] | null {
  const read = parsed(state, text, out)
  if (read.kind === 'plain') return [Object.keys(style).length ? ({ ...style, type: 'text', text: read.text } as Inline) : read.text]
  if (read.kind === 'raw') {
    const value = lone(state, read.field, scope, out)
    if (value.empty) {
      const rule = missingOf(read.field.filters)
      const placeholder = (rule?.mode ?? state.missing) === 'placeholder'
      if (empty === 'null' && !placeholder) {
        state.lastNullSource = value.source
        return null
      }
      if (empty === 'omit') {
        if (value.absent) gap(state, read.field.text, value.path, value.source)
        return null
      }
      return [missingRun(state, style, read.field.text, value.source, placeholder, rule?.name)]
    }
    const plain = unwrapDeep(value.value)
    const rich = asRich(plain)
    if (rich) return rich.map((inline) => withSource(inline, value.source))
    const printed = valueText(plain, read.field.text, where(out)) ?? ''
    return [value.source || Object.keys(style).length ? ({ ...style, type: 'text', text: printed, ...(value.source ? { source: value.source } : {}) } as Inline) : printed]
  }
  const runs: Inline[] = []
  for (const piece of pieces(state, read.segs, scope, out)) {
    if ('missing' in piece)
      runs.push(missingRun(state, style, piece.missing, piece.source, (piece.rule?.mode ?? state.missing) === 'placeholder', piece.rule?.name))
    else if ('runs' in piece) runs.push(...piece.runs.map((run) => withSource(Object.keys(style).length && typeof run === 'string' ? ({ ...style, type: 'text', text: run } as Inline) : run, piece.source)))
    else if (piece.source || Object.keys(style).length) runs.push({ ...style, type: 'text', text: piece.text, ...(piece.source ? { source: piece.source } : {}) } as Inline)
    else runs.push(piece.text)
  }
  return joinRuns(merge(runs))
}

/**
 * One written run's pieces joined back where they print alike: a text run whose tags all filled stays one run, as it
 * was written. Pieces stay apart only where they differ (a blank for a missing value, a value with another source,
 * runs the data gave whole).
 */
function joinRuns(runs: Inline[]): Inline[] {
  const out: Inline[] = []
  for (const run of runs) {
    const last = out[out.length - 1]
    if (last !== undefined && typeof last === 'object' && typeof run === 'object' && last.type === 'text' && run.type === 'text') {
      const { text: a, ...lookA } = last
      const { text: b, ...lookB } = run
      if (sameJson(lookA as never, lookB as never)) {
        out[out.length - 1] = { ...last, text: a + b }
        continue
      }
    }
    out.push(run)
  }
  return out
}

/** A value with none, printed: a blank rule, or under the placeholder rule its bracketed name, marked. */
function missingRun(state: State, style: Style, tag: string, source: string | undefined, placeholder: boolean, name?: string): Inline {
  const run = {
    ...style,
    type: 'blank',
    what: name ?? tag,
    ...(source ? { source } : {}),
    ...(placeholder ? { show: 'placeholder' } : {}),
  } as Inline
  if (placeholder && name === undefined) state.autoNamed.add(run as object)
  return run
}

const INLINE_TYPES = new Set(['text', 'field', 'break', 'tab', 'blank', 'image', 'sentence', 'checkbox', 'footnote'])
const isInlineObject = (value: unknown) => isRecord(value) && typeof value.type === 'string' && INLINE_TYPES.has(value.type)

/** A data value that is already rich text (a run, or runs), or null. */
function asRich(value: unknown): Inline[] | null {
  if (isInlineObject(value)) return [value as unknown as Inline]
  if (Array.isArray(value) && value.length && value.every((item) => typeof item === 'string' || isInlineObject(item))) return value as Inline[]
  return null
}

const withSource = (inline: Inline, source: string | undefined): Inline => {
  if (!source) return inline
  if (typeof inline === 'string') return { type: 'text', text: inline, source }
  if ((inline.type === 'text' || inline.type === 'blank' || inline.type === 'sentence' || inline.type === 'checkbox') && inline.source === undefined) return { ...inline, source }
  return inline
}

function merge(runs: Inline[]): Inline[] {
  const out: Inline[] = []
  for (const run of runs) {
    const last = out[out.length - 1]
    if (typeof run === 'string' && typeof last === 'string') out[out.length - 1] = last + run
    else if (run !== '') out.push(run)
  }
  return out.length ? out : ['']
}

const richValue = (runs: Inline[]): Inline[] | string => (runs.length === 1 && typeof runs[0] === 'string' ? runs[0] : runs)

/** Plain text from a string: a value with none prints as nothing (listed when the data lacks it). OMIT for a lone empty tag. */
function plainFromString(state: State, text: string, scope: Scope, out: Out): string | typeof OMIT {
  const read = parsed(state, text, out)
  if (read.kind === 'plain') return read.text
  const segs = read.kind === 'raw' ? [read.field] : read.segs
  let printed = ''
  let any = false
  for (const piece of pieces(state, segs, scope, out)) {
    if ('missing' in piece) {
      if (piece.absent) gap(state, piece.missing, piece.path, piece.source)
    } else {
      printed += 'runs' in piece ? plainText(piece.runs) : piece.text
      any = true
    }
  }
  return read.kind === 'raw' && !any ? OMIT : printed
}

/** A lone tag's raw value (wrappers kept when `keepWrappers`), or a mixed string's text. OMIT when a lone tag is empty. */
function rawFromString(state: State, text: string, scope: Scope, out: Out, keepWrappers = false): unknown {
  const read = parsed(state, text, out)
  if (read.kind === 'plain') return read.text
  if (read.kind === 'raw') {
    const value = lone(state, read.field, scope, out)
    if (value.value === null || value.value === undefined) return OMIT
    if (keepWrappers && value.source !== undefined) return { $value: value.value, $source: value.source }
    return keepWrappers ? value.value : unwrapDeep(value.value)
  }
  return plainFromString(state, text, scope, out)
}

/** Places where null stays null (the model prints it blank, or a picture's placeholder); elsewhere a key with none is left out. */
const KEEPS_NULL: ReadonlySet<Pos> = new Set(['richNull', 'cell', 'image', 'file', 'slots'])

function evalValue(state: State, value: unknown, pos: Pos, scope: Scope, out: Out): unknown {
  if (isListPos(pos) && Array.isArray(value)) {
    const items = evalList(state, value, pos, scope, out)
    return pos === 'rich' || pos === 'richNull' || pos === 'richOpt' ? richValue(merge(items as Inline[])) : items
  }
  if (typeof value === 'string') {
    switch (pos) {
      case 'inline':
        return richFromString(state, value, scope, out, {}, 'blank') ?? []
      case 'rich':
      case 'listItem':
        return richValue(richFromString(state, value, scope, out, {}, 'blank') ?? [''])
      case 'richOpt': {
        const runs = richFromString(state, value, scope, out, {}, 'omit')
        return runs === null ? OMIT : richValue(runs)
      }
      case 'caption': {
        // A caption the data gives whole ({ text, numbered }) is taken as it is.
        const read = parsed(state, value, out)
        if (read.kind === 'raw') {
          const whole = unwrapDeep(lone(state, read.field, scope, out).value)
          if (isRecord(whole) && !('type' in whole) && 'text' in whole) return whole
        }
        const runs = richFromString(state, value, scope, out, {}, 'omit')
        return runs === null ? OMIT : { text: richValue(runs) }
      }
      case 'richNull':
      case 'cell': {
        if (pos === 'cell') {
          // A cell the data gives whole ({ content, shading, … }) is taken as it is.
          const read = parsed(state, value, out)
          if (read.kind === 'raw') {
            const whole = lone(state, read.field, scope, out)
            const plain = unwrapDeep(whole.value)
            if (isRecord(plain) && !('type' in plain) && ('content' in plain || 'blocks' in plain)) return plain
          }
        }
        state.lastNullSource = undefined
        const runs = richFromString(state, value, scope, out, {}, 'null')
        // A bare cell that is blank keeps the source it would have had, as a cell object.
        if (runs === null && pos === 'cell' && state.lastNullSource) return { content: null, source: state.lastNullSource }
        return runs === null ? null : richValue(runs)
      }
      case 'text':
        return plainFromString(state, value, scope, out)
      case 'literal':
        return value
      case 'scalar': {
        const read = parsed(state, value, out)
        if (read.kind === 'plain') return read.text
        if (read.kind === 'mixed') return plainFromString(state, value, scope, out)
        const one = lone(state, read.field, scope, out)
        if (one.empty) {
          if (one.absent) gap(state, read.field.text, one.path, one.source)
          return OMIT
        }
        const raw = unwrapDeep(one.value)
        if (raw !== null && typeof raw === 'object' && !(raw instanceof Date))
          throw new RecipeError('bad-value', `${value} gives ${Array.isArray(raw) ? 'a list' : 'an object'} where a single value goes`, where(out))
        return raw
      }
      case 'image': {
        const raw = rawFromString(state, value, scope, out)
        return picture(raw === OMIT ? null : raw, value, out)
      }
      case 'file': {
        const raw = rawFromString(state, value, scope, out)
        return fileBytes(raw === OMIT ? null : raw, value, out)
      }
      case 'slots': {
        const raw = rawFromString(state, value, scope, out)
        return raw === OMIT ? null : raw
      }
      case 'any':
        return rawFromString(state, value, scope, out)
      default: {
        // A lone tag standing for a whole structure (rows, cells, blocks): the data gives it in the model's shape.
        const raw = rawFromString(state, value, scope, out)
        return raw === OMIT ? (isListPos(pos) ? [] : OMIT) : raw
      }
    }
  }
  if (value === null || typeof value !== 'object') return value
  if (ArrayBuffer.isView(value)) return value
  if (Array.isArray(value)) {
    if (pos === 'listItem' || pos === 'cell') return richValue(merge(evalList(state, value, 'rich', scope, out) as Inline[]))
    return evalList(state, value, 'any', scope, out)
  }
  const record = value as Record<string, unknown>
  switch (pos) {
    case 'any':
    case 'slots': {
      const result: Record<string, unknown> = {}
      for (const [key, member] of Object.entries(record)) {
        const evaluated = evalValue(state, member, pos === 'slots' ? 'slots' : 'any', scope, [...out, key])
        if (evaluated === OMIT) continue
        result[key] = pos === 'slots' ? slotValue(evaluated, key, out) : evaluated
      }
      return result
    }
    case 'image':
    case 'file':
      return record
    case 'block':
      return evalBlock(state, record, scope, out)
    case 'caption': {
      // A caption whose text has no value is left out whole.
      const caption = evalObject(state, record, keysOf('caption', record) ?? {}, scope, out)
      return caption.text === undefined ? OMIT : caption
    }
    case 'inline':
    case 'rich':
    case 'richNull':
    case 'richOpt':
    case 'cell':
    case 'listItem': {
      if (typeof record.type === 'string') {
        const runs = evalInline(state, record, scope, out)
        if (pos === 'inline') return runs
        // Written as one run, it stays one run when it fills to one.
        return runs.length === 1 && typeof runs[0] === 'object' ? runs[0] : richValue(runs)
      }
      if (pos === 'cell') return evalObject(state, record, keysOf('cell', record) ?? {}, scope, out)
      if (pos === 'listItem') return evalObject(state, record, keysOf('listItem', record) ?? {}, scope, out)
      return record
    }
    default:
      return evalObject(state, record, keysOf(pos, record) ?? {}, scope, out)
  }
}

function slotValue(value: unknown, key: string, out: Out): string | number | null {
  if (value === null || value === undefined || value === '') return null
  if (typeof value === 'string' || typeof value === 'number') return value
  return valueText(value, `the slot "${key}"`, where(out))
}

function picture(value: unknown, text: string, out: Out): unknown {
  if (value === null || value === undefined) return null
  if (value instanceof Uint8Array) return { bytes: value }
  if (isRecord(value) && value.bytes instanceof Uint8Array) return value
  throw new RecipeError('bad-value', `${text} is not a picture (bytes, or { bytes })`, where(out))
}

function fileBytes(value: unknown, text: string, out: Out): unknown {
  if (value === null || value === undefined) return null
  if (value instanceof Uint8Array) return { bytes: value }
  if (isRecord(value) && value.bytes instanceof Uint8Array) return value
  throw new RecipeError('bad-value', `${text} is not a file (bytes, or { bytes })`, where(out))
}

function evalObject(state: State, record: Record<string, unknown>, keys: Record<string, Pos>, scope: Scope, out: Out): Record<string, unknown> {
  const result: Record<string, unknown> = {}
  let nullSource: string | undefined
  for (const [key, member] of Object.entries(record)) {
    if (key === 'type') {
      result.type = member
      continue
    }
    const pos = keys[key] ?? 'any'
    state.lastNullSource = undefined
    const value = evalValue(state, member, pos, scope, [...out, key])
    if (value === OMIT || (value === null && !KEEPS_NULL.has(pos))) continue
    if (value === null && pos === 'richNull' && state.lastNullSource) nullSource = state.lastNullSource
    result[key] = value
  }
  // A blank that the data says would come from somewhere carries that source onto its row or cell.
  if (nullSource && result.source === undefined && ('value' in keys || 'content' in keys)) result.source = nullSource
  // A field's placeholder is named after the field, unless its tag named it.
  if ('value' in keys && 'label' in keys && Array.isArray(result.value) && result.value.length === 1) {
    const only = result.value[0] as Record<string, unknown> | undefined
    if (only && state.autoNamed.has(only)) {
      const label = plainText(result.label as never).replace(/:\s*$/, '').trim()
      if (label) result.value = [{ ...only, what: typeof result.what === 'string' ? result.what : label }]
    }
  }
  return result
}

function evalBlock(state: State, record: Record<string, unknown>, scope: Scope, out: Out): Block {
  state.blocks += 1
  if (state.blocks > state.limits.blocks) throw new RecipeError('limit', `the document holds more than ${state.limits.blocks} blocks`, where(out))
  const keys = keysOf('block', record)
  if (!keys) throw new RecipeError('bad-node', `a block of type ${String(record.type)} the model does not have`, where(out))
  return evalObject(state, record, keys, scope, out) as unknown as Block
}

function evalInline(state: State, record: Record<string, unknown>, scope: Scope, out: Out): Inline[] {
  const type = record.type as string
  const keys = keysOf('inline', record) ?? {}
  if (type === 'text' && typeof record.text === 'string') {
    const { text, ...rest } = record
    const style = evalObject(state, rest, keys, scope, out)
    delete style.type
    const runs = richFromString(state, text as string, scope, [...out, 'text'], style, 'blank') ?? []
    return runs.map((run) => (typeof run === 'string' ? ({ ...style, type: 'text', text: run } as Inline) : run))
  }
  if (type === 'image') {
    const run = evalObject(state, record, keys, scope, out)
    if (run.image === null || run.image === undefined) return []
    return [run as unknown as Inline]
  }
  if (type === 'sentence') {
    const run = evalObject(state, record, keys, scope, out) as Record<string, unknown>
    if (run.missing === undefined && state.missing === 'placeholder') run.missing = 'placeholder'
    if (run.source === undefined) {
      const sources = new Set<string>()
      for (const member of Object.values(isRecord(record.values) ? record.values : {})) {
        if (typeof member !== 'string') continue
        const read = parsed(state, member, out)
        if (read.kind !== 'raw') continue
        const found = lookup({ ...state, check: false }, read.field.path, scope, 'value', out)
        if (found !== MISSING && found.source) sources.add(found.source)
      }
      if (sources.size === 1) run.source = [...sources][0]
    }
    return [run as unknown as Inline]
  }
  return [evalObject(state, record, keys, scope, out) as unknown as Inline]
}

/** A list place: each item, control nodes expanded and spliced. */
function evalList(state: State, list: readonly unknown[], pos: Pos, scope: Scope, out: Out): unknown[] {
  const item = ITEM_OF[pos] ?? 'any'
  const into: unknown[] = []
  list.forEach((entry, i) => evalItem(state, entry, item, scope, [...out, i], into))
  return into
}

function evalItem(state: State, entry: unknown, pos: Pos, scope: Scope, out: Out, into: unknown[]): void {
  const kind = isRecord(entry) ? controlKind(entry) : null
  if (kind && isRecord(entry)) {
    for (const key of Object.keys(entry))
      if (!CONTROL_KEYS[kind].includes(key)) throw new RecipeError('unknown-key', `a "${kind}" node has no key "${key}"`, where([...out, key]))
    const body = (key: string, inner: Scope) => {
      const list = entry[key]
      if (list === undefined) return
      if (!Array.isArray(list)) throw new RecipeError('bad-node', `a ${kind} node's ${key} is a list`, where([...out, key]))
      list.forEach((item, i) => evalItem(state, item, pos, inner, [...out, key, i], into))
    }
    switch (kind) {
      case 'if':
        body(test(state, String(entry.if), scope, [...out, 'if']) ? 'then' : 'else', scope)
        return
      case 'for': {
        const runs = passes(state, loopOf(state, String(entry.for), [...out, 'for']), scope, [...out, 'for'])
        if (!runs.length) body('empty', scope)
        for (const pass of runs) body('do', pass)
        return
      }
      case 'component': {
        const inner = bind(state, entry.params as JsonObject | undefined, scope, [...out, 'params'], scope.stack)
        body('blocks', inner)
        return
      }
      case 'code':
        if (pos !== 'block') throw new RecipeError('bad-node', 'a code component stands only where a block does', where(out))
        into.push(...renderCode(state, String(entry.code), entry.params as JsonObject | undefined, entry.variant as string | undefined, scope, out))
        return
      case 'use':
        throw new RecipeError('bad-node', 'a use node in a compiled recipe: compile the recipe first', where(out))
    }
  }
  const value = evalValue(state, entry, pos, scope, out)
  if (value === OMIT) return
  if (pos === 'inline') {
    if (Array.isArray(value)) into.push(...value)
    else if (value !== null) into.push(value)
    return
  }
  into.push(value)
}

/** A component's scope: its parameters' bindings read in the caller's scope, with provenance kept. */
function bind(state: State, params: JsonObject | undefined, caller: Scope, out: Out, stack: readonly string[]): Scope {
  const bound = new Map<string, Bound>()
  for (const [name, binding] of Object.entries(params ?? {})) {
    const at = [...out, name]
    if (typeof binding === 'string') {
      const read = parsed(state, binding, at)
      if (read.kind === 'raw') {
        const found = filtered(state, read.field, caller, at)
        if (found !== MISSING) bound.set(name, found)
        else {
          // Bound to data that is not there: the component reads it as a value the data lacks, at its data path.
          const root = read.field.path[0]
          if (root !== 'params' && root !== 'options' && root !== 'loop')
            bound.set(name, { value: undefined, path: dataPath(state, read.field.path, caller).split('.'), depth: -1 })
        }
        continue
      }
    }
    const value = bindValue(state, binding, caller, at)
    if (value !== OMIT) bound.set(name, { value, path: null, depth: -1 })
  }
  return { params: bound, vars: new Map(), loop: null, stack }
}

/** A binding's value: like an `any` place, but a lone tag with a source stays wrapped so the component can print it. */
function bindValue(state: State, value: Json | undefined, scope: Scope, out: Out): unknown {
  if (typeof value === 'string') {
    const read = parsed(state, value, out)
    // A lone tag whose value is null stays null (a blank to complete); one the data lacks leaves the key out.
    if (read.kind === 'raw') {
      const one = lone(state, read.field, scope, out)
      if (one.value === undefined) return OMIT
      if (one.value === null) return null
    }
    return rawFromString(state, value, scope, out, true)
  }
  if (Array.isArray(value)) {
    const into: unknown[] = []
    value.forEach((entry, i) => {
      const kind = isRecord(entry) ? controlKind(entry) : null
      if (kind) evalItem(state, entry, 'any', scope, [...out, i], into)
      else {
        const one = bindValue(state, entry, scope, [...out, i])
        if (one !== OMIT) into.push(one)
      }
    })
    return into
  }
  if (isRecord(value)) {
    const result: Record<string, unknown> = {}
    for (const [key, member] of Object.entries(value)) {
      const one = bindValue(state, member, scope, [...out, key])
      if (one !== OMIT) result[key] = one
    }
    return result
  }
  return value
}

/* ------------------------------------------------------------------ code components */

function renderCode(state: State, id: string, params: JsonObject | undefined, variant: string | undefined, scope: Scope, out: Out): Block[] {
  const component = state.code?.get(id)
  if (!component) {
    if (state.check) {
      // One the compiled recipe lists is reported once, up front (missingExtensions).
      if (!state.compiled.code.some((need) => need.id === id)) problem(state, { code: 'missing-code', at: where(out), message: `the code component "${id}" was not given` })
      return []
    }
    throw new RecipeError('missing-code', `the code component "${id}" was not given to the render (pass it in code)`, where(out))
  }
  const stack = [...scope.stack, id]
  if (stack.length > state.limits.includeDepth) throw new RecipeError('limit', `code components nest more than ${state.limits.includeDepth} deep (${stack.join(' → ')})`, where(out))
  const inner = bind(state, params, scope, [...out, 'params'], stack)
  const values = Object.fromEntries([...inner.params].map(([name, bound]) => [name, unwrapDeep(bound.value)]))
  const root: Scope = { params: new Map(), vars: new Map(), loop: null, stack }
  let built: unknown
  try {
    built = component.build({
      params: values,
      variant: variant ?? null,
      data: state.data,
      options: state.options,
      value: (path) => {
        const found = lookup({ ...state, check: false }, path.split('.'), root, 'value', out)
        return found === MISSING ? undefined : unwrapDeep(found.value)
      },
      source: (path) => {
        const found = lookup({ ...state, check: false }, path.split('.'), root, 'value', out)
        return found === MISSING ? mapSource(state.sources, path.split('.'))?.source : found.source
      },
    })
  } catch (error) {
    if (error instanceof RecipeError) throw error
    if (state.check) {
      problem(state, { code: 'code-failed', at: where(out), message: `"${id}" failed: ${error instanceof Error ? error.message : String(error)}` })
      return []
    }
    throw new RecipeError('code-failed', `the code component "${id}" failed: ${error instanceof Error ? error.message : String(error)}`, where(out))
  }
  if (!Array.isArray(built)) throw new RecipeError('code-failed', `the code component "${id}" returned something that is not a list of blocks`, where(out))
  const compiled = compileCodeBlocks(built, state.compiled.embedded ?? {}, state.code, stack, state.part, { includeDepth: state.limits.includeDepth })
  return evalList(state, compiled, 'blocks', inner, [...out, 'build']) as Block[]
}

/* ------------------------------------------------------------------ the render */

function newState(compiled: CompiledRecipe, data: unknown, options: RenderOptions, check: boolean): State {
  if (!isRecord(data)) throw new RecipeError('bad-value', 'the data is not an object', '')
  return {
    compiled,
    data,
    options: { ...compiled.options, ...options.options },
    sources: sourceMap(options.sources),
    filters: options.filters ?? {},
    code: options.code,
    format: formatting(options),
    limits: { ...RENDER_LIMITS, ...options.limits },
    check,
    problems: [],
    problemKeys: new Set(),
    gaps: [],
    lastNullSource: undefined,
    part: null,
    section: 0,
    where: 'body',
    loopPasses: 0,
    blocks: 0,
    strings: new Map(),
    conditions: new Map(),
    loops: new Map(),
    missing: compiled.missing ?? 'blank',
    rules: compiled.document.textRules?.length ? compileTextRules(compiled.document.textRules as unknown as TextRule[]) : null,
    autoNamed: new WeakSet(),
  }
}

/** What the render will need from the application that it was not given. */
function missingExtensions(compiled: CompiledRecipe, options: RenderOptions): CheckProblem[] {
  const out: CheckProblem[] = []
  for (const need of compiled.code)
    if (!options.code?.has(need.id)) out.push({ code: 'missing-code', part: need.parts[0] ?? null, at: 'code', message: `the code component "${need.id}" (${need.title}) was not given` })
  for (const name of compiled.filters)
    if (!options.filters || !hasOwn(options.filters, name)) out.push({ code: 'missing-filter', part: null, at: 'filters', message: `the filter "${name}" was not given` })
  if (options.registry)
    for (const kind of compiled.customKinds)
      if (!options.registry.has(kind)) out.push({ code: 'missing-custom-kind', part: null, at: 'customKinds', message: `the custom block kind "${kind}" is not in the registry` })
  return out
}

const EMPTY_SCOPE: Scope = { params: new Map(), vars: new Map(), loop: null, stack: [] }

function renderPart(state: State, part: CompiledPart, scope: Scope, out: Out): Block[] {
  if (part.code !== undefined) return renderCode(state, part.code, part.params, part.variant, scope, out)
  const inner = bind(state, part.params, scope, [...out, 'params'], [part.part.component])
  return evalList(state, part.blocks ?? [], 'blocks', inner, [...out, 'blocks']) as Block[]
}

function sectionBlocks(state: State, blocks: Json[] | undefined, out: Out, part: State['where']): Block[] | undefined {
  if (!blocks) return undefined
  state.part = null
  state.where = part
  return evalList(state, blocks, 'blocks', EMPTY_SCOPE, out) as Block[]
}

const isToCompletePlaceholder = (block: Block) => block.type === 'custom' && block.kind === TO_COMPLETE_KIND

/**
 * Renders a compiled recipe against data. Throws a `RecipeError` for what the application did not give (a code
 * component, a filter, a custom kind its registry lacks), a value that cannot go where it is printed, or a limit.
 */
export function renderRecipe(compiledIn: CompiledRecipe, data: unknown, options: RenderOptions = {}): RenderResult {
  const compiled = readCompiledRecipe(compiledIn)
  const lacking = missingExtensions(compiled, options)
  const first = lacking[0]
  if (first) throw new RecipeError(first.code === 'missing-filter' ? 'unknown-filter' : first.code === 'missing-custom-kind' ? 'missing-custom-kind' : 'missing-code', first.message, first.at)
  const state = newState(compiled, data, options, false)
  const editSource = options.editSource ?? 'edited'
  const edits = options.edits ?? {}

  const metadata = compiled.document.metadata ? (evalValue(state, compiled.document.metadata, 'any', EMPTY_SCOPE, ['document', 'metadata']) as Record<string, unknown>) : undefined
  const theme = compiled.document.theme ? (evalValue(state, compiled.document.theme, 'any', EMPTY_SCOPE, ['document', 'theme']) as Record<string, unknown>) : undefined
  const styles = compiled.document.styles ? (evalValue(state, compiled.document.styles, 'any', EMPTY_SCOPE, ['document', 'styles']) as Record<string, unknown>) : undefined

  const parts: RenderedPart[] = []
  const items: RecipeToCompleteItem[] = []
  const targets: EditTarget[] = []
  const sections: Section[] = compiled.sections.map((section, s) => {
    state.section = s
    const out = ['sections', s]
    state.part = null
    const page = section.page === undefined ? undefined : evalValue(state, section.page, 'any', EMPTY_SCOPE, [...out, 'page'])
    const pageNumbers = section.pageNumbers === undefined ? undefined : evalValue(state, section.pageNumbers, 'any', EMPTY_SCOPE, [...out, 'pageNumbers'])
    const header = sectionBlocks(state, section.header, [...out, 'header'], 'header')
    const footer = sectionBlocks(state, section.footer, [...out, 'footer'], 'footer')
    const firstHeader = sectionBlocks(state, section.firstPage?.header, [...out, 'firstPage', 'header'], 'firstHeader')
    const firstFooter = sectionBlocks(state, section.firstPage?.footer, [...out, 'firstPage', 'footer'], 'firstFooter')
    const headerGaps = state.gaps.splice(0)
    const body: Block[] = []
    let lastHeading: string | undefined
    section.parts.forEach((part, p) => {
      const pout = [...out, 'parts', p]
      state.part = part.part.id
      state.where = 'body'
      state.missing = part.missing ?? compiled.missing ?? 'blank'
      const runs = part.for !== undefined ? passes(state, loopOf(state, part.for, [...pout, 'for']), EMPTY_SCOPE, [...pout, 'for']) : [EMPTY_SCOPE]
      let produced: Block[] = []
      let repeats = 0
      for (const run of runs) {
        if (part.if !== undefined && !test(state, part.if, run, [...pout, 'if'])) continue
        repeats += 1
        produced.push(...renderPart(state, part, run, pout))
      }
      // A pinned part's blocks go in one block pinned to the foot of the page.
      if (part.pin !== undefined && produced.length) {
        const spaceBelow = typeof part.pin === 'object' ? part.pin.spaceBelow : undefined
        produced = [{ type: 'pinned', blocks: produced, ...(spaceBelow !== undefined ? { spaceBelow } : {}) }]
      }
      if (part.pageBreakBefore && produced[0]) produced[0] = { ...produced[0], pageBreakBefore: true }
      state.missing = compiled.missing ?? 'blank'
      const edited = applyEdits({ id: part.part.id, title: part.part.title }, produced, edits, editSource)
      produced = edited.blocks
      targets.push(...edited.targets)
      const start = body.length
      body.push(...produced)
      // The part's own to-complete items, read by the model's rules, under the heading before it.
      const context: Block[] = lastHeading ? [{ type: 'heading', level: 1, content: lastHeading }] : []
      const own = toCompleteOf(options, page, theme, [...context, ...produced.filter((b) => !isToCompletePlaceholder(b))], s, part.part.id)
      const gaps = state.gaps.splice(0)
      items.push(...own, ...gaps)
      for (const block of produced) if (block.type === 'heading') lastHeading = plainHeading(block)
      parts.push({
        id: part.part.id,
        component: part.part.component,
        title: part.part.title,
        section: s,
        start,
        end: body.length,
        blocks: produced.length,
        toComplete: own.length + gaps.length,
        repeats,
        ...(part.part.meta ? { meta: part.part.meta } : {}),
        ...(part.part.componentMeta ? { componentMeta: part.part.componentMeta } : {}),
      })
    })
    const built: Record<string, unknown> = { body }
    if (page !== undefined) built.page = page
    if (pageNumbers !== undefined) built.pageNumbers = pageNumbers
    if (header) built.header = header
    if (footer) built.footer = footer
    if (firstHeader || firstFooter) built.firstPage = { ...(firstHeader ? { header: firstHeader } : {}), ...(firstFooter ? { footer: firstFooter } : {}) }
    const result = built as unknown as Section
    const running = toCompleteOf(options, page, theme, [], s, null, result)
    items.push(...running, ...headerGaps)
    return result
  })

  // The to-complete list printed in the document (the `to-complete` component) lists everything else.
  for (const section of sections)
    (section as { body: Block[] }).body = section.body.map((block) => (isToCompletePlaceholder(block) ? toCompleteBlock(block, items) : block))

  const document: Document = { sections, ...(compiled.document.footnotes && { footnotes: compiled.document.footnotes }) }
  if (metadata && Object.keys(metadata).length) (document as { metadata?: unknown }).metadata = metadata
  if (theme && Object.keys(theme).length) (document as { theme?: unknown }).theme = theme
  if (styles && Object.keys(styles).length) (document as { styles?: unknown }).styles = styles
  if (compiled.document.textRules?.length) (document as { textRules?: unknown }).textRules = compiled.document.textRules

  // An edit counts as applied only where its target took it (a tick must be one of the question's options).
  const keys = new Set(targets.filter((target) => target.edited !== undefined).map((target) => target.key))
  const partIds = new Set(compiled.sections.flatMap((section) => section.parts.map((part) => part.part.id)))
  const orphans: string[] = []
  const otherParts: string[] = []
  for (const key of Object.keys(edits)) {
    if (keys.has(key)) continue
    if (partIds.has(key.slice(0, key.indexOf('|')))) orphans.push(key)
    else otherParts.push(key)
  }
  return { document, parts, toComplete: items, edits: { targets, orphans, otherParts } }
}

const plainHeading = (block: Extract<Block, { type: 'heading' }>) => plainText(block.content).trim() || undefined

function toCompleteOf(
  options: RenderOptions,
  page: unknown,
  theme: Record<string, unknown> | undefined,
  body: Block[],
  section: number,
  partId: string | null,
  running?: Section,
): RecipeToCompleteItem[] {
  const mini: Section = running
    ? { ...running, body: [] }
    : ({ body, ...(page !== undefined ? { page } : {}) } as Section)
  const doc: Document = { sections: [mini], ...(theme ? { theme } : {}) } as Document
  const found = modelToComplete(doc, options.registry ? { registry: options.registry } : {})
  return found.map((item) => ({ ...item, section, partId }))
}

/** The printed to-complete list: one list block (or a line saying nothing is left), in place of its placeholder. */
function toCompleteBlock(block: Block, items: readonly RecipeToCompleteItem[]): Block {
  const data = (block.type === 'custom' && isRecord(block.data) ? block.data : {}) as Record<string, unknown>
  const none = typeof data.none === 'string' ? data.none : 'Nothing is left to complete.'
  const base = { ...(block.pageBreakBefore ? { pageBreakBefore: true } : {}), ...(block.id ? { id: block.id } : {}) }
  if (!items.length) return { type: 'paragraph', content: none, ...base }
  const lines = items.map((item) => {
    const needs = item.missing?.length ? ` (needs ${item.missing.join(', ')})` : ''
    return `${item.heading ? `${item.heading}: ` : ''}${item.what}${needs}`
  })
  return { type: 'list', items: lines, ...base }
}

/**
 * What the data lacks for a compiled recipe, without making the document: every data path a tag, condition or loop
 * reads that the data does not have (in every pass of every loop, and the branch each condition takes), and every
 * code component, filter or custom kind the application did not give. Empty when the data is enough.
 */
export function checkRecipe(compiledIn: CompiledRecipe, data: unknown, options: RenderOptions = {}): CheckProblem[] {
  const compiled = readCompiledRecipe(compiledIn)
  const state = newState(compiled, data, options, true)
  state.problems.push(...missingExtensions(compiled, options))
  const guard = (fn: () => void) => {
    try {
      fn()
    } catch (error) {
      if (!(error instanceof RecipeError) || !['bad-value', 'not-text', 'code-failed'].includes(error.code)) throw error
      problem(state, { code: error.code as CheckProblem['code'], at: error.path, message: error.message })
    }
  }
  if (compiled.document.metadata) guard(() => evalValue(state, compiled.document.metadata, 'any', EMPTY_SCOPE, ['document', 'metadata']))
  compiled.sections.forEach((section, s) => {
    const out = ['sections', s]
    state.part = null
    for (const [key, blocks] of [['header', section.header], ['footer', section.footer]] as const)
      if (blocks) guard(() => evalList(state, blocks, 'blocks', EMPTY_SCOPE, [...out, key]))
    section.parts.forEach((part, p) => {
      const pout = [...out, 'parts', p]
      state.part = part.part.id
      guard(() => {
        const runs = part.for !== undefined ? passes(state, loopOf(state, part.for, [...pout, 'for']), EMPTY_SCOPE, [...pout, 'for']) : [EMPTY_SCOPE]
        for (const run of runs) {
          if (part.if !== undefined && !test(state, part.if, run, [...pout, 'if'])) continue
          guard(() => renderPart(state, part, run, pout))
        }
      })
    })
  })
  return state.problems
}
