/**
 * Compiling a recipe: its `extends` resolved, every `use` inlined (a code component kept as a marked reference),
 * variants and parameter defaults applied, every tag read and every key checked against the block model, into one
 * self-contained JSON object (`CompiledRecipe`). Deterministic: the same inputs give the same JSON text (keys in
 * sorted order). Pure: no DOM, no network.
 */
import type { CodeComponent, CodeComponents } from './code'
import { jsonPath, RecipeError } from './errors'
import { canonical, cloneJson, contentHash, hasOwn, isRecord } from './json'
import { readLibrary, readRecipe } from './read'
import { BLOCK_KEYS, CONTROL_KEYS, controlKind, ITEM_OF, isListPos, keysOf, type Pos } from './schema'
import { BUILT_IN_FILTERS, customFilterNames, readCondition, readLoop, readString, ruleIdsOf, TagSyntaxError, walkSegs, type CompiledFilter, type Loop } from './tags'
import { compileTextRules, TextRuleError } from '../model/textRules'
import type { TextRule } from '../model/types'
import type { TemplateExpression } from '../template/templateTags'
import {
  RECIPE_FORMAT_VERSION,
  type CodeNeed,
  type CompiledPart,
  type CompiledRecipe,
  type CompiledSection,
  type Component,
  type Json,
  type JsonObject,
  type Library,
  type NeedUse,
  type PartRef,
  type Recipe,
  type RecipeChange,
  type RecipeDocument,
  type RecipeNeed,
  type RecipeSection,
} from './types'

/** Bounds on a compile. Each can be lowered or raised by the caller (`limits`). */
export const COMPILE_LIMITS = {
  /** How deep `use` may nest (a component in a component …), code components and their `uses` included. */
  includeDepth: 16,
  /** How long an `extends` chain may be. */
  extendsDepth: 16,
  /** JSON nodes in the compiled recipe's templates: every value walked counts one. */
  nodes: 200_000,
}
export type CompileLimits = typeof COMPILE_LIMITS

/** The custom block kind the render fills with the document's to-complete list (`to-complete` component). */
export const TO_COMPLETE_KIND = 'recipe.toComplete'

export interface CompileOptions {
  /** Code components the recipe may use. A JSON component with one's id is an error (`component-conflict`). */
  code?: CodeComponents
  /** The application's filter names: given, a tag using any other is an error at compile, not at render. */
  filters?: readonly string[]
  limits?: Partial<CompileLimits>
}

/** Where a node came from, for errors: a component or recipe, and the steps into it. */
interface Loc {
  root: string
  steps: ReadonlyArray<string | number>
}
const at = (loc: Loc, ...more: Array<string | number>): Loc => ({ root: loc.root, steps: [...loc.steps, ...more] })
const where = (loc: Loc) => `${loc.root}${loc.steps.length ? `:${jsonPath(loc.steps)}` : ''}`
type Out = ReadonlyArray<string | number>

/** What a name means for listing needs: the data path it stands for (`tables[]`), or null when it is not data. */
interface Sym {
  params: ReadonlyMap<string, string | null>
  vars: ReadonlyMap<string, string | null>
  inLoop: boolean
  part: string | null
  stack: readonly string[]
}

interface Ctx {
  components: ReadonlyMap<string, Component>
  code: CodeComponents | undefined
  filtersKnown: ReadonlySet<string> | null
  limits: CompileLimits
  nodes: number
  needs: Map<string, RecipeNeed>
  customKinds: Set<string>
  filters: Set<string>
  codeUsed: Map<string, CodeNeed & { component: CodeComponent }>
  componentsUsed: Map<string, Component>
  embedded: Map<string, Component>
  /** Walking a code component's `uses`: every JSON component reached is embedded too. */
  embedding: boolean
  /** Walking blocks a code component returned (at render): pictures as bytes pass, `use` reads `embedded`. */
  fromCode: boolean
  /** The text rule ids the recipe's document defines, for the `rules` filter; null when not checked. */
  ruleIds: ReadonlySet<string> | null
}

const fail = (code: ConstructorParameters<typeof RecipeError>[0], message: string, loc: Loc): never => {
  throw new RecipeError(code, message, where(loc))
}

/* ------------------------------------------------------------------ extends and changes */

/** A part's id in its recipe. */
export const partId = (part: PartRef): string => part.id ?? part.use

function findPart(sections: RecipeSection[], id: string): { s: number; p: number } | null {
  for (let s = 0; s < sections.length; s += 1) {
    const p = sections[s]?.parts.findIndex((part) => partId(part) === id) ?? -1
    if (p >= 0) return { s, p }
  }
  return null
}

function findSection(sections: RecipeSection[], ref: string | number | undefined): number {
  if (ref === undefined) return sections.length - 1
  if (typeof ref === 'number') return Number.isInteger(ref) && ref >= 0 && ref < sections.length ? ref : -1
  return sections.findIndex((section) => section.id === ref)
}

/**
 * Sections with `changes` applied, in order: `add`, `remove`, `replace`, `move`, `set`, `setSection`, `addSection`.
 * A change that names a part or section the sections do not have is an error (`unknown-part`). Returns new sections.
 */
export function applyChanges(sectionsIn: readonly RecipeSection[], changes: readonly RecipeChange[], root = 'recipe'): RecipeSection[] {
  const sections = cloneJson(sectionsIn as unknown as Json) as unknown as RecipeSection[]
  changes.forEach((change, i) => {
    const loc: Loc = { root, steps: ['changes', i] }
    const need = (id: string) => findPart(sections, id) ?? fail('unknown-part', `no part "${id}" to change`, loc)
    const insert = (part: PartRef, anchor: { before?: string; after?: string; section?: string | number }) => {
      if (anchor.before !== undefined || anchor.after !== undefined) {
        const found = need((anchor.before ?? anchor.after) as string)
        sections[found.s]?.parts.splice(found.p + (anchor.after !== undefined ? 1 : 0), 0, part)
        return
      }
      const s = findSection(sections, anchor.section)
      if (s < 0) fail('unknown-part', `no section ${JSON.stringify(anchor.section)}`, loc)
      sections[s]?.parts.push(part)
    }
    if ('add' in change) insert(cloneJson(change.add as unknown as Json) as unknown as PartRef, change)
    else if ('remove' in change) {
      const found = need(change.remove)
      sections[found.s]?.parts.splice(found.p, 1)
    } else if ('replace' in change) {
      const found = need(change.replace)
      const by = cloneJson(change.by as unknown as Json) as unknown as PartRef
      // The part keeps its id (and so its hand edits) unless the replacement names its own.
      if (by.id === undefined && by.use !== change.replace) by.id = change.replace
      sections[found.s]?.parts.splice(found.p, 1, by)
    } else if ('move' in change) {
      const found = need(change.move)
      const [part] = sections[found.s]?.parts.splice(found.p, 1) ?? []
      if (part) insert(part, change)
    } else if ('set' in change) {
      const found = need(change.set)
      const part = sections[found.s]?.parts[found.p]
      if (!part) return
      for (const key of ['variant', 'if', 'for', 'title', 'pageBreakBefore', 'missing', 'pin', 'meta'] as const) {
        const value = change[key]
        if (value === undefined) continue
        if (value === null) delete part[key]
        else (part as unknown as Record<string, unknown>)[key] = cloneJson(value as Json)
      }
      if (change.with) {
        const next: JsonObject = { ...(part.with ?? {}) }
        for (const [name, value] of Object.entries(change.with)) {
          if (value === null) delete next[name]
          else next[name] = cloneJson(value)
        }
        if (Object.keys(next).length) part.with = next
        else delete part.with
      }
    } else if ('setSection' in change) {
      const s = findSection(sections, change.setSection)
      const section = sections[s]
      if (!section) return fail('unknown-part', `no section ${JSON.stringify(change.setSection)}`, loc)
      for (const key of ['page', 'header', 'footer', 'firstPage', 'pageNumbers'] as const) {
        const value = change[key]
        if (value === undefined) continue
        if (value === null) delete section[key]
        else (section as unknown as Record<string, unknown>)[key] = cloneJson(value as Json)
      }
    } else if ('addSection' in change) {
      const added = cloneJson(change.addSection as unknown as Json) as unknown as RecipeSection
      if (change.after === undefined) sections.push(added)
      else {
        const s = findSection(sections, change.after)
        if (s < 0) fail('unknown-part', `no section ${JSON.stringify(change.after)}`, loc)
        sections.splice(s + 1, 0, added)
      }
    }
  })
  return sections
}

const mergeDeep = (base: Json | undefined, over: Json | undefined): Json | undefined => {
  if (over === undefined) return base
  if (!isRecord(base) || !isRecord(over)) return cloneJson(over as Json)
  const out: JsonObject = { ...(base as JsonObject) }
  for (const [key, value] of Object.entries(over as JsonObject)) out[key] = mergeDeep(out[key], value) as Json
  return out
}

/** A recipe with its `extends` chain resolved: its own sections, document and options, and the chain, nearest first. */
export function resolveRecipe(
  recipe: Recipe,
  recipes: ReadonlyMap<string, Recipe>,
  limits: Pick<CompileLimits, 'extendsDepth'> = COMPILE_LIMITS,
): { sections: RecipeSection[]; document: RecipeDocument; options: JsonObject; chain: Recipe[]; missing?: Recipe['missing'] } {
  const chain: Recipe[] = [recipe]
  const seen = new Set([recipe.id])
  for (let current = recipe; current.extends !== undefined; ) {
    const loc: Loc = { root: `recipe/${current.id}`, steps: ['extends'] }
    if (seen.has(current.extends)) fail('cycle', `recipe "${current.id}" extends "${current.extends}", which leads back to it`, loc)
    if (chain.length > limits.extendsDepth) fail('limit', `an extends chain longer than ${limits.extendsDepth}`, loc)
    const base = recipes.get(current.extends) ?? fail('unknown-recipe', `no recipe "${current.extends}" to extend`, loc)
    seen.add(base.id)
    chain.push(base)
    current = base
  }
  let sections: RecipeSection[] = []
  let document: Json | undefined = {}
  let options: JsonObject = {}
  for (const r of [...chain].reverse()) {
    if (r.sections) sections = cloneJson(r.sections as unknown as Json) as unknown as RecipeSection[]
    document = mergeDeep(document, r.document as Json | undefined)
    options = { ...options, ...cloneJson((r.options ?? {}) as Json) as JsonObject }
    if (r.changes?.length) sections = applyChanges(sections, r.changes, `recipe/${r.id}`)
  }
  // The nearest recipe that states a missing-value rule gives it.
  const missing = chain.find((r) => r.missing !== undefined)?.missing
  return { sections, document: document as RecipeDocument, options, chain, ...(missing ? { missing } : {}) }
}

/* ------------------------------------------------------------------ needs */

function expressionPaths(expression: TemplateExpression, out: string[][] = []): string[][] {
  switch (expression.kind) {
    case 'path':
      out.push(expression.path)
      break
    case 'literal':
      break
    case 'filtered':
      expressionPaths(expression.operand, out)
      break
    case 'not':
      expressionPaths(expression.operand, out)
      break
    case 'and':
    case 'or':
      for (const operand of expression.operands) expressionPaths(operand, out)
      break
    case 'compare':
      expressionPaths(expression.left, out)
      expressionPaths(expression.right, out)
      break
  }
  return out
}

/** The data path a tag's path stands for, or null when it reads parameters bound to no data, options or `loop`. */
function symbolic(path: readonly string[], sym: Sym, local: ReadonlyMap<string, string | null> = new Map()): string | null {
  const [first, ...rest] = path
  if (first === undefined || first === 'options' || first === 'loop') return null
  let base: string | null
  if (local.has(first)) base = local.get(first) ?? null
  else if (sym.vars.has(first)) base = sym.vars.get(first) ?? null
  else if (first === 'params') {
    const [name, ...after] = rest
    if (name === undefined) return null
    base = sym.params.get(name) ?? null
    return base === null ? null : [base, ...after].join('.')
  } else return path.join('.')
  return base === null ? null : [base, ...rest].join('.')
}

function addNeed(ctx: Ctx, path: string | null, use: NeedUse, sym: Sym, out: Out, inLoop: boolean, declared = false) {
  if (path === null) return
  let need = ctx.needs.get(path)
  if (!need) ctx.needs.set(path, (need = { path, uses: [], inLoop: false, places: [] }))
  if (!need.uses.includes(use)) need.uses.push(use)
  const loop = sym.inLoop || inLoop
  need.inLoop ||= loop
  need.places.push({ part: sym.part, at: jsonPath(out), use, inLoop: loop, ...(declared ? { declared: true as const } : {}) })
}

function noteFilters(ctx: Ctx, filters: readonly CompiledFilter[], loc: Loc) {
  for (const id of ruleIdsOf(filters))
    if (ctx.ruleIds && !ctx.ruleIds.has(id))
      fail('bad-tag', `the rules filter names "${id}", which is not a text rule of the recipe's document (${[...ctx.ruleIds].join(', ') || 'it has none with an id'})`, loc)
  for (const name of customFilterNames(filters)) {
    if (ctx.filtersKnown && !ctx.filtersKnown.has(name))
      fail('unknown-filter', `the filter "${name}" is not format, upper, lower or one of the application's (${[...ctx.filtersKnown].join(', ') || 'none given'})`, loc)
    ctx.filters.add(name)
  }
}

/** Reads a string's tags, recording what they read. Returns the parsed string (null for plain text). */
function readStr(text: string, ctx: Ctx, sym: Sym, loc: Loc, out: Out): ReturnType<typeof readString> {
  let parsed: ReturnType<typeof readString>
  try {
    parsed = readString(text)
  } catch (error) {
    if (error instanceof TagSyntaxError) return fail('bad-tag', error.message, loc)
    throw error
  }
  if (parsed.kind === 'plain') return parsed
  const segs = parsed.kind === 'raw' ? [parsed.field] : parsed.segs
  const local = new Map<string, string | null>()
  for (const item of walkSegs(segs)) {
    const inLoop = item.loops.length > 0
    for (const loop of item.loops) local.set(loop.name, symbolic(loop.list, sym, local) === null ? null : `${symbolic(loop.list, sym, local)}[]`)
    if (item.kind === 'field') {
      noteFilters(ctx, item.field.filters, loc)
      addNeed(ctx, symbolic(item.field.path, sym, local), 'value', sym, out, inLoop)
    } else if (item.kind === 'condition') {
      for (const chain of item.condition.chains) noteFilters(ctx, chain, loc)
      for (const path of expressionPaths(item.condition.expression)) addNeed(ctx, symbolic(path, sym, local), 'condition', sym, out, inLoop)
    } else addNeed(ctx, symbolic(item.loop.list, sym, local), 'list', sym, out, inLoop)
  }
  return parsed
}

function condition(text: string, ctx: Ctx, sym: Sym, loc: Loc, out: Out) {
  try {
    const read = readCondition(text)
    for (const chain of read.chains) noteFilters(ctx, chain, loc)
    for (const path of expressionPaths(read.expression)) addNeed(ctx, symbolic(path, sym), 'condition', sym, out, false)
  } catch (error) {
    if (error instanceof TagSyntaxError) fail('bad-tag', error.message, loc)
    throw error
  }
}

function loop(text: string, ctx: Ctx, sym: Sym, loc: Loc, out: Out): { loop: Loop; sym: Sym } {
  let read: Loop
  try {
    read = readLoop(text)
  } catch (error) {
    if (error instanceof TagSyntaxError) return fail('bad-tag', error.message, loc)
    throw error
  }
  const list = symbolic(read.list, sym)
  addNeed(ctx, list, 'list', sym, out, false)
  const vars = new Map(sym.vars)
  vars.set(read.name, list === null ? null : `${list}[]`)
  return { loop: read, sym: { ...sym, vars, inLoop: true } }
}

/* ------------------------------------------------------------------ walking templates */

function count(ctx: Ctx, loc: Loc) {
  ctx.nodes += 1
  if (ctx.nodes > ctx.limits.nodes) fail('limit', `the compiled templates hold more than ${ctx.limits.nodes} nodes`, loc)
}

const isBytes = (value: unknown) => ArrayBuffer.isView(value) || value instanceof ArrayBuffer

/** Walks a list place: each item, control nodes included. */
function walkList(value: unknown, pos: Pos, ctx: Ctx, sym: Sym, loc: Loc, out: Out): Json[] {
  if (!Array.isArray(value)) return fail('bad-node', `a ${pos} place takes a list`, loc)
  const item = ITEM_OF[pos] ?? 'any'
  return value.map((entry, i) => walkItem(entry, item, ctx, sym, at(loc, i), [...out, i]))
}

function walkItem(value: unknown, pos: Pos, ctx: Ctx, sym: Sym, loc: Loc, out: Out): Json {
  if (!isRecord(value)) return walkValue(value, pos, ctx, sym, loc, out)
  const kind = controlKind(value)
  if (!kind) return walkValue(value, pos, ctx, sym, loc, out)
  count(ctx, loc)
  for (const key of Object.keys(value))
    if (!CONTROL_KEYS[kind].includes(key)) fail('unknown-key', `a "${kind}" node has no key "${key}" (it takes ${CONTROL_KEYS[kind].join(', ')})`, at(loc, key))
  const listOf = (key: string, inner: Sym) =>
    value[key] === undefined ? undefined : walkBodies(value[key], pos, ctx, inner, at(loc, key), [...out, key])
  switch (kind) {
    case 'if': {
      if (typeof value.if !== 'string') return fail('bad-node', 'an if node\'s condition is not a string', at(loc, 'if'))
      condition(value.if, ctx, sym, at(loc, 'if'), [...out, 'if'])
      const then = listOf('then', sym) ?? fail('bad-node', 'an if node has no then', loc)
      const otherwise = listOf('else', sym)
      return { if: value.if, then, ...(otherwise ? { else: otherwise } : {}) }
    }
    case 'for': {
      if (typeof value.for !== 'string') return fail('bad-node', 'a for node\'s loop is not a string', at(loc, 'for'))
      const read = loop(value.for, ctx, sym, at(loc, 'for'), [...out, 'for'])
      const body = listOf('do', read.sym) ?? fail('bad-node', 'a for node has no do', loc)
      const empty = listOf('empty', sym)
      return { for: value.for, do: body, ...(empty ? { empty } : {}) }
    }
    case 'use':
      return useNode(value, pos, ctx, sym, loc, out)
    default:
      return fail('bad-node', `a "${kind}" node is written only by the compile; write { "use": … }`, loc)
  }
}

function walkBodies(value: unknown, itemPos: Pos, ctx: Ctx, sym: Sym, loc: Loc, out: Out): Json[] {
  if (!Array.isArray(value)) return fail('bad-node', 'a control node\'s branch is a list', loc)
  return value.map((entry, i) => walkItem(entry, itemPos, ctx, sym, at(loc, i), [...out, i]))
}

function walkKeys(value: Record<string, unknown>, keys: Record<string, Pos>, what: string, ctx: Ctx, sym: Sym, loc: Loc, out: Out, skip: readonly string[] = []): JsonObject {
  const result: JsonObject = {}
  for (const [key, member] of Object.entries(value)) {
    if (skip.includes(key)) {
      result[key] = member as Json
      continue
    }
    const pos = keys[key] ?? fail('unknown-key', `${what} has no key "${key}" (it takes ${Object.keys(keys).join(', ') || 'none'})`, at(loc, key))
    if (member === undefined) continue
    result[key] = walkValue(member, pos, ctx, sym, at(loc, key), [...out, key])
  }
  return result
}

function walkValue(value: unknown, pos: Pos, ctx: Ctx, sym: Sym, loc: Loc, out: Out): Json {
  count(ctx, loc)
  if (isListPos(pos) && Array.isArray(value)) return walkList(value, pos, ctx, sym, loc, out)
  if (typeof value === 'string') {
    if (pos === 'literal') {
      if (value.includes('{{')) fail('bad-tag', 'a sentence\'s template takes {slots}, not tags: put the tags in its values', loc)
      return value
    }
    readStr(value, ctx, sym, loc, out)
    return value
  }
  if (value === null) {
    if (!['richNull', 'richOpt', 'cell', 'image', 'any', 'slots', 'scalar', 'text', 'caption'].includes(pos))
      return fail('bad-node', `null does not stand in a ${pos} place`, loc)
    return null
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    if (pos !== 'scalar' && pos !== 'any' && pos !== 'slots' && pos !== 'text')
      return fail('bad-node', `${String(value)} does not stand in a ${pos} place`, loc)
    return value
  }
  if (ctx.fromCode && isBytes(value)) return value as unknown as Json
  if (Array.isArray(value)) {
    if (pos === 'any') return value.map((entry, i) => walkItem(entry, 'any', ctx, sym, at(loc, i), [...out, i]))
    if (pos === 'listItem') return walkList(value, 'rich', ctx, sym, loc, out)
    if (pos === 'cell') return walkList(value, 'rich', ctx, sym, loc, out)
    return fail('bad-node', `a list does not stand in a ${pos} place`, loc)
  }
  if (!isRecord(value)) return fail('bad-node', `${typeof value} is not JSON`, loc)
  if (controlKind(value)) return fail('bad-node', 'a control node stands only in a list', loc)
  switch (pos) {
    case 'any':
    case 'slots': {
      const result: JsonObject = {}
      for (const [key, member] of Object.entries(value))
        result[key] = walkValue(member, 'any', ctx, sym, at(loc, key), [...out, key])
      return result
    }
    case 'image':
      if (ctx.fromCode) return value as unknown as Json
      return fail('bad-node', 'a picture comes from the data ("{{ path }}"); JSON cannot hold its bytes', loc)
    case 'block': {
      if (typeof value.type !== 'string' || !keysOf('block', value))
        return fail('bad-node', `a block's type is one of ${Object.keys(BLOCK_KEYS).join(', ')}, written as such`, at(loc, 'type'))
      const keys = keysOf('block', value) ?? {}
      if (value.type === 'custom' && typeof value.kind === 'string' && !value.kind.includes('{{') && value.kind !== TO_COMPLETE_KIND)
        ctx.customKinds.add(value.kind)
      return { type: value.type, ...walkKeys(value, keys, `a ${value.type} block`, ctx, sym, loc, out, ['type']) } as JsonObject
    }
    case 'inline':
    case 'rich':
    case 'richNull':
    case 'richOpt':
    case 'cell':
    case 'listItem': {
      if (typeof value.type === 'string') {
        const keys = keysOf('inline', value) ?? fail('bad-node', `an inline run's type is one of text, field, break, tab, blank, image, sentence, checkbox, footnote`, at(loc, 'type'))
        return { type: value.type, ...walkKeys(value, keys, `a ${value.type} run`, ctx, sym, loc, out, ['type']) } as JsonObject
      }
      if (pos === 'cell') return walkKeys(value, keysOf('cell', value) ?? {}, 'a table cell', ctx, sym, loc, out)
      if (pos === 'listItem') return walkKeys(value, keysOf('listItem', value) ?? {}, 'a list item', ctx, sym, loc, out)
      return fail('bad-node', 'an inline run has a type', loc)
    }
    default: {
      const keys = keysOf(pos, value) ?? fail('bad-node', `an object does not stand in a ${pos} place`, loc)
      return walkKeys(value, keys, `a ${pos === 'fieldRow' ? 'field row' : pos === 'gridItem' ? 'picture' : pos}`, ctx, sym, loc, out)
    }
  }
}

/* ------------------------------------------------------------------ use */

interface Resolved {
  kind: 'json' | 'code'
  id: string
  title: string
  variant?: string
  params: JsonObject
  blocks?: Json[]
}

function lookupComponent(id: string, ctx: Ctx, loc: Loc): { json?: Component; code?: CodeComponent } {
  const json = ctx.components.get(id)
  const code = ctx.code?.get(id)
  if (json && code) fail('component-conflict', `"${id}" is both a JSON component and a code component`, loc)
  if (!json && !code) {
    const known = [...ctx.components.keys(), ...(ctx.code?.ids ?? [])].sort()
    fail('unknown-component', `no component "${id}"${known.length ? ` (the library has ${known.slice(0, 12).join(', ')}${known.length > 12 ? ', …' : ''})` : ''}`, loc)
  }
  return { ...(json ? { json } : {}), ...(code ? { code } : {}) }
}

/** The parameter bindings: defaults < variant < caller, checked against what the component declares. */
function bindings(
  id: string,
  declared: Record<string, { default?: Json; required?: boolean }>,
  variantWith: JsonObject | undefined,
  given: JsonObject | undefined,
  loc: Loc,
): JsonObject {
  const result: JsonObject = {}
  for (const [source, values] of [['variant', variantWith], ['with', given]] as const)
    for (const name of Object.keys(values ?? {}))
      if (!hasOwn(declared, name))
        fail('unknown-param', `"${id}" takes no parameter "${name}"${Object.keys(declared).length ? ` (it takes ${Object.keys(declared).join(', ')})` : ' (it takes none)'}`, at(loc, source === 'with' ? 'with' : 'variant', name))
  for (const [name, spec] of Object.entries(declared)) {
    const value = given && hasOwn(given, name) ? given[name] : variantWith && hasOwn(variantWith, name) ? variantWith[name] : spec.default
    if (value === undefined) {
      if (spec.required) fail('missing-param', `"${id}" needs the parameter "${name}"`, loc)
      continue
    }
    result[name] = cloneJson(value as Json)
  }
  return result
}

function resolveUse(
  id: string,
  variant: string | undefined,
  given: JsonObject | undefined,
  pos: Pos,
  ctx: Ctx,
  sym: Sym,
  loc: Loc,
  out: Out,
): Resolved {
  const found = lookupComponent(id, ctx, loc)
  if (sym.stack.includes(id)) fail('cycle', `"${id}" uses itself through ${[...sym.stack.slice(sym.stack.indexOf(id)), id].join(' → ')}`, loc)
  if (sym.stack.length >= ctx.limits.includeDepth) fail('limit', `components nest more than ${ctx.limits.includeDepth} deep (${[...sym.stack, id].join(' → ')})`, loc)
  const walkParams = (params: JsonObject) => {
    const result: JsonObject = {}
    for (const [name, value] of Object.entries(params))
      result[name] = walkValue(value, 'any', ctx, sym, at(loc, 'with', name), [...out, 'params', name])
    return result
  }
  if (found.code) {
    const code = found.code
    if (pos !== 'block') fail('bad-node', `the code component "${id}" builds blocks, and stands only where a block does`, loc)
    if (variant !== undefined && !hasOwn(code.variants ?? {}, variant))
      fail('unknown-variant', `"${id}" has no variant "${variant}"`, at(loc, 'variant'))
    const params = walkParams(bindings(id, code.params ?? {}, variant !== undefined ? code.variants?.[variant]?.with : undefined, given, loc))
    for (const path of code.needs ?? []) addNeed(ctx, path, 'value', sym, out, false, true)
    const used = ctx.codeUsed.get(id) ?? { id, title: code.title, ...(code.version ? { version: code.version } : {}), parts: [], component: code }
    if (!used.parts.includes(sym.part)) used.parts.push(sym.part)
    ctx.codeUsed.set(id, used)
    // The JSON components its blocks may use travel with the compiled recipe; each is checked here.
    for (const usedId of code.uses ?? []) {
      const component = ctx.components.get(usedId) ?? fail('unknown-component', `the code component "${id}" declares it uses "${usedId}", which the library does not have`, loc)
      if (ctx.embedded.has(usedId)) continue
      ctx.embedded.set(usedId, component)
      const inner: Sym = { params: new Map(), vars: new Map(), inLoop: sym.inLoop, part: sym.part, stack: [...sym.stack, id] }
      const was = ctx.embedding
      ctx.embedding = true
      resolveUse(usedId, undefined, undefined, 'block', ctx, inner, { root: `component/${usedId}`, steps: [] }, [...out, 'embedded', usedId])
      ctx.embedding = was
    }
    return { kind: 'code', id, title: code.title, ...(variant !== undefined ? { variant } : {}), params }
  }
  const component = found.json as Component
  const chosen = variant === undefined ? undefined : component.variants?.[variant] ?? fail('unknown-variant', `"${id}" has no variant "${variant}"${component.variants ? ` (it has ${Object.keys(component.variants).join(', ')})` : ''}`, at(loc, 'variant'))
  // An embedded component (reached from a code component) is used at render with whatever its caller gives, so its
  // required parameters are not asked for here.
  const declared = ctx.embedding ? Object.fromEntries(Object.entries(component.params ?? {}).map(([name, spec]) => [name, { ...spec, required: false }])) : component.params ?? {}
  const params = walkParams(bindings(id, declared, chosen?.with, given, loc))
  const inner: Sym = {
    params: new Map(Object.entries(params).map(([name, value]) => [name, bindingSym(value, sym)])),
    vars: new Map(),
    inLoop: sym.inLoop,
    part: sym.part,
    stack: [...sym.stack, id],
  }
  const body = chosen?.blocks ?? component.blocks
  const bodyLoc: Loc = { root: `component/${id}`, steps: chosen?.blocks ? ['variants', variant as string, 'blocks'] : ['blocks'] }
  const listPos: Pos = pos === 'block' ? 'blocks' : pos === 'row' ? 'rows' : pos === 'cell' ? 'cells' : pos === 'fieldRow' ? 'fieldRows' : pos === 'listItem' ? 'listItems' : pos === 'gridItem' ? 'gridItems' : pos === 'inline' ? 'inlines' : 'any'
  const blocks = walkList(body, listPos, ctx, inner, bodyLoc, [...out, 'blocks'])
  ctx.componentsUsed.set(id, component)
  if (ctx.embedding) ctx.embedded.set(id, component)
  return { kind: 'json', id, title: component.title, ...(variant !== undefined ? { variant } : {}), params, blocks }
}

/** The data path a parameter binding stands for: a lone unfiltered tag reading data, else null. */
function bindingSym(value: Json, sym: Sym): string | null {
  if (typeof value !== 'string') return null
  try {
    const parsed = readString(value)
    return parsed.kind === 'raw' && !parsed.field.filters.length ? symbolic(parsed.field.path, sym) : null
  } catch {
    return null
  }
}

function useNode(value: Record<string, unknown>, pos: Pos, ctx: Ctx, sym: Sym, loc: Loc, out: Out): Json {
  if (typeof value.use !== 'string') return fail('bad-node', 'a use node names a component', at(loc, 'use'))
  if (value.variant !== undefined && typeof value.variant !== 'string') fail('bad-node', 'a use node\'s variant is an id', at(loc, 'variant'))
  if (value.with !== undefined && !isRecord(value.with)) fail('bad-node', 'a use node\'s with is an object', at(loc, 'with'))
  if (ctx.fromCode && !ctx.components.has(value.use) && !ctx.code?.has(value.use))
    fail('undeclared-use', `blocks from a code component use "${value.use}", which it did not declare in uses`, loc)
  const resolved = resolveUse(value.use, value.variant as string | undefined, value.with as JsonObject | undefined, pos, ctx, sym, loc, out)
  if (resolved.kind === 'code') return { code: resolved.id, params: resolved.params, ...(resolved.variant ? { variant: resolved.variant } : {}) }
  return { component: resolved.id, params: resolved.params, blocks: resolved.blocks ?? [], ...(resolved.variant ? { variant: resolved.variant } : {}) }
}

/* ------------------------------------------------------------------ the compile */

function newCtx(components: ReadonlyMap<string, Component>, options: CompileOptions): Ctx {
  return {
    components,
    code: options.code,
    filtersKnown: options.filters ? new Set(options.filters) : null,
    limits: { ...COMPILE_LIMITS, ...options.limits },
    nodes: 0,
    needs: new Map(),
    customKinds: new Set(),
    filters: new Set(),
    codeUsed: new Map(),
    componentsUsed: new Map(),
    embedded: new Map(),
    embedding: false,
    fromCode: false,
    ruleIds: null,
  }
}

/**
 * Compiles a recipe (by id from the library, or given) against a library into one self-contained JSON object. Throws a
 * `RecipeError` naming the place for an unknown component or variant, a missing required parameter, a cycle, a tag
 * that does not read, a key the block model does not have, or a limit passed.
 */
export function compileRecipe(recipe: string | Recipe, libraryIn: Library, options: CompileOptions = {}): CompiledRecipe {
  const library = readLibrary(libraryIn)
  const components = new Map(library.components.map((c) => [c.id, c]))
  const recipes = new Map(library.recipes.map((r) => [r.id, r]))
  const given = typeof recipe === 'string' ? recipes.get(recipe) ?? fail('unknown-recipe', `no recipe "${recipe}" in the library`, { root: 'library', steps: ['recipes'] }) : readRecipe(recipe)
  if (typeof recipe !== 'string') recipes.set(given.id, given)
  for (const id of options.code?.ids ?? [])
    if (components.has(id)) fail('component-conflict', `"${id}" is both a JSON component in the library and a code component`, { root: `component/${id}`, steps: [] })
  const ctx = newCtx(components, options)
  const resolved = resolveRecipe(given, recipes, ctx.limits)
  const root: Loc = { root: `recipe/${given.id}`, steps: [] }
  // The text rules are checked here, every problem named by its place, and carried as written.
  const textRules = resolved.document.textRules
  if (textRules !== undefined) {
    try {
      compileTextRules(textRules as unknown as TextRule[])
    } catch (error) {
      if (!(error instanceof TextRuleError)) throw error
      const problems = error.problems.map((p) => ({
        path: `${where(at(root, 'document', 'textRules', p.rule))}${p.key ? `.${p.key}` : ''}`,
        message: `text rule ${p.rule} ${p.message}`,
      }))
      throw new RecipeError('invalid', error.message.replace(/^document-kit: /, ''), problems[0]?.path ?? where(root), problems)
    }
    ctx.ruleIds = new Set(textRules.flatMap((rule) => (isRecord(rule) && typeof rule.id === 'string' ? [rule.id] : [])))
  } else ctx.ruleIds = new Set()

  const ids = new Set<string>()
  const sections: CompiledSection[] = resolved.sections.map((section, s) => {
    const loc = at(root, 'sections', s)
    const out = ['sections', s]
    const top: Sym = { params: new Map(), vars: new Map(), inLoop: false, part: null, stack: [] }
    const compiled: CompiledSection = { parts: [] }
    if (section.id !== undefined) compiled.id = section.id
    if (section.page !== undefined) compiled.page = walkValue(section.page, 'any', ctx, top, at(loc, 'page'), [...out, 'page'])
    if (section.pageNumbers !== undefined) compiled.pageNumbers = walkValue(section.pageNumbers, 'any', ctx, top, at(loc, 'pageNumbers'), [...out, 'pageNumbers'])
    if (section.header) compiled.header = walkList(section.header, 'blocks', ctx, top, at(loc, 'header'), [...out, 'header'])
    if (section.footer) compiled.footer = walkList(section.footer, 'blocks', ctx, top, at(loc, 'footer'), [...out, 'footer'])
    if (section.firstPage) {
      compiled.firstPage = {}
      if (section.firstPage.header) compiled.firstPage.header = walkList(section.firstPage.header, 'blocks', ctx, top, at(loc, 'firstPage', 'header'), [...out, 'firstPage', 'header'])
      if (section.firstPage.footer) compiled.firstPage.footer = walkList(section.firstPage.footer, 'blocks', ctx, top, at(loc, 'firstPage', 'footer'), [...out, 'firstPage', 'footer'])
    }
    compiled.parts = section.parts.map((part, p) => {
      const ploc = at(loc, 'parts', p)
      const pout = [...out, 'parts', p]
      const id = partId(part)
      if (ids.has(id)) fail('duplicate-part', `two parts are "${id}"; give one an id`, ploc)
      ids.add(id)
      let sym: Sym = { ...top, part: id }
      if (part.for !== undefined) sym = loop(part.for, ctx, sym, at(ploc, 'for'), [...pout, 'for']).sym
      if (part.if !== undefined) condition(part.if, ctx, sym, at(ploc, 'if'), [...pout, 'if'])
      const use = resolveUse(part.use, part.variant, part.with, 'block', ctx, sym, ploc, pout)
      const componentMeta = use.kind === 'json' ? ctx.components.get(use.id)?.meta : ctx.code?.get(use.id)?.meta
      const result: CompiledPart = {
        part: {
          id,
          component: use.id,
          title: part.title ?? use.title,
          ...(part.meta ? { meta: cloneJson(part.meta) as JsonObject } : {}),
          ...(componentMeta ? { componentMeta: cloneJson(componentMeta) as JsonObject } : {}),
        },
        params: use.params,
      }
      if (part.missing !== undefined) result.missing = part.missing
      if (part.pin !== undefined) result.pin = cloneJson(part.pin as Json) as CompiledPart['pin'] & Json
      if (part.if !== undefined) result.if = part.if
      if (part.for !== undefined) result.for = part.for
      if (part.pageBreakBefore !== undefined) result.pageBreakBefore = part.pageBreakBefore
      if (use.variant !== undefined) result.variant = use.variant
      if (use.kind === 'code') result.code = use.id
      else result.blocks = use.blocks ?? []
      return result
    })
    return compiled
  })

  const document: RecipeDocument = {}
  const top: Sym = { params: new Map(), vars: new Map(), inLoop: false, part: null, stack: [] }
  if (resolved.document.metadata) document.metadata = walkValue(resolved.document.metadata, 'any', ctx, top, at(root, 'document', 'metadata'), ['document', 'metadata']) as JsonObject
  if (resolved.document.theme) document.theme = walkValue(resolved.document.theme, 'any', ctx, top, at(root, 'document', 'theme'), ['document', 'theme']) as JsonObject
  if (resolved.document.footnotes) document.footnotes = cloneJson(resolved.document.footnotes) as Record<string, string[]>
  if (textRules !== undefined) document.textRules = cloneJson(textRules as Json) as Json[]

  const recipesFrom = resolved.chain.map((r) => ({ id: r.id, hash: contentHash(r) }))
  const componentsFrom = [...ctx.componentsUsed.values()].sort((a, b) => a.id.localeCompare(b.id)).map((c) => ({ id: c.id, hash: contentHash(c) }))
  const codeFrom = [...ctx.codeUsed.values()].sort((a, b) => a.id.localeCompare(b.id)).map((c) => ({ id: c.id, ...(c.version ? { version: c.version } : {}) }))
  const compiled: CompiledRecipe = {
    kind: 'compiled-recipe',
    version: RECIPE_FORMAT_VERSION,
    recipe: {
      id: given.id,
      name: given.name,
      ...(given.description ? { description: given.description } : {}),
      ...(given.meta ? { meta: cloneJson(given.meta) as JsonObject } : {}),
    },
    ...(resolved.missing && resolved.missing !== 'blank' ? { missing: resolved.missing } : {}),
    from: { recipes: recipesFrom, components: componentsFrom, code: codeFrom, hash: contentHash({ recipes: recipesFrom, components: componentsFrom, code: codeFrom }) },
    document,
    options: resolved.options,
    sections,
    needs: [...ctx.needs.values()].sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0)),
    customKinds: [...ctx.customKinds].sort(),
    filters: [...ctx.filters].filter((name) => !BUILT_IN_FILTERS.includes(name)).sort(),
    code: [...ctx.codeUsed.values()].sort((a, b) => a.id.localeCompare(b.id)).map(({ component: _component, ...need }) => need),
    ...(ctx.embedded.size ? { embedded: Object.fromEntries([...ctx.embedded].sort(([a], [b]) => a.localeCompare(b))) } : {}),
  }
  return canonical(compiled)
}

/**
 * Checks and compiles blocks a code component returned, at render: the same walk as the compile, with the compiled
 * recipe's embedded components and the code components as the library. Returns the compiled blocks.
 */
export function compileCodeBlocks(
  blocks: readonly unknown[],
  embedded: Readonly<Record<string, Component>>,
  code: CodeComponents | undefined,
  stack: readonly string[],
  part: string | null,
  limits: Partial<CompileLimits> = {},
): Json[] {
  const ctx = newCtx(new Map(Object.entries(embedded)), { ...(code ? { code } : {}), limits })
  ctx.fromCode = true
  const sym: Sym = { params: new Map(), vars: new Map(), inLoop: false, part, stack }
  return walkList([...blocks], 'blocks', ctx, sym, { root: `code/${stack[stack.length - 1] ?? '?'}`, steps: ['build'] }, [])
}
