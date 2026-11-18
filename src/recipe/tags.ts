/**
 * The tag language of JSON components: the Word template engine's own (`src/template/templateTags.ts` reads every tag
 * and condition, `templateFormat.ts` reads and writes pictures), so a `{{ tag }}` means the same in a Word template and
 * in a component. A string is read once into segments: text, fields, and the inline `{{ if }}` and `{{ for }}` blocks
 * the Word engine takes within a paragraph. Pure: no DOM, no network.
 */
import {
  conditionFilters,
  readTag,
  type TemplateExpression,
  type TemplateFilter,
  type TemplateTag,
} from '../template/templateTags'
import { readPicture, type Picture } from '../template/templateFormat'

/** The filters the language has itself; any other is the application's (`filters` at render). */
export const BUILT_IN_FILTERS: readonly string[] = ['format', 'upper', 'lower']

export type CompiledFilter =
  | { kind: 'format'; picture: Picture; text: string }
  | { kind: 'upper' | 'lower' }
  | { kind: 'custom'; name: string; args: (string | number)[] }

export interface FieldSeg {
  t: 'field'
  path: string[]
  text: string
  filters: CompiledFilter[]
}

export interface Condition {
  text: string
  expression: TemplateExpression
  /** Each filter chain of the condition, by its place (the template engine's `chain`). */
  chains: CompiledFilter[][]
}

export interface Loop {
  name: string
  list: string[]
  listText: string
}

export type Seg =
  | { t: 'text'; text: string }
  | FieldSeg
  | { t: 'if'; branches: Array<{ condition: Condition; body: Seg[] }>; otherwise: Seg[] | null }
  | ({ t: 'for'; body: Seg[] } & Loop)

/** A string as read: plain text, exactly one field (its raw value), or a mix. */
export type ParsedString =
  | { kind: 'plain'; text: string }
  | { kind: 'raw'; field: FieldSeg }
  | { kind: 'mixed'; segs: Seg[] }

/** Thrown by the readers below with what is wrong; the caller adds where. */
export class TagSyntaxError extends Error {}

const fail = (message: string): never => {
  throw new TagSyntaxError(message)
}

function compileFilters(filters: readonly TemplateFilter[] | undefined, typed: string): CompiledFilter[] {
  return (filters ?? []).map((filter): CompiledFilter => {
    switch (filter.name) {
      case 'format': {
        const [picture, ...extra] = filter.args
        if (typeof picture !== 'string' || extra.length)
          return fail(`${typed} formats with something other than one quoted picture, as in format "#,##0.00"`)
        const reading = readPicture(picture)
        if (!reading.ok) return fail(`${typed} formats with the picture "${picture}", which ${reading.reason}`)
        return { kind: 'format', picture: reading.picture, text: picture }
      }
      case 'upper':
      case 'lower':
        if (filter.args.length) return fail(`${typed} writes something after ${filter.name}, which takes nothing`)
        return { kind: filter.name }
      default:
        return { kind: 'custom', name: filter.name, args: [...filter.args] }
    }
  })
}

/** A condition's text (no braces): `status == "final" and total > 0`. */
export function readCondition(text: string): Condition {
  const reading = readTag(`if ${text}`)
  if (!reading || !reading.ok || reading.tag.kind !== 'if')
    return fail(`the condition "${text}" ${reading && !reading.ok ? reading.reason : 'cannot be read'}`)
  const expression = reading.tag.condition
  const chains = conditionFilters(expression).map((chain) => compileFilters(chain, `the condition "${text}"`))
  return { text, expression, chains }
}

/** A loop's text: `item in list`. */
export function readLoop(text: string): Loop {
  const reading = readTag(`for ${text}`)
  if (!reading || !reading.ok || reading.tag.kind !== 'for')
    return fail(`the loop "${text}" ${reading && !reading.ok ? reading.reason : 'is not written as item in list'}`)
  const { name, list, listText } = reading.tag
  if (name === 'params' || name === 'options') return fail(`the loop "${text}" names its item ${name}, a reserved name`)
  return { name, list, listText }
}

/** Reads a string's tags. Every `{{` starts a tag; a tag that does not read is an error. */
export function readString(text: string): ParsedString {
  if (!text.includes('{{')) {
    if (text.includes('}}')) fail(`"}}" with no "{{" before it`)
    return { kind: 'plain', text }
  }
  type Frame =
    | { kind: 'root'; body: Seg[] }
    | { kind: 'if'; node: Extract<Seg, { t: 'if' }>; body: Seg[] }
    | { kind: 'for'; node: Extract<Seg, { t: 'for' }>; body: Seg[] }
  const root: Frame = { kind: 'root', body: [] }
  const stack: Frame[] = [root]
  const top = () => stack[stack.length - 1] ?? root
  let at = 0
  while (at < text.length) {
    const open = text.indexOf('{{', at)
    const before = text.slice(at, open < 0 ? text.length : open)
    if (before.includes('}}')) fail(`"}}" with no "{{" before it`)
    if (before) top().body.push({ t: 'text', text: before })
    if (open < 0) break
    const close = text.indexOf('}}', open + 2)
    if (close < 0) fail(`a tag with no closing "}}": ${text.slice(open, open + 40)}`)
    const inner = text.slice(open + 2, close)
    const typed = `{{${inner}}}`
    if (inner.includes('{{')) fail(`a tag inside a tag: ${typed}`)
    const reading = readTag(inner)
    if (reading === null) fail(`${typed} is not a tag (a name has no spaces; a filter is format, upper, lower or one the application gives)`)
    if (reading && !reading.ok) fail(`${typed} ${reading.reason}`)
    const tag = (reading as { ok: true; tag: TemplateTag }).tag
    const frame = top()
    switch (tag.kind) {
      case 'field':
        frame.body.push({ t: 'field', path: tag.path, text: tag.text, filters: compileFilters(tag.filters, typed) })
        break
      case 'if': {
        const condition = readCondition(inner.trim().replace(/^if\s+/i, ''))
        const node: Extract<Seg, { t: 'if' }> = { t: 'if', branches: [{ condition, body: [] }], otherwise: null }
        frame.body.push(node)
        stack.push({ kind: 'if', node, body: node.branches[0]?.body ?? [] })
        break
      }
      case 'elseif':
      case 'else': {
        if (frame.kind !== 'if') fail(`${typed} with no open {{ if }}`)
        const node = (frame as Extract<Frame, { kind: 'if' }>).node
        if (node.otherwise) fail(`${typed} after the {{ else }}`)
        stack.pop()
        if (tag.kind === 'else') {
          node.otherwise = []
          stack.push({ kind: 'if', node, body: node.otherwise })
        } else {
          const condition = readCondition(inner.trim().replace(/^else\s+if\s+/i, ''))
          const branch = { condition, body: [] as Seg[] }
          node.branches.push(branch)
          stack.push({ kind: 'if', node, body: branch.body })
        }
        break
      }
      case 'endif':
        if (frame.kind !== 'if') fail(`${typed} with no open {{ if }}`)
        stack.pop()
        break
      case 'for': {
        const loop = readLoop(inner.trim().replace(/^for\s+/i, ''))
        const node: Extract<Seg, { t: 'for' }> = { t: 'for', ...loop, body: [] }
        frame.body.push(node)
        stack.push({ kind: 'for', node, body: node.body })
        break
      }
      case 'endfor':
        if (frame.kind !== 'for') fail(`${typed} with no open {{ for }}`)
        stack.pop()
        break
    }
    at = close + 2
  }
  if (stack.length > 1) fail(`an {{ ${top().kind} }} that is never closed`)
  const segs = root.body
  // Exactly one field and nothing else: the field's raw value. Text around it, even a space, makes it text.
  const only = segs.length === 1 ? segs[0] : undefined
  if (only?.t === 'field') return { kind: 'raw', field: only }
  return { kind: 'mixed', segs }
}

/** Every field, condition and loop of segments, for listing what a string reads. */
export function* walkSegs(
  segs: readonly Seg[],
  loops: readonly Loop[] = [],
): Generator<
  | { kind: 'field'; field: FieldSeg; loops: readonly Loop[] }
  | { kind: 'condition'; condition: Condition; loops: readonly Loop[] }
  | { kind: 'loop'; loop: Loop; loops: readonly Loop[] }
> {
  for (const seg of segs) {
    if (seg.t === 'field') yield { kind: 'field', field: seg, loops }
    else if (seg.t === 'if') {
      for (const branch of seg.branches) {
        yield { kind: 'condition', condition: branch.condition, loops }
        yield* walkSegs(branch.body, loops)
      }
      if (seg.otherwise) yield* walkSegs(seg.otherwise, loops)
    } else if (seg.t === 'for') {
      yield { kind: 'loop', loop: seg, loops }
      yield* walkSegs(seg.body, [...loops, seg])
    }
  }
}

/** The custom filter names a list of compiled filters uses. */
export const customFilterNames = (filters: readonly CompiledFilter[]): string[] =>
  filters.flatMap((filter) => (filter.kind === 'custom' ? [filter.name] : []))
