/**
 * Values as the render reads them: data paths with provenance (a `sources` map by path, or a value wrapped as
 * `{ "$value": …, "$source": … }`), filters (the Word template engine's `format`, `upper` and `lower`, through its own
 * picture code, and the application's), and text. Pure: no DOM, no network.
 */
import { dateOf, formatDate, formatNumber, numberOf, type NumberTextSeparators, type PictureSettings } from '../template/templateFormat'
import { RecipeError } from './errors'
import { hasOwn, isRecord } from './json'
import type { CompiledFilter } from './tags'

/** A filter of the application's own, used in a tag as `{{ value | name "arg" }}`. */
export type RecipeFilter = (value: unknown, ...args: Array<string | number>) => unknown

/** How pictures write numbers and dates, and how text in the data reads as a number. Defaults as the Word engine's. */
export interface FormatOptions {
  numberSeparators?: { decimal?: string; group?: string }
  /** How numbers written as text in the data are read (`{ decimal: ',', group: '.' }` reads `1.234,50`). */
  dataNumberSeparators?: NumberTextSeparators | readonly NumberTextSeparators[]
  dateNames?: Partial<Pick<PictureSettings, 'months' | 'monthsShort' | 'days' | 'daysShort' | 'am' | 'pm'>>
  utcOffsetMinutes?: number
  /** `number` compares two texts that both read as numbers as numbers; text by default, as the Word engine does. */
  numericTextComparison?: 'text' | 'number'
}

export const ENGLISH_DATE_NAMES = {
  months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  monthsShort: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  daysShort: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  am: 'AM',
  pm: 'PM',
}

export interface Formatting {
  pictures: PictureSettings
  ways: readonly NumberTextSeparators[] | null
  numericText: 'text' | 'number'
}

export function formatting(options: FormatOptions = {}): Formatting {
  const ways = options.dataNumberSeparators
  return {
    pictures: {
      decimal: options.numberSeparators?.decimal ?? '.',
      group: options.numberSeparators?.group ?? ',',
      ...ENGLISH_DATE_NAMES,
      ...options.dateNames,
      utcOffsetMinutes: options.utcOffsetMinutes ?? 0,
    },
    ways: ways === undefined ? null : Array.isArray(ways) ? (ways as readonly NumberTextSeparators[]) : [ways as NumberTextSeparators],
    numericText: options.numericTextComparison ?? 'text',
  }
}

/* ------------------------------------------------------------------ provenance */

/** Whether a value is a wrapped value: `{ "$value": …, "$source": "…" }`. */
export function isWrapped(value: unknown): value is { $value: unknown; $source?: string } {
  return isRecord(value) && hasOwn(value, '$value') && Object.keys(value).every((key) => key === '$value' || key === '$source')
}

/** A value with every wrapper inside it taken off. */
export function unwrapDeep(value: unknown): unknown {
  if (isWrapped(value)) return unwrapDeep(value.$value)
  if (Array.isArray(value)) return value.map(unwrapDeep)
  if (isRecord(value)) {
    const out: Record<string, unknown> = {}
    for (const [key, member] of Object.entries(value)) out[key] = unwrapDeep(member)
    return out
  }
  return value
}

/** A `sources` map read once: each key's steps, `*` matching any one step. */
export interface SourceMap {
  entries: Array<{ steps: string[]; source: string; wild: number }>
}

export function sourceMap(sources: Readonly<Record<string, string>> | undefined): SourceMap {
  return {
    entries: Object.entries(sources ?? {}).map(([key, source]) => {
      const steps = key.replace(/\[\]/g, '.*').split('.').filter(Boolean)
      return { steps, source, wild: steps.filter((step) => step === '*').length }
    }),
  }
}

/** The source the map gives a data path: the most specific key that is a prefix of it, and that key's length. */
export function mapSource(map: SourceMap, path: readonly string[]): { source: string; depth: number } | null {
  let best: SourceMap['entries'][number] | null = null
  for (const entry of map.entries) {
    if (entry.steps.length > path.length) continue
    if (!entry.steps.every((step, i) => step === '*' || step === path[i])) continue
    if (!best || entry.steps.length > best.steps.length || (entry.steps.length === best.steps.length && entry.wild < best.wild)) best = entry
  }
  return best ? { source: best.source, depth: best.steps.length } : null
}

/* ------------------------------------------------------------------ filters and text */

/** A value's text, or null when it has none (null, undefined) — and a `not-text` error for a list or an object. */
export function valueText(value: unknown, what: string, where: string): string | null {
  if (value === null || value === undefined) return null
  if (typeof value === 'string') return value
  if (typeof value === 'number' || typeof value === 'bigint' || typeof value === 'boolean') return String(value)
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value.toISOString()
  throw new RecipeError('not-text', `${what} is ${Array.isArray(value) ? 'a list' : 'an object'}, printed where text goes; give a field of it, or loop over it`, where)
}

/** A value through filters, in order. A built-in filter leaves an empty value empty; the application's get it as is. */
export function applyFilters(
  value: unknown,
  filters: readonly CompiledFilter[],
  format: Formatting,
  custom: Readonly<Record<string, RecipeFilter>>,
  what: string,
  where: string,
  onUnknown?: (name: string) => void,
): unknown {
  let current = value
  for (const filter of filters) {
    if (filter.kind === 'custom') {
      const run = hasOwn(custom, filter.name) ? custom[filter.name] : undefined
      if (!run) {
        if (onUnknown) {
          onUnknown(filter.name)
          continue
        }
        throw new RecipeError('unknown-filter', `no filter "${filter.name}" was given to the render`, where)
      }
      current = run(current, ...filter.args)
      continue
    }
    if (current === null || current === undefined) continue
    if (filter.kind === 'format') {
      if (filter.picture.kind === 'number') {
        const n = numberOf(current, format.ways)
        const written = n === null ? null : formatNumber(n, filter.picture, format.pictures)
        if (written === null) throw new RecipeError('bad-value', `${what} (${JSON.stringify(current)}) is not a number for the picture "${filter.text}"`, where)
        current = written
      } else {
        const date = dateOf(current)
        const written = date === null ? null : formatDate(date, filter.picture, format.pictures)
        if (written === null) throw new RecipeError('bad-value', `${what} (${JSON.stringify(current)}) is not a date for the picture "${filter.text}"`, where)
        current = written
      }
      continue
    }
    const text = valueText(current, what, where) ?? ''
    current = filter.kind === 'upper' ? text.toUpperCase() : text.toLowerCase()
  }
  return current
}

/** Whether a value counts as no value when printed: missing, null, or empty text. */
export const isEmptyValue = (value: unknown): boolean => value === null || value === undefined || value === ''

export { numberOf }
