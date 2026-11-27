/**
 * A line height as the writers read it: a multiple of single spacing, an exact height in points or a least height in
 * points ({@link LineHeight}). One reading for the Word writer and the HTML writer, so both take the same values and
 * refuse the same mistakes.
 */
import type { LineHeight } from './types'

/** A line height, read: `multiple` (of the font's own line), `exact` or `atLeast` (points). */
export interface ResolvedLineHeight {
  readonly rule: 'multiple' | 'exact' | 'atLeast'
  readonly value: number
}

/** Word's largest line height, 1584 points (22 inches); its multiples go to 132 lines. */
const MAX_POINTS = 1584
const MAX_MULTIPLE = 132

const fail = (value: unknown, why: string): never => {
  throw new Error(`document-kit: the line height ${JSON.stringify(value)} ${why}`)
}

/**
 * A line height read and checked: a positive multiple (at most 132), or `{ exact }` or `{ atLeast }` with a positive
 * number of points (at most 1584), one of the two. Throws an error naming the value on anything else.
 */
export function readLineHeight(value: LineHeight): ResolvedLineHeight {
  if (typeof value === 'number') {
    if (!Number.isFinite(value) || value <= 0 || value > MAX_MULTIPLE) {
      fail(value, `is not a multiple of single spacing from more than 0 to ${MAX_MULTIPLE}`)
    }
    return { rule: 'multiple', value }
  }
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    return fail(value, 'is neither a multiple nor { exact } or { atLeast } in points')
  }
  const keys = Object.keys(value).filter((key) => (value as Record<string, unknown>)[key] !== undefined)
  if (keys.length !== 1 || (keys[0] !== 'exact' && keys[0] !== 'atLeast')) {
    return fail(value, 'takes one of exact or atLeast, in points')
  }
  const rule = keys[0] as 'exact' | 'atLeast'
  const points = (value as Record<string, unknown>)[rule]
  if (typeof points !== 'number' || !Number.isFinite(points) || points <= 0 || points > MAX_POINTS) {
    fail(value, `needs a number of points from more than 0 to ${MAX_POINTS}`)
  }
  return { rule, value: points as number }
}

/** True for a line height in points (`exact` or `atLeast`), false for a multiple or none. */
export const isFixedLineHeight = (value: LineHeight | undefined): boolean =>
  value !== undefined && typeof value === 'object'
