/**
 * Kerning as the writers read it: the size from which a run's letter pairs are kerned. One reading for the Word writer
 * and the HTML writer, so both take the same values and refuse the same mistakes.
 */
import type { Kerning } from './types'

/** Word's largest font size, 1638 points; a threshold above it kerns nothing. */
const MAX_POINTS = 1638

/**
 * The smallest size, in points, that a kerning setting kerns text at: `true` kerns at every size (from 1 point,
 * Word's "Kerning for fonts: 1 pt and above"), a number from that many points up, `false` or `0` not at all (`0`).
 * Throws an error naming the value on anything else.
 */
export function kerningFrom(value: Kerning): number {
  if (value === true) return 1
  if (value === false) return 0
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0 || value > MAX_POINTS) {
    throw new Error(
      `document-kit: the kerning ${JSON.stringify(value)} is not true, false or a size in points from 0 to ${MAX_POINTS}`,
    )
  }
  return value
}

/** Whether text of `size` points is kerned under a setting (an unknown size counts as kerned when anything is). */
export function kernsAt(value: Kerning, size: number | undefined): boolean {
  const from = kerningFrom(value)
  if (from === 0) return false
  return size === undefined || size >= from
}
