/** Points to the units a Word file counts in. */
import { BorderStyle, LineRuleType, type IBorderOptions } from 'docx'
import { kerningFrom } from '../model/kerning'
import { isFixedLineHeight, readLineHeight } from '../model/lineHeight'
import type { Border, Kerning, LineHeight } from '../model/types'

/** Twentieths of a point: page sizes, margins, indents, spacing, table widths. */
export const twips = (points: number): number => Math.round(points * 20)
/** Half-points: font sizes. */
export const halfPoints = (points: number): number => Math.round(points * 2)
/** Eighths of a point: border widths, from 2 (a quarter point) to 96 (twelve points). */
export const eighths = (points: number): number => Math.min(96, Math.max(2, Math.round(points * 8)))
/** Pixels at 96 to the inch, which docx takes for a picture's drawn size. */
export const pixels = (points: number): number => Math.max(1, Math.round((points * 96) / 72))
/**
 * Kerning as Word's `w:kern`: the size kerning starts at, in half-points. The docx library leaves out a 0; a value under
 * one writes the explicit 0 that undoes a style's kerning.
 */
export const kernHalfPoints = (value: Kerning): number => halfPoints(kerningFrom(value)) || Number.EPSILON
/** Line height as a multiple, in Word's 240ths. */
export const lineMultiple = (multiple: number): number => Math.round(multiple * 240)

/**
 * A line height as Word's `w:spacing` `w:line` and `w:lineRule`: a multiple in 240ths, an exact or least height in
 * twips (`exact`, `atLeast`). A multiple is written without a rule, as the kit always has, unless it stands over a
 * height in points it would otherwise be read as (`inherited`, the line height of the style beneath): Word reads a
 * `w:line` without a rule as a multiple, but the layout and LibreOffice take the rule from the style, so there it is
 * written `auto`.
 */
export function lineSpacing(
  value: LineHeight | undefined,
  inherited?: LineHeight,
): { line?: number; lineRule?: (typeof LineRuleType)[keyof typeof LineRuleType] } {
  if (value === undefined) return {}
  const read = readLineHeight(value)
  if (read.rule === 'exact') return { line: twips(read.value), lineRule: LineRuleType.EXACT }
  if (read.rule === 'atLeast') return { line: twips(read.value), lineRule: LineRuleType.AT_LEAST }
  return isFixedLineHeight(inherited)
    ? { line: lineMultiple(read.value), lineRule: LineRuleType.AUTO }
    : { line: lineMultiple(read.value) }
}

/**
 * A colour as Word writes it: six upper-case hex digits. Takes `#rrggbb`, `rrggbb`, `#rgb` or `rgb` in any case (the
 * short form doubles each digit, as CSS does: `#0a3` is `00AA33`); throws on anything else, naming the value.
 */
export function hex(color: string): string {
  const cleaned = typeof color === 'string' ? color.trim().replace(/^#/, '') : ''
  if (/^[0-9a-fA-F]{6}$/.test(cleaned)) return cleaned.toUpperCase()
  if (/^[0-9a-fA-F]{3}$/.test(cleaned)) return [...cleaned].map((digit) => digit + digit).join('').toUpperCase()
  throw new Error(
    `document-kit: ${JSON.stringify(color)} is not a hex colour; write it as #rrggbb or #rgb (the # is optional)`,
  )
}

const STYLES = {
  single: BorderStyle.SINGLE,
  double: BorderStyle.DOUBLE,
  dashed: BorderStyle.DASHED,
  dotted: BorderStyle.DOTTED,
  none: BorderStyle.NONE,
} as const

export function border(value: Border | undefined): IBorderOptions | undefined {
  if (!value) return undefined
  const style = STYLES[value.style ?? 'single']
  if (style === BorderStyle.NONE) return { style, size: 0, color: 'auto' }
  return {
    style,
    size: eighths(value.width ?? 0.5),
    color: value.color ? hex(value.color) : 'auto',
    space: Math.max(0, Math.round(value.space ?? 0)),
  }
}
