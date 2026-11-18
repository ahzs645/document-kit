/** Points to the units a Word file counts in. */
import { BorderStyle, type IBorderOptions } from 'docx'
import type { Border } from '../model/types'

/** Twentieths of a point: page sizes, margins, indents, spacing, table widths. */
export const twips = (points: number): number => Math.round(points * 20)
/** Half-points: font sizes. */
export const halfPoints = (points: number): number => Math.round(points * 2)
/** Eighths of a point: border widths, from 2 (a quarter point) to 96 (twelve points). */
export const eighths = (points: number): number => Math.min(96, Math.max(2, Math.round(points * 8)))
/** Pixels at 96 to the inch, which docx takes for a picture's drawn size. */
export const pixels = (points: number): number => Math.max(1, Math.round((points * 96) / 72))
/** Line height as a multiple, in Word's 240ths. */
export const lineMultiple = (multiple: number): number => Math.round(multiple * 240)

/** A colour as six upper-case hex digits; throws on anything else, naming it. */
export function hex(color: string): string {
  const cleaned = color.trim().replace(/^#/, '')
  if (!/^[0-9a-fA-F]{6}$/.test(cleaned)) throw new Error(`document-kit: "${color}" is not a six-digit hex colour`)
  return cleaned.toUpperCase()
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
    space: 0,
  }
}
