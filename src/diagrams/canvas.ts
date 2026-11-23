import { line } from 'd3-shape'
import type { DiagramFonts, DiagramTheme } from './types'

export const defaults: DiagramTheme = {
  ink: '#151515',
  paper: '#ffffff',
  accent: '#222222',
  onAccent: '#ffffff',
  muted: '#505050',
  panel: '#f3f3f3',
  rule: '#bcbcbc',
  font: 'Arial',
  fontSize: 17,
}
export const escape = (s: string) =>
  s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]!)
export function check(ok: unknown, message: string): asserts ok {
  if (!ok) throw new Error('Diagram: ' + message)
}
export function label(v: unknown): string {
  check(typeof v === 'string' && v.trim(), 'expected a nonempty label')
  return v
}
export function number(v: number, min: number, max = Infinity) {
  check(Number.isFinite(v) && v >= min && v <= max, 'number outside supported range: ' + v)
  return v
}
export type Point = [number, number]
export class Canvas {
  parts: string[] = []
  constructor(
    readonly w: number,
    readonly t: DiagramTheme,
    readonly fonts?: DiagramFonts,
  ) {}
  get s() {
    return this.t.fontSize
  }
  measure(text: string, size: number, bold: boolean) {
    return (
      this.fonts?.measure(text, size, bold) ??
      Array.from(text).reduce(
        (sum, c) => sum + (/[MW@%]/.test(c) ? 0.95 : /[il.,' ]/.test(c) ? 0.33 : 0.63) * size,
        0,
      )
    )
  }
  wrap(text: string, max: number, size = this.s, bold = false): string[] {
    label(text)
    const lines: string[] = []
    for (const paragraph of text.split('\n')) {
      let current = ''
      for (const word of paragraph.trim().split(/\s+/)) {
        if (current && this.measure(current + ' ' + word, size, bold) > max) {
          lines.push(current)
          current = ''
        }
        if (this.measure(word, size, bold) > max) {
          for (const c of word) {
            if (current && this.measure(current + c, size, bold) > max) {
              lines.push(current)
              current = ''
            }
            current += c
          }
        } else current += (current ? ' ' : '') + word
      }
      lines.push(current)
    }
    return lines
  }
  text(
    text: string,
    x: number,
    y: number,
    size = this.s,
    bold = false,
    color = this.t.ink,
    anchor: 'start' | 'middle' | 'end' = 'start',
  ) {
    if (!text) return
    const offset = anchor === 'start' ? 0 : this.measure(text, size, bold) * (anchor === 'middle' ? 0.5 : 1)
    if (this.fonts) {
      this.parts.push(
        '<path fill="' +
          color +
          '" d="' +
          escape(this.fonts.outline(text, size, bold, x - offset, y)) +
          '"><title>' +
          escape(text) +
          '</title></path>',
      )
    } else {
      this.parts.push(
        '<text x="' +
          x +
          '" y="' +
          y +
          '" text-anchor="' +
          anchor +
          '" font-family="' +
          escape(this.t.font) +
          '" font-size="' +
          size +
          '" font-weight="' +
          (bold ? 700 : 400) +
          '" fill="' +
          color +
          '">' +
          escape(text) +
          '</text>',
      )
    }
  }
  lines(
    lines: string[],
    x: number,
    y: number,
    size = this.s,
    bold = false,
    color = this.t.ink,
    anchor: 'start' | 'middle' = 'start',
  ) {
    for (const text of lines) {
      this.text(text, x, y, size, bold, color, anchor)
      y += size * 1.26
    }
    return y
  }
  rect(x: number, y: number, w: number, h: number, fill: string, stroke = 'none', radius = 0) {
    this.parts.push(
      `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${radius}" fill="${fill}" stroke="${stroke}"/>`,
    )
  }
  circle(x: number, y: number, r: number, fill: string, stroke = 'none') {
    this.parts.push(`<circle cx="${x}" cy="${y}" r="${r}" fill="${fill}" stroke="${stroke}"/>`)
  }
  path(points: Point[], color = this.t.rule, width = 1, dashed = false, arrow = false) {
    this.parts.push(
      `<path d="${line()(points)}" fill="none" stroke="${color}" stroke-width="${width}" stroke-linejoin="round"${dashed ? ' stroke-dasharray="6 5"' : ''}${arrow ? ' marker-end="url(#arrow)"' : ''}/>`,
    )
  }
}
