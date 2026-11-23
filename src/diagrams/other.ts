import { pointRadial } from 'd3-shape'
import { scaleLinear } from 'd3-scale'
import { Canvas, check, number } from './canvas'
import type { DiagramData } from './types'

export function cycle(c: Canvas, d: DiagramData['circular-process']) {
  check(d.steps.length > 1, 'cycle needs at least two steps')
  const radius = c.w * 0.16,
    cx = radius + 24
  const descriptions = d.steps.map((s) => ({
    head: c.wrap(s.label, c.w - 2 * cx - 75, c.s, true),
    body: s.detail ? c.wrap(s.detail, c.w - 2 * cx - 75, c.s * 0.9) : [],
  }))
  const heights = descriptions.map((d) => d.head.length * c.s * 1.26 + d.body.length * c.s * 1.134 + 21)
  const h = Math.max(2 * radius + 65, heights.reduce((a, b) => a + b, 0) + 20),
    cy = h / 2
  let y = 15
  const count = d.steps.length
  const nodeRadius = Math.min(17, radius * Math.sin(Math.PI / count) * 0.65)
  for (let i = 0; i < count; i++) {
    const angle = (2 * Math.PI * i) / count
    const p = pointRadial(angle, radius),
      next = pointRadial(angle + (2 * Math.PI) / count - nodeRadius / radius, radius)
    c.parts.push(
      `<path d="M${cx + p[0]},${cy + p[1]} A${radius},${radius} 0 0 1 ${cx + next[0]},${cy + next[1]}" fill="none" stroke="${c.t.muted}" stroke-width="2" marker-end="url(#arrow)"/>`,
    )
    c.circle(cx + p[0], cy + p[1], nodeRadius, c.t.accent)
    c.text(String(i + 1), cx + p[0], cy + p[1] + c.s * 0.3, c.s * 0.9, true, c.t.onAccent, 'middle')
    const x = 2 * cx + 24,
      item = descriptions[i]!
    c.text(String(i + 1), x, y + c.s, c.s, true, c.t.muted)
    const bottom = c.lines(item.head, x + 32, y + c.s, c.s, true)
    c.lines(item.body, x + 32, bottom, c.s * 0.9, false, c.t.muted)
    y += heights[i]!
  }
  return h
}
export function metrics(c: Canvas, d: DiagramData['metrics']) {
  check(d.values.length, 'empty metrics')
  const cols = number(d.columns ?? 3, 1, d.values.length)
  check(Number.isInteger(cols), 'columns must be an integer')
  const w = (c.w - 12) / cols
  let top = 6
  for (let r = 0; r < d.values.length; r += cols) {
    const row = d.values
      .slice(r, r + cols)
      .map((v) => ({ value: c.wrap(v.value, w - 26, c.s * 1.8, true), name: c.wrap(v.label, w - 26, c.s) }))
    const h = Math.max(...row.map((v) => v.value.length * c.s * 2.27 + v.name.length * c.s * 1.26)) + 25
    row.forEach((v, i) => {
      const left = 6 + i * w
      c.rect(left, top, w - 6, h, c.t.panel)
      const end = c.lines(v.value, left + 12, top + c.s * 2, c.s * 1.8, true, c.t.accent)
      c.lines(v.name, left + 12, end, c.s)
    })
    top += h + 6
  }
  return top
}
export function overlap(c: Canvas, d: DiagramData['overlap']) {
  check(d.sets.length === 2, 'expected two sets')
  const r = c.w * 0.235,
    cy = r + 8
  for (const x of [c.w * 0.33, c.w * 0.67]) c.circle(x, cy, r, 'none', c.t.muted)
  d.sets.forEach((s, i) => {
    const lines = c.wrap(s, r * 1.1, c.s * 1.2, true)
    c.lines(
      lines,
      c.w * (i ? 0.72 : 0.28),
      cy - (lines.length - 1) * c.s * 0.76 + c.s * 0.4,
      c.s * 1.2,
      true,
      c.t.ink,
      'middle',
    )
  })
  return cy * 2
}
export function footprint(c: Canvas, d: DiagramData['footprint']) {
  check(d.points.length, 'empty scatter plot')
  const h = c.w * 0.48,
    left = 53,
    right = c.w - 40,
    top = 32,
    bottom = h - 40
  const x = scaleLinear([0, 1], [left, right]),
    y = scaleLinear([0, 1], [bottom, top])
  c.rect(left, top, right - left, bottom - top, c.t.panel, c.t.rule)
  for (const p of d.points)
    c.circle(
      x(number(p.x, 0, 1)),
      y(number(p.y, 0, 1)),
      number(p.radius, 1, 30),
      p.emphasis ? c.t.accent : c.t.muted,
      c.t.paper,
    )
  c.text(d.xLabels[0], left, h - 9, c.s, false, c.t.muted)
  c.text(d.xLabels[1], right, h - 9, c.s, false, c.t.muted, 'end')
  c.text(d.yLabels[1], left, 22, c.s, false, c.t.muted)
  c.text(d.yLabels[0], left - 6, bottom, c.s * 0.8, false, c.t.muted, 'end')
  return h
}
