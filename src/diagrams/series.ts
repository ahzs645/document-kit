import { scaleLinear, scaleUtc } from 'd3-scale'
import { Canvas, check, label, number } from './canvas'
import type { DiagramData } from './types'

export function gantt(c: Canvas, d: DiagramData['gantt']) {
  check(d.tasks.length, 'empty schedule')
  const dates = typeof d.tasks[0]!.start === 'string'
  function position(value: string | number) {
    if (!dates) {
      check(typeof value === 'number', 'mixed schedule axes')
      return number(value, -Infinity)
    }
    check(typeof value === 'string' && /^\d{4}-\d\d-\d\d$/.test(value), 'expected ISO calendar date')
    const t = Date.parse(value + 'T00:00:00Z')
    check(Number.isFinite(t) && new Date(t).toISOString().slice(0, 10) === value, 'invalid calendar date')
    return t
  }
  const tasks = d.tasks.map((t) => {
    const a = position(t.start),
      b = position(t.end ?? t.start)
    check(b >= a, 'task ends before its start')
    return { a, b, label: label(t.label) }
  })
  const left = Math.min(c.w * 0.36, 295),
    right = c.w - 42,
    lo = Math.min(...tasks.map((t) => t.a)),
    hi = Math.max(...tasks.map((t) => t.b))
  const domain = [lo, hi === lo ? lo + (dates ? 86400000 : 1) : hi]
  const x = scaleLinear(domain, [left, right])
  const dateTicks = scaleUtc(domain, [left, right]).ticks(5)
  const ticks = dates
    ? [
        ...new Set(dateTicks.map((v) => Date.UTC(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate()))),
      ].filter((v) => v >= lo && v <= hi)
    : x.ticks(6)
  const rows = tasks.map((t) => c.wrap(t.label, left - 30, c.s, true))
  const heights = rows.map((r) => Math.max(48, r.length * c.s * 1.26 + 14))
  const height = heights.reduce((s, h) => s + h, 0) + 62
  for (const tick of ticks) {
    c.path([
      [x(tick), 40],
      [x(tick), height - 14],
    ])
    const caption = dates ? new Date(tick).toISOString().slice(0, 10) : String(tick)
    c.text(caption, x(tick), 26, c.s * 0.83, false, c.t.muted, 'middle')
  }
  if (d.unit) c.text(d.unit, 5, 26, c.s * 0.9, true)
  let top = 43
  tasks.forEach((t, i) => {
    const h = heights[i]!,
      mid = top + h / 2
    c.lines(rows[i]!, 5, top + 6 + c.s, c.s, true)
    if (t.a === t.b) {
      const cx = x(t.a)
      c.parts.push(`<path fill="${c.t.accent}" d="M${cx} ${mid - 8}l8 8 -8 8 -8 -8z"/>`)
    } else c.rect(x(t.a), mid - 7, x(t.b) - x(t.a), 14, c.t.accent, 'none', 3)
    top += h
    c.path([
      [5, top],
      [right, top],
    ])
  })
  return height
}
export function bars(c: Canvas, d: DiagramData['bars']) {
  check(d.values.length, 'empty bar series')
  const vals = d.values.map((v) => number(v.value, -Infinity))
  const min = Math.min(0, ...vals),
    max = Math.max(0, ...vals)
  const x = scaleLinear([min, max === min ? 1 : max], [c.w * 0.36, c.w - 100])
  let y = 16
  d.values.forEach((item, i) => {
    const lines = c.wrap(item.label, c.w * 0.31, c.s, true)
    const h = Math.max(42, lines.length * c.s * 1.26 + 14),
      mid = y + h / 2
    c.lines(lines, 8, y + c.s, c.s, true)
    const vx = x(vals[i]!),
      zero = x(0)
    c.rect(Math.min(vx, zero), mid - 7, Math.abs(vx - zero), 14, c.t.accent)
    c.path(
      [
        [zero, y + 2],
        [zero, y + h - 2],
      ],
      c.t.muted,
    )
    c.text(String(item.value) + (d.unit ?? ''), c.w - 8, mid + c.s * 0.32, c.s, true, c.t.ink, 'end')
    y += h
  })
  return y + 8
}
export function timeline(c: Canvas, d: DiagramData['timeline']) {
  check(d.steps.length, 'empty timeline')
  let y = 16
  for (const [i, item] of d.steps.entries()) {
    const lines = c.wrap(item.label, c.w - 70, c.s * 1.1, true)
    const detail = item.detail ? c.wrap(item.detail, c.w - 70, c.s) : []
    const h = lines.length * c.s * 1.4 + detail.length * c.s * 1.26 + 22
    if (i + 1 < d.steps.length)
      c.path(
        [
          [22, y + 12],
          [22, y + h + 12],
        ],
        c.t.rule,
        2,
      )
    c.circle(22, y + 12, 13, c.t.accent)
    c.text(String(i + 1), 22, y + 17, c.s * 0.76, true, c.t.onAccent, 'middle')
    const end = c.lines(lines, 53, y + c.s * 1.05, c.s * 1.1, true)
    c.lines(detail, 53, end, c.s, false, c.t.muted)
    y += h
  }
  return y
}
