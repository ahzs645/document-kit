import { scaleLinear, scalePoint } from 'd3-scale'
import { pointRadial } from 'd3-shape'
import { Canvas, check, label, number, type Point } from './canvas'
import type { DiagramData, Edge, Node } from './types'

function draw(
  c: Canvas,
  nodes: Node[],
  edges: Edge[],
  positions: Point[],
  width: number,
  height: number,
  waypoint?: (v: Point) => Point,
) {
  const cards = nodes.map((n, i) => {
    const title = c.wrap(n.label, width - 20, c.s, true)
    const detail = n.detail ? c.wrap(n.detail, width - 20, c.s * 0.86) : []
    return {
      n,
      p: positions[i]!,
      title,
      detail,
      h: 18 + title.length * c.s * 1.26 + detail.length * c.s * 1.09,
    }
  })
  const map = new Map(cards.map((card) => [label(card.n.id), card]))
  check(map.size === cards.length, 'duplicate graph node')
  const border = (card: (typeof cards)[number], towards: Point): Point => {
    const [x, y] = card.p,
      dx = towards[0] - x,
      dy = towards[1] - y
    const k = Math.max(Math.abs(dx) / (width / 2), Math.abs(dy) / (card.h / 2))
    check(k > 0, 'a graph edge needs distinct endpoints or a bend')
    return [x + dx / k, y + dy / k]
  }
  for (const e of edges) {
    const a = map.get(e.from),
      b = map.get(e.to)
    check(a && b, 'edge references unknown node')
    const via = (e.via ?? []).map((p) => (waypoint ? waypoint(p) : p))
    c.path(
      [border(a, via[0] ?? b.p), ...via, border(b, via.at(-1) ?? a.p)],
      e.dashed ? c.t.muted : c.t.ink,
      1.8,
      e.dashed,
      e.arrow !== false,
    )
  }
  for (const card of cards) {
    const [x, y] = card.p,
      top = y - card.h / 2
    c.rect(
      x - width / 2,
      top,
      width,
      card.h,
      card.n.emphasis ? c.t.accent : c.t.panel,
      card.n.emphasis ? c.t.accent : c.t.rule,
      4,
    )
    const end = c.lines(
      card.title,
      x,
      top + 9 + c.s * 0.9,
      c.s,
      true,
      card.n.emphasis ? c.t.onAccent : c.t.ink,
      'middle',
    )
    c.lines(card.detail, x, end, c.s * 0.86, false, card.n.emphasis ? c.t.onAccent : c.t.muted, 'middle')
  }
  return height
}
export function flow(c: Canvas, d: DiagramData['flowchart']) {
  check(d.nodes.length, 'empty flowchart')
  const cols = d.nodes.map((n) => number(n.column, 0)),
    rows = d.nodes.map((n) => number(n.row, 0))
  const count = Math.max(...cols) + 1
  const width = Math.min(255, ((c.w - 24) / count) * 0.85)
  check(width > 60, 'increase width for this many flowchart columns')
  const maxH = Math.max(
    ...d.nodes.map(
      (n) =>
        18 +
        c.wrap(n.label, width - 20, c.s, true).length * c.s * 1.26 +
        (n.detail ? c.wrap(n.detail, width - 20, c.s * 0.86).length * c.s * 1.09 : 0),
    ),
  )
  const step = maxH + 35,
    height = step * (Math.max(...rows) + 1)
  const x = scaleLinear([0, count - 1], [width / 2 + 6, c.w - width / 2 - 6])
  const pos = ([col, row]: Point): Point => [x(col), (row + 0.5) * step]
  return draw(
    c,
    d.nodes,
    d.edges,
    d.nodes.map((n) => pos([n.column, n.row])),
    width,
    height,
    pos,
  )
}
export function network(c: Canvas, d: DiagramData['network']) {
  check(d.nodes.length, 'empty network')
  const w = Math.min(c.w * 0.25, 200),
    rx = (c.w - w - 20) / 2,
    ry = Math.max(110, d.nodes.length * 25)
  const h = 2 * ry + 112,
    a = scalePoint<string>()
      .domain(d.nodes.map((n) => n.id))
      .range([0, 2 * Math.PI * (1 - 1 / d.nodes.length)])
  const points = d.nodes.map((n) => {
    const [x, y] = pointRadial(a(n.id)!, 1)
    return [c.w / 2 + x * rx, h / 2 + y * ry] as Point
  })
  return draw(c, d.nodes, d.edges, points, w, h)
}
