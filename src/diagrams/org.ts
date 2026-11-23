import { cluster, stratify } from 'd3-hierarchy'
import { Canvas, check, label } from './canvas'
import type { DiagramData, Item, OrgNode } from './types'

export function org(c: Canvas, data: DiagramData['org-chart']) {
  const ids = new Set(data.nodes.map((n) => label(n.id)))
  check(ids.size === data.nodes.length, 'duplicate organization node')
  const root = stratify<OrgNode>()
    .id((n) => n.id)
    .parentId((n) => n.parent)(data.nodes)
  const gap = 12,
    margin = 8,
    leaves = root.leaves().length
  const width = (c.w - 2 * margin - gap * (leaves - 1)) / leaves
  check(width > 65, 'increase width for this many department columns')
  const tree = cluster<OrgNode>()
    .nodeSize([width + gap, 1])
    .separation(() => 1)(root)
  const x0 = Math.min(...tree.leaves().map((n) => n.x))
  const x = (v: number) => margin + v - x0
  const pad = 10
  const lineHeight = (rows: string[], size: number) => rows.length * size * 1.26
  function card(n: Item, size: number) {
    const title = c.wrap(n.label, width - 2 * pad, size, true)
    const detail = n.detail ? c.wrap(n.detail, width - 2 * pad, size * 0.87) : []
    return { title, detail, height: lineHeight(title, size) + lineHeight(detail, size * 0.87) + pad * 1.5 }
  }
  const nodes = tree.descendants().map((n) => {
    const tier = tree.height - n.height
    const head = card(n.data, c.s)
    const members = (n.data.members ?? []).map((m) => card(m, c.s * 0.93))
    return { n, tier, head, members, h: head.height + members.reduce((sum, m) => sum + m.height, 0) }
  })
  const rows: number[] = [8]
  for (let i = 0; i <= tree.height; i++)
    rows.push(rows[i]! + Math.max(0, ...nodes.filter((n) => n.tier === i).map((n) => n.h)) + 30)
  const lookup = new Map(nodes.map((n) => [n.n.id, n]))
  for (const n of nodes) {
    if (!n.n.parent) continue
    const p = lookup.get(n.n.parent.id)!
    const sx = x(p.n.x) + width / 2,
      tx = x(n.n.x) + width / 2,
      top = rows[p.tier]! + p.h,
      bottom = rows[n.tier]!
    const bend = bottom - 15
    c.path(
      [
        [sx, top],
        [sx, bend],
        [tx, bend],
        [tx, bottom],
      ],
      c.t.muted,
      1.6,
    )
  }
  for (const node of nodes) {
    const left = x(node.n.x),
      top = rows[node.tier]!
    const { head } = node
    c.rect(left, top, width, head.height, c.t.accent)
    let y = c.lines(head.title, left + pad, top + pad + c.s * 0.85, c.s, true, c.t.onAccent)
    c.lines(head.detail, left + pad, y, c.s * 0.87, false, c.t.onAccent)
    y = top + head.height
    for (const m of node.members) {
      const size = c.s * 0.93
      const next = c.lines(m.title, left + pad, y + size + 5, size, true)
      c.lines(m.detail, left + pad, next, size * 0.87, false, c.t.muted)
      y += m.height
      c.path([
        [left, y],
        [left + width, y],
      ])
    }
  }
  return rows.at(-1)! - 20
}
