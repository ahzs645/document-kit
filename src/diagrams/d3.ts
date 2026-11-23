import type { DiagramAdapter } from './index'
import type { DiagramData, DiagramFonts, DiagramSpec, DiagramTheme, DiagramType } from './types'
import { Canvas, check, defaults, escape, label, number } from './canvas'
import { org } from './org'
import { flow, network } from './graph'
import { bars, gantt, timeline } from './series'
import { cycle, footprint, metrics, overlap } from './other'
export type * from './types'

const templates: { [K in DiagramType]: (canvas: Canvas, data: DiagramData[K]) => number } = {
  'org-chart': org,
  gantt,
  flowchart: flow,
  timeline,
  'circular-process': cycle,
  network,
  bars,
  metrics,
  overlap,
  footprint,
}
export const diagramTypes = Object.freeze(Object.keys(templates) as DiagramType[])

export function renderD3(spec: DiagramSpec, theme: Partial<DiagramTheme> = {}, fonts?: DiagramFonts): string {
  check(spec && Object.hasOwn(templates, spec.type), 'unknown D3 template')
  const title = label(spec.title),
    w = number(spec.width ?? 960, 320, 5000)
  const t = { ...defaults, ...theme, ...spec.theme }
  for (const key of ['ink', 'paper', 'accent', 'onAccent', 'muted', 'panel', 'rule'] as const)
    check(/^#[0-9a-f]{3}(?:[0-9a-f]{3})?$/i.test(t[key]), 'invalid colour ' + key)
  number(t.fontSize, 8, 60)
  const c = new Canvas(w, t, fonts)
  // spec is a discriminated union. Each registry value receives its corresponding data.
  const draw = templates[spec.type] as (canvas: Canvas, data: DiagramSpec['data']) => number
  const h = draw(c, spec.data)
  return `<svg xmlns="http://www.w3.org/2000/svg" role="img" width="${w}" height="${spec.height === undefined ? h : number(spec.height, 1, 10000)}" viewBox="0 0 ${w} ${h}"><title>${escape(title)}</title><defs><marker id="arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="${t.ink}"/></marker></defs><rect width="${w}" height="${h}" fill="${t.paper}"/>${c.parts.join('')}</svg>`
}
export function d3Adapter(theme: Partial<DiagramTheme> = {}, fonts?: DiagramFonts): DiagramAdapter {
  return { id: 'd3', render: (source) => ({ svg: renderD3(source as DiagramSpec, theme, fonts) }) }
}
