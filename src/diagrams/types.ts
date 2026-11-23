export interface Item {
  label: string
  detail?: string
}
export interface OrgNode extends Item {
  id: string
  parent?: string
  members?: Item[]
}
export interface Node extends Item {
  id: string
  emphasis?: boolean
}
export interface FlowNode extends Node {
  column: number
  row: number
}
export interface Edge {
  from: string
  to: string
  dashed?: boolean
  arrow?: boolean
  via?: [number, number][]
}
export interface ScheduleTask {
  label: string
  start: number | string
  end?: number | string
}
export interface DiagramData {
  'org-chart': { nodes: OrgNode[] }
  flowchart: { nodes: FlowNode[]; edges: Edge[] }
  network: { nodes: Node[]; edges: Edge[] }
  gantt: { tasks: ScheduleTask[]; unit?: string }
  timeline: { steps: Item[] }
  'circular-process': { steps: Item[] }
  bars: { values: { label: string; value: number }[]; unit?: string }
  metrics: { values: { label: string; value: string }[]; columns?: number }
  overlap: { sets: [string, string] }
  footprint: {
    points: { x: number; y: number; radius: number; emphasis?: boolean }[]
    xLabels: [string, string]
    yLabels: [string, string]
  }
}
export interface DiagramTheme {
  ink: string
  paper: string
  accent: string
  onAccent: string
  muted: string
  panel: string
  rule: string
  font: string
  fontSize: number
}
export type DiagramType = keyof DiagramData
export type DiagramSpec = {
  [K in DiagramType]: {
    type: K
    title: string
    data: DiagramData[K]
    width?: number
    height?: number
    theme?: Partial<DiagramTheme>
  }
}[DiagramType]
/** Optional host font metrics and glyph outlines for identical rendering in workers, browsers and print. */
export interface DiagramFonts {
  measure: (text: string, size: number, bold: boolean) => number
  outline: (text: string, size: number, bold: boolean, x: number, y: number) => string
}
