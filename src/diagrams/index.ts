import type { FigureBlock, ImageData } from '../model/types'
import { imageInfo } from '../model/image'

export interface DiagramSource {
  renderer: string
  source: unknown
}
export type DiagramArtwork = { svg: string } | ImageData
export interface DiagramAdapter {
  id: string
  render: (source: unknown) => DiagramArtwork | Promise<DiagramArtwork>
}
export type Rasterize = (svg: string, width: number) => Uint8Array | Promise<Uint8Array>

/** A registry belongs to its consumer. Adapters can capture a theme, worker, or font provider. */
export function createDiagramRegistry(adapters: readonly DiagramAdapter[], rasterize?: Rasterize) {
  const byId = new Map<string, DiagramAdapter>()
  for (const adapter of adapters) {
    if (!/^[a-z][a-z0-9-]*$/.test(adapter.id) || byId.has(adapter.id))
      throw new Error('Invalid or duplicate diagram renderer: ' + adapter.id)
    byId.set(adapter.id, adapter)
  }
  async function render(spec: DiagramSource) {
    const adapter = byId.get(spec.renderer)
    if (!adapter) throw new Error('Unknown diagram renderer: ' + spec.renderer)
    return adapter.render(spec.source)
  }
  async function image(spec: DiagramSource, pixelWidth = 2400): Promise<ImageData> {
    const art = await render(spec)
    if ('svg' in art && !rasterize) throw new Error('SVG diagram output needs a rasterizer for Word.')
    const result = 'svg' in art ? { bytes: await rasterize!(art.svg, pixelWidth) } : art
    if (!imageInfo(result.bytes))
      throw new Error('Diagram renderer returned an unsupported image: ' + spec.renderer)
    return result
  }
  async function figure(
    spec: DiagramSource,
    options: Omit<FigureBlock, 'type' | 'image'> = {},
  ): Promise<FigureBlock> {
    return { ...options, type: 'figure', image: await image(spec) }
  }
  return { render, image, figure }
}

/** Use an initialized Mermaid instance in a browser, or supply a host function for Node or a worker. */
export function mermaidAdapter(
  engine:
    | { render(id: string, source: string): Promise<{ svg: string }> }
    | ((definition: string) => Promise<DiagramArtwork>),
): DiagramAdapter {
  return {
    id: 'mermaid',
    async render(source) {
      if (typeof source !== 'string' || !source.trim()) throw new Error('Expected Mermaid diagram text.')
      if (typeof engine === 'function') return engine(source)
      const result = await engine.render('diagram-' + crypto.randomUUID(), source)
      return { svg: result.svg }
    },
  }
}
