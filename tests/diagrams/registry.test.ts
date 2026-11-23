import { describe, expect, it } from 'vitest'
import { createDiagramRegistry, mermaidAdapter } from '../../src/diagrams'
import { d3Adapter, diagramTypes, renderD3, type DiagramSpec } from '../../src/diagrams/d3'
import gallery from '../../examples/diagrams/gallery.json'

describe('diagram adapters', () => {
  it('renders the complete example library', async () => {
    const registry = createDiagramRegistry([d3Adapter()])
    expect(gallery.map((s) => s.source.type).sort()).toEqual([...diagramTypes].sort())
    for (const spec of gallery) {
      const art = await registry.render(spec)
      expect(art).toHaveProperty('svg')
      if ('svg' in art) {
        expect(art.svg).toContain(spec.source.title)
        expect(art.svg).not.toMatch(/NaN|Infinity/)
      }
    }
  })
  it('sizes every metric card for the longest wrapped content across rows', () => {
    const svg = renderD3({
      type: 'metrics',
      title: 'Review',
      width: 400,
      data: { columns: 2, values: [
        { label: 'Items', value: '8' }, { label: 'Drafts', value: '5' },
        { label: 'Items ready for review by the whole team', value: '3' },
      ] },
    })
    const cards = [...svg.matchAll(/<rect x="[^"]+" y="[^"]+" width="([^"]+)" height="([^"]+)"/g)]
    expect(cards).toHaveLength(3)
    expect(new Set(cards.map((card) => card.slice(1).join('x'))).size).toBe(1)
  })

  it('uses the same explicit viewport for SVG, browser canvas and native rasterizers', () => {
    const svg = renderD3({
      type: 'bars',
      title: 'Hours',
      width: 600,
      height: 120,
      data: { values: [{ label: 'Review', value: 0 }] },
    })
    expect(svg).toContain('width="600" height="120" viewBox="0 0 600 120"')
    expect(svg).not.toMatch(/NaN|Infinity/)
  })

  it('selects Mermaid and independent renderers and rejects duplicate or unknown names', async () => {
    const external = { id: 'external', render: () => ({ svg: '<svg/>' }) }
    const registry = createDiagramRegistry([
      external,
      mermaidAdapter({ render: async (_id, s) => ({ svg: s }) }),
    ])
    expect(await registry.render({ renderer: 'mermaid', source: 'graph TD' })).toEqual({ svg: 'graph TD' })
    expect(await registry.render({ renderer: 'external', source: null })).toEqual({ svg: '<svg/>' })
    expect(() => createDiagramRegistry([external, external])).toThrow(/duplicate/)
    await expect(registry.render({ renderer: 'missing', source: null })).rejects.toThrow(/Unknown/)
    await expect(registry.figure({ renderer: 'external', source: null })).rejects.toThrow(/rasterizer/)
  })
  it('validates references, dates and theme values before drawing', () => {
    expect(() =>
      renderD3({
        type: 'gantt',
        title: 'Schedule',
        data: { tasks: [{ label: 'Test', start: '2025-02-30' }] },
      }),
    ).toThrow(/calendar date/)
    expect(() =>
      renderD3({
        type: 'flowchart',
        title: 'Flow',
        data: { nodes: [{ id: 'a', label: 'One', row: 0, column: 0 }], edges: [{ from: 'a', to: 'b' }] },
      }),
    ).toThrow(/unknown node/)
    expect(() => renderD3(gallery[0]!.source as DiagramSpec, { accent: 'url(external)' })).toThrow(/colour/)
  })
})
