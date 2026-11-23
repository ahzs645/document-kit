# Diagram renderers

A diagram is a JSON object with a `renderer` ID and a renderer-specific `source`. A local registry resolves that object to SVG or image bytes. Use `registry.figure` for a native Document Kit figure, `registry.image` to bind a picture to a recipe, or `registry.render` to keep the artwork. The host supplies fonts and SVG rasterization for Word. The core registry has no dependency on D3, Mermaid, a browser DOM or a network connection.

```ts
import { createDiagramRegistry, mermaidAdapter } from 'document-kit/diagrams'
import { d3Adapter } from 'document-kit/diagrams/d3'

// mermaid is the host's configured, locally bundled Mermaid instance.
// svgToPng returns PNG bytes at the requested width.
const registry = createDiagramRegistry([
  d3Adapter({ accent: '#17426b', onAccent: '#ffffff' }),
  mermaidAdapter(mermaid),
  { id: 'another-engine', render: source => localEngine.render(source) }
], svgToPng)

const graph = await registry.figure({
  renderer: 'd3', source: { type: 'gantt', title: 'Delivery schedule',
    data: { unit: 'Week', tasks: [{ label: 'Build', start: 0, end: 3 }, { label: 'Launch', start: 3 }] } }
}, { alt: 'Three weeks of implementation followed by launch.', fit: { width: 450, height: 250 } })
```

Each adapter returns `{svg}` or Document Kit image data such as `{bytes: pngBytes}`. A Node host can pass an asynchronous drawing callback to `mermaidAdapter`. Adapters are registered by application code. JSON documents carry only renderer names and data.

## D3 templates

[The gallery](../examples/diagrams/gallery.json) contains runnable sources for all ten templates. D3 specifications accept `type`, `title`, `data`, and optional `width`, `height` and `theme`. Geometry uses the requested width. Height defaults to the space needed for content. Brand colours, body size and font can be set once on the adapter or overridden on a diagram. Supply a font measurement and outline provider for reproducible print typography.

| Type | Data |
|---|---|
| org-chart | Nodes with IDs, parent IDs, labels, details and optional member lists |
| gantt | Tasks with numeric offsets or ISO dates on a single axis. Omit the end for a milestone |
| flowchart | ID-labelled nodes on a column and row grid and directed connections. Optional via points use grid coordinates |
| timeline | Ordered steps with labels and optional detail |
| circular-process | A numbered loop and its step descriptions |
| network | ID-labelled nodes and edges on a deterministic radial layout |
| bars | Labelled signed values and an optional unit suffix |
| metrics | Labelled display values in a grid |
| overlap | Two named intersecting sets |
| footprint | Normalized x/y points, marker radii and direction labels |

Use `diagramTypes` to enumerate the built-in templates. A new engine implements the same adapter interface. D3's [example gallery](https://observablehq.com/@d3/gallery) provides additional layout ideas.
