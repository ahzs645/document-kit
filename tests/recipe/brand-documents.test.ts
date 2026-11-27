// A designed document from a recipe: named styles, tracked runs, a frame and a picture behind the text, custom
// properties, all from JSON, and the same document written as Word and as HTML.
import { describe, expect, it } from 'vitest'

import { writeDocx } from '../../src/docx'
import { writeHtml } from '../../src/html'
import { compileRecipe, renderRecipe } from '../../src/recipe'
import { makePng, readPackage } from '../docx/helpers'
import { comp, lib, recipe } from './helpers'

const cover = comp('cover', [
  { type: 'floatingImage', image: '{{ artwork }}', box: 'page', alt: 'Cover artwork' },
  {
    type: 'frame',
    x: 54,
    y: '{{ titleTop }}',
    width: 400,
    wrap: 'through',
    blocks: [{ type: 'paragraph', style: 'CoverTitle', content: [{ type: 'text', text: '{{ title }}', characterSpacing: 1.5 }] }],
  },
])

const compiled = () =>
  compileRecipe(
    recipe('brand', [{ use: 'cover', id: 'cover' }], {
      document: {
        metadata: { title: '{{ title }}', custom: { 'app.publicationId': '{{ id }}' } },
        styles: { paragraph: [{ id: 'CoverTitle', text: { size: 28, color: '{{ ink }}' } }] },
      },
    }),
    lib([cover]),
  )

describe('a designed document from a recipe', () => {
  it('renders styles, frames, floating pictures and custom properties from the data', async () => {
    const rendered = renderRecipe(compiled(), { artwork: makePng(51, 66), title: 'Made-up title', titleTop: 120, id: 'pub-9', ink: '#123456' })
    const doc = rendered.document
    expect(doc.styles).toEqual({ paragraph: [{ id: 'CoverTitle', text: { size: 28, color: '#123456' } }] })
    expect(doc.metadata?.custom).toEqual({ 'app.publicationId': 'pub-9' })
    expect(doc.sections[0]!.body[1]).toMatchObject({ type: 'frame', x: 54, y: 120, width: 400 })
    const pkg = readPackage(await writeDocx(doc))
    expect(pkg.document).toMatch(/<w:framePr[^>]*w:y="2400"/)
    expect(pkg.document).toMatch(/behindDoc="1"/)
    expect(pkg.document).toMatch(/<w:pStyle w:val="CoverTitle"\/>/)
    expect(pkg.document).toMatch(/<w:spacing w:val="30"\/>/)
    expect(pkg.text('docProps/custom.xml')).toMatch(/name="app.publicationId"><vt:lpwstr>pub-9</)
    const html = writeHtml(doc)
    expect(html).toMatch(/\.dk-ps-CoverTitle\{[^}]*color:#123456/)
    expect(html).toMatch(/<div class="dk-frame" style="left:54pt;top:120pt/)
  })

  it('places no picture, and lists it, when the data has none', () => {
    const rendered = renderRecipe(compiled(), { artwork: null, title: 'T', titleTop: 100, id: 'x', ink: '000' })
    expect(rendered.document.sections[0]!.body[0]).toMatchObject({ type: 'floatingImage', image: null })
    expect(rendered.toComplete).toEqual([expect.objectContaining({ kind: 'figure', what: 'Cover artwork' })])
  })

  it('refuses a key the frame does not have', () => {
    expect(() => compileRecipe(recipe('r', [{ use: 'c' }]), lib([comp('c', [{ type: 'frame', blocks: [], x: 0, y: 0, width: 9, top: 1 }])]))).toThrow(/top/)
  })
})
