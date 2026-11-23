import { describe, expect, it } from 'vitest'
import { writeDocx } from '../../src/docx'
import { compileRecipe, renderRecipe, type Json } from '../../src/recipe'
import { readPackage } from '../docx/helpers'
import { comp, lib, recipe } from './helpers'

const blocks: Json[] = [
  { for: 'item in items', do: [{ type: 'paragraph', content: { type: 'text', link: '#{{ item.id }}', text: '{{ item.name }}' } }] },
  { for: 'item in items', do: [{ if: 'loop.odd',
    then: [{ type: 'heading', level: 2, content: '{{ item.name }}', bookmark: '{{ item.id }}' }],
    else: [{ type: 'paragraph', content: '{{ item.name }}', bookmark: '{{ item.id }}' }],
  }] },
]
const compiled = compileRecipe(recipe('r', [{ use: 'c' }]), lib([comp('c', blocks)]))
const render = (ids: string[]) => writeDocx(renderRecipe(compiled, { items: ids.map(id => ({ id, name: 'Section ' + id })) }).document)

describe('document navigation', () => {
  it('links to heading and paragraph bookmarks after reordering records', async () => {
    const pkg = readPackage(await render(['Second', 'First']))
    expect(pkg.document).toContain('w:anchor="First"')
    expect(pkg.document).toContain('w:anchor="Second"')
    expect(pkg.document.indexOf('w:name="Second"')).toBeLessThan(pkg.document.indexOf('w:name="First"'))
    expect(pkg.text('word/_rels/document.xml.rels')).not.toContain('Target="#')
  })
  it('rejects duplicate and invalid Word bookmark names', async () => {
    await expect(render(['Repeat', 'repeat'])).rejects.toThrow('Duplicate Word bookmark')
    await expect(render(['bad:name'])).rejects.toThrow('Invalid Word bookmark')
  })
})
