import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const root = join(import.meta.dirname, '..', '..', 'src', 'recipe')

describe('browser safety', () => {
  it('keeps Node, the DOM and the network out of the recipe system', () => {
    const sources = readdirSync(root).filter((name) => name.endsWith('.ts'))
    expect(sources.length).toBeGreaterThan(10)
    for (const name of sources) {
      const code = readFileSync(join(root, name), 'utf8')
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .replace(/\/\/.*$/gm, '')
      const found = /\bfrom ['"]node:|\brequire\(|\bBuffer\b|\bprocess\.|\bwindow\b|\bglobalThis\.document\b|\bdocument\.(createElement|body|querySelector)|\bnew Image\(|\bOffscreenCanvas\b|\bfetch\(|\bXMLHttpRequest\b/.exec(code)
      expect([name, found?.[0] ?? null]).toEqual([name, null])
    }
  })
})
