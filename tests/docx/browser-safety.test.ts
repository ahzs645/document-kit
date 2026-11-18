import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { writeDocx } from '../../src/docx/index'
import { laboratoryReport } from './fixtures'
import { readPackage } from './helpers'

const root = join(import.meta.dirname, '..', '..', 'src')
const files = (dir: string): string[] =>
  readdirSync(dir, { withFileTypes: true }).flatMap((entry) =>
    entry.isDirectory() ? files(join(dir, entry.name)) : entry.name.endsWith('.ts') ? [join(dir, entry.name)] : [],
  )

describe('browser safety', () => {
  it('keeps Node and the DOM out of the model and the Word writer', () => {
    const sources = [...files(join(root, 'model')), ...files(join(root, 'docx')), join(root, 'index.ts')]
    expect(sources.length).toBeGreaterThan(10)
    for (const file of sources) {
      const code = readFileSync(file, 'utf8')
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .replace(/\/\/.*$/gm, '')
      const found = /\bfrom ['"]node:|\brequire\(|\bBuffer\b|\bprocess\.|\bwindow\b|\bglobalThis\.document\b|\bdocument\.(createElement|body|querySelector)|\bnew Image\(|\bOffscreenCanvas\b|\bfetch\(/.exec(code)
      expect([file, found?.[0] ?? null]).toEqual([file, null])
    }
  })

  it('writes a file with no Buffer in the global scope', async () => {
    const saved = globalThis.Buffer
    let bytes: Uint8Array
    try {
      // @ts-expect-error removed to stand in for a browser
      delete globalThis.Buffer
      bytes = await writeDocx(laboratoryReport())
    } finally {
      globalThis.Buffer = saved
    }
    expect(bytes).toBeInstanceOf(Uint8Array)
    expect(readPackage(bytes).document).toContain('Certificate of Analysis')
  })
})
