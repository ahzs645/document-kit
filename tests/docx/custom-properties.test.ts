// Custom document properties, each written in its own type.
import { describe, expect, it } from 'vitest'

import { document, paragraph, section, type Document } from '../../src/model'
import { writeDocx } from '../../src/docx'
import { all, readPackage } from './helpers'

const write = async (extra: Partial<Document>) => readPackage(await writeDocx(document([section([paragraph('x')])], extra)))

describe('custom properties', () => {
  it('writes each property in its own type', async () => {
    const pkg = await write({
      metadata: {
        created: '2026-01-01T00:00:00Z',
        custom: { 'app.publicationId': 'pub-1 & co', count: 3, ratio: 0.5, draft: false, checked: new Date('2026-02-03T04:05:06Z') },
      },
    })
    const custom = pkg.text('docProps/custom.xml')
    expect(custom).toMatch(/name="app.publicationId"><vt:lpwstr>pub-1 &amp; co<\/vt:lpwstr>/)
    expect(custom).toMatch(/name="count"><vt:i4>3<\/vt:i4>/)
    expect(custom).toMatch(/name="ratio"><vt:r8>0.5<\/vt:r8>/)
    expect(custom).toMatch(/name="draft"><vt:bool>false<\/vt:bool>/)
    expect(custom).toMatch(/name="checked"><vt:filetime>2026-02-03T04:05:06Z<\/vt:filetime>/)
    expect(all(custom, /pid="(\d+)"/)).toEqual(['2', '3', '4', '5', '6'])
    expect(pkg.text('_rels/.rels')).toMatch(/docProps\/custom\.xml/)
    expect(pkg.text('[Content_Types].xml')).toMatch(/\/docProps\/custom\.xml/)
  })

  it('leaves the package\'s own empty properties part when there are none', async () => {
    const pkg = await write({})
    expect(pkg.text('docProps/custom.xml')).not.toMatch(/<property/)
  })

  it('refuses a value it cannot write', async () => {
    await expect(write({ metadata: { custom: { bad: Number.NaN } } })).rejects.toThrow(/"bad"/)
    await expect(write({ metadata: { custom: { when: new Date('nope') } } })).rejects.toThrow(/"when"/)
  })
})
