// Fonts embedded in the Word file by their own licence flags, obfuscated, with a key made from their bytes.
import { describe, expect, it } from 'vitest'

import { document, paragraph, section } from '../../src/model'
import { fontEmbeddingReport, obfuscateFont, writeDocx } from '../../src/docx'
import { all, readPackage } from './helpers'
import { otf, ttf, withFsType } from '../fonts'

describe('embedded fonts', () => {
  const fonts = [
    { family: 'Brand Sans', bytes: ttf() },
    { family: 'Brand Sans', weight: 700, bytes: ttf('Bold') },
    { family: 'Brand Sans', style: 'italic' as const, bytes: ttf('Italic') },
    { family: 'Brand Mono', bytes: otf() },
    { family: 'Locked', bytes: withFsType(ttf(), 2) },
  ]
  const doc = document([section([paragraph('x')])], { metadata: { created: '2026-01-01T00:00:00Z' } })

  it('embeds the TrueType faces the licence allows, obfuscated, and leaves out the rest', async () => {
    const pkg = readPackage(await writeDocx(doc, { embedFonts: fonts }))
    const table = pkg.text('word/fontTable.xml')
    expect(all(table, /<w:font w:name="([^"]+)"/)).toEqual(['Brand Sans'])
    expect(all(table, /<w:(embed\w+) /)).toEqual(['embedRegular', 'embedBold', 'embedItalic'])
    expect(pkg.text('word/settings.xml')).toMatch(/<w:embedTrueTypeFonts\/>/)
    const rels = pkg.text('word/_rels/fontTable.xml.rels')
    const key = /w:embedBold r:id="([^"]+)" w:fontKey="\{([^}]+)\}"/.exec(table)!
    const target = new RegExp(`<Relationship [^>]*Id="${key[1]}"[^>]*>`).exec(rels)![0].match(/Target="([^"]+)"/)![1]!
    const file = pkg.files[`word/${target}`]!
    expect(obfuscateFont(file, key[2]!)).toEqual(ttf('Bold'))
    expect(file.subarray(0, 32)).not.toEqual(ttf('Bold').subarray(0, 32))
    // What was left out, and why, is the report's to say.
    const report = fontEmbeddingReport(fonts, 'docx')
    expect(report.filter((face) => !face.embedded).map((face) => face.family)).toEqual(['Brand Mono', 'Locked'])
  })

  it('writes the same bytes for the same fonts', async () => {
    const a = await writeDocx(doc, { embedFonts: fonts })
    const b = await writeDocx(doc, { embedFonts: fonts })
    expect(a).toEqual(b)
  })

  it('writes no font table entries when nothing can be embedded', async () => {
    const pkg = readPackage(await writeDocx(doc, { embedFonts: [{ family: 'Brand Mono', bytes: otf() }] }))
    expect(pkg.text('word/fontTable.xml')).not.toMatch(/<w:embed/)
  })
})
