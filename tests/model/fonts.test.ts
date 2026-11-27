// What a font file is and what its licence flags let each writer do with it.
import { describe, expect, it } from 'vitest'

import { fontEmbeddingReport, fontFileInfo, fontSlot, groupFonts } from '../../src/model'
import { cidKeyedCff, otf, ttf, withFsType } from '../fonts'

describe('fontFileInfo', () => {
  it('tells TrueType, CFF and CID-keyed CFF outlines apart and reads the embedding permission', () => {
    expect(fontFileInfo(ttf())).toMatchObject({ format: 'truetype', fsType: 0, level: 'installable', bitmapOnly: false })
    expect(fontFileInfo(otf())).toMatchObject({ format: 'cff', fsType: 8, level: 'editable' })
    expect(fontFileInfo(cidKeyedCff()).format).toBe('cff-cid')
    expect(fontFileInfo(new Uint8Array([1, 2, 3, 4, 5])).format).toBe('unknown')
    expect(fontFileInfo(ttf().buffer as ArrayBuffer).format).toBe('truetype')
  })

  it('reads each permission level, and bitmap-only', () => {
    expect(fontFileInfo(withFsType(ttf(), 0x0002)).level).toBe('restricted')
    expect(fontFileInfo(withFsType(ttf(), 0x0004)).level).toBe('previewPrint')
    expect(fontFileInfo(withFsType(ttf(), 0x0008)).level).toBe('editable')
    expect(fontFileInfo(withFsType(ttf(), 0x0208)).bitmapOnly).toBe(true)
  })
})

describe('font slots', () => {
  it('puts 600 and heavier in the bold slot and italic in the italic ones', () => {
    expect(fontSlot({})).toBe('regular')
    expect(fontSlot({ weight: 500 })).toBe('regular')
    expect(fontSlot({ weight: 600 })).toBe('bold')
    expect(fontSlot({ weight: 700, style: 'italic' })).toBe('boldItalic')
    expect(fontSlot({ style: 'italic' })).toBe('italic')
  })

  it('gives a contested slot to the face nearest its weight, case-insensitively by family', () => {
    const light = { family: 'Brand', weight: 300, bytes: ttf() }
    const book = { family: 'brand ', weight: 400, bytes: ttf() }
    const { families, extra } = groupFonts([light, book])
    expect(families.get('brand')?.faces.regular).toBe(book)
    expect(extra).toEqual([light])
  })
})

describe('fontEmbeddingReport', () => {
  it('embeds TrueType installable faces in Word and refuses CFF there, with reasons', () => {
    const report = fontEmbeddingReport(
      [
        { family: 'Brand', bytes: ttf() },
        { family: 'Brand', weight: 700, bytes: ttf('Bold') },
        { family: 'Code', bytes: otf() },
      ],
      'docx',
    )
    expect(report.map((r) => [r.family, r.slot, r.embedded])).toEqual([
      ['Brand', 'regular', true],
      ['Brand', 'bold', true],
      ['Code', 'regular', false],
    ])
    expect(report[2]!.reason).toMatch(/Word embeds TrueType outlines only/)
    expect(report[2]!.format).toBe('cff')
  })

  it('lets an HTML print view carry a CFF face and a preview & print face, but never a restricted one', () => {
    const preview = withFsType(ttf(), 0x0004)
    const restricted = withFsType(ttf(), 0x0002)
    const docx = fontEmbeddingReport([{ family: 'P', bytes: preview }, { family: 'R', bytes: restricted }], 'docx')
    const html = fontEmbeddingReport(
      [{ family: 'P', bytes: preview }, { family: 'R', bytes: restricted }, { family: 'C', bytes: otf() }],
      'html',
    )
    expect(docx.map((r) => r.embedded)).toEqual([false, false])
    expect(docx[0]!.reason).toMatch(/preview & print/)
    expect(docx[1]!.reason).toMatch(/forbids embedding/)
    expect(html.map((r) => r.embedded)).toEqual([true, false, true])
  })

  it('refuses a family that has only an italic face in Word, which needs a regular or bold one', () => {
    const report = fontEmbeddingReport([{ family: 'Slant', style: 'italic', bytes: ttf('Italic') }], 'docx')
    expect(report[0]).toMatchObject({ embedded: false })
    expect(report[0]!.reason).toMatch(/no regular or bold face/)
  })

  it('names the face that lost its slot to another', () => {
    const report = fontEmbeddingReport(
      [
        { family: 'Brand', weight: 400, bytes: ttf() },
        { family: 'Brand', weight: 450, bytes: ttf() },
      ],
      'docx',
    )
    expect(report[1]).toMatchObject({ embedded: false })
    expect(report[1]!.reason).toMatch(/fills its slot/)
  })
})
