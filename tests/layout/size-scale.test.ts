// A size Word holds only as a half-point with a character scale (units' fontSize): laid out as wide as the size asked,
// so a 9.3 pt line breaks where the HTML writer's does, not 2% early at 9.5 pt.
import { describe, expect, it } from 'vitest'

import { document, paragraph, section, text, type Block, type Document } from '../../src/model'
import { writeDocx } from '../../src/docx'
import { createFontProvider, layoutDocx } from '../../src/layout'
import { diskLoader } from './sample'

const fonts = createFontProvider(diskLoader)
const page = { paper: 'letter' as const, margins: { top: 72, bottom: 72, left: 72, right: 72, header: 36, footer: 36 } }
const LINE = 'Made-up words set to measure'

// A right-aligned line ends at the margin, so its start says its width.
const widthOf = async (blocks: Block[], extra: Partial<Document> = {}) =>
  540 - (await layoutDocx(await writeDocx(document([section(blocks, { page })], extra)), { fonts })).pageLines[0]![0]!.x

describe('a size between half-points', () => {
  it('sets a run as wide as its size, not the half-point it is written at', async () => {
    const at = (size: number) => widthOf([paragraph([text(LINE, { size })], { align: 'right' })])
    const asked = (await at(10)) * 9.3 / 10
    // 9.5 pt at 98%: within the 0.5% a whole percent allows, where 9.5 pt alone is 2% wider.
    expect(Math.abs((await at(9.3)) / asked - 1)).toBeLessThan(0.005)
    expect((await at(9.5)) / asked).toBeGreaterThan(1.02)
  })

  it('sets a style\'s size so, and an exact size in it at its full width', async () => {
    const theme = { body: { size: 9.3 } }
    const plain = await widthOf([paragraph([text(LINE, { size: 10 })], { align: 'right' })])
    expect(Math.abs((await widthOf([paragraph(LINE, { align: 'right' })], { theme })) / (plain * 0.93) - 1)).toBeLessThan(0.005)
    expect(await widthOf([paragraph([text(LINE, { size: 10 })], { align: 'right' })], { theme })).toBeCloseTo(plain, 2)
  })
})
