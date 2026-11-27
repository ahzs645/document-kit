// The PDF's text state after a picture: ET ends a text object, not the text state, so a run's character spacing
// (Tc) and horizontal scale (Tz) stand until set again (ISO 32000-1 §9.3.1). A plain run after a picture must set
// them back itself, or it keeps the tracking and scale of the run before the picture.
import { describe, expect, it } from 'vitest'

import { document, inlineImage, paragraph, pinned, section, text } from '../../src/model'
import { writeDocx } from '../../src/docx'
import { createFontProvider, layoutDocx } from '../../src/layout'
import { diskLoader } from './sample'
import { makePng } from '../docx/helpers'
import { pageDrawings } from './pdf'

const fonts = createFontProvider(diskLoader)
const page = { paper: 'letter' as const, margins: { top: 72, bottom: 72, left: 72, right: 72, header: 36, footer: 36 } }

describe('text state after a picture', () => {
  it('sets a plain run\'s spacing and scale again after a tracked, scaled run and a picture', async () => {
    const rule = inlineImage({ bytes: makePng(40, 1), format: 'png' }, 200, 1, '')
    const docx = await writeDocx(document([section([
      paragraph([text('TRACKED LABEL', { characterSpacing: 2, size: 5.7 })]),
      // A foot of the page: a rule drawn as a picture, then plain words, in one frame.
      pinned([paragraph([rule]), paragraph([text('Plain words', { size: 10 })])]),
    ], { page })]))
    const drawing = pageDrawings((await layoutDocx(docx, { fonts })).toPdf().bytes)[0]!
    const spacing = [...drawing.matchAll(/(-?[\d.]+) Tc\b/g)].map((m) => Number(m[1]))
    const scale = [...drawing.matchAll(/(-?[\d.]+) Tz\b/g)].map((m) => Number(m[1]))
    // The label: 5.7 pt is 5.5 pt at 104%, its 2 pt of spacing stated in that scaled space (2 / 1.04). The words,
    // drawn last, state 0 and 100 again.
    expect(scale).toContain(104)
    expect(spacing).toContain(1.923)
    expect(spacing.at(-1)).toBe(0)
    expect(scale.at(-1)).toBe(100)
  })
})
