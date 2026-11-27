// Tracked text in the layout: a run's character spacing, a named paragraph style's and a character style's.
import { describe, expect, it } from 'vitest'

import { document, paragraph, section, text, type Block, type Document } from '../../src/model'
import { writeDocx } from '../../src/docx'
import { createFontProvider, layoutDocx } from '../../src/layout'
import { diskLoader } from './sample'

const fonts = createFontProvider(diskLoader)
const page = { paper: 'letter' as const, margins: { top: 72, bottom: 72, left: 72, right: 72, header: 36, footer: 36 } }
const docxOf = (blocks: Block[], extra: Partial<Document> = {}, options: Parameters<typeof writeDocx>[1] = {}) =>
  writeDocx(document([section(blocks, { page })], extra), options)

describe('tracking', () => {
  // A centred line starts half its extra width further left.
  const startOf = async (blocks: Block[], extra: Partial<Document> = {}) =>
    (await layoutDocx(await docxOf(blocks, extra), { fonts })).pageLines[0]![0]!.x

  it('widens a run by its spacing after every character', async () => {
    const word = 'TRACKED'
    const plain = await startOf([paragraph(word, { align: 'center' })])
    const wide = await startOf([paragraph([text(word, { characterSpacing: 2 })], { align: 'center' })])
    expect(plain - wide).toBeGreaterThan(((word.length - 1) * 2) / 2 - 0.5)
    expect(plain - wide).toBeLessThan((word.length * 2) / 2 + 0.5)
  })

  it('takes a named style\'s tracking and a character style\'s', async () => {
    const word = 'DEPARTMENT'
    const styles: Document['styles'] = {
      paragraph: [{ id: 'Spaced', text: { characterSpacing: 1, align: 'center' } }],
      character: [{ id: 'Wider', text: { characterSpacing: 3 } }],
    }
    const plain = await startOf([paragraph(word, { align: 'center' })])
    const styled = await startOf([paragraph(word, { style: 'Spaced' })], { styles })
    const both = await startOf([paragraph([text(word, { characterStyle: 'Wider' })], { style: 'Spaced' })], { styles })
    expect(plain - styled).toBeGreaterThan(4)
    expect(styled - both).toBeGreaterThan(9)
  })
})
