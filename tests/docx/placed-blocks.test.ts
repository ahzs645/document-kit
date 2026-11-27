// Frames placed anywhere on the page and pictures floating behind or in front of the text. Made-up text.
import { describe, expect, it } from 'vitest'

import { document, floatingImage, frame, heading, paragraph, pinned, row, section, table, type Block } from '../../src/model'
import { writeDocx } from '../../src/docx'
import { all, makePng, readPackage } from './helpers'

const write = async (blocks: Block[]) => readPackage(await writeDocx(document([section(blocks)])))
const paragraphsOf = (xml: string) => all(xml, /(<w:p\b[^>]*>[\s\S]*?<\/w:p>)/)

describe('frames placed on the page', () => {
  it('gives every paragraph of the frame one frame at x, y against the page', async () => {
    const { document: xml } = await write([
      paragraph('Flow.'),
      frame([heading(1, 'Cover title'), paragraph('Subtitle')], { x: 54, y: 120, width: 300 }, { wrap: 'through' }),
    ])
    const framed = paragraphsOf(xml).filter((p) => p.includes('<w:framePr'))
    expect(framed).toHaveLength(2)
    const frames = framed.map((p) => /<w:framePr[^>]*\/>/.exec(p)?.[0])
    expect(new Set(frames).size).toBe(1)
    expect(frames[0]).toMatch(/w:x="1080"/)
    expect(frames[0]).toMatch(/w:y="2400"/)
    expect(frames[0]).toMatch(/w:w="6000"/)
    expect(frames[0]).toMatch(/w:hAnchor="page"/)
    expect(frames[0]).toMatch(/w:vAnchor="page"/)
    expect(frames[0]).toMatch(/w:wrap="through"/)
  })

  it('measures from the margins, with a fixed height and room round it, when asked', async () => {
    const { document: xml } = await write([
      frame([paragraph('Boxed')], { x: 0, y: 36, width: 144, height: 72 }, { relativeTo: 'margin', heightRule: 'exact', space: 6 }),
    ])
    const framePr = /<w:framePr[^>]*\/>/.exec(xml)?.[0] ?? ''
    expect(framePr).toMatch(/w:hAnchor="margin"/)
    expect(framePr).toMatch(/w:h="1440"/)
    expect(framePr).toMatch(/w:hRule="exact"/)
    expect(framePr).toMatch(/w:hSpace="120"/)
    expect(framePr).toMatch(/w:wrap="around"/)
  })

  it('refuses a table, as a pinned block does', async () => {
    await expect(write([frame([table([row(['a'])])], { x: 0, y: 0, width: 100 })])).rejects.toThrow(/frame cannot hold a table/)
  })
})

describe('floating pictures', () => {
  const art = { bytes: makePng(51, 66) }

  it('anchors a full-page picture behind the text at the page\'s corner', async () => {
    const pkg = await write([floatingImage(art, 'page', { alt: 'Cover artwork' }), paragraph('Over it.')])
    const anchor = /<wp:anchor\b[^>]*>[\s\S]*?<\/wp:anchor>/.exec(pkg.document)?.[0] ?? ''
    expect(anchor).toMatch(/behindDoc="1"/)
    expect(anchor).toMatch(/<wp:positionH relativeFrom="page"><wp:posOffset>0<\/wp:posOffset>/)
    expect(anchor).toMatch(/<wp:extent cx="7772400" cy="10058400"\/>/)
    expect(anchor).toMatch(/<wp:wrapNone\/>/)
    expect(anchor).toMatch(/descr="Cover artwork"/)
    // Its paragraph takes no room in the flow.
    expect(pkg.document).toMatch(/<w:spacing w:after="0" w:before="0" w:line="1" w:lineRule="exact"\/>/)
  })

  it('places a picture in front, against the margins, and in a header', async () => {
    const bytes = await writeDocx(
      document([
        section([floatingImage(art, { x: 36, y: 18, width: 72, height: 93 }, { relativeTo: 'margin', behind: false })], {
          header: [floatingImage(art, 'page')],
        }),
      ]),
    )
    const pkg = readPackage(bytes)
    const anchor = /<wp:anchor\b[^>]*>/.exec(pkg.document)?.[0] ?? ''
    expect(anchor).toMatch(/behindDoc="0"/)
    expect(pkg.document).toMatch(/<wp:positionH relativeFrom="margin"><wp:posOffset>457200<\/wp:posOffset>/)
    expect(pkg.document).toMatch(/<wp:positionV relativeFrom="margin"><wp:posOffset>228600<\/wp:posOffset>/)
    const header = pkg.names.find((name) => /^word\/header\d+\.xml$/.test(name))!
    expect(pkg.text(header)).toMatch(/behindDoc="1"/)
  })

  it('places nothing for a missing picture', async () => {
    const pkg = await write([floatingImage(null, 'page'), paragraph('Text.')])
    expect(pkg.document).not.toMatch(/<wp:anchor/)
  })

  it('refuses a floating picture inside a frame', async () => {
    await expect(write([pinned([floatingImage(art, 'page')])])).rejects.toThrow(/cannot hold a floatingImage/)
  })
})
