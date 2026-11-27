// Test fonts read from the suite's own fixtures (open licences, see tests/ream/fixtures/fonts/NOTICE.md), and font
// files edited in memory to carry other licence flags or a CID-keyed CFF table.
import { readFileSync } from 'node:fs'

/** The bytes of a font in `tests/ream/fixtures/fonts`. */
export function testFont(file: string): Uint8Array {
  return new Uint8Array(readFileSync(new URL(`./ream/fixtures/fonts/${file}`, import.meta.url)))
}

/** A TrueType face (Roboto, Apache-2.0, installable). */
export const ttf = (style: 'Regular' | 'Bold' | 'Italic' | 'BoldItalic' = 'Regular') => testFont(`Roboto-${style}.ttf`)

/** An OpenType face with CFF outlines (Inconsolata, OFL-1.1, editable embedding). */
export const otf = () => testFont('Inconsolata-Regular.otf')

/** A copy of a font with its OS/2 `fsType` set to `value`. */
export function withFsType(bytes: Uint8Array, value: number): Uint8Array {
  const out = bytes.slice()
  const view = new DataView(out.buffer)
  const count = view.getUint16(4)
  for (let i = 0; i < count; i += 1) {
    const at = 12 + i * 16
    if (String.fromCharCode(...out.subarray(at, at + 4)) !== 'OS/2') continue
    view.setUint16(view.getUint32(at + 8) + 8, value)
    return out
  }
  throw new Error('no OS/2 table')
}

/**
 * The smallest OpenType file whose CFF Top DICT carries the ROS operator, as a CID-keyed font's does. It holds nothing
 * else of use: enough to be told apart, not to draw with.
 */
export function cidKeyedCff(): Uint8Array {
  // CFF: header (4), Name INDEX with one name "X", Top DICT INDEX with one DICT: ROS = 0 0 0 (12 30).
  const name = [0x00, 0x01, 0x01, 0x01, 0x02, 0x58]
  const dict = [0x8b, 0x8b, 0x8b, 0x0c, 0x1e]
  const top = [0x00, 0x01, 0x01, 0x01, 1 + dict.length, ...dict]
  const cff = Uint8Array.from([0x01, 0x00, 0x04, 0x01, ...name, ...top])
  const os2 = new Uint8Array(78)
  new DataView(os2.buffer).setUint16(0, 4)
  const tables: Array<[string, Uint8Array]> = [
    ['CFF ', cff],
    ['OS/2', os2],
  ]
  const headerSize = 12 + tables.length * 16
  let offset = headerSize
  const placed = tables.map(([tag, data]) => {
    const at = offset
    offset += (data.length + 3) & ~3
    return { tag, data, at }
  })
  const out = new Uint8Array(offset)
  const view = new DataView(out.buffer)
  out.set([0x4f, 0x54, 0x54, 0x4f], 0)
  view.setUint16(4, tables.length)
  placed.forEach(({ tag, data, at }, i) => {
    const entry = 12 + i * 16
    for (let k = 0; k < 4; k += 1) out[entry + k] = tag.charCodeAt(k)
    view.setUint32(entry + 8, at)
    view.setUint32(entry + 12, data.length)
    out.set(data, at)
  })
  return out
}
