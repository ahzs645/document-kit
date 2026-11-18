/** Test helpers: made-up pictures, and reading a written Word package back. */
import { strFromU8, unzipSync, zlibSync } from 'fflate'

const CRC_TABLE = Array.from({ length: 256 }, (_, n) => {
  let c = n
  for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
  return c >>> 0
})

function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff
  for (const byte of bytes) c = (CRC_TABLE[(c ^ byte) & 0xff] ?? 0) ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length)
  const view = new DataView(out.buffer)
  view.setUint32(0, data.length)
  for (let i = 0; i < 4; i += 1) out[4 + i] = type.charCodeAt(i)
  out.set(data, 8)
  view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)))
  return out
}

/** A valid RGB PNG of one flat colour. */
export function makePng(width: number, height: number, rgb: [number, number, number] = [200, 220, 210]): Uint8Array {
  const header = new Uint8Array(13)
  const view = new DataView(header.buffer)
  view.setUint32(0, width)
  view.setUint32(4, height)
  header.set([8, 2, 0, 0, 0], 8)
  const raw = new Uint8Array(height * (1 + width * 3))
  for (let y = 0; y < height; y += 1) {
    const row = y * (1 + width * 3)
    for (let x = 0; x < width; x += 1) raw.set(rgb, row + 1 + x * 3)
  }
  const parts = [
    Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', header),
    chunk('IDAT', zlibSync(raw)),
    chunk('IEND', new Uint8Array(0)),
  ]
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0))
  let at = 0
  for (const part of parts) {
    out.set(part, at)
    at += part.length
  }
  return out
}

/** A JPEG's header up to its frame: enough for the size to be read, with an APP0 segment before it. */
export function makeJpegHeader(width: number, height: number): Uint8Array {
  return Uint8Array.from([
    0xff, 0xd8,
    0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00,
    0xff, 0xc0, 0x00, 0x11, 0x08, height >> 8, height & 0xff, width >> 8, width & 0xff, 0x03,
    0x01, 0x22, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01,
    0xff, 0xd9,
  ])
}

export interface Package {
  files: Record<string, Uint8Array>
  text(name: string): string
  document: string
  names: string[]
}

export function readPackage(bytes: Uint8Array): Package {
  const files = unzipSync(bytes)
  const text = (name: string) => {
    const file = files[name]
    if (!file) throw new Error(`no part ${name} in the package (has ${Object.keys(files).join(', ')})`)
    return strFromU8(file)
  }
  return { files, text, document: text('word/document.xml'), names: Object.keys(files) }
}

/** Every match of a pattern's first group. */
export const all = (source: string, pattern: RegExp): string[] =>
  [...source.matchAll(new RegExp(pattern.source, pattern.flags.includes('g') ? pattern.flags : `${pattern.flags}g`))].map(
    (m) => m[1] ?? m[0],
  )

/** The body's paragraphs' plain text, in order. */
export function paragraphTexts(xml: string): string[] {
  return all(xml, /<w:p\b[^>]*>([\s\S]*?)<\/w:p>/).map((p) => all(p, /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/).join(''))
}
