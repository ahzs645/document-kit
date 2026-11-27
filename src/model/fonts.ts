/**
 * Fonts a caller hands in by family name, and what their files allow: the same faces the Word writer embeds, the
 * layout draws with and the HTML writer loads, so the three outputs set text in one face. Reads the font file's
 * header and its OS/2 table only; nothing here draws or fetches.
 */

/** One face of a font family, as bytes, named as the document names it. */
export interface NamedFont {
  /** The family name runs and styles use (`font: 'Brand Sans'`). Matched without regard to case. */
  family: string
  /** The font file: TrueType (`.ttf`) or OpenType (`.otf`, CFF outlines). */
  bytes: Uint8Array | ArrayBuffer
  /**
   * The face's weight, 100 to 900; 400 when absent. Word knows two weights a family: 600 and over is the family's
   * bold face, under 600 its regular one. A light or black cut is its own family (`Brand Sans Light`).
   */
  weight?: number
  /** `italic` for the family's italic (or bold italic) face. */
  style?: 'normal' | 'italic'
}

/** One of a family's four faces, as Word and the layout know them. */
export type FontSlot = 'regular' | 'bold' | 'italic' | 'boldItalic'

/** What a font file's outlines are. */
export type FontFormat = 'truetype' | 'cff' | 'cff-cid' | 'collection' | 'unknown'

/** The OS/2 table's embedding permission (`fsType`, bits 0–3). */
export type EmbeddingLevel = 'installable' | 'editable' | 'previewPrint' | 'restricted' | 'unknown'

/** What a font file is and what its licence flags let a document do with it. */
export interface FontFileInfo {
  readonly format: FontFormat
  /** The OS/2 `fsType`; absent when the file has no OS/2 table. */
  readonly fsType?: number
  readonly level: EmbeddingLevel
  /** `fsType` bit 9: only bitmaps may be embedded (no outlines). */
  readonly bitmapOnly: boolean
}

/** Where a face is to be embedded: a Word file (editable) or an HTML print view. */
export type EmbeddingTarget = 'docx' | 'html'

/** One face of a {@link fontEmbeddingReport}. */
export interface FontEmbedding {
  readonly family: string
  readonly weight: number
  readonly style: 'normal' | 'italic'
  /** The slot the face fills in its family. */
  readonly slot: FontSlot
  readonly format: FontFormat
  readonly fsType?: number
  /** Whether the writer embeds it. */
  readonly embedded: boolean
  /** Why, in a sentence: what was embedded, or why not and what to do instead. */
  readonly reason: string
}

const asBytes = (value: Uint8Array | ArrayBuffer): Uint8Array =>
  value instanceof Uint8Array ? value : new Uint8Array(value)

const tag = (bytes: Uint8Array, at: number): string => String.fromCharCode(...bytes.subarray(at, at + 4))

/** The offset and length of one table of an sfnt font, or undefined. */
function table(bytes: Uint8Array, name: string): { offset: number; length: number } | undefined {
  if (bytes.length < 12) return undefined
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const count = view.getUint16(4)
  for (let i = 0; i < count; i += 1) {
    const at = 12 + i * 16
    if (at + 16 > bytes.length) return undefined
    if (tag(bytes, at) !== name) continue
    const offset = view.getUint32(at + 8)
    const length = view.getUint32(at + 12)
    return offset + length <= bytes.length ? { offset, length } : undefined
  }
  return undefined
}

/** Whether a CFF table's first Top DICT says it is CID-keyed (the ROS operator, 12 30). */
function cidKeyed(bytes: Uint8Array, cff: { offset: number; length: number }): boolean {
  const data = bytes.subarray(cff.offset, cff.offset + cff.length)
  if (data.length < 4) return false
  const headerSize = data[2]!
  // An INDEX: count (2), offSize (1), offsets, then the data. Returns where its data starts and ends.
  const index = (at: number): { items: Array<[number, number]>; end: number } | undefined => {
    if (at + 2 > data.length) return undefined
    const count = (data[at]! << 8) | data[at + 1]!
    if (count === 0) return { items: [], end: at + 2 }
    const offSize = data[at + 2]!
    if (offSize < 1 || offSize > 4) return undefined
    const offsetAt = (i: number) => {
      let value = 0
      for (let k = 0; k < offSize; k += 1) value = value * 256 + data[at + 3 + i * offSize + k]!
      return value
    }
    const base = at + 3 + (count + 1) * offSize - 1
    const items: Array<[number, number]> = []
    for (let i = 0; i < count; i += 1) items.push([base + offsetAt(i), base + offsetAt(i + 1)])
    return { items, end: base + offsetAt(count) }
  }
  const names = index(headerSize)
  const top = names ? index(names.end) : undefined
  const first = top?.items[0]
  if (!first) return false
  // Walk the DICT's operands and operators; ROS must be the first operator of a CID font, but any place is read.
  for (let i = first[0]; i < first[1] && i < data.length; ) {
    const b = data[i]!
    if (b === 12) {
      if (data[i + 1] === 30) return true
      i += 2
    } else if (b <= 21) i += 1
    else if (b === 28) i += 3
    else if (b === 29) i += 5
    else if (b === 30) {
      // A real number: nibbles until one is 0xf.
      i += 1
      while (i < first[1] && (data[i]! & 0x0f) !== 0x0f && (data[i]! >> 4) !== 0x0f) i += 1
      i += 1
    } else if (b >= 32 && b <= 246) i += 1
    else if (b >= 247 && b <= 254) i += 2
    else i += 1
  }
  return false
}

/**
 * What a font file is: TrueType or CFF outlines (and whether a CFF font is CID-keyed), a collection, or something
 * else; and its OS/2 embedding permission (OpenType spec, OS/2 `fsType`). For a table of version 0 to 2 several
 * permission bits may be set, and the least restrictive applies, as the specification says.
 */
export function fontFileInfo(input: Uint8Array | ArrayBuffer): FontFileInfo {
  const bytes = asBytes(input)
  const signature = bytes.length >= 4 ? tag(bytes, 0) : ''
  let format: FontFormat = 'unknown'
  if (signature === 'ttcf') format = 'collection'
  else if (signature === 'OTTO') {
    const cff = table(bytes, 'CFF ')
    format = cff && cidKeyed(bytes, cff) ? 'cff-cid' : 'cff'
  } else if (signature === 'true' || (bytes[0] === 0 && bytes[1] === 1 && bytes[2] === 0 && bytes[3] === 0)) {
    format = table(bytes, 'glyf') ? 'truetype' : 'unknown'
  }
  const os2 = format === 'collection' ? undefined : table(bytes, 'OS/2')
  if (!os2 || os2.length < 10) return { format, level: 'unknown', bitmapOnly: false }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const version = view.getUint16(os2.offset)
  const fsType = view.getUint16(os2.offset + 8)
  const bits = fsType & 0x000f
  let level: EmbeddingLevel
  if (bits === 0) level = 'installable'
  else if (version <= 2 ? (bits & 0x8) !== 0 : bits === 0x8) level = 'editable'
  else if (version <= 2 ? (bits & 0x4) !== 0 : bits === 0x4) level = 'previewPrint'
  else if (bits & 0x2) level = 'restricted'
  else level = 'unknown'
  return { format, fsType, level, bitmapOnly: (fsType & 0x0200) !== 0 }
}

/** The slot a face fills in its family: 600 and heavier is bold. */
export function fontSlot(font: Pick<NamedFont, 'weight' | 'style'>): FontSlot {
  const bold = (font.weight ?? 400) >= 600
  const italic = font.style === 'italic'
  return bold && italic ? 'boldItalic' : bold ? 'bold' : italic ? 'italic' : 'regular'
}

/** How far a face's weight is from the weight its slot stands for (400 or 700). */
const slotDistance = (font: NamedFont): number => Math.abs((font.weight ?? 400) - ((font.weight ?? 400) >= 600 ? 700 : 400))

/** The family name as the writers and the layout key it: trimmed, lower case. */
export const familyKey = (family: string): string => family.trim().toLowerCase()

/**
 * The faces of each family, one a slot: where two faces want the same slot, the one whose weight is nearer 400
 * (regular) or 700 (bold) has it, and the other is returned in `extra`.
 */
export function groupFonts(fonts: readonly NamedFont[]): {
  families: Map<string, { family: string; faces: Partial<Record<FontSlot, NamedFont>> }>
  extra: NamedFont[]
} {
  const families = new Map<string, { family: string; faces: Partial<Record<FontSlot, NamedFont>> }>()
  const extra: NamedFont[] = []
  for (const font of fonts) {
    if (typeof font.family !== 'string' || !font.family.trim()) throw new Error('document-kit: a font needs a family name')
    const key = familyKey(font.family)
    const entry = families.get(key) ?? { family: font.family.trim(), faces: {} }
    families.set(key, entry)
    const slot = fontSlot(font)
    const held = entry.faces[slot]
    if (!held) entry.faces[slot] = font
    else if (slotDistance(font) < slotDistance(held)) {
      extra.push(held)
      entry.faces[slot] = font
    } else extra.push(font)
  }
  return { families, extra }
}

const LEVEL_WORDS: Readonly<Record<EmbeddingLevel, string>> = {
  installable: 'installable embedding',
  editable: 'editable embedding',
  previewPrint: 'preview & print embedding only',
  restricted: 'restricted licence embedding (no embedding)',
  unknown: 'no embedding permission it states',
}

/**
 * Which faces a writer embeds, and why each is or is not, by the font's own licence flags (OS/2 `fsType`) and what the
 * target can carry:
 *
 * - `docx`: TrueType outlines only (Word embeds no CFF), and only installable or editable embedding, since a Word file
 *   stays editable. Embedded whole, never subset.
 * - `html`: TrueType or OpenType; installable, editable or preview & print embedding (a print view is viewed and
 *   printed). Bitmap-only and restricted faces are never embedded.
 *
 * `writeDocx` (option `embedFonts`) and `writeHtml` (option `fonts`) embed exactly the faces this marks `embedded`. A
 * family is embedded only when one of its faces can fill the regular slot (Word needs a regular face; the bold one
 * stands in when there is no regular).
 */
export function fontEmbeddingReport(fonts: readonly NamedFont[], target: EmbeddingTarget): FontEmbedding[] {
  const { families, extra } = groupFonts(fonts)
  const out: FontEmbedding[] = []
  const describe = (font: NamedFont, embedded: boolean, reason: string, info: FontFileInfo): FontEmbedding => ({
    family: font.family.trim(),
    weight: font.weight ?? 400,
    style: font.style ?? 'normal',
    slot: fontSlot(font),
    format: info.format,
    ...(info.fsType !== undefined ? { fsType: info.fsType } : {}),
    embedded,
    reason,
  })
  for (const font of fonts) {
    const info = fontFileInfo(font.bytes)
    if (extra.includes(font)) {
      out.push(describe(font, false, `another face of "${font.family.trim()}" nearer the ${fontSlot(font)} weight fills its slot`, info))
      continue
    }
    const refusal = refuse(info, target)
    if (refusal) {
      out.push(describe(font, false, refusal, info))
      continue
    }
    out.push(describe(font, true, `${info.format === 'truetype' ? 'TrueType' : 'OpenType'} face with ${LEVEL_WORDS[info.level]}; embedded whole, not subset`, info))
  }
  if (target === 'docx') {
    // Word reads a family's faces only when it has a regular one (the bold face stands in for it when alone).
    for (const [, family] of families) {
      const ok = (slot: FontSlot) => out.some((e) => e.embedded && e.family === family.family && e.slot === slot)
      if (ok('regular') || ok('bold')) continue
      for (let i = 0; i < out.length; i += 1) {
        const entry = out[i]!
        if (entry.embedded && familyKey(entry.family) === familyKey(family.family)) {
          out[i] = { ...entry, embedded: false, reason: `"${family.family}" has no regular or bold face that can be embedded, and Word needs one to read the family` }
        }
      }
    }
  }
  return out
}

function refuse(info: FontFileInfo, target: EmbeddingTarget): string | undefined {
  if (info.format === 'collection') return 'a font collection (.ttc) cannot be embedded; pass its faces one by one'
  if (info.format === 'unknown') return 'the file is not a TrueType or OpenType font'
  if (target === 'docx' && info.format !== 'truetype') {
    return 'OpenType (CFF) outlines: Word embeds TrueType outlines only; install the font where the document is opened, or supply a TrueType build'
  }
  if (info.fsType === undefined) return 'the font has no OS/2 table stating its embedding permission; install it where the document is opened'
  if (info.bitmapOnly) return 'the font permits bitmap embedding only'
  if (info.level === 'restricted') return 'the font forbids embedding (restricted licence)'
  if (info.level === 'unknown') return `the font's embedding flags (fsType ${info.fsType}) grant no embedding`
  if (target === 'docx' && info.level === 'previewPrint') {
    return 'the font permits preview & print embedding only, which an editable Word file cannot use; install it where the document is opened'
  }
  return undefined
}
