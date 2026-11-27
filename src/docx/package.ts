/**
 * Pins what the docx library leaves to chance, so the same document gives the same file:
 *
 * - hyperlink relationship ids (random nanoid strings) become `rIdL1`, `rIdL2`… in the order the part lists them;
 * - drawing ids (`wp:docPr id`), bookmark ids and the contents' `_Toc` bookmark names, which come from counters that
 *   run on across every document made in the same page, are numbered from 1 in each file;
 * - the created and modified times in `docProps/core.xml` are the document's own when it gives them;
 * - the zip entries carry one fixed time instead of the moment of writing;
 * - an embedded font's obfuscation key (`w:fontKey`, a random GUID) is made from the font's own bytes, and the font file
 *   re-obfuscated with it;
 * - the custom properties (`docProps/custom.xml`), when given, are written with their types.
 *
 * Runs on the package's bytes with fflate and string replacement only, so it works in a browser.
 */
import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate'
import type { CustomPropertyValue } from '../model/types'

export interface NormalizeOptions {
  /** ISO 8601 creation time for `docProps/core.xml`; left as the library wrote it when absent. */
  created?: string
  /** ISO 8601 modified time; `created` when absent. */
  modified?: string
  /** Custom document properties, written as `docProps/custom.xml` in place of what the package has. */
  customProperties?: Readonly<Record<string, CustomPropertyValue>>
}

const XML = /\.(xml|rels)$/

/** A fixed zip entry time, built from local fields so every time zone writes the same DOS date: 1 January 1980. */
const ZIP_TIME = (): Date => new Date(1980, 0, 1, 0, 0, 0)

export function normalizePackage(bytes: Uint8Array, options: NormalizeOptions = {}): Uint8Array {
  const files = unzipSync(bytes)
  const names = Object.keys(files)
  const text = new Map<string, string>()
  for (const name of names) if (XML.test(name)) text.set(name, strFromU8(files[name]!))

  // Relationship ids: per part, the ids that are not rId<number>.
  for (const relsName of names.filter((name) => name.endsWith('.rels'))) {
    const rels = text.get(relsName)
    if (rels === undefined) continue
    const owner = relsName.replace(/_rels\/([^/]+)\.rels$/, '$1')
    const random = [...rels.matchAll(/\bId="([^"]+)"/g)].map((m) => m[1]!).filter((id) => !/^rId\d+$/.test(id))
    if (!random.length) continue
    const renamed = new Map(random.map((id, index) => [id, `rIdL${index + 1}`]))
    const swap = (source: string) => source.replace(/="([^"]*)"/g, (whole, id: string) => {
      const next = renamed.get(id)
      return next ? `="${next}"` : whole
    })
    text.set(relsName, swap(rels))
    const part = text.get(owner)
    if (part !== undefined) text.set(owner, swap(part))
  }

  // Drawing ids, numbered through the parts in a fixed order: the body first, then the rest by name.
  const parts = [...text.keys()].filter((name) => name.startsWith('word/') && name.endsWith('.xml'))
  parts.sort((a, b) => (a === 'word/document.xml' ? -1 : b === 'word/document.xml' ? 1 : a < b ? -1 : a > b ? 1 : 0))
  let drawing = 0
  for (const name of parts) {
    const source = text.get(name)!
    if (!source.includes('docPr')) continue
    text.set(
      name,
      source.replace(/(<wp:docPr\b[^>]*?\bid=")(\d+)(")/g, (_, open: string, __: string, close: string) => {
        drawing += 1
        return `${open}${drawing}${close}`
      }),
    )
  }

  // Bookmarks and the contents' `_Toc` names, per part, numbered in the order they appear.
  for (const name of parts) {
    let source = text.get(name)!
    if (source.includes('_Toc')) {
      const tocNames = new Map<string, string>()
      source = source.replace(/_Toc(\d+)/g, (whole) => {
        if (!tocNames.has(whole)) tocNames.set(whole, `_Toc${tocNames.size + 1}`)
        return tocNames.get(whole)!
      })
    }
    if (source.includes('w:bookmark')) {
      const ids = new Map<string, number>()
      source = source.replace(/(<w:bookmark(?:Start|End)\b[^>]*?\bw:id=")(\d+)(")/g, (_, open: string, id: string, close: string) => {
        if (!ids.has(id)) ids.set(id, ids.size + 1)
        return `${open}${ids.get(id)}${close}`
      })
    }
    text.set(name, source)
  }

  // Creation and modified times.
  const core = text.get('docProps/core.xml')
  if (core !== undefined && options.created) {
    const created = isoTime(options.created)
    const modified = isoTime(options.modified ?? options.created)
    text.set(
      'docProps/core.xml',
      core
        .replace(/(<dcterms:created\b[^>]*>)[^<]*(<\/dcterms:created>)/, `$1${created}$2`)
        .replace(/(<dcterms:modified\b[^>]*>)[^<]*(<\/dcterms:modified>)/, `$1${modified}$2`),
    )
  }

  // Embedded fonts: each face's key made from its own bytes, so the same fonts give the same file.
  const fontTable = text.get('word/fontTable.xml')
  const fontRels = text.get('word/_rels/fontTable.xml.rels')
  if (fontTable !== undefined && fontRels !== undefined && fontTable.includes('w:fontKey')) {
    const targets = new Map<string, string>()
    for (const m of fontRels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = /\bId="([^"]+)"/.exec(m[0])?.[1]
      const target = /\bTarget="([^"]+)"/.exec(m[0])?.[1]
      if (id && target) targets.set(id, target.startsWith('/') ? target.slice(1) : `word/${target}`)
    }
    let table = fontTable
    for (const m of fontTable.matchAll(/<w:embed(?:Regular|Bold|Italic|BoldItalic)\b[^>]*\/>/g)) {
      const element = m[0]
      const id = /\br:id="([^"]+)"/.exec(element)?.[1]
      const key = /\bw:fontKey="\{?([0-9A-Fa-f-]{36})\}?"/.exec(element)?.[1]
      const file = id ? targets.get(id) : undefined
      if (!key || !file || !files[file]) continue
      const plain = obfuscateFont(files[file]!, key)
      const next = fontKeyOf(plain, file)
      files[file] = obfuscateFont(plain, next)
      table = table.replace(element, element.replace(/(\bw:fontKey=")[^"]*(")/, `$1{${next}}$2`))
    }
    text.set('word/fontTable.xml', table)
  }

  // Custom properties, with their types.
  if (options.customProperties) {
    text.set('docProps/custom.xml', customXml(options.customProperties))
    if (!names.includes('docProps/custom.xml')) names.push('docProps/custom.xml')
    const rels = text.get('_rels/.rels')
    if (rels !== undefined && !rels.includes('docProps/custom.xml')) {
      text.set(
        '_rels/.rels',
        rels.replace(
          '</Relationships>',
          '<Relationship Id="rIdCustom" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>',
        ),
      )
    }
    const types = text.get('[Content_Types].xml')
    if (types !== undefined && !types.includes('/docProps/custom.xml')) {
      text.set(
        '[Content_Types].xml',
        types.replace(
          '</Types>',
          '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>',
        ),
      )
    }
  }

  // [Content_Types].xml first, as Word writes it; the rest in the library's order.
  const order = ['[Content_Types].xml', ...names.filter((name) => name !== '[Content_Types].xml')]
  const zippable: Zippable = {}
  const mtime = ZIP_TIME()
  for (const name of order) {
    if (!(name in files) && !text.has(name)) continue
    const data = text.has(name) ? strToU8(text.get(name)!) : files[name]!
    zippable[name] = [data, { mtime }]
  }
  return zipSync(zippable, { level: 6 })
}

/** An ISO time as W3CDTF in UTC, seconds precision; throws on a time that does not parse. */
function isoTime(value: string): string {
  const time = Date.parse(value)
  if (Number.isNaN(time)) throw new Error(`document-kit: "${value}" is not an ISO 8601 time`)
  return new Date(time).toISOString().replace(/\.\d{3}Z$/, 'Z')
}

/**
 * ECMA-376 Part 1 §17.8.1 font obfuscation: the first 32 bytes of the file XORed with the key's 16 bytes, the key read
 * from its GUID's hex digits in reverse. The same operation hides and reveals.
 */
export function obfuscateFont(bytes: Uint8Array, key: string): Uint8Array<ArrayBuffer> {
  const hex = key.replace(/[{}-]/g, '')
  if (!/^[0-9A-Fa-f]{32}$/.test(hex)) throw new Error(`document-kit: "${key}" is not a font key GUID`)
  const digits = Array.from({ length: 16 }, (_, i) => parseInt(hex.slice(i * 2, i * 2 + 2), 16)).reverse()
  const out = new Uint8Array(bytes)
  for (let i = 0; i < 32 && i < out.length; i += 1) out[i] = out[i]! ^ digits[i % 16]!
  return out
}

/** A GUID made from a font's bytes and its part name: four FNV-1a hashes with different seeds. */
function fontKeyOf(bytes: Uint8Array, part: string): string {
  const words: string[] = []
  for (const seed of [0x811c9dc5, 0x01000193, 0x9e3779b9, 0x85ebca6b]) {
    let h = seed >>> 0
    for (let i = 0; i < part.length; i += 1) h = Math.imul(h ^ part.charCodeAt(i), 0x01000193) >>> 0
    for (let i = 0; i < bytes.length; i += 1) h = Math.imul(h ^ bytes[i]!, 0x01000193) >>> 0
    words.push(h.toString(16).padStart(8, '0').toUpperCase())
  }
  const all = words.join('')
  return `${all.slice(0, 8)}-${all.slice(8, 12)}-${all.slice(12, 16)}-${all.slice(16, 20)}-${all.slice(20, 32)}`
}

const escapeXml = (value: string): string =>
  value.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!)

/** `docProps/custom.xml` for the properties, each value in its own type. Throws on a name or value it cannot write. */
export function customXml(properties: Readonly<Record<string, CustomPropertyValue>>): string {
  const items = Object.entries(properties).map(([name, value], index) => {
    if (!name.trim() || name.length > 255) throw new Error(`document-kit: a custom property's name must be 1 to 255 characters (${JSON.stringify(name)})`)
    let typed: string
    if (typeof value === 'string') typed = `<vt:lpwstr>${escapeXml(value)}</vt:lpwstr>`
    else if (typeof value === 'boolean') typed = `<vt:bool>${value ? 'true' : 'false'}</vt:bool>`
    else if (typeof value === 'number') {
      if (!Number.isFinite(value)) throw new Error(`document-kit: the custom property "${name}" is not a finite number`)
      typed = Number.isInteger(value) && Math.abs(value) <= 0x7fffffff ? `<vt:i4>${value}</vt:i4>` : `<vt:r8>${value}</vt:r8>`
    } else if (value instanceof Date) {
      if (Number.isNaN(value.getTime())) throw new Error(`document-kit: the custom property "${name}" is not a valid time`)
      typed = `<vt:filetime>${value.toISOString().replace(/\.\d{3}Z$/, 'Z')}</vt:filetime>`
    } else {
      throw new Error(`document-kit: the custom property "${name}" must be text, a number, a yes/no or a Date`)
    }
    return `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${index + 2}" name="${escapeXml(name)}">${typed}</property>`
  })
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ' +
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    items.join('') +
    '</Properties>'
  )
}
