/**
 * Pins what the docx library leaves to chance, so the same document gives the same file:
 *
 * - hyperlink relationship ids (random nanoid strings) become `rIdL1`, `rIdL2`… in the order the part lists them;
 * - drawing ids (`wp:docPr id`), bookmark ids and the contents' `_Toc` bookmark names, which come from counters that
 *   run on across every document made in the same page, are numbered from 1 in each file;
 * - the created and modified times in `docProps/core.xml` are the document's own when it gives them;
 * - the zip entries carry one fixed time instead of the moment of writing.
 *
 * Runs on the package's bytes with fflate and string replacement only, so it works in a browser.
 */
import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate'

export interface NormalizeOptions {
  /** ISO 8601 creation time for `docProps/core.xml`; left as the library wrote it when absent. */
  created?: string
  /** ISO 8601 modified time; `created` when absent. */
  modified?: string
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
    const swap = (source: string) => source.replace(/"([^"]+)"/g, (whole, id: string) => {
      const next = renamed.get(id)
      return next ? `"${next}"` : whole
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

  // [Content_Types].xml first, as Word writes it; the rest in the library's order.
  const order = ['[Content_Types].xml', ...names.filter((name) => name !== '[Content_Types].xml')]
  const zippable: Zippable = {}
  const mtime = ZIP_TIME()
  for (const name of order) {
    if (!(name in files)) continue
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
