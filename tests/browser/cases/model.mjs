/**
 * Browser case: build a document with the block model from the built package, write it as Word with
 * `document-kit/docx`, and check the file in the page itself. Made-up content only. Throws on failure; returns a
 * JSON-serialisable summary.
 */
import {
  caption,
  custom,
  defineCustomBlocks,
  document,
  field,
  fields,
  figure,
  groupRow,
  headerRow,
  heading,
  image,
  imageInfo,
  noteRow,
  numberCaptions,
  pageCount,
  pageNumber,
  paragraph,
  row,
  section,
  sup,
  table,
  toComplete,
  toc,
  UnknownCustomBlockError,
} from 'document-kit'
import { writeDocx } from 'document-kit/docx'

/** A 2 × 1 PNG (two grey pixels), so no canvas or decoder is needed. */
const PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAC0lEQVR4nGNoAAMACocDAcTGH2YAAAAASUVORK5CYII='

function bytesOf(base64) {
  const binary = atob(base64)
  return Uint8Array.from(binary, (c) => c.charCodeAt(0))
}

/** The names in a zip file's central directory, read without unzipping. */
function zipNames(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  // End of central directory: the last record with signature 0x06054b50.
  let end = -1
  for (let at = bytes.length - 22; at >= Math.max(0, bytes.length - 22 - 65535); at -= 1) {
    if (view.getUint32(at, true) === 0x06054b50) {
      end = at
      break
    }
  }
  if (end < 0) throw new Error('model case: the file has no zip end record')
  const count = view.getUint16(end + 10, true)
  let at = view.getUint32(end + 16, true)
  const names = []
  const decoder = new TextDecoder()
  for (let i = 0; i < count; i += 1) {
    if (view.getUint32(at, true) !== 0x02014b50) throw new Error('model case: a broken central directory entry')
    const nameLength = view.getUint16(at + 28, true)
    const extra = view.getUint16(at + 30, true)
    const comment = view.getUint16(at + 32, true)
    names.push(decoder.decode(bytes.subarray(at + 46, at + 46 + nameLength)))
    at += 46 + nameLength + extra + comment
  }
  return names
}

export default async function modelCase(ctx) {
  void ctx
  const picture = image(bytesOf(PNG_BASE64))
  const info = imageInfo(picture.bytes)
  if (!info || info.width !== 2 || info.height !== 1) throw new Error(`model case: read the PNG as ${JSON.stringify(info)}`)

  const registry = defineCustomBlocks({
    marked: (data) => [
      table([headerRow(['', 'A', 'B']), row(['Row', data.at === 0 ? { content: 'x', shading: 'f4c7c3' } : '', data.at === 1 ? { content: 'x', shading: 'f4c7c3' } : ''])], {
        caption: caption('Made-up matrix'),
      }),
    ],
  })
  const results = table(
    [
      headerRow([{ content: 'Item', rowSpan: 2 }, { content: 'Values', colSpan: 2 }]),
      headerRow(['V1', 'V2']),
      groupRow([{ content: 'Group', colSpan: 3 }]),
      ...Array.from({ length: 20 }, (_, i) => row([[`Item ${i + 1}`, ...(i === 0 ? [sup('a')] : [])], String(i), i === 5 ? null : String(i * 2)])),
      noteRow('a Made-up note.'),
    ],
    { caption: caption('Made-up results'), columnWidths: [2, 1, 1] },
  )
  const doc = document(
    [
      section([toc({ title: 'Contents' }), heading(1, 'Made-up report'), fields([field('Date', '1 May'), field('Checked by', null)]), results, custom('marked', { at: 1 })], {
        footer: [paragraph(['Page ', pageNumber(), ' of ', pageCount()], { align: 'center' })],
      }),
      section([figure(picture, { fit: 'page', caption: caption('Made-up sheet') })], { page: { paper: 'tabloid', orientation: 'landscape' } }),
    ],
    { metadata: { title: 'Browser case', created: '2026-01-01T00:00:00Z' } },
  )

  const bytes = await writeDocx(doc, { registry })
  if (!(bytes instanceof Uint8Array) || bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new Error('model case: not a zip file')
  const names = zipNames(bytes)
  for (const part of ['[Content_Types].xml', 'word/document.xml', 'word/styles.xml', 'word/numbering.xml']) {
    if (!names.includes(part)) throw new Error(`model case: the file has no ${part}`)
  }
  if (!names.some((name) => /^word\/footer\d+\.xml$/.test(name))) throw new Error('model case: no footer part')
  if (!names.some((name) => /^word\/media\/.+\.png$/.test(name))) throw new Error('model case: no picture in the file')

  const again = await writeDocx(doc, { registry })
  const same = again.length === bytes.length && again.every((byte, i) => byte === bytes[i])
  if (!same) throw new Error('model case: the same document wrote different bytes')

  let unknownKind = null
  try {
    await writeDocx(document([section([custom('notRegistered', {})])]))
  } catch (error) {
    if (!(error instanceof UnknownCustomBlockError)) throw error
    unknownKind = error.kind
  }
  if (unknownKind !== 'notRegistered') throw new Error('model case: an unregistered custom block was not refused')

  const toDo = toComplete(doc, { registry })
  return {
    byteLength: bytes.length,
    parts: names.length,
    deterministic: same,
    captions: numberCaptions(doc).size,
    toComplete: toDo.map((item) => item.what),
    unknownKind,
  }
}
