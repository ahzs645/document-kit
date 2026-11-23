import { Bookmark, FootnoteReferenceRun, Paragraph, SimpleField, TextRun } from 'docx'
import type { Document } from '../model/types'

/** Keep the cached field result readable even in viewers that do not evaluate Word fields. */
class NoteReferenceField extends SimpleField {
  constructor(bookmark: string, number: number) {
    super(`NOTEREF ${bookmark} \\h \\f`)
    this.root.push(new TextRun({ style: 'FootnoteReference', text: String(number) }))
  }
}

/** One registry per exported document, shared by sections and nested table cells in reading order. */
export class Footnotes {
  readonly definitions: Record<number, { children: Paragraph[] }> = {}
  private readonly numbers = new Map<string, number>()

  constructor(private readonly catalog: Document['footnotes']) {}

  get size() { return this.numbers.size }

  reference(id: string): Bookmark | SimpleField {
    // This alphabet maps one-to-one onto valid Word bookmark names up to 40 characters.
    if (typeof id !== 'string' || !/^[a-z][a-z0-9-]{0,32}$/.test(id))
      throw new Error(`document-kit: invalid footnote ID "${id}" (use 1–33 lowercase letters, digits or hyphens, starting with a letter)`)
    const bookmark = `dkNote_${id.replace(/-/g, '_')}`
    const existing = this.numbers.get(id)
    if (existing !== undefined) return new NoteReferenceField(bookmark, existing)
    const paragraphs = this.catalog && Object.hasOwn(this.catalog, id) ? this.catalog[id] : undefined
    if (!Array.isArray(paragraphs) || !paragraphs.length || !paragraphs.every(p => typeof p === 'string' && p.trim()))
      throw new Error(`document-kit: footnote "${id}" needs a definition with one or more nonempty paragraphs`)
    const number = this.size + 1
    this.numbers.set(id, number)
    this.definitions[number] = {
      children: paragraphs.map((p, index) => new Paragraph({
        style: 'FootnoteText',
        text: index === 0 ? ' ' + p : p,
        spacing: { after: 45 },
      })),
    }
    return new Bookmark({ id: bookmark, children: [new FootnoteReferenceRun(number)] })
  }
}
