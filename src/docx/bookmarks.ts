import { Bookmark, type ParagraphChild } from 'docx'
import type { WriterContext } from './context'

export function bookmarked(name: string | undefined, children: ParagraphChild[], context: WriterContext): ParagraphChild[] {
  if (name === undefined) return children
  if (!/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(name) || /^dkNote_/i.test(name))
    throw new Error(`Invalid Word bookmark name: ${name}`)
  if (context.bookmarks.has(name.toLowerCase())) throw new Error(`Duplicate Word bookmark: ${name}`)
  context.bookmarks.add(name.toLowerCase())
  return [new Bookmark({ id: name, children })]
}
