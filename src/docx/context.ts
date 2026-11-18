/** What every part of the writer needs to know while it writes one section. */
import type { CaptionedBlock, PageBox } from '../model/index'
import type { TextStyle, Theme } from '../model/types'

export interface WriterContext {
  theme: Theme
  page: PageBox
  /** Width the current container gives its content, points: the text width, or a table cell's inside. */
  width: number
  part: 'body' | 'header' | 'footer'
  captions: ReadonlyMap<CaptionedBlock, number>
  sourceStyles: Readonly<Record<string, TextStyle>>
  nullCells: 'empty' | 'rule'
  lists: ListRegistry
}

/** The numbering definitions the lists of a document need, gathered while the body is written. */
export interface ListRegistry {
  /** A new instance of an ordered or bulleted list, so each ordered list counts from its own start. */
  instance(): number
  /** The numbering reference for a kind and first number. */
  reference(ordered: boolean, start: number): string
}
