/** What every part of the writer needs to know while it writes one section. */
import type { IFrameOptions } from 'docx'
import type { CaptionedBlock, PageBox } from '../model/index'
import type { TextStyle, Theme } from '../model/types'

import type { Footnotes } from './footnotes'

export interface WriterContext {
  bookmarks: Set<string>
  footnotes: Footnotes
  theme: Theme
  page: PageBox
  /** Width the current container gives its content, points: the text width, or a table cell's inside. */
  width: number
  part: 'body' | 'header' | 'footer'
  captions: ReadonlyMap<CaptionedBlock, number>
  sourceStyles: Readonly<Record<string, TextStyle>>
  nullCells: 'empty' | 'rule'
  lists: ListRegistry
  /** Inside a pinned block: the text frame every paragraph written there carries (one frame together). */
  frame?: IFrameOptions
  /** Writing the blocks of a table cell. */
  inCell?: boolean
}

/** The numbering definitions the lists of a document need, gathered while the body is written. */
export interface ListRegistry {
  /** A new instance of an ordered or bulleted list, so each ordered list counts from its own start. */
  instance(): number
  /** The numbering reference for a kind and first number. */
  reference(ordered: boolean, start: number): string
}
