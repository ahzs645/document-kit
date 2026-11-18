/**
 * Small functions that build the model, so application code reads as the document it makes:
 * `section([heading(1, 'Results'), table([headerRow(['Item', 'Value']), row(['A', '1'])])])`.
 * Each returns plain data; nothing here is needed to use the model.
 */
import type {
  Block,
  BlankRun,
  Caption,
  CellInput,
  CustomBlock,
  Document,
  FieldRow,
  FieldRun,
  FieldsBlock,
  FigureBlock,
  FigureGridBlock,
  FigureGridItem,
  HeadingBlock,
  ImageData,
  InlineImage,
  LineBreak,
  ListBlock,
  ListItem,
  PageBreakBlock,
  ParagraphBlock,
  RichText,
  RowRole,
  Section,
  SentenceRun,
  SignatureBlock,
  SlotValues,
  SpacerBlock,
  TabRun,
  TableBlock,
  TableCell,
  TableRow,
  TextRun,
  TextStyle,
  TocBlock,
} from './types'

type Options<T, Omitted extends keyof T> = Omit<T, Omitted | 'type'>

/* ------------------------------------------------------------------ document */

export const document = (sections: readonly Section[], options: Omit<Document, 'sections'> = {}): Document => ({
  ...options,
  sections,
})

export const section = (body: readonly Block[], options: Omit<Section, 'body'> = {}): Section => ({ ...options, body })

/* ------------------------------------------------------------------ inline */

export const text = (value: string, style: TextStyle & { link?: string; source?: string } = {}): TextRun => ({
  ...style,
  type: 'text',
  text: value,
})
export const bold = (value: string, style: TextStyle = {}): TextRun => text(value, { ...style, bold: true })
export const italic = (value: string, style: TextStyle = {}): TextRun => text(value, { ...style, italic: true })
export const sup = (value: string, style: TextStyle = {}): TextRun => text(value, { ...style, superscript: true })
export const sub = (value: string, style: TextStyle = {}): TextRun => text(value, { ...style, subscript: true })
export const link = (value: string, href: string, style: TextStyle = {}): TextRun =>
  text(value, { ...style, link: href })
export const lineBreak = (): LineBreak => ({ type: 'break' })
export const tab = (): TabRun => ({ type: 'tab' })
export const pageNumber = (style: TextStyle = {}): FieldRun => ({ ...style, type: 'field', field: 'page' })
export const pageCount = (style: TextStyle = {}): FieldRun => ({ ...style, type: 'field', field: 'pages' })
export const sectionPageCount = (style: TextStyle = {}): FieldRun => ({ ...style, type: 'field', field: 'sectionPages' })
export const blank = (what?: string, options: Options<BlankRun, 'what'> = {}): BlankRun => ({
  ...options,
  type: 'blank',
  ...(what === undefined ? {} : { what }),
})
export const inlineImage = (image: ImageData, width: number, height: number, alt?: string): InlineImage => ({
  type: 'image',
  image,
  width,
  height,
  ...(alt === undefined ? {} : { alt }),
})
export const sentence = (
  template: string,
  values: SlotValues,
  options: Options<SentenceRun, 'template' | 'values'> = {},
): SentenceRun => ({ ...options, type: 'sentence', template, values })

/**
 * Sentences joined by a space, as one paragraph's content. A sentence left out for a missing slot takes its space
 * with it ({@link resolveSentences}).
 */
export const sentences = (...items: ReadonlyArray<SentenceRun | TextRun | string>): RichText =>
  items.flatMap((item, index) => (index ? [' ', item] : [item]))

/** A picture's bytes, with its pixel size when the caller knows it (else read from the header). */
export const image = (bytes: Uint8Array, size: Omit<ImageData, 'bytes'> = {}): ImageData => ({ ...size, bytes })

/* ------------------------------------------------------------------ blocks */

export const heading = (
  level: HeadingBlock['level'],
  content: RichText,
  options: Options<HeadingBlock, 'level' | 'content'> = {},
): HeadingBlock => ({ ...options, type: 'heading', level, content })

export const paragraph = (content: RichText, options: Options<ParagraphBlock, 'content'> = {}): ParagraphBlock => ({
  ...options,
  type: 'paragraph',
  content,
})

export const title = (content: RichText, options: Options<ParagraphBlock, 'content' | 'style'> = {}) =>
  paragraph(content, { ...options, style: 'title' })
export const subtitle = (content: RichText, options: Options<ParagraphBlock, 'content' | 'style'> = {}) =>
  paragraph(content, { ...options, style: 'subtitle' })
export const note = (content: RichText, options: Options<ParagraphBlock, 'content' | 'style'> = {}) =>
  paragraph(content, { ...options, style: 'note' })

export const list = (
  items: readonly (ListItem | RichText)[],
  options: Options<ListBlock, 'items'> = {},
): ListBlock => ({ ...options, type: 'list', items })
export const orderedList = (items: readonly (ListItem | RichText)[], options: Options<ListBlock, 'items' | 'ordered'> = {}) =>
  list(items, { ...options, ordered: true })
export const item = (content: RichText, items?: readonly (ListItem | RichText)[], ordered?: boolean): ListItem => ({
  content,
  ...(items ? { items } : {}),
  ...(ordered === undefined ? {} : { ordered }),
})

export const field = (
  label: RichText,
  value: RichText | null | undefined,
  options: Omit<FieldRow, 'label' | 'value'> = {},
): FieldRow => ({ ...options, label, value: value === undefined || value === '' ? null : value })

export const fields = (rows: readonly FieldRow[], options: Options<FieldsBlock, 'rows'> = {}): FieldsBlock => ({
  ...options,
  type: 'fields',
  rows,
})

export const cell = (content: RichText | null, options: Omit<TableCell, 'content'> = {}): TableCell => ({
  ...options,
  content,
})
export const row = (cells: readonly CellInput[], options: Omit<TableRow, 'cells'> = {}): TableRow => ({
  ...options,
  cells,
})
const roleRow =
  (role: RowRole) =>
  (cells: readonly CellInput[], options: Omit<TableRow, 'cells' | 'role'> = {}): TableRow => ({
    ...options,
    role,
    cells,
  })
export const headerRow = roleRow('header')
export const groupRow = roleRow('group')
/** A note across the whole table. */
export const noteRow = (content: RichText, options: Omit<TableRow, 'cells' | 'role'> = {}): TableRow => ({
  ...options,
  role: 'note',
  cells: [content],
})
export const spacerRow = (options: Omit<TableRow, 'cells' | 'role'> = {}): TableRow => ({
  ...options,
  role: 'spacer',
  cells: [],
})

export const table = (rows: readonly TableRow[], options: Options<TableBlock, 'rows'> = {}): TableBlock => ({
  ...options,
  type: 'table',
  rows,
})

export const caption = (content: RichText, numbered?: Caption['numbered']): Caption => ({
  text: content,
  ...(numbered === undefined ? {} : { numbered }),
})

export const figure = (image: ImageData | null, options: Options<FigureBlock, 'image'> = {}): FigureBlock => ({
  ...options,
  type: 'figure',
  image,
})

export const figureGrid = (
  items: readonly FigureGridItem[],
  columns: number,
  options: Options<FigureGridBlock, 'items' | 'columns'> = {},
): FigureGridBlock => ({ ...options, type: 'figureGrid', items, columns })

export const signature = (options: Options<SignatureBlock, never> = {}): SignatureBlock => ({
  ...options,
  type: 'signature',
})

export const pageBreak = (): PageBreakBlock => ({ type: 'pageBreak' })

export const toc = (options: Options<TocBlock, never> = {}): TocBlock => ({ ...options, type: 'toc' })

export const spacer = (height: number): SpacerBlock => ({ type: 'spacer', height })

export const custom = <Data>(
  kind: string,
  data: Data,
  options: Options<CustomBlock<Data>, 'kind' | 'data'> = {},
): CustomBlock<Data> => ({ ...options, type: 'custom', kind, data })
