/**
 * The block model: a document as plain data before it is laid out. Both applications build this, and every writer
 * (Word here, others later) reads it. Nothing in this file renders, measures or touches the DOM.
 *
 * Units: every length is in points (1/72 inch). Font sizes are points. Picture sizes given in pixels are pixels of
 * the picture itself, read at 96 to the inch. Colours are hex, `#rrggbb` or `#rgb` in any case, with or without the `#`.
 */

/* ------------------------------------------------------------------ inline text */

/** Word's highlight colours: the only colours a highlight can take in a Word file. */
export type Highlight =
  | 'yellow'
  | 'green'
  | 'cyan'
  | 'magenta'
  | 'blue'
  | 'red'
  | 'darkBlue'
  | 'darkCyan'
  | 'darkGreen'
  | 'darkMagenta'
  | 'darkRed'
  | 'darkYellow'
  | 'darkGray'
  | 'lightGray'
  | 'black'
  | 'white'

/** How a stretch of text looks. Anything left out comes from the paragraph's style and the theme. */
export interface TextStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  superscript?: boolean
  subscript?: boolean
  smallCaps?: boolean
  allCaps?: boolean
  /** Six hex digits. */
  color?: string
  /** Points. */
  size?: number
  font?: string
  highlight?: Highlight
}

/**
 * A stretch of text in one style. In `text`, a newline is a line break and a tab character is a tab. `source` is the
 * application's own tag for where the value came from ("run", "reviewer", "lab"); the kit never interprets it, but a
 * writer can be told how to mark text by its source.
 */
export interface TextRun extends TextStyle {
  type: 'text'
  text: string
  /** A web address, `mailto:` address or other URI: the run is written as a real hyperlink. */
  link?: string
  source?: string
}

/** A field Word fills in as it lays the page out: this page's number, the document's page count, the section's. */
export interface FieldRun extends TextStyle {
  type: 'field'
  field: 'page' | 'pages' | 'sectionPages'
}

export interface LineBreak {
  type: 'break'
}

export interface TabRun {
  type: 'tab'
}

/**
 * A space to be filled by hand: printed as a rule. `what` says what goes there, for the list of things still to
 * complete ({@link toComplete}).
 */
export interface BlankRun extends TextStyle {
  type: 'blank'
  what?: string
  /** Points; the theme's blank width when absent. */
  width?: number
  source?: string
}

/** A picture in a line of text, such as a logo in a header, at a stated size. */
export interface InlineImage {
  type: 'image'
  image: ImageData
  /** Points. */
  width: number
  /** Points. */
  height: number
  alt?: string
}

/**
 * A sentence written over named slots: `"The block covers {area} ha."`. A slot with no value (null, undefined or an
 * empty string) means the sentence is not printed at all, and is listed as still to complete with the slots it is
 * missing. Never printed with braces.
 */
export interface SentenceRun extends TextStyle {
  type: 'sentence'
  template: string
  values: SlotValues
  /** What the sentence is, for the to-complete list ("Visibility statement"); the template when absent. */
  label?: string
  source?: string
}

export type SlotValues = Readonly<Record<string, string | number | null | undefined>>

export type Inline = string | TextRun | FieldRun | LineBreak | TabRun | BlankRun | InlineImage | SentenceRun

/** Text with its formatting: a plain string, one run, or a list of runs. */
export type RichText = string | Exclude<Inline, string> | readonly Inline[]

/* ------------------------------------------------------------------ pictures */

export type ImageFormat = 'png' | 'jpeg'

/**
 * A picture as bytes. `width` and `height` are its size in pixels; when absent they are read from the PNG or JPEG
 * header ({@link imageInfo}). The format is read from the bytes when absent.
 */
export interface ImageData {
  bytes: Uint8Array
  format?: ImageFormat
  width?: number
  height?: number
}

/* ------------------------------------------------------------------ shared block parts */

export type Align = 'left' | 'center' | 'right' | 'justify'
export type VerticalAlign = 'top' | 'center' | 'bottom'

export interface Border {
  style?: 'single' | 'double' | 'dashed' | 'dotted' | 'none'
  /** Points. */
  width?: number
  color?: string
}

/** Something the application knows is still to complete but the kit cannot see (a sentence it left out). */
export interface Pending {
  what: string
  /** Slots or inputs it is waiting on, in plain words. */
  missing?: readonly string[]
  source?: string
}

/** Properties every block may carry. */
export interface BlockBase {
  /** The application's own id for the block; the kit never reads it. */
  id?: string
  /** Where the block's content came from (the application's tag). */
  source?: string
  /** Items still to complete that this block knows of and cannot show. */
  pending?: readonly Pending[]
  /** Start the block on a new page. */
  pageBreakBefore?: boolean
}

export interface Spacing {
  /** Points before the paragraph. */
  before?: number
  /** Points after the paragraph. */
  after?: number
  /** Line height as a multiple of single spacing (1.15). */
  line?: number
}

export interface Indent {
  left?: number
  right?: number
  firstLine?: number
  hanging?: number
}

/**
 * A tab stop: `at` in points from the left margin, or `'middle'` / `'end'` of the text width ("Page 2 of 9" at the
 * right edge of a footer).
 */
export interface TabStop {
  at: number | 'middle' | 'end'
  align?: 'left' | 'center' | 'right'
  leader?: 'none' | 'dot'
}

/**
 * A caption for a table or a figure. Numbered captions take the next number of their kind in reading order
 * ({@link numberCaptions}): `"Table 3: Results"`. `numbered: false` prints the text alone and takes no number;
 * `'silent'` prints the text alone but still takes a number, for a caption that states its own.
 */
export interface Caption {
  text: RichText
  numbered?: boolean | 'silent'
}

/* ------------------------------------------------------------------ blocks */

export interface HeadingBlock extends BlockBase {
  type: 'heading'
  /** 1 to 6. Written as Word's own Heading styles, so a table of contents and the navigation pane read them. */
  level: 1 | 2 | 3 | 4 | 5 | 6
  content: RichText
  align?: Align
}

/** Paragraph styles every theme defines. */
export type ParagraphStyleName = 'normal' | 'title' | 'subtitle' | 'caption' | 'note' | 'small'

export interface ParagraphBlock extends BlockBase {
  type: 'paragraph'
  content: RichText
  style?: ParagraphStyleName
  align?: Align
  spacing?: Spacing
  indent?: Indent
  keepWithNext?: boolean
  keepLinesTogether?: boolean
  /** A base style for every run in the paragraph. */
  text?: TextStyle
  tabs?: readonly TabStop[]
  borders?: { top?: Border; bottom?: Border; left?: Border; right?: Border }
  /** Fill behind the paragraph, a hex colour. */
  shading?: string
}

export interface ListItem {
  content: RichText
  /** Items under this one, one level deeper. */
  items?: readonly (ListItem | RichText)[]
  /** The nested items' kind, when it differs from this list's. */
  ordered?: boolean
}

export interface ListBlock extends BlockBase {
  type: 'list'
  ordered?: boolean
  items: readonly (ListItem | RichText)[]
  /** The first number of an ordered list. */
  start?: number
}

export interface FieldRow {
  label: RichText
  /** Null prints a blank rule, to be filled by hand, and is listed as still to complete. */
  value: RichText | null
  source?: string
  /** What the blank is for the to-complete list; the label when absent. */
  what?: string
}

/** Label and value rows: "Date: 12 March", either as lines of text or as a two-column table. */
export interface FieldsBlock extends BlockBase {
  type: 'fields'
  title?: RichText
  rows: readonly FieldRow[]
  /** `lines` (default): one paragraph a row; `table`: a ruled two-column table. */
  layout?: 'lines' | 'table'
  /** `lines`: a tab stop in points the values line up on; `table`: the label column's share of the width, 0–100. */
  labelWidth?: number
  /** Printed after the label in `lines` layout; `": "` by default. */
  separator?: string
  /** Text size for the rows (points). */
  size?: number
}

/**
 * A table cell. A cell covered by another cell's `colSpan` or `rowSpan` is not written in the rows at all, as in
 * HTML: a cell spanning two rows appears once, in the upper row.
 */
export interface TableCell {
  /** The cell's text. Null prints an empty cell (or the theme's blank rule) and is listed as still to complete. */
  content?: RichText | null
  /** Blocks in the cell instead of text: paragraphs, lists, figures, a nested table. */
  blocks?: readonly Block[]
  colSpan?: number
  rowSpan?: number
  align?: Align
  verticalAlign?: VerticalAlign
  /** Fill, a hex colour. */
  shading?: string
  text?: TextStyle
  borders?: { top?: Border; bottom?: Border; left?: Border; right?: Border }
  /** A rule corner to corner: `down` top-left to bottom-right, `up` bottom-left to top-right (a struck cell). */
  diagonal?: 'down' | 'up'
  /**
   * Turn the cell's text a quarter: `up` reads from the bottom of the cell to the top (Word's bottom-to-top, `btLr`),
   * `down` from the top to the bottom (`tbRl`). The row grows to the length of the longest line; give the row a
   * `height` to fix it.
   */
  textDirection?: 'up' | 'down'
  source?: string
  /** What the blank is for the to-complete list. */
  what?: string
}

/** A cell as text alone (`"12.5"`, rich runs, or null for a blank) or as a {@link TableCell}. */
export type CellInput = RichText | null | TableCell

/**
 * A row's role decides how it is drawn: `header` in the theme's header look (and repeated at the top of each page
 * while it is among the table's first rows), `group` as a shaded band heading the rows under it, `note` as one small
 * cell across the table, `spacer` as a short empty row, `body` as ordinary content.
 */
export type RowRole = 'header' | 'body' | 'group' | 'note' | 'spacer'

export interface TableRow {
  role?: RowRole
  cells: readonly CellInput[]
  /**
   * Repeat the row at the top of each page. Defaults to true for `header` rows. Word repeats only rows at the top of a
   * table, so this has effect only in the table's first run of rows.
   */
  repeat?: boolean
  /** Never break the row across a page. */
  cantSplit?: boolean
  /**
   * Keep the row on the same page as the row after it (Word's keep-with-next on the row's paragraphs). Defaults to true
   * for a `header` row that has a row after it, so a heading row is never left alone at the foot of a page with its
   * body on the next. False lets a header row part from the rows under it.
   */
  keepWithNext?: boolean
  /** Least height, points. */
  height?: number
}

export interface TableBorders {
  top?: Border
  bottom?: Border
  left?: Border
  right?: Border
  insideHorizontal?: Border
  insideVertical?: Border
}

export interface TableBlock extends BlockBase {
  type: 'table'
  rows: readonly TableRow[]
  /** Relative column widths (any unit: they are shares). Equal columns when absent. */
  columnWidths?: readonly number[]
  /** Width as a share of the text width, 0–100. 100 by default. */
  width?: number
  align?: 'left' | 'center' | 'right'
  caption?: Caption
  /** Above (default) or below the table. */
  captionPosition?: 'above' | 'below'
  /** A note printed under the table in the theme's note style. */
  note?: RichText
  /** Text size in the table, points; the theme's table size when absent. */
  size?: number
  /** Borders for this table, over the theme's. */
  borders?: TableBorders
  /** Inner padding of every cell, points. */
  cellPadding?: { top?: number; bottom?: number; left?: number; right?: number }
  /** Keep every row whole (no row breaks across a page). */
  cantSplit?: boolean
  /**
   * Keep the whole table on one page: every row whole and kept with the next. A table taller than a page still breaks
   * (Word's rule as well), between rows. A row's own `keepWithNext` or `cantSplit` overrides it for that row.
   */
  keepTogether?: boolean
  /** Turn off repeating header rows for this table. */
  repeatHeader?: boolean
  /** Alignment of body cells by column, when a cell does not say. */
  columnAlign?: readonly (Align | undefined)[]
}

/** The box a picture is fitted into, keeping its proportions. */
export interface FitBox {
  /** Points; the text width when absent. */
  width?: number
  /** Points. */
  height?: number
  /** Enlarge a picture smaller than the box. False by default: a small picture keeps its own size. */
  upscale?: boolean
}

export interface FigureBlock extends BlockBase {
  type: 'figure'
  /** Null when the picture is missing: a placeholder box is printed and the figure listed as still to complete. */
  image: ImageData | null
  /** `'page'` fills what the section's page has left after the caption. */
  fit?: FitBox | 'page'
  caption?: Caption
  /** Text printed in the placeholder box when the picture is missing. */
  placeholder?: RichText
  align?: 'left' | 'center' | 'right'
  alt?: string
  /** A note printed under the caption. */
  note?: RichText
  /** Caption above (default below) the picture. */
  captionPosition?: 'above' | 'below'
  /**
   * Show part of the picture (Word's crop, `a:srcRect`): the share, 0 to 1, of its width or height cut away at each
   * edge. `{ top: 0.25, bottom: 0.5 }` shows the band from a quarter to half way down. The part shown is what is fitted
   * into the box. What is cut away stays in the file, as Word keeps it.
   */
  crop?: PictureCrop
  /** A solid line round the picture, drawn on its edge (Word's picture border). */
  border?: PictureBorder
}

/** The share, 0 to 1, of a picture's width (left, right) or height (top, bottom) cut away at each edge. */
export interface PictureCrop {
  top?: number
  right?: number
  bottom?: number
  left?: number
}

/** A solid frame round a picture. Word draws it centred on the picture's edge. */
export interface PictureBorder {
  /** Points; half a point when absent. */
  width?: number
  /** Hex colour; the body text's colour when absent. */
  color?: string
}

export interface FigureGridItem {
  image: ImageData | null
  label?: RichText
  placeholder?: RichText
  alt?: string
}

/** Pictures side by side in `columns` columns, each with a label, under one caption. One column stacks them. */
export interface FigureGridBlock extends BlockBase {
  type: 'figureGrid'
  items: readonly FigureGridItem[]
  columns: number
  /** Tallest a picture may be, points. Three quarters of its column's width when absent. */
  itemHeight?: number
  labelPosition?: 'above' | 'below'
  caption?: Caption
  upscale?: boolean
}

/** A signature block: a label, the signature picture or room to sign, the rule, the name, and lines under it. */
export interface SignatureBlock extends BlockBase {
  type: 'signature'
  label?: RichText
  /** The signature as a picture. Absent or null: room is left to sign by hand. */
  image?: ImageData | null
  /** The box the signature picture is fitted into, points (190 × 74 by default). */
  imageBox?: { width: number; height: number }
  /** Printed under the rule. Null is a blank to complete. */
  name?: RichText | null
  /** Rows under the name: "Title", "Date". A null value is a blank to complete. */
  lines?: readonly FieldRow[]
  align?: 'left' | 'center'
  /** Length of the signing rule, points. */
  ruleWidth?: number
  /** Room to sign by hand when there is no picture, points. */
  signingSpace?: number
}

export interface PageBreakBlock extends BlockBase {
  type: 'pageBreak'
}

/** A table of contents, filled from the headings. */
export interface TocBlock extends BlockBase {
  type: 'toc'
  title?: RichText
  /** Deepest heading level listed; 3 by default. */
  levels?: number
}

export interface SpacerBlock extends BlockBase {
  type: 'spacer'
  /** Points. */
  height: number
}

/**
 * A block the kit does not know: a rating matrix, a scale bar, a results table with its own rules. A writer expands
 * it through the application's registry ({@link defineCustomBlocks}) into the kit's own blocks; a kind nobody
 * registered stops the writer with its name.
 */
export interface CustomBlock<Data = unknown> extends BlockBase {
  type: 'custom'
  kind: string
  data: Data
}

export type Block =
  | HeadingBlock
  | ParagraphBlock
  | ListBlock
  | FieldsBlock
  | TableBlock
  | FigureBlock
  | FigureGridBlock
  | SignatureBlock
  | PageBreakBlock
  | TocBlock
  | SpacerBlock
  | CustomBlock

export type BlockType = Block['type']

/* ------------------------------------------------------------------ pages and sections */

/** Letter 8.5 × 11 in, legal 8.5 × 14 in, A4, tabloid 11 × 17 in. */
export type PaperName = 'letter' | 'legal' | 'a4' | 'tabloid'

/** A paper size in points, given portrait (width the shorter side). */
export interface PaperSize {
  width: number
  height: number
}

export interface Margins {
  top: number
  right: number
  bottom: number
  left: number
  /** From the top edge to the header. */
  header: number
  /** From the bottom edge to the footer. */
  footer: number
}

export interface PageSetup {
  paper?: PaperName | PaperSize
  orientation?: 'portrait' | 'landscape'
  margins?: Partial<Margins>
  /** Where the body sits between the margins: `center` for a cover page. */
  verticalAlign?: 'top' | 'center' | 'bottom'
}

export interface Section {
  page?: PageSetup
  /**
   * The header's blocks. Absent: the same as the section before (Word's own rule). An empty list: no header.
   */
  header?: readonly Block[]
  footer?: readonly Block[]
  /**
   * A different header and footer on the section's first page (a cover). A part left out here is empty on that page;
   * it does not fall back to `header` or `footer`.
   */
  firstPage?: { header?: readonly Block[]; footer?: readonly Block[] }
  /** Page numbering in this section: restart at `start`, and the number style. */
  pageNumbers?: { start?: number; format?: 'decimal' | 'lowerRoman' | 'upperRoman' | 'lowerLetter' | 'upperLetter' }
  body: readonly Block[]
}

export interface Metadata {
  title?: string
  author?: string
  subject?: string
  description?: string
  keywords?: readonly string[]
  /** ISO 8601. Written as the file's creation time; pins the file's bytes when given. */
  created?: string
  /** ISO 8601. The modified time; `created` when absent. */
  modified?: string
}

/** Text in one role in the theme. */
export interface ThemeText {
  font?: string
  /** Points. */
  size?: number
  color?: string
  bold?: boolean
  italic?: boolean
  allCaps?: boolean
  smallCaps?: boolean
  align?: Align
  /** Points before. */
  before?: number
  /** Points after. */
  after?: number
  /** Line height multiple. */
  line?: number
}

/** How tables look unless a table says otherwise. */
export interface TableLook {
  /** Text size in tables, points. */
  size: number
  headerFill: string
  headerText: ThemeText
  /** Header cell alignment. */
  headerAlign: Align
  groupFill: string
  groupText: ThemeText
  noteText: ThemeText
  /** Fill of the label column in a `fields` table. */
  labelFill: string
  outerBorder: Border
  innerBorder: Border
  cellPadding: { top: number; bottom: number; left: number; right: number }
  /** Spacer row height, points. */
  spacerHeight: number
  /** Fill of a missing figure's placeholder box. */
  placeholderFill: string
  placeholderBorder: Border
  placeholderText: ThemeText
}

/** The look of a document: one application can be serif with a green band, another sans-serif and grey. */
export interface Theme {
  /** Body text: font, size, colour, spacing. */
  body: Required<Pick<ThemeText, 'font' | 'size' | 'color'>> & ThemeText
  title: ThemeText
  subtitle: ThemeText
  /** Heading levels 1 to 6, in order. */
  headings: readonly ThemeText[]
  caption: ThemeText
  note: ThemeText
  small: ThemeText
  header: ThemeText
  footer: ThemeText
  link: { color: string; underline: boolean }
  table: TableLook
  /** Caption words and the separator after the number: `"Table 3: Results"`. */
  captions: { table: string; figure: string; separator: string }
  /** A blank rule's width, points. */
  blankWidth: number
  /** Default page set-up for sections that do not give one. */
  page: { paper: PaperName | PaperSize; orientation: 'portrait' | 'landscape'; margins: Margins }
}

/** A theme as given: any part may be left out and is taken from the default. */
export type ThemeInput = {
  [K in keyof Theme]?: Theme[K] extends readonly ThemeText[]
    ? readonly ThemeText[]
    : Theme[K] extends object
      ? { [P in keyof Theme[K]]?: Theme[K][P] extends object ? Partial<Theme[K][P]> : Theme[K][P] }
      : Theme[K]
}

export interface Document {
  metadata?: Metadata
  theme?: ThemeInput
  sections: readonly Section[]
}
