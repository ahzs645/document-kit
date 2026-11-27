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
  /**
   * Character spacing (tracking), points: added after every character, negative to condense (Word's `w:spacing` on a
   * run). Tracking given in em is `em × size`: 0.05 em at 12 pt is 0.6.
   */
  characterSpacing?: number
  /**
   * A character style the document defines (`document.styles.character`, by id): its look applies under this run's
   * own. Unknown ids stop the writer with the id named.
   */
  characterStyle?: string
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
 *
 * `show: 'placeholder'` prints the bracketed name instead of a rule (`[metals run date]`), in the theme's
 * `placeholderMark` look (a yellow highlight by default), so an unfilled value inside a sentence cannot be missed. It
 * is listed to complete either way.
 */
export interface BlankRun extends TextStyle {
  type: 'blank'
  what?: string
  /** Points; the theme's blank width when absent. Not used by a placeholder. */
  width?: number
  source?: string
  /** `rule` (default): a rule to write on; `placeholder`: `[what]`, marked. */
  show?: 'rule' | 'placeholder'
}

/**
 * A tick box in a line of text. Word shows a real check box (a content control the reader can tick); the layout draws
 * a box, crossed when ticked, and `pageText` reads it as ☐ or ☒.
 */
export interface CheckboxRun extends TextStyle {
  type: 'checkbox'
  checked: boolean
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
  /**
   * What a missing slot does: `drop` (default) leaves the whole sentence out; `placeholder` prints the sentence with
   * each missing slot as its bracketed name, marked (`[visibility]`, see {@link BlankRun}). Listed to complete either
   * way.
   */
  missing?: 'drop' | 'placeholder'
}

export type SlotValues = Readonly<Record<string, string | number | null | undefined>>

export type Inline =
  | string
  | TextRun
  | FieldRun
  | { type: 'footnote'; id: string }
  | LineBreak
  | TabRun
  | BlankRun
  | InlineImage
  | SentenceRun
  | CheckboxRun

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
  /**
   * A paragraph's border: points between the text and the line (Word's `w:space`), a rule set a little under a
   * letterhead. 0 by default. Table and cell borders take none.
   */
  space?: number
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
  /**
   * Which of the document's text rules ({@link TextRule}) touch this block's text: absent, the automatic ones; a list
   * of rule ids, the automatic ones and these (rules with `auto: false` apply only where named); `false`, none.
   */
  textRules?: false | readonly string[]
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
  /** Named destination for an internal link written as #name. */
  bookmark?: string
  type: 'heading'
  /** 1 to 6. Written as Word's own Heading styles, so a table of contents and the navigation pane read them. */
  level: 1 | 2 | 3 | 4 | 5 | 6
  content: RichText
  align?: Align
}

/** Paragraph styles every theme defines. */
export type ParagraphStyleName = 'normal' | 'title' | 'subtitle' | 'caption' | 'note' | 'small'

/**
 * A paragraph's style: one the theme defines ({@link ParagraphStyleName}), or the id of a paragraph style the document
 * defines in `document.styles.paragraph`.
 */
export type ParagraphStyleRef = ParagraphStyleName | (string & {})

export interface ParagraphBlock extends BlockBase {
  /** Named destination for an internal link written as #name. */
  bookmark?: string
  type: 'paragraph'
  content: RichText
  /** A theme style, or the id of a paragraph style the document defines (`document.styles.paragraph`). */
  style?: ParagraphStyleRef
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
  /**
   * `lines` layout: print this row on the line of the row before it, after a gap ("E-mail: …   Phone: …"), when it
   * fits; when it does not, it moves whole to a line of its own (its spaces and hyphens do not break). Ignored on the
   * first row and in `table` layout.
   */
  sameLine?: boolean
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
  /** `lines` layout: space after each row, points (the body's paragraph spacing when absent). */
  rowSpacing?: number
  /** The gap before a `sameLine` row, points; 24 by default. */
  sameLineGap?: number
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
  /** This cell's inner padding, points, over the table's `cellPadding` side by side. */
  padding?: { top?: number; bottom?: number; left?: number; right?: number }
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
  /** Line height of the table's text as a multiple of single spacing; the theme's table line height when absent. */
  line?: number
  /** Keep every row whole (no row breaks across a page). */
  cantSplit?: boolean
  /** Space after this table, points, over the theme's `table.spaceAfter`; 0 for none. */
  spaceAfter?: number
  /**
   * Keep the whole table on one page: every row whole and kept with the next. A table taller than a page still breaks
   * (Word's rule as well), between rows. A row's own `keepWithNext` or `cantSplit` overrides it for that row.
   */
  keepTogether?: boolean
  /** Turn off repeating header rows for this table. */
  repeatHeader?: boolean
  /** Alignment of body cells by column, when a cell does not say. */
  columnAlign?: readonly (Align | undefined)[]
  /**
   * Let the columns fit their content (Word's AutoFit to contents, `w:tblLayout autofit`), the column shares only a
   * starting point. Fixed by default: the columns keep their shares whatever they hold.
   */
  autoFit?: boolean
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

/**
 * A signature block: a label, the signature picture (fitted to a box, sitting on the rule) or room to sign by hand,
 * the signing rule, the name (and title) under it, and lines under that. It is kept on one page unless `keepTogether`
 * is false.
 */
export interface SignatureBlock extends BlockBase {
  type: 'signature'
  label?: RichText
  /** `above` (default): the label heads the block; `below`: it is printed under the name and lines. */
  labelPosition?: 'above' | 'below'
  /** The label's look; bold by default. */
  labelStyle?: TextStyle
  /** Space between the label and the picture or signing room, points; 4 by default. */
  labelGap?: number
  /** The signature as a picture. Absent or null: room is left to sign by hand. */
  image?: ImageData | null
  /** The box the signature picture is fitted into, keeping its shape, points (140 × 55 by default). */
  imageBox?: { width: number; height: number }
  /** Enlarge a picture smaller than its box (true by default). */
  imageUpscale?: boolean
  /** Printed under the rule. Null is a blank to complete. */
  name?: RichText | null
  /** The signer's title, printed with the name: on its line (`titleLayout: 'sameLine'`) or on the next. */
  title?: RichText | null
  /** `sameLine` (default): "Name, Title"; `nextLine`: the title on its own line under the name. */
  titleLayout?: 'sameLine' | 'nextLine'
  /** Between the name and the title on one line; `", "` by default. */
  titleSeparator?: string
  /** The name's (and title's) look; bold by default. */
  nameStyle?: TextStyle
  /** Rows under the name: "Title", "Date". A null value is a blank to complete. A row with an empty label prints its value alone. */
  lines?: readonly FieldRow[]
  align?: 'left' | 'center' | 'right'
  /** Length of the signing rule, points; 216 by default. */
  ruleWidth?: number
  /** The rule's thickness, points; 0.75 by default. */
  ruleThickness?: number
  /** The rule's colour; the body text's by default. */
  ruleColor?: string
  /** Between the rule and the name under it, points; 2 by default. */
  ruleGap?: number
  /** Room to sign by hand when there is no picture, points; 36 by default. */
  signingSpace?: number
  /** Space above the block, points; 12 by default. */
  spaceAbove?: number
  /** Space under the block, points; 2 by default. */
  spaceBelow?: number
  /** Keep the whole block on one page (true by default). */
  keepTogether?: boolean
}

/**
 * A tick-box question: its options, each with a box, the one chosen ticked. `chosen` null (or absent) is unanswered:
 * every box empty, and listed to complete. `answer` adds a free-text answer line under it (null: a blank to complete;
 * absent: no answer line). A hand edit can change which option is ticked and the answer (see the recipes' edits).
 */
export interface ChoicesBlock extends BlockBase {
  type: 'choices'
  /** The question. */
  label?: RichText
  /** The options, as printed. */
  options: readonly string[]
  /** The option ticked: one of `options`, or several for a question that takes more than one. */
  chosen?: string | readonly string[] | null
  answer?: RichText | null
  /** `inline` (default): the boxes on one line after the question; `list`: one option a line under it. */
  layout?: 'inline' | 'list'
  /** Drawn as a ruled box: the question (and answer) at the left, the tick boxes at the right. */
  boxed?: boolean
}

/**
 * Blocks pinned to the foot of the page the block falls on, out of the text's flow: a disclaimer at the bottom of the
 * first page, a contact block at the bottom of the last. Written as a Word text frame anchored to the bottom margin
 * (consecutive paragraphs sharing one frame are one box). It holds headings, paragraphs, lists, `lines` fields, a
 * signature, tick boxes, spacers and pictures; not tables. The text of the page flows above it; text that would run
 * into it goes on to the next page.
 */
export interface PinnedBlock extends BlockBase {
  type: 'pinned'
  /** Where: the foot of the page (the bottom margin), the only place for now. */
  at?: 'bottom'
  blocks: readonly Block[]
  /** Empty room kept under the blocks inside the frame, points, so the last line clears the margin; 0 by default. */
  spaceBelow?: number
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

/** A workbook as bytes: the `.xlsx` file itself. */
export interface WorkbookData {
  bytes: Uint8Array
}

/**
 * One sheet of an Excel workbook (`.xlsx`), printed as a native table: the sheet's column widths, merged cells, fonts,
 * fills, borders, number formats (currency, percent, dates), wrapped text and alignment. A formula prints its cached
 * value; a formula the file holds no value for (a workbook written by openpyxl, say) is evaluated from the sheet,
 * other sheets of the workbook included (SUM, arithmetic, cell and range references, and most of Excel's common
 * functions). `formulas: 'evaluate'` recalculates every formula and ignores cached values.
 *
 * The leading `headerRows` rows repeat at the top of each page the table runs onto. Give the section a page set-up to
 * print the sheet as the workbook does (`{ paper: 'tabloid', orientation: 'landscape' }` for 17 × 11 in).
 *
 * A null `workbook` prints the placeholder and is listed as still to complete.
 */
export interface SpreadsheetBlock extends BlockBase {
  type: 'spreadsheet'
  workbook: WorkbookData | null
  /** The sheet by tab name, or by 0-based position. The first visible sheet when absent. */
  sheet?: string | number
  /**
   * The cells printed, an A1 range (`A1:F60`). Absent: the sheet's print area when the workbook names one, else from A1
   * to the last row and column holding a value or a fill.
   */
  range?: string
  /**
   * How many leading rows of the range repeat at the top of each page. Absent: the sheet's print titles (rows to
   * repeat) when the workbook names them, else its frozen rows, else none.
   */
  headerRows?: number
  /** `cached` (default): the values the file stored, evaluating only formulas without one; `evaluate`: every formula. */
  formulas?: 'cached' | 'evaluate'
  /**
   * Print scale, percent, as Excel's page set-up: text, row heights and column widths together. Absent: the sheet's
   * own, fitted to the text width when the sheet is set to fit its page width. Columns still wider than the text width
   * are narrowed to fit, the text keeping its size.
   */
  scale?: number
  /** Draw the sheet's grid lines as hairline cell borders. The sheet's own print setting when absent. */
  gridLines?: boolean
  align?: 'left' | 'center' | 'right'
  caption?: Caption
  /** Above (default) or below the sheet. */
  captionPosition?: 'above' | 'below'
  note?: RichText
  /** Printed in place of the sheet when the workbook is missing. */
  placeholder?: RichText
  /** Space after the sheet, points, as a table's. */
  spaceAfter?: number
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
  | SpreadsheetBlock
  | FigureBlock
  | FigureGridBlock
  | SignatureBlock
  | ChoicesBlock
  | PinnedBlock
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
  /** Character spacing (tracking), points; see {@link TextStyle.characterSpacing}. */
  characterSpacing?: number
  /**
   * A heading level: keep it on the page of what follows it (and its own lines together). True by default; false lets a
   * heading end a page, as a document whose own heading style does not keep with the next sets it.
   */
  keepWithNext?: boolean
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
  /** Least height of a body, group or note row, points (a row's own `height` wins); none when absent. */
  rowHeight?: number
  /** Least height of a header row, points (a row's own `height` wins); none when absent. */
  headerRowHeight?: number
  /** Line height of table text as a multiple of single spacing; 1 by default. */
  line?: number
  /** Spacer row height, points. */
  spacerHeight: number
  /**
   * Space between a table (with its caption or note) and the block after it, points. Applies in the body, not
   * inside a table cell, pinned frame, header or footer. A table's own `spaceAfter` wins.
   */
  spaceAfter?: number
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
  /** How a placeholder (`[name]` for a value nobody gave, see {@link BlankRun}) is marked; a yellow highlight by default. */
  placeholderMark: TextStyle
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

/**
 * A paragraph style the document defines, written as a named Word style (it shows in Word's style gallery and keeps
 * its name when the document is edited) and referenced by a paragraph's `style`. `text` is the look: font, size,
 * colour, bold, italic, caps, tracking, alignment, spacing before and after, line height and keep-with-next.
 */
export interface ParagraphStyleDefinition {
  /** The Word style id: a letter, then letters and digits (`AffiliationDepartment`). Not one of the theme's ids. */
  id: string
  /** The name Word shows; the id when absent. */
  name?: string
  /**
   * The style it builds on: a theme style (`normal`, `title`, `subtitle`, `caption`, `note`, `small`, `heading1` to
   * `heading6`) or another style the document defines, by id. `normal` when absent.
   */
  basedOn?: string
  /** The style of the paragraph after it, when one is typed in Word: a theme style or a defined id. */
  next?: string
  text?: ThemeText
  indent?: Indent
  borders?: { top?: Border; bottom?: Border; left?: Border; right?: Border }
  /** Fill behind the paragraph, a hex colour. */
  shading?: string
  keepLinesTogether?: boolean
}

/** A character style the document defines, referenced by a run's `characterStyle`. */
export interface CharacterStyleDefinition {
  /** The Word style id: a letter, then letters and digits. */
  id: string
  name?: string
  /** Another character style the document defines, by id. */
  basedOn?: string
  text: Omit<TextStyle, 'characterStyle'>
}

/** The named styles a document defines beyond the theme's. */
export interface DocumentStyles {
  paragraph?: readonly ParagraphStyleDefinition[]
  character?: readonly CharacterStyleDefinition[]
}

export interface Document {
  /** Paragraphs of each note, keyed by a permanent ID. Numbers follow first use in the body. */
  footnotes?: Readonly<Record<string, readonly string[]>>
  metadata?: Metadata
  theme?: ThemeInput
  /** Named paragraph and character styles beyond the theme's, for blocks and runs to reference by id. */
  styles?: DocumentStyles
  /**
   * Automatic text touches, stated as data and applied in order by the writer to the text's ordinary runs: links,
   * terms in italics, raised footnote markers, formula digits set low. See {@link TextRule} and `docs/model.md`.
   */
  textRules?: readonly TextRule[]
  sections: readonly Section[]
}

/* ------------------------------------------------------------------ text rules */

/** The kinds of text a rule may be limited to (`in`). Text in a header or footer is `header` or `footer`. */
export type TextPlace =
  | 'paragraph'
  | 'heading'
  | 'list'
  | 'label'
  | 'value'
  | 'cell'
  | 'caption'
  | 'note'
  | 'signature'
  | 'choice'
  | 'header'
  | 'footer'

/**
 * One automatic text touch: what to find, and what to do with what is found. Exactly one of `pattern`, `terms` or
 * `builtIn` says what to find:
 *
 * - `pattern`: a regular expression from a safe subset (see `docs/model.md`, "Text rules": no back-references or
 *   look-arounds, no repeat inside a repeat, `*` and `+` repeat at most 100 times);
 * - `terms`: words or phrases, matched whole (longest first);
 * - `builtIn`: `links` (web and e-mail addresses become links), `trailingMarker` ("Total a" → the `a` raised),
 *   `leadingMarker` ("a Dried at 105 °C" → the `a` raised), `formulaDigits` (the digits of the listed formulas set
 *   low: CaCO3 → CaCO₃).
 *
 * What is done: `style` on the match (or on its `group`), and/or `link: true` (the match becomes a link, if it reads
 * as an address of an allowed scheme). A built-in has its own default (links: a link; markers: superscript; formula
 * digits: subscript), which `style` replaces.
 */
export interface TextRule {
  /** The rule's name, for a block's or a filter's list of rules to use. */
  id?: string
  pattern?: string
  /** Pattern flags: `i` (ignore case) is the only one. */
  flags?: string
  terms?: readonly string[]
  /** Terms: match regardless of case. */
  ignoreCase?: boolean
  builtIn?: 'links' | 'trailingMarker' | 'leadingMarker' | 'formulaDigits'
  /** `pattern`: the capture group, by number or name, the style or link goes on; the whole match when absent. */
  group?: number | string
  style?: TextStyle
  link?: boolean
  /** `links` and `link: true`: the URI schemes made into links; `https`, `http` and `mailto` by default. */
  schemes?: readonly string[]
  /** `links`: also link bare domains (`example.org/path`) on these top-level domains; a common list by default, `[]` for none. */
  domains?: readonly string[]
  /** `trailingMarker`, `leadingMarker`: the marker letters, as a character list with ranges (`a-h`); `a-z` by default. */
  markers?: string
  /** `formulaDigits`: the formulas whose digits are set low (`CaCO3`, `NO3`). Required. */
  formulas?: readonly string[]
  /** Only in these kinds of text. Everywhere when absent. */
  in?: readonly TextPlace[]
  /** Applied wherever the rule's places allow (true, the default), or only where a block or filter names it (false). */
  auto?: boolean
}
