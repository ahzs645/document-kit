/** Every kit block as docx paragraphs and tables. Custom blocks are expanded before this runs. */
import {
  FrameAnchorType,
  FrameWrap,
  HeadingLevel,
  HeightRule,
  HorizontalPositionAlign,
  HorizontalPositionRelativeFrom,
  ImageRun,
  LineRuleType,
  Paragraph,
  Table,
  TableOfContents,
  TextRun,
  TextWrappingType,
  VerticalPositionAlign,
  VerticalPositionRelativeFrom,
  type IFrameOptions,
  type IParagraphOptions,
  type ParagraphChild,
} from 'docx'
import { fitImage, isRichText, plainText, resolveImage } from '../model/index'
import { choiceBlocks } from '../model/choices'
import type {
  Block,
  FieldRow,
  FieldsBlock,
  FigureBlock,
  FigureGridBlock,
  FloatingImageBlock,
  FrameBlock,
  HeadingBlock,
  ListBlock,
  Inline,
  ListItem,
  PinnedBlock,
  RichText,
  SignatureBlock,
  SpreadsheetBlock,
  TableBlock,
  TableRow,
  TocBlock,
} from '../model/types'
import type { WriterContext } from './context'
import { blankChildren, croppedPixels, imageRun, inlineChildren } from './inline'
import {
  alignment,
  captionParagraph,
  compact,
  gapParagraph,
  styledParagraph,
  TOC_TITLE_STYLE_ID,
  writeParagraph,
} from './paragraphs'
import { bookmarked } from './bookmarks'
import { spreadsheetTable } from './spreadsheet'
import { autoLineTables, tableGaps, writeTable } from './table'
import { border, pixels, twips } from './units'

export type FileChild = Paragraph | Table | TableOfContents

const HEADINGS = [
  HeadingLevel.HEADING_1,
  HeadingLevel.HEADING_2,
  HeadingLevel.HEADING_3,
  HeadingLevel.HEADING_4,
  HeadingLevel.HEADING_5,
  HeadingLevel.HEADING_6,
] as const

/** Writes blocks in order. A page break block moves the next block onto a new page. */
export function writeBlocks(blocks: readonly Block[], context: WriterContext): FileChild[] {
  const out: FileChild[] = []
  let breakNext = false
  for (const [index, block] of blocks.entries()) {
    if (block.type === 'pageBreak') {
      breakNext = true
      continue
    }
    const breakBefore = (breakNext || block.pageBreakBefore === true) && context.part === 'body' && !context.frame
    const written = writeBlock(block, context, breakBefore)
    // A table's trailing space is pointless when the next block starts a new page, and harmful when the table
    // ends at the foot of its page: the empty paragraph would spill onto a blank page of its own.
    const next = blocks[index + 1]
    const nextStartsPage = !next || next.type === 'pageBreak' || next.pageBreakBefore === true
    const last = written[written.length - 1]
    if (nextStartsPage && last && tableGaps.has(last)) written.pop()
    if (written.length) breakNext = false
    else if (breakBefore) breakNext = true
    out.push(...written)
  }
  // A break with nothing after it still starts a new page.
  if (breakNext && !context.frame) out.push(gapParagraph(1, true))
  return out
}

const tableBlocks = (blocks: NonNullable<import('../model/types').TableCell['blocks']>, context: WriterContext) =>
  writeBlocks(blocks, context).filter((child): child is Paragraph | Table => !(child instanceof TableOfContents))

function writeBlock(block: Block, context: WriterContext, breakBefore: boolean): FileChild[] {
  switch (block.type) {
    case 'heading':
      return [writeHeading(block, context, breakBefore)]
    case 'paragraph': {
      const paragraph = writeParagraph(block, context, breakBefore)
      return paragraph ? [paragraph] : []
    }
    case 'list':
      return writeList(block, context, breakBefore)
    case 'fields':
      return writeFields(block, context, breakBefore)
    case 'table':
      return writeTable(block, context, tableBlocks, breakBefore)
    case 'spreadsheet':
      return writeSpreadsheet(block, context, breakBefore)
    case 'figure':
      return writeFigure(block, context, breakBefore)
    case 'figureGrid':
      return writeFigureGrid(block, context, breakBefore)
    case 'signature':
      return writeSignature(block, context, breakBefore)
    case 'choices':
    {
      const blocks = choiceBlocks(block, context.theme)
      if (breakBefore && blocks[0]) blocks[0] = { ...blocks[0], pageBreakBefore: true }
      return writeBlocks(blocks, context)
    }
    case 'pinned':
      return writePinned(block, context, breakBefore)
    case 'frame':
      return writeFrame(block, context, breakBefore)
    case 'floatingImage':
      return writeFloatingImage(block, context, breakBefore)
    case 'toc':
      return writeToc(block, context, breakBefore)
    case 'spacer':
      if (block.height <= 0) return breakBefore ? [gapParagraph(1, true)] : []
      return [gapParagraph(block.height, breakBefore, context.frame)]
    case 'pageBreak':
      return context.frame ? [] : [gapParagraph(1, true)]
    case 'custom':
      throw new Error(`document-kit: custom block "${block.kind}" reached the Word writer unexpanded`)
  }
}

function writeHeading(block: HeadingBlock, context: WriterContext, breakBefore: boolean): Paragraph {
  const level = Math.min(6, Math.max(1, Math.round(block.level))) as HeadingBlock['level']
  const children = inlineChildren(block.content, context, compact({ source: block.source }))
  return new Paragraph(
    compact({
      heading: HEADINGS[level - 1],
      alignment: alignment(block.align),
      pageBreakBefore: breakBefore || undefined,
      frame: context.frame,
      children: bookmarked(block.bookmark, children, context),
    }),
  )
}

/* ------------------------------------------------------------------ lists */

const asItem = (item: ListItem | RichText): ListItem => (isRichText(item) ? { content: item } : item)

function writeList(block: ListBlock, context: WriterContext, breakBefore: boolean): Paragraph[] {
  const out: Paragraph[] = []
  const visit = (
    items: readonly (ListItem | RichText)[],
    ordered: boolean,
    level: number,
    reference: string,
    instance: number,
  ) => {
    for (const raw of items) {
      const item = asItem(raw)
      out.push(
        new Paragraph(
          compact({
            numbering: { reference, level: Math.min(level, 8), instance },
            pageBreakBefore: (breakBefore && out.length === 0) || undefined,
            frame: context.frame,
            children: inlineChildren(item.content, context, compact({ source: block.source })),
          }),
        ),
      )
      if (item.items?.length) {
        const childOrdered = item.ordered ?? ordered
        // Nested items of the same kind share the list's numbering (Word restarts them under each item); a different
        // kind takes numbering of its own.
        const sameKind = childOrdered === ordered
        const childReference = sameKind ? reference : context.lists.reference(childOrdered, 1)
        visit(item.items, childOrdered, level + 1, childReference, sameKind ? instance : context.lists.instance())
      }
    }
  }
  const ordered = block.ordered === true
  visit(block.items, ordered, 0, context.lists.reference(ordered, block.start ?? 1), context.lists.instance())
  return out
}

/* ------------------------------------------------------------------ fields */

function writeFields(block: FieldsBlock, context: WriterContext, breakBefore: boolean): FileChild[] {
  const out: FileChild[] = []
  const titleFirst = block.title !== undefined
  if (block.title !== undefined) {
    out.push(
      new Paragraph(
        compact({
          keepNext: true,
          pageBreakBefore: breakBefore || undefined,
          frame: context.frame,
          children: inlineChildren(block.title, context, { style: { bold: true } }),
        }),
      ),
    )
  }
  const size = block.size
  if (block.layout === 'table') {
    const look = context.theme.table
    const share = Math.min(90, Math.max(10, block.labelWidth ?? 32))
    const rows: TableRow[] = block.rows.map((row) => ({
      cells: [
        { content: row.label, shading: look.labelFill, text: { bold: true } },
        {
          content: row.value === null ? [{ type: 'blank', ...(row.what ? { what: row.what } : {}) }] : row.value,
          ...(row.source ? { source: row.source } : {}),
        },
      ],
      cantSplit: true,
    }))
    const table: TableBlock = compact({ type: 'table' as const, rows, columnWidths: [share, 100 - share], size })
    out.push(...writeTable(table, context, tableBlocks, breakBefore && !titleFirst))
    return out
  }
  const separator = block.separator ?? ': '
  const tabAt = block.labelWidth
  // Rows that go on the line of the row before them join its paragraph.
  const lines: FieldRow[][] = []
  block.rows.forEach((row, index) => {
    const last = lines[lines.length - 1]
    if (row.sameLine && index > 0 && last) last.push(row)
    else lines.push([row])
  })
  const style = compact({ size })
  const rowChildren = (row: FieldRow, unbreakable: boolean): ParagraphChild[] => {
    const label = inlineChildren(keep(row.label, unbreakable), context, { style: { ...style, bold: true } })
    const gap =
      tabAt === undefined || unbreakable
        ? inlineChildren(keep(separator, unbreakable), context, { style: { ...style, bold: true } })
        : inlineChildren([separator.trimEnd(), { type: 'tab' }], context, { style: { ...style, bold: true } })
    const value =
      row.value === null
        ? blankChildren(context, style)
        : inlineChildren(keep(row.value, unbreakable), context, compact({ style, source: row.source }))
    return [...label, ...gap, ...value]
  }
  const gapSpaces = ' '.repeat(Math.max(1, Math.round((block.sameLineGap ?? 24) / ((size ?? context.theme.body.size) * 0.25))))
  lines.forEach((line, index) => {
    const children: ParagraphChild[] = []
    line.forEach((row, at) => {
      if (at > 0) children.push(new TextRun(compact({ text: gapSpaces, size: size === undefined ? undefined : size * 2 })))
      children.push(...rowChildren(row, at > 0))
    })
    out.push(
      new Paragraph(
        compact({
          keepNext: index < lines.length - 1 || undefined,
          pageBreakBefore: (breakBefore && !titleFirst && index === 0) || undefined,
          tabStops: tabAt === undefined ? undefined : [{ type: 'left' as const, position: twips(tabAt) }],
          indent: tabAt === undefined ? undefined : { left: twips(tabAt), hanging: twips(tabAt) },
          spacing: block.rowSpacing === undefined ? undefined : { after: twips(block.rowSpacing) },
          frame: context.frame,
          children,
        }),
      ),
    )
  })
  return out
}

const NO_BREAK_HYPHEN = '\u2011'

/**
 * Text that does not break: its spaces made no-break spaces and its hyphens no-break hyphens (written as
 * Word's own `w:noBreakHyphen`), so a row on another row's line moves whole when it does not fit.
 */
function keep(text: RichText, unbreakable: boolean): RichText {
  if (!unbreakable) return text
  const fix = (value: string) => value.replace(/ /g, '\u00a0').replace(/-/g, NO_BREAK_HYPHEN)
  const list = typeof text === 'string' ? [text] : Array.isArray(text) ? (text as readonly Inline[]) : [text as Inline]
  return list.map((inline): Inline => (typeof inline === 'string' ? fix(inline) : inline.type === 'text' ? { ...inline, text: fix(inline.text) } : inline))
}

/* ------------------------------------------------------------------ figures */

/** Room a caption line and its spacing take, points; what a full-page figure leaves for it. */
function captionRoom(context: WriterContext, lines: number): number {
  const caption = context.theme.caption
  const size = caption.size ?? context.theme.body.size
  return lines * size * 1.3 + (caption.before ?? 0) + (caption.after ?? 0)
}

function figureBox(block: FigureBlock, context: WriterContext) {
  if (block.fit === 'page') {
    const lines = (block.caption ? 2 : 0) + (block.note !== undefined ? 2 : 0)
    // A little slack: a picture one point too tall for the page pushes itself onto the next one.
    const height = context.page.contentHeight - captionRoom(context, lines) - 12
    return { width: context.width, height, upscale: true, explicitHeight: true }
  }
  const fit = block.fit ?? {}
  return {
    width: Math.min(fit.width ?? context.width, context.width),
    height: fit.height ?? context.page.contentHeight * 0.5,
    upscale: fit.upscale ?? false,
    explicitHeight: fit.height !== undefined,
  }
}

/* ------------------------------------------------------------------ floating pictures */

const EMU_PER_POINT = 12700

/**
 * A picture anchored out of the flow: an empty paragraph, one twip tall, holding a drawing placed against the page or
 * the margins, behind the text (or in front), with no wrapping, so the text takes no notice of it.
 */
function writeFloatingImage(block: FloatingImageBlock, context: WriterContext, breakBefore: boolean): FileChild[] {
  if (context.frame) throw new Error('document-kit: a floating picture cannot go in a pinned block or a frame')
  if (!block.image) return breakBefore ? [gapParagraph(1, true)] : []
  const info = resolveImage(block.image)
  const page = context.page
  const box = block.box === 'page' ? { x: 0, y: 0, width: page.width, height: page.height } : block.box
  for (const [key, value] of Object.entries(box)) {
    if (!Number.isFinite(value) || ((key === 'width' || key === 'height') && value <= 0)) {
      throw new Error(`document-kit: a floating picture's ${key} must be a ${key === 'width' || key === 'height' ? 'positive ' : ''}number of points`)
    }
  }
  const margin = block.box !== 'page' && block.relativeTo === 'margin'
  const run = new ImageRun({
    type: info.format === 'jpeg' ? 'jpg' : 'png',
    data: block.image.bytes,
    transformation: { width: pixels(box.width), height: pixels(box.height) },
    ...(block.alt ? { altText: { name: block.alt, description: block.alt, title: block.alt } } : {}),
    floating: {
      horizontalPosition: {
        relative: margin ? HorizontalPositionRelativeFrom.MARGIN : HorizontalPositionRelativeFrom.PAGE,
        offset: Math.round(box.x * EMU_PER_POINT),
      },
      verticalPosition: {
        relative: margin ? VerticalPositionRelativeFrom.MARGIN : VerticalPositionRelativeFrom.PAGE,
        offset: Math.round(box.y * EMU_PER_POINT),
      },
      behindDocument: block.behind !== false,
      allowOverlap: true,
      lockAnchor: true,
      layoutInCell: false,
      wrap: { type: TextWrappingType.NONE },
    },
  })
  return [
    new Paragraph(
      compact({
        spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
        pageBreakBefore: breakBefore || undefined,
        children: [run],
      }),
    ),
  ]
}

/** A box marking where a missing picture goes, with its placeholder text. */
function placeholderBox(text: RichText | undefined, width: number, height: number, context: WriterContext): Table[] {
  const look = context.theme.table
  const placeholder: TableBlock = {
    type: 'table',
    width: Math.min(100, (width / context.width) * 100),
    align: 'center',
    borders: {
      top: look.placeholderBorder,
      bottom: look.placeholderBorder,
      left: look.placeholderBorder,
      right: look.placeholderBorder,
    },
    rows: [
      {
        height,
        cantSplit: true,
        cells: [
          {
            content: text ?? '',
            align: 'center',
            verticalAlign: 'center',
            shading: look.placeholderFill,
            text: compact({
              size: look.placeholderText.size,
              color: look.placeholderText.color,
              italic: look.placeholderText.italic,
              bold: look.placeholderText.bold,
            }),
          },
        ],
      },
    ],
  }
  return writeTable(placeholder, context, tableBlocks, false).filter((child): child is Table => child instanceof Table)
}

/* ------------------------------------------------------------------ spreadsheets */

/** A sheet of a workbook as a table, numbered as the block's caption is; a missing workbook as a placeholder box. */
function writeSpreadsheet(block: SpreadsheetBlock, context: WriterContext, breakBefore: boolean): FileChild[] {
  if (!block.workbook) {
    const out: FileChild[] = []
    if (breakBefore) out.push(gapParagraph(1, true))
    const label = block.placeholder ?? (typeof block.sheet === 'string' ? `Spreadsheet: ${block.sheet}` : 'Spreadsheet')
    out.push(...placeholderBox(label, context.width, 72, context))
    return out
  }
  const table = spreadsheetTable(block, { width: context.width })
  autoLineTables.add(table)
  const number = context.captions.get(block)
  const captions = number === undefined ? context.captions : new Map([...context.captions, [table, number]])
  return writeTable(table, { ...context, captions }, tableBlocks, breakBefore)
}

function writeFigure(block: FigureBlock, context: WriterContext, breakBefore: boolean): FileChild[] {
  const out: FileChild[] = []
  const box = figureBox(block, context)
  const captionAbove = block.caption && block.captionPosition === 'above'
  const align = block.align ?? 'center'
  if (block.caption && captionAbove) {
    out.push(captionParagraph(block, block.caption, context, { keepNext: true, breakBefore, align }))
  }
  const breakHere = breakBefore && !captionAbove
  if (block.image) {
    const shown = croppedPixels(resolveImage(block.image), block.crop)
    const size = fitImage(shown, box, box.upscale)
    const alt = block.alt ?? (block.caption ? plainText(block.caption.text) : undefined)
    const look = compact({
      crop: block.crop,
      border: block.border
        ? { width: block.border.width ?? 0.5, color: block.border.color ?? context.theme.body.color }
        : undefined,
    })
    out.push(
      new Paragraph(
        compact({
          alignment: alignment(align),
          keepNext: (block.caption !== undefined && !captionAbove) || undefined,
          pageBreakBefore: breakHere || undefined,
          frame: context.frame,
          spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO },
          children: [imageRun(block.image, size.width, size.height, alt, look)],
        }),
      ),
    )
  } else {
    if (breakHere) out.push(gapParagraph(1, true))
    const height = box.explicitHeight ? box.height : Math.min(box.height, 144)
    out.push(...placeholderBox(block.placeholder, box.width, height, context))
  }
  if (block.caption && !captionAbove) out.push(captionParagraph(block, block.caption, context, { align }))
  if (block.note !== undefined) out.push(styledParagraph('note', block.note, context, compact({ align, source: block.source })))
  return out
}

function writeFigureGrid(block: FigureGridBlock, context: WriterContext, breakBefore: boolean): FileChild[] {
  const columns = Math.max(1, Math.round(block.columns))
  const look = context.theme.table
  const columnWidth = context.width / columns
  const inner = columnWidth - look.cellPadding.left - look.cellPadding.right
  const itemHeight = block.itemHeight ?? inner * 0.75
  const labelAbove = block.labelPosition === 'above'
  const none = { style: 'none' as const }
  const rows: TableRow[] = []
  for (let start = 0; start < block.items.length; start += columns) {
    const items = block.items.slice(start, start + columns)
    rows.push({
      cantSplit: true,
      cells: Array.from({ length: columns }, (_, index) => {
        const item = items[index]
        if (!item) return ''
        const label = item.label === undefined ? [] : [{ type: 'paragraph' as const, style: 'note' as const, align: 'center' as const, content: item.label }]
        const picture: Block = {
          type: 'figure',
          image: item.image,
          fit: { width: inner, height: itemHeight, upscale: block.upscale ?? false },
          ...(item.placeholder !== undefined ? { placeholder: item.placeholder } : {}),
          ...(item.alt !== undefined ? { alt: item.alt } : {}),
        }
        return { blocks: labelAbove ? [...label, picture] : [picture, ...label], verticalAlign: 'top' as const }
      }),
    })
  }
  const grid: TableBlock = {
    type: 'table',
    rows,
    borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
  }
  const out: FileChild[] = []
  out.push(...writeTable(grid, context, tableBlocks, breakBefore))
  if (block.caption) out.push(captionParagraph(block, block.caption, context, { align: 'center' }))
  return out
}

/* ------------------------------------------------------------------ signature */

/**
 * The signature block as paragraphs: the label, the picture (fitted to its box, sitting on the rule) or room to sign,
 * the rule as the top border of the name's paragraph (as long as `ruleWidth`), the name and title, the lines, and the
 * label when it goes below. Kept on one page unless `keepTogether` is false.
 */
function writeSignature(block: SignatureBlock, context: WriterContext, breakBefore: boolean): Paragraph[] {
  const items: Array<Record<string, unknown>> = []
  const align = block.align ?? 'left'
  const ruleWidth = Math.min(block.ruleWidth ?? 216, context.width)
  const spare = Math.max(0, context.width - ruleWidth)
  const indent =
    align === 'center'
      ? { left: twips(spare / 2), right: twips(spare / 2) }
      : align === 'right'
        ? { left: twips(spare) }
        : { right: twips(spare) }
  const labelStyle = block.labelStyle ?? { bold: true }
  const nameStyle = block.nameStyle ?? { bold: true }
  const labelBelow = block.labelPosition === 'below'
  const label = () =>
    items.push({
      alignment: alignment(align),
      spacing: { before: 0, after: twips(labelBelow ? 2 : (block.labelGap ?? 4)) },
      children: inlineChildren(block.label!, context, { style: labelStyle }),
    })
  if (block.label !== undefined && !labelBelow) label()
  if (block.image) {
    const box = block.imageBox ?? { width: 140, height: 55 }
    const size = fitImage(resolveImage(block.image), box, block.imageUpscale ?? true)
    items.push({
      alignment: alignment(align),
      spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO },
      children: [imageRun(block.image, size.width, size.height, 'Signature')],
    })
  } else {
    items.push({
      spacing: { before: 0, after: 0, line: twips(Math.max(1, block.signingSpace ?? 36)), lineRule: LineRuleType.EXACT },
      children: [],
    })
  }
  // The signing rule is the top border of the name's paragraph, as wide as the rule.
  const rule = border({ style: 'single', width: block.ruleThickness ?? 0.75, color: block.ruleColor ?? context.theme.body.color })
  const name = block.name === null || block.name === undefined ? [] : inlineChildren(block.name, context, compact({ style: nameStyle, source: block.source }))
  const title = block.title === null || block.title === undefined ? [] : inlineChildren(block.title, context, compact({ style: nameStyle, source: block.source }))
  const sameLine = (block.titleLayout ?? 'sameLine') === 'sameLine'
  const separator = name.length && title.length ? inlineChildren(block.titleSeparator ?? ', ', context, { style: nameStyle }) : []
  items.push({
    alignment: alignment(align),
    indent,
    ...(rule ? { border: { top: { ...rule, space: Math.max(0, Math.round(block.ruleGap ?? 2)) } } } : {}),
    spacing: { before: 0, after: twips(2) },
    children: sameLine ? [...name, ...separator, ...title] : name,
  })
  if (!sameLine && title.length) items.push({ alignment: alignment(align), indent, spacing: { before: 0, after: twips(2) }, children: title })
  for (const line of block.lines ?? []) {
    const labelText = plainText(line.label).trim()
    const separator = labelText ? (/[:]\s*$/.test(plainText(line.label)) ? ' ' : ': ') : ''
    items.push({
      alignment: alignment(align),
      indent,
      spacing: { before: 0, after: twips(2) },
      children: [
        ...(labelText ? inlineChildren(line.label, context, {}) : []),
        ...(separator ? inlineChildren(separator, context, {}) : []),
        ...(line.value === null
          ? blankChildren(context, {}, ruleWidth * 0.6)
          : inlineChildren(line.value, context, compact({ source: line.source }))),
      ],
    })
  }
  if (block.label !== undefined && labelBelow) label()
  const keepTogether = block.keepTogether !== false
  return items.map((item, index) => {
    const first = index === 0
    const last = index === items.length - 1
    const spacing = { ...(item.spacing ?? {}) } as Record<string, unknown>
    if (first) spacing.before = twips(block.spaceAbove ?? 12)
    if (last) spacing.after = twips(block.spaceBelow ?? 2)
    return new Paragraph(
      compact({
        ...item,
        spacing,
        keepNext: (keepTogether && !last) || undefined,
        keepLines: keepTogether || undefined,
        pageBreakBefore: (breakBefore && first) || undefined,
        frame: context.frame,
      }) as IParagraphOptions,
    )
  })
}

/* ------------------------------------------------------------------ pinned blocks */

const FRAMED = new Set<Block['type']>([
  'heading',
  'paragraph',
  'list',
  'fields',
  'signature',
  'choices',
  'spacer',
  'figure',
  'pageBreak',
  'pinned',
  'frame',
])

/** Refuses, naming it, a block a Word text frame cannot hold. */
function checkFramed(blocks: readonly Block[], holder: string): void {
  const refuse = (what: string) => {
    throw new Error(`document-kit: ${holder} cannot hold ${what}: Word's frames take paragraphs only`)
  }
  for (const inner of blocks) {
    if (!FRAMED.has(inner.type)) refuse(`a ${inner.type} block`)
    if (inner.type === 'fields' && inner.layout === 'table') refuse('fields in table layout')
    if (inner.type === 'choices' && inner.boxed) refuse('boxed tick boxes')
    if (inner.type === 'figure' && (!inner.image || inner.caption)) refuse('a figure without its picture, or with a caption')
    if (inner.type === 'pinned' || inner.type === 'frame') checkFramed(inner.blocks, holder)
  }
}

/**
 * Blocks pinned to the foot of the page: every paragraph they make carries one text frame anchored to the bottom
 * margin, centred across the text width, so Word (and the layout) set them as one box at the foot of the page the
 * pinned block falls on. Tables cannot sit in such a frame, so a block that would make one is refused.
 */
function writePinned(block: PinnedBlock, context: WriterContext, breakBefore: boolean): FileChild[] {
  if (context.frame) return writeBlocks(block.blocks, context)
  const refuse = (what: string) => {
    throw new Error(`document-kit: a pinned block cannot hold ${what}: Word's frames take paragraphs only`)
  }
  checkFramed(block.blocks, 'a pinned block')
  const frame: IFrameOptions = {
    type: 'alignment',
    width: twips(context.width),
    height: 20,
    rule: HeightRule.AUTO,
    anchor: { horizontal: FrameAnchorType.MARGIN, vertical: FrameAnchorType.MARGIN },
    alignment: { x: HorizontalPositionAlign.CENTER, y: VerticalPositionAlign.BOTTOM },
    wrap: FrameWrap.AROUND,
  }
  const out: FileChild[] = []
  if (breakBefore) out.push(gapParagraph(1, true))
  const inner = { ...context, frame }
  out.push(...writeBlocks(block.blocks, inner))
  if ((block.spaceBelow ?? 0) > 0) out.push(gapParagraph(block.spaceBelow!, false, frame))
  if (out.some((child) => child instanceof Table)) refuse('a table')
  return out
}

const FRAME_WRAP = {
  around: FrameWrap.AROUND,
  notBeside: FrameWrap.NOT_BESIDE,
  through: FrameWrap.THROUGH,
  none: FrameWrap.NONE,
} as const

/**
 * Blocks in a text frame placed on the page: every paragraph they make carries one frame (`w:framePr` with `w:x`,
 * `w:y` and `w:w`, anchored to the page or the margins), so Word and the layout set them as one box there. Inside a
 * pinned block or another frame, the blocks join that frame.
 */
function writeFrame(block: FrameBlock, context: WriterContext, breakBefore: boolean): FileChild[] {
  if (context.frame) return writeBlocks(block.blocks, context)
  checkFramed(block.blocks, 'a frame')
  for (const key of ['x', 'y', 'width'] as const) {
    if (!Number.isFinite(block[key])) throw new Error(`document-kit: a frame's ${key} must be a number of points`)
  }
  if (block.width <= 0) throw new Error("document-kit: a frame's width must be more than 0")
  const anchor = block.relativeTo === 'margin' ? FrameAnchorType.MARGIN : FrameAnchorType.PAGE
  const space = block.space === undefined ? undefined : twips(Math.max(0, block.space))
  const frame: IFrameOptions = {
    type: 'absolute',
    position: { x: twips(block.x), y: twips(block.y) },
    anchor: { horizontal: anchor, vertical: anchor },
    width: twips(block.width),
    ...(block.height !== undefined && block.height > 0
      ? { height: twips(block.height), rule: block.heightRule === 'exact' ? HeightRule.EXACT : HeightRule.ATLEAST }
      : {}),
    wrap: FRAME_WRAP[block.wrap ?? 'around'],
    ...(space !== undefined ? { space: { horizontal: space, vertical: space } } : {}),
  }
  const out: FileChild[] = []
  if (breakBefore) out.push(gapParagraph(1, true))
  out.push(...writeBlocks(block.blocks, { ...context, frame, width: block.width }))
  if (out.some((child) => child instanceof Table)) throw new Error("document-kit: a frame cannot hold a table: Word's frames take paragraphs only")
  return out
}

/* ------------------------------------------------------------------ table of contents */

function writeToc(block: TocBlock, context: WriterContext, breakBefore: boolean): FileChild[] {
  const out: FileChild[] = []
  const levels = Math.min(6, Math.max(1, Math.round(block.levels ?? 3)))
  if (block.title !== undefined) {
    out.push(
      new Paragraph(
        compact({
          style: TOC_TITLE_STYLE_ID,
          pageBreakBefore: breakBefore || undefined,
          frame: context.frame,
          children: inlineChildren(block.title, context, {}),
        }),
      ),
    )
  } else if (breakBefore) {
    out.push(gapParagraph(1, true))
  }
  out.push(
    new TableOfContents(plainText(block.title) || 'Contents', {
      hyperlink: true,
      headingStyleRange: `1-${levels}`,
    }),
  )
  return out
}

