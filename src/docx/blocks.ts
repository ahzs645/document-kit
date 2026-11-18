/** Every kit block as docx paragraphs and tables. Custom blocks are expanded before this runs. */
import {
  HeadingLevel,
  LineRuleType,
  Paragraph,
  Table,
  TableOfContents,
} from 'docx'
import { fitImage, isRichText, plainText, resolveImage } from '../model/index'
import type {
  Block,
  FieldsBlock,
  FigureBlock,
  FigureGridBlock,
  HeadingBlock,
  ListBlock,
  ListItem,
  RichText,
  SignatureBlock,
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
import { writeTable } from './table'
import { border, twips } from './units'

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
  for (const block of blocks) {
    if (block.type === 'pageBreak') {
      breakNext = true
      continue
    }
    const breakBefore = (breakNext || block.pageBreakBefore === true) && context.part === 'body'
    const written = writeBlock(block, context, breakBefore)
    if (written.length) breakNext = false
    else if (breakBefore) breakNext = true
    out.push(...written)
  }
  // A break with nothing after it still starts a new page.
  if (breakNext) out.push(gapParagraph(1, true))
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
    case 'figure':
      return writeFigure(block, context, breakBefore)
    case 'figureGrid':
      return writeFigureGrid(block, context, breakBefore)
    case 'signature':
      return writeSignature(block, context, breakBefore)
    case 'toc':
      return writeToc(block, context, breakBefore)
    case 'spacer':
      if (block.height <= 0) return breakBefore ? [gapParagraph(1, true)] : []
      return [gapParagraph(block.height, breakBefore)]
    case 'pageBreak':
      return [gapParagraph(1, true)]
    case 'custom':
      throw new Error(`document-kit: custom block "${block.kind}" reached the Word writer unexpanded`)
  }
}

function writeHeading(block: HeadingBlock, context: WriterContext, breakBefore: boolean): Paragraph {
  const level = Math.min(6, Math.max(1, Math.round(block.level))) as HeadingBlock['level']
  return new Paragraph(
    compact({
      heading: HEADINGS[level - 1],
      alignment: alignment(block.align),
      pageBreakBefore: breakBefore || undefined,
      children: inlineChildren(block.content, context, compact({ source: block.source })),
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
  block.rows.forEach((row, index) => {
    const style = compact({ size })
    const label = inlineChildren(row.label, context, { style: { ...style, bold: true } })
    const gap =
      tabAt === undefined
        ? inlineChildren(separator, context, { style: { ...style, bold: true } })
        : inlineChildren([separator.trimEnd(), { type: 'tab' }], context, { style: { ...style, bold: true } })
    const value =
      row.value === null
        ? blankChildren(context, style)
        : inlineChildren(row.value, context, compact({ style, source: row.source }))
    out.push(
      new Paragraph(
        compact({
          keepNext: index < block.rows.length - 1 || undefined,
          pageBreakBefore: (breakBefore && !titleFirst && index === 0) || undefined,
          tabStops: tabAt === undefined ? undefined : [{ type: 'left' as const, position: twips(tabAt) }],
          indent: tabAt === undefined ? undefined : { left: twips(tabAt), hanging: twips(tabAt) },
          children: [...label, ...gap, ...value],
        }),
      ),
    )
  })
  return out
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

function writeSignature(block: SignatureBlock, context: WriterContext, breakBefore: boolean): Paragraph[] {
  const out: Paragraph[] = []
  const align = block.align ?? 'left'
  const ruleWidth = Math.min(block.ruleWidth ?? 216, context.width)
  const spare = Math.max(0, context.width - ruleWidth)
  const indent = align === 'center' ? { left: twips(spare / 2), right: twips(spare / 2) } : { right: twips(spare) }
  const first = () => (breakBefore && out.length === 0) || undefined
  if (block.label !== undefined) {
    out.push(
      new Paragraph(
        compact({
          alignment: alignment(align),
          keepNext: true,
          pageBreakBefore: first(),
          spacing: { before: twips(12), after: twips(4) },
          children: inlineChildren(block.label, context, { style: { bold: true } }),
        }),
      ),
    )
  }
  if (block.image) {
    const box = block.imageBox ?? { width: 140, height: 55 }
    const size = fitImage(resolveImage(block.image), box, true)
    out.push(
      new Paragraph(
        compact({
          alignment: alignment(align),
          keepNext: true,
          pageBreakBefore: first(),
          spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO },
          children: [imageRun(block.image, size.width, size.height, 'Signature')],
        }),
      ),
    )
  } else {
    out.push(
      new Paragraph(
        compact({
          keepNext: true,
          pageBreakBefore: first(),
          spacing: { before: 0, after: 0, line: twips(block.signingSpace ?? 36), lineRule: LineRuleType.EXACT },
          children: [],
        }),
      ),
    )
  }
  // The signing rule is the top border of the name's paragraph, as wide as the rule.
  const rule = border({ style: 'single', width: 0.75, color: context.theme.body.color })
  out.push(
    new Paragraph(
      compact({
        alignment: alignment(align),
        indent,
        keepNext: (block.lines?.length ?? 0) > 0 || undefined,
        border: rule ? { top: { ...rule, space: 2 } } : undefined,
        spacing: { before: 0, after: twips(2) },
        children:
          block.name === null || block.name === undefined
            ? []
            : inlineChildren(block.name, context, compact({ style: { bold: true }, source: block.source })),
      }),
    ),
  )
  const lines = block.lines ?? []
  lines.forEach((line, index) => {
    const label = inlineChildren(line.label, context, {})
    const separator = plainText(line.label).trim() && !/[:]\s*$/.test(plainText(line.label)) ? ': ' : ' '
    out.push(
      new Paragraph(
        compact({
          alignment: alignment(align),
          indent,
          keepNext: index < lines.length - 1 || undefined,
          spacing: { before: 0, after: twips(2) },
          children: [
            ...label,
            ...inlineChildren(separator, context, {}),
            ...(line.value === null
              ? blankChildren(context, {}, ruleWidth * 0.6)
              : inlineChildren(line.value, context, compact({ source: line.source }))),
          ],
        }),
      ),
    )
  })
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

