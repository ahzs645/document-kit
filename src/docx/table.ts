/** Tables: the grid, merges, row roles, repeating header rows, borders, shading, captions and notes. */
import {
  HeightRule,
  Paragraph,
  Table,
  TableCell as DocxCell,
  TableLayoutType,
  TableRow as DocxRow,
  TextDirection,
  VerticalAlignTable,
  VerticalMergeType,
  WidthType,
  type ITableCellOptions,
} from 'docx'
import { tableGrid, type PlacedCell } from '../model/index'
import type { Align, RowRole, TableBlock, TableCell, TextStyle, ThemeText, VerticalAlign } from '../model/types'
import type { WriterContext } from './context'
import { blankChildren, inlineChildren } from './inline'
import { alignment, captionParagraph, compact, gapParagraph, styledParagraph, TABLE_STYLE_ID } from './paragraphs'
import { border, hex, lineMultiple, twips } from './units'

/** Writes nested blocks inside a cell. Set by the block writer, which tables call back into. */
export type BlockWriter = (blocks: TableCell['blocks'] & object, context: WriterContext) => Array<Paragraph | Table>

/** Splits `total` twips by `weights`, rounding so the parts add up to exactly `total`. */
export function splitWidths(total: number, weights: readonly number[]): number[] {
  const clean = weights.map((weight) => (Number.isFinite(weight) && weight > 0 ? weight : 0))
  const sum = clean.reduce((a, b) => a + b, 0)
  const shares = sum > 0 ? clean.map((weight) => weight / sum) : clean.map(() => 1 / clean.length)
  const out: number[] = []
  let used = 0
  let cumulative = 0
  shares.forEach((share) => {
    cumulative += share
    const edge = Math.round(total * cumulative)
    out.push(edge - used)
    used = edge
  })
  return out
}

const textStyleOf = (text: ThemeText | undefined): TextStyle =>
  compact({
    font: text?.font,
    size: text?.size,
    color: text?.color,
    bold: text?.bold,
    italic: text?.italic,
    allCaps: text?.allCaps,
    smallCaps: text?.smallCaps,
  })

const verticalAlign = (value: VerticalAlign | undefined) =>
  value === 'center' ? VerticalAlignTable.CENTER : value === 'bottom' ? VerticalAlignTable.BOTTOM : undefined

/** The leading rows Word repeats at the top of each page. */
export function repeatedRows(table: TableBlock): number {
  if (table.repeatHeader === false) return 0
  let count = 0
  for (const row of table.rows) {
    const repeat = row.repeat ?? row.role === 'header'
    if (!repeat) break
    count += 1
  }
  return count
}

function roleLook(role: RowRole, context: WriterContext) {
  const look = context.theme.table
  switch (role) {
    case 'header':
      return { fill: look.headerFill, text: textStyleOf(look.headerText), align: look.headerAlign as Align }
    case 'group':
      return { fill: look.groupFill, text: textStyleOf(look.groupText), align: (look.groupText.align ?? 'left') as Align }
    case 'note':
      return { fill: undefined, text: textStyleOf(look.noteText), align: (look.noteText.align ?? 'left') as Align }
    default:
      return { fill: undefined, text: {} as TextStyle, align: undefined }
  }
}

export function writeTable(
  table: TableBlock,
  context: WriterContext,
  writeBlocks: BlockWriter,
  breakBefore: boolean,
): Array<Paragraph | Table> {
  const grid = tableGrid(table)
  const look = context.theme.table
  const widthPoints = (context.width * Math.min(100, Math.max(1, table.width ?? 100))) / 100
  const columnTwips = splitWidths(twips(widthPoints), table.columnWidths ?? Array.from({ length: grid.columns }, () => 1))
  const padding = { ...look.cellPadding, ...compact(table.cellPadding ?? {}) }
  // The table text style carries the theme's size; a table's own size is written on its runs.
  const size = table.size
  const repeat = repeatedRows(table)
  const spanTwips = (column: number, span: number) =>
    columnTwips.slice(column, column + span).reduce((a, b) => a + b, 0)

  const cellBorders = (cell: TableCell) => {
    const borders = compact({
      top: border(cell.borders?.top),
      bottom: border(cell.borders?.bottom),
      left: border(cell.borders?.left),
      right: border(cell.borders?.right),
      diagonalDown: cell.diagonal === 'down' ? border(look.outerBorder) : undefined,
      diagonalUp: cell.diagonal === 'up' ? border(look.outerBorder) : undefined,
    })
    return Object.keys(borders).length ? borders : undefined
  }

  // A cell's own padding, side by side over the table's; written only for a cell that gives one.
  const cellMargins = (cell: TableCell) => {
    if (!cell.padding) return undefined
    const own = { ...padding, ...compact(cell.padding) }
    return { top: twips(own.top), bottom: twips(own.bottom), left: twips(own.left), right: twips(own.right) }
  }

  const rows = grid.rows.map((line, r) => {
    const row = table.rows[r]!
    const role = row.role ?? 'body'
    const roleStyle = roleLook(role, context)
    const cells: DocxCell[] = []
    let column = 0
    const shadingOf = (cell: TableCell) => {
      const fill = cell.shading ?? roleStyle.fill
      return fill ? { fill: hex(fill), color: 'auto', type: 'clear' as const } : undefined
    }
    for (const entry of line) {
      if (entry.kind === 'covered') {
        const { by } = entry
        cells.push(
          new DocxCell(
            compact({
              children: [new Paragraph({ style: TABLE_STYLE_ID, children: [] })],
              columnSpan: by.colSpan > 1 ? by.colSpan : undefined,
              verticalMerge: VerticalMergeType.CONTINUE,
              width: { size: spanTwips(by.column, by.colSpan), type: WidthType.DXA },
              shading: shadingOf(by.cell),
              borders: cellBorders(by.cell),
            }) as ITableCellOptions,
          ),
        )
        column = by.column + by.colSpan
        continue
      }
      const placed = entry.placed
      cells.push(writeCell(placed, row.role ?? 'body', roleStyle, column))
      column = placed.column + placed.colSpan
    }
    // A row shorter than the grid is padded with empty cells.
    while (column < grid.columns) {
      cells.push(
        new DocxCell({
          children: [new Paragraph({ style: TABLE_STYLE_ID, children: [] })],
          width: { size: columnTwips[column] ?? 0, type: WidthType.DXA },
        }),
      )
      column += 1
    }
    // A row's own height wins; else the theme's least height for header rows or for the others.
    const least = row.height ?? (role === 'header' ? look.headerRowHeight : look.rowHeight)
    const height =
      role === 'spacer'
        ? { value: twips(row.height ?? look.spacerHeight), rule: HeightRule.EXACT }
        : least !== undefined
          ? { value: twips(least), rule: HeightRule.ATLEAST }
          : undefined
    // Keep with next is Word's: `w:keepNext` on the paragraphs of the row's cells (the docx library puts it on every
    // one, the rows it adds under a merged cell too). The last row has no row to keep with.
    const last = r === table.rows.length - 1
    const keepNext = row.keepWithNext ?? (!last && (table.keepTogether === true || role === 'header'))
    return new DocxRow(
      compact({
        children: cells,
        tableHeader: r < repeat || undefined,
        cantSplit: row.cantSplit ?? (table.keepTogether || table.cantSplit) ?? undefined,
        keepNext: keepNext || undefined,
        height,
      }),
    )

    function writeCell(placed: PlacedCell, rowRole: RowRole, style: ReturnType<typeof roleLook>, at: number) {
      const cell = placed.cell
      const width = spanTwips(placed.column, placed.colSpan)
      const own = { ...padding, ...compact(cell.padding ?? {}) }
      const inner = Math.max(12, width / 20 - own.left - own.right)
      const cellAlign = cell.align ?? style.align ?? (rowRole === 'body' ? table.columnAlign?.[at] : undefined)
      const base: TextStyle = { ...(size === undefined ? {} : { size }), ...style.text, ...(cell.text ?? {}) }
      let children: Array<Paragraph | Table>
      if (rowRole === 'spacer') {
        children = [gapParagraph(1)]
      } else if (cell.blocks?.length) {
        children = writeBlocks(cell.blocks, { ...context, width: inner, inCell: true })
        if (!(children[children.length - 1] instanceof Paragraph)) children.push(new Paragraph({ children: [] }))
      } else {
        const content = cell.content
        const runs =
          content === null && context.nullCells === 'rule'
            ? blankChildren(context, base, Math.min(inner * 0.8, context.theme.blankWidth))
            : content === null || content === undefined
              ? []
              : inlineChildren(content, context, compact({ style: base, source: cell.source }))
        children = [
          new Paragraph(
            compact({
              style: TABLE_STYLE_ID,
              alignment: alignment(cellAlign),
              spacing: table.line === undefined ? undefined : { line: lineMultiple(table.line) },
              children: runs,
            }),
          ),
        ]
      }
      return new DocxCell(
        compact({
          children,
          columnSpan: placed.colSpan > 1 ? placed.colSpan : undefined,
          verticalMerge: placed.rowSpan > 1 ? VerticalMergeType.RESTART : undefined,
          width: { size: width, type: WidthType.DXA },
          shading: shadingOf(cell),
          verticalAlign: verticalAlign(cell.verticalAlign),
          textDirection:
            cell.textDirection === 'up'
              ? TextDirection.BOTTOM_TO_TOP_LEFT_TO_RIGHT
              : cell.textDirection === 'down'
                ? TextDirection.TOP_TO_BOTTOM_RIGHT_TO_LEFT
                : undefined,
          borders: cellBorders(cell),
          margins: cellMargins(cell),
        }) as ITableCellOptions,
      )
    }
  })

  const borders = { ...tableBorders(context), ...compact(table.borders ?? {}) }
  const docxTable = new Table(
    compact({
      rows,
      width: { size: columnTwips.reduce((a, b) => a + b, 0), type: WidthType.DXA },
      columnWidths: columnTwips,
      layout: table.autoFit ? TableLayoutType.AUTOFIT : TableLayoutType.FIXED,
      alignment: alignment(table.align),
      margins: {
        top: twips(padding.top),
        bottom: twips(padding.bottom),
        left: twips(padding.left),
        right: twips(padding.right),
      },
      borders: compact({
        top: border(borders.top),
        bottom: border(borders.bottom),
        left: border(borders.left),
        right: border(borders.right),
        insideHorizontal: border(borders.insideHorizontal),
        insideVertical: border(borders.insideVertical),
      }),
    }),
  )

  const out: Array<Paragraph | Table> = []
  const above = table.caption && (table.captionPosition ?? 'above') === 'above'
  if (table.caption && above) {
    out.push(captionParagraph(table, table.caption, context, { keepNext: true, breakBefore }))
  } else if (breakBefore) {
    out.push(gapParagraph(1, true))
  }
  out.push(docxTable)
  if (table.caption && !above) out.push(captionParagraph(table, table.caption, context))
  if (table.note !== undefined) out.push(styledParagraph('note', table.note, context, compact({ source: table.source })))
  // Space before the next block, so text does not sit on the table's last rule.
  const after = table.spaceAfter ?? context.theme.table.spaceAfter ?? 0
  if (after > 0 && context.part === 'body' && !context.inCell && !context.frame) out.push(gapParagraph(after))
  return out
}

function tableBorders(context: WriterContext) {
  const look = context.theme.table
  return {
    top: look.outerBorder,
    bottom: look.outerBorder,
    left: look.outerBorder,
    right: look.outerBorder,
    insideHorizontal: look.innerBorder,
    insideVertical: look.innerBorder,
  }
}
