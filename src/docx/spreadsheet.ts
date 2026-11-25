/**
 * A spreadsheet block as a kit table: one sheet of an `.xlsx`, read by the reamkit Excel reader, with its column
 * widths, merges, fonts, fills, borders, number formats, alignment and wrapped text. Formulas print their cached value,
 * or are evaluated when the file cached none ({@link sheetValues}).
 */
import type { Sheet, SheetDoc } from '../ream/core/ir/sheet'
import type {
  SheetRichRun,
  WorksheetCell,
  XlsxBorderEdge,
  XlsxCellXf,
  XlsxStyles,
} from '../ream/core/spreadsheet-model'
import { applyNumberFormat, numberFormatColorHex } from '../ream/core/number-format'
import { parseCellRef } from '../ream/excel/cell-reference'
import { parseAreaRef, parseTitleRowRange } from '../ream/excel/defined-name-ref'
import { columnUnitTwips } from '../ream/excel/print-model'
import { makeColWidthPt } from '../ream/excel/sheet-drawing'
import { readXlsxToSheetDoc } from '../ream/excel/xlsx-reader'
import type {
  Align,
  Border,
  Inline,
  SpreadsheetBlock,
  TableBlock,
  TableCell,
  TableRow,
  TextStyle,
  VerticalAlign,
} from '../model/types'
import { sheetValues, type CellValue, type SheetValues } from './sheetValues'

export interface SpreadsheetTableOptions {
  /** The width the sheet is printed in, points: the text width. Columns wider in all are narrowed to it. */
  width: number
}

/** The cells a sheet range covers, 0-based and inclusive. */
interface Area {
  r0: number
  c0: number
  r1: number
  c1: number
}

const workbooks = new WeakMap<Uint8Array, SheetDoc>()

/** Reads a workbook once per byte array, so several sheets of one file are parsed once. */
function readWorkbook(bytes: Uint8Array): SheetDoc {
  let doc = workbooks.get(bytes)
  if (!doc) {
    doc = readXlsxToSheetDoc(bytes)
    workbooks.set(bytes, doc)
  }
  return doc
}

function pickSheet(doc: SheetDoc, wanted: SpreadsheetBlock['sheet']): number {
  if (typeof wanted === 'number') {
    if (!Number.isInteger(wanted) || wanted < 0 || wanted >= doc.sheets.length) {
      throw new Error(`document-kit: the workbook has no sheet ${wanted} (it has ${doc.sheets.length})`)
    }
    return wanted
  }
  if (typeof wanted === 'string' && wanted !== '') {
    const index = doc.sheets.findIndex((sheet) => sheet.name.toLowerCase() === wanted.toLowerCase())
    if (index < 0) {
      throw new Error(
        `document-kit: the workbook has no sheet "${wanted}" (it has ${doc.sheets.map((s) => `"${s.name}"`).join(', ')})`,
      )
    }
    return index
  }
  const visible = doc.sheets.findIndex((sheet) => !sheet.hidden && !sheet.grid.chartSheet)
  return visible < 0 ? 0 : visible
}

function parseRange(range: string): Area {
  const text = range.replace(/\$/g, '').replace(/^.*!/, '').trim()
  const [a, b = a] = text.split(':')
  try {
    const start = parseCellRef(a!.trim())
    const end = parseCellRef(b!.trim())
    return {
      r0: Math.min(start.row, end.row),
      c0: Math.min(start.column, end.column),
      r1: Math.max(start.row, end.row),
      c1: Math.max(start.column, end.column),
    }
  } catch {
    throw new Error(`document-kit: "${range}" is not a sheet range (A1:F60)`)
  }
}

const definedName = (doc: SheetDoc, index: number, name: string) =>
  doc.definedNames.find((d) => d.name.toLowerCase() === name.toLowerCase() && d.localSheetId === index)?.value

/** Every cell holding a value, a fill or a border, merges included: what the sheet prints when no range is given. */
function usedArea(sheet: Sheet, styles: XlsxStyles, values: SheetValues, index: number): Area | undefined {
  let r1 = -1
  let c1 = -1
  for (const cell of sheet.grid.cells) {
    const xf = styles.cellXfs[cell.styleIndex ?? 0]
    const painted = xf !== undefined && (hasFill(styles, xf) || hasBorder(styles, xf))
    if (!painted && values.value(index, cell.row, cell.column).t === 'blank') continue
    r1 = Math.max(r1, cell.row)
    c1 = Math.max(c1, cell.column)
  }
  for (const merge of sheet.grid.merges) {
    r1 = Math.max(r1, merge.endRow)
    c1 = Math.max(c1, merge.endColumn)
  }
  return r1 < 0 ? undefined : { r0: 0, c0: 0, r1, c1 }
}

const hasFill = (styles: XlsxStyles, xf: XlsxCellXf) => {
  const fill = styles.fills[xf.fillId]
  return fill !== undefined && fill.patternType !== undefined && fill.patternType !== 'none' && fill.fgColorHex !== undefined
}

const hasBorder = (styles: XlsxStyles, xf: XlsxCellXf) => {
  const border = styles.borders[xf.borderId]
  return border !== undefined && [border.top, border.bottom, border.left, border.right].some((e) => edgeOf(e) !== undefined)
}

function edgeOf(edge: XlsxBorderEdge | undefined): Border | undefined {
  if (!edge?.style || edge.style === 'none') return undefined
  const color = edge.colorHex ?? '000000'
  switch (edge.style) {
    case 'hair':
      return { style: 'single', width: 0.25, color }
    case 'thin':
      return { style: 'single', width: 0.5, color }
    case 'medium':
      return { style: 'single', width: 1, color }
    case 'thick':
      return { style: 'single', width: 1.5, color }
    case 'double':
      return { style: 'double', width: 0.5, color }
    case 'dotted':
      return { style: 'dotted', width: 0.5, color }
    case 'mediumDashed':
    case 'mediumDashDot':
    case 'mediumDashDotDot':
    case 'slantDashDot':
      return { style: 'dashed', width: 1, color }
    default:
      return { style: 'dashed', width: 0.5, color }
  }
}

const HORIZONTAL: Record<string, Align | undefined> = {
  left: 'left',
  center: 'center',
  centerContinuous: 'center',
  right: 'right',
  justify: 'justify',
  distributed: 'justify',
  fill: 'left',
}

const VERTICAL: Record<string, VerticalAlign> = {
  top: 'top',
  center: 'center',
  bottom: 'bottom',
  justify: 'top',
  distributed: 'center',
}

/** Rounds away binary noise (0.1 + 0.2) before a number is formatted. */
const clean = (n: number) => (Number.isFinite(n) ? Number(n.toPrecision(15)) : n)

/** The value as the raw cell the number formats read. */
function rawOf(value: CellValue): { type: WorksheetCell['type']; raw: string } {
  switch (value.t) {
    case 'num':
      return { type: 'n', raw: String(clean(value.v)) }
    case 'str':
      return { type: 'str', raw: value.v }
    case 'bool':
      return { type: 'b', raw: value.v ? '1' : '0' }
    case 'err':
      return { type: 'e', raw: value.v }
    default:
      return { type: 'n', raw: '' }
  }
}

/**
 * The sheet as a kit table. Throws when the workbook cannot be read, or names no such sheet. A missing workbook is the
 * writer's to print as a placeholder.
 */
export function spreadsheetTable(block: SpreadsheetBlock, options: SpreadsheetTableOptions): TableBlock {
  if (!block.workbook) throw new Error('document-kit: a spreadsheet block needs its workbook')
  const doc = readWorkbook(block.workbook.bytes)
  const index = pickSheet(doc, block.sheet)
  const sheet = doc.sheets[index]!
  const { grid } = sheet
  const styles = doc.styles
  const values = sheetValues(doc, block.formulas ?? 'cached')

  const printArea = definedName(doc, index, '_xlnm.Print_Area')
  const areaFromName = printArea ? parseAreaRef(printArea) : undefined
  const area: Area | undefined = block.range
    ? parseRange(block.range)
    : areaFromName
      ? { r0: areaFromName.startRow, c0: areaFromName.startColumn, r1: areaFromName.endRow, c1: areaFromName.endColumn }
      : usedArea(sheet, styles, values, index)
  if (!area) return { type: 'table', rows: [{ cells: [''] }], columnWidths: [1] }

  // Hidden rows and columns are not printed, as in Excel.
  const hiddenColumns = new Set<number>()
  for (const column of grid.columns) {
    if (column.hidden) for (let c = column.min - 1; c <= column.max - 1; c += 1) hiddenColumns.add(c)
  }
  const rowInfo = new Map(grid.rowHeights.map((row) => [row.row, row]))
  const columns: number[] = []
  for (let c = area.c0; c <= area.c1; c += 1) if (!hiddenColumns.has(c)) columns.push(c)
  const rows: number[] = []
  for (let r = area.r0; r <= area.r1; r += 1) if (!rowInfo.get(r)?.hidden) rows.push(r)
  if (!columns.length || !rows.length) return { type: 'table', rows: [{ cells: [''] }], columnWidths: [1] }

  // Widths in points, the sheet's own (its Normal font's digit width, as Excel measures columns).
  const widthOf = makeColWidthPt(grid, columnUnitTwips(styles))
  const natural = columns.map((c) => widthOf(c))
  const total = natural.reduce((a, b) => a + b, 0)
  const setup = grid.pageSetup
  const fitWidth = grid.fitToPage === true && (setup?.fitToWidth ?? 1) >= 1
  const sheetScale = fitWidth
    ? Math.min(1, options.width / total)
    : setup?.scale !== undefined && setup.scale > 0
      ? setup.scale / 100
      : 1
  const scale = block.scale !== undefined && block.scale > 0 ? block.scale / 100 : sheetScale
  const printed = total * scale
  const width = Math.min(100, (printed / options.width) * 100)

  // Merges by their top-left cell, and the cells they cover.
  const merges = new Map<number, { rowSpan: number; colSpan: number; endRow: number; endColumn: number }>()
  const covered = new Set<number>()
  const key = (r: number, c: number) => r * 16384 + c
  for (const merge of grid.merges) {
    const r0 = Math.max(merge.startRow, area.r0)
    const c0 = Math.max(merge.startColumn, area.c0)
    const r1 = Math.min(merge.endRow, area.r1)
    const c1 = Math.min(merge.endColumn, area.c1)
    if (r0 > r1 || c0 > c1) continue
    const spanRows = rows.filter((r) => r >= r0 && r <= r1)
    const spanColumns = columns.filter((c) => c >= c0 && c <= c1)
    if (!spanRows.length || !spanColumns.length) continue
    const top = spanRows[0]!
    const left = spanColumns[0]!
    merges.set(key(top, left), {
      rowSpan: spanRows.length,
      colSpan: spanColumns.length,
      endRow: spanRows[spanRows.length - 1]!,
      endColumn: spanColumns[spanColumns.length - 1]!,
    })
    for (const r of spanRows) for (const c of spanColumns) if (r !== top || c !== left) covered.add(key(r, c))
  }

  const cellAt = new Map(grid.cells.map((cell) => [key(cell.row, cell.column), cell]))
  const rowStyle = new Map((grid.rowStyles ?? []).map((row) => [row.row, row.styleIndex]))
  const styleIndexAt = (r: number, c: number) => {
    const own = cellAt.get(key(r, c))?.styleIndex
    if (own !== undefined) return own
    const ofRow = rowStyle.get(r)
    if (ofRow !== undefined) return ofRow
    return grid.columnStyles?.find((col) => c >= col.min - 1 && c <= col.max - 1)?.styleIndex ?? 0
  }
  const xfAt = (r: number, c: number): XlsxCellXf | undefined => styles.cellXfs[styleIndexAt(r, c)]
  const gridLines = block.gridLines ?? grid.printOptions?.gridLines === true
  const gridRule: Border | undefined = gridLines ? { style: 'single', width: 0.25, color: 'A6A6A6' } : undefined

  const baseFont = styles.fonts[0]
  const textOf = (r: number, c: number, xf: XlsxCellXf | undefined): { runs: Inline[]; style: TextStyle; numeric: boolean; kind: CellValue['t'] } => {
    const font = styles.fonts[xf?.fontId ?? 0] ?? baseFont
    const style: TextStyle = {}
    if (font?.bold) style.bold = true
    if (font?.italic) style.italic = true
    if (font?.underline) style.underline = true
    if (font?.strike) style.strike = true
    // Excel's automatic colour is black, whatever the document's body text is.
    style.color = font?.colorHex ?? '000000'
    const size = font?.sizePt ?? baseFont?.sizePt ?? 11
    style.size = Math.round(size * scale * 2) / 2
    if (font?.name) style.font = font.name
    const value = values.value(index, r, c)
    if (value.t === 'blank') return { runs: [], style, numeric: false, kind: 'blank' }
    const cell = cellAt.get(key(r, c))
    // Rich text: the shared string's own runs, each over the cell's font.
    if (cell?.type === 's' && !values.hasFormula(index, r, c)) {
      const runs = doc.sharedStringRuns?.[Number(cell.rawValue)]
      if (runs?.length) return { runs: runs.map((run) => richRun(run, scale)), style, numeric: false, kind: 'str' }
    }
    const { type, raw } = rawOf(value)
    const numFmtId = xf?.numFmtId ?? 0
    let text: string
    if (type === 'n') {
      text = applyNumberFormat(raw, numFmtId, styles.numFmts, doc.date1904)
      const tint = numberFormatColorHex(raw, numFmtId, styles.numFmts)
      if (tint) style.color = tint
    } else if (type === 'b') {
      text = raw === '1' ? 'TRUE' : 'FALSE'
    } else {
      text = raw
    }
    return { runs: text === '' ? [] : [text], style, numeric: type === 'n', kind: value.t }
  }

  const borderAt = (row: number, column: number) => {
    const at = xfAt(row, column)
    return at ? styles.borders[at.borderId] : undefined
  }
  const fillAt = (row: number, column: number) => {
    const at = xfAt(row, column)
    return at && hasFill(styles, at) ? styles.fills[at.fillId]!.fgColorHex : undefined
  }
  const widthAt = (column: number) => widthOf(column) * scale

  const tableRows: TableRow[] = rows.map((r) => {
    // The row's cells by visible column, before overflow is settled.
    const slots = columns.map((c) => {
      if (covered.has(key(r, c))) return undefined
      const xf = xfAt(r, c)
      const alignment = xf?.applyAlignment === false ? undefined : xf?.alignment
      return { c, xf, alignment, merge: merges.get(key(r, c)), ...textOf(r, c, xf) }
    })
    // Excel lets unwrapped text run on into empty cells beside it (right for text set left, left for text set right,
    // both ways for centred). The cells it runs into are joined to it here, so the text prints whole on one line.
    const span: number[] = slots.map((slot) => (slot ? 1 : 0))
    const absorbed = new Set<number>()
    slots.forEach((slot, i) => {
      if (!slot || slot.merge || slot.numeric || !slot.runs.length || slot.alignment?.wrapText || absorbed.has(i)) return
      const horizontal = slot.alignment?.horizontal
      if (horizontal === 'fill' || horizontal === 'justify' || horizontal === 'distributed') return
      const needed = textWidth(slot.runs, slot.style) + 4
      let have = widthAt(slot.c)
      if (needed <= have) return
      const free = (j: number) => {
        const other = slots[j]
        if (!other || other.merge || other.runs.length || absorbed.has(j)) return false
        if (fillAt(r, other.c) !== fillAt(r, slot.c)) return false
        const b = borderAt(r, other.c)
        return !b || [b.left, b.right].every((e) => edgeOf(e) === undefined)
      }
      const direction = horizontal === 'right' ? -1 : horizontal === 'center' || horizontal === 'centerContinuous' ? 0 : 1
      let lo = i
      let hi = i
      while (needed > have) {
        const right = direction >= 0 && hi + 1 < slots.length && free(hi + 1)
        const left = direction <= 0 && lo - 1 >= 0 && free(lo - 1)
        if (right && (!left || direction > 0 || hi - i <= i - lo)) {
          hi += 1
          have += widthAt(slots[hi]!.c)
        } else if (left) {
          lo -= 1
          have += widthAt(slots[lo]!.c)
        } else break
      }
      if (lo === i && hi === i) return
      // The text moves to the first cell of the run; the cells it covers are dropped.
      for (let j = lo; j <= hi; j += 1) if (j !== lo) absorbed.add(j)
      if (lo !== i) {
        slots[lo] = { ...slot, c: slots[lo]!.c }
        absorbed.add(i)
      }
      span[lo] = hi - lo + 1
    })

    const cells: TableCell[] = []
    slots.forEach((slot, i) => {
      if (!slot || absorbed.has(i)) return
      const { c, xf, alignment, merge, runs, style, numeric, kind } = slot
      const cell: TableCell = { content: runs.length ? runs : '', text: style }
      const colSpan = merge?.colSpan ?? span[i]!
      if (colSpan > 1) cell.colSpan = colSpan
      if (merge && merge.rowSpan > 1) cell.rowSpan = merge.rowSpan
      const horizontal = alignment?.horizontal ? HORIZONTAL[alignment.horizontal] : undefined
      // Excel's General alignment: numbers right, logical values and errors centred, text left.
      const align = horizontal ?? (numeric ? 'right' : kind === 'bool' || kind === 'err' ? 'center' : undefined)
      if (align && align !== 'left') cell.align = align
      cell.verticalAlign = alignment?.vertical ? (VERTICAL[alignment.vertical] ?? 'bottom') : 'bottom'
      if (alignment?.indent) cell.padding = { left: 2 + alignment.indent * 9 * scale }
      const fill = fillAt(r, c)
      if (fill) cell.shading = fill
      // A merged (or overflowed) cell's outer edges are the edge cells' borders.
      const endRow = merge?.endRow ?? r
      const endColumn = merge?.endColumn ?? slots[i + colSpan - 1]?.c ?? c
      const borders: NonNullable<TableCell['borders']> = {}
      const top = edgeOf(borderAt(r, c)?.top) ?? gridRule
      const left = edgeOf(borderAt(r, c)?.left) ?? gridRule
      const bottom = edgeOf(borderAt(endRow, c)?.bottom) ?? gridRule
      const right = edgeOf(borderAt(r, endColumn)?.right) ?? gridRule
      if (top) borders.top = top
      if (left) borders.left = left
      if (bottom) borders.bottom = bottom
      if (right) borders.right = right
      if (Object.keys(borders).length) cell.borders = borders
      cells.push(cell)
    })
    // A row is never split across a page, as Excel prints it.
    const row: TableRow = { cells, cantSplit: true }
    const height = rowInfo.get(r)
    if (height?.customHeight && height.heightPt > 0) row.height = height.heightPt * scale
    return row
  })

  // Rows repeated on each page: the block's count, else the sheet's print titles, else its frozen rows.
  const titles = definedName(doc, index, '_xlnm.Print_Titles')
  const titleRows = titles ? parseTitleRowRange(titles) : undefined
  const fromSheet =
    titleRows && titleRows.startRow <= area.r0
      ? rows.filter((r) => r <= titleRows.endRow).length
      : grid.pane?.frozenRows
        ? rows.filter((r) => r < grid.pane!.frozenRows).length
        : 0
  const headerRows = Math.max(0, Math.min(tableRows.length - 1, Math.floor(block.headerRows ?? fromSheet)))
  for (let r = 0; r < headerRows; r += 1) {
    tableRows[r] = { ...tableRows[r]!, repeat: true, keepWithNext: true, cantSplit: true }
  }

  const none: Border = { style: 'none' }
  const table: TableBlock = {
    type: 'table',
    rows: tableRows,
    columnWidths: natural,
    width,
    borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
    cellPadding: { top: 1, bottom: 1, left: 2, right: 2 },
    line: 1,
  }
  if (block.align) table.align = block.align
  if (block.caption) table.caption = block.caption
  if (block.captionPosition) table.captionPosition = block.captionPosition
  if (block.note !== undefined) table.note = block.note
  if (block.spaceAfter !== undefined) table.spaceAfter = block.spaceAfter
  if (block.source !== undefined) table.source = block.source
  return table
}

/**
 * About how wide the text sets, points: an average glyph of half an em (a little more in bold). Enough to tell whether a
 * label needs the empty cells beside it.
 */
function textWidth(runs: readonly Inline[], base: TextStyle): number {
  let width = 0
  for (const run of runs) {
    const text = typeof run === 'string' ? run : run.type === 'text' ? run.text : ''
    const style = typeof run === 'string' ? base : { ...base, ...(run as TextStyle) }
    const size = style.size ?? 11
    let line = 0
    for (const part of text.split('\n')) line = Math.max(line, part.length)
    width += line * size * (style.bold ? 0.54 : 0.5)
  }
  return width
}

function richRun(run: SheetRichRun, scale: number): Inline {
  const out: Inline = { type: 'text', text: run.text }
  if (run.bold) out.bold = true
  if (run.italic) out.italic = true
  if (run.underline) out.underline = true
  if (run.strike) out.strike = true
  if (run.colorHex) out.color = run.colorHex
  if (run.sizePt) out.size = Math.round(run.sizePt * scale * 2) / 2
  if (run.fontName) out.font = run.fontName
  if (run.vertAlign === 'superscript') out.superscript = true
  if (run.vertAlign === 'subscript') out.subscript = true
  return out
}
