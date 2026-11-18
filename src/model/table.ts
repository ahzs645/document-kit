/**
 * A table's rows placed on its grid: where each written cell starts, how far it spans, and which grid places are
 * covered by a cell from a row above. Cells are listed as in HTML: a cell covered by another's span is not written.
 */
import type { CellInput, RichText, TableBlock, TableCell } from './types'

export interface PlacedCell {
  cell: TableCell
  row: number
  column: number
  colSpan: number
  rowSpan: number
}

export interface TableGrid {
  columns: number
  /**
   * For each row, its cells in grid order: a written cell, or the continuation of a cell from a row above (once per
   * covering cell, at its first column; it covers `by.colSpan` columns).
   */
  rows: Array<Array<{ kind: 'cell'; placed: PlacedCell } | { kind: 'covered'; by: PlacedCell }>>
  /** Every written cell. */
  cells: PlacedCell[]
}

/** A cell object has no `type`; a run (`{ type: 'text', … }`) is the cell's content. */
const isCellObject = (input: CellInput): input is TableCell =>
  input !== null && typeof input === 'object' && !Array.isArray(input) && !('type' in input)

/** A cell input as a {@link TableCell}. A string, a list of runs or null is the cell's content. */
export function toCell(input: CellInput): TableCell {
  return isCellObject(input) ? input : { content: input as RichText | null }
}

/**
 * The number of grid columns: the column widths' count when given, else the widest row counting spans and the cells
 * that rows above carry down into it.
 */
export function tableColumns(table: Pick<TableBlock, 'rows' | 'columnWidths'>): number {
  if (table.columnWidths?.length) return table.columnWidths.length
  return placeRows(table.rows, Number.POSITIVE_INFINITY).columns
}

/**
 * Places the rows on the grid. A `note` or `spacer` row with a single unspanned cell spans the whole width. Throws
 * when a row holds more than the columns, or a span runs past the last row.
 */
export function tableGrid(table: Pick<TableBlock, 'rows' | 'columnWidths'>): TableGrid {
  const columns = tableColumns(table)
  const placed = placeRows(table.rows, columns)
  return { ...placed, columns }
}

function placeRows(rows: TableBlock['rows'], limit: number): TableGrid {
  // covered[r] maps a column to the cell (from a row above) that covers it.
  const covered: Array<Map<number, PlacedCell>> = rows.map(() => new Map())
  const out: TableGrid['rows'] = []
  const cells: PlacedCell[] = []
  let widest = 0
  rows.forEach((row, r) => {
    const line: TableGrid['rows'][number] = []
    const spanAll = (row.role === 'note' || row.role === 'spacer') && row.cells.length <= 1 && Number.isFinite(limit)
    let column = 0
    const advanceOverCovered = () => {
      while (covered[r]?.has(column)) {
        const by = covered[r]?.get(column)
        // One entry for the covering cell, at its first column; it spans the rest.
        if (by && by.column === column) line.push({ kind: 'covered', by })
        column += 1
      }
    }
    const inputs = row.cells.length ? row.cells : row.role === 'spacer' || row.role === 'note' ? [''] : []
    for (const input of inputs) {
      advanceOverCovered()
      const cell = toCell(input)
      const colSpan = spanAll && !cell.colSpan ? Math.max(1, limit - column) : Math.max(1, Math.floor(cell.colSpan ?? 1))
      const rowSpan = Math.max(1, Math.floor(cell.rowSpan ?? 1))
      if (column + colSpan > limit) {
        throw new Error(
          `document-kit: table row ${r + 1} needs ${column + colSpan} columns but the table has ${limit}`,
        )
      }
      if (r + rowSpan > rows.length) {
        throw new Error(`document-kit: a cell in table row ${r + 1} spans ${rowSpan} rows, past the table's last row`)
      }
      const placedCell: PlacedCell = { cell, row: r, column, colSpan, rowSpan }
      cells.push(placedCell)
      line.push({ kind: 'cell', placed: placedCell })
      for (let below = 1; below < rowSpan; below += 1) {
        for (let c = column; c < column + colSpan; c += 1) {
          const map = covered[r + below]
          if (map?.has(c)) throw new Error(`document-kit: two cells overlap in table row ${r + below + 1}`)
          map?.set(c, placedCell)
        }
      }
      column += colSpan
    }
    advanceOverCovered()
    widest = Math.max(widest, column)
    out.push(line)
  })
  return { columns: Number.isFinite(limit) ? limit : widest, rows: out, cells }
}
