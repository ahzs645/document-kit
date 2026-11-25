/**
 * The value of every cell of a workbook, formulas included. A formula prints the value the file cached for it; a
 * formula with none (openpyxl writes none) is evaluated by the reamkit formula engine against the other cells, which
 * are evaluated in turn, on any sheet of the workbook. Shared formulas take their group's text, shifted. A formula that
 * refers back to itself is `#REF!`.
 */
import type { WorksheetCell } from '../ream/core/spreadsheet-model'
import type { SheetDoc } from '../ream/core/ir/sheet'
import {
  BLANK,
  NO_SHIFT,
  compileFormula,
  err,
  evaluate,
  num,
  type CompiledFormula,
  type EvalContext,
  type FValue,
  type Rect,
  type Scalar,
} from '../ream/excel/formula'

export type FormulaMode = 'cached' | 'evaluate'

/** A cell's value as the sheet shows it, before its number format. */
export type CellValue = Scalar

const key = (row: number, col: number) => row * 16384 + col

interface SheetIndex {
  cells: Map<number, WorksheetCell>
  /** The populated cells, row by row, for range walks. */
  ordered: WorksheetCell[]
  /** Shared formula groups: their master cell and its compiled text. */
  shared: Map<number, { cell: WorksheetCell; formula: CompiledFormula | undefined }>
}

export interface SheetValues {
  /** The value of a cell, 0-based row and column, on the sheet at `sheet` (0-based). */
  value(sheet: number, row: number, col: number): CellValue
  /** True when the cell holds a formula. */
  hasFormula(sheet: number, row: number, col: number): boolean
}

/** Reads (and evaluates where needed) the workbook's cell values. */
export function sheetValues(doc: SheetDoc, mode: FormulaMode = 'cached'): SheetValues {
  const sheets: SheetIndex[] = doc.sheets.map((sheet) => {
    const cells = new Map<number, WorksheetCell>()
    const shared = new Map<number, { cell: WorksheetCell; formula: CompiledFormula | undefined }>()
    for (const cell of sheet.grid.cells) {
      cells.set(key(cell.row, cell.column), cell)
      if (cell.sharedFormula !== undefined && cell.formula !== undefined && !shared.has(cell.sharedFormula)) {
        shared.set(cell.sharedFormula, { cell, formula: compileFormula(cell.formula) })
      }
    }
    const ordered = [...sheet.grid.cells].sort((a, b) => a.row - b.row || a.column - b.column)
    return { cells, ordered, shared }
  })
  const byName = new Map(doc.sheets.map((sheet, index) => [sheet.name.toLowerCase(), index]))
  const compiled = new Map<string, CompiledFormula | undefined>()
  const compile = (text: string) => {
    if (!compiled.has(text)) compiled.set(text, compileFormula(text))
    return compiled.get(text)
  }
  const memo = sheets.map(() => new Map<number, Scalar>())
  const busy = sheets.map(() => new Set<number>())
  const contexts: EvalContext[] = []
  let nameDepth = 0

  const stored = (cell: WorksheetCell): Scalar => {
    switch (cell.type) {
      case 's': {
        const index = Number(cell.rawValue)
        const text = Number.isInteger(index) ? doc.sharedStrings[index] : undefined
        return text === undefined ? BLANK : { t: 'str', v: text }
      }
      case 'inlineStr':
        return { t: 'str', v: cell.inlineText ?? '' }
      case 'str':
      case 'd':
        return cell.rawValue === '' ? BLANK : { t: 'str', v: cell.rawValue }
      case 'b':
        return { t: 'bool', v: cell.rawValue === '1' || cell.rawValue.toLowerCase() === 'true' }
      case 'e':
        return err((cell.rawValue || '#N/A') as never)
      default: {
        if (cell.rawValue === '') return BLANK
        const n = Number(cell.rawValue)
        return Number.isFinite(n) ? num(n) : { t: 'str', v: cell.rawValue }
      }
    }
  }

  const formulaOf = (s: number, cell: WorksheetCell): { formula: CompiledFormula | undefined; dRow: number; dCol: number } | undefined => {
    if (cell.sharedFormula !== undefined) {
      const group = sheets[s]!.shared.get(cell.sharedFormula)
      if (group) return { formula: group.formula, dRow: cell.row - group.cell.row, dCol: cell.column - group.cell.column }
    }
    if (cell.formula !== undefined) return { formula: compile(cell.formula), dRow: 0, dCol: 0 }
    return undefined
  }

  const hasCached = (cell: WorksheetCell) => cell.rawValue !== '' || cell.type === 'inlineStr'

  const value = (s: number, row: number, col: number): Scalar => {
    const sheet = sheets[s]
    if (!sheet) return err('#REF!')
    const k = key(row, col)
    const known = memo[s]!.get(k)
    if (known) return known
    const cell = sheet.cells.get(k)
    if (!cell) return BLANK
    const formula = formulaOf(s, cell)
    let result: Scalar
    if (!formula || (mode === 'cached' && hasCached(cell))) {
      result = stored(cell)
    } else if (busy[s]!.has(k)) {
      return err('#REF!')
    } else if (!formula.formula) {
      result = hasCached(cell) ? stored(cell) : err('#NAME?')
    } else {
      busy[s]!.add(k)
      try {
        result = scalarOf(
          evaluate(formula.formula.ast, contexts[s]!, { dRow: formula.dRow, dCol: formula.dCol, curRow: row, curCol: col }),
          contexts[s]!,
        )
      } finally {
        busy[s]!.delete(k)
      }
    }
    memo[s]!.set(k, result)
    return result
  }

  const eachCellOn = (s: number, rect: Rect, visit: (row: number, col: number, value: Scalar) => void) => {
    for (const cell of sheets[s]?.ordered ?? []) {
      if (cell.row < rect.r0 || cell.row > rect.r1 || cell.column < rect.c0 || cell.column > rect.c1) continue
      const v = value(s, cell.row, cell.column)
      if (v.t !== 'blank') visit(cell.row, cell.column, v)
    }
  }

  const resolveName = (s: number) => (name: string): FValue | undefined => {
    const lower = name.toLowerCase()
    const found =
      doc.definedNames.find((d) => d.name.toLowerCase() === lower && d.localSheetId === s) ??
      doc.definedNames.find((d) => d.name.toLowerCase() === lower && d.localSheetId === undefined)
    if (!found || nameDepth > 16) return undefined
    const formula = compile(found.value.replace(/^=/, ''))
    if (!formula) return undefined
    nameDepth += 1
    try {
      return evaluate(formula.ast, contexts[found.localSheetId ?? s] ?? contexts[s]!, NO_SHIFT)
    } finally {
      nameDepth -= 1
    }
  }

  sheets.forEach((_, s) => {
    contexts.push({
      getCell: (row, col) => value(s, row, col),
      eachCell: (rect, visit) => eachCellOn(s, rect, visit),
      nowSerial: undefined,
      date1904: doc.date1904,
      sheetIndex: (name) => byName.get(name.toLowerCase()),
      getCellOn: (sheet, row, col) => value(sheet, row, col),
      eachCellOn,
      resolveName: resolveName(s),
    })
  })

  return {
    value,
    hasFormula: (s, row, col) => {
      const cell = sheets[s]?.cells.get(key(row, col))
      return cell !== undefined && (cell.formula !== undefined || cell.sharedFormula !== undefined)
    },
  }
}

/** A formula's result as one value: a one-cell reference read, a range or array taken at its first cell. */
function scalarOf(result: FValue, ctx: EvalContext): Scalar {
  if (result.t === 'arr') return result.rows[0]?.[0] ?? BLANK
  if (result.t === 'ref') {
    const { r0, c0 } = result.rect
    const v = result.sheet !== undefined ? (ctx.getCellOn?.(result.sheet, r0, c0) ?? BLANK) : ctx.getCell(r0, c0)
    // A formula pointing at an empty cell shows 0, as Excel's does.
    return v.t === 'blank' ? num(0) : v
  }
  return result
}
