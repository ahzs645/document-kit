/** document-kit/docx: writes the block model as a Word file with the `docx` library. Browser-safe: no Node, no DOM. */
export { writeDocx, type WriteDocxOptions } from './write'
export { customXml, normalizePackage, obfuscateFont, type NormalizeOptions } from './package'
export { spreadsheetTable, type SpreadsheetTableOptions } from './spreadsheet'
export { sheetValues, type CellValue, type FormulaMode, type SheetValues } from './sheetValues'
export { fontEmbeddingReport, type FontEmbedding, type NamedFont } from '../model/index'
