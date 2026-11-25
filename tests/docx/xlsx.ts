/**
 * A small made-up workbook, written by hand as SpreadsheetML: two sheets, a styled header row, merged section bands,
 * currency and percent formats, wrapped text, a frozen first row, and formulas with no cached values (as openpyxl
 * writes them), one cross-sheet and one shared.
 */
import { strToU8, zipSync } from 'fflate'

const esc = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

export type XCell =
  | { s?: number; text: string }
  | { s?: number; n: number; f?: string }
  | { s?: number; f: string; v?: number; shared?: { si: number; ref?: string } }
  | { s: number }

function cellXml(ref: string, cell: XCell): string {
  const s = cell.s ? ` s="${cell.s}"` : ''
  if ('text' in cell) return `<c r="${ref}"${s} t="inlineStr"><is><t>${esc(cell.text)}</t></is></c>`
  if ('n' in cell) return `<c r="${ref}"${s}>${cell.f ? `<f>${esc(cell.f)}</f>` : ''}<v>${cell.n}</v></c>`
  if ('f' in cell) {
    const shared = cell.shared
      ? ` t="shared" si="${cell.shared.si}"${cell.shared.ref ? ` ref="${cell.shared.ref}"` : ''}`
      : ''
    const formula = shared && !cell.shared!.ref ? `<f${shared}/>` : `<f${shared}>${esc(cell.f)}</f>`
    return `<c r="${ref}"${s}>${formula}<v>${cell.v ?? ''}</v></c>`
  }
  return `<c r="${ref}"${s}/>`
}

export interface XSheet {
  name: string
  /** Column widths in characters, from A. */
  widths?: number[]
  rows: Array<{ cells: Record<string, XCell>; height?: number }>
  merges?: string[]
  frozenRows?: number
  fitToWidth?: boolean
}

function sheetXml(sheet: XSheet): string {
  const cols = sheet.widths?.length
    ? `<cols>${sheet.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>`
    : ''
  const pane = sheet.frozenRows
    ? `<sheetViews><sheetView workbookViewId="0"><pane ySplit="${sheet.frozenRows}" topLeftCell="A${sheet.frozenRows + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>`
    : ''
  const rows = sheet.rows
    .map((row, i) => {
      const r = i + 1
      const height = row.height ? ` ht="${row.height}" customHeight="1"` : ''
      return `<row r="${r}"${height}>${Object.entries(row.cells)
        .map(([column, cell]) => cellXml(`${column}${r}`, cell))
        .join('')}</row>`
    })
    .join('')
  const merges = sheet.merges?.length
    ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>`
    : ''
  const fit = sheet.fitToWidth ? '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' : ''
  const setup = sheet.fitToWidth ? '<pageSetup orientation="landscape" paperSize="3" fitToWidth="1" fitToHeight="0"/>' : ''
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${fit}${pane}<sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${rows}</sheetData>${merges}${setup}</worksheet>`
}

/**
 * Styles by index: 0 plain Calibri 11; 1 bold on grey, thin box (header); 2 currency; 3 bold currency; 4 wrapped text,
 * top; 5 bold on peach (band); 6 percent 0.0%; 7 bold 14pt currency on yellow; 8 red bold text, right.
 */
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/><numFmt numFmtId="165" formatCode="0.0%"/></numFmts>
<fonts count="5">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="14"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/></font>
<font><sz val="20"/><name val="Calibri"/></font>
</fonts>
<fills count="5">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFE7E6E6"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="10">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="164" fontId="2" fillId="4" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="right"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
</styleSheet>`

/** The workbook as `.xlsx` bytes. */
export function makeXlsx(sheets: XSheet[]): Uint8Array {
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('\n')}
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets
      .map((sheet, i) => `<sheet name="${esc(sheet.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
      .join('')}</sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets
      .map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
      .join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`),
    'xl/styles.xml': strToU8(STYLES),
  }
  sheets.forEach((sheet, i) => {
    files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(sheetXml(sheet))
  })
  return zipSync(files)
}

/**
 * A made-up project budget: four tasks at hours × rate, a total, 5% tax, a grand total, and a second sheet that reads
 * the first. Total 9,000.00; tax 450.00; grand total 9,450.00.
 */
export function budgetWorkbook(): Uint8Array {
  const task = (r: number, name: string, text: string, hours: number, shared?: { si: number; ref?: string }) => ({
    cells: {
      A: { s: 4, text: name },
      B: { s: 4, text: text },
      C: { n: hours },
      D: { s: 2, n: 75 },
      E: shared ? { s: 3, f: `C${r}*D${r}`, shared } : { s: 3, f: `C${r}*D${r}` },
    } as Record<string, XCell>,
  })
  return makeXlsx([
    {
      name: 'Budget',
      widths: [30, 60, 12, 14, 18],
      frozenRows: 1,
      fitToWidth: true,
      merges: ['A7:E7', 'A10:D10'],
      rows: [
        {
          height: 30,
          cells: {
            A: { s: 1, text: 'TASK' },
            B: { s: 1, text: 'DESCRIPTION' },
            C: { s: 1, text: 'HOURS' },
            D: { s: 1, text: 'RATE' },
            E: { s: 1, text: 'TOTAL' },
          },
        },
        task(2, 'Planning', 'Work plan, schedule and weekly meetings with the client team.', 20, { si: 0, ref: 'E2:E5' }),
        task(3, 'Design', 'Templates, components and a tested prototype.', 40, { si: 0 }),
        task(4, 'Content', 'Rewriting of the priority pages.', 30, { si: 0 }),
        task(5, 'Handover', 'Training and the final acceptance package.', 30, { si: 0 }),
        { cells: { A: { s: 3, text: 'Total before tax' }, C: { f: 'SUM(C2:C5)' }, E: { s: 3, f: 'SUM(E2:E5)' } } },
        { cells: { A: { s: 5, text: 'TAXES' } } },
        { cells: { A: { text: 'Tax' }, B: { text: '5% of the total' }, E: { s: 2, f: 'E6*5/100' } } },
        { cells: { A: { text: 'Deposit' }, E: { s: 8, text: 'No charge' } } },
        { cells: { A: { s: 7, text: 'TOTAL INCLUDING TAX' }, E: { s: 7, f: 'E6+E8' } } },
      ],
    },
    {
      name: 'Breakdown',
      widths: [30, 14, 14],
      rows: [
        { cells: { A: { s: 9, text: 'Budget breakdown for the made-up project' } } },
        { cells: { A: { s: 1, text: 'TASK' }, B: { s: 1, text: 'FEE' }, C: { s: 1, text: 'SHARE' } } },
        { cells: { A: { text: 'Planning' }, B: { s: 2, f: 'Budget!E2' }, C: { s: 6, f: 'Budget!E2/Budget!$E$6' } } },
        { cells: { A: { text: 'Design' }, B: { s: 2, f: 'Budget!E3' }, C: { s: 6, f: 'Budget!E3/Budget!$E$6' } } },
        // A cached value the file states: printed as it is unless the block asks to evaluate.
        { cells: { A: { text: 'Cached' }, B: { s: 2, n: 1234, f: 'Budget!E4' } } },
      ],
    },
  ])
}
