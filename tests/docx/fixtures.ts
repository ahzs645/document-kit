/**
 * Two made-up reports shaped like the two applications' reports, built only from the kit's model. Every name,
 * number and code in them is invented.
 */
import {
  blank,
  caption,
  cell,
  custom,
  defineCustomBlocks,
  document,
  field,
  fields,
  figure,
  groupRow,
  headerRow,
  heading,
  image,
  inlineImage,
  item,
  link,
  list,
  noteRow,
  pageBreak,
  pageCount,
  pageNumber,
  paragraph,
  row,
  section,
  sentence,
  sentences,
  signature,
  sub,
  subtitle,
  sup,
  tab,
  table,
  text,
  title,
  toc,
  toComplete,
  type Block,
  type Document,
  type TableRow,
} from '../../src/index'
import { makePng } from './helpers'

/* ------------------------------------------------------------------ laboratory-style */

export function laboratoryReport(): Document {
  const letterhead = image(makePng(940, 120, [20, 90, 70]))
  const chart = image(makePng(600, 300, [230, 236, 240]))
  const rows: TableRow[] = [
    headerRow([cell('Parameter', { rowSpan: 2 }), cell('Units', { rowSpan: 2 }), cell('Samples', { colSpan: 3 })]),
    headerRow(['S-1', 'S-2', 'S-3']),
  ]
  for (const group of ['Group A', 'Group B']) {
    rows.push(groupRow([cell(group, { colSpan: 5 })]))
    for (let i = 1; i <= 30; i += 1) {
      const name = i % 10 === 0 ? ['Analyte ', `${group.slice(-1)}${i}`, sup('a')] : `Analyte ${group.slice(-1)}${i}`
      const last = i === 30 && group === 'Group B' ? cell('', { diagonal: 'up' }) : (i * 0.37).toFixed(2)
      rows.push(row([name, i % 3 ? 'mg/L' : ['mg CaCO', sub('3'), '/L'], (i * 0.11).toFixed(2), (i * 0.23).toFixed(2), last]))
    }
  }
  rows.push(noteRow([sup('a'), ' Made-up footnote: reported on a dry basis.']))
  return document(
    [
      section(
        [
          title('Certificate of Analysis'),
          fields([
            field('Attn', 'A. Client'),
            field('Address', '1 Example Road, Sampletown'),
            field('E-mail', link('client@example.com', 'mailto:client@example.com')),
            field('Phone', null),
            field('Re', 'Water samples, made-up project'),
            field('Date', '12 March 2026'),
            field('Reference File', 'LAB-0001'),
          ]),
          signature({ label: 'Report authorized by:', name: 'B. Analyst, Laboratory Lead', align: 'center' }),
          pageBreak(),
          heading(2, 'Sample and analysis notes'),
          paragraph(['Alkalinity is reported as mg CaCO', sub('3'), '/L. Three samples were received in good order.']),
          table(rows, {
            columnWidths: [3, 1.6, 1, 1, 1],
            caption: caption('Results for made-up samples'),
            columnAlign: ['left', 'left', 'center', 'center', 'center'],
          }),
          paragraph([sup('a'), ' Made-up footnote shown again under the table.'], { style: 'note' }),
          table([headerRow(['Method', 'Reference']), row(['Method X', 'Made-up method 1'])], {
            caption: caption('Methods'),
          }),
          figure(chart, { fit: { width: 400 }, caption: caption('Made-up chart of the results') }),
        ],
        {
          page: { paper: 'letter', margins: { top: 54, bottom: 54 } },
          header: [
            paragraph([inlineImage(letterhead, 352, 45, 'Made-up letterhead')], {
              align: 'center',
              borders: { bottom: { style: 'single', width: 0.75, color: '145a46' } },
            }),
            paragraph('LAB-0001 R0', { align: 'right' }),
          ],
          footer: [
            paragraph(['LAB-0001', tab(), 'Page ', pageNumber(), ' of ', pageCount()], {
              tabs: [{ at: 'end', align: 'right' }],
            }),
          ],
        },
      ),
    ],
    {
      metadata: { title: 'Certificate of Analysis', author: 'Made-up Lab', created: '2026-03-12T09:00:00Z' },
      theme: {
        body: { font: 'Times New Roman', size: 10.5 },
        title: { size: 24, color: '145a46' },
        headings: [{}, { color: '145a46', size: 12 }],
        table: { headerFill: 'd5e8dc', groupFill: 'e8f1eb' },
      },
    },
  )
}

/* ------------------------------------------------------------------ review-style */

export interface MatrixData {
  title: string
  rowHeader: string
  rows: string[]
  columns: string[]
  cells: string[][]
  mark: { row: number; column: number } | null
}

export interface ChoicesData {
  label: string
  options: string[]
  chosen: string | null
}

/** How the review application turns its own blocks into the kit's. */
export const reviewRegistry = defineCustomBlocks<{ matrix: MatrixData; choices: ChoicesData }>({
  matrix: (data) => [
    table(
      [
        headerRow([data.rowHeader, ...data.columns]),
        ...data.rows.map((label, r) =>
          row([
            cell(label, { text: { bold: true } }),
            ...data.columns.map((_, c) => {
              const marked = data.mark?.row === r && data.mark.column === c
              const value = data.cells[r]?.[c] ?? ''
              return marked ? cell(value, { shading: 'f4c7c3', text: { bold: true }, align: 'center' }) : cell(value, { align: 'center' })
            }),
          ]),
        ),
      ],
      {
        caption: caption(data.title),
        ...(data.mark ? {} : { pending: [{ what: `${data.title}: the cell the assessment reads` }] }),
      },
    ),
  ],
  choices: (data) => [
    paragraph([
      text(`${data.label} `, { bold: true }),
      data.options.map((option) => `${option === data.chosen ? '☒' : '☐'} ${option}`).join('   '),
    ], data.chosen === null ? { pending: [{ what: data.label }] } : {}),
  ],
})

export function reviewReport(): Document {
  const sheet = image(makePng(612, 396, [210, 225, 215]))
  const wideColumns = ['Unit', ...Array.from({ length: 9 }, (_, i) => `C${i + 1}`)]
  const wide = table(
    [
      headerRow(wideColumns),
      ...Array.from({ length: 12 }, (_, r) => row([`Unit ${r + 1}`, ...wideColumns.slice(1).map((_, c) => String((r + 1) * (c + 2)))])),
    ],
    { caption: caption('Made-up inventory across nine measures'), size: 8 },
  )
  const cover = section(
    [
      title('Made-up Visual Review'),
      subtitle('Block 0 from an invented viewpoint'),
      fields(
        [field('Prepared for', 'Example Forest Company'), field('Prepared by', null, { source: 'reviewer' }), field('Date', '2 April 2026')],
        { layout: 'table' },
      ),
    ],
    { page: { verticalAlign: 'center' }, header: [], footer: [] },
  )
  const body: Block[] = [
    toc({ title: 'Contents' }),
    heading(1, 'Summary', { pageBreakBefore: true }),
    paragraph(
      sentences(
        sentence('The proposal alters {percent}% of the landform seen from the viewpoint.', { percent: 4.2 }, { source: 'run' }),
        sentence('It is {visibility} from the viewpoint.', { visibility: null }, { label: 'Visibility statement', source: 'reviewer' }),
        'The figures below are made up.',
      ),
    ),
    heading(1, 'Assessment'),
    fields([field('Viewing distance', '2.1 km', { source: 'run' }), field('Visibility term', null, { source: 'reviewer' })]),
    custom<MatrixData>('matrix', {
      title: 'Made-up rating matrix',
      rowHeader: 'Size',
      rows: ['Small', 'Medium', 'Large'],
      columns: ['Low', 'Moderate', 'High'],
      cells: [
        ['A', 'A', 'B'],
        ['A', 'B', 'C'],
        ['B', 'C', 'C'],
      ],
      mark: { row: 1, column: 2 },
    }),
    custom<ChoicesData>('choices', { label: 'Objective met?', options: ['Yes', 'No', 'Inconclusive'], chosen: null }),
    heading(2, 'Viewpoints'),
    list(['Viewpoint 1, a made-up road', item('Viewpoint 2', ['From the made-up lake shore'])]),
  ]
  const running = {
    header: [paragraph('Made-up Visual Review', { align: 'right' })],
    footer: [paragraph(['Page ', pageNumber(), ' of ', pageCount()], { align: 'center' })],
  }
  const draft = document(
    [
      cover,
      section(body, running),
      section([heading(1, 'Inventory'), wide], { page: { orientation: 'landscape' } }),
      section([figure(sheet, { fit: 'page', caption: caption('Made-up analysis sheet') })], {
        page: { paper: 'tabloid', orientation: 'landscape', margins: { top: 36, bottom: 36, left: 36, right: 36 } },
        header: [],
        footer: [],
      }),
    ],
    {
      metadata: { title: 'Made-up Visual Review', author: 'Test Reviewer', created: '2026-04-02T12:00:00Z' },
      theme: {
        body: { font: 'Arial', size: 10 },
        headings: [{ size: 15, color: '1f3a5f' }],
        table: { headerFill: 'd9dee4' },
        captions: { table: 'Table', figure: 'Figure', separator: '. ' },
      },
    },
  )
  // The to-complete list closes the report, built from the document itself.
  const items = toComplete(draft, { registry: reviewRegistry })
  const closing = section(
    [
      heading(1, 'To complete before signing'),
      list(items.map((entry) => [entry.heading ? `${entry.heading}: ` : '', entry.what, entry.missing ? ` (needs ${entry.missing.join(', ')})` : ''].join(''))),
      paragraph(['Reviewer initials: ', blank('reviewer initials')]),
    ],
    running,
  )
  return { ...draft, sections: [...draft.sections, closing] }
}
