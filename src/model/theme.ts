/**
 * The default theme and the merge of an application's partial theme over it. Neutral on purpose: a sans-serif body,
 * grey table header, no brand colour. An application gives its own fonts, sizes and colours; the kit knows neither.
 */
import type { PaperName, PaperSize, Margins, Theme, ThemeInput, ThemeText } from './types'

/** Paper sizes in points, portrait. A4 is 210 × 297 mm, written to the twip Word uses (11906 × 16838). */
export const PAPER: Readonly<Record<PaperName, PaperSize>> = {
  letter: { width: 612, height: 792 },
  legal: { width: 612, height: 1008 },
  a4: { width: 595.3, height: 841.9 },
  tabloid: { width: 792, height: 1224 },
}

const INCH = 72

export const DEFAULT_MARGINS: Readonly<Margins> = {
  top: INCH,
  right: INCH,
  bottom: INCH,
  left: INCH,
  header: INCH / 2,
  footer: INCH / 2,
}

export const DEFAULT_THEME: Theme = {
  body: { font: 'Calibri', size: 11, color: '1f2933', after: 6, line: 1.15 },
  title: { size: 24, bold: true, after: 12, align: 'center' },
  subtitle: { size: 14, color: '52606d', after: 12, align: 'center' },
  headings: [
    { size: 16, bold: true, before: 18, after: 6 },
    { size: 13, bold: true, before: 14, after: 4 },
    { size: 12, bold: true, before: 12, after: 4 },
    { size: 11, bold: true, italic: true, before: 10, after: 3 },
    { size: 11, italic: true, before: 8, after: 2 },
    { size: 10, italic: true, before: 6, after: 2 },
  ],
  caption: { size: 10, bold: true, before: 6, after: 6 },
  note: { size: 9, color: '52606d', after: 6 },
  small: { size: 8, color: '52606d', after: 4 },
  header: { size: 8, color: '52606d', after: 0 },
  footer: { size: 8, color: '52606d', after: 0 },
  link: { color: '1d4ed8', underline: true },
  table: {
    size: 9,
    headerFill: 'e4e7eb',
    headerText: { bold: true },
    headerAlign: 'center',
    groupFill: 'f0f2f5',
    groupText: { bold: true },
    noteText: { size: 8, color: '52606d' },
    labelFill: 'f5f7fa',
    outerBorder: { style: 'single', width: 0.5, color: '9aa5b1' },
    innerBorder: { style: 'single', width: 0.25, color: 'cbd2d9' },
    cellPadding: { top: 3, bottom: 3, left: 5, right: 5 },
    spacerHeight: 6,
    spaceAfter: 9,
    placeholderFill: 'f5f7fa',
    placeholderBorder: { style: 'dashed', width: 0.75, color: '9aa5b1' },
    placeholderText: { size: 10, color: '7b8794' },
  },
  captions: { table: 'Table', figure: 'Figure', separator: ': ' },
  blankWidth: 144,
  placeholderMark: { highlight: 'yellow' },
  page: { paper: 'letter', orientation: 'portrait', margins: { ...DEFAULT_MARGINS } },
}

const isPlain = (value: unknown): value is Record<string, unknown> =>
  typeof value === 'object' && value !== null && !Array.isArray(value) && !(value instanceof Uint8Array)

function merge<T>(base: T, over: unknown): T {
  if (over === undefined) return base
  if (!isPlain(base) || !isPlain(over)) return over as T
  const out: Record<string, unknown> = { ...base }
  for (const [key, value] of Object.entries(over)) {
    if (value === undefined) continue
    out[key] = merge(out[key], value)
  }
  return out as T
}

/**
 * The theme a writer uses: the application's parts over the default. Heading levels merge one by one, so a theme can
 * restyle level 1 alone.
 */
export function resolveTheme(input?: ThemeInput): Theme {
  if (!input) return DEFAULT_THEME
  const { headings, ...rest } = input
  const theme = merge(DEFAULT_THEME, rest)
  if (!headings) return theme
  const levels: ThemeText[] = DEFAULT_THEME.headings.map((level, index) => merge(level, headings[index]))
  return { ...theme, headings: levels }
}
