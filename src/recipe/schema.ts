/**
 * Where each key of the block model sits, as the recipe system reads a block template: rich text, a nullable value,
 * a list of rows, plain text, a scalar. One table serves the compile's check (a key the model does not have is an
 * error naming it) and the render (how a tag's value or a missing value is printed at that place). Keep it in step
 * with `src/model/types.ts`.
 */

/**
 * A place in a block template:
 * - `blocks`, `rows`, `cells`, `fieldRows`, `listItems`, `gridItems`, `inlines`: lists, whose items may be control nodes;
 * - `block`, `row`, `cell`, `fieldRow`, `listItem`, `gridItem`, `inline`, `caption`: one item;
 * - `rich`: rich text (a missing value prints a blank rule); `richNull`: the same, but a missing whole value is null
 *   (a field's value, a cell, a signature's name), which the model prints blank and lists; `richOpt`: optional rich
 *   text (a table's note, a fields block's title), left out when a lone tag has no value;
 * - `text`: a plain string (`alt`, `link`, `what`); `scalar`: a number, a flag or a word (`level`, `align`);
 * - `image`: a picture from the data (null when missing); `slots`: a sentence's values (null when missing);
 * - `any`: free JSON (`data` of a custom block, `page`, `theme`), raw values kept;
 * - `literal`: a string taken as written, with no tags (a sentence's template).
 */
export type Pos =
  | 'blocks'
  | 'block'
  | 'rows'
  | 'row'
  | 'cells'
  | 'cell'
  | 'fieldRows'
  | 'fieldRow'
  | 'listItems'
  | 'listItem'
  | 'gridItems'
  | 'gridItem'
  | 'inlines'
  | 'inline'
  | 'caption'
  | 'rich'
  | 'richNull'
  | 'richOpt'
  | 'text'
  | 'scalar'
  | 'image'
  | 'slots'
  | 'any'
  | 'literal'

/** The item place of a list place. */
export const ITEM_OF: Partial<Record<Pos, Pos>> = {
  blocks: 'block',
  rows: 'row',
  cells: 'cell',
  fieldRows: 'fieldRow',
  listItems: 'listItem',
  gridItems: 'gridItem',
  inlines: 'inline',
  rich: 'inline',
  richNull: 'inline',
  richOpt: 'inline',
}

const BASE: Record<string, Pos> = { id: 'text', source: 'text', pending: 'any', pageBreakBefore: 'scalar', textRules: 'any' }

const STYLE: Record<string, Pos> = {
  bold: 'scalar',
  italic: 'scalar',
  underline: 'scalar',
  strike: 'scalar',
  superscript: 'scalar',
  subscript: 'scalar',
  smallCaps: 'scalar',
  allCaps: 'scalar',
  color: 'scalar',
  size: 'scalar',
  font: 'scalar',
  highlight: 'scalar',
}

/** Each block type's keys (besides `type`). */
export const BLOCK_KEYS: Record<string, Record<string, Pos>> = {
  heading: { ...BASE, level: 'scalar', content: 'rich', align: 'scalar' },
  paragraph: {
    ...BASE,
    content: 'rich',
    style: 'scalar',
    align: 'scalar',
    spacing: 'any',
    indent: 'any',
    keepWithNext: 'scalar',
    keepLinesTogether: 'scalar',
    text: 'any',
    tabs: 'any',
    borders: 'any',
    shading: 'scalar',
  },
  list: { ...BASE, ordered: 'scalar', items: 'listItems', start: 'scalar' },
  fields: {
    ...BASE,
    title: 'richOpt',
    rows: 'fieldRows',
    layout: 'scalar',
    labelWidth: 'scalar',
    separator: 'text',
    size: 'scalar',
    rowSpacing: 'scalar',
    sameLineGap: 'scalar',
  },
  table: {
    ...BASE,
    rows: 'rows',
    columnWidths: 'any',
    width: 'scalar',
    align: 'scalar',
    caption: 'caption',
    captionPosition: 'scalar',
    note: 'richOpt',
    size: 'scalar',
    borders: 'any',
    cellPadding: 'any',
    line: 'scalar',
    cantSplit: 'scalar',
    keepTogether: 'scalar',
    repeatHeader: 'scalar',
    columnAlign: 'any',
    autoFit: 'scalar',
  },
  figure: {
    ...BASE,
    image: 'image',
    fit: 'any',
    caption: 'caption',
    placeholder: 'richOpt',
    align: 'scalar',
    alt: 'text',
    note: 'richOpt',
    captionPosition: 'scalar',
    crop: 'any',
    border: 'any',
  },
  figureGrid: {
    ...BASE,
    items: 'gridItems',
    columns: 'scalar',
    itemHeight: 'scalar',
    labelPosition: 'scalar',
    caption: 'caption',
    upscale: 'scalar',
  },
  signature: {
    ...BASE,
    label: 'richOpt',
    labelPosition: 'scalar',
    labelStyle: 'any',
    labelGap: 'scalar',
    image: 'image',
    imageBox: 'any',
    imageUpscale: 'scalar',
    name: 'richNull',
    title: 'richNull',
    titleLayout: 'scalar',
    titleSeparator: 'text',
    nameStyle: 'any',
    lines: 'fieldRows',
    align: 'scalar',
    ruleWidth: 'scalar',
    ruleThickness: 'scalar',
    ruleColor: 'scalar',
    ruleGap: 'scalar',
    signingSpace: 'scalar',
    spaceAbove: 'scalar',
    spaceBelow: 'scalar',
    keepTogether: 'scalar',
  },
  choices: {
    ...BASE,
    label: 'richOpt',
    options: 'any',
    chosen: 'any',
    answer: 'richNull',
    layout: 'scalar',
    boxed: 'scalar',
  },
  pinned: { ...BASE, at: 'scalar', blocks: 'blocks', spaceBelow: 'scalar' },
  pageBreak: { ...BASE },
  toc: { ...BASE, title: 'richOpt', levels: 'scalar' },
  spacer: { ...BASE, height: 'scalar' },
  custom: { ...BASE, kind: 'text', data: 'any' },
}

export const ROW_KEYS: Record<string, Pos> = {
  role: 'scalar',
  cells: 'cells',
  repeat: 'scalar',
  cantSplit: 'scalar',
  keepWithNext: 'scalar',
  height: 'scalar',
}

export const CELL_KEYS: Record<string, Pos> = {
  content: 'richNull',
  blocks: 'blocks',
  colSpan: 'scalar',
  rowSpan: 'scalar',
  align: 'scalar',
  verticalAlign: 'scalar',
  shading: 'scalar',
  text: 'any',
  borders: 'any',
  diagonal: 'scalar',
  textDirection: 'scalar',
  source: 'text',
  what: 'text',
}

export const FIELD_ROW_KEYS: Record<string, Pos> = { label: 'rich', value: 'richNull', source: 'text', what: 'text', sameLine: 'scalar' }
export const LIST_ITEM_KEYS: Record<string, Pos> = { content: 'rich', items: 'listItems', ordered: 'scalar' }
export const GRID_ITEM_KEYS: Record<string, Pos> = { image: 'image', label: 'richOpt', placeholder: 'richOpt', alt: 'text' }
/** A caption's `text` is optional rich text: a caption whose lone tag has no value is left out whole. */
export const CAPTION_KEYS: Record<string, Pos> = { text: 'richOpt', numbered: 'scalar' }

/** Each inline run type's keys (besides `type`). A text run's `text` is read as rich text in that run's style. */
export const INLINE_KEYS: Record<string, Record<string, Pos>> = {
  text: { ...STYLE, text: 'rich', link: 'text', source: 'text' },
  field: { ...STYLE, field: 'scalar' },
  break: {},
  tab: {},
  blank: { ...STYLE, what: 'text', width: 'scalar', source: 'text', show: 'scalar' },
  checkbox: { ...STYLE, checked: 'scalar', source: 'text' },
  image: { image: 'image', width: 'scalar', height: 'scalar', alt: 'text' },
  sentence: { ...STYLE, template: 'literal', values: 'slots', label: 'text', source: 'text', missing: 'scalar' },
}

/** Keys of a recipe section and the document, beyond parts. */
export const SECTION_KEYS: Record<string, Pos> = { page: 'any', header: 'blocks', footer: 'blocks', pageNumbers: 'any' }

/** Keys of an object place, or undefined for a place that is not an object of known keys. */
export function keysOf(pos: Pos, value: Record<string, unknown>): Record<string, Pos> | undefined {
  switch (pos) {
    case 'block':
      return typeof value.type === 'string' ? BLOCK_KEYS[value.type] : undefined
    case 'inline':
      return typeof value.type === 'string' ? INLINE_KEYS[value.type] : undefined
    case 'row':
      return ROW_KEYS
    case 'cell':
      return CELL_KEYS
    case 'fieldRow':
      return FIELD_ROW_KEYS
    case 'listItem':
      return LIST_ITEM_KEYS
    case 'gridItem':
      return GRID_ITEM_KEYS
    case 'caption':
      return CAPTION_KEYS
    default:
      return undefined
  }
}

/* ------------------------------------------------------------------ control nodes */

export type ControlKind = 'if' | 'for' | 'use' | 'component' | 'code'

/** The keys each control node takes. `component` and `code` appear only in a compiled recipe. */
export const CONTROL_KEYS: Record<ControlKind, readonly string[]> = {
  if: ['if', 'then', 'else'],
  for: ['for', 'do', 'empty'],
  use: ['use', 'with', 'variant'],
  component: ['component', 'variant', 'params', 'blocks'],
  code: ['code', 'variant', 'params'],
}

/** The control node an object is, by its marker key; null for any other object. */
export function controlKind(value: Record<string, unknown>): ControlKind | null {
  if ('type' in value) return null
  for (const kind of ['if', 'for', 'use', 'component', 'code'] as const) if (kind in value) return kind
  return null
}

/** Places whose items may be control nodes: every list place. */
export const isListPos = (pos: Pos): boolean => ITEM_OF[pos] !== undefined
