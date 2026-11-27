/** The theme as Word styles, and the numbering definitions the document's lists use. */
import {
  AlignmentType,
  LevelFormat,
  type ILevelsOptions,
  type INumberingOptions,
  type IStylesOptions,
} from 'docx'
import type { StyleTable } from '../model/index'
import type { LineHeight, TextStyle, Theme, ThemeText } from '../model/types'
import type { ListRegistry } from './context'
import { runOptions } from './inline'
import { alignment, compact, TABLE_STYLE_ID, TOC_TITLE_STYLE_ID } from './paragraphs'
import { border, halfPoints, hex, lineSpacing, twips } from './units'

function runOf(text: ThemeText) {
  return compact({
    font: text.font,
    size: text.size === undefined ? undefined : halfPoints(text.size),
    color: text.color ? hex(text.color) : undefined,
    bold: text.bold,
    italics: text.italic,
    allCaps: text.allCaps,
    smallCaps: text.smallCaps,
    characterSpacing:
      text.characterSpacing === undefined ? undefined : twips(text.characterSpacing) || Number.EPSILON,
  })
}

/**
 * A style's paragraph options. `inherited` is the line height of the style it is based on, so a multiple written over
 * a height in points says its rule.
 */
function paragraphOf(text: ThemeText, keepNext = false, inherited?: LineHeight) {
  const spacing = compact({
    before: text.before === undefined ? undefined : twips(text.before),
    after: text.after === undefined ? undefined : twips(text.after),
    ...lineSpacing(text.line, inherited),
  })
  return compact({
    alignment: alignment(text.align),
    spacing: Object.keys(spacing).length ? spacing : undefined,
    keepNext: keepNext || undefined,
    keepLines: keepNext || undefined,
  })
}

const style = (
  id: string,
  name: string,
  text: ThemeText,
  extra: { keepNext?: boolean; next?: string; normalLine?: LineHeight } = {},
) => ({
  id,
  name,
  basedOn: 'Normal',
  next: extra.next ?? 'Normal',
  quickFormat: true,
  run: runOf(text),
  paragraph: paragraphOf(text, extra.keepNext, extra.normalLine),
})

/** The paragraph and character styles a document defines, as the docx library's style options. */
function definedStyles(table: StyleTable) {
  const paragraphStyles = [...table.paragraph.values()]
    .filter((style) => style.defined && style.definition)
    .map((style) => {
      const own = style.definition!
      const text = own.text ?? {}
      const borders = own.borders
        ? compact({
            top: border(own.borders.top),
            bottom: border(own.borders.bottom),
            left: border(own.borders.left),
            right: border(own.borders.right),
          })
        : undefined
      const indent = own.indent
        ? compact({
            left: own.indent.left === undefined ? undefined : twips(own.indent.left),
            right: own.indent.right === undefined ? undefined : twips(own.indent.right),
            firstLine: own.indent.firstLine === undefined ? undefined : twips(own.indent.firstLine),
            hanging: own.indent.hanging === undefined ? undefined : twips(own.indent.hanging),
          })
        : undefined
      return compact({
        id: style.id,
        name: style.name,
        basedOn: style.basedOn,
        next: style.next ?? style.id,
        quickFormat: true,
        run: runOf(text),
        paragraph: {
          ...paragraphOf(text, text.keepWithNext ?? false, style.basedOn ? table.paragraph.get(style.basedOn)?.text.line : undefined),
          // keepWithNext asks for the lines kept together too only when the style says so itself.
          ...(own.keepLinesTogether !== undefined ? { keepLines: own.keepLinesTogether } : {}),
          ...(text.keepWithNext === false ? { keepNext: false } : {}),
          ...(indent ? { indent } : {}),
          ...(borders && Object.keys(borders).length ? { border: borders } : {}),
          ...(own.shading ? { shading: { fill: hex(own.shading), color: 'auto', type: 'clear' as const } } : {}),
        },
      })
    })
  const characterStyles = [...table.character.values()].map((style) =>
    compact({
      id: style.id,
      name: style.name,
      basedOn: style.basedOn,
      quickFormat: true,
      run: runOptions(style.definition.text as TextStyle),
    }),
  )
  return { paragraphStyles, characterStyles }
}

export function stylesFor(theme: Theme, named?: StyleTable): IStylesOptions {
  const defined = named ? definedStyles(named) : { paragraphStyles: [], characterStyles: [] }
  // Every theme style is based on Normal, whose line height is the body's.
  const normalLine = theme.body.line ?? 1.15
  const heading = (level: number) => {
    const text = theme.headings[level - 1] ?? {}
    return {
      run: runOf({ ...text, color: text.color ?? theme.body.color }),
      // The outline level is what the navigation pane and a table of contents read.
      paragraph: { ...paragraphOf(text, text.keepWithNext ?? true, normalLine), outlineLevel: level - 1 },
    }
  }
  const table = theme.table
  return {
    default: {
      document: {
        run: runOf(theme.body),
        paragraph: paragraphOf({ after: theme.body.after ?? 6, line: theme.body.line ?? 1.15, ...compact({ before: theme.body.before }) }),
      },
      title: {
        run: runOf({ ...theme.title, color: theme.title.color ?? theme.body.color }),
        paragraph: paragraphOf(theme.title, false, normalLine),
      },
      footnoteText: { run: runOf({ ...theme.note, font: theme.note.font ?? theme.body.font, color: theme.note.color ?? theme.body.color }) },
      heading1: heading(1),
      heading2: heading(2),
      heading3: heading(3),
      heading4: heading(4),
      heading5: heading(5),
      heading6: heading(6),
    },
    paragraphStyles: [
      style('Subtitle', 'Subtitle', theme.subtitle, { normalLine }),
      style('Caption', 'caption', theme.caption, { normalLine }),
      style('KitNote', 'Note', theme.note, { normalLine }),
      style('KitSmall', 'Small', theme.small, { normalLine }),
      style('Header', 'header', { after: 0, line: 1, ...theme.header }, { normalLine }),
      style('Footer', 'footer', { after: 0, line: 1, ...theme.footer }, { normalLine }),
      style(TABLE_STYLE_ID, 'Table text', { size: table.size, before: 0, after: 0, line: table.line ?? 1 }, { normalLine }),
      // The contents' own title looks like a first-level heading but is not one, so it does not list itself.
      style(TOC_TITLE_STYLE_ID, 'Contents title', theme.headings[0] ?? {}, { keepNext: true, normalLine }),
      ...defined.paragraphStyles,
    ],
    ...(defined.characterStyles.length ? { characterStyles: defined.characterStyles } : {}),
  }
}

const BULLETS = ['•', '◦', '▪']
const ORDERED = [LevelFormat.DECIMAL, LevelFormat.LOWER_LETTER, LevelFormat.LOWER_ROMAN] as const

function levels(ordered: boolean, start: number): ILevelsOptions[] {
  return Array.from({ length: 9 }, (_, level) => {
    const indent = { left: twips(18 * (level + 1)), hanging: twips(14) }
    if (!ordered) {
      return {
        level,
        format: LevelFormat.BULLET,
        text: BULLETS[level % BULLETS.length] ?? '•',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent } },
      }
    }
    return {
      level,
      format: ORDERED[level % ORDERED.length] ?? LevelFormat.DECIMAL,
      text: `%${level + 1}.`,
      alignment: AlignmentType.LEFT,
      start: level === 0 ? start : 1,
      style: { paragraph: { indent } },
    }
  })
}

/** Collects the list references a document uses while it is written, then gives the numbering they need. */
export function createListRegistry(): ListRegistry & { numbering(): INumberingOptions } {
  const references = new Map<string, { ordered: boolean; start: number }>()
  let instances = 0
  return {
    instance: () => (instances += 1),
    reference(ordered, start) {
      const at = Math.max(0, Math.floor(start))
      const key = ordered ? (at === 1 ? 'kit-ordered' : `kit-ordered-${at}`) : 'kit-bullet'
      if (!references.has(key)) references.set(key, { ordered, start: at })
      return key
    },
    numbering: () => ({
      config: [...references].map(([reference, { ordered, start }]) => ({ reference, levels: levels(ordered, start) })),
    }),
  }
}
