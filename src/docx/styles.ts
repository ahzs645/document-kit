/** The theme as Word styles, and the numbering definitions the document's lists use. */
import {
  AlignmentType,
  LevelFormat,
  type ILevelsOptions,
  type INumberingOptions,
  type IStylesOptions,
} from 'docx'
import type { Theme, ThemeText } from '../model/types'
import type { ListRegistry } from './context'
import { alignment, compact, TABLE_STYLE_ID, TOC_TITLE_STYLE_ID } from './paragraphs'
import { halfPoints, hex, lineMultiple, twips } from './units'

function runOf(text: ThemeText) {
  return compact({
    font: text.font,
    size: text.size === undefined ? undefined : halfPoints(text.size),
    color: text.color ? hex(text.color) : undefined,
    bold: text.bold,
    italics: text.italic,
    allCaps: text.allCaps,
    smallCaps: text.smallCaps,
  })
}

function paragraphOf(text: ThemeText, keepNext = false) {
  const spacing = compact({
    before: text.before === undefined ? undefined : twips(text.before),
    after: text.after === undefined ? undefined : twips(text.after),
    line: text.line === undefined ? undefined : lineMultiple(text.line),
  })
  return compact({
    alignment: alignment(text.align),
    spacing: Object.keys(spacing).length ? spacing : undefined,
    keepNext: keepNext || undefined,
    keepLines: keepNext || undefined,
  })
}

const style = (id: string, name: string, text: ThemeText, extra: { keepNext?: boolean; next?: string } = {}) => ({
  id,
  name,
  basedOn: 'Normal',
  next: extra.next ?? 'Normal',
  quickFormat: true,
  run: runOf(text),
  paragraph: paragraphOf(text, extra.keepNext),
})

export function stylesFor(theme: Theme): IStylesOptions {
  const heading = (level: number) => {
    const text = theme.headings[level - 1] ?? {}
    return {
      run: runOf({ ...text, color: text.color ?? theme.body.color }),
      // The outline level is what the navigation pane and a table of contents read.
      paragraph: { ...paragraphOf(text, true), outlineLevel: level - 1 },
    }
  }
  const table = theme.table
  return {
    default: {
      document: {
        run: runOf(theme.body),
        paragraph: paragraphOf({ after: theme.body.after ?? 6, line: theme.body.line ?? 1.15, ...compact({ before: theme.body.before }) }),
      },
      title: { run: runOf({ ...theme.title, color: theme.title.color ?? theme.body.color }), paragraph: paragraphOf(theme.title) },
      heading1: heading(1),
      heading2: heading(2),
      heading3: heading(3),
      heading4: heading(4),
      heading5: heading(5),
      heading6: heading(6),
    },
    paragraphStyles: [
      style('Subtitle', 'Subtitle', theme.subtitle),
      style('Caption', 'caption', theme.caption),
      style('KitNote', 'Note', theme.note),
      style('KitSmall', 'Small', theme.small),
      style('Header', 'header', { after: 0, line: 1, ...theme.header }),
      style('Footer', 'footer', { after: 0, line: 1, ...theme.footer }),
      style(TABLE_STYLE_ID, 'Table text', { size: table.size, before: 0, after: 0, line: table.line ?? 1 }),
      // The contents' own title looks like a first-level heading but is not one, so it does not list itself.
      style(TOC_TITLE_STYLE_ID, 'Contents title', theme.headings[0] ?? {}, { keepNext: true }),
    ],
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
