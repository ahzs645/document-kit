/** Paragraph-level helpers shared by the block writers. */
import {
  AlignmentType,
  LeaderType,
  LineRuleType,
  Paragraph,
  SimpleField,
  TabStopType,
  TextRun,
  type IParagraphOptions,
  type ParagraphChild,
} from 'docx'
import { captionKind, captionPrefix, inlines, paragraphStyleId, resolveSentences, type CaptionedBlock } from '../model/index'
import type { Align, Caption, ParagraphBlock, TabStop, TextStyle } from '../model/types'
import { bookmarked } from './bookmarks'
import type { WriterContext } from './context'
import { inlineChildren } from './inline'
import { border, hex, lineMultiple, twips } from './units'

export function alignment(align: Align | undefined) {
  switch (align) {
    case 'center':
      return AlignmentType.CENTER
    case 'right':
      return AlignmentType.RIGHT
    case 'justify':
      return AlignmentType.JUSTIFIED
    case 'left':
      return AlignmentType.LEFT
    default:
      return undefined
  }
}

/** Word style ids for the model's paragraph styles. */
export const STYLE_IDS = {
  normal: undefined,
  title: 'Title',
  subtitle: 'Subtitle',
  caption: 'Caption',
  note: 'KitNote',
  small: 'KitSmall',
} as const

export const TABLE_STYLE_ID = 'KitTable'
export const TOC_TITLE_STYLE_ID = 'KitTocTitle'

function tabStop(stop: TabStop, context: WriterContext) {
  const at = stop.at === 'end' ? context.width : stop.at === 'middle' ? context.width / 2 : stop.at
  const type =
    stop.align === 'right' ? TabStopType.RIGHT : stop.align === 'center' ? TabStopType.CENTER : TabStopType.LEFT
  return { type, position: twips(at), ...(stop.leader === 'dot' ? { leader: LeaderType.DOT } : {}) }
}

/** An object's type with its possibly-undefined keys made optional and never undefined. */
export type Compact<T> = { [K in keyof T as undefined extends T[K] ? never : K]: T[K] } & {
  [K in keyof T as undefined extends T[K] ? K : never]?: Exclude<T[K], undefined>
}

/** The object without its undefined keys, so optional options are left out rather than given as undefined. */
export function compact<T extends object>(value: T): Compact<T> {
  return Object.fromEntries(Object.entries(value).filter(([, v]) => v !== undefined)) as Compact<T>
}

/** A model paragraph as a docx paragraph; null when every sentence in it was left out. */
export function writeParagraph(block: ParagraphBlock, context: WriterContext, breakBefore: boolean): Paragraph | null {
  const list = inlines(block.content)
  // A paragraph of sentences that were all left out for missing slots prints nothing at all.
  if (list.some((inline) => typeof inline !== 'string' && inline.type === 'sentence') && !resolveSentences(list).length) {
    return null
  }
  const children = inlineChildren(block.content, context, compact({ style: block.text, source: block.source }))
  const style =
    paragraphStyleId(context.styles, block.style) ??
    (context.part === 'header' ? 'Header' : context.part === 'footer' ? 'Footer' : undefined)
  const spacing = block.spacing
    ? compact({
        before: block.spacing.before === undefined ? undefined : twips(block.spacing.before),
        after: block.spacing.after === undefined ? undefined : twips(block.spacing.after),
        line: block.spacing.line === undefined ? undefined : lineMultiple(block.spacing.line),
      })
    : undefined
  const indent = block.indent
    ? compact({
        left: block.indent.left === undefined ? undefined : twips(block.indent.left),
        right: block.indent.right === undefined ? undefined : twips(block.indent.right),
        firstLine: block.indent.firstLine === undefined ? undefined : twips(block.indent.firstLine),
        hanging: block.indent.hanging === undefined ? undefined : twips(block.indent.hanging),
      })
    : undefined
  const borders = block.borders
    ? compact({
        top: border(block.borders.top),
        bottom: border(block.borders.bottom),
        left: border(block.borders.left),
        right: border(block.borders.right),
      })
    : undefined
  const options: IParagraphOptions = compact({
    style,
    alignment: alignment(block.align),
    spacing,
    indent,
    keepNext: block.keepWithNext || undefined,
    keepLines: block.keepLinesTogether || undefined,
    pageBreakBefore: breakBefore || block.pageBreakBefore || undefined,
    border: borders,
    shading: block.shading ? { fill: hex(block.shading), color: 'auto', type: 'clear' as const } : undefined,
    tabStops: block.tabs?.map((stop) => tabStop(stop, context)),
    frame: context.frame,
    children: bookmarked(block.bookmark, children, context),
  })
  return new Paragraph(options)
}

/** An empty paragraph exactly `height` points tall (a spacer, or the carrier of a page break before a table). */
export function gapParagraph(height: number, breakBefore = false, frame?: WriterContext['frame']): Paragraph {
  return new Paragraph(
    compact({
      spacing: { before: 0, after: 0, line: Math.max(1, twips(height)), lineRule: LineRuleType.EXACT },
      pageBreakBefore: breakBefore || undefined,
      frame,
      children: [],
    }),
  )
}

/**
 * A caption paragraph: `Table`, the number as a SEQ field (so Word's table of figures and cross-references find it)
 * with its value already written, the separator, then the caption's own text.
 */
export function captionParagraph(
  block: CaptionedBlock,
  caption: Caption,
  context: WriterContext,
  options: { keepNext?: boolean; breakBefore?: boolean; align?: Align } = {},
): Paragraph {
  const kind = captionKind(block) ?? 'figure'
  const number = context.captions.get(block)
  const own = inlineChildren(caption.text, context, compact({ source: block.source }))
  const prefix = captionPrefix(kind, caption, number, context.theme, own.length > 0)
  const children: ParagraphChild[] = []
  if (prefix && number !== undefined) {
    const word = kind === 'table' ? context.theme.captions.table : context.theme.captions.figure
    const identifier = kind === 'table' ? 'Table' : 'Figure'
    children.push(new TextRun(`${word} `))
    children.push(new SimpleField(`SEQ ${identifier} \\* ARABIC`, String(number)))
    if (own.length) children.push(new TextRun(context.theme.captions.separator))
  }
  children.push(...own)
  return new Paragraph(
    compact({
      style: 'Caption',
      keepNext: options.keepNext || undefined,
      pageBreakBefore: options.breakBefore || undefined,
      alignment: alignment(options.align),
      frame: context.frame,
      children,
    }),
  )
}

/** A paragraph of rich text in a named style, for notes and labels the writers add. */
export function styledParagraph(
  style: keyof typeof STYLE_IDS,
  content: Parameters<typeof inlineChildren>[0],
  context: WriterContext,
  options: { align?: Align; keepNext?: boolean; text?: TextStyle; source?: string } = {},
): Paragraph {
  return new Paragraph(
    compact({
      style: STYLE_IDS[style],
      alignment: alignment(options.align),
      keepNext: options.keepNext || undefined,
      frame: context.frame,
      children: inlineChildren(content, context, compact({ style: options.text, source: options.source })),
    }),
  )
}
