/** Inline text to docx runs: styles, links, fields, breaks, tabs, blanks and pictures in a line. */
import {
  ExternalHyperlink,
  ImageRun,
  PageNumber,
  Tab,
  TextRun,
  type IRunOptions,
  type ParagraphChild,
} from 'docx'
import { resolveImage, resolveSentences } from '../model/index'
import type { InlineImage, RichText, TextStyle } from '../model/types'
import type { WriterContext } from './context'
import { halfPoints, hex, pixels } from './units'

/** A model text style as docx run options. Only what is set is written, so styles underneath show through. */
export function runOptions(style: TextStyle): IRunOptions {
  const out: Record<string, unknown> = {}
  if (style.bold !== undefined) out.bold = style.bold
  if (style.italic !== undefined) out.italics = style.italic
  if (style.underline) out.underline = {}
  if (style.strike !== undefined) out.strike = style.strike
  if (style.superscript) out.superScript = true
  if (style.subscript) out.subScript = true
  if (style.smallCaps !== undefined) out.smallCaps = style.smallCaps
  if (style.allCaps !== undefined) out.allCaps = style.allCaps
  if (style.color) out.color = hex(style.color)
  if (style.size !== undefined) out.size = halfPoints(style.size)
  if (style.font) out.font = style.font
  if (style.highlight) out.highlight = style.highlight
  return out as IRunOptions
}

/** The style keys of a run, without its content keys. */
function styleOf(value: TextStyle & Record<string, unknown>): TextStyle {
  const out: TextStyle = {}
  const keys: Array<keyof TextStyle> = [
    'bold',
    'italic',
    'underline',
    'strike',
    'superscript',
    'subscript',
    'smallCaps',
    'allCaps',
    'color',
    'size',
    'font',
    'highlight',
  ]
  for (const key of keys) {
    if (value[key] !== undefined) (out as Record<string, unknown>)[key] = value[key]
  }
  return out
}

export interface InlineBase {
  style?: TextStyle
  /** The source of the container (a cell, a field row, a block), for runs that do not name their own. */
  source?: string
}

/** A blank rule: underlined no-break spaces, about `width` points long at the text size. */
function blankRun(width: number, size: number, style: TextStyle): TextRun {
  const count = Math.max(4, Math.round(width / (size * 0.25)))
  return new TextRun({ ...runOptions(style), underline: {}, text: ' '.repeat(count) })
}

/** Plain text with its newlines as line breaks and its tab characters as tabs. */
function textRuns(text: string, options: IRunOptions): TextRun[] {
  const out: TextRun[] = []
  text.split('\n').forEach((line, index) => {
    if (index) out.push(new TextRun({ ...options, break: 1 }))
    const parts = line.split('\t')
    parts.forEach((part, partIndex) => {
      if (partIndex) out.push(new TextRun({ ...options, children: [new Tab()] }))
      if (part) out.push(new TextRun({ ...options, text: part }))
    })
  })
  return out
}

export function imageRun(image: InlineImage['image'], width: number, height: number, alt?: string): ImageRun {
  const info = resolveImage(image)
  return new ImageRun({
    type: info.format === 'jpeg' ? 'jpg' : 'png',
    data: image.bytes,
    transformation: { width: pixels(width), height: pixels(height) },
    ...(alt ? { altText: { name: alt, description: alt, title: alt } } : {}),
  })
}

/**
 * Rich text as docx paragraph children. Sentences with a missing slot are left out ({@link resolveSentences}). Runs
 * that carry a source (their own, else the container's) take the writer's style for that source over their own.
 */
export function inlineChildren(text: RichText, context: WriterContext, base: InlineBase = {}): ParagraphChild[] {
  const children: ParagraphChild[] = []
  const baseStyle = base.style ?? {}
  const sourced = (style: TextStyle, source: string | undefined): TextStyle => {
    const from = source ?? base.source
    const mark = from ? context.sourceStyles[from] : undefined
    return mark ? { ...style, ...mark } : style
  }
  for (const inline of resolveSentences(text)) {
    if (typeof inline === 'string') {
      children.push(...textRuns(inline, runOptions(sourced(baseStyle, undefined))))
      continue
    }
    switch (inline.type) {
      case 'text': {
        const style = sourced({ ...baseStyle, ...styleOf(inline as unknown as TextStyle & Record<string, unknown>) }, inline.source)
        if (inline.link) {
          const linkStyle: TextStyle = {
            color: context.theme.link.color,
            ...(context.theme.link.underline ? { underline: true } : {}),
            ...style,
          }
          children.push(new ExternalHyperlink({ link: inline.link, children: textRuns(inline.text, runOptions(linkStyle)) }))
        } else {
          children.push(...textRuns(inline.text, runOptions(style)))
        }
        break
      }
      case 'field': {
        const style = sourced({ ...baseStyle, ...styleOf(inline as unknown as TextStyle & Record<string, unknown>) }, undefined)
        const field =
          inline.field === 'page'
            ? PageNumber.CURRENT
            : inline.field === 'pages'
              ? PageNumber.TOTAL_PAGES
              : PageNumber.TOTAL_PAGES_IN_SECTION
        children.push(new TextRun({ ...runOptions(style), children: [field] }))
        break
      }
      case 'break':
        children.push(new TextRun({ ...runOptions(baseStyle), break: 1 }))
        break
      case 'tab':
        children.push(new TextRun({ ...runOptions(baseStyle), children: [new Tab()] }))
        break
      case 'blank': {
        const style = sourced({ ...baseStyle, ...styleOf(inline as unknown as TextStyle & Record<string, unknown>) }, inline.source)
        const size = style.size ?? context.theme.body.size
        children.push(blankRun(inline.width ?? context.theme.blankWidth, size, style))
        break
      }
      case 'image':
        children.push(imageRun(inline.image, inline.width, inline.height, inline.alt))
        break
      case 'sentence':
        // resolveSentences turned every printed sentence into a text run.
        break
    }
  }
  return children
}

/** A blank rule on its own, for a null field value or signature line. */
export function blankChildren(context: WriterContext, style: TextStyle, width?: number): ParagraphChild[] {
  return [blankRun(width ?? context.theme.blankWidth, style.size ?? context.theme.body.size, style)]
}
