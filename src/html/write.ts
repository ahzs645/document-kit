/**
 * The block model as one self-contained, print-styled HTML document: the same document `writeDocx` writes, read the
 * same way (custom blocks expanded, text rules applied, captions numbered, the theme's and the document's named
 * styles), so the browser's print view looks like the Word file and the PDF.
 *
 * Pages are approximated with CSS: each section is a run of sheets of its paper size, a new sheet at every page break,
 * with the margins as padding, the header and footer in the margins, frames and floating pictures placed on the
 * sheet, pinned blocks at its foot. The browser does the line breaking, so a sheet that overflows grows (on screen)
 * or runs on to another sheet of paper (in print) without its header; page numbers count sheets.
 */
import { spreadsheetTable } from '../docx/spreadsheet'
import {
  applyTextRules,
  captionKind,
  captionPrefix,
  characterStyleId,
  CHECKED_BOX,
  EMPTY_BOX,
  expandDocument,
  fitImage,
  isRichText,
  numberCaptions,
  pageBox,
  paragraphStyleId,
  plainText,
  resolveImage,
  resolveSentences,
  resolveTheme,
  styleTable,
  tableGrid,
  type CaptionedBlock,
  type CustomRegistry,
  type NamedFont,
  type PageBox,
  type StyleTable,
} from '../model/index'
import { choiceBlocks } from '../model/choices'
import type {
  Align,
  Block,
  Border,
  Caption,
  CustomPropertyValue,
  Document,
  FieldRow,
  FieldsBlock,
  FigureBlock,
  FigureGridBlock,
  FloatingImageBlock,
  FrameBlock,
  HeadingBlock,
  ImageData,
  Inline,
  ListBlock,
  ListItem,
  ParagraphBlock,
  PictureCrop,
  RichText,
  Section,
  SignatureBlock,
  SpreadsheetBlock,
  TabStop,
  TableBlock,
  TableCell,
  TextStyle,
  Theme,
  ThemeText,
  TocBlock,
} from '../model/types'
import {
  BASE_CSS,
  base64,
  borderCss,
  color,
  escapeHtml,
  fontFaceCss,
  fontStack,
  LINE,
  paragraphCss,
  pt,
  styleAttr,
  styleCss,
  textCss,
} from './css'

export interface WriteHtmlOptions {
  /** Expanders for the document's custom blocks, as `writeDocx` takes them. */
  registry?: CustomRegistry
  /** How text carrying a source tag is marked, as `writeDocx` takes it. */
  sourceStyles?: Readonly<Record<string, TextStyle>>
  /** A null table cell prints empty (default) or as a blank rule. */
  nullCells?: 'empty' | 'rule'
  /**
   * Fonts to carry in the page as `@font-face` rules (base64), by family name: the faces `writeDocx` embeds and the
   * layout draws with. Each is checked first (`fontEmbeddingReport(fonts, 'html')`): a face whose licence flags
   * forbid embedding is left out, and the page falls back to the face installed, its stand-in, then the generic kind.
   */
  fonts?: readonly NamedFont[]
  /** The page's language (`<html lang>`); `en` by default. */
  lang?: string
  /** The colour behind the sheets on screen; a light grey by default. Not printed. */
  desk?: string
}

/* ------------------------------------------------------------------ the document's sheets */

interface SheetPlan {
  section: number
  /** Index of the sheet within its section. */
  index: number
  blocks: Block[]
  /** The page number printed on it, in its section's numbering. */
  number: number
  label: string
}

const PAGE_FORMATS = ['decimal', 'lowerRoman', 'upperRoman', 'lowerLetter', 'upperLetter'] as const

function formatPage(n: number, format: (typeof PAGE_FORMATS)[number] | undefined): string {
  switch (format) {
    case 'lowerRoman':
    case 'upperRoman': {
      const table: Array<[number, string]> = [
        [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'],
        [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
      ]
      let rest = Math.max(1, n)
      let out = ''
      for (const [value, digits] of table) while (rest >= value) (out += digits), (rest -= value)
      return format === 'upperRoman' ? out.toUpperCase() : out
    }
    case 'lowerLetter':
    case 'upperLetter': {
      // Word's letters: a … z, then aa … zz.
      const letter = String.fromCharCode(97 + ((Math.max(1, n) - 1) % 26)).repeat(Math.floor((Math.max(1, n) - 1) / 26) + 1)
      return format === 'upperLetter' ? letter.toUpperCase() : letter
    }
    default:
      return String(n)
  }
}

/** Each section's body cut into sheets at its page breaks, numbered as Word numbers its pages. */
function planSheets(sections: readonly Section[]): SheetPlan[] {
  const out: SheetPlan[] = []
  let number = 0
  sections.forEach((section, s) => {
    if (section.pageNumbers?.start !== undefined) number = section.pageNumbers.start - 1
    const format = section.pageNumbers?.format
    const sheets: Block[][] = [[]]
    let breakNext = false
    for (const block of section.body) {
      if (block.type === 'pageBreak') {
        breakNext = true
        continue
      }
      const current = sheets[sheets.length - 1]!
      if ((breakNext || block.pageBreakBefore) && current.length) sheets.push([])
      breakNext = false
      sheets[sheets.length - 1]!.push(block)
    }
    // A break with nothing after it still starts a new page.
    if (breakNext) sheets.push([])
    sheets.forEach((blocks, index) => {
      number += 1
      out.push({ section: s, index, blocks, number, label: formatPage(number, format) })
    })
  })
  return out
}

/* ------------------------------------------------------------------ writer state */

interface Sheet {
  plan: SheetPlan
  /** Frames and floating pictures, placed on the sheet. */
  placed: string[]
  pinned: string[]
  notes: string[]
}

interface State {
  theme: Theme
  styles: StyleTable
  captions: ReadonlyMap<CaptionedBlock, number>
  sourceStyles: Readonly<Record<string, TextStyle>>
  nullCells: 'empty' | 'rule'
  footnotes: Document['footnotes']
  noteNumbers: Map<string, number>
  headings: Array<{ level: number; id: string; text: string; page: string }>
  ids: Set<string>
  tocs: TocBlock[]
  pages: number
  sectionPages: Map<number, number>
}

interface Ctx {
  state: State
  sheet: Sheet
  page: PageBox
  /** Width the container gives its content, points. */
  width: number
  part: 'body' | 'header' | 'footer'
  inCell?: boolean
  inFrame?: boolean
}

/* ------------------------------------------------------------------ entry */

/**
 * Writes the document as one self-contained HTML page (pictures and the given fonts inline as `data:` URLs), styled
 * for print: `@page` sizes per section, a sheet a page, the theme's and the document's named styles as CSS classes.
 * Custom blocks are expanded through `options.registry`, text rules applied and captions numbered, as `writeDocx`
 * does. Synchronous; runs in a browser or in Node.
 */
export function writeHtml(source: Document, options: WriteHtmlOptions = {}): string {
  if (!source.sections.length) throw new Error('document-kit: a document needs at least one section')
  const document = applyTextRules(expandDocument(source, options.registry)).document
  const theme = resolveTheme(document.theme)
  const styles = styleTable(theme, document.styles)
  const plans = planSheets(document.sections)
  const sectionPages = new Map<number, number>()
  for (const plan of plans) sectionPages.set(plan.section, (sectionPages.get(plan.section) ?? 0) + 1)
  const state: State = {
    theme,
    styles,
    captions: numberCaptions(document),
    sourceStyles: options.sourceStyles ?? {},
    nullCells: options.nullCells ?? 'empty',
    footnotes: document.footnotes,
    noteNumbers: new Map(),
    headings: [],
    ids: new Set(),
    tocs: [],
    pages: plans.length,
    sectionPages,
  }

  const sectionCss: string[] = []
  const boxes = document.sections.map((section) => pageBox(section.page, theme))
  boxes.forEach((box, s) => {
    const m = box.margins
    const align = document.sections[s]!.page?.verticalAlign
    sectionCss.push(
      `@page dk-s${s}{size:${pt(box.width)} ${pt(box.height)};margin:0}`,
      `.dk-s${s}{page:dk-s${s};width:${pt(box.width)};min-height:${pt(box.height)};padding:${pt(m.top)} ${pt(m.right)} ${pt(m.bottom)} ${pt(m.left)}}`,
      `.dk-s${s}>.dk-header{top:${pt(m.header)};left:${pt(m.left)};width:${pt(box.contentWidth)}}`,
      `.dk-s${s}>.dk-footer{bottom:${pt(m.footer)};left:${pt(m.left)};width:${pt(box.contentWidth)}}`,
      ...(align === 'center' || align === 'bottom'
        ? [`.dk-s${s}>.dk-body{justify-content:${align === 'center' ? 'center' : 'flex-end'}}`]
        : []),
    )
  })

  // Headers carry on from the section before when a section gives none (Word's own rule).
  let header: readonly Block[] = []
  let footer: readonly Block[] = []
  const sheets: string[] = []
  let lastSection = -1
  for (const plan of plans) {
    const section = document.sections[plan.section]!
    if (plan.section !== lastSection) {
      if (section.header) header = section.header
      if (section.footer) footer = section.footer
      lastSection = plan.section
    }
    const page = boxes[plan.section]!
    const sheet: Sheet = { plan, placed: [], pinned: [], notes: [] }
    const ctx: Ctx = { state, sheet, page, width: page.contentWidth, part: 'body' }
    const first = plan.index === 0 && section.firstPage
    const head = first ? (section.firstPage!.header ?? []) : header
    const foot = first ? (section.firstPage!.footer ?? []) : footer
    const headerHtml = blocksHtml(head, { ...ctx, part: 'header' })
    const footerHtml = blocksHtml(foot, { ...ctx, part: 'footer' })
    const body = blocksHtml(plan.blocks, ctx)
    const notes = sheet.notes.length ? `<div class="dk-footnotes">${sheet.notes.join('')}</div>` : ''
    const footArea = notes || sheet.pinned.length ? `<div class="dk-foot">${notes}${sheet.pinned.join('')}</div>` : ''
    sheets.push(
      `<section class="dk-sheet dk-s${plan.section}" data-page="${escapeHtml(plan.label)}">` +
        sheet.placed.join('') +
        (headerHtml ? `<header class="dk-header">${headerHtml}</header>` : '') +
        `<div class="dk-body dk-flow">${body}${footArea}</div>` +
        (footerHtml ? `<footer class="dk-footer">${footerHtml}</footer>` : '') +
        '</section>',
    )
  }

  let html = sheets.join('\n')
  // A contents list is written once the headings it lists, and the sheets they fall on, are known.
  state.tocs.forEach((block, index) => {
    html = html.replace(`<!--dk-toc-${index}-->`, () => tocEntries(block, state))
  })

  const metadata = document.metadata ?? {}
  const body = theme.body
  const css = [
    fontFaceCss(options.fonts ?? []),
    BASE_CSS,
    `body{--dk-desk:${options.desk ? color(options.desk) : '#e5e7eb'}}`,
    `.dk-sheet{font-family:${fontStack(body.font)};font-size:${pt(body.size)};color:${color(body.color)};line-height:${Math.round((body.line ?? 1.15) * LINE * 1000) / 1000}}`,
    `.dk-sheet a{color:${color(theme.link.color)};text-decoration:${theme.link.underline ? 'underline' : 'none'}}`,
    styleCss(styles),
    `.dk-ps-KitTocTitle{${paragraphCss({ ...styles.paragraph.get('Heading1')!.text }).join(';')}}`,
    ...sectionCss,
  ]
    .filter(Boolean)
    .join('\n')
  const meta = [
    metadata.author ? `<meta name="author" content="${escapeHtml(metadata.author)}">` : '',
    metadata.subject || metadata.description
      ? `<meta name="description" content="${escapeHtml(metadata.description ?? metadata.subject ?? '')}">`
      : '',
    metadata.keywords?.length ? `<meta name="keywords" content="${escapeHtml(metadata.keywords.join(', '))}">` : '',
    ...Object.entries(metadata.custom ?? {}).map(
      ([name, value]) => `<meta name="${escapeHtml(name)}" content="${escapeHtml(propertyText(value))}">`,
    ),
  ].join('')
  return (
    `<!doctype html><html lang="${escapeHtml(options.lang ?? 'en')}"><head><meta charset="utf-8">` +
    `<meta name="viewport" content="width=device-width,initial-scale=1">${meta}` +
    // No `<` in the style sheet (CSS needs none outside strings, which escape it), so nothing in it can close it.
    `<title>${escapeHtml(metadata.title ?? '')}</title><style>\n${css.replace(/</g, '\\3c ')}\n</style></head>` +
    `<body class="dk-doc">\n${html}\n</body></html>`
  )
}

function propertyText(value: CustomPropertyValue): string {
  return value instanceof Date ? value.toISOString() : String(value)
}

/* ------------------------------------------------------------------ blocks */

function blocksHtml(blocks: readonly Block[], ctx: Ctx): string {
  let out = ''
  for (const block of blocks) {
    // Page breaks cut the body into sheets before this runs; inside a cell, a frame or a header they do nothing.
    if (block.type === 'pageBreak') continue
    out += blockHtml(block, ctx)
  }
  return out
}

function blockHtml(block: Block, ctx: Ctx): string {
  switch (block.type) {
    case 'heading':
      return headingHtml(block, ctx)
    case 'paragraph':
      return paragraphHtml(block, ctx)
    case 'list':
      return listHtml(block, ctx)
    case 'fields':
      return fieldsHtml(block, ctx)
    case 'table':
      return tableHtml(block, ctx)
    case 'spreadsheet':
      return spreadsheetHtml(block, ctx)
    case 'figure':
      return figureHtml(block, ctx)
    case 'figureGrid':
      return figureGridHtml(block, ctx)
    case 'signature':
      return signatureHtml(block, ctx)
    case 'choices':
      return blocksHtml(choiceBlocks(block, ctx.state.theme), ctx)
    case 'pinned': {
      if (ctx.inFrame || ctx.part !== 'body') return blocksHtml(block.blocks, ctx)
      const space = block.spaceBelow ? `padding-bottom:${pt(block.spaceBelow)}` : ''
      ctx.sheet.pinned.push(`<div class="dk-pinned dk-flow"${styleAttr(space)}>${blocksHtml(block.blocks, { ...ctx, inFrame: true })}</div>`)
      return ''
    }
    case 'frame':
      return frameHtml(block, ctx)
    case 'floatingImage':
      return floatingImageHtml(block, ctx)
    case 'toc':
      return tocHtml(block, ctx)
    case 'spacer':
      return block.height > 0 ? `<div class="dk-spacer"${styleAttr(`height:${pt(block.height)};flex:none`)}></div>` : ''
    case 'pageBreak':
      return ''
    case 'custom':
      throw new Error(`document-kit: custom block "${block.kind}" reached the HTML writer unexpanded`)
  }
}

/** A unique id for an element, from a bookmark name or made up. */
function idFor(state: State, wanted: string | undefined, prefix: string): string {
  if (wanted && !state.ids.has(wanted)) {
    state.ids.add(wanted)
    return wanted
  }
  let n = state.ids.size + 1
  while (state.ids.has(`${prefix}${n}`)) n += 1
  state.ids.add(`${prefix}${n}`)
  return `${prefix}${n}`
}

function headingHtml(block: HeadingBlock, ctx: Ctx): string {
  const level = Math.min(6, Math.max(1, Math.round(block.level)))
  const id = idFor(ctx.state, block.bookmark, 'dk-h')
  const content = inlineHtml(block.content, ctx, block.source ? { source: block.source } : {})
  if (ctx.part === 'body') {
    ctx.state.headings.push({ level, id, text: plainText(block.content), page: ctx.sheet.plan.label })
  }
  const style = block.align ? styleAttr(`text-align:${block.align}`) : ''
  return `<h${level} class="dk-p dk-ps-Heading${level}" id="${escapeHtml(id)}"${style}>${tabbed(content, undefined, ctx)}</h${level}>`
}

/** The Word style id a paragraph is set in: its own, else the part's. */
function paragraphStyle(block: Pick<ParagraphBlock, 'style'>, ctx: Ctx): string {
  return paragraphStyleId(ctx.state.styles, block.style) ?? (ctx.part === 'header' ? 'Header' : ctx.part === 'footer' ? 'Footer' : 'Normal')
}

function paragraphHtml(block: ParagraphBlock, ctx: Ctx): string {
  const list = resolveSentences(block.content)
  const hadSentences = (typeof block.content === 'string' ? [] : Array.isArray(block.content) ? block.content : [block.content]).some(
    (inline) => typeof inline !== 'string' && (inline as Inline & { type: string }).type === 'sentence',
  )
  // A paragraph of sentences that were all left out for missing slots prints nothing at all.
  if (hadSentences && !list.length) return ''
  const style = paragraphStyle(block, ctx)
  const css: string[] = []
  if (block.align) css.push(`text-align:${block.align}`)
  if (block.spacing?.before !== undefined) css.push(`margin-top:${pt(block.spacing.before)}`)
  if (block.spacing?.after !== undefined) css.push(`margin-bottom:${pt(block.spacing.after)}`)
  if (block.spacing?.line !== undefined) css.push(`line-height:${Math.round(block.spacing.line * LINE * 1000) / 1000}`)
  const indent = block.indent
  if (indent) {
    if (indent.left !== undefined) css.push(`margin-left:${pt(indent.left)}`)
    if (indent.right !== undefined) css.push(`margin-right:${pt(indent.right)}`)
    if (indent.hanging) css.push(`text-indent:${pt(-indent.hanging)}`)
    else if (indent.firstLine !== undefined) css.push(`text-indent:${pt(indent.firstLine)}`)
  }
  if (block.keepWithNext) css.push('break-after:avoid')
  if (block.keepLinesTogether) css.push('break-inside:avoid')
  if (block.shading) css.push(`background-color:${color(block.shading)}`)
  for (const side of ['top', 'bottom', 'left', 'right'] as const) {
    const value = block.borders?.[side]
    const border = borderCss(value)
    if (border) css.push(`border-${side}:${border}`, `padding-${side}:${pt(value?.space ?? 0)}`)
  }
  if (block.text) css.push(...textCss(block.text))
  const base: InlineBase = { ...(block.text ? { style: block.text } : {}), ...(block.source ? { source: block.source } : {}) }
  const content = tabbed(inlineHtml(block.content, ctx, base), block.tabs, ctx, indent?.left ?? 0)
  const id = block.bookmark ? ` id="${escapeHtml(idFor(ctx.state, block.bookmark, 'dk-b'))}"` : ''
  const tabbedClass = content.startsWith(TABBED) ? ' dk-tabbed' : ''
  return `<p class="dk-p dk-ps-${style}${tabbedClass}"${id}${styleAttr(css)}>${content.replace(TABBED, '')}</p>`
}

/** A paragraph of rich text in one of the theme's styles (a note, a caption), as the writers add them. */
function styled(styleId: string, content: string, css: string[] = []): string {
  return `<p class="dk-p dk-ps-${styleId}"${styleAttr(css)}>${content}</p>`
}

const TAB = '\u0001'
const TABBED = '\u0002'

/**
 * A paragraph's content with its tabs set: at the paragraph's tab stops (the text after each tab placed at its stop,
 * left, centred or right; "Page 2 of 9" at the right edge of a footer), else at Word's default stops every half inch.
 * The result starts with a marker when the paragraph needs the `dk-tabbed` class.
 */
function tabbed(content: string, stops: readonly TabStop[] | undefined, ctx: Ctx, indent = 0): string {
  if (!content.includes(TAB)) return content
  const parts = content.split(TAB)
  if (!stops?.length) return parts.join('<span class="dk-tab" style="tab-size:36pt">\t</span>')
  const ordered = stops
    .map((stop) => ({ ...stop, at: stop.at === 'end' ? ctx.width : stop.at === 'middle' ? ctx.width / 2 : stop.at }))
    .sort((a, b) => a.at - b.at)
  let out = parts[0]!
  parts.slice(1).forEach((part, i) => {
    const stop = ordered[i]
    if (!stop) {
      out += `<span class="dk-tab" style="tab-size:36pt">\t</span>${part}`
      return
    }
    const at = stop.at - indent
    const place =
      stop.align === 'right'
        ? `right:${pt(Math.max(0, ctx.width - indent - at))}`
        : stop.align === 'center'
          ? `left:${pt(at)};transform:translateX(-50%)`
          : `left:${pt(at)}`
    out += `<span class="dk-at"${styleAttr(place)}>${part}</span>`
  })
  if (!parts[0]) out = `&#8203;${out}`
  return TABBED + out
}

/* ------------------------------------------------------------------ inline text */

interface InlineBase {
  style?: TextStyle
  source?: string
}

const STYLE_KEYS: ReadonlyArray<keyof TextStyle> = [
  'bold', 'italic', 'underline', 'strike', 'superscript', 'subscript', 'smallCaps', 'allCaps',
  'color', 'size', 'font', 'highlight', 'characterSpacing', 'characterStyle',
]

function styleOf(value: object): TextStyle {
  const out: Record<string, unknown> = {}
  for (const key of STYLE_KEYS) {
    const v = (value as Record<string, unknown>)[key]
    if (v !== undefined) out[key] = v
  }
  return out as TextStyle
}

/** Text in a style: a span with the run's own look (and its character style's class), raised or lowered. */
function styledText(html: string, style: TextStyle, ctx: Ctx): string {
  characterStyleId(ctx.state.styles, style.characterStyle)
  const css = textCss(style)
  const cls = style.characterStyle ? ` class="dk-cs-${escapeHtml(style.characterStyle)}"` : ''
  let out = css.length || cls ? `<span${cls}${styleAttr(css)}>${html}</span>` : html
  if (style.superscript) out = `<sup>${out}</sup>`
  else if (style.subscript) out = `<sub>${out}</sub>`
  return out
}

/** Plain text as HTML: newlines as line breaks, tabs as the tab marker, Word's no-break hyphen kept. */
function textHtml(text: string, style: TextStyle, ctx: Ctx, topLevel: boolean): string {
  return text
    .split('\n')
    .map((line) =>
      line
        .split('\t')
        .map((part) => (part ? styledText(escapeHtml(part), style, ctx) : ''))
        .join(topLevel ? TAB : '<span class="dk-tab" style="tab-size:36pt">\t</span>'),
    )
    .join('<br>')
}

function inlineHtml(text: RichText, ctx: Ctx, base: InlineBase = {}): string {
  const state = ctx.state
  const baseStyle = base.style ?? {}
  const sourced = (style: TextStyle, source: string | undefined): TextStyle => {
    const from = source ?? base.source
    const mark = from ? state.sourceStyles[from] : undefined
    return mark ? { ...style, ...mark } : style
  }
  let out = ''
  for (const inline of resolveSentences(text)) {
    if (typeof inline === 'string') {
      out += textHtml(inline, sourced(baseStyle, undefined), ctx, true)
      continue
    }
    switch (inline.type) {
      case 'text': {
        const style = sourced({ ...baseStyle, ...styleOf(inline) }, inline.source)
        if (inline.link) {
          const link = inline.link
          const href = link.startsWith('#') ? link : /^(https?|mailto|tel|ftp):/i.test(link) ? link : undefined
          const body = textHtml(inline.text, style, ctx, false)
          out += href ? `<a href="${escapeHtml(href)}">${body}</a>` : body
        } else {
          out += textHtml(inline.text, style, ctx, true)
        }
        break
      }
      case 'footnote':
        out += footnoteHtml(inline.id, ctx)
        break
      case 'field': {
        const style = sourced({ ...baseStyle, ...styleOf(inline) }, undefined)
        const plan = ctx.sheet.plan
        const value =
          inline.field === 'page'
            ? plan.label
            : inline.field === 'pages'
              ? String(state.pages)
              : String(state.sectionPages.get(plan.section) ?? 1)
        out += styledText(`<span class="dk-field" data-field="${escapeHtml(String(inline.field))}">${escapeHtml(value)}</span>`, style, ctx)
        break
      }
      case 'break':
        out += '<br>'
        break
      case 'tab':
        out += TAB
        break
      case 'blank': {
        const style = sourced({ ...baseStyle, ...styleOf(inline) }, inline.source)
        if (inline.show === 'placeholder') {
          out += textHtml(`[${inline.what ?? ''}]`, { ...style, ...state.theme.placeholderMark }, ctx, true)
          break
        }
        out += styledText(`<span class="dk-blank"${styleAttr(`width:${pt(inline.width ?? state.theme.blankWidth)}`)}>&nbsp;</span>`, style, ctx)
        break
      }
      case 'checkbox': {
        const style = sourced({ ...baseStyle, ...styleOf(inline) }, inline.source)
        out += styledText(`<span class="dk-box">${inline.checked ? CHECKED_BOX : EMPTY_BOX}</span>`, style, ctx)
        break
      }
      case 'image':
        out += imgHtml(inline.image, inline.width, inline.height, inline.alt)
        break
      case 'sentence':
        break
    }
  }
  return out
}

function dataUrl(image: ImageData): string {
  const info = resolveImage(image)
  return `data:image/${info.format === 'jpeg' ? 'jpeg' : 'png'};base64,${base64(image.bytes)}`
}

function imgHtml(image: ImageData, width: number, height: number, alt: string | undefined, css: string[] = []): string {
  return `<img src="${dataUrl(image)}" alt="${escapeHtml(alt ?? '')}"${styleAttr([`width:${pt(width)}`, `height:${pt(height)}`, 'vertical-align:bottom', ...css])}>`
}

function footnoteHtml(id: string, ctx: Ctx): string {
  if (ctx.part !== 'body') throw new Error('document-kit: footnote citations belong in the document body')
  if (typeof id !== 'string' || !/^[a-z][a-z0-9-]{0,32}$/.test(id))
    throw new Error(`document-kit: invalid footnote ID "${id}" (use 1–33 lowercase letters, digits or hyphens, starting with a letter)`)
  const state = ctx.state
  const known = state.noteNumbers.get(id)
  if (known !== undefined) return `<sup class="dk-noteref"><a href="#dk-fn-${known}">${known}</a></sup>`
  const paragraphs = state.footnotes && Object.hasOwn(state.footnotes, id) ? state.footnotes[id] : undefined
  if (!Array.isArray(paragraphs) || !paragraphs.length || !paragraphs.every((p) => typeof p === 'string' && p.trim()))
    throw new Error(`document-kit: footnote "${id}" needs a definition with one or more nonempty paragraphs`)
  const number = state.noteNumbers.size + 1
  state.noteNumbers.set(id, number)
  ctx.sheet.notes.push(
    paragraphs
      .map((p, i) =>
        styled('KitNote', `${i === 0 ? `<sup>${number}</sup> ` : ''}${escapeHtml(p)}`, ['margin-bottom:2.25pt']),
      )
      .join('')
      .replace('<p class', `<p id="dk-fn-${number}" class`),
  )
  return `<sup class="dk-noteref"><a href="#dk-fn-${number}">${number}</a></sup>`
}

/* ------------------------------------------------------------------ lists and fields */

const BULLETS = ['disc', 'circle', 'square']
const ORDERED = ['decimal', 'lower-alpha', 'lower-roman']

function listHtml(block: ListBlock, ctx: Ctx): string {
  const visit = (items: readonly (ListItem | RichText)[], ordered: boolean, level: number, start: number): string => {
    const kind = ordered ? ORDERED[level % 3]! : BULLETS[level % 3]!
    const tag = ordered ? 'ol' : 'ul'
    const lis = items
      .map((raw) => {
        const item: ListItem = isRichText(raw) ? { content: raw } : raw
        const content = tabbed(inlineHtml(item.content, ctx, block.source ? { source: block.source } : {}), undefined, ctx)
        const nested = item.items?.length ? visit(item.items, item.ordered ?? ordered, level + 1, 1) : ''
        return `<li class="dk-ps-${paragraphStyle({}, ctx)}">${content}${nested}</li>`
      })
      .join('')
    return `<${tag} class="dk-list"${styleAttr(`list-style-type:${kind}`)}${ordered && start !== 1 ? ` start="${start}"` : ''}>${lis}</${tag}>`
  }
  return visit(block.items, block.ordered === true, 0, Math.floor(block.start ?? 1))
}

function fieldsHtml(block: FieldsBlock, ctx: Ctx): string {
  let out = ''
  if (block.title !== undefined) {
    out += `<p class="dk-p dk-ps-${paragraphStyle({}, ctx)}" style="break-after:avoid">${inlineHtml(block.title, ctx, { style: { bold: true } })}</p>`
  }
  const size = block.size
  const style: TextStyle = size === undefined ? {} : { size }
  if (block.layout === 'table') {
    const look = ctx.state.theme.table
    const share = Math.min(90, Math.max(10, block.labelWidth ?? 32))
    const table: TableBlock = {
      type: 'table',
      rows: block.rows.map((row) => ({
        cells: [
          { content: row.label, shading: look.labelFill, text: { bold: true } },
          {
            content: row.value === null ? [{ type: 'blank', ...(row.what ? { what: row.what } : {}) }] : row.value,
            ...(row.source ? { source: row.source } : {}),
          },
        ],
        cantSplit: true,
      })),
      columnWidths: [share, 100 - share],
      ...(size !== undefined ? { size } : {}),
    }
    return out + tableHtml(table, ctx)
  }
  const separator = block.separator ?? ': '
  const tabAt = block.labelWidth
  const lines: FieldRow[][] = []
  block.rows.forEach((row, index) => {
    const last = lines[lines.length - 1]
    if (row.sameLine && index > 0 && last) last.push(row)
    else lines.push([row])
  })
  const rowHtml = (row: FieldRow, joined: boolean) => {
    const label = inlineHtml(row.label, ctx, { style: { ...style, bold: true } }).split(TAB).join(' ')
    const value =
      row.value === null
        ? styledText(`<span class="dk-blank"${styleAttr(`width:${pt(ctx.state.theme.blankWidth)}`)}>&nbsp;</span>`, style, ctx)
        : inlineHtml(row.value, ctx, { style, ...(row.source ? { source: row.source } : {}) }).split(TAB).join(' ')
    const sep = escapeHtml(tabAt === undefined || joined ? separator : separator.trimEnd())
    const head =
      tabAt === undefined || joined
        ? `${label}${styledText(sep, { ...style, bold: true }, ctx)}`
        : `<span${styleAttr(`display:inline-block;min-width:${pt(tabAt)};text-indent:0`)}>${label}${styledText(sep, { ...style, bold: true }, ctx)}</span>`
    return joined ? `<span style="white-space:nowrap">${head}${value}</span>` : `${head}${value}`
  }
  const gap = block.sameLineGap ?? 24
  lines.forEach((line, index) => {
    const css: string[] = []
    if (index < lines.length - 1) css.push('break-after:avoid')
    if (tabAt !== undefined) css.push(`padding-left:${pt(tabAt)}`, `text-indent:${pt(-tabAt)}`)
    if (block.rowSpacing !== undefined) css.push(`margin-bottom:${pt(block.rowSpacing)}`)
    const content = line.map((row, at) => (at ? `<span${styleAttr(`display:inline-block;width:${pt(gap)}`)}></span>` : '') + rowHtml(row, at > 0)).join('')
    out += `<p class="dk-p dk-ps-${paragraphStyle({}, ctx)}"${styleAttr(css)}>${content}</p>`
  })
  return out
}

/* ------------------------------------------------------------------ tables */

function textStyleOf(text: ThemeText | undefined): TextStyle {
  const out: TextStyle = {}
  if (!text) return out
  if (text.font !== undefined) out.font = text.font
  if (text.size !== undefined) out.size = text.size
  if (text.color !== undefined) out.color = text.color
  if (text.bold !== undefined) out.bold = text.bold
  if (text.italic !== undefined) out.italic = text.italic
  if (text.allCaps !== undefined) out.allCaps = text.allCaps
  if (text.smallCaps !== undefined) out.smallCaps = text.smallCaps
  return out
}

function captionHtml(block: CaptionedBlock, caption: Caption, ctx: Ctx, css: string[] = []): string {
  const kind = captionKind(block) ?? 'figure'
  const number = ctx.state.captions.get(block)
  const own = inlineHtml(caption.text, ctx, block.source ? { source: block.source } : {})
  const prefix = captionPrefix(kind, caption, number, ctx.state.theme, own.length > 0)
  return styled('Caption', `${escapeHtml(prefix)}${own}`, css)
}

function tableHtml(table: TableBlock, ctx: Ctx, number?: number): string {
  const theme = ctx.state.theme
  const look = theme.table
  const grid = tableGrid(table)
  const share = Math.min(100, Math.max(1, table.width ?? 100))
  const weights = table.columnWidths ?? Array.from({ length: grid.columns }, () => 1)
  const clean = weights.map((w) => (Number.isFinite(w) && w > 0 ? w : 0))
  const total = clean.reduce((a, b) => a + b, 0) || clean.length
  const widthPt = (ctx.width * share) / 100
  const columnPt = clean.map((w) => (total ? (widthPt * (w || (total ? 0 : 1))) / total : widthPt / clean.length))
  const padding = { ...look.cellPadding, ...Object.fromEntries(Object.entries(table.cellPadding ?? {}).filter(([, v]) => v !== undefined)) }
  const borders = {
    top: look.outerBorder,
    bottom: look.outerBorder,
    left: look.outerBorder,
    right: look.outerBorder,
    insideHorizontal: look.innerBorder,
    insideVertical: look.innerBorder,
    ...Object.fromEntries(Object.entries(table.borders ?? {}).filter(([, v]) => v !== undefined)),
  } as Record<'top' | 'bottom' | 'left' | 'right' | 'insideHorizontal' | 'insideVertical', Border>
  let repeat = 0
  if (table.repeatHeader !== false) for (const row of table.rows) if (row.repeat ?? row.role === 'header') repeat += 1; else break
  const lastRow = table.rows.length - 1
  const rows = table.rows.map((row, r) => {
    const role = row.role ?? 'body'
    const roleLook =
      role === 'header'
        ? { fill: look.headerFill, text: textStyleOf(look.headerText), align: look.headerAlign as Align }
        : role === 'group'
          ? { fill: look.groupFill, text: textStyleOf(look.groupText), align: (look.groupText.align ?? 'left') as Align }
          : role === 'note'
            ? { fill: undefined, text: textStyleOf(look.noteText), align: (look.noteText.align ?? 'left') as Align }
            : { fill: undefined, text: {} as TextStyle, align: undefined }
    const cells = (grid.rows[r] ?? [])
      .flatMap((entry) => (entry.kind === 'cell' ? [entry.placed] : []))
      .map((placed) => {
        const cell: TableCell = placed.cell
        const tag = role === 'header' ? 'th' : 'td'
        const css: string[] = []
        const fill = cell.shading ?? roleLook.fill
        if (fill) css.push(`background-color:${color(fill)}`)
        const own = { ...padding, ...Object.fromEntries(Object.entries(cell.padding ?? {}).filter(([, v]) => v !== undefined)) }
        css.push(`padding:${pt(own.top)} ${pt(own.right)} ${pt(own.bottom)} ${pt(own.left)}`)
        if (cell.verticalAlign) css.push(`vertical-align:${cell.verticalAlign === 'center' ? 'middle' : cell.verticalAlign}`)
        const edge = {
          top: placed.row === 0 ? borders.top : borders.insideHorizontal,
          bottom: placed.row + placed.rowSpan - 1 === lastRow ? borders.bottom : borders.insideHorizontal,
          left: placed.column === 0 ? borders.left : borders.insideVertical,
          right: placed.column + placed.colSpan === grid.columns ? borders.right : borders.insideVertical,
        }
        for (const side of ['top', 'bottom', 'left', 'right'] as const) {
          const value = borderCss(cell.borders?.[side] ?? edge[side])
          if (value) css.push(`border-${side}:${value}`)
        }
        if (cell.diagonal) {
          const line = color(look.outerBorder.color ?? theme.body.color)
          const direction = cell.diagonal === 'down' ? 'to top right' : 'to bottom right'
          css.push(`background-image:linear-gradient(${direction},transparent calc(50% - .5pt),${line} calc(50% - .5pt),${line} calc(50% + .5pt),transparent calc(50% + .5pt))`)
        }
        const inner = Math.max(12, columnPt.slice(placed.column, placed.column + placed.colSpan).reduce((a, b) => a + b, 0) - own.left - own.right)
        const align = cell.align ?? roleLook.align ?? (role === 'body' ? table.columnAlign?.[placed.column] : undefined)
        const base: TextStyle = { ...(table.size === undefined ? {} : { size: table.size }), ...roleLook.text, ...(cell.text ?? {}) }
        let content: string
        if (role === 'spacer') content = ''
        else if (cell.blocks?.length) {
          content = `<div class="dk-cell">${blocksHtml(cell.blocks, { ...ctx, width: inner, inCell: true })}</div>`
        } else {
          const pcss = [...textCss(base)]
          if (align) pcss.push(`text-align:${align}`)
          if (table.line !== undefined) pcss.push(`line-height:${Math.round(table.line * LINE * 1000) / 1000}`)
          const text =
            cell.content === null && ctx.state.nullCells === 'rule'
              ? `<span class="dk-blank"${styleAttr(`width:${pt(Math.min(inner * 0.8, theme.blankWidth))}`)}>&nbsp;</span>`
              : cell.content === null || cell.content === undefined
                ? ''
                : tabbed(inlineHtml(cell.content, { ...ctx, width: inner }, { style: base, ...(cell.source ? { source: cell.source } : {}) }), undefined, ctx)
          content = styled('KitTable', text, pcss)
        }
        if (cell.textDirection) {
          content = `<div${styleAttr(`writing-mode:vertical-rl;${cell.textDirection === 'up' ? 'transform:rotate(180deg);' : ''}margin:auto`)}>${content}</div>`
        }
        const span = `${placed.colSpan > 1 ? ` colspan="${placed.colSpan}"` : ''}${placed.rowSpan > 1 ? ` rowspan="${placed.rowSpan}"` : ''}`
        return `<${tag}${span}${tag === 'th' ? ' scope="col"' : ''}${styleAttr(css)}>${content}</${tag}>`
      })
    const least = role === 'spacer' ? (row.height ?? look.spacerHeight) : (row.height ?? (role === 'header' ? look.headerRowHeight : look.rowHeight))
    const trCss: string[] = []
    if (least !== undefined) trCss.push(`height:${pt(least)}`)
    if (row.cantSplit ?? table.keepTogether ?? table.cantSplit) trCss.push('break-inside:avoid')
    const keepNext = row.keepWithNext ?? (r < lastRow && (table.keepTogether === true || role === 'header'))
    if (keepNext) trCss.push('break-after:avoid')
    return `<tr${styleAttr(trCss)}>${cells.join('')}</tr>`
  })
  const cols = `<colgroup>${columnPt.map((w) => `<col${styleAttr(`width:${Math.round((w / widthPt) * 1e5) / 1e3}%`)}>`).join('')}</colgroup>`
  const head = repeat ? `<thead>${rows.slice(0, repeat).join('')}</thead>` : ''
  const self = table.align === 'center' ? 'center' : table.align === 'right' ? 'flex-end' : 'flex-start'
  const tcss = [
    `width:${Math.round(share * 1000) / 1000}%`,
    `table-layout:${table.autoFit ? 'auto' : 'fixed'}`,
    `align-self:${self}`,
    ...(table.keepTogether ? ['break-inside:avoid'] : []),
  ]
  let out = ''
  const above = table.caption && (table.captionPosition ?? 'above') === 'above'
  const numbered = number === undefined ? ctx : { ...ctx, state: { ...ctx.state, captions: new Map([...ctx.state.captions, [table, number]]) } }
  if (table.caption && above) out += captionHtml(table, table.caption, numbered, ['break-after:avoid'])
  out += `<table class="dk-table"${styleAttr(tcss)}>${cols}${head}<tbody>${rows.slice(repeat).join('')}</tbody></table>`
  if (table.caption && !above) out += captionHtml(table, table.caption, numbered)
  if (table.note !== undefined) out += styled('KitNote', inlineHtml(table.note, ctx, table.source ? { source: table.source } : {}))
  const after = table.spaceAfter ?? theme.table.spaceAfter ?? 0
  if (after > 0 && ctx.part === 'body' && !ctx.inCell && !ctx.inFrame) out += `<div class="dk-gap"${styleAttr(`height:${pt(after)};flex:none`)}></div>`
  return out
}

function spreadsheetHtml(block: SpreadsheetBlock, ctx: Ctx): string {
  if (!block.workbook) {
    const label = block.placeholder ?? (typeof block.sheet === 'string' ? `Spreadsheet: ${block.sheet}` : 'Spreadsheet')
    return placeholderHtml(label, ctx.width, 72, ctx)
  }
  const table = spreadsheetTable(block, { width: ctx.width })
  return tableHtml(table, ctx, ctx.state.captions.get(block))
}

/* ------------------------------------------------------------------ pictures */

function placeholderHtml(text: RichText | undefined, width: number, height: number, ctx: Ctx): string {
  const look = ctx.state.theme.table
  const css = [
    `width:${pt(Math.min(width, ctx.width))}`,
    `min-height:${pt(height)}`,
    `background-color:${color(look.placeholderFill)}`,
    `border:${borderCss(look.placeholderBorder) ?? 'none'}`,
    ...textCss(textStyleOf(look.placeholderText)),
  ]
  return `<div class="dk-placeholder" role="img" aria-label="${escapeHtml(plainText(text ?? ''))}"${styleAttr(css)}>${text === undefined ? '' : inlineHtml(text, ctx)}</div>`
}

/** The picture's size and look: fitted into its box as the Word writer fits it, cropped and framed. */
function pictureHtml(image: ImageData, box: { width: number; height: number; upscale: boolean }, alt: string | undefined, crop: PictureCrop | undefined, border: { width: number; color: string } | undefined): string {
  const info = resolveImage(image)
  const { top = 0, right = 0, bottom = 0, left = 0 } = crop ?? {}
  const ok = [top, right, bottom, left].every((share) => Number.isFinite(share) && share >= 0 && share < 1)
  if (!ok || left + right >= 1 || top + bottom >= 1) {
    throw new Error(`document-kit: crop ${JSON.stringify(crop)} must cut away shares from 0 to 1 that leave part of the picture`)
  }
  const shown = { width: info.width * (1 - left - right), height: info.height * (1 - top - bottom) }
  const size = fitImage(shown, box, box.upscale)
  const frame = border ? [`outline:${pt(border.width)} solid ${color(border.color)}`, `outline-offset:${pt(-border.width / 2)}`] : []
  if (!crop || !(top || right || bottom || left)) return imgHtml(image, size.width, size.height, alt, frame)
  const whole = { width: size.width / (1 - left - right), height: size.height / (1 - top - bottom) }
  return (
    `<span class="dk-crop"${styleAttr([`width:${pt(size.width)}`, `height:${pt(size.height)}`, ...frame])}>` +
    imgHtml(image, whole.width, whole.height, alt, [`left:${pt(-left * whole.width)}`, `top:${pt(-top * whole.height)}`]) +
    '</span>'
  )
}

function figureHtml(block: FigureBlock, ctx: Ctx): string {
  const theme = ctx.state.theme
  let box: { width: number; height: number; upscale: boolean; explicit: boolean }
  if (block.fit === 'page') {
    const lines = (block.caption ? 2 : 0) + (block.note !== undefined ? 2 : 0)
    const size = theme.caption.size ?? theme.body.size
    const room = lines * size * 1.3 + (theme.caption.before ?? 0) + (theme.caption.after ?? 0)
    box = { width: ctx.width, height: ctx.page.contentHeight - room - 12, upscale: true, explicit: true }
  } else {
    const fit = block.fit ?? {}
    box = {
      width: Math.min(fit.width ?? ctx.width, ctx.width),
      height: fit.height ?? ctx.page.contentHeight * 0.5,
      upscale: fit.upscale ?? false,
      explicit: fit.height !== undefined,
    }
  }
  const align = block.align ?? 'center'
  const above = block.caption && block.captionPosition === 'above'
  let out = ''
  if (block.caption && above) out += captionHtml(block, block.caption, ctx, ['break-after:avoid', `text-align:${align}`])
  if (block.image) {
    const alt = block.alt ?? (block.caption ? plainText(block.caption.text) : undefined)
    const border = block.border ? { width: block.border.width ?? 0.5, color: block.border.color ?? theme.body.color } : undefined
    const keep = block.caption !== undefined && !above ? ';break-after:avoid' : ''
    out += `<p class="dk-p dk-figure"${styleAttr(`text-align:${align};line-height:1;margin:0${keep}`)}>${pictureHtml(block.image, box, alt, block.crop, border)}</p>`
  } else {
    out += placeholderHtml(block.placeholder, box.width, box.explicit ? box.height : Math.min(box.height, 144), ctx)
  }
  if (block.caption && !above) out += captionHtml(block, block.caption, ctx, [`text-align:${align}`])
  if (block.note !== undefined) out += styled('KitNote', inlineHtml(block.note, ctx, block.source ? { source: block.source } : {}), [`text-align:${align}`])
  return out
}

function figureGridHtml(block: FigureGridBlock, ctx: Ctx): string {
  const columns = Math.max(1, Math.round(block.columns))
  const look = ctx.state.theme.table
  const inner = ctx.width / columns - look.cellPadding.left - look.cellPadding.right
  const itemHeight = block.itemHeight ?? inner * 0.75
  const labelAbove = block.labelPosition === 'above'
  const items = block.items.map((item) => {
    const label = item.label === undefined ? '' : styled('KitNote', inlineHtml(item.label, ctx), ['text-align:center'])
    const picture = item.image
      ? `<p class="dk-p" style="text-align:center;line-height:1">${pictureHtml(item.image, { width: inner, height: itemHeight, upscale: block.upscale ?? false }, item.alt, undefined, undefined)}</p>`
      : placeholderHtml(item.placeholder, inner, Math.min(itemHeight, 144), { ...ctx, width: inner })
    const pad = `padding:${pt(look.cellPadding.top)} ${pt(look.cellPadding.right)} ${pt(look.cellPadding.bottom)} ${pt(look.cellPadding.left)}`
    return `<div class="dk-flow"${styleAttr(`${pad};break-inside:avoid`)}>${labelAbove ? label + picture : picture + label}</div>`
  })
  let out = `<div class="dk-grid"${styleAttr(`grid-template-columns:repeat(${columns},minmax(0,1fr))`)}>${items.join('')}</div>`
  if (block.caption) out += captionHtml(block, block.caption, ctx, ['text-align:center'])
  return out
}

function floatingImageHtml(block: FloatingImageBlock, ctx: Ctx): string {
  if (!block.image) return ''
  const page = ctx.page
  const box = block.box === 'page' ? { x: 0, y: 0, width: page.width, height: page.height } : block.box
  const margin = block.box !== 'page' && block.relativeTo === 'margin'
  const x = box.x + (margin ? page.margins.left : 0)
  const y = box.y + (margin ? page.margins.top : 0)
  ctx.sheet.placed.push(
    `<img class="dk-float" src="${dataUrl(block.image)}" alt="${escapeHtml(block.alt ?? '')}"${styleAttr(`left:${pt(x)};top:${pt(y)};width:${pt(box.width)};height:${pt(box.height)};z-index:${block.behind === false ? 2 : -1}`)}>`,
  )
  return ''
}

function frameHtml(block: FrameBlock, ctx: Ctx): string {
  if (ctx.inFrame) return blocksHtml(block.blocks, ctx)
  const page = ctx.page
  const margin = block.relativeTo === 'margin'
  const x = block.x + (margin ? page.margins.left : 0)
  const y = block.y + (margin ? page.margins.top : 0)
  const css = [`left:${pt(x)}`, `top:${pt(y)}`, `width:${pt(block.width)}`, 'z-index:1']
  if (block.height !== undefined && block.height > 0) {
    css.push(block.heightRule === 'exact' ? `height:${pt(block.height)};overflow:hidden` : `min-height:${pt(block.height)}`)
  }
  const inner = blocksHtml(block.blocks, { ...ctx, width: block.width, inFrame: true })
  ctx.sheet.placed.push(`<div class="dk-frame"${styleAttr(css)}>${inner}</div>`)
  return ''
}

/* ------------------------------------------------------------------ signature */

function signatureHtml(block: SignatureBlock, ctx: Ctx): string {
  const align = block.align ?? 'left'
  const ruleWidth = Math.min(block.ruleWidth ?? 216, ctx.width)
  const self = align === 'center' ? 'center' : align === 'right' ? 'flex-end' : 'flex-start'
  const labelStyle = block.labelStyle ?? { bold: true }
  const nameStyle = block.nameStyle ?? { bold: true }
  const labelBelow = block.labelPosition === 'below'
  const normal = paragraphStyle({}, ctx)
  const p = (content: string, css: string[]) => `<p class="dk-p dk-ps-${normal}"${styleAttr([`text-align:${align}`, ...css])}>${content}</p>`
  const parts: string[] = []
  const label = () => p(inlineHtml(block.label!, ctx, { style: labelStyle }), ['margin-top:0', `margin-bottom:${pt(labelBelow ? 2 : (block.labelGap ?? 4))}`])
  if (block.label !== undefined && !labelBelow) parts.push(label())
  if (block.image) {
    const box = block.imageBox ?? { width: 140, height: 55 }
    parts.push(p(pictureHtml(block.image, { ...box, upscale: block.imageUpscale ?? true }, 'Signature', undefined, undefined), ['margin:0', 'line-height:1']))
  } else {
    parts.push(`<div${styleAttr(`height:${pt(Math.max(1, block.signingSpace ?? 36))};flex:none`)}></div>`)
  }
  const rule = borderCss({ style: 'single', width: block.ruleThickness ?? 0.75, color: block.ruleColor ?? ctx.state.theme.body.color })
  const name = block.name === null || block.name === undefined ? '' : inlineHtml(block.name, ctx, { style: nameStyle, ...(block.source ? { source: block.source } : {}) })
  const title = block.title === null || block.title === undefined ? '' : inlineHtml(block.title, ctx, { style: nameStyle, ...(block.source ? { source: block.source } : {}) })
  const sameLine = (block.titleLayout ?? 'sameLine') === 'sameLine'
  const sep = name && title ? inlineHtml(block.titleSeparator ?? ', ', ctx, { style: nameStyle }) : ''
  const lineCss = [`width:${pt(ruleWidth)}`, `align-self:${self}`, 'margin-top:0', 'margin-bottom:2pt']
  parts.push(p(sameLine ? name + sep + title : name, [...lineCss, `border-top:${rule}`, `padding-top:${pt(Math.max(0, block.ruleGap ?? 2))}`]))
  if (!sameLine && title) parts.push(p(title, lineCss))
  for (const line of block.lines ?? []) {
    const labelText = plainText(line.label).trim()
    const separator = labelText ? (/[:]\s*$/.test(plainText(line.label)) ? ' ' : ': ') : ''
    const value =
      line.value === null
        ? `<span class="dk-blank"${styleAttr(`width:${pt(ruleWidth * 0.6)}`)}>&nbsp;</span>`
        : inlineHtml(line.value, ctx, line.source ? { source: line.source } : {})
    parts.push(p(`${labelText ? inlineHtml(line.label, ctx) : ''}${escapeHtml(separator)}${value}`, lineCss))
  }
  if (block.label !== undefined && labelBelow) parts.push(label())
  const css = [
    `margin-top:${pt(block.spaceAbove ?? 12)}`,
    `margin-bottom:${pt(block.spaceBelow ?? 2)}`,
    ...(block.keepTogether !== false ? ['break-inside:avoid'] : []),
  ]
  return `<div class="dk-signature dk-flow"${styleAttr(css)}>${parts.join('')}</div>`
}

/* ------------------------------------------------------------------ contents */

function tocHtml(block: TocBlock, ctx: Ctx): string {
  const index = ctx.state.tocs.length
  ctx.state.tocs.push(block)
  const title = block.title !== undefined ? `<p class="dk-p dk-ps-KitTocTitle" style="break-after:avoid">${inlineHtml(block.title, ctx)}</p>` : ''
  return `${title}<nav class="dk-toc dk-flow"><!--dk-toc-${index}--></nav>`
}

function tocEntries(block: TocBlock, state: State): string {
  const levels = Math.min(6, Math.max(1, Math.round(block.levels ?? 3)))
  return state.headings
    .filter((h) => h.level <= levels)
    .map(
      (h) =>
        `<p class="dk-p dk-ps-Normal"${styleAttr(`display:flex;gap:4pt;margin-left:${pt((h.level - 1) * 11)}`)}>` +
        `<a href="#${escapeHtml(h.id)}" style="color:inherit;text-decoration:none">${escapeHtml(h.text)}</a>` +
        `<span style="flex:1;border-bottom:.75pt dotted currentColor;margin-bottom:.3em"></span>` +
        `<span>${escapeHtml(h.page)}</span></p>`,
    )
    .join('')
}
