/**
 * The named styles of a document, resolved once for every writer: the theme's own (Normal, Title, the six headings,
 * Caption, Note, Small, Header, Footer, table text) and the paragraph and character styles the document defines
 * (`document.styles`). The Word writer writes them as Word styles; the HTML writer as CSS classes. Both read this
 * table, so a style looks the same in either.
 */
import type {
  Border,
  CharacterStyleDefinition,
  DocumentStyles,
  Indent,
  ParagraphStyleDefinition,
  ParagraphStyleRef,
  TextStyle,
  Theme,
  ThemeText,
} from './types'

/** The Word style id of each theme style a block or a definition can name. */
export const THEME_STYLE_IDS: Readonly<Record<string, string>> = {
  normal: 'Normal',
  title: 'Title',
  subtitle: 'Subtitle',
  caption: 'Caption',
  note: 'KitNote',
  small: 'KitSmall',
  heading1: 'Heading1',
  heading2: 'Heading2',
  heading3: 'Heading3',
  heading4: 'Heading4',
  heading5: 'Heading5',
  heading6: 'Heading6',
  header: 'Header',
  footer: 'Footer',
  table: 'KitTable',
}

/** Ids the theme's styles take in a Word file; a defined style may not take one (in any case). */
const RESERVED = new Set(
  [...Object.keys(THEME_STYLE_IDS), ...Object.values(THEME_STYLE_IDS), 'KitTocTitle', 'ListParagraph', 'Hyperlink', 'FootnoteText', 'FootnoteReference']
    .map((id) => id.toLowerCase()),
)

const ID = /^[A-Za-z][A-Za-z0-9]{0,99}$/

/** A paragraph style, every look it takes from the styles under it merged in. */
export interface ResolvedParagraphStyle {
  /** The Word style id. */
  readonly id: string
  readonly name: string
  /** The Word id of the style it is based on; none for Normal. */
  readonly basedOn?: string
  readonly next?: string
  /** True for a style the document defines; false for the theme's. */
  readonly defined: boolean
  /** The look, merged through the styles it is based on down to the body text. */
  readonly text: ThemeText
  readonly indent?: Indent
  readonly borders?: { top?: Border; bottom?: Border; left?: Border; right?: Border }
  readonly shading?: string
  readonly keepLinesTogether?: boolean
  /** Its own definition, for the Word writer (which writes only what the style itself says). */
  readonly definition?: ParagraphStyleDefinition
}

/** A character style the document defines, merged through the styles it is based on. */
export interface ResolvedCharacterStyle {
  readonly id: string
  readonly name: string
  readonly basedOn?: string
  readonly text: Omit<TextStyle, 'characterStyle'>
  readonly definition: CharacterStyleDefinition
}

/** Every named style of a document. */
export interface StyleTable {
  readonly paragraph: ReadonlyMap<string, ResolvedParagraphStyle>
  readonly character: ReadonlyMap<string, ResolvedCharacterStyle>
}

/** An error in the styles a document defines, or a reference to one it does not. */
export class StyleError extends Error {
  constructor(message: string) {
    super(`document-kit: ${message}`)
    this.name = 'StyleError'
  }
}

const defined = (value: ThemeText | undefined): ThemeText =>
  Object.fromEntries(Object.entries(value ?? {}).filter(([, v]) => v !== undefined)) as ThemeText

/** The theme's paragraph styles, resolved, in the order the Word writer writes them. */
function themeStyles(theme: Theme): ResolvedParagraphStyle[] {
  const body: ThemeText = { after: 6, line: 1.15, ...defined(theme.body) }
  const over = (id: string, name: string, text: ThemeText, extra: Partial<ResolvedParagraphStyle> = {}) => ({
    id,
    name,
    basedOn: 'Normal',
    next: 'Normal',
    defined: false,
    text: { ...body, ...defined(text) },
    ...extra,
  })
  const headings = theme.headings.map((level, index) =>
    over(`Heading${index + 1}`, `heading ${index + 1}`, {
      ...level,
      color: level.color ?? theme.body.color,
      keepWithNext: level.keepWithNext ?? true,
    }),
  )
  return [
    { id: 'Normal', name: 'Normal', defined: false, text: body },
    over('Title', 'Title', { ...theme.title, color: theme.title.color ?? theme.body.color }),
    over('Subtitle', 'Subtitle', theme.subtitle),
    ...headings,
    over('Caption', 'caption', theme.caption),
    over('KitNote', 'Note', theme.note),
    over('KitSmall', 'Small', theme.small),
    over('Header', 'header', { after: 0, line: 1, ...defined(theme.header) }),
    over('Footer', 'footer', { after: 0, line: 1, ...defined(theme.footer) }),
    over('KitTable', 'Table text', { size: theme.table.size, before: 0, after: 0, line: theme.table.line ?? 1 }),
  ]
}

/** The Word id a definition's `basedOn` or `next` names: a theme style by its model name, or a defined id. */
function referenced(name: string, ids: ReadonlySet<string>, what: string): string {
  const themeId = THEME_STYLE_IDS[name]
  if (themeId) return themeId
  if (ids.has(name)) return name
  throw new StyleError(`${what} names the style "${name}", which is neither a theme style nor one the document defines`)
}

function checkId(id: unknown, seen: Set<string>, kind: string): asserts id is string {
  if (typeof id !== 'string' || !ID.test(id)) {
    throw new StyleError(`the ${kind} style id ${JSON.stringify(id)} must be a letter followed by letters and digits`)
  }
  if (RESERVED.has(id.toLowerCase())) throw new StyleError(`the ${kind} style id "${id}" is the theme's own; choose another`)
  if (seen.has(id.toLowerCase())) throw new StyleError(`the style id "${id}" is defined twice`)
  seen.add(id.toLowerCase())
}

/**
 * The document's named styles: the theme's and its own, each resolved through the styles it is based on. Throws a
 * {@link StyleError} on an id that is malformed, taken twice or the theme's own, on a `basedOn` or `next` naming no
 * style, and on styles based on each other in a circle.
 */
export function styleTable(theme: Theme, styles: DocumentStyles | undefined): StyleTable {
  const paragraph = new Map<string, ResolvedParagraphStyle>()
  for (const style of themeStyles(theme)) paragraph.set(style.id, style)
  const character = new Map<string, ResolvedCharacterStyle>()
  const seen = new Set<string>()
  const paragraphs = styles?.paragraph ?? []
  const characters = styles?.character ?? []
  for (const definition of paragraphs) checkId(definition.id, seen, 'paragraph')
  for (const definition of characters) checkId(definition.id, seen, 'character')
  const paragraphIds = new Set(paragraphs.map((d) => d.id))
  const characterIds = new Set(characters.map((d) => d.id))

  const byId = new Map(paragraphs.map((d) => [d.id, d]))
  const resolving = new Set<string>()
  const resolve = (definition: ParagraphStyleDefinition): ResolvedParagraphStyle => {
    const done = paragraph.get(definition.id)
    if (done) return done
    if (resolving.has(definition.id)) throw new StyleError(`the paragraph style "${definition.id}" is based on itself`)
    resolving.add(definition.id)
    const basedOn = referenced(definition.basedOn ?? 'normal', paragraphIds, `the paragraph style "${definition.id}"`)
    const next = definition.next === undefined ? undefined : referenced(definition.next, paragraphIds, `the paragraph style "${definition.id}"`)
    const base = byId.has(basedOn) ? resolve(byId.get(basedOn)!) : paragraph.get(basedOn)!
    const own = defined(definition.text)
    const resolved: ResolvedParagraphStyle = {
      id: definition.id,
      name: definition.name ?? definition.id,
      basedOn,
      ...(next ? { next } : {}),
      defined: true,
      text: { ...base.text, ...own },
      ...((definition.indent ?? base.indent) ? { indent: { ...base.indent, ...definition.indent } } : {}),
      ...((definition.borders ?? base.borders) ? { borders: { ...base.borders, ...definition.borders } } : {}),
      ...((definition.shading ?? base.shading) ? { shading: (definition.shading ?? base.shading)! } : {}),
      ...((definition.keepLinesTogether ?? base.keepLinesTogether) !== undefined
        ? { keepLinesTogether: (definition.keepLinesTogether ?? base.keepLinesTogether)! }
        : {}),
      definition,
    }
    resolving.delete(definition.id)
    paragraph.set(definition.id, resolved)
    return resolved
  }
  for (const definition of paragraphs) resolve(definition)

  const characterById = new Map(characters.map((d) => [d.id, d]))
  const resolveCharacter = (definition: CharacterStyleDefinition): ResolvedCharacterStyle => {
    const done = character.get(definition.id)
    if (done) return done
    if (resolving.has(definition.id)) throw new StyleError(`the character style "${definition.id}" is based on itself`)
    resolving.add(definition.id)
    let base: Omit<TextStyle, 'characterStyle'> = {}
    if (definition.basedOn !== undefined) {
      const parent = characterById.get(definition.basedOn)
      if (!parent || !characterIds.has(definition.basedOn)) {
        throw new StyleError(`the character style "${definition.id}" is based on "${definition.basedOn}", which the document does not define`)
      }
      base = resolveCharacter(parent).text
    }
    const resolved: ResolvedCharacterStyle = {
      id: definition.id,
      name: definition.name ?? definition.id,
      ...(definition.basedOn !== undefined ? { basedOn: definition.basedOn } : {}),
      text: { ...base, ...(definition.text as Omit<TextStyle, 'characterStyle'>) },
      definition,
    }
    resolving.delete(definition.id)
    character.set(definition.id, resolved)
    return resolved
  }
  for (const definition of characters) resolveCharacter(definition)
  return { paragraph, character }
}

/**
 * The Word style id of a paragraph's `style`: the theme style's (none for `normal`, which is the default), or a
 * defined style's id. Throws a {@link StyleError} naming a style the table does not have.
 */
export function paragraphStyleId(table: StyleTable, style: ParagraphStyleRef | undefined): string | undefined {
  if (style === undefined || style === 'normal') return undefined
  const themeId = THEME_STYLE_IDS[style]
  if (themeId && !style.startsWith('heading') && !['header', 'footer', 'table'].includes(style)) return themeId
  if (table.paragraph.get(style)?.defined) return style
  throw new StyleError(`a paragraph names the style "${style}", which is neither a theme style nor one the document defines`)
}

/** Checks a run's `characterStyle` against the table; throws a {@link StyleError} naming one it does not have. */
export function characterStyleId(table: StyleTable, id: string | undefined): string | undefined {
  if (id === undefined) return undefined
  if (table.character.has(id)) return id
  throw new StyleError(`a run names the character style "${id}", which the document does not define`)
}
