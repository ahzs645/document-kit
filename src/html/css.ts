/**
 * The look of an HTML document from the same data the Word writer reads: the theme, the style table
 * (`src/model/styles.ts`) and the caller's fonts. One CSS class a named style, its look merged through the styles it
 * is based on, so a paragraph in style `AffiliationDepartment` looks in the browser as Word sets it.
 */
import { fontEmbeddingReport, groupFonts, type NamedFont, type StyleTable } from '../model/index'
import type { Border, Highlight, TextStyle, Theme, ThemeText } from '../model/types'

/** A number of points as CSS, at most three decimals. */
export const pt = (value: number): string => `${Math.round(value * 1000) / 1000}pt`

/** A model colour as CSS `#rrggbb`; throws on what is not a hex colour, as the Word writer does. */
export function color(value: string): string {
  const cleaned = typeof value === 'string' ? value.trim().replace(/^#/, '') : ''
  if (/^[0-9a-fA-F]{6}$/.test(cleaned)) return `#${cleaned.toLowerCase()}`
  if (/^[0-9a-fA-F]{3}$/.test(cleaned)) return `#${[...cleaned].map((d) => d + d).join('').toLowerCase()}`
  throw new Error(`document-kit: ${JSON.stringify(value)} is not a hex colour; write it as #rrggbb or #rgb (the # is optional)`)
}

/** Text for HTML content or a quoted attribute value: `&`, `<`, `>`, `"` and `'` written as references. */
export const escapeHtml = (value: string): string =>
  value.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!)

/**
 * A CSS string literal. Quotes, backslashes, line ends and angle brackets are written as CSS escapes, so the literal
 * neither ends early nor closes the `<style>` element it may stand in. In a `style` attribute it still needs the
 * attribute's own escaping ({@link styleAttr}).
 */
export const cssString = (value: string): string =>
  `"${value.replace(/["\\\n\r<>]/g, (c) => `\\${c.charCodeAt(0).toString(16)} `)}"`

/**
 * A ` style="…"` attribute from CSS declarations, escaped as an attribute value (a font stack's quoted names would
 * otherwise end it); nothing when there are no declarations.
 */
export function styleAttr(css: string | readonly string[]): string {
  const text = typeof css === 'string' ? css : css.filter(Boolean).join(';')
  return text ? ` style="${escapeHtml(text)}"` : ''
}

/** Word's highlight colours as CSS. */
export const HIGHLIGHT: Readonly<Record<Highlight, string>> = {
  yellow: '#ffff00',
  green: '#00ff00',
  cyan: '#00ffff',
  magenta: '#ff00ff',
  blue: '#0000ff',
  red: '#ff0000',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#c0c0c0',
  black: '#000000',
  white: '#ffffff',
}

/** The open faces `document-kit/fonts` bundles for the Office names, tried after the named face. */
const STAND_INS: ReadonlyArray<[RegExp, string]> = [
  [/^calibri$/i, 'Carlito'],
  [/^cambria$/i, 'Caladea'],
  [/^(arial|helvetica)$/i, 'Arimo'],
  [/^times new roman$/i, 'Tinos'],
  [/^courier new$/i, 'Cousine'],
]

/** A font name as a CSS font stack: the face, its bundled stand-in, then the generic family of its kind. */
export function fontStack(name: string): string {
  const standIn = STAND_INS.find(([pattern]) => pattern.test(name.trim()))?.[1]
  const generic = /mono|courier|consol|code/i.test(name)
    ? 'monospace'
    : /serif|times|roman|cambria|georgia|garamond|book|minion|caslon|caladea|tinos/i.test(name) && !/sans/i.test(name)
      ? 'serif'
      : 'sans-serif'
  return [cssString(name.trim()), ...(standIn ? [cssString(standIn)] : []), generic].join(', ')
}

/**
 * Word's single line is the face's own height, about 1.2 em for the faces documents use; a line multiple is set on
 * that, so 1.15 lines is 1.38 em. The browser's metrics are not the layout's, so this is near, not exact.
 */
export const LINE = 1.2

/** The declarations for a text style: face, size, colour, weight, slant, decoration, caps, highlight, tracking. */
export function textCss(style: TextStyle | ThemeText): string[] {
  const out: string[] = []
  if (style.font) out.push(`font-family:${fontStack(style.font)}`)
  if (style.size !== undefined) out.push(`font-size:${pt(style.size)}`)
  if (style.color) out.push(`color:${color(style.color)}`)
  if (style.bold !== undefined) out.push(`font-weight:${style.bold ? 700 : 400}`)
  if (style.italic !== undefined) out.push(`font-style:${style.italic ? 'italic' : 'normal'}`)
  if (style.allCaps !== undefined) out.push(`text-transform:${style.allCaps ? 'uppercase' : 'none'}`)
  if (style.smallCaps !== undefined) out.push(`font-variant-caps:${style.smallCaps ? 'small-caps' : 'normal'}`)
  if (style.characterSpacing !== undefined) out.push(`letter-spacing:${pt(style.characterSpacing)}`)
  const run = style as TextStyle
  const lines = [run.underline ? 'underline' : '', run.strike ? 'line-through' : ''].filter(Boolean)
  if (lines.length) out.push(`text-decoration:${lines.join(' ')}`)
  if (run.highlight) out.push(`background-color:${HIGHLIGHT[run.highlight] ?? 'transparent'}`)
  return out
}

/** A paragraph look's declarations: the text, alignment, spacing before and after, line height and keep-with-next. */
export function paragraphCss(text: ThemeText): string[] {
  const out = textCss(text)
  if (text.align) out.push(`text-align:${text.align}`)
  if (text.before !== undefined) out.push(`margin-top:${pt(text.before)}`)
  if (text.after !== undefined) out.push(`margin-bottom:${pt(text.after)}`)
  if (text.line !== undefined) out.push(`line-height:${Math.round(text.line * LINE * 1000) / 1000}`)
  if (text.keepWithNext) out.push('break-after:avoid')
  return out
}

const BORDER_STYLES: Readonly<Record<NonNullable<Border['style']>, string>> = {
  single: 'solid',
  double: 'double',
  dashed: 'dashed',
  dotted: 'dotted',
  none: 'none',
}

/** A border as a CSS `border` value. */
export function borderCss(value: Border | undefined, fallbackColor = 'currentColor'): string | undefined {
  if (!value) return undefined
  const style = BORDER_STYLES[value.style ?? 'single']
  if (style === 'none') return 'none'
  const width = Math.max(0.25, value.width ?? 0.5)
  return `${pt(style === 'double' ? Math.max(width, 1.5) : width)} ${style} ${value.color ? color(value.color) : fallbackColor}`
}

/** The style table as CSS classes: `.dk-ps-<id>` for each paragraph style, `.dk-cs-<id>` for each character style. */
export function styleCss(table: StyleTable): string {
  const rules: string[] = []
  for (const style of table.paragraph.values()) {
    const declarations = paragraphCss(style.text)
    if (style.indent) {
      const { left = 0, right = 0, firstLine, hanging } = style.indent
      if (left) declarations.push(`margin-left:${pt(left)}`)
      if (right) declarations.push(`margin-right:${pt(right)}`)
      if (hanging) declarations.push(`text-indent:${pt(-hanging)}`)
      else if (firstLine) declarations.push(`text-indent:${pt(firstLine)}`)
    }
    if (style.keepLinesTogether) declarations.push('break-inside:avoid')
    if (style.shading) declarations.push(`background-color:${color(style.shading)}`)
    for (const side of ['top', 'bottom', 'left', 'right'] as const) {
      const value = style.borders?.[side]
      const css = borderCss(value)
      if (!css) continue
      declarations.push(`border-${side}:${css}`, `padding-${side}:${pt(value?.space ?? 0)}`)
    }
    rules.push(`.dk-ps-${style.id}{${declarations.join(';')}}`)
  }
  for (const style of table.character.values()) {
    rules.push(`.dk-cs-${style.id}{${textCss(style.text).join(';')}}`)
  }
  return rules.join('\n')
}

/** Base64 without Node's Buffer or a DOM: the bytes of a picture or a font for a `data:` URL. */
export function base64(bytes: Uint8Array): string {
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
  let out = ''
  let i = 0
  for (; i + 2 < bytes.length; i += 3) {
    const n = (bytes[i]! << 16) | (bytes[i + 1]! << 8) | bytes[i + 2]!
    out += alphabet[n >> 18]! + alphabet[(n >> 12) & 63]! + alphabet[(n >> 6) & 63]! + alphabet[n & 63]!
  }
  if (i < bytes.length) {
    const n = (bytes[i]! << 16) | ((bytes[i + 1] ?? 0) << 8)
    out += alphabet[n >> 18]! + alphabet[(n >> 12) & 63]!
    out += i + 1 < bytes.length ? `${alphabet[(n >> 6) & 63]!}=` : '=='
  }
  return out
}

/**
 * `@font-face` rules for the caller's fonts, each face the report lets an HTML print view carry, declared at the
 * weight of the slot it fills (400 or 700) so the browser picks the faces Word and the layout pick.
 */
export function fontFaceCss(fonts: readonly NamedFont[]): string {
  if (!fonts.length) return ''
  const report = fontEmbeddingReport(fonts, 'html')
  const { families } = groupFonts(fonts.filter((_, i) => report[i]!.embedded))
  const rules: string[] = []
  for (const { family, faces } of families.values()) {
    for (const [slot, font] of Object.entries(faces)) {
      if (!font) continue
      const bytes = font.bytes instanceof Uint8Array ? font.bytes : new Uint8Array(font.bytes)
      const otf = bytes[0] === 0x4f && bytes[1] === 0x54 && bytes[2] === 0x54 && bytes[3] === 0x4f
      const weight = slot === 'bold' || slot === 'boldItalic' ? 700 : 400
      const italic = slot === 'italic' || slot === 'boldItalic'
      rules.push(
        `@font-face{font-family:${cssString(family)};font-weight:${weight};font-style:${italic ? 'italic' : 'normal'};` +
          `src:url(data:font/${otf ? 'otf' : 'ttf'};base64,${base64(bytes)}) format("${otf ? 'opentype' : 'truetype'}");font-display:block}`,
      )
    }
  }
  return rules.join('\n')
}

/** The page furniture's CSS that does not depend on the document. */
export const BASE_CSS = `*,*::before,*::after{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--dk-desk,#e5e7eb)}
.dk-sheet{position:relative;isolation:isolate;margin:16px auto;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.2);overflow:hidden;display:flex;flex-direction:column;break-after:page}
.dk-sheet:last-child{break-after:auto}
.dk-flow{display:flex;flex-direction:column;min-width:0}
.dk-body{flex:1 1 auto}
.dk-header,.dk-footer{position:absolute;display:flex;flex-direction:column}
.dk-foot{margin-top:auto;display:flex;flex-direction:column}
.dk-body>:first-child,.dk-frame>:first-child,.dk-header>:first-child{margin-top:0}
.dk-noteref>a{color:inherit;text-decoration:none}
.dk-p{margin:0;overflow-wrap:break-word}
.dk-p:empty::after{content:"\\200b"}
.dk-tabbed{position:relative}
.dk-tabbed>.dk-at{position:absolute;top:0;white-space:nowrap}
.dk-tab{white-space:pre}
.dk-list{margin:0;padding:0 0 0 18pt}
.dk-list>li{padding-left:0}
.dk-list>li>.dk-p{display:block}
.dk-table{border-collapse:collapse;border-spacing:0}
.dk-table td,.dk-table th{font-weight:inherit;text-align:inherit;vertical-align:top}
.dk-table thead{display:table-header-group}
.dk-cell{display:flex;flex-direction:column}
.dk-crop{overflow:hidden;position:relative;display:inline-block;vertical-align:bottom}
.dk-crop>img{position:absolute;max-width:none}
.dk-placeholder{display:flex;align-items:center;justify-content:center;text-align:center;margin:0 auto}
.dk-grid{display:grid;gap:0}
.dk-blank{display:inline-block;border-bottom:.75pt solid currentColor;vertical-align:baseline}
.dk-box{font-family:"Segoe UI Symbol","MS Gothic","DejaVu Sans",sans-serif}
.dk-frame{position:absolute;display:flex;flex-direction:column}
.dk-float{position:absolute;display:block;max-width:none}
.dk-footnotes{margin-top:6pt}
.dk-footnotes::before{content:"";display:block;width:144pt;max-width:100%;border-top:.5pt solid currentColor;margin-bottom:3pt}
@media print{
body{background:none}
.dk-sheet{margin:0;box-shadow:none}
*{-webkit-print-color-adjust:exact;print-color-adjust:exact}
}`
