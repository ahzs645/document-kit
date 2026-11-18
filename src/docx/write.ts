/** The Word writer: the block model to a .docx file's bytes, in a browser or in Node. */
import {
  Document as DocxDocument,
  Footer,
  Header,
  Packer,
  PageOrientation,
  Paragraph,
  Table,
  TableOfContents,
  VerticalAlignSection,
  type ISectionOptions,
} from 'docx'
import {
  expandDocument,
  numberCaptions,
  pageBox,
  resolveTheme,
  type CustomRegistry,
} from '../model/index'
import type { Block, Document, Section, TextStyle } from '../model/types'
import { writeBlocks } from './blocks'
import type { WriterContext } from './context'
import { normalizePackage } from './package'
import { createListRegistry, stylesFor } from './styles'
import { twips } from './units'

export interface WriteDocxOptions {
  /** Expanders for the document's custom blocks. A custom block of a kind not registered here stops the writer. */
  registry?: CustomRegistry
  /**
   * How text carrying a source tag is marked, by tag: `{ reviewer: { highlight: 'yellow' } }` sets apart every value
   * the reviewer gave. Unlisted tags, and every tag when this is absent, print plainly.
   */
  sourceStyles?: Readonly<Record<string, TextStyle>>
  /** A null table cell prints empty (default) or as a blank rule. */
  nullCells?: 'empty' | 'rule'
  /** The application named in the file's properties (`docProps/core.xml` creator), when the author is not given. */
  creator?: string
}

const SECTION_ALIGN = {
  top: VerticalAlignSection.TOP,
  center: VerticalAlignSection.CENTER,
  bottom: VerticalAlignSection.BOTTOM,
} as const

/**
 * Writes the document as a Word file. Custom blocks are expanded through `options.registry` first; tables and figures
 * are numbered in reading order. Resolves to the file's bytes.
 */
export async function writeDocx(document: Document, options: WriteDocxOptions = {}): Promise<Uint8Array> {
  const docx = buildDocx(document, options)
  const bytes = await Packer.pack(docx, 'uint8array')
  const created = document.metadata?.created
  const modified = document.metadata?.modified
  return normalizePackage(bytes, {
    ...(created ? { created } : {}),
    ...(modified ? { modified } : {}),
  })
}

/** The document as the docx library's object, before it is packed. */
function buildDocx(source: Document, options: WriteDocxOptions): DocxDocument {
  if (!source.sections.length) throw new Error('document-kit: a document needs at least one section')
  const document = expandDocument(source, options.registry)
  const theme = resolveTheme(document.theme)
  const captions = numberCaptions(document)
  const lists = createListRegistry()
  const sections = document.sections.map((section) => {
    const page = pageBox(section.page, theme)
    const context: WriterContext = {
      theme,
      page,
      width: page.contentWidth,
      part: 'body',
      captions,
      sourceStyles: options.sourceStyles ?? {},
      nullCells: options.nullCells ?? 'empty',
      lists,
    }
    return writeSection(section, context)
  })
  const metadata = document.metadata ?? {}
  return new DocxDocument({
    sections,
    styles: stylesFor(theme),
    numbering: lists.numbering(),
    ...(metadata.title ? { title: metadata.title } : {}),
    ...(metadata.subject ? { subject: metadata.subject } : {}),
    ...(metadata.description ? { description: metadata.description } : {}),
    ...(metadata.keywords?.length ? { keywords: metadata.keywords.join(', ') } : {}),
    creator: metadata.author ?? options.creator ?? 'document-kit',
    lastModifiedBy: metadata.author ?? options.creator ?? 'document-kit',
  })
}

function writeSection(section: Section, context: WriterContext): ISectionOptions {
  const { page } = context
  const part = (blocks: readonly Block[], kind: 'header' | 'footer'): Array<Paragraph | Table> => {
    // A table of contents has no place in a header or footer.
    const written = writeBlocks(blocks, { ...context, part: kind }).filter(
      (child): child is Paragraph | Table => !(child instanceof TableOfContents),
    )
    return written.length ? written : [new Paragraph({ style: kind === 'header' ? 'Header' : 'Footer', children: [] })]
  }
  const headers: { default?: Header; first?: Header } = {}
  const footers: { default?: Footer; first?: Footer } = {}
  if (section.header) headers.default = new Header({ children: part(section.header, 'header') })
  if (section.footer) footers.default = new Footer({ children: part(section.footer, 'footer') })
  if (section.firstPage) {
    headers.first = new Header({ children: part(section.firstPage.header ?? [], 'header') })
    footers.first = new Footer({ children: part(section.firstPage.footer ?? [], 'footer') })
  }
  const children = writeBlocks(section.body, context)
  if (!children.length || !(children[children.length - 1] instanceof Paragraph)) {
    // Word wants a paragraph to close a section; a table cannot be last.
    children.push(new Paragraph({ spacing: { before: 0, after: 0 }, children: [] }))
  }
  const numbers = section.pageNumbers
  return {
    properties: {
      page: {
        size: {
          // docx takes the portrait size and turns it for a landscape orientation.
          width: twips(Math.min(page.width, page.height)),
          height: twips(Math.max(page.width, page.height)),
          orientation: page.orientation === 'landscape' ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT,
        },
        margin: {
          top: twips(page.margins.top),
          right: twips(page.margins.right),
          bottom: twips(page.margins.bottom),
          left: twips(page.margins.left),
          header: twips(page.margins.header),
          footer: twips(page.margins.footer),
        },
        ...(numbers
          ? {
              pageNumbers: {
                ...(numbers.start !== undefined ? { start: numbers.start } : {}),
                ...(numbers.format ? { formatType: numbers.format } : {}),
              },
            }
          : {}),
      },
      ...(section.firstPage ? { titlePage: true } : {}),
      ...(section.page?.verticalAlign ? { verticalAlign: SECTION_ALIGN[section.page.verticalAlign] } : {}),
    },
    ...(Object.keys(headers).length ? { headers } : {}),
    ...(Object.keys(footers).length ? { footers } : {}),
    children,
  }
}
