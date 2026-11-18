export type {
  HyperlinkResolver,
  ImageResolver,
  ParseContext,
  ResolvedDiagram,
} from './document-parser';
export type { BlockCounter, CommentExtension } from './document-parser';
export { extractParagraphs } from './text-extractor';
export {
  parseDocument,
  parseBackgroundColor,
  parseBackgroundFill,
  parseBodyElements,
  parseSection,
  parseSections,
  parseHeaderFooter,
  parseNotes,
  parseComments,
  parseCommentThreads,
  parseCommentsExtended,
  parsePeople,
  applyAuthorIds,
  bodyIndexForBlock,
  newBlockCounter,
  EMPTY_SECTION,
} from './document-parser';
export { parseRunProperties } from './run-properties';
export { HTML_AUTO_SPACING_PT, parseParagraphProperties } from './paragraph-properties';
export { parseStyles } from './styles-parser';
export { parseTable } from './table-parser';
export { loadEmbeddedFonts, deobfuscateEmbeddedFont, parseFontTable } from './font-table';
export { parseNumbering, EMPTY_NUMBERING } from './numbering-parser';
export { parseSettings, EMPTY_SETTINGS } from './settings-parser';
export type { DocumentSettings } from './settings-parser';
