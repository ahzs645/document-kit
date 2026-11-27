export type * from './types'
export { DEFAULT_MARGINS, DEFAULT_THEME, PAPER, resolveTheme } from './theme'
export { pageBox, type PageBox } from './page'
export {
  CHECKED_BOX,
  EMPTY_BOX,
  fillSlots,
  inlines,
  isEmptyText,
  isRichText,
  placeholderSentence,
  plainText,
  resolveSentences,
  sentenceText,
  type FilledSentence,
} from './inline'
export { tableColumns, tableGrid, toCell, type PlacedCell, type TableGrid } from './table'
export { fitImage, imageInfo, POINTS_PER_PIXEL, resolveImage, type ImageInfo } from './image'
export { walkBlocks, walkDocument, type BlockVisit } from './walk'
export {
  defineCustomBlocks,
  expandBlocks,
  expandDocument,
  hasCustomBlocks,
  MAX_CUSTOM_DEPTH,
  UnknownCustomBlockError,
  type CustomExpander,
  type CustomRegistry,
  type ExpandContext,
} from './custom'
export {
  captionKind,
  captionLabel,
  captionPrefix,
  numberCaptions,
  takesNumber,
  type CaptionedBlock,
  type CaptionKind,
} from './captions'
export { toComplete, type ToCompleteItem, type ToCompleteKind, type ToCompleteOptions } from './complete'
export {
  applyTextRules,
  compileTextRules,
  linkHref,
  ruleText,
  TEXT_RULE_LIMITS,
  TextRuleError,
  type AppliedTextRules,
  type CompiledTextRules,
  type SkippedText,
  type TextRuleLimits,
  type TextRuleProblem,
} from './textRules'
export * from './builders'
export {
  characterStyleId,
  paragraphStyleId,
  styleTable,
  StyleError,
  THEME_STYLE_IDS,
  type ResolvedCharacterStyle,
  type ResolvedParagraphStyle,
  type StyleTable,
} from './styles'
