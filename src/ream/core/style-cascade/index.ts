export {
  EMPTY_STYLE_SHEET,
  mergeRunPartial,
  resolveBodyStyles,
  resolveHeadersFootersStyles,
  resolveParagraphProperties,
  resolveRunProperties,
} from './resolver';
export type {
  ResolvedRunProperties,
  ResolvedParagraphProperties,
} from './types';
export { DEFAULT_RESOLVED_RUN, DEFAULT_RESOLVED_PARAGRAPH } from './types';
export { resolveTableStyles } from './table';
