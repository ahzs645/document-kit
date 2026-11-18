export { parseTtf } from './ttf-parser';
export type { ParsedTtf, TtfTableInfo } from './ttf-parser';
export { subsetTtf, glyphClosure } from './ttf-subset';
export { FontRegistry } from './font-registry';
export type { FontVariant, FontBytesByVariant } from './font-registry';
export { pickVariant } from './font-registry';
export { parseGposKerning, parseGsubLigatures, shapeText } from './opentype-layout';
export type { KerningMap, LigatureMap, ShapedRun } from './opentype-layout';
export { hasSubstitutable, lettersForLigature, substituteLetters } from './ligatures';
export { createFontMeasure, shapeGlyphs, sharedGlyphCodes } from './measure';
export type { FontMeasure } from './measure';
export { buildTrueType, editableEmbedding, MAX_BUILT_GLYPHS } from './ttf-build';
export type {
  BuiltFace,
  BuiltGlyph,
  BuiltKernPair,
  BuiltLigature,
  GlyphSeg,
} from './ttf-build';
