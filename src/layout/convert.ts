// Word file → PDF, page pictures (SVG) and flowed HTML, over the reamkit engine.
//
// The engine's own front door (`Ream.convert`) resolves fonts by downloading
// open substitutes from a CDN. This one never does: fonts embedded in the
// document come first, then the stand-ins the caller supplies through a
// FontProvider, and a face nobody supplied is a loss in the report. The calls
// below are the engine's own pipeline (reader → layout → writer), driven
// directly so the font set can be built from the caller's bytes.

import type { FontRegistry as EngineFontRegistry, ParsedTtf } from '../ream/core/font';
import type { DocumentInfo } from '../ream/core/document-model';
import type { FamilyKey, SubstituteKey } from '../ream/core/fonts/remote-fonts';
import type { Loss } from '../ream/core/ir';
import type { FlowDoc } from '../ream/core/ir/flow';
import type { LaidOutDocument, LaidOutPage } from '../ream/layout/page-doc';
import type { PdfALevel, StyledRenderOptions } from '../ream/layout/styled-layout';
import type { FaceStyle, FontProvider, ProvidedFamily, StandInFamily } from './fonts';

import { flowRenderOptions } from '../ream/core/converter/project';
import { decryptPackage, isEncryptedPackage } from '../ream/core/crypto/offcrypto';
import { WrongPasswordError } from '../ream/core/crypto/password-error';
import { FontRegistry } from '../ream/core/font';
import { familyNamesInFlow } from '../ream/core/fonts/families';
import { resolveFamilyKey } from '../ream/core/fonts/remote-fonts';
import {
  companionsInFlow,
  forEachTextInFlow,
  scriptForCodepoint,
  scriptsInFlow,
} from '../ream/core/fonts/scripts';
import { FEATURES } from '../ream/core/ir';
import { writeHtml } from '../ream/html/html-writer';
import { layoutStyledDocument } from '../ream/layout/styled-layout';
import { emitStyledPdf } from '../ream/pdf/styled-page-emitter';
import { PdfDocument } from '../ream/pdf/writer';
import { writeSvg } from '../ream/svg/svg-writer';
import { docxReader } from '../ream/word/docx-reader';
import { STAND_IN_NAMES } from './fonts';
import { circledNumberRegistry, isCircledNumber } from './circledNumbers';
import { TICK_BOX_CHARACTERS, tickBoxRegistry } from './tickBoxes';

/** Something the conversion dropped, degraded or substituted (the engine's `Loss`). */
export type LayoutLoss = Loss;

/** Options every conversion takes. */
export interface LayoutOptions {
  /**
   * The password of an encrypted Word file (ECMA-376 §2.3). An encrypted file
   * without one, or with the wrong one, throws `WrongPasswordError`.
   */
  readonly password?: string;
}

/** Options for the paged conversions, which need fonts. */
export interface PagedOptions extends LayoutOptions {
  /**
   * The stand-in fonts (see `createFontProvider`). Fonts embedded in the
   * document are used first; a face neither supplies is reported in `losses`.
   * Without a provider only embedded fonts are used, and a document with none
   * cannot be laid out.
   */
  readonly fonts?: FontProvider;
}

/** Options for {@link docxToPdf}. */
export interface PdfOptions extends PagedOptions {
  /** Document properties for the PDF, laid over the Word file's own. */
  readonly info?: DocumentInfo;
  /** Write a PDF/A file at this conformance level. */
  readonly pdfA?: PdfALevel;
  /** Write a tagged PDF (reading order and structure for assistive technology). */
  readonly tagged?: boolean;
}

/** A page's size in points (1/72 inch). */
export interface PageSize {
  readonly widthPt: number;
  readonly heightPt: number;
}

/**
 * One laid-out line of a page: its text, where its first character starts (`x`) and its baseline (`y`), in points
 * from the page's top-left corner, rounded to a hundredth. Lines are in layout order: the body, then frames, headers
 * and footers as the engine sets them.
 */
export interface PageLine {
  readonly text: string;
  readonly x: number;
  readonly y: number;
}

/** How {@link LaidOutDocx.toPages} and {@link docxToPages} write the page pictures. */
export interface PagesWriteOptions {
  /**
   * Put the glyph outlines once, in {@link PagesResult.glyphDefs}, and have
   * every page refer to them, instead of each page carrying its own copy. For
   * an application that places every page inline in ONE HTML document: the
   * pages then draw text only where `glyphDefs` is in the same document. Off
   * by default, which keeps each page a standalone picture.
   */
  readonly sharedGlyphs?: boolean;
}

/** The result of {@link docxToPdf}. */
export interface PdfResult {
  /** The PDF file. */
  readonly bytes: Uint8Array;
  /** How many pages it has. */
  readonly pageCount: number;
  /** The physical page dimensions in points, in PDF page order. */
  readonly pageSizes: ReadonlyArray<PageSize>;
  /** What the conversion dropped, degraded or substituted. */
  readonly losses: ReadonlyArray<LayoutLoss>;
}

/** The result of {@link docxToPages}. */
export interface PagesResult {
  /**
   * One standalone SVG document per page, the page exactly as the PDF lays it
   * out. Text is drawn as glyph outlines (`<path>`s), so a page needs no font
   * to show and has no selectable text; each page carries its own glyph
   * definitions, its ids prefixed `dk-p<page>-` so that pages placed inline in
   * one HTML document do not take each other's glyphs. Each page is drawn on a
   * white rectangle with a light grey outline, as the engine's preview draws it.
   */
  readonly pages: ReadonlyArray<string>;
  /** Each page's size, in the order of `pages`. */
  readonly pageSizes: ReadonlyArray<PageSize>;
  /**
   * The text of each page's lines, in layout order, one line per `\n` — for
   * search and checks, since the SVG's text is outlines. Text inside charts and
   * drawn shapes is not included.
   */
  readonly pageText: ReadonlyArray<string>;
  /** Each page's lines with their positions, the same lines as `pageText` (see {@link PageLine}). */
  readonly pageLines: ReadonlyArray<ReadonlyArray<PageLine>>;
  /**
   * With `sharedGlyphs`: one hidden `<svg>` holding every glyph outline the
   * pages draw, each once, to be placed in the same HTML document as the pages
   * (anywhere; it takes no room). Absent otherwise.
   */
  readonly glyphDefs?: string;
  /** What the conversion dropped, degraded or substituted. */
  readonly losses: ReadonlyArray<LayoutLoss>;
}

/** The result of {@link docxToHtml}. */
export interface HtmlResult {
  /** A self-contained HTML document (pictures inlined). */
  readonly html: string;
  /** What the conversion dropped, degraded or substituted. */
  readonly losses: ReadonlyArray<LayoutLoss>;
}

/** Options for {@link layoutDocx}. */
export interface LayoutDocxOptions extends PagedOptions {
  /**
   * Write the PDF as PDF/A at this level. It is a layout option because the
   * accessible levels (`…a`) need the structure tree the layout builds.
   */
  readonly pdfA?: PdfALevel;
  /**
   * Build the structure tree a tagged PDF needs while laying out. It changes
   * nothing on the pages; only {@link LaidOutDocx.toPdf} uses it.
   */
  readonly tagged?: boolean;
}

/** Options for {@link LaidOutDocx.toPdf}. */
export interface PdfWriteOptions {
  /** Document properties for the PDF, laid over the Word file's own. */
  readonly info?: DocumentInfo;
}

/**
 * A Word file laid out once, ready to be written as a PDF, as page pictures or
 * both, without laying it out again.
 */
export interface LaidOutDocx {
  /** How many pages the layout made. */
  readonly pageCount: number;
  /** Each page's size. */
  readonly pageSizes: ReadonlyArray<PageSize>;
  /** The text of each page's lines, one per `\n` (see {@link PagesResult.pageText}). */
  readonly pageText: ReadonlyArray<string>;
  /** Each page's lines with their positions (see {@link PageLine}). */
  readonly pageLines: ReadonlyArray<ReadonlyArray<PageLine>>;
  /** What reading and laying out the file dropped, degraded or substituted. */
  readonly losses: ReadonlyArray<LayoutLoss>;
  /**
   * Write the layout as a PDF. Can be called more than once; each call writes
   * the same pages.
   */
  toPdf(options?: PdfWriteOptions): PdfResult;
  /** Draw each page of the layout as an SVG picture. Can be called more than once. */
  toPages(options?: PagesWriteOptions): PagesResult;
}

/**
 * Lay a Word file out once. The result writes the PDF and the page pictures
 * from that one layout, so an application that previews and then exports pays
 * for the layout once. {@link docxToPdf} and {@link docxToPages} are this
 * followed by one writer.
 *
 * @param docx    The `.docx` bytes.
 * @param options Fonts, password, and the PDF settings the layout depends on.
 * @returns The laid-out document.
 */
export async function layoutDocx(
  docx: Uint8Array,
  options: LayoutDocxOptions = {},
): Promise<LaidOutDocx> {
  const { flow, losses } = readWord(docx, options.password);
  const render = await renderOptions(flow, options.fonts, losses);
  const styled: StyledRenderOptions = {
    ...render,
    // Compressed page drawing and palette pictures: the same pages, a fraction of the bytes.
    compact: true,
    ...(options.pdfA ? { pdfA: options.pdfA } : {}),
    ...(options.tagged ? { tagged: true } : {}),
  };
  const laid = layoutStyledDocument(flow.body, styled);
  const pageSizes = laid.pages.map((p) => ({ widthPt: p.width, heightPt: p.height }));
  const pageLines = laid.pages.map(linesOfPage);
  const pageTexts = pageLines.map(pageText);
  const layoutLosses: ReadonlyArray<LayoutLoss> = [...losses];
  return {
    pageCount: laid.pages.length,
    pageSizes,
    pageText: pageTexts,
    pageLines,
    losses: layoutLosses,
    toPdf(write: PdfWriteOptions = {}): PdfResult {
      const info = flow.info || write.info ? { ...flow.info, ...write.info } : undefined;
      const bytes = emitStyledPdf(laid, { ...styled, ...(info ? { info } : {}) }, new PdfDocument());
      return { bytes, pageCount: laid.pages.length, pageSizes, losses: [...layoutLosses] };
    },
    toPages(write: PagesWriteOptions = {}): PagesResult {
      const out = [...layoutLosses];
      const drawn = drawPages(laid, write.sharedGlyphs === true, out);
      return {
        pages: drawn.pages,
        ...(drawn.glyphDefs !== undefined ? { glyphDefs: drawn.glyphDefs } : {}),
        pageSizes,
        pageText: pageTexts,
        pageLines,
        losses: out,
      };
    },
  };
}

/**
 * Lay a Word file out and write it as a PDF.
 *
 * @param docx    The `.docx` bytes.
 * @param options Fonts, password and PDF settings.
 * @returns The PDF bytes, its page count and the losses.
 */
export async function docxToPdf(docx: Uint8Array, options: PdfOptions = {}): Promise<PdfResult> {
  const laid = await layoutDocx(docx, options);
  return laid.toPdf(options.info ? { info: options.info } : {});
}

/**
 * Lay a Word file out and draw each page as an SVG picture — the same layout
 * the PDF gets, for a true-to-print preview.
 *
 * @param docx    The `.docx` bytes.
 * @param options Fonts, password, and how the pictures are written.
 * @returns One SVG per page, the page sizes, each page's text and the losses.
 */
export async function docxToPages(
  docx: Uint8Array,
  options: PagedOptions & PagesWriteOptions = {},
): Promise<PagesResult> {
  const laid = await layoutDocx(docx, options);
  return laid.toPages(options.sharedGlyphs ? { sharedGlyphs: true } : {});
}

/** Draw every page of a layout as its own SVG document, ids scoped per page. */
function drawPages(
  laid: LaidOutDocument,
  shared: boolean,
  losses: Array<LayoutLoss>,
): { pages: Array<string>; glyphDefs?: string } {
  const decoder = new TextDecoder();
  const pages: Array<string> = [];
  for (const page of laid.pages) {
    const one: LaidOutDocument = { ...laid, pages: [page] };
    const svg = writeSvg(one);
    losses.push(...svg.losses);
    pages.push(scopeIds(decoder.decode(svg.bytes), `dk-p${pages.length + 1}-`));
  }
  if (!shared) return { pages };
  return shareGlyphs(pages);
}

/**
 * Write a Word file as flowed HTML: no pages, no layout and no fonts. Headers,
 * footers and page numbers have no place in a flowed page and are reported as
 * losses.
 *
 * @param docx    The `.docx` bytes.
 * @param options Password.
 * @returns The HTML text and the losses.
 */
export function docxToHtml(docx: Uint8Array, options: LayoutOptions = {}): Promise<HtmlResult> {
  try {
    const { flow, losses } = readWord(docx, options.password);
    const out = writeHtml(flow);
    losses.push(...out.losses);
    return Promise.resolve({ html: new TextDecoder().decode(out.bytes), losses });
  } catch (error) {
    return Promise.reject(error instanceof Error ? error : new Error(String(error)));
  }
}

/** Open (and if need be decrypt) a Word file and read it into the engine's flow model. */
function readWord(
  bytes: Uint8Array,
  password: string | undefined,
): { flow: FlowDoc; losses: Array<LayoutLoss> } {
  let source = bytes;
  if (isEncryptedPackage(bytes)) {
    if (password === undefined) {
      throw new WrongPasswordError('This Word file is password-protected: pass `password`.');
    }
    source = decryptPackage(bytes, password);
  }
  if (!docxReader.sniff(source)) {
    throw new Error('Not a Word file (.docx): no word/document.xml in the package.');
  }
  const { doc, losses } = docxReader.read(source);
  return { flow: doc, losses: [...losses] };
}

const FACE_WORDS: Readonly<Record<FaceStyle, string>> = {
  regular: 'Regular',
  bold: 'Bold',
  italic: 'Italic',
  boldItalic: 'Bold Italic',
};

const DRAWN_FROM: Readonly<Record<Exclude<FaceStyle, 'regular'>, string>> = {
  bold: 'bold text in it is drawn from another of its faces, thickened by the layout',
  italic: 'italic text in it is drawn from another of its faces, slanted by the layout',
  boldItalic: 'bold italic text in it is drawn from another of its faces',
};

/**
 * Build the engine's render options for a document: its own projection plus a
 * font set from the document's embedded fonts and the caller's stand-ins.
 * Records every substitution and every face nobody supplied in `losses`.
 */
async function renderOptions(
  flow: FlowDoc,
  provider: FontProvider | undefined,
  losses: Array<LayoutLoss>,
): Promise<StyledRenderOptions> {
  const embedded = flow.embeddedFonts ?? new Map<string, EngineFontRegistry>();

  // Which stand-in each named family needs. A family the document carries
  // itself is drawn with its own embedded face and needs none. Arimo is always
  // asked for, as the engine does: it draws unstyled runs, math and chart text.
  const namesByFamily = new Map<StandInFamily, Array<string>>([['arimo', []]]);
  for (const name of familyNamesInFlow(flow)) {
    if (embedded.has(name)) continue;
    const key = standInKey(resolveFamilyKey(name));
    const list = namesByFamily.get(key) ?? [];
    list.push(name);
    namesByFamily.set(key, list);
  }

  const provided = new Map<StandInFamily, ProvidedFamily>();
  for (const family of namesByFamily.keys()) {
    provided.set(
      family,
      provider
        ? await provider.family(family)
        : { missing: [{ style: 'regular', reason: 'no font provider was given' }] },
    );
  }

  // The engine's per-family registries, in the order it falls back through
  // them (the first is the fallback for a family it has no registry for).
  const registriesByFamily = new Map<SubstituteKey, EngineFontRegistry>();
  for (const [family, got] of provided) {
    if (got.faces) registriesByFamily.set(family, FontRegistry.fromBytes(got.faces));
  }
  const fallbackFamily = registriesByFamily.keys().next().value as StandInFamily | undefined;
  const registry =
    (fallbackFamily ? registriesByFamily.get(fallbackFamily) : undefined) ??
    embedded.values().next().value;
  if (!registry) {
    throw new Error(
      'No fonts to lay the document out with: it embeds none and no stand-in was supplied ' +
        '(pass `fonts: createFontProvider(…)`).',
    );
  }
  const fallbackName = fallbackFamily
    ? STAND_IN_NAMES[fallbackFamily]
    : `the embedded font "${embedded.keys().next().value ?? ''}"`;

  for (const [family, names] of namesByFamily) {
    const got = provided.get(family)!;
    const standIn = STAND_IN_NAMES[family];
    if (!got.faces) {
      const why = got.missing.find((m) => m.style === 'regular')?.reason ?? 'not supplied';
      const subjects =
        names.length > 0 ? names.map((n) => `"${n}"`).join(', ') : 'text with no named font';
      losses.push({
        severity: 'substituted',
        feature: FEATURES.fontsSubstitution,
        detail: `${subjects} drawn with ${fallbackName}: ${standIn} Regular was not available (${why})`,
      });
      continue;
    }
    for (const name of names) {
      losses.push({
        severity: 'substituted',
        feature: FEATURES.fontsSubstitution,
        detail: `"${name}" drawn with its stand-in ${standIn}`,
      });
    }
    for (const m of got.missing) {
      if (m.style === 'regular') continue;
      losses.push({
        severity: 'degraded',
        feature: FEATURES.fontsSubstitution,
        detail: `${standIn} ${FACE_WORDS[m.style]} was not available (${m.reason}); ${DRAWN_FROM[m.style]}`,
      });
    }
  }

  // Letters the bundled families do not have. The engine would download a face
  // for each; the kit names what is missing instead.
  const loaded = [...registriesByFamily.keys()] as Array<FamilyKey>;
  for (const companion of companionsInFlow(flow, loaded)) {
    losses.push({
      severity: 'degraded',
      feature: FEATURES.fontsSubstitution,
      detail:
        `Greek or Cyrillic letters in Caladea (Cambria) text want ${COMPANION_NAMES[companion] ?? companion}, ` +
        'which document-kit does not bundle; they are drawn from another face on hand, at its own widths',
    });
  }
  // Install the built-in symbol faces only when a document needs their glyphs.
  const before = [...registriesByFamily.values(), ...embedded.values()].flatMap(facesOf);
  const boxes = new Set<number>();
  let hasCircledNumbers = false;
  forEachTextInFlow(flow, (text) => {
    for (const ch of text) {
      const cp = ch.codePointAt(0) ?? 0;
      if (TICK_BOX_CHARACTERS.includes(cp)) boxes.add(cp);
      if (isCircledNumber(cp)) hasCircledNumbers = true;
    }
  });
  const undrawn = (cp: number) => !before.some((face) => face.glyphForCodepoint(cp) !== 0);
  if ([...boxes].some(undrawn)) registriesByFamily.set('symbols2', tickBoxRegistry());
  if (hasCircledNumbers) registriesByFamily.set('symbols1', circledNumberRegistry());


  // A script is named only for the characters of it that no face on hand
  // draws: the engine sends ▼ ■ ↑ to a symbol face it would download, but the
  // stand-ins draw most of them themselves, and a loss for those misled.
  const faces = [...registriesByFamily.values(), ...embedded.values()].flatMap(facesOf);
  const drawn = (cp: number): boolean => faces.some((face) => face.glyphForCodepoint(cp) !== 0);
  for (const [script, cps] of scriptCharacters(flow)) {
    const missing = [...cps].filter((cp) => !drawn(cp));
    if (missing.length === 0) continue;
    const shown = missing
      .slice(0, 8)
      .map((cp) => `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`)
      .join(', ');
    losses.push({
      severity: 'degraded',
      feature: FEATURES.fontsSubstitution,
      detail:
        `the document has ${script} text that no supplied face draws ` +
        `(${shown}${missing.length > 8 ? ', …' : ''}); document-kit bundles no face for that ` +
        'script, and those letters are drawn as empty boxes',
    });
  }

  return {
    ...flowRenderOptions(flow),
    registry,
    ...(registriesByFamily.size > 0 ? { registriesByFamily } : {}),
    losses,
  };
}

/** Every face of a registry (the regular one and those of the styles it has). */
function facesOf(registry: EngineFontRegistry): Array<ParsedTtf> {
  const out = new Set<ParsedTtf>();
  for (const [bold, italic] of [
    [false, false],
    [true, false],
    [false, true],
    [true, true],
  ] as const) {
    out.add(registry.resolveByStyle(bold, italic).parsed);
  }
  return [...out];
}

/**
 * The characters of each script the engine has no bundled face for, by the
 * script's name in loss messages. Han is named for the face the engine would
 * draw it with (Japanese, Korean or Chinese, by what else the document holds).
 */
function scriptCharacters(flow: FlowDoc): Map<string, Set<number>> {
  const { hanFace } = scriptsInFlow(flow);
  const out = new Map<string, Set<number>>();
  forEachTextInFlow(flow, (text) => {
    for (const ch of text) {
      const cp = ch.codePointAt(0) ?? 0;
      if (cp < 0x0590) continue;
      const key = scriptForCodepoint(cp, hanFace);
      if (key === undefined) continue;
      const name = SCRIPT_NAMES[key] ?? key;
      let set = out.get(name);
      if (!set) out.set(name, (set = new Set()));
      set.add(cp);
    }
  });
  return out;
}

/** Names for the engine's companion keys, for loss messages. */
const COMPANION_NAMES: Readonly<Partial<Record<FamilyKey, string>>> = { notoserif: 'Noto Serif' };

/** Names for the engine's script keys (`ScriptKey`), for loss messages. */
const SCRIPT_NAMES: Readonly<Record<string, string>> = {
  jp: 'Japanese',
  kr: 'Korean',
  sc: 'Chinese',
  arabic: 'Arabic',
  hebrew: 'Hebrew',
  thai: 'Thai',
  devanagari: 'Devanagari',
  bengali: 'Bengali',
  gurmukhi: 'Gurmukhi',
  gujarati: 'Gujarati',
  oriya: 'Oriya',
  tamil: 'Tamil',
  telugu: 'Telugu',
  kannada: 'Kannada',
  malayalam: 'Malayalam',
  symbols1: 'symbol',
  symbols2: 'symbol',
};

/** The engine names `notoserif` only as a companion, never for a family name. */
function standInKey(key: FamilyKey): StandInFamily {
  return key === 'notoserif' ? 'tinos' : key;
}

/**
 * Prefix every id an SVG defines, and every reference to one, so several pages
 * can sit inline in one HTML document: each page is written on its own and
 * numbers its glyphs and clips from zero, and inline SVGs share one id space.
 */
function scopeIds(svg: string, prefix: string): string {
  const ids = new Set<string>();
  for (const m of svg.matchAll(/\sid="([^"]+)"/g)) ids.add(m[1]!);
  if (ids.size === 0) return svg;
  return svg
    .replace(/(\sid=")([^"]+)"/g, (_m, pre: string, id: string) => `${pre}${prefix}${id}"`)
    .replace(/(href="#|url\(#)([^")]+)/g, (m, pre: string, id: string) =>
      ids.has(id) ? `${pre}${prefix}${id}` : m,
    );
}

/** The id prefix of a glyph outline the shared definitions hold. */
const SHARED_GLYPH_PREFIX = 'dk-g';

/**
 * Move the glyph outlines out of every page into one set of definitions. Each
 * page numbers its glyphs on its own (`dk-p3-g0-36`), so an outline is known
 * by its path data, which is the glyph's shape in em units whatever the size
 * it is drawn at: the same shape on two pages becomes one `dk-gN`. A page's
 * other definitions (clips, gradients, pictures) stay on the page.
 */
function shareGlyphs(pages: ReadonlyArray<string>): { pages: Array<string>; glyphDefs: string } {
  const byPath = new Map<string, string>();
  const defs: Array<string> = [];
  const glyphPath = /<path id="(dk-p\d+-g\d+-\d+)" d="([^"]*)"\/>/g;
  const out = pages.map((page) => {
    const renamed = new Map<string, string>();
    const stripped = page.replace(glyphPath, (_whole, id: string, d: string) => {
      let shared = byPath.get(d);
      if (shared === undefined) {
        shared = `${SHARED_GLYPH_PREFIX}${byPath.size + 1}`;
        byPath.set(d, shared);
        defs.push(`<path id="${shared}" d="${d}"/>`);
      }
      renamed.set(id, shared);
      return '';
    });
    if (renamed.size === 0) return page;
    return stripped
      .replace(/<defs><\/defs>/g, '')
      .replace(/href="#(dk-p\d+-g\d+-\d+)"/g, (whole, id: string) => {
        const shared = renamed.get(id);
        return shared === undefined ? whole : `href="#${shared}"`;
      });
  });
  const glyphDefs =
    '<svg xmlns="http://www.w3.org/2000/svg" width="0" height="0" aria-hidden="true" ' +
    `style="position:absolute;width:0;height:0;overflow:hidden"><defs>${defs.join('')}</defs></svg>`;
  return { pages: out, glyphDefs };
}

/**
 * A page's laid-out lines: each line's text, where it starts and its baseline. A no-break hyphen (Word's
 * `w:noBreakHyphen`, which keeps "Phone: 250-555-0100" from breaking) is read as the hyphen it prints: a person
 * copying a phone number out of the page gets the hyphen they would type.
 */
function linesOfPage(page: LaidOutPage): Array<PageLine> {
  const lines: Array<PageLine> = [];
  for (const item of page.commands) {
    if (item.type !== 'line') continue;
    let text = '';
    for (const token of item.line.tokens) if (token.kind === 'text') text += token.text;
    lines.push({ text: text.replace(/\u2011/g, '-'), x: round2(item.originX), y: round2(item.baselineY) });
  }
  return lines;
}

const round2 = (value: number): number => Math.round(value * 100) / 100;

/** The text of a laid-out page's lines, one per line. */
function pageText(lines: ReadonlyArray<PageLine>): string {
  return lines.map((line) => line.text).join('\n');
}
