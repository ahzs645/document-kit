/**
 * document-kit/layout — a Word file to PDF, page pictures (SVG) and flowed
 * HTML, laid out in the browser by the reamkit engine (`document-kit/ream`).
 *
 * Nothing here touches the network. Fonts embedded in the document come
 * first, then the stand-ins the caller hands in through `createFontProvider`;
 * a face nobody supplied is reported in `losses`, never fetched.
 *
 * @packageDocumentation
 */

export { docxToHtml, docxToPages, docxToPdf, layoutDocx } from './convert';
export type {
  HtmlResult,
  LaidOutDocx,
  LayoutDocxOptions,
  LayoutLoss,
  LayoutOptions,
  PagedOptions,
  PageLine,
  PageSize,
  PagesResult,
  PagesWriteOptions,
  PdfOptions,
  PdfResult,
  PdfWriteOptions,
} from './convert';
export { STAND_IN_FILES, STAND_IN_NAMES, createFontProvider, standInFor } from './fonts';
export type {
  FaceStyle,
  FontFileBytes,
  FontLoader,
  FontProvider,
  ProvidedFamily,
  StandInFamily,
  StandInFile,
} from './fonts';
export { WrongPasswordError } from '../ream/core/crypto/password-error';
export type { DocumentInfo } from '../ream/core/document-model';
export type { PdfALevel } from '../ream/layout/styled-layout';
