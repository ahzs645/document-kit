/**
 * document-kit/html: writes the block model as one self-contained, print-styled HTML page, from the same theme and
 * named styles the Word writer uses. Browser-safe: no Node, no DOM, no network.
 */
export { writeHtml, type WriteHtmlOptions } from './write'
export { fontEmbeddingReport, type FontEmbedding, type NamedFont } from '../model/index'
