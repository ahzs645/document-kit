/**
 * XML escaping for filled values, taken from officeParser 8.9.0's `src/utils/sanitize.ts` (MIT,
 * Harsh Ankur): the two functions the template engine uses, unchanged.
 *
 * @module template/sanitize
 */

/**
 * Escapes text for an XML text node or attribute (XHTML/OPF/NCX). Same as
 * escapeHtml but emits the XML-canonical `&apos;` for the single quote.
 */
export function escapeXml(text: string): string {
    if (typeof text !== 'string') return '';
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * Removes characters that are illegal in XML 1.0 even when escaped, so a single stray control byte
 * from a hostile source document cannot make an OOXML part unparseable (which makes Word refuse the
 * whole file). Strips C0/C1 controls except tab/newline/carriage-return, U+FFFE/U+FFFF, and lone
 * surrogates. Apply before escaping at every text and attribute sink.
 */
export function stripInvalidXmlChars(text: string): string {
    if (typeof text !== 'string') return '';
    return text
        .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x84\x86-\x9F\uFFFE\uFFFF]/g, '')
        .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])/g, '')   // high surrogate not followed by low
        .replace(/(^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/g, '$1'); // low surrogate not preceded by high
}

