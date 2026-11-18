/**
 * UTF-8 text from and to bytes, in place of officeParser's `Buffer.from(bytes).toString('utf8')` and
 * `Buffer.from(text, 'utf8')`. document-kit's own. The decoder keeps a byte-order mark, as Node's
 * Buffer does (a `TextDecoder` drops it by default), and writes U+FFFD for bytes that are not UTF-8.
 *
 * @module template/utf8
 */

const decoder = new TextDecoder('utf-8', { ignoreBOM: true });
const encoder = new TextEncoder();

/** The bytes read as UTF-8 text. */
export function utf8Text(bytes: Uint8Array): string {
    return decoder.decode(bytes);
}

/** The text written as UTF-8 bytes. */
export function utf8Bytes(text: string): Uint8Array {
    return encoder.encode(text);
}
