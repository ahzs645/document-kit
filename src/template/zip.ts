/**
 * Unzipping a template under decompression limits, taken from officeParser 8.9.0's
 * `src/utils/zipUtils.ts` (`extractFiles`, MIT, Harsh Ankur). Changed for the browser: bytes are
 * `Uint8Array`s, not Node `Buffer`s, and the abort-signal path (event-loop turns) is gone, because a
 * template call takes no abort signal.
 *
 * @module template/zip
 */

import { Unzip, UnzipInflate, type UnzipFile } from 'fflate';
import { DecompressionLimits, IssueHandler, OfficeErrorType } from './types.js';
import { getOfficeError } from './errors.js';

/**
 * Signature of the End Of Central Directory record ("PK\x05\x06"), the trailer every ZIP
 * archive ends with. Its presence is what distinguishes a complete archive from one that
 * was cut off in transfer.
 */
const EOCD_SIGNATURE = [0x50, 0x4b, 0x05, 0x06] as const;

/** Size of the fixed portion of an End Of Central Directory record, in bytes. */
const EOCD_RECORD_MIN_BYTES = 22;

/** Maximum size of the optional archive comment that may trail the EOCD record, in bytes. */
const ZIP_MAX_COMMENT_BYTES = 65535;

/**
 * How far back from the end of the input the EOCD record may start. The record is last in
 * the file apart from its own variable-length comment, so searching this window is
 * sufficient and bounded regardless of archive size.
 */
const EOCD_SEARCH_WINDOW_BYTES = EOCD_RECORD_MIN_BYTES + ZIP_MAX_COMMENT_BYTES;

/** A file extracted from a ZIP archive: its path in the archive and its bytes. */
export interface ZipFileContent {
    path: string;
    content: Uint8Array;
}

/** The bytes one archive may inflate to: `limits.maxUncompressedBytes` when it is a valid limit, else 512 MB. */
export const maxUncompressedBytesOf = (limits: DecompressionLimits | undefined): number =>
    limits?.maxUncompressedBytes !== undefined && Number.isFinite(limits.maxUncompressedBytes) && limits.maxUncompressedBytes >= 0
        ? limits.maxUncompressedBytes
        : 512 * 1024 * 1024;

/** The last index of the EOCD signature in `bytes`, or -1 (Buffer.lastIndexOf, for a Uint8Array). */
function lastIndexOfEocd(bytes: Uint8Array): number {
    for (let i = bytes.length - EOCD_SIGNATURE.length; i >= 0; i--) {
        if (bytes[i] === EOCD_SIGNATURE[0] && bytes[i + 1] === EOCD_SIGNATURE[1] && bytes[i + 2] === EOCD_SIGNATURE[2] && bytes[i + 3] === EOCD_SIGNATURE[3]) return i;
    }
    return -1;
}

/** The chunks of one entry joined (Buffer.concat, for Uint8Arrays). */
function concat(chunks: Uint8Array[]): Uint8Array {
    if (chunks.length === 1) return chunks[0]!;
    let length = 0;
    for (const chunk of chunks) length += chunk.length;
    const out = new Uint8Array(length);
    let offset = 0;
    for (const chunk of chunks) { out.set(chunk, offset); offset += chunk.length; }
    return out;
}

/**
 * Extracts the files `filterFn` accepts from a ZIP archive, counting the bytes actually inflated
 * against `maxUncompressedBytes` (the size a header declares is the attacker's to choose) and the
 * entries against `maxZipEntries`. Input that is no archive rejects with `ZIP_NO_ENTRIES_FOUND`, and an
 * archive cut short with `ZIP_TRUNCATED`.
 */
export const extractFiles = (
    zipInput: Uint8Array,
    filterFn: (fileName: string) => boolean,
    limits: DecompressionLimits,
    config?: IssueHandler
): Promise<ZipFileContent[]> => {
    const maxUncompressedBytes = maxUncompressedBytesOf(limits);

    const maxZipEntries = limits?.maxZipEntries !== undefined && Number.isFinite(limits.maxZipEntries) && limits.maxZipEntries >= 0
        ? limits.maxZipEntries
        : 10000;

    return new Promise((resolve, reject) => {
        // Decompress as a stream and cap on the ACTUAL inflated byte count rather than
        // the size declared in the ZIP header. The declared size is attacker-controlled,
        // so a "zip bomb" can understate it and still inflate to gigabytes; counting real
        // output bytes and aborting once the limit is crossed is the only reliable guard.
        const results: ZipFileContent[] = [];
        let totalEntryCount = 0;
        let actualTotalBytes = 0;
        let pendingFiles = 0;
        let pushComplete = false;
        let settled = false;
        const activeFiles = new Map<UnzipFile, Uint8Array[]>();

        const cleanup = () => {
            for (const [file, chunks] of activeFiles) {
                chunks.length = 0;
                file.terminate();
            }
            activeFiles.clear();
        };

        // (A cleanup that threw between `settled` and the settle left the promise pending for good.)
        const release = () => { try { cleanup(); } catch { /* the promise settles regardless */ } };
        const fail = (err: unknown) => {
            if (settled) return;
            settled = true;
            release();
            results.length = 0;
            reject(err);
        };
        const maybeResolve = () => {
            if (!settled && pushComplete && pendingFiles === 0) {
                settled = true;
                release();
                resolve(results);
            }
        };

        const unzipper = new Unzip((file) => {
            if (settled) return;
            totalEntryCount++;
            if (totalEntryCount > maxZipEntries) {
                fail(getOfficeError(OfficeErrorType.ZIP_ENTRY_COUNT_LIMIT_EXCEEDED, config, maxZipEntries));
                return;
            }
            if (!filterFn(file.name) || settled) return;

            const name = file.name;
            const chunks: Uint8Array[] = [];
            activeFiles.set(file, chunks);
            pendingFiles++;
            file.ondata = (err, chunk, final) => {
                if (settled) return;
                if (err) { fail(err); return; }
                if (chunk && chunk.length) {
                    actualTotalBytes += chunk.length;
                    if (actualTotalBytes > maxUncompressedBytes) {
                        fail(getOfficeError(OfficeErrorType.ZIP_SIZE_LIMIT_EXCEEDED, config, maxUncompressedBytes));
                        return;
                    }
                    // A copy, as Buffer.from(chunk) made one: fflate may reuse its output buffer.
                    chunks.push(chunk.slice());
                }
                if (final) {
                    results.push({ path: name, content: concat(chunks) });
                    chunks.length = 0;
                    activeFiles.delete(file);
                    pendingFiles--;
                    maybeResolve();
                }
            };
            try {
                file.start();
            } catch (e) {
                fail(e);
            }
        });
        unzipper.register(UnzipInflate);

        const src = zipInput;
        // Feed the compressed input in bounded chunks so the actual-size check can abort a
        // bomb early; the worst-case overshoot is one chunk's worth of inflation.
        const PUSH_CHUNK = 1 << 16; // 64 KiB
        let offset = 0;
        try {
            while (offset < src.length && !settled) {
                const end = Math.min(offset + PUSH_CHUNK, src.length);
                unzipper.push(src.subarray(offset, end), end >= src.length);
                offset = end;
            }
        } catch (e) {
            fail(e);
            return;
        }
        if (settled) return;
        pushComplete = true;
        // fflate's streaming Unzip emits nothing, and no error, when the input is not a
        // ZIP archive. Zero entries can never be a valid document, so treat it as corrupt input.
        // Checked before the truncation check below because it produces the better message for
        // input that is not an archive at all.
        if (totalEntryCount === 0) {
            fail(getOfficeError(OfficeErrorType.ZIP_NO_ENTRIES_FOUND, config));
            return;
        }
        // The streaming reader recovers entries from local file headers alone, so an archive
        // cut short still yields whatever entries preceded the cut. Requiring the trailer that
        // terminates every complete archive rejects such input. Deliberately not gated on
        // pendingFiles: when a cut lands inside an entry's compressed data that entry's final
        // callback never fires, so this is also what settles the promise.
        const tail = zipInput.subarray(Math.max(0, zipInput.length - EOCD_SEARCH_WINDOW_BYTES));
        const eocdIndex = lastIndexOfEocd(tail);
        if (eocdIndex === -1 || eocdIndex + EOCD_RECORD_MIN_BYTES > tail.length) {
            fail(getOfficeError(OfficeErrorType.ZIP_TRUNCATED, config));
            return;
        }
        maybeResolve();
    });
};
