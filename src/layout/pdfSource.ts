import type { PdfResult } from './convert';

/**
 * Keep one in-flight or finished PDF for an immutable document snapshot and output variant.
 * Preview and download can share the result even when they request it at the same time.
 * A failed build can be retried, and cached files follow the lifetime of the snapshot.
 *
 * The generator can call the layout API directly or return a PDF produced by a worker.
 */
export function createPdfSource<Snapshot extends object, Variant, Pdf extends Pick<PdfResult, 'bytes'>>(
  generate: (snapshot: Snapshot, variant: Variant) => Pdf | Promise<Pdf>,
): (snapshot: Snapshot, variant: Variant) => Promise<Pdf> {
  const snapshots = new WeakMap<Snapshot, Map<Variant, Promise<Pdf>>>();
  return (snapshot, variant) => {
    const variants = snapshots.get(snapshot) ?? new Map<Variant, Promise<Pdf>>();
    const existing = variants.get(variant);
    if (existing) return existing;

    const pending = Promise.resolve().then(() => generate(snapshot, variant));
    variants.set(variant, pending);
    snapshots.set(snapshot, variants);
    void pending.catch(() => variants.delete(variant));
    return pending;
  };
}

/**
 * Open the generated PDF in a caller-supplied viewer such as PDF.js.
 * A viewer can transfer its copy to a worker while the original stays available for download.
 */
export function openPdfPreview<Task>(
  pdf: Pick<PdfResult, 'bytes'>,
  open: (data: Uint8Array) => Task,
): Task {
  return open(pdf.bytes.slice());
}
