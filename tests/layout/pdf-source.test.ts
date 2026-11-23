import { describe, expect, it, vi } from 'vitest';
import { createPdfSource, openPdfPreview } from '../../src/layout';

describe('a shared PDF source', () => {
  it('deduplicates generation for the same snapshot and variant', async () => {
    const snapshot = { title: 'Example' };
    const generate = vi.fn(async (_snapshot: typeof snapshot, _variant: string) => ({ bytes: new Uint8Array([1, 2]) }));
    const getPdf = createPdfSource(generate);
    const preview = getPdf(snapshot, 'a');
    expect(getPdf(snapshot, 'a')).toBe(preview);
    const pdf = await preview;
    expect(await getPdf(snapshot, 'a')).toBe(pdf);
    expect(generate).toHaveBeenCalledTimes(1);

    expect(await getPdf(snapshot, 'b')).not.toBe(pdf);
    expect(await getPdf({ ...snapshot }, 'a')).not.toBe(pdf);
    expect(generate).toHaveBeenCalledTimes(3);
  });

  it('allows a failed generation to be retried', async () => {
    const snapshot = {};
    const generate = vi.fn(() => ({ bytes: new Uint8Array([3]) })).mockImplementationOnce(() => {
      throw new Error('Unable to generate PDF');
    });
    const getPdf = createPdfSource(generate);
    await expect(getPdf(snapshot, undefined)).rejects.toThrow('Unable to generate PDF');
    await expect(getPdf(snapshot, undefined)).resolves.toEqual({ bytes: new Uint8Array([3]) });
    expect(generate).toHaveBeenCalledTimes(2);
  });

  it('keeps the download bytes when the viewer transfers its copy', () => {
    const pdf = { bytes: Uint8Array.from([1, 2, 3]) };
    const viewer = openPdfPreview(pdf, (data) => {
      const moved = structuredClone(data, { transfer: [data.buffer] });
      expect(data.byteLength).toBe(0);
      return moved;
    });
    expect(viewer).toEqual(pdf.bytes);
    expect(pdf.bytes.byteLength).toBe(3);
  });
});
