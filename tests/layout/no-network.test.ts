// The kit's hard rule: a conversion never touches the network. Every entry is
// run with `fetch` replaced by a spy that fails the test if it is called, on
// documents that would send the engine's own font path to its CDN: a family
// nobody supplied, Cyrillic in Cambria (its companion face) and Japanese text.

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createFontProvider, docxToHtml, docxToPages, docxToPdf } from '../../src/layout';
import { Ream } from '../../src/ream';
import { clearFontCache } from '../../src/ream/core/fonts';
import { diskLoader, fontFile, sampleDocx } from './sample';

const fetchSpy = vi.fn(() => {
  throw new Error('fetch was called during a conversion');
});

beforeEach(() => {
  clearFontCache();
  fetchSpy.mockClear();
  vi.stubGlobal('fetch', fetchSpy);
});
afterEach(() => {
  vi.unstubAllGlobals();
});

const hungry = (): Promise<Uint8Array> =>
  sampleDocx({
    font: 'Calibri',
    rows: 5,
    extra: [
      { text: 'Привет, мир', font: 'Cambria' },
      { text: '日本語のテキスト' },
      { text: 'Lucida text', font: 'Lucida Console' },
    ],
  });

describe('no network', () => {
  it('docxToPdf with every stand-in supplied', async () => {
    const { losses } = await docxToPdf(await hungry(), { fonts: createFontProvider(diskLoader) });
    expect(fetchSpy).not.toHaveBeenCalled();
    // What the engine would have downloaded is named instead.
    expect(losses.some((l) => l.detail.includes('Noto Serif'))).toBe(true);
    expect(losses.some((l) => l.detail.includes('Japanese text'))).toBe(true);
  });

  it('docxToPdf with only Arimo supplied', async () => {
    const onlyArimo = createFontProvider((f) => (f.family === 'arimo' ? fontFile(f.file) : null));
    const { losses } = await docxToPdf(await hungry(), { fonts: onlyArimo });
    expect(fetchSpy).not.toHaveBeenCalled();
    expect(losses.some((l) => l.detail.includes('Carlito Regular was not available'))).toBe(true);
    expect(losses.some((l) => l.detail.startsWith('"cambria" drawn with Arimo'))).toBe(true);
    expect(losses.some((l) => l.detail.includes('Japanese text'))).toBe(true);
  });

  it('docxToPages', async () => {
    await docxToPages(await hungry(), { fonts: createFontProvider(diskLoader) });
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it('docxToHtml', async () => {
    await docxToHtml(await hungry());
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it('the engine itself refuses to download when no fetch is injected', async () => {
    // Upstream fell back to the global fetch here; this copy never does.
    const error = await Ream.parse(await hungry())
      .convertWithReport('pdf')
      .then(
        () => undefined,
        (e: unknown) => e as Error & { cause?: Error },
      );
    expect(error?.cause?.message).toMatch(/does not download fonts/);
    expect(fetchSpy).not.toHaveBeenCalled();
  });
});
