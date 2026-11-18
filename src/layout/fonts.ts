// Fonts for the layout API: bytes the caller hands in, never a download.
//
// A Word file names its faces (Calibri, Times New Roman, Arial…), most of
// which cannot be shipped. The engine already knows which open face stands in
// for each name — the LibreOffice set of metric-compatible twins, then a
// serif / sans / mono class fallback — and that table is the one used here
// (`resolveFamilyKey`, src/ream/core/fonts/remote-fonts.ts). Upstream it feeds a
// CDN downloader; here it only names a file in `fonts/` that the caller loads
// however it likes (an app fetches its own bundled copy, a test reads the disk).

import type { FontBytesByVariant } from '../ream/core/font';
import { resolveFamilyKey } from '../ream/core/fonts/remote-fonts';

/** The open families document-kit bundles in `fonts/`, by the engine's key. */
export type StandInFamily = 'tinos' | 'arimo' | 'carlito' | 'caladea' | 'cousine';

/** One of a family's four faces. */
export type FaceStyle = 'regular' | 'bold' | 'italic' | 'boldItalic';

/** A bundled font file: which family and face it is, and its name in `fonts/`. */
export interface StandInFile {
  readonly family: StandInFamily;
  readonly style: FaceStyle;
  /** File name in the package's `fonts/` folder, e.g. `Carlito-Bold.ttf`. */
  readonly file: string;
}

/** The family name each stand-in is published under, for messages and file names. */
export const STAND_IN_NAMES: Readonly<Record<StandInFamily, string>> = {
  tinos: 'Tinos',
  arimo: 'Arimo',
  carlito: 'Carlito',
  caladea: 'Caladea',
  cousine: 'Cousine',
};

const STYLE_SUFFIX: Readonly<Record<FaceStyle, string>> = {
  regular: 'Regular',
  bold: 'Bold',
  italic: 'Italic',
  boldItalic: 'BoldItalic',
};

const FAMILIES = Object.keys(STAND_IN_NAMES) as ReadonlyArray<StandInFamily>;
const STYLES = Object.keys(STYLE_SUFFIX) as ReadonlyArray<FaceStyle>;

/** Every bundled font file: five families, four faces each. */
export const STAND_IN_FILES: ReadonlyArray<StandInFile> = FAMILIES.flatMap((family) =>
  STYLES.map((style) => ({
    family,
    style,
    file: `${STAND_IN_NAMES[family]}-${STYLE_SUFFIX[style]}.ttf`,
  })),
);

/**
 * The bundled family that stands in for a font name a document uses, by the
 * engine's own substitution table: Calibri → Carlito, Cambria → Caladea,
 * Arial/Helvetica → Arimo, Times New Roman → Tinos, Courier New → Cousine,
 * other serif names → Tinos, other monospaced names → Cousine, anything else
 * → Arimo.
 *
 * @param fontName The family name as the document spells it.
 * @returns The stand-in family's key.
 */
export function standInFor(fontName: string): StandInFamily {
  const key = resolveFamilyKey(fontName);
  // The engine maps a NAME only to one of the five; `notoserif` is reached
  // only as Caladea's companion for Greek and Cyrillic, never from a name.
  return key === 'notoserif' ? 'tinos' : key;
}

/** What a font loader may return: the file's bytes, or nothing when it has none. */
export type FontFileBytes = Uint8Array | ArrayBuffer;

/**
 * Loads one bundled font file. Return `undefined` (or `null`) for a file you do
 * not supply; the conversion then reports the face as a loss. A loader that
 * throws is treated the same way, with its message kept in the loss.
 */
export type FontLoader = (
  file: StandInFile,
) => FontFileBytes | null | undefined | Promise<FontFileBytes | null | undefined>;

/** The faces of one family a provider could supply, and what it could not. */
export interface ProvidedFamily {
  /** The faces, ready for the engine; absent when the regular face is missing. */
  readonly faces?: FontBytesByVariant;
  /** Faces that were not supplied, each with the reason. */
  readonly missing: ReadonlyArray<{ readonly style: FaceStyle; readonly reason: string }>;
}

/**
 * Hands the layout the bytes of the bundled stand-in families. Build one with
 * {@link createFontProvider}; it loads a family only when a document needs it,
 * and each family once.
 */
export interface FontProvider {
  /** Marks the object as a document-kit font provider. */
  readonly kind: 'document-kit/font-provider';
  /** The faces of one stand-in family (loaded on first use, then kept). */
  readonly family: (family: StandInFamily) => Promise<ProvidedFamily>;
}

function asBytes(value: FontFileBytes): Uint8Array {
  return value instanceof Uint8Array ? value : new Uint8Array(value);
}

/**
 * Build a {@link FontProvider} from a loader. The loader is asked for files by
 * name (see {@link STAND_IN_FILES}); it is never asked for anything the
 * document does not need, and nothing here touches the network.
 *
 * ```ts
 * const fonts = createFontProvider(async ({ file }) =>
 *   new Uint8Array(await (await fetch(`/assets/fonts/${file}`)).arrayBuffer()),
 * );
 * const { bytes, losses } = await docxToPdf(docx, { fonts });
 * ```
 *
 * @param load Returns the bytes of one font file, or nothing.
 * @returns The provider.
 */
export function createFontProvider(load: FontLoader): FontProvider {
  const cache = new Map<StandInFamily, Promise<ProvidedFamily>>();
  const loadFamily = async (family: StandInFamily): Promise<ProvidedFamily> => {
    const files = STAND_IN_FILES.filter((f) => f.family === family);
    const results = await Promise.all(
      files.map(async (file) => {
        try {
          const bytes = await load(file);
          return bytes ? { file, bytes: asBytes(bytes) } : { file, reason: 'not supplied' };
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error);
          return { file, reason: `the loader failed: ${message}` };
        }
      }),
    );
    const got: Partial<Record<FaceStyle, Uint8Array>> = {};
    const missing: Array<{ style: FaceStyle; reason: string }> = [];
    for (const r of results) {
      if ('bytes' in r && r.bytes) got[r.file.style] = r.bytes;
      else missing.push({ style: r.file.style, reason: 'reason' in r ? r.reason : 'not supplied' });
    }
    if (!got.regular) return { missing };
    return {
      faces: {
        regular: got.regular,
        ...(got.bold ? { bold: got.bold } : {}),
        ...(got.italic ? { italic: got.italic } : {}),
        ...(got.boldItalic ? { boldItalic: got.boldItalic } : {}),
      },
      missing,
    };
  };
  return {
    kind: 'document-kit/font-provider',
    family: (family) => {
      let pending = cache.get(family);
      if (!pending) {
        const loading = loadFamily(family);
        pending = loading;
        cache.set(family, loading);
        // A face the loader failed on (a dropped request, say) is asked for
        // again next time; one it answered with nothing is not.
        void loading.then((result) => {
          const failed = result.missing.some((m) => m.reason.startsWith('the loader failed'));
          if (failed && cache.get(family) === loading) cache.delete(family);
        });
      }
      return pending;
    },
  };
}
