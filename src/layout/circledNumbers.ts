/// <reference path="./font-assets.d.ts" />
// Small, bundled subsets of Noto Sans Symbols. See fonts/README.md for the source and license.
import bold from '../../fonts/DocumentKitCircledNumbers-Bold.ttf?inline';
import regular from '../../fonts/DocumentKitCircledNumbers-Regular.ttf?inline';
import { FontRegistry } from '../ream/core/font';

export function isCircledNumber(cp: number): boolean {
  return cp === 0x24ea || (cp >= 0x2460 && cp <= 0x2473);
}

function fontBytes(url: string): Uint8Array {
  return Uint8Array.from(atob(url.slice(url.indexOf(',') + 1)), (ch) => ch.charCodeAt(0));
}

let built: FontRegistry | undefined;

/** Decode and load the built-in faces only when a document needs a circled number. */
export function circledNumberRegistry(): FontRegistry {
  built ??= FontRegistry.fromBytes({ regular: fontBytes(regular), bold: fontBytes(bold) });
  return built;
}
