/**
 * A picture's format and pixel size, read from its own header bytes (no DOM, no decoding), and the size it is drawn
 * at inside a box.
 */
import type { ImageData, ImageFormat } from './types'

export interface ImageInfo {
  format: ImageFormat
  width: number
  height: number
}

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]

function pngInfo(bytes: Uint8Array): ImageInfo | null {
  if (bytes.length < 24 || PNG_SIGNATURE.some((byte, i) => bytes[i] !== byte)) return null
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  // IHDR is the first chunk: length (4), "IHDR" (4), then width and height, big-endian.
  return { format: 'png', width: view.getUint32(16), height: view.getUint32(20) }
}

function jpegInfo(bytes: Uint8Array): ImageInfo | null {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return null
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let at = 2
  while (at + 4 <= bytes.length) {
    if (bytes[at] !== 0xff) {
      at += 1
      continue
    }
    const marker = bytes[at + 1] ?? 0
    if (marker === 0xff) {
      at += 1 // fill byte
      continue
    }
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd8)) {
      at += 2
      continue
    }
    if (marker === 0xd9 || marker === 0xda) return null
    const length = view.getUint16(at + 2)
    // Start of frame: SOF0–SOF15 except DHT (C4), JPG (C8) and DAC (CC). Its header holds the true size; an EXIF
    // thumbnail earlier in the file would not.
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      if (at + 9 > bytes.length) return null
      return { format: 'jpeg', height: view.getUint16(at + 5), width: view.getUint16(at + 7) }
    }
    if (length < 2) return null
    at += 2 + length
  }
  return null
}

/** The format and pixel size in a PNG or JPEG file's header, or null when the bytes are neither or are cut short. */
export function imageInfo(bytes: Uint8Array): ImageInfo | null {
  const info = pngInfo(bytes) ?? jpegInfo(bytes)
  return info && info.width > 0 && info.height > 0 ? info : null
}

/**
 * A picture's format and size: what the caller gave, the rest read from its header. Throws when the bytes are not a
 * PNG or JPEG the kit can read and the caller did not say.
 */
export function resolveImage(image: ImageData): ImageInfo {
  const read = image.format && image.width && image.height ? null : imageInfo(image.bytes)
  const format = image.format ?? read?.format
  const width = image.width ?? read?.width
  const height = image.height ?? read?.height
  if (!format || !width || !height) {
    throw new Error('document-kit: a picture is not a PNG or JPEG whose size can be read; give its format and size')
  }
  return { format, width, height }
}

/** Points per pixel, at 96 pixels to the inch. */
export const POINTS_PER_PIXEL = 72 / 96

/**
 * The size in points a picture of `width` × `height` pixels is drawn at to fit inside `box`, keeping its
 * proportions. Without `upscale`, a picture smaller than the box keeps its own size (pixels at 96 to the inch).
 */
export function fitImage(
  pixels: { width: number; height: number },
  box: { width?: number; height?: number },
  upscale = false,
): { width: number; height: number } {
  const natural = { width: pixels.width * POINTS_PER_PIXEL, height: pixels.height * POINTS_PER_PIXEL }
  const scales = [box.width === undefined ? Infinity : box.width / natural.width]
  scales.push(box.height === undefined ? Infinity : box.height / natural.height)
  let scale = Math.min(...scales)
  if (!Number.isFinite(scale)) scale = 1
  if (!upscale) scale = Math.min(scale, 1)
  return { width: natural.width * scale, height: natural.height * scale }
}
