/** Helpers for the kit's own template tests: Word templates built with the `docx` library, read back as XML. */
import { Document, Packer, type ISectionOptions } from 'docx';
import { strFromU8, unzipSync } from 'fflate';

/** A Word template of one section, as bytes. */
export async function wordTemplate(section: ISectionOptions): Promise<Uint8Array> {
    return new Uint8Array(await Packer.toArrayBuffer(new Document({ sections: [section] })));
}

/** Every part of a package, by name, as text. */
export function partsOf(bytes: Uint8Array): Record<string, string> {
    return Object.fromEntries(Object.entries(unzipSync(bytes)).map(([name, data]) => [name, strFromU8(data)]));
}

/** A part's paragraphs as plain text, one line each (Word's `w:p`, tabs and breaks kept). */
export function paragraphs(xml: string): string[] {
    return [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(([p]) => p
        .replace(/<w:tab\/>/g, '\t').replace(/<w:br\/>/g, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, '\'').replace(/&amp;/g, '&'));
}

/** The text of the document body, a paragraph a line. */
export function bodyText(bytes: Uint8Array): string {
    return paragraphs(partsOf(bytes)['word/document.xml']!).join('\n');
}
