/**
 * What officeParser's template tests take from the rest of officeParser, for their port here
 * (see docs/template-upstream.md, "Tests"):
 *
 * - `OfficeTemplate`: officeParser's class, as the kit's `render`, `tags` and `check`. A path (the
 *   upstream tests name their fixtures by path) is read here, in the test, since the kit takes bytes.
 * - `OfficeParser.parseOffice`: a stand-in, not officeParser's parser. It reads a filled `.docx` or
 *   `.odt` back as the paragraphs, tables and headers the tests look at: each body paragraph's text
 *   (`w:t`/`text:s`/tabs/breaks), a table's rows as lines of tab-separated cells, each run as a text
 *   node with `formatting.bold` when its own properties say bold. The tests use it for `includes`
 *   checks and one bold check; it is not a general reader.
 * - `parseXmlString`: a well-formedness check (fast-xml-parser's validator) in place of officeParser's
 *   XML reader.
 * - `withinTime`/`scaledMs`: officeParser's timing helpers with the calibration (which times
 *   officeParser's parser) replaced by a fixed factor, `DOCUMENT_KIT_TEST_SPEED_FACTOR` or 3.
 */
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';
import { strFromU8, unzipSync } from 'fflate';
import { XMLParser, XMLValidator } from 'fast-xml-parser';
import { check, render, tags } from '../../../src/template/officeTemplate.js';

/** The folder the upstream tests' `path.join(__dirname, 'files', ...)` resolves against. */
export const FIXTURE_ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');

const bytesOf = (template: any): any => typeof template === 'string' ? new Uint8Array(fs.readFileSync(template)) : template;

export const OfficeTemplate = {
    render: (template: any, config: any): Promise<any> => render(bytesOf(template), config),
    tags: (template: any, config?: any): Promise<any> => tags(bytesOf(template), config),
    check: (template: any, config: any): Promise<any> => check(bytesOf(template), config),
};

export function parseXmlString(xml: string): void {
    const result = XMLValidator.validate(xml);
    if (result !== true) throw new Error(`not well-formed: ${result.err.msg} (line ${result.err.line})`);
}

// ─── a reader of filled documents, for the tests' checks ───────────────────

type Node = Record<string, any>;
interface TextNode { type: 'text'; text: string; formatting?: { bold?: boolean } }
interface BlockNode { type: 'paragraph' | 'table'; text: string; children: TextNode[] }

const reader = new XMLParser({ preserveOrder: true, ignoreAttributes: false, attributeNamePrefix: '@', trimValues: false, parseTagValue: false, processEntities: true, htmlEntities: false });
const tagOf = (node: Node): string => Object.keys(node).find(k => k !== ':@') ?? '';
const kids = (node: Node): Node[] => node[tagOf(node)] ?? [];
const attr = (node: Node, name: string): string | undefined => node[':@']?.[`@${name}`];

/** A Word run's (or ODT span's) text, with tabs and breaks. */
function wordText(nodes: Node[], out: string[]): void {
    for (const n of nodes) {
        const tag = tagOf(n);
        if (tag === '#text') out.push(String(n['#text']));
        else if (tag === 'w:t') wordText(kids(n), out);
        else if (tag === 'w:tab') out.push('\t');
        else if (tag === 'w:br' || tag === 'w:cr') out.push('\n');
        else if (tag === 'w:delText' || tag === 'w:instrText' || tag === 'w:rPr' || tag === 'w:pPr' || tag === 'w:txbxContent') continue;
        else wordText(kids(n), out);
    }
}

function isBold(run: Node): boolean {
    const props = kids(run).find(k => tagOf(k) === 'w:rPr');
    const b = props && kids(props).find(k => tagOf(k) === 'w:b');
    if (!b) return false;
    const val = attr(b, 'w:val');
    return val === undefined || !['0', 'false', 'off'].includes(val);
}

function wordParagraph(p: Node): BlockNode {
    const children: TextNode[] = [];
    const walk = (nodes: Node[]) => {
        for (const n of nodes) {
            const tag = tagOf(n);
            if (tag === 'w:r') {
                const out: string[] = [];
                wordText(kids(n), out);
                children.push({ type: 'text', text: out.join(''), formatting: { bold: isBold(n) } });
            } else if (tag !== 'w:pPr' && tag !== '#text') walk(kids(n));
        }
    };
    walk(kids(p));
    return { type: 'paragraph', text: children.map(c => c.text).join(''), children };
}

function wordBlocks(nodes: Node[]): BlockNode[] {
    const blocks: BlockNode[] = [];
    for (const n of nodes) {
        const tag = tagOf(n);
        if (tag === 'w:p') blocks.push(wordParagraph(n));
        else if (tag === 'w:tbl') {
            const children: TextNode[] = [];
            const rows = kids(n).filter(r => tagOf(r) === 'w:tr').map(row => kids(row).filter(c => tagOf(c) === 'w:tc').map(cell => {
                const paragraphs = wordBlocks(kids(cell));
                for (const p of paragraphs) children.push(...p.children);
                return paragraphs.map(p => p.text).join('\n');
            }).join('\t'));
            blocks.push({ type: 'table', text: rows.join('\n'), children });
        } else if (tag === 'w:sdt' || tag === 'w:sdtContent' || tag === 'w:customXml') blocks.push(...wordBlocks(kids(n)));
    }
    return blocks;
}

function odtText(nodes: Node[], out: string[]): void {
    for (const n of nodes) {
        const tag = tagOf(n);
        if (tag === '#text') out.push(String(n['#text']));
        else if (tag === 'text:s') out.push(' '.repeat(Number(attr(n, 'text:c') ?? 1)));
        else if (tag === 'text:tab') out.push('\t');
        else if (tag === 'text:line-break') out.push('\n');
        else if (tag === 'office:annotation' || tag === 'text:note' || tag === 'draw:frame' || tag === 'text:tracked-changes') continue;
        else odtText(kids(n), out);
    }
}

function odtBlocks(nodes: Node[]): BlockNode[] {
    const blocks: BlockNode[] = [];
    for (const n of nodes) {
        const tag = tagOf(n);
        if (tag === 'text:p' || tag === 'text:h') {
            const out: string[] = [];
            odtText(kids(n), out);
            blocks.push({ type: 'paragraph', text: out.join(''), children: [{ type: 'text', text: out.join('') }] });
        } else if (tag === 'table:table') {
            const rows: string[] = [];
            const walkRows = (list: Node[]) => {
                for (const r of list) {
                    if (tagOf(r) === 'table:table-row') rows.push(kids(r).filter(c => tagOf(c) === 'table:table-cell').map(c => odtBlocks(kids(c)).map(p => p.text).join('\n')).join('\t'));
                    else if (tagOf(r).startsWith('table:table-')) walkRows(kids(r));
                }
            };
            walkRows(kids(n));
            blocks.push({ type: 'table', text: rows.join('\n'), children: [] });
        } else if (tag === 'text:list' || tag === 'text:list-item' || tag === 'text:section') blocks.push(...odtBlocks(kids(n)));
    }
    return blocks;
}

const find = (nodes: Node[], tag: string): Node | undefined => {
    for (const n of nodes) {
        if (tagOf(n) === tag) return n;
        const inner = find(kids(n), tag);
        if (inner) return inner;
    }
    return undefined;
};
const findAll = (nodes: Node[], tag: string, out: Node[] = []): Node[] => {
    for (const n of nodes) {
        if (tagOf(n) === tag) out.push(n);
        else findAll(kids(n), tag, out);
    }
    return out;
};

/** The stand-in for officeParser's parseOffice: see the module's comment. */
export const OfficeParser = {
    async parseOffice(bytes: Uint8Array, config: { fileType: 'docx' | 'odt' }): Promise<{ content: BlockNode[]; auxiliary: { headers: { text: string }[] } }> {
        const entries = unzipSync(bytes);
        const xml = (name: string) => entries[name] ? reader.parse(strFromU8(entries[name]!)) as Node[] : [];
        if (config.fileType === 'docx') {
            const body = find(xml('word/document.xml'), 'w:body');
            const headers = Object.keys(entries).filter(n => /^word\/header\d*\.xml$/.test(n))
                .map(n => ({ text: wordBlocks(kids(find(xml(n), 'w:hdr') ?? {})).map(b => b.text).join('\n') }));
            return { content: body ? wordBlocks(kids(body)) : [], auxiliary: { headers } };
        }
        const text = find(xml('content.xml'), 'office:text');
        const headers = findAll(xml('styles.xml'), 'style:header').map(h => ({ text: odtBlocks(kids(h)).map(b => b.text).join('\n') }));
        return { content: text ? odtBlocks(kids(text)) : [], auxiliary: { headers } };
    },
};

// ─── timing ────────────────────────────────────────────────────────────────

const factor = (() => {
    const set = process.env.DOCUMENT_KIT_TEST_SPEED_FACTOR;
    const value = set ? Number(set) : 3;
    if (!(value > 0) || !Number.isFinite(value)) throw new Error(`DOCUMENT_KIT_TEST_SPEED_FACTOR must be a number above 0, not "${set}"`);
    return value;
})();

/** `ms` of officeParser's reference machine's time, on this one. */
export const scaledMs = async (ms: number): Promise<number> => ms * factor;

/** `run` timed against `limitMs` (scaled), and run once more when over, the faster try kept. */
export const withinTime = async <T>(limitMs: number, run: () => T | Promise<T>): Promise<{ result: T; ms: number; within: boolean; limitMs: number; detail: string }> => {
    const limit = limitMs * factor;
    let started = performance.now();
    let result = await run();
    let ms = performance.now() - started;
    if (ms >= limit) {
        started = performance.now();
        result = await run();
        ms = Math.min(ms, performance.now() - started);
    }
    return { result, ms: Math.round(ms), within: ms < limit, limitMs: Math.round(limit), detail: `${Math.round(ms)} ms, limit ${Math.round(limit)} ms (${limitMs} ms x ${factor})` };
};
