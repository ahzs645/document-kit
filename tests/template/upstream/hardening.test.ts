/**
 * The template cases of officeParser 8.9.0's test/security/testSanitization.ts (MIT, Harsh Ankur),
 * ported to vitest: limits, linear time and bounded memory. Its helpers (`check`, `linear` and the
 * memory readings) are copied as they are; `check` collects failures and each vitest case fails when
 * any did. Changes are marked `document-kit:` and listed in docs/template-upstream.md, "Tests".
 */
import { describe, expect, it } from 'vitest';
import { createRequire } from 'node:module';
import { zipSync, strToU8, unzipSync, strFromU8 } from 'fflate';
import { OfficeTemplate, scaledMs, withinTime } from './harness.js';

// document-kit: an ES module has no require; V8's and vm's modules are reached through one.
const require = createRequire(import.meta.url);
type OfficeParserAST = any;

// document-kit: the failed checks' names, for the vitest case's message.
let failures: string[] = [];
let passed = 0;
let failed = 0;
const check = (name: string, cond: boolean, detail = '') => {
    if (cond) { passed++; }
    else { failed++; failures.push(`${name}${detail ? ` — ${detail}` : ''}`); console.error(`  ✗ FAIL: ${name}${detail ? ` — ${detail}` : ''}`); }
};

/**
 * Linear time, checked by how a case's time grows rather than against a clock: it runs at a quarter of
 * its size and then at its full size, here and now, and the full run may take at most eight times as long,
 * plus a quarter of a second for fixed costs (linear work takes four times as long, the collector's work
 * on a large heap can add some, quadratic work takes sixteen).
 * The machine's speed, and a slower power mode, cancel out. A case past that runs at both sizes once
 * more and the faster time of each is kept, so a pause of the machine (sleep, another process at work)
 * is not taken for the code's. A full run of over two minutes fails whatever its ratio. A case builds
 * its input with `n(count)`, which is `count` at full size and a quarter of it in the smaller run, and
 * the last run made is always the full one, so what it leaves behind is the full case's.
 */
let scale = 1;
const n = (count: number): number => Math.max(1, Math.round(count * scale));
const LINEAR_RATIO = 8;
const LINEAR_SLACK_MS = 250;
const LINEAR_BACKSTOP_MS = 120_000;
/**
 * The collector, called when a case asks (V8's flag, set as the suite runs). A measure of memory starts
 * from what is live: a reading taken without it counted what earlier cases left for the collector, and
 * one taken over a case counts what the case allocated and the collector has not yet freed, so whether
 * the collector happened to run decided the result (Node 22 failed in CI where Node 24 passed). From a
 * collected start, what a case adds is at most what it allocates, whenever the collector runs.
 */
const collectGarbage: () => void = (() => {
    try {
        (require('node:v8') as typeof import('node:v8')).setFlagsFromString('--expose-gc');
        const gc = (require('node:vm') as typeof import('node:vm')).runInNewContext('gc');
        if (typeof gc === 'function') return gc as () => void;
    } catch { /* no collector to call: readings start where they are */ }
    return () => { };
})();
/** The bytes V8 has allocated in all, where it counts them (`total_allocated_bytes`), else undefined. */
const allocatedSoFar = (): number | undefined => (require('node:v8') as typeof import('node:v8')).getHeapStatistics().total_allocated_bytes as number | undefined;
/**
 * What a run adds to the heap and to array buffers, from a collected start; what of the heap it still
 * holds once collected again with its result alive (`retained`: what it keeps, whenever the collector
 * runs); and what it allocated where V8 counts it (for a case's message: a limit at or above what a case
 * allocates holds whenever the collector runs).
 */
const memoryOf = async <T>(run: () => T | Promise<T>): Promise<{ result: T; heap: number; retained: number; buffers: number; allocated: number | undefined }> => {
    collectGarbage();
    const before = process.memoryUsage(), allocatedBefore = allocatedSoFar();
    const result = await run();
    const after = process.memoryUsage(), allocatedAfter = allocatedSoFar();
    collectGarbage();
    const retained = process.memoryUsage().heapUsed - before.heapUsed;
    return {
        result, heap: after.heapUsed - before.heapUsed, retained, buffers: after.arrayBuffers - before.arrayBuffers,
        allocated: allocatedBefore === undefined || allocatedAfter === undefined ? undefined : allocatedAfter - allocatedBefore,
    };
};
/** A memory reading as a case's message gives it: what was added, and what was allocated where known. */
const megabytes = (bytes: number | undefined) => (bytes === undefined ? '?' : `${Math.round(bytes / 1e6)}MB`);
const memoryNote = (grew: number, allocated: number | undefined) => `+${megabytes(grew)} (allocated ${megabytes(allocated)})`;

/** Holds a run's result past the collection that measures what it keeps. */
let keptResult: unknown;
const linear = async (name: string, run: () => unknown, heap?: { label: string; limit: number }): Promise<void> => {
    // The heap a full run keeps, when asked: from a collected start to a collected end with its result
    // alive (see memoryOf), the smaller one's leftovers not counted. What it allocated and let go is the
    // collector's to free: many of these cases allocate past their limit, so a reading without the
    // collector measured when it happened to run.
    let kept = 0, grown = 0;
    let allocated: number | undefined;
    const time = async (size: number): Promise<number> => {
        scale = size;
        const measuring = size === 1 && !!heap;
        if (measuring) collectGarbage();
        const heapBefore = process.memoryUsage().heapUsed, allocatedBefore = allocatedSoFar();
        const started = performance.now();
        let result: unknown;
        try { result = await run(); } finally { scale = 1; }
        const took = performance.now() - started;
        if (measuring) {
            grown = process.memoryUsage().heapUsed - heapBefore;
            const allocatedAfter = allocatedSoFar();
            allocated = allocatedBefore === undefined || allocatedAfter === undefined ? undefined : allocatedAfter - allocatedBefore;
            collectGarbage();
            kept = process.memoryUsage().heapUsed - heapBefore;
            keptResult = result;
            keptResult = undefined;
        }
        return took;
    };
    // The slack and the backstop are time, so they are this machine's (see test/helpers/timing.ts); the ratio is not.
    const slack = await scaledMs(LINEAR_SLACK_MS), backstop = await scaledMs(LINEAR_BACKSTOP_MS);
    let small = await time(0.25);
    let large = await time(1);
    const within = () => large <= LINEAR_RATIO * small + slack && large < backstop;
    if (!within()) {
        small = Math.min(small, await time(0.25));
        large = Math.min(large, await time(1));
    }
    check(`${name} in linear time`, within(), `${Math.round(small)}ms at a quarter of the size, ${Math.round(large)}ms at full size`);
    if (heap) check(heap.label, kept < heap.limit, `${megabytes(kept)} kept, ${memoryNote(grown, allocated)}`);
};

// document-kit: each vitest case fails when a check in it failed, naming them.
const expectNoFailures = () => expect(failed, failures.join('\n')).toBe(0);

/** The template cases of upstream's parserHardeningTests. */
async function templateHardeningTests() {
    const enc = (s: string) => new TextEncoder().encode(s);
    const timed = linear;
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
    const docxOf = (body: string, extra: Record<string, string | Uint8Array> = {}) => Buffer.from(zipSync({
        '[Content_Types].xml': enc('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
        'word/document.xml': enc(`<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`),
        ...Object.fromEntries(Object.entries(extra).map(([k, v]) => [k, typeof v === 'string' ? enc(v) : v])),
    }));
    // Templates, raw source and chunks.
    await timed('template: a part of 160,000 unclosed <w:p is rendered', () => OfficeTemplate.render(Buffer.from(zipSync({ '[Content_Types].xml': enc('<Types/>'), 'word/document.xml': enc('<w:p '.repeat(n(160000))) })), { data: { a: 1 } }).catch(() => undefined));
    // The template renderer rebuilds a paragraph from slices, not a character at a time.
    await timed('template: one paragraph of 20 million characters renders', () => OfficeTemplate.render(docxOf(`<w:p><w:r><w:t>{{a}}${'x'.repeat(n(20_000_000))}</w:t></w:r></w:p>`), { data: { a: 'y' } }),
        { label: 'template: one paragraph of 20 million characters around a placeholder renders without copying it per reference', limit: 300_000_000 });
    // Templates: a value repeated at many placeholders is charged past its first use.
    const repeatedTemplate = docxOf(`<w:p><w:r><w:t>${'{{a}}'.repeat(200000)}</w:t></w:r></w:p>`);
    let repeatError: any;
    const repeatStarted = Date.now();
    try { await OfficeTemplate.render(repeatedTemplate, { data: { a: 'x'.repeat(100) } }); } catch (e) { repeatError = e; }
    check('template: a 100-character value at 200,000 placeholders is refused', repeatError?.officeIssue?.code === 'OUTPUT_TOO_LARGE', `${Date.now() - repeatStarted}ms ${repeatError}`);
    const fewRepeats = await OfficeTemplate.render(docxOf(`<w:p><w:r><w:t>${'{{a}} '.repeat(10)}</w:t></w:r></w:p>`), { data: { a: 'x'.repeat(1000) } });
    check('template: a long value at ten placeholders renders', fewRepeats instanceof Uint8Array && fewRepeats.length > 0, String(fewRepeats));
    // Templates with block tags: every scan is forward-only, nesting is bounded, repetition is charged.
    const renderQuiet = async (body: string, data: any, cfg: any = {}, parts: Record<string, string> = {}) => {
        try { return { out: await OfficeTemplate.render(docxOf(body, parts), { data, onWarning: () => { }, ...cfg }) as Uint8Array, code: undefined }; }
        catch (e: any) { return { out: undefined, code: e?.officeIssue?.code ?? String(e) }; }
    };
    const tplPara = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
    for (const [label, make] of [
        ['200,000 unclosed if tags in one paragraph', () => [tplPara('{{ if x }}'.repeat(n(200000))), { x: 1 }, 'MAX_NESTING_DEPTH_EXCEEDED']],
        ['200,000 closing tags with no block', () => [tplPara('{{ endif }}'.repeat(n(200000))), {}, 'TEMPLATE_SYNTAX']],
        ['60,000 sibling inline blocks', () => [tplPara('{{ if x }}a{{ endif }}'.repeat(n(60000))), { x: 1 }, undefined]],
        ['30,000 paragraph blocks', () => [(tplPara('{{ if x }}') + tplPara('a') + tplPara('{{ endif }}')).repeat(n(30000)), { x: 0 }, undefined]],
        ['a table of 30,000 rows each in a row loop', () => [`<w:tbl>${`<w:tr><w:tc>${tplPara('{{ for v in vs }}{{ v }}')}</w:tc><w:tc>${tplPara('{{ endfor }}')}</w:tc></w:tr>`.repeat(n(30000))}</w:tbl>`, { vs: [1] }, undefined]],
        ['a condition of 100,000 ands', () => [tplPara(`{{ if ${'a and '.repeat(n(100000))}a }}y{{ endif }}`), { a: true }, undefined]],
        ['a condition of 100,000 open parentheses', () => [tplPara(`{{ if ${'('.repeat(n(100000))}a }}y{{ endif }}`), { a: true }, 'TEMPLATE_SYNTAX']],
        ['a condition of 100,000 nots', () => [tplPara(`{{ if ${'not '.repeat(n(100000))}a }}y{{ endif }}`), { a: true }, 'TEMPLATE_SYNTAX']],
        ['a condition with an unclosed 100,000-character string', () => [tplPara(`{{ if a == ${'"'.repeat(1)}${'x'.repeat(n(100000))} }}y{{ endif }}`), { a: 'x' }, 'TEMPLATE_SYNTAX']],
        ['a path of 50,000 steps', () => [tplPara(`{{ if ${'a.'.repeat(n(50000))}a }}y{{ endif }}`), { a: 1 }, undefined]],
        ['a block before 160,000 unclosed <w:p', () => [tplPara('{{ if a }}x{{ endif }}') + '<w:p '.repeat(n(160000)), { a: 1 }, 'TEMPLATE_SYNTAX']],
        ['a block before 100,000 start tags with an unclosed quote', () => [tplPara('{{ if a }}x{{ endif }}') + '<w:x a="'.repeat(n(100000)), { a: 1 }, 'TEMPLATE_SYNTAX']],
        ['a block before 100,000 unclosed elements', () => [tplPara('{{ if a }}x{{ endif }}') + '<w:x>'.repeat(n(100000)), { a: 1 }, 'TEMPLATE_SYNTAX']],
        ['a block among 200,000 fields in one paragraph', () => [tplPara('{{ if a }}' + '{{a}}'.repeat(n(200000)) + '{{ endif }}'), { a: 1 }, undefined]],
        ['three nested loops over 1,000 items with nothing to write', () => [tplPara('{{ for a in xs }}{{ for b in xs }}{{ for c in xs }}{{ endfor }}{{ endfor }}{{ endfor }}'), { xs: Array(n(1000)).fill(0) }, 'OUTPUT_TOO_LARGE']],
        ['two nested loops over 3,000 items writing a paragraph each', () => [tplPara('{{ for a in xs }}') + tplPara('{{ for b in xs }}') + tplPara('row') + tplPara('{{ endfor }}') + tplPara('{{ endfor }}'), { xs: Array(n(3000)).fill(0) }, 'OUTPUT_TOO_LARGE']],
        ['a loop over 1,000,000 items', () => [tplPara('{{ for a in xs }}{{ endfor }}'), { xs: Array(n(1000000)).fill(0) }, undefined]],
        ['an if with 48,000 else ifs over paragraphs', () => [tplPara('x {{ if c }}') + Array.from({ length: n(48000) }, (_, i) => tplPara(`{{ else if c${i} }}`) + tplPara('y')).join('') + tplPara('{{ endif }}'), { c: true }, undefined]],
        ['8,000 tags in 8,000 nested smart tags', () => ['<w:p><w:r><w:t>{{ if a }}A</w:t></w:r>' + '<w:smartTag>'.repeat(n(8000)) + Array.from({ length: n(8000) }, (_, i) => `<w:r><w:t>{{ else if c${i} }}x</w:t></w:r>`).join('') + '</w:smartTag>'.repeat(n(8000)) + '<w:r><w:t>{{ endif }}</w:t></w:r></w:p>', { a: 1 }, 'TEMPLATE_SYNTAX']],
        ['two nested loops over 400 items evaluating 1,000 empty fields', () => [tplPara('{{ for a in xs }}{{ for b in xs }}' + '{{ v }}'.repeat(n(1000)) + '{{ endfor }}{{ endfor }}'), { xs: Array(n(400)).fill(0), v: '' }, 'OUTPUT_TOO_LARGE']],
        ['two nested loops over 300 items evaluating 1,000 empty branches', () => [tplPara('{{ for a in xs }}{{ for b in xs }}' + '{{ if f }}{{ endif }}'.repeat(n(1000)) + '{{ endfor }}{{ endfor }}'), { xs: Array(n(300)).fill(0), f: false }, 'OUTPUT_TOO_LARGE']],
        ['a loop referring 2,000 times to a 300 KB footnote', () => [tplPara('{{ for i in xs }}') + '<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>' + tplPara('{{ endfor }}'), { xs: Array(n(2000)).fill(0) }, 'OUTPUT_TOO_LARGE',
            { 'word/footnotes.xml': `<?xml version="1.0"?><w:footnotes ${W}><w:footnote w:id="1">${tplPara('x'.repeat(n(300000)))}</w:footnote></w:footnotes>` }]],
        ['a missing 100,000-character field kept as typed in 1,000 passes', () => [tplPara(`{{ for i in xs }}{{ ${'z'.repeat(n(100000))} }}{{ endfor }}`), { xs: Array(n(1000)).fill(0) }, 'OUTPUT_TOO_LARGE']],
        ['a block with a missing 100,000-character condition kept as typed in 1,000 passes', () => [tplPara(`{{ for i in xs }}{{ if ${'z'.repeat(n(100000))} }}x{{ endif }}{{ endfor }}`), { xs: Array(n(1000)).fill(0) }, 'OUTPUT_TOO_LARGE']],
        ['a 10,000-operand condition tested in nested loops over 300 items', () => [tplPara(`{{ for x in xs }}{{ for y in xs }}{{ if ${'a and '.repeat(n(9999))}a }}{{ endif }}{{ endfor }}{{ endfor }}`), { xs: Array(n(300)).fill(0), a: true }, 'OUTPUT_TOO_LARGE']],
        ['16,000 tags in one run with 16 KB of run properties', () => [`<w:p><w:r><w:rPr><w:rFonts w:ascii="${'A'.repeat(n(16000))}"/></w:rPr><w:t>${'{{ if a }}x{{ endif }}'.repeat(n(16000))}</w:t></w:r></w:p>`, { a: true }, 'OUTPUT_TOO_LARGE']],
        ['a kept block passing 50,000 else ifs in each of 20,000 passes', () => [tplPara('{{ for o in xs }}') + tplPara('{{ if a }}') + tplPara('{{ for y in nope }}') + tplPara('B') + tplPara('{{ endfor }}' + Array.from({ length: n(50000) }, (_, i) => `{{ else if b${i} }}`).join('')) + tplPara('{{ endif }}') + tplPara('{{ endfor }}'), { xs: Array(n(20000)).fill(0), a: true }, 'OUTPUT_TOO_LARGE']],
        ['two 10 MB strings compared in 10,000 passes', () => [tplPara('{{ for i in xs }}{{ if a == b }}Y{{ endif }}{{ endfor }}'), { xs: Array(n(10000)).fill(0), a: 'x'.repeat(n(10_000_000)) + 'a', b: 'x'.repeat(n(10_000_000)) + 'b' }, 'OUTPUT_TOO_LARGE']],
        ['200,000 paragraphs before one block', () => [tplPara('a').repeat(n(200000)) + tplPara('{{ if x }}y{{ endif }}'), { x: true }, undefined]],
        ['400,000 comments before one block', () => ['<!---->'.repeat(n(400000)) + tplPara('{{ if x }}y{{ endif }}'), { x: true }, undefined]],
        ['a 1,000,000-character loop name read 10 times in each of 100,000 passes', () => [tplPara(`{{ for ${'n'.repeat(n(1000000))} in xs }}${`{{ ${'n'.repeat(n(1000000))} }}`.repeat(10)}{{ endfor }}`), { xs: Array(n(100000)).fill(0) }, 'OUTPUT_TOO_LARGE']],
        ['a row 10,000 wrappers deep repeated 100,000 times', () => ['<w:customXml w:element="e">'.repeat(n(10000)) + '<w:p><w:r><w:t>{{ for x in xs }}</w:t></w:r><w:tr/><w:r><w:t>{{ endfor }}</w:t></w:r></w:p>' + '</w:customXml>'.repeat(n(10000)), { xs: Array(n(100000)).fill(0) }, undefined]],
        ['60 nested blocks around 1 MB, each kept on its own pass', () => [tplPara('{{ for x in xs }}') + Array.from({ length: 60 }, (_, i) => tplPara(`{{ if x.c${i} }}`)).join('') + tplPara('Z'.repeat(n(1000000))) + tplPara('{{ endif }}').repeat(60) + tplPara('{{ endfor }}'), { xs: Array.from({ length: 60 }, (_, i) => Object.fromEntries(Array.from({ length: i }, (_, j) => [`c${j}`, true]))) }, 'OUTPUT_TOO_LARGE']],
        ['a quoted string holding 80,000 no-break spaces inside guillemets', () => [tplPara(`{{ if x == \u00ABa${'\u00A0'.repeat(n(80000))}b\u00BB }}Y{{ endif }}`), { x: 'a' }, undefined]],
        ['16 headers each copying run properties into 1,900 cuts', () => [tplPara('body'), { a: true }, 'OUTPUT_TOO_LARGE',
            Object.fromEntries(Array.from({ length: 16 }, (_, i) => [`word/header${i + 1}.xml`, `<?xml version="1.0"?><w:hdr ${W}><w:p><w:r><w:rPr><w:rFonts w:ascii="${'A'.repeat(n(4000))}"/></w:rPr><w:t>${'{{ if a }}x{{ endif }}y'.repeat(n(1900))}</w:t></w:r></w:p></w:hdr>`]))]],
        ['a value of 15,000,000 line breaks', () => [tplPara('{{ v }}'), { v: '\n'.repeat(n(15_000_000)) }, 'OUTPUT_TOO_LARGE']],
        ['3,000,000 empty text runs beside a field', () => ['<w:p><w:r><w:t>{{ a }}</w:t></w:r>' + '<w:r><w:t/></w:r>'.repeat(n(3_000_000)) + '</w:p>', { a: 1 }, 'XML_ELEMENT_LIMIT_EXCEEDED']],
        ['a 100,000-placeholder number picture written in 1,000 passes', () => [tplPara(`{{ for i in xs }}{{ v | format "${'0'.repeat(n(100000))}" }}{{ endfor }}`), { xs: Array(n(1000)).fill(0), v: 1 }, 'OUTPUT_TOO_LARGE']],
        ['a date picture of 100,000 parts', () => [tplPara(`{{ v | format "${'d M y '.repeat(n(20000))}" }}`), { v: '2026-11-01' }, undefined]],
        ['a chain of 10,000 filters', () => [tplPara(`{{ v ${'| upper '.repeat(n(10000))}}}`), { v: 'abc' }, undefined]],
        ['upper on 1 MB of text in 1,000 passes', () => [tplPara('{{ for i in xs }}{{ v | upper }}{{ endfor }}'), { xs: Array(n(1000)).fill(0), v: 'x'.repeat(n(1_000_000)) }, 'OUTPUT_TOO_LARGE']],
        ['a 1,000,000-digit text that is no number given to a number picture', () => [tplPara('{{ v | format "0" }}'), { v: '1'.repeat(n(1_000_000)) + 'x' }, 'TEMPLATE_DATA_INVALID']],
        ['prose of 100,000 pipes between delimiters', () => [tplPara(`{{ a ${'| '.repeat(n(100000))}}} {{ b }}`), { b: 1 }, undefined]],
        ['a picture with 100,000 unclosed quotes', () => [tplPara(`{{ v | format "0${"'".repeat(1)}${'x'.repeat(n(100000))}" }}`), { v: 1 }, 'TEMPLATE_SYNTAX']],
        ['the smallest number through a number picture in 400,000 passes', () => [tplPara('{{ for i in xs }}{{ v | format "0" }}{{ endfor }}'), { xs: Array(n(400000)).fill(0), v: 5e-324 }, 'OUTPUT_TOO_LARGE']],
        ['a text of a tiny number through a number picture in 100,000 passes', () => [tplPara('{{ for i in xs }}{{ v | format "#,##0.00" }}{{ endfor }}'), { xs: Array(n(100000)).fill(0), v: '1.2345e-300' }, undefined]],
        ['an if with 32,000 else ifs 32,000 elements deep, closed outside them', () => ['<w:customXml w:element="e">'.repeat(n(32000)) + tplPara('{{ if a }}' + '{{ else if a }}'.repeat(n(32000))) + '</w:customXml>'.repeat(n(32000)) + tplPara('{{ endif }}'), { a: true }, 'TEMPLATE_SYNTAX']],
        ['32,000 else-if paragraphs 32,000 elements deep, closed outside them', () => ['<w:customXml w:element="e">'.repeat(n(32000)) + tplPara('{{ if a }}') + tplPara('{{ else if a }}').repeat(n(32000)) + '</w:customXml>'.repeat(n(32000)) + tplPara('{{ endif }}'), { a: true }, 'TEMPLATE_SYNTAX']],
        ['64,000 nested elements with paragraph ids in a loop', () => [tplPara('{{ for a in xs }}') + '<w:p>' + '<w:customXml w:element="e" w14:paraId="00000001">'.repeat(n(64000)) + '<w:r><w:t>x</w:t></w:r>' + '</w:customXml>'.repeat(n(64000)) + '</w:p>' + tplPara('{{ endfor }}'), { xs: [1] }, undefined]],
        ['64,000 nested bookmark ends of a range crossing a loop, around a field', () => ['<w:p><w:bookmarkStart w:id="5" w:name="bm"/></w:p>' + tplPara('{{ for a in xs }}') + '<w:p>' + '<w:bookmarkEnd w:id="5">'.repeat(n(64000)) + '<w:r><w:t>{{ a }}</w:t></w:r>' + '</w:bookmarkEnd>'.repeat(n(64000)) + '</w:p>' + tplPara('{{ endfor }}'), { xs: [1] }, undefined]],
        ['16,000,000 fields in one run (117 KB of zip)', () => [tplPara('{{a}}'.repeat(n(16_000_000))), { a: 1 }, 'XML_ELEMENT_LIMIT_EXCEEDED']],
        ['4,000,000 inline blocks in one run', () => [tplPara('{{ if a }}k{{ endif }}'.repeat(n(4_000_000))), { a: 1 }, 'XML_ELEMENT_LIMIT_EXCEEDED']],
        ['one condition of 4,000,000 operands', () => [tplPara(`{{ if a${' and a'.repeat(n(4_000_000))} }}x{{ endif }}`), { a: 1 }, 'XML_ELEMENT_LIMIT_EXCEEDED']],
        ['one field through 3,000,000 filters', () => [tplPara(`{{ a${' | upper'.repeat(n(3_000_000))} }}`), { a: 'x' }, 'XML_ELEMENT_LIMIT_EXCEEDED']],
        ['one filter given 3,000,000 arguments', () => [tplPara(`{{ a | upper${' 1'.repeat(n(3_000_000))} }}`), { a: 'x' }, 'XML_ELEMENT_LIMIT_EXCEEDED']],
        ['one date picture of 4,000,000 parts', () => [tplPara(`{{ a | format "${'dM'.repeat(n(2_000_000))}" }}`), { a: '2026-01-01' }, 'XML_ELEMENT_LIMIT_EXCEEDED']],
        ['one 1 MB number text through a number picture in 50,000 passes', () => [tplPara('{{ for i in xs }}{{ v | format "0" }}{{ endfor }}'), { xs: Array(n(50000)).fill(0), v: '1' + ' '.repeat(n(1_000_000)) }, undefined]],
        ['a list of 1 MB number texts through a number picture', () => [tplPara('{{ for i in xs }}{{ i | format "0" }}{{ endfor }}'), { xs: Array(n(2000)).fill('1' + ' '.repeat(n(1_000_000))) }, 'OUTPUT_TOO_LARGE']],
        ['a list of dates each followed by 1 MB of spaces through a date picture', () => [tplPara('{{ for i in xs }}{{ i | format "d" }}{{ endfor }}'), { xs: Array(n(2000)).fill('2026-01-01' + ' '.repeat(n(1_000_000))) }, 'OUTPUT_TOO_LARGE']],
        ['an if, then 40,000 else ifs in a run 40,000 links deep', () => [`<w:p><w:r><w:t>{{ if a }}</w:t></w:r>${'<w:hyperlink>'.repeat(n(40000))}<w:r><w:t>${'{{ else if a }}'.repeat(n(40000))}{{ endif }}</w:t></w:r>${'</w:hyperlink>'.repeat(n(40000))}</w:p>`, { a: false }, 'TEMPLATE_SYNTAX']],
        ['an if, then 40,000 else-if paragraphs 40,000 content controls deep', () => [tplPara('{{ if a }}') + '<w:sdt><w:sdtContent>'.repeat(n(40000)) + tplPara('{{ else if a }}').repeat(n(40000)) + '</w:sdtContent></w:sdt>'.repeat(n(40000)) + tplPara('{{ endif }}'), { a: false }, 'TEMPLATE_SYNTAX']],
        ['40,000 inline blocks 40,000 content controls deep', () => ['<w:sdt><w:sdtContent>'.repeat(n(40000)) + tplPara('{{ if a }}x{{ endif }}').repeat(n(40000)) + '</w:sdtContent></w:sdt>'.repeat(n(40000)), { a: 1 }, undefined]],
        ['20,000 paragraph blocks 20,000 content controls deep', () => ['<w:sdt><w:sdtContent>'.repeat(n(20000)) + (tplPara('{{ if a }}') + tplPara('x') + tplPara('{{ endif }}')).repeat(n(20000)) + '</w:sdtContent></w:sdt>'.repeat(n(20000)), { a: 1 }, undefined]],
        ['3,000,000 nested text runs beside a block', () => [tplPara('{{ if a }}x{{ endif }}') + '<w:t>'.repeat(n(3_000_000)) + '</w:t>'.repeat(n(3_000_000)), { a: 1 }, 'XML_ELEMENT_LIMIT_EXCEEDED']],
        ['3,000,000 nested text runs in a header without a tag (copied, never read)', () => [tplPara('{{ a }}'), { a: 1 }, undefined, { 'word/header1.xml': `<?xml version="1.0"?><w:hdr ${W}><w:p><w:r>${'<w:t>'.repeat(n(3_000_000))}${'</w:t>'.repeat(n(3_000_000))}</w:r></w:p></w:hdr>` }]],
        ['a loop repeating an id beside a header of 300,000 paragraph ids, the bound 200,000', () => [tplPara('{{ for x in xs }}') + '<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="p"/></wp:inline></w:drawing></w:r></w:p>' + tplPara('{{ endfor }}'), { xs: [1, 2] }, 'XML_ELEMENT_LIMIT_EXCEEDED',
            { 'word/header1.xml': `<?xml version="1.0"?><w:hdr ${W} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">${'<w:p w14:paraId="0000000A"/>'.repeat(n(300_000))}</w:hdr>` }, { decompressionLimits: { maxXmlElements: 200_000 } }]],
        ['a loop repeating an id beside a header of one field and 300,000 distinct paragraph ids, the bound 200,000', () => [tplPara('{{ for x in xs }}') + '<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="p"/></wp:inline></w:drawing></w:r></w:p>' + tplPara('{{ endfor }}'), { xs: [1, 2], a: 1 }, 'XML_ELEMENT_LIMIT_EXCEEDED',
            { 'word/header1.xml': `<?xml version="1.0"?><w:hdr ${W} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">${tplPara('{{ a }}')}${Array.from({ length: n(300_000) }, (_, i) => `<w:p w14:paraId="${(i + 1).toString(16).toUpperCase().padStart(8, '0')}"/>`).join('')}</w:hdr>` }, { decompressionLimits: { maxXmlElements: 200_000 } }]],
        ['a loop repeating an id beside 300,000 footnotes without a tag, the bound 200,000', () => [tplPara('{{ for x in xs }}') + '<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="p"/></wp:inline></w:drawing></w:r></w:p>' + tplPara('{{ endfor }}'), { xs: [1, 2] }, 'XML_ELEMENT_LIMIT_EXCEEDED',
            { 'word/footnotes.xml': `<?xml version="1.0"?><w:footnotes ${W}>${Array.from({ length: n(300_000) }, (_, i) => `<w:footnote w:id="${i + 1}"/>`).join('')}</w:footnotes>` }, { decompressionLimits: { maxXmlElements: 200_000 } }]],
        ['a loop repeating an id beside a paragraph of 3,000,000 distinct attributes', () => [tplPara('{{ for x in xs }}') + '<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="p"/></wp:inline></w:drawing></w:r></w:p>' + tplPara('{{ endfor }}'), { xs: [1, 2] }, undefined,
            { 'word/header1.xml': `<?xml version="1.0"?><w:hdr ${W}><w:p ${Array.from({ length: n(3_000_000) }, (_, i) => `a${i.toString(36)}=""`).join(' ')}/></w:hdr>` }]],
        ['64,000 tags with a quote outside their values, holding paragraph ids, in a loop', () => [tplPara('{{ for x in xs }}') + `<w:p>${'<w:proofErr \'b w14:paraId="0000000A" \'/>'.repeat(n(64000))}<w:r><w:t>{{ x }}</w:t></w:r></w:p>` + tplPara('{{ endfor }}'), { xs: [1] }, undefined]],
        ['400,000 comments before a field in one run', () => [`<w:p><w:r>${'<!---->'.repeat(n(400000))}<w:t>{{ a }}</w:t></w:r></w:p>`, { a: 1 }, undefined]],
        ['400,000 processing instructions before a field in one run', () => [`<w:p><w:r>${'<?a?>'.repeat(n(400000))}<w:t>{{ a }}</w:t></w:r></w:p>`, { a: 1 }, undefined]],
        ['200,000 comments among runs, one processing instruction at the end', () => [`<w:p>${'<w:r><!----><w:t>{</w:t></w:r>'.repeat(n(200000))}<w:r><w:t>{{ a }}</w:t></w:r><?a?></w:p>`, { a: 1 }, undefined]],
        ['a loop copying 50,000 pictures twice', () => [tplPara('{{ for a in xs }}') + '<w:p><w:r><w:drawing><wp:docPr id="1"/></w:drawing></w:r></w:p>'.repeat(n(50000)) + tplPara('{{ endfor }}'), { xs: [0, 1] }, undefined]],
    ] as [string, () => [string, any, string | undefined, Record<string, string>?, any?]][]) {
        // Each row builds its case when it runs, at the size the run asks for (see linear); what is
        // checked is the last run's, the full one.
        let result: { out: Uint8Array | undefined; code: string | undefined } = { out: undefined, code: undefined };
        let wanted: string | undefined;
        // A case whose own work allocates more than the 300 MB the others are held to is held to a little
        // more than it allocates (measured with V8's count), so the collector's timing cannot fail it, and
        // still far less than copying per reference costs there (its blocks times its depth or its runs).
        const ownWork: Record<string, number> = {
            'a block among 200,000 fields in one paragraph': 500_000_000, // allocates about 376 MB
            '20,000 paragraph blocks 20,000 content controls deep': 750_000_000, // about 548 MB
            '40,000 inline blocks 40,000 content controls deep': 1_000_000_000, // about 807 MB
            '3,000,000 nested text runs beside a block': 750_000_000, // about 548 MB
        };
        await timed(`template: ${label}`, async () => {
            const [body, data, expected, parts, cfg] = make();
            wanted = expected;
            result = await renderQuiet(body, data, cfg ?? {}, parts);
        },
            { label: `template: ${label} without copying it per reference`, limit: ownWork[label] ?? 300_000_000 });
        check(`template: ${label} ${wanted ? `is refused with ${wanted}` : 'renders'}`, result.code === wanted && (wanted !== undefined || result.out instanceof Uint8Array), String(result.code));
    }
    // Prototype keys in a template's paths are never read from the prototype; a data key named
    // __proto__ (as JSON.parse makes it) is the data's own.
    const protoList = await renderQuiet(tplPara('{{ for x in __proto__ }}[{{ x }}]{{ endfor }}{{ if constructor }}C{{ endif }}{{ if toString }}T{{ endif }}'), {}, { onMissing: 'empty' });
    check('template: __proto__, constructor and toString are missing fields, not the prototype', !protoList.code && !strFromU8(unzipSync(protoList.out!)['word/document.xml']).match(/\[|C<|T</), String(protoList.code));
    const ownProto = await renderQuiet(tplPara('{{ for x in __proto__ }}[{{ x }}]{{ endfor }}'), JSON.parse('{"__proto__": ["own"]}'));
    check('template: a data key named __proto__ is the data\'s own list', !ownProto.code && strFromU8(unzipSync(ownProto.out!)['word/document.xml']).includes('[own]'), String(ownProto.code));
}

async function odtTemplateHardeningTests() {
    console.log('- ODT templates, text that looks like a tag, numbers written as text (linear time, bounds)...');
    const enc = (s: string) => new TextEncoder().encode(s);
    const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"';
    const odtOf = (body: string, styles = '<text:p>H</text:p>') => Buffer.from(zipSync({
        mimetype: [enc('application/vnd.oasis.opendocument.text'), { level: 0 }],
        'META-INF/manifest.xml': enc('<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/></manifest:manifest>'),
        'content.xml': enc(`<?xml version="1.0"?><office:document-content ${NS}><office:body><office:text>${body}</office:text></office:body></office:document-content>`),
        'styles.xml': enc(`<?xml version="1.0"?><office:document-styles ${NS}><office:master-styles><style:master-page style:name="S"><style:header>${styles}</style:header></style:master-page></office:master-styles></office:document-styles>`),
    }));
    const docxOf = (body: string) => Buffer.from(zipSync({
        '[Content_Types].xml': enc('<Types/>'),
        'word/document.xml': enc(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`),
    }));
    const renderQuiet = async (bytes: Buffer, data: any, cfg: any = {}) => {
        try { return { out: await OfficeTemplate.render(bytes, { data, onWarning: () => { }, ...cfg }) as Uint8Array, code: undefined }; }
        catch (e: any) { return { out: undefined, code: e?.officeIssue?.code ?? String(e) }; }
    };
    const p = (t: string) => `<text:p>${t}</text:p>`;
    const wp = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
    for (const [label, make] of [
        ['200,000 unclosed if tags in one paragraph', () => [odtOf(p('{{ if x }}'.repeat(n(200000)))), { x: 1 }, 'MAX_NESTING_DEPTH_EXCEEDED']],
        ['200,000 closing tags with no block', () => [odtOf(p('{{ endif }}'.repeat(n(200000)))), {}, 'TEMPLATE_SYNTAX']],
        ['60,000 sibling inline blocks', () => [odtOf(p('{{ if x }}a{{ endif }}'.repeat(n(60000)))), { x: 1 }, undefined]],
        ['30,000 paragraph blocks', () => [odtOf((p('{{ if x }}') + p('a') + p('{{ endif }}')).repeat(n(30000))), { x: 0 }, undefined]],
        ['30,000 list-item blocks', () => [odtOf(`<text:list>${('<text:list-item>' + p('{{ if x }}') + '</text:list-item><text:list-item>' + p('a') + '</text:list-item><text:list-item>' + p('{{ endif }}') + '</text:list-item>').repeat(n(30000))}</text:list>`), { x: 0 }, undefined]],
        ['a table of 30,000 rows each in a row loop', () => [odtOf(`<table:table>${`<table:table-row><table:table-cell>${p('{{ for v in vs }}{{ v }}')}</table:table-cell><table:table-cell>${p('{{ endfor }}')}</table:table-cell></table:table-row>`.repeat(n(30000))}</table:table>`), { vs: [1] }, undefined]],
        ['a block before 160,000 unclosed <text:p', () => [odtOf(p('{{ if a }}x{{ endif }}') + '<text:p '.repeat(n(160000))), { a: 1 }, 'TEMPLATE_SYNTAX']],
        ['a field before 160,000 unclosed <text:p', () => [odtOf(p('{{ a }}') + '<text:p '.repeat(n(160000))), { a: 1 }, undefined]],
        ['a block before 100,000 start tags with an unclosed quote', () => [odtOf(p('{{ if a }}x{{ endif }}') + '<text:x a="'.repeat(n(100000))), { a: 1 }, 'TEMPLATE_SYNTAX']],
        ['a block before 100,000 unclosed elements', () => [odtOf(p('{{ if a }}x{{ endif }}') + '<text:x>'.repeat(n(100000))), { a: 1 }, 'TEMPLATE_SYNTAX']],
        ['a field among 200,000 stray end tags', () => [odtOf(p('{{ a }}' + '</text:span>'.repeat(n(200000)))), { a: 1 }, undefined]],
        ['three nested loops over 1,000 items with nothing to write', () => [odtOf(p('{{ for a in xs }}{{ for b in xs }}{{ for c in xs }}{{ endfor }}{{ endfor }}{{ endfor }}')), { xs: Array(n(1000)).fill(0) }, 'OUTPUT_TOO_LARGE']],
        ['two nested loops over 3,000 items writing a paragraph each', () => [odtOf(p('{{ for a in xs }}') + p('{{ for b in xs }}') + p('row') + p('{{ endfor }}') + p('{{ endfor }}')), { xs: Array(n(3000)).fill(0) }, 'OUTPUT_TOO_LARGE']],
        ['a loop over 1,000,000 items', () => [odtOf(p('{{ for a in xs }}{{ endfor }}')), { xs: Array(n(1000000)).fill(0) }, undefined]],
        ['a loop copying a 300 KB footnote 2,000 times', () => [odtOf(p('{{ for i in xs }}') + p(`<text:note text:id="n1"><text:note-body>${p('x'.repeat(n(300000)))}</text:note-body></text:note>`) + p('{{ endfor }}')), { xs: Array(n(2000)).fill(0) }, 'OUTPUT_TOO_LARGE']],
        // Inline blocks inside one span are cut inside it: nothing of the span is copied per tag (Word's run is).
        ['16,000 tags in one span with 16 KB of attributes', () => [odtOf(p(`<text:span text:style-name="${'A'.repeat(n(16000))}">${'{{ if a }}x{{ endif }}'.repeat(n(16000))}</text:span>`)), { a: true }, undefined]],
        // A block cut out of a span copies the span's start tag, charged as repetition: here 2,000 spans of 1 KB each, against 100 KB.
        ['2,000 blocks each cut out of a span of 1 KB of attributes', () => [odtOf(p(`<text:span text:style-name="${'A'.repeat(1000)}">{{ if a }}x</text:span>{{ endif }}`.repeat(n(2000)))), { a: true }, 'OUTPUT_TOO_LARGE', { decompressionLimits: { maxRepeatedContent: 100_000 } }]],
        ['a value of 30,000,000 line breaks', () => [odtOf(p('{{ v }}')), { v: '\n'.repeat(n(30_000_000)) }, 'OUTPUT_TOO_LARGE']],
        ['a value of 15,000,000 spaces', () => [odtOf(p('{{ v }}')), { v: ' '.repeat(n(15_000_000)) }, undefined]],
        ['3,000,000 empty spans beside a field', () => [odtOf(p('{{ a }}' + '<text:span/>'.repeat(n(3_000_000)))), { a: 1 }, 'XML_ELEMENT_LIMIT_EXCEEDED']],
        ['3,000,000 text:s beside a field', () => [odtOf(p('{{ a }}' + '<text:s/>'.repeat(n(3_000_000)))), { a: 1 }, 'XML_ELEMENT_LIMIT_EXCEEDED']],
        ['a missing tag split across 400,000 spans', () => [odtOf(p('{{' + '<text:span>a</text:span>'.repeat(n(400000)) + '}}')), { a: 1 }, 'OUTPUT_TOO_LARGE']],
        ['a tag split across 400,000 empty spans', () => [odtOf(p('{{' + '<text:span></text:span>'.repeat(n(400000)) + 'a}}')), { a: 1 }, undefined]],
        ['20,000 text:s of 999,999 spaces each beside a field', () => [odtOf(p('{{ a }}' + '<text:s text:c="999999"/>'.repeat(n(20000)))), { a: 1 }, undefined]],
        ['16,000,000 fields in one paragraph', () => [odtOf(p('{{a}}'.repeat(n(16_000_000)))), { a: 1 }, 'XML_ELEMENT_LIMIT_EXCEEDED']],
        ['an if, then 40,000 else ifs in a run 40,000 spans deep', () => [odtOf(p(`{{ if a }}${'<text:span>'.repeat(n(40000))}${'{{ else if a }}'.repeat(n(40000))}{{ endif }}${'</text:span>'.repeat(n(40000))}`)), { a: false }, 'TEMPLATE_SYNTAX']],
        ['40,000 inline blocks 40,000 sections deep', () => [odtOf('<text:section>'.repeat(n(40000)) + p('{{ if a }}x{{ endif }}').repeat(n(40000)) + '</text:section>'.repeat(n(40000))), { a: 1 }, undefined]],
        ['64,000 nested spans with xml:ids in a loop', () => [odtOf(p('{{ for a in xs }}') + '<text:p>' + '<text:span xml:id="s1">'.repeat(n(64000)) + 'x' + '</text:span>'.repeat(n(64000)) + '</text:p>' + p('{{ endfor }}')), { xs: [1, 2] }, undefined]],
        ['a loop copying 50,000 named frames twice', () => [odtOf(p('{{ for a in xs }}') + p('<draw:frame draw:name="F" xml:id="f" draw:id="f"/>').repeat(n(50000)) + p('{{ endfor }}')), { xs: [0, 1] }, undefined]],
        ['a loop copying 50,000 bookmark ranges twice', () => [odtOf(p('{{ for a in xs }}') + p('<text:bookmark-start text:name="b"/>x<text:bookmark-end text:name="b"/>').repeat(n(50000)) + p('{{ endfor }}')), { xs: [0, 1] }, undefined]],
        ['100,000 orphaned bookmark ends after a branch', () => [odtOf(p('{{ if a }}x{{ endif }}') + p('<text:bookmark-end text:name="b"/>'.repeat(n(100000)))), { a: 1 }, undefined]],
        ['100,000 open bookmark starts after a branch', () => [odtOf(p('{{ if a }}x{{ endif }}') + p('<text:bookmark-start text:name="b"/>'.repeat(n(100000)))), { a: 1 }, undefined]],
        ['100,000 empty cells after a branch', () => [odtOf(p('{{ if a }}x{{ endif }}') + `<table:table><table:table-row>${'<table:table-cell/>'.repeat(n(100000))}</table:table-row></table:table>`), { a: 1 }, undefined]],
        ['400,000 comments before a block', () => [odtOf('<!---->'.repeat(n(400000)) + p('{{ if x }}y{{ endif }}')), { x: true }, undefined]],
        ['400,000 deleted paragraphs before a field', () => [odtOf(`<text:tracked-changes>${p('{{ a }}').repeat(n(400000))}</text:tracked-changes>` + p('{{ a }}')), { a: 1 }, undefined]],
        ['a header of 60,000 paragraph blocks', () => [odtOf(p('x'), (p('{{ if a }}') + p('b') + p('{{ endif }}')).repeat(n(60000))), { a: false }, undefined]],
        // Text that looks like a tag: every scan forward-only, each start looking at most a name ahead.
        ['ODT: 200,000 unclosed {{ a', () => [odtOf(p('{{ a '.repeat(n(200000)))), {}, undefined]],
        ['ODT: 200,000 { name }} in braces never closed', () => [odtOf(p('{'.repeat(n(200000)) + '{ name }}'.repeat(n(200000)))), {}, undefined]],
        ['ODT: 200,000 full-width {{ starts', () => [odtOf(p('\uFF5B\uFF5B nom '.repeat(n(200000)))), {}, undefined]],
        ['ODT: 200,000 names with spaces between delimiters', () => [odtOf(p('{{a b}}'.repeat(n(200000)))), {}, undefined]],
        ['ODT: 1,000,000 full-width names', () => [odtOf(p('\uFF5B\uFF5Ba\uFF5D\uFF5D'.repeat(n(1_000_000)))), {}, undefined]],
        // Look-alikes are no elements: past the 1,000 a template lists, they are counted and not kept.
        ['ODT: 3,000,000 full-width names, 1,000 of them listed', () => [odtOf(p('\uFF5B\uFF5Ba\uFF5D\uFF5D'.repeat(n(3_000_000)))), {}, undefined]],
        ['ODT: a name of 1,000,000 characters a brace short', () => [odtOf(p('{{ ' + 'a'.repeat(n(1_000_000)) + ' }')), {}, undefined]],
        ['DOCX: 200,000 unclosed {{ a', () => [docxOf(wp('{{ a '.repeat(n(200000)))), {}, undefined]],
        ['DOCX: 200,000 { name }} in braces never closed', () => [docxOf(wp('{'.repeat(n(200000)) + '{ name }}'.repeat(n(200000)))), {}, undefined]],
        ['DOCX: 200,000 full-width braces in 200,000 runs', () => [docxOf('<w:p>' + '<w:r><w:t>\uFF5B</w:t></w:r>'.repeat(n(200000)) + '</w:p>'), {}, undefined]],
        ['DOCX: 200,000 names with spaces split across runs', () => [docxOf('<w:p>' + '<w:r><w:t>{{a</w:t></w:r><w:r><w:t> b}}</w:t></w:r>'.repeat(n(200000)) + '</w:p>'), {}, undefined]],
        ['DOCX: 200,000 unknown filters', () => [docxOf(wp('{{ a | nope }}'.repeat(n(200000)))), {}, undefined]],
        // Numbers written as text, read with separators.
        ['a 1 MB text of digit groups read with dataNumberSeparators in 50,000 passes', () => [odtOf(p('{{ for i in xs }}{{ v | format "0" }}{{ endfor }}')), { xs: Array(n(50000)).fill(0), v: '1' + '.000'.repeat(n(250_000)) + 'x' }, 'TEMPLATE_DATA_INVALID', { dataNumberSeparators: { decimal: ',', group: '.' } }]],
        ['a list of 1 MB grouped number texts compared with a number', () => [odtOf(p('{{ for i in xs }}{{ if i > 1 }}y{{ endif }}{{ endfor }}')), { xs: Array(n(2000)).fill('1 000' + ' '.repeat(n(1_000_000))) }, 'OUTPUT_TOO_LARGE', { dataNumberSeparators: { decimal: ',', group: ' ' } }]],
        ['a 1 MB text of spaces compared with a number in 10,000 passes', () => [odtOf(p('{{ for i in xs }}{{ if v > 1 }}y{{ endif }}{{ endfor }}')), { xs: Array(n(10000)).fill(0), v: ' '.repeat(n(1_000_000)) + '2' }, 'OUTPUT_TOO_LARGE']],
    ] as [string, () => [Buffer, any, string | undefined, any?]][]) {
        let result: { out: Uint8Array | undefined; code: string | undefined } = { out: undefined, code: undefined };
        let wanted: string | undefined;
        await linear(`template: ${label}`, async () => {
            const [bytes, data, expected, cfg] = make();
            wanted = expected;
            result = await renderQuiet(bytes, data, cfg ?? {});
        }, { label: `template: ${label} without copying it per reference`, limit: 300_000_000 });
        check(`template: ${label} ${wanted ? `is refused with ${wanted}` : 'renders'}`, result.code === wanted && (wanted !== undefined || result.out instanceof Uint8Array), String(result.code));
    }
    // Prototype keys in an ODT template's paths are never read from the prototype; a data key named
    // __proto__ (as JSON.parse makes it) is the data's own.
    const protoList = await renderQuiet(odtOf(p('{{ for x in __proto__ }}[{{ x }}]{{ endfor }}{{ if constructor }}C{{ endif }}{{ if toString }}T{{ endif }}')), {}, { onMissing: 'empty' });
    check('ODT template: __proto__, constructor and toString are missing fields, not the prototype', !protoList.code && !strFromU8(unzipSync(protoList.out!)['content.xml']).match(/\[|C<|T</), String(protoList.code));
    const ownProto = await renderQuiet(odtOf(p('{{ for x in __proto__ }}[{{ x }}]{{ endfor }}')), JSON.parse('{"__proto__": ["own"]}'));
    check('ODT template: a data key named __proto__ is the data\'s own list', !ownProto.code && strFromU8(unzipSync(ownProto.out!)['content.xml']).includes('[own]'), String(ownProto.code));
    // Ids and names a template holds named as a prototype's keys are ids like any other.
    const protoIds = await renderQuiet(odtOf(p('{{ for a in xs }}<text:bookmark text:name="__proto__"/><text:section text:name="constructor"/><draw:frame draw:name="toString" xml:id="__proto__"/>{{ endfor }}')), { xs: [1, 2, 3] });
    const protoXml = protoIds.out ? strFromU8(unzipSync(protoIds.out)['content.xml']) : '';
    check('ODT template: ids named __proto__ or constructor are renumbered as any other', !protoIds.code && (protoXml.match(/text:name="__proto__/g) ?? []).length === 3 && new Set(protoXml.match(/text:name="__proto__[^"]*"/g)).size === 3, String(protoIds.code));
    // A zip bomb is bounded as a parse's is.
    const bomb = Buffer.from(zipSync({ mimetype: enc('application/vnd.oasis.opendocument.text'), 'content.xml': new Uint8Array(40_000_000) }));
    const bombed = await renderQuiet(bomb, {}, { decompressionLimits: { maxUncompressedBytes: 1_000_000 } });
    check('ODT template: an entry past the decompression limit is refused, typed', !!bombed.code && bombed.code !== 'TEMPLATE_UNSUPPORTED_FORMAT', String(bombed.code));
}

/**
 * Template text bounded by the XML it is read from: a `<text:s text:c="N"/>` is read as at most 16
 * spaces with delimiters of at most 7 spaces in a row, so an ODT paragraph's text is never longer than its
 * XML, and every later pass over it costs what the same text costs in a Word document.
 */

async function templateTextBoundTests() {
    console.log('- Templates: text read no longer than its XML, the look-alike scan and its list bounded, cut ids renewed...');
    const enc = (s: string) => new TextEncoder().encode(s);
    const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"';
    const odtOf = (body: string) => Buffer.from(zipSync({
        mimetype: [enc('application/vnd.oasis.opendocument.text'), { level: 0 }],
        'META-INF/manifest.xml': enc('<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/></manifest:manifest>'),
        'content.xml': enc(`<?xml version="1.0"?><office:document-content ${NS}><office:body><office:text>${body}</office:text></office:body></office:document-content>`),
    }));
    const docxOf = (body: string) => Buffer.from(zipSync({
        '[Content_Types].xml': enc('<Types/>'),
        'word/document.xml': enc(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`),
    }));
    const renderQuiet = async (bytes: Buffer, data: any, cfg: any = {}) => {
        const warnings: any[] = [];
        try { return { out: await OfficeTemplate.render(bytes, { data, onWarning: (w: any) => warnings.push(w), ...cfg }) as Uint8Array, code: undefined, warnings }; }
        catch (e: any) { return { out: undefined, code: e?.officeIssue?.code ?? String(e), warnings }; }
    };
    const p = (t: string) => `<text:p>${t}</text:p>`;
    const wp = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;

    // 22 bytes of XML read as 256 spaces made every later pass run over 11.6 times the text the limits allowed.
    const longSpaces = () => odtOf(p('{{' + '<text:s text:c="256"/>'.repeat(n(400_000)) + 'a}}'));
    let spaced: Awaited<ReturnType<typeof renderQuiet>> | undefined;
    await linear('template: an ODT paragraph of 400,000 <text:s text:c="256"/> inside a tag', async () => { spaced = await renderQuiet(longSpaces(), { a: 1 }); },
        { label: 'template: an ODT paragraph of 400,000 <text:s text:c="256"/> inside a tag, its heap', limit: 600_000_000 });
    check('template: an ODT paragraph of 400,000 <text:s text:c="256"/> inside a tag renders', spaced?.out instanceof Uint8Array, String(spaced?.code));
    const spacedRun = await withinTime(4000, () => renderQuiet(longSpaces(), { a: 1 }));
    check('template: an ODT paragraph of 400,000 <text:s text:c="256"/> inside a tag, read at most as long as its XML', spacedRun.within, `${spacedRun.ms}ms`);

    // The look-alike scan stops only where a delimiter can start: over text with no brace it costs about
    // what reading the text does (it allocated at every character, five times the reading).
    const spaces = ' '.repeat(20_000_000);
    const scanned = await withinTime(30_000, () => renderQuiet(docxOf(wp(`{{a}}${spaces}`)), { a: 1 }));
    const unscanned = await withinTime(30_000, () => renderQuiet(docxOf(wp('{{a}}') + wp(spaces)), { a: 1 }));
    check('template: the look-alike scan of a paragraph of 20,000,000 spaces costs about what reading it does',
        !scanned.result.code && !unscanned.result.code && scanned.ms <= 1.5 * unscanned.ms + await scaledMs(250), `${scanned.ms}ms with a brace in the paragraph, ${unscanned.ms}ms without; ${scanned.result.code ?? unscanned.result.code ?? ''}`);
    // Text made of braces, the scan's worst case: linear, each a few compares.
    for (const [label, make, texts] of [
        ['10,000,000 {} in one paragraph', () => docxOf(wp('{{ v }}') + wp('{}'.repeat(n(10_000_000)))), 0],
        ['20,000,000 full-width { in one paragraph', () => docxOf(wp('{{ v }}') + wp('\uFF5B'.repeat(n(20_000_000)))), 0],
        ['5,000,000 {a}} in one paragraph', () => docxOf(wp('{{ v }}') + wp('{a}}'.repeat(n(5_000_000)))), 1000],
        // Look-alikes are not XML elements: 8.6.0 rendered this at the default limits, and so does this.
        ['2,100,000 {a}} in one paragraph at the default limits', () => docxOf(wp('{{ v }}') + wp('{a}}'.repeat(n(2_100_000)))), 1000],
    ] as [string, () => Buffer, number][]) {
        let result: Awaited<ReturnType<typeof renderQuiet>> | undefined;
        await linear(`template: ${label}`, async () => { result = await renderQuiet(make(), { v: 1 }); },
            { label: `template: ${label}, its heap`, limit: 400_000_000 });
        const listed = result?.warnings.find(w => w.code === 'TEMPLATE_NOT_A_TAG')?.details.texts.length ?? 0;
        check(`template: ${label} renders, listing ${texts} look-alikes`, result?.out instanceof Uint8Array && listed === texts, `${result?.code ?? ''} listed ${listed}`);
    }
    // Braces are counted as a stack that never falls below zero, and pending missing ends as a stack of
    // levels: no allocation depends on a depth. Adversarial paragraphs of look-alikes and braces, each about
    // 2,000,000 characters, linear and flat outside the heap.
    const escaped = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    for (const [label, delimiters, make] of [
        ['500,000 missing ends {{a} left pending', { start: '{{', end: '}}' }, () => '{{a}'.repeat(n(500_000))],
        ['250,000 missing ends {{a} then 250,000 }', { start: '{{', end: '}}' }, () => '{{a}'.repeat(n(250_000)) + '}'.repeat(n(250_000))],
        ['250,000 {{ a }}} with the end }}}}', { start: '{{', end: '}}}}' }, () => '{{ a }}}'.repeat(n(250_000))],
        ['285,000 {{ a }} with the end }}}', { start: '{{', end: '}}}' }, () => '{{ a }}'.repeat(n(285_000))],
        ['1,000,000 { then 1,000,000 }', { start: '{{', end: '}}' }, () => '{'.repeat(n(1_000_000)) + '}'.repeat(n(1_000_000))],
        ['500,000 $$a$ with $$ for both delimiters', { start: '$$', end: '$$' }, () => '$$a$'.repeat(n(500_000))],
        ['500,000 <<a> with << and >>', { start: '<<', end: '>>' }, () => '<<a>'.repeat(n(500_000))],
        ['400,000 full-width starts', { start: '{{', end: '}}' }, () => '\uFF5B{a}}'.repeat(n(400_000))],
        ['500,000 missing starts {a}}', { start: '{{', end: '}}' }, () => '{a}}'.repeat(n(500_000))],
    ] as [string, { start: string; end: string }, () => string][]) {
        const bytes = () => docxOf(wp(escaped(`${delimiters.start} v ${delimiters.end}`)) + wp(escaped(make())));
        let result: Awaited<ReturnType<typeof renderQuiet>> | undefined;
        await linear(`template: look-alikes, ${label}`, async () => { result = await renderQuiet(bytes(), { v: 1 }, { delimiters }); },
            { label: `template: look-alikes, ${label}, its heap`, limit: 400_000_000 });
        // Outside the heap a render holds the document's own bytes a few times (unzipped, and the output);
        // a depth's allocation was 512 MB at a time.
        const input = bytes();
        const memory = await memoryOf(() => renderQuiet(input, { v: 1 }, { delimiters }));
        check(`template: look-alikes, ${label} render, flat outside the heap`, result?.out instanceof Uint8Array && memory.buffers < 128_000_000, `${result?.code ?? ''} buffers ${megabytes(memory.buffers)}`);
    }
    // The reviewer's case: 1,001 look-alikes use the room, then paragraphs of missing ends whose end is two
    // longer than the start (each allocated 512 MB, by a depth of -1, and took 5 seconds for 1,000).
    const negdepth = () => docxOf(wp('{{ v }}}}') + wp('{{ a }}}'.repeat(1001)) + wp('{{ a }}}{{ b }}}').repeat(n(1000)));
    await linear('template: 1,000 paragraphs of two missing ends past the room, end }}}}', () => renderQuiet(negdepth(), { v: 1 }, { delimiters: { start: '{{', end: '}}}}' } }));
    const negdepthInput = negdepth();
    const negdepthRun = await memoryOf(() => renderQuiet(negdepthInput, { v: 1 }, { delimiters: { start: '{{', end: '}}}}' } }));
    const negdepthWarning = negdepthRun.result.warnings.find(w => w.code === 'TEMPLATE_NOT_A_TAG');
    check('template: 1,000 paragraphs of two missing ends past the room are counted, flat outside the heap',
        negdepthRun.result.out instanceof Uint8Array && negdepthWarning?.details.found === 3001 && negdepthWarning?.details.texts.length === 1000 && negdepthRun.buffers < 16_000_000,
        `found ${negdepthWarning?.details.found}, listed ${negdepthWarning?.details.texts.length}, buffers ${megabytes(negdepthRun.buffers)}`);
    const negdepthTime = await withinTime(3000, () => renderQuiet(negdepth(), { v: 1 }, { delimiters: { start: '{{', end: '}}}}' } }));
    check('template: 1,000 paragraphs of two missing ends past the room render quickly', negdepthTime.within, `${negdepthTime.ms}ms`);
    // A template lists at most 1,000 look-alikes, and says how many it found.
    const many = await renderQuiet(docxOf(wp('{{ v }}') + wp('{a}}'.repeat(2_100_000))), { v: 1 });
    const manyWarning = many.warnings.find(w => w.code === 'TEMPLATE_NOT_A_TAG');
    check('template: 2,100,000 look-alikes are listed 1,000 times and counted in full', manyWarning?.details.texts.length === 1000 && manyWarning?.details.found === 2_100_000 && /2100000 texts .*only the first 1000 are listed/.test(manyWarning?.message ?? ''),
        JSON.stringify(manyWarning?.details.found));
    const wide = await renderQuiet(odtOf(p('\uFF5B\uFF5Ba\uFF5D\uFF5D'.repeat(3000))), {});
    const wideWarning = wide.warnings.find(w => w.code === 'TEMPLATE_NOT_A_TAG');
    // The half of a meta after a block's cut gets a fresh id: found once per cut, written once per pass.
    for (const [label, make, data] of [
        ['ODT: 20,000 inline blocks each cutting a meta with an id', () => odtOf(p('<text:meta xml:id="m">a{{ if x }}b</text:meta>c{{ endif }}'.repeat(n(20_000)))), { x: true }],
        ['ODT: a loop of 2,000 passes cutting 20 metas with ids', () => odtOf(p('{{ for i in xs }}' + '<text:meta xml:id="m">a{{ if x }}b</text:meta>c{{ endif }}'.repeat(20) + '{{ endfor }}')), { x: true, xs: Array(n(2000)).fill(0) }],
    ] as [string, () => Buffer, any][]) {
        let cut: Awaited<ReturnType<typeof renderQuiet>> | undefined;
        await linear(`template: ${label}`, async () => { cut = await renderQuiet(make(), data); });
        check(`template: ${label} renders`, cut?.out instanceof Uint8Array, String(cut?.code));
    }
    check('template: ODT, 3,000 full-width names are listed 1,000 times and counted in full', wideWarning?.details.texts.length === 1000 && wideWarning?.details.found === 3000, JSON.stringify(wideWarning?.details.found));
}

describe('officeParser template hardening tests (ported)', () => {
    it('templateHardeningTests', async () => { failed = 0; failures = []; await templateHardeningTests(); expectNoFailures(); }, 600_000);
    it('odtTemplateHardeningTests', async () => { failed = 0; failures = []; await odtTemplateHardeningTests(); expectNoFailures(); }, 600_000);
    it('templateTextBoundTests', async () => { failed = 0; failures = []; await templateTextBoundTests(); expectNoFailures(); }, 600_000);
});
