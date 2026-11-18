/**
 * The rest of officeParser 8.9.0's template tests (MIT, Harsh Ankur), ported to vitest:
 *
 * - test/generator/testOfficeGenerator.ts, runTemplatePropertyTests: the look-alike scan against a
 *   plain reference (lookalikeReference.ts) on random texts, and ODT runs of spaces as `<text:s/>`
 *   read as if typed out (odtSpaceRuns.ts). Upstream's fast mode sizes by default (20,000 texts, 1,000
 *   paragraphs); DOCUMENT_KIT_TEMPLATE_FULL=1 runs its full sizes (200,000 and 4,000).
 * - test/parser/testOfficeParser.ts: a hostile `<w:t>` run renders fast, and a tag split across runs
 *   is still filled (read back from document.xml, where upstream read it through its parser).
 * - test/testExhaustive.ts, testConfigValidation: a template's invalid `maxRepeatedContent` is
 *   reported and its default kept; `maxPasswordAttempts`, no option of the kit's, is reported as
 *   unrecognized (upstream: an invalid value of it), and an encrypted template is refused.
 */
import { describe, expect, it } from 'vitest';
import * as path from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { findLookalikeTags, PartContext, triggersOf } from '../../../src/template/templateEngine.js';
import { checkLookalikeScan } from './lookalikeReference.js';
import { checkOdtSpaceRuns } from './odtSpaceRuns.js';
import { FIXTURE_ROOT, OfficeTemplate, withinTime } from './harness.js';

const FULL = process.env.DOCUMENT_KIT_TEMPLATE_FULL === '1';

describe('officeParser template property tests (ported)', () => {
    it('the look-alike scan lists and counts what the reference finds', () => {
        const scan = (joined: string, start: string, end: string, room: number) => {
            const ctx = { opts: { start, end }, triggers: triggersOf({ start }), nearRoom: { left: room }, nearUnlisted: 0 } as unknown as PartContext;
            const found = findLookalikeTags(joined, ctx);
            return { found, unlisted: (ctx as any).nearUnlisted };
        };
        const { withLookalikes, failures } = checkLookalikeScan(scan, FULL ? 200_000 : 20_000, 1);
        expect(failures.slice(0, 3)).toEqual([]);
        expect(withLookalikes).toBeGreaterThan(0);
    }, 300_000);

    it('ODT spaces as <text:s/> render as typed out, next to tags with space edges', async () => {
        const failures = await checkOdtSpaceRuns((bytes, config) => OfficeTemplate.render(bytes, config) as Promise<Uint8Array>, FULL ? 4_000 : 1_000, 1);
        expect(failures).toEqual([]);
    }, 300_000);
});

describe('officeParser parser-suite template checks (ported)', () => {
    // buildDocxBytes, from test/parser/testOfficeParser.ts.
    const buildDocxBytes = (bodyXml: string): Uint8Array => {
        const xml = '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            + `<w:body>${bodyXml}</w:body></w:document>`;
        return zipSync({ 'word/document.xml': [strToU8(xml), { level: 0 }] }, { level: 0 });
    };

    it('a hostile <w:t> run is rendered fast', async () => {
        const hostileTemplate = buildDocxBytes('<w:p>' + '<w:t>'.repeat(80000) + '{{v}}</w:p>');
        const run = await withinTime(2000, async () => { await OfficeTemplate.render(hostileTemplate, { data: { v: 'OK' } }); });
        expect(run.within, run.detail).toBe(true);
    });

    it('a placeholder split across runs still substitutes', async () => {
        const split = buildDocxBytes('<w:p><w:r><w:t>Hi {{na</w:t></w:r><w:r><w:t>me}}</w:t></w:r><w:r><w:t/></w:r></w:p>');
        const out = await OfficeTemplate.render(split, { data: { name: 'World' } }) as Uint8Array;
        // document-kit: upstream read the text back with its parser.
        const text = strFromU8(unzipSync(out)['word/document.xml']!).replace(/<[^>]+>/g, '');
        expect(text).toContain('Hi World');
    });
});

describe('officeParser config checks of templates (ported)', () => {
    it('maxPasswordAttempts is not an option of the kit, and an encrypted template is refused', async () => {
        const agile = path.join(FIXTURE_ROOT, 'files', 'encrypted', 'agile.docx');
        const seen: any[] = [];
        let code: string | undefined;
        await OfficeTemplate.render(agile, { data: {}, maxPasswordAttempts: 1.5, onPassword: () => 'wrong', onWarning: (w: any) => seen.push(w) } as any).catch((e: any) => { code = e?.officeIssue?.code; });
        expect(code).toBe('TEMPLATE_ENCRYPTED');
        const unknown = seen.filter(w => w.code === 'UNRECOGNIZED_CONFIG_OPTION').flatMap(w => w.details.keys);
        expect(unknown.sort()).toEqual(['maxPasswordAttempts', 'onPassword']);
    });

    it('a template with an invalid maxRepeatedContent reports it and keeps the default bound', async () => {
        const loops = zipSync({ '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
            'word/document.xml': strToU8('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>{{ for a in xs }}{{ for b in xs }}{{ for c in xs }}{{ endfor }}{{ endfor }}{{ endfor }}</w:t></w:r></w:p></w:body></w:document>') });
        for (const bad of ['x', NaN, -1]) {
            const warned: any[] = [];
            let code = 'rendered';
            await OfficeTemplate.render(loops, { data: { xs: Array(300).fill(0) }, decompressionLimits: { maxRepeatedContent: bad }, onWarning: (w: any) => warned.push(w) } as any).catch((e: any) => { code = e?.officeIssue?.code; });
            expect(code, `maxRepeatedContent ${String(bad)}`).toBe('OUTPUT_TOO_LARGE');
            expect(warned.some(w => w.code === 'INVALID_CONFIG_VALUE')).toBe(true);
        }
    });
});
