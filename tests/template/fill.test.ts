/**
 * document-kit's own tests of the template engine, on templates built with the `docx` library as an
 * application builds them. Made-up data only.
 */
import { describe, expect, it } from 'vitest';
import { AlignmentType, Footer, Header, Paragraph, Table, TableCell, TableRow, TextRun } from 'docx';
import {
    checkTemplate, checkTemplates, fillTemplate, fillTemplates, isTemplateError, OfficeErrorType, OfficeWarningType,
    TemplateError, templateTags, type OfficeIssue,
} from '../../src/template/index.js';
import { bodyText, paragraphs, partsOf, wordTemplate } from './helpers.js';

const sample = {
    site: 'Example Field Station',
    contact: 'A. Sample',
    total: 1234567.891,
    issued: '2026-03-05',
};

describe('fillTemplate', () => {
    it('fills a tag split across two runs of different formatting, in the first run\'s formatting', async () => {
        const template = await wordTemplate({
            children: [new Paragraph({ children: [new TextRun({ text: 'Prepared for {{ si', bold: true }), new TextRun({ text: 'te }}, with thanks.', italics: true })] })],
        });
        const out = await fillTemplate(template, { data: sample });
        const xml = partsOf(out)['word/document.xml']!;
        expect(bodyText(out)).toBe('Prepared for Example Field Station, with thanks.');
        // The value sits in the bold run its tag began in; the italic run keeps only what followed the tag.
        const runs = [...xml.matchAll(/<w:r>([\s\S]*?)<\/w:r>/g)].map(([, r]) => ({ bold: /<w:b\/>/.test(r!), italic: /<w:i\/>/.test(r!), text: r!.replace(/<[^>]+>/g, '') }));
        expect(runs).toEqual([
            { bold: true, italic: false, text: 'Prepared for Example Field Station' },
            { bold: false, italic: true, text: ', with thanks.' },
        ]);
    });

    it('keeps one branch of an inline condition and of a block condition', async () => {
        const template = await wordTemplate({
            children: [
                new Paragraph('Status: {{ if approved }}approved{{ else }}pending{{ endif }}.'),
                new Paragraph('{{ if remarks }}'),
                new Paragraph({ children: [new TextRun({ text: 'Remarks: {{ remarks }}', bold: true })] }),
                new Paragraph('{{ else }}'),
                new Paragraph('No remarks.'),
                new Paragraph('{{ endif }}'),
                new Paragraph('End.'),
            ],
        });
        const approved = bodyText(await fillTemplate(template, { data: { approved: true, remarks: 'Samples arrived cold.' } }));
        expect(approved).toBe('Status: approved.\nRemarks: Samples arrived cold.\nEnd.');
        const pending = bodyText(await fillTemplate(template, { data: { approved: false, remarks: '' } }));
        expect(pending).toBe('Status: pending.\nNo remarks.\nEnd.');
    });

    it('repeats a table\'s single body row once per item, for 60 items', async () => {
        const cell = (text: string) => new TableCell({ children: [new Paragraph(text)] });
        const template = await wordTemplate({
            children: [new Table({
                rows: [
                    new TableRow({ tableHeader: true, children: [cell('Sample'), cell('Value')] }),
                    new TableRow({ children: [cell('{{ for row in rows }}{{ row.id }}'), cell('{{ row.value | format "0.0" }}{{ endfor }}')] }),
                ],
            })],
        });
        const rows = Array.from({ length: 60 }, (_, i) => ({ id: `S-${String(i + 1).padStart(3, '0')}`, value: i / 4 }));
        const xml = partsOf(await fillTemplate(template, { data: { rows } }))['word/document.xml']!;
        const tableRows = [...xml.matchAll(/<w:tr>[\s\S]*?<\/w:tr>|<w:tr [\s\S]*?<\/w:tr>/g)].map(([row]) => paragraphs(row));
        expect(tableRows).toHaveLength(61);
        expect(tableRows[0]).toEqual(['Sample', 'Value']);
        expect(tableRows[1]).toEqual(['S-001', '0.0']);
        expect(tableRows[60]).toEqual(['S-060', '14.8']);
        expect(xml).not.toContain('{{');
        // The repeated row keeps the template row's properties; the heading row stays a heading row.
        expect((xml.match(/<w:tblHeader\/>/g) ?? []).length).toBe(1);
    });

    it('fills tags in a header and a footer', async () => {
        const template = await wordTemplate({
            headers: { default: new Header({ children: [new Paragraph('{{ site }} report')] }) },
            footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun('Contact: {{ contact }}')] })] }) },
            children: [new Paragraph('Body')],
        });
        const parts = partsOf(await fillTemplate(template, { data: sample }));
        const header = Object.entries(parts).find(([name]) => /^word\/header\d*\.xml$/.test(name))?.[1];
        const footer = Object.entries(parts).find(([name]) => /^word\/footer\d*\.xml$/.test(name))?.[1];
        expect(paragraphs(header!)).toEqual(['Example Field Station report']);
        expect(paragraphs(footer!)).toEqual(['Contact: A. Sample']);
        expect(footer).toContain('<w:jc w:val="right"/>');
    });

    it('writes numbers and dates through a format picture', async () => {
        const template = await wordTemplate({
            children: [
                new Paragraph('Total: {{ total | format "#,##0.00" }}'),
                new Paragraph('Issued: {{ issued | format "d MMMM yyyy" }} ({{ issued | format "ddd dd/MM/yy" }})'),
                new Paragraph('Received: {{ received | format "yyyy-MM-dd HH:mm" }}'),
            ],
        });
        const data = { ...sample, received: new Date(Date.UTC(2026, 2, 5, 23, 30)) };
        expect(bodyText(await fillTemplate(template, { data }))).toBe('Total: 1,234,567.89\nIssued: 5 March 2026 (Thu 05/03/26)\nReceived: 2026-03-05 23:30');
        // A fixed UTC offset and the document's own separators. A date written without a time is midnight
        // UTC, so eight hours west of Greenwich it is the evening before (as officeParser reads it).
        const local = bodyText(await fillTemplate(template, { data, utcOffsetMinutes: -480, numberSeparators: { decimal: ',', group: ' ' } }));
        expect(local).toBe('Total: 1 234 567,89\nIssued: 4 March 2026 (Wed 04/03/26)\nReceived: 2026-03-05 15:30');
    });

    it('takes a Uint8Array, an ArrayBuffer or a Blob', async () => {
        const template = await wordTemplate({ children: [new Paragraph('{{ site }}')] });
        const fromBytes = await fillTemplate(template, { data: sample });
        const fromBuffer = await fillTemplate(template.slice().buffer, { data: sample });
        const fromBlob = await fillTemplate(new Blob([template.slice()]), { data: sample });
        expect(fromBuffer).toEqual(fromBytes);
        expect(fromBlob).toEqual(fromBytes);
        await expect(fillTemplate('template.docx' as any, { data: sample, onWarning: () => {} })).rejects.toSatisfy((e: unknown) => isTemplateError(e, OfficeErrorType.INVALID_INPUT));
    });

    it('fills a batch, one document per entry, and refuses a list given to fillTemplate', async () => {
        const template = await wordTemplate({ children: [new Paragraph('Dear {{ contact }},')] });
        const docs = await fillTemplates(template, { data: [{ contact: 'A. Sample' }, { contact: 'B. Example' }] });
        expect(docs.map(bodyText)).toEqual(['Dear A. Sample,', 'Dear B. Example,']);
        await expect(fillTemplate(template, { data: [{ contact: 'x' }] as any, onWarning: () => {} })).rejects.toSatisfy((e: unknown) => isTemplateError(e, OfficeErrorType.IMPROPER_ARGUMENTS));
        await expect(fillTemplates(template, { data: { contact: 'x' } as any, onWarning: () => {} })).rejects.toSatisfy((e: unknown) => isTemplateError(e, OfficeErrorType.IMPROPER_ARGUMENTS));
    });

    it('rejects a missing field with onMissing: \'error\', keeps it by default, empties it with \'empty\'', async () => {
        const template = await wordTemplate({ children: [new Paragraph('Site: {{ site }}; lot {{ lot }}.')] });
        const error = await fillTemplate(template, { data: sample, onMissing: 'error', onWarning: () => {} }).then(() => null, (e: unknown) => e);
        expect(error).toBeInstanceOf(TemplateError);
        expect((error as TemplateError).code).toBe(OfficeErrorType.TEMPLATE_FIELD_MISSING);
        expect((error as TemplateError).officeIssue.details).toBe('lot');
        expect((error as TemplateError).message).toContain("Template field 'lot' has no matching value");
        expect(bodyText(await fillTemplate(template, { data: sample }))).toBe('Site: Example Field Station; lot {{ lot }}.');
        expect(bodyText(await fillTemplate(template, { data: sample, onMissing: 'empty' }))).toBe('Site: Example Field Station; lot .');
    });

    it('reports a template\'s own tag errors, typed', async () => {
        const template = await wordTemplate({ children: [new Paragraph('{{ if approved }}never closed')] });
        const error = await fillTemplate(template, { data: { approved: true }, onWarning: () => {} }).then(() => null, (e: unknown) => e) as TemplateError;
        expect(error.code).toBe(OfficeErrorType.TEMPLATE_SYNTAX);
        expect(error.officeIssue.details).toMatchObject({ reason: 'unclosed', part: 'word/document.xml' });
    });

    it('changes nothing but the tags: every other part is copied byte for byte', async () => {
        const template = await wordTemplate({ children: [new Paragraph({ children: [new TextRun({ text: '{{ site }}', font: 'Arial', size: 28 })] })] });
        const before = partsOf(template), after = partsOf(await fillTemplate(template, { data: sample }));
        expect(Object.keys(after)).toEqual(Object.keys(before));
        for (const name of Object.keys(before)) if (name !== 'word/document.xml') expect(after[name], name).toBe(before[name]);
        expect(after['word/document.xml']).toBe(before['word/document.xml']!.replace('{{ site }}', 'Example Field Station'));
    });
});

describe('templateTags', () => {
    it('lists fields, blocks and loops with what each reads, and text that looks like a tag', async () => {
        const warnings: OfficeIssue[] = [];
        const template = await wordTemplate({
            headers: { default: new Header({ children: [new Paragraph('{{ site | upper }}')] }) },
            children: [
                new Paragraph('{{ if approved and total > 100 }}Approved{{ else }}Pending{{ endif }}'),
                new Paragraph('{{ for row in rows }}{{ row.id }}{{ if not loop.last }}, {{ endif }}{{ endfor }}'),
                new Paragraph('{{first name}}'),
            ],
        });
        const tags = await templateTags(template, { onWarning: w => warnings.push(w) });
        expect(tags.map(t => [t.kind, t.part, t.tag, t.fields, t.loops.map(l => `${l.item} in ${l.list}`)])).toEqual([
            ['if', 'word/document.xml', '{{ if approved and total > 100 }}', ['approved', 'total'], []],
            ['else', 'word/document.xml', '{{ else }}', [], []],
            ['for', 'word/document.xml', '{{ for row in rows }}', ['rows'], []],
            ['field', 'word/document.xml', '{{ row.id }}', ['row.id'], ['row in rows']],
            ['if', 'word/document.xml', '{{ if not loop.last }}', ['loop.last'], ['row in rows']],
            ['text', 'word/document.xml', '{{first name}}', [], []],
            ['field', 'word/header1.xml', '{{ site | upper }}', ['site'], []],
        ]);
        expect(tags.find(t => t.kind === 'for')?.item).toBe('row');
        expect(tags.find(t => t.kind === 'text')?.reason).toBe('space-in-name');
        expect(tags.find(t => t.part === 'word/header1.xml')?.filters).toEqual(['upper']);
        expect(warnings).toEqual([]);
    });
});

describe('checkTemplate', () => {
    it('reports a missing field, where it is and in which pass of a loop, and nothing when the data is enough', async () => {
        const template = await wordTemplate({
            children: [
                new Paragraph('{{ site }} / {{ contact }}'),
                new Paragraph('{{ for row in rows }}{{ row.id }}={{ row.value }} {{ endfor }}'),
            ],
        });
        const problems = await checkTemplate(template, { data: { site: 'Example Field Station', rows: [{ id: 'S-1', value: 2 }, { id: 'S-2' }] } });
        expect(problems.map(p => [p.code, p.field, p.part, p.loops])).toEqual([
            [OfficeErrorType.TEMPLATE_FIELD_MISSING, 'contact', 'word/document.xml', []],
            [OfficeErrorType.TEMPLATE_FIELD_MISSING, 'row.value', 'word/document.xml', [{ list: 'rows', index: 1 }]],
        ]);
        expect(problems[0]!.message).toContain("Template field 'contact'");
        expect(await checkTemplate(template, { data: { site: 'x', contact: 'y', rows: [] } })).toEqual([]);
        const batch = await checkTemplates(template, { data: [{ site: 'x', contact: 'y', rows: [] }, { contact: 'y', rows: [] }] });
        expect(batch.map(list => list.map(p => p.field))).toEqual([[], ['site']]);
    });

    it('reports a value that does not fit its tag', async () => {
        const template = await wordTemplate({ children: [new Paragraph('{{ total | format "#,##0.00" }}')] });
        const [problem] = await checkTemplate(template, { data: { total: 'about ten' } });
        expect(problem?.code).toBe(OfficeErrorType.TEMPLATE_DATA_INVALID);
        expect(problem?.details).toMatchObject({ reason: 'unformattable', filter: 'format', takes: 'number' });
    });
});

describe('encrypted and other formats', () => {
    it('refuses a password-protected template with TEMPLATE_ENCRYPTED, whatever options are given', async () => {
        const { readFileSync } = await import('node:fs');
        const { join } = await import('node:path');
        for (const name of ['agile.docx', 'standard.docx', 'encrypted.odt']) {
            const bytes = new Uint8Array(readFileSync(join(__dirname, 'files', 'encrypted', name)));
            const error = await fillTemplate(bytes, { data: {}, onWarning: () => {} }).then(() => null, (e: unknown) => e) as TemplateError;
            expect(error?.code, name).toBe(OfficeErrorType.TEMPLATE_ENCRYPTED);
            expect(error.message).toContain('password-protected');
            await expect(templateTags(bytes, { onWarning: () => {} })).rejects.toSatisfy((e: unknown) => isTemplateError(e, OfficeErrorType.TEMPLATE_ENCRYPTED));
        }
    });

    it('refuses what is not a .docx, .odt or .ott', async () => {
        const warnings: OfficeIssue[] = [];
        const error = await fillTemplate(new TextEncoder().encode('plain text, not a document'), { data: {}, onWarning: w => warnings.push(w) }).then(() => null, (e: unknown) => e) as TemplateError;
        expect(error.code).toBe(OfficeErrorType.TEMPLATE_UNSUPPORTED_FORMAT);
        // (As in officeParser, the unzip's own refusal is reported too, before the call's.)
        expect(warnings.map(w => w.code)).toEqual([OfficeErrorType.ZIP_NO_ENTRIES_FOUND, OfficeErrorType.TEMPLATE_UNSUPPORTED_FORMAT]);
        expect(OfficeWarningType.TEMPLATE_NOT_A_TAG).toBe('TEMPLATE_NOT_A_TAG');
    });
});
