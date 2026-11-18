/**
 * The sentence filters shared by Word templates and recipes: capitalize, plural, prose and ranges. Made-up data only.
 */
import { describe, expect, it } from 'vitest';
import { Paragraph } from 'docx';
import { fillTemplate, templateTags } from '../../src/template/index.js';
import { capitalize, plural, prose, ranges, SentenceFilterError } from '../../src/template/sentenceFilters.js';
import { bodyText, wordTemplate } from './helpers.js';

describe('the sentence filters', () => {
    it('capitalises a plain lower-case first word, never one cased on purpose or a lone letter', () => {
        expect(capitalize('three samples')).toBe('Three samples');
        expect(capitalize('  two')).toBe('  Two');
        expect(capitalize('pH of 7')).toBe('pH of 7');
        expect(capitalize('mL')).toBe('mL');
        expect(capitalize('a')).toBe('a');
        expect(capitalize('a dried sample')).toBe('A dried sample');
        expect(capitalize('12 bottles')).toBe('12 bottles');
        expect(capitalize(null)).toBe(null);
    });

    it('chooses a word by a count, a list counting as its length', () => {
        expect(plural(1, 'sample', 'samples')).toBe('sample');
        expect(plural(0, 'sample', 'samples')).toBe('samples');
        expect(plural('2', 'batch', 'batches')).toBe('batches');
        expect(plural(['x'], 'bottle')).toBe('bottle');
        expect(plural(['x', 'y'], 'bottle')).toBe('bottles');
        expect(() => plural('several', 'bottle')).toThrow(SentenceFilterError);
    });

    it('joins a list as prose', () => {
        expect(prose(['a'])).toBe('a');
        expect(prose(['a', 'b'])).toBe('a and b');
        expect(prose(['a', 'b', 'c'])).toBe('a, b and c');
        expect(prose(['a', '', null, 'c'], 'or')).toBe('a or c');
        expect(prose(['a', 'b', 'c'], 'and', '; ')).toBe('a; b and c');
        expect(prose('just one')).toBe('just one');
    });

    it('collapses runs of consecutive values to ranges', () => {
        expect(ranges([1, 2, 3, 5, 7, 8])).toEqual(['1–3', '5', '7', '8']);
        expect(ranges(['S-1', 'S-2', 'S-3', 'S-4', 'T-5', 'S-6'])).toEqual(['S-1–S-4', 'T-5', 'S-6']);
        expect(ranges([7, 8], ' to ', 2)).toEqual(['7 to 8']);
        expect(ranges(['007', '008', '009', '010'])).toEqual(['007–010']);
        expect(prose(ranges(['W-1', 'W-2', 'W-3', 'W-9']) as string[])).toBe('W-1–W-3 and W-9');
    });
});

describe('in a Word template', () => {
    it('reads them as filters, chained, and lets a caller\'s own filter of the same name win', async () => {
        const template = await wordTemplate({
            children: [
                new Paragraph('{{ count | plural "sample was" "samples were" | capitalize }} received: {{ ids | ranges | prose }}.'),
                new Paragraph('{{ n }} {{ n | plural "bottle" }}, {{ word | capitalize }}.'),
            ],
        });
        expect((await templateTags(template)).filter((tag) => tag.kind === 'text')).toEqual([]);
        const out = await fillTemplate(template, { data: { count: 3, ids: ['S-1', 'S-2', 'S-3', 'S-7'], n: 1, word: 'fine' } });
        expect(bodyText(out)).toBe('Samples were received: S-1–S-3 and S-7.\n1 bottle, Fine.');
        const own = await fillTemplate(template, { data: { count: 3, ids: [], n: 1, word: 'fine' }, formatters: { capitalize: (v) => String(v).toUpperCase() } });
        expect(bodyText(own)).toBe('SAMPLES WERE received: .\n1 bottle, FINE.');
    });
});
