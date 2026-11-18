/**
 * document-kit/template: fills a Word (`.docx`) or OpenDocument text (`.odt`, `.ott`) template's tags
 * from data. `{{ field }}`, `{{ if … }} … {{ else }} … {{ endif }}`, `{{ for item in list }} … {{ endfor }}`
 * (down to one table row per item) and value formats such as `{{ total | format "#,##0.00" }}`. The
 * template is copied and only its tags change, so its formatting survives.
 *
 * The engine is officeParser 8.9.0's (MIT, Harsh Ankur; see NOTICE-OFFICEPARSER.md), moved here to run
 * in a browser. The options keep officeParser's names, so its documentation of `OfficeTemplate` applies:
 * `fillTemplate`/`fillTemplates` are its `render`, `templateTags` its `tags`, `checkTemplate`/
 * `checkTemplates` its `check`. A guide for template authors is docs/template.md.
 *
 * Runs anywhere with `TextDecoder`: no Node, no DOM, no network.
 *
 * @module template
 */

import { buildOfficeError } from './errors.js';
import { check, render, tags, TemplateInput } from './officeTemplate.js';
import { OfficeErrorType, TemplateConfig, TemplateData, TemplateProblem, TemplateTagInfo } from './types.js';

export { isTemplateError, TemplateError } from './errors.js';
export type { TemplateInput } from './officeTemplate.js';
export { OfficeErrorType, OfficeWarningType } from './types.js';
export type {
    BlobLike,
    DecompressionLimits,
    OfficeIssue,
    TemplateConfig,
    TemplateData,
    TemplateDataDetails,
    TemplateDataReason,
    TemplateDateNames,
    TemplateFormatter,
    TemplateLoopPlace,
    TemplateNumberReading,
    TemplateNumbersAsTextDetails,
    TemplateProblem,
    TemplateSyntaxDetails,
    TemplateSyntaxReason,
    TemplateTagInfo,
    TemplateTextReason,
    TemplateValue,
} from './types.js';

/** The options of a fill or a check besides `data` (officeParser's `TemplateConfig`). */
export type TemplateOptions = Omit<TemplateConfig, 'data'>;

/** A typed refusal of data of the wrong shape for the function called, reported as officeParser reports one. */
function wrongShape(options: TemplateOptions | undefined, detail: string): Error {
    const error = buildOfficeError(OfficeErrorType.IMPROPER_ARGUMENTS, detail);
    options?.onWarning?.(error.officeIssue);
    return error;
}

/**
 * Fills a template from one data object and returns the new document's bytes (a `.docx` for a Word
 * template, an `.odt` for an OpenDocument one). Rejects with a {@link TemplateError}.
 */
export async function fillTemplate(template: TemplateInput, options: TemplateOptions & { data: TemplateData }): Promise<Uint8Array> {
    if (Array.isArray(options?.data)) throw wrongShape(options, 'fillTemplate takes one data object; use fillTemplates for a list');
    return await render(template, options) as Uint8Array;
}

/**
 * Fills a template once per data object (a mail-merge): one document per entry, in order. The template
 * is read and compiled once. Rejects with a {@link TemplateError}.
 */
export async function fillTemplates(template: TemplateInput, options: TemplateOptions & { data: TemplateData[] }): Promise<Uint8Array[]> {
    if (!Array.isArray(options?.data)) throw wrongShape(options, 'fillTemplates takes a list of data objects; use fillTemplate for one');
    return await render(template, options) as Uint8Array[];
}

/**
 * Lists a template's tags in the order of each part: each field, and each tag that opens a block or a
 * branch, with what it reads from the data and the loops it is inside; and text that looks like a tag
 * but is not one (`kind: 'text'`, with why). No data is needed. A template's own tag errors reject.
 */
export function templateTags(template: TemplateInput, options: TemplateOptions = {}): Promise<TemplateTagInfo[]> {
    return tags(template, options);
}

/**
 * Checks one data object against a template without writing a document: every problem a fill with
 * `onMissing: 'error'` would reject for (a missing field, a value that does not fit its tag), whatever
 * `onMissing` says. An empty list means the data is enough.
 */
export async function checkTemplate(template: TemplateInput, options: TemplateOptions & { data: TemplateData }): Promise<TemplateProblem[]> {
    if (Array.isArray(options?.data)) throw wrongShape(options, 'checkTemplate takes one data object; use checkTemplates for a list');
    return await check(template, options) as TemplateProblem[];
}

/** Checks each data object of a batch against a template: one list of problems per entry. */
export async function checkTemplates(template: TemplateInput, options: TemplateOptions & { data: TemplateData[] }): Promise<TemplateProblem[][]> {
    if (!Array.isArray(options?.data)) throw wrongShape(options, 'checkTemplates takes a list of data objects; use checkTemplate for one');
    return await check(template, options) as TemplateProblem[][];
}
