/**
 * Errors and warnings of the template engine, taken from officeParser 8.9.0's `src/utils/errorUtils.ts`
 * (MIT, Harsh Ankur): the template's messages only, with the same wording. Errors are
 * {@link TemplateError}s (document-kit's own class) carrying officeParser's `officeIssue`, so code that
 * branches on `error.officeIssue.code` works unchanged.
 *
 * @module template/errors
 */

import { IssueHandler, OfficeErrorType, OfficeIssue, OfficeWarningType, TemplateDataDetails, TemplateNumbersAsTextDetails, TemplateSyntaxDetails } from './types.js';

/** Error header prefix for all error messages. */
const ERRORHEADER = '[document-kit template]: ';

/**
 * An error a template call rejects with. `officeIssue` is the structured issue that produced it (as in
 * officeParser), and `code` is its code.
 */
export class TemplateError extends Error {
    readonly officeIssue: OfficeIssue;
    constructor(issue: OfficeIssue) {
        super(ERRORHEADER + issue.message);
        this.name = 'TemplateError';
        this.officeIssue = issue;
    }
    /** The error's code (`TEMPLATE_FIELD_MISSING`, `TEMPLATE_SYNTAX`, ...). */
    get code(): OfficeErrorType {
        return this.officeIssue.code as OfficeErrorType;
    }
}

/** True when `error` is a {@link TemplateError}, optionally of one code. */
export function isTemplateError(error: unknown, code?: OfficeErrorType): error is TemplateError {
    return error instanceof TemplateError && (code === undefined || error.code === code);
}

/** The end of a TEMPLATE_SYNTAX message: what is wrong with the tag. */
function templateSyntaxReason(info: TemplateSyntaxDetails): string {
    switch (info.reason) {
        case 'unreadable': return `cannot be read: it ${info.detail}`;
        case 'unclosed': return `has no closing ${info.detail} tag`;
        case 'unopened': return `has no block to belong to${info.detail ? `: it needs an earlier ${info.detail}` : ''}`;
        case 'mismatched': return `closes a block that is not open: the open block needs ${info.detail} first`;
        case 'else-after-else': return 'follows its block\'s else; else comes last, before endif';
        case 'crossing': return `and the other tags of its block do not enclose content together: ${info.detail}`;
        case 'row-branch': return 'shares a table row with another tag of its block; in a table, each branch of an if starts a row of its own';
        case 'overlap': return 'opens a block that overlaps another block without holding it or lying inside it';
        case 'malformed-xml': return `is in a part whose XML is not well-formed${info.detail ? ` (${info.detail})` : ''}, so its blocks cannot be found`;
    }
}

/** The end of a TEMPLATE_DATA_INVALID message: how the data does not fit the tag. */
function templateDataReason(info: TemplateDataDetails): string {
    switch (info.reason) {
        case 'not-a-list': return `loops over a value that is ${info.kind}, not a list`;
        case 'not-text': return `fills in a value that is a${info.kind === 'object' ? 'n' : ''} ${info.kind}, which has no text of its own; loop over a list with for, or name a field inside an object with a path`;
        case 'incomparable': return `compares a value that is ${info.kind} with one that is ${info.otherKind}; a comparison takes two values of the same kind (or empty)`;
        case 'unformattable': return `formats a value that is ${info.kind} with ${info.filter}, which takes ${info.takes === 'date' ? 'a date, or text that is one in ISO 8601 form (2026-11-01)' : info.takes === 'number' ? 'a number, or text that is one' : 'text, a number, true/false or a date'}`;
        case 'ambiguous-number': {
            const readings = (info.readings ?? []).map(r => `${r.value} (decimal '${r.decimal}'${r.group ? `, group '${r.group}'` : ''})`);
            const listed = readings.length > 1 ? `${readings.slice(0, -1).join(', ')} and as ${readings[readings.length - 1]}` : readings.join('');
            return `reads text as a number ${['two', 'three', 'four', 'five', 'six', 'seven', 'eight'][readings.length - 2] ?? readings.length} ways with dataNumberSeparators, as ${listed}; write the number so that one way reads it, or list only the ways the data is written`;
        }
    }
}

/**
 * Lookup table for error messages.
 * Some entries are functions that take parameters to build dynamic messages.
 */
const ERROR_MESSAGES: Record<OfficeErrorType, string | ((...args: any[]) => string)> = {
    [OfficeErrorType.IMPROPER_ARGUMENTS]: (detail?: string) => typeof detail === 'string' && detail ? `Improper arguments: ${detail}` : `Improper arguments`,
    [OfficeErrorType.INVALID_INPUT]: `Invalid input type: Expected a Uint8Array, an ArrayBuffer or a Blob`,
    [OfficeErrorType.TEMPLATE_ENCRYPTED]: `This template is password-protected. document-kit does not decrypt templates: open it in Word or LibreOffice and save it without a password, then use that copy.`,
    [OfficeErrorType.TEMPLATE_UNSUPPORTED_FORMAT]: (detail: string) => `This document type cannot be used as a template${detail ? `: ${detail}` : ''}. Template rendering supports Word documents (.docx) and OpenDocument texts and text templates (.odt, .ott).`,
    [OfficeErrorType.TEMPLATE_FIELD_MISSING]: (field: string) => `Template field '${field}' has no matching value in the data, and onMissing is 'error'. Provide the field, or set onMissing to 'keep' or 'empty'.`,
    [OfficeErrorType.TEMPLATE_SYNTAX]: (info: TemplateSyntaxDetails) => `Template tag '${info.tag}' in ${info.part} ${templateSyntaxReason(info)}.`,
    [OfficeErrorType.TEMPLATE_DATA_INVALID]: (info: TemplateDataDetails) => `Template tag '${info.tag}' in ${info.part} ${templateDataReason(info)}.`,
    [OfficeErrorType.ZIP_ENTRY_COUNT_LIMIT_EXCEEDED]: (limit: number) => `ZIP entry count exceeds limit (${limit})`,
    [OfficeErrorType.ZIP_SIZE_LIMIT_EXCEEDED]: (limit: number) => `ZIP uncompressed size limit exceeded (${limit} bytes)`,
    [OfficeErrorType.XML_ELEMENT_LIMIT_EXCEEDED]: (limit: number) => `XML element limit exceeded: the document's XML holds more than ${limit} elements, comments and processing instructions, counting the lines, control words or parts of content embedded in it (decompressionLimits.maxXmlElements). Each takes several hundred bytes of memory to read; raise the limit, with memory to match, to parse larger documents.`,
    [OfficeErrorType.ZIP_NO_ENTRIES_FOUND]: `No readable entries found in ZIP data. The input is corrupt, truncated, or not a ZIP archive: every ZIP-based document format requires at least one entry.`,
    [OfficeErrorType.ZIP_TRUNCATED]: `Malformed ZIP data: no End of Central Directory record was found at the end of the input. Either the file was cut off during download or transfer, or extra data follows the archive; in both cases the entries recovered from it cannot be trusted to be the whole document.`,
    [OfficeErrorType.MAX_NESTING_DEPTH_EXCEEDED]: `Document nesting depth exceeded the safe limit (possible denial-of-service input)`,
    [OfficeErrorType.OUTPUT_TOO_LARGE]: (info?: { repeatedLimit?: number }) => info?.repeatedLimit !== undefined
        ? `The output is too large to write: the template's loops and values, repeated, add more than decompressionLimits.maxRepeatedContent (${info.repeatedLimit} characters) to the document. Raise the limit for a template that legitimately repeats a lot of content.`
        : `The output is too large to write: it grew past what a string or array can hold, or the AST shares its nodes along more paths than a writer follows (each shared node is written under every node holding it). Write the content to another format or in parts.`,
};

/** `text` shown to 80 characters: when longer, its first 77 and `...`, not splitting a surrogate pair. */
export const shownTo80 = (text: string): string => {
    if (text.length <= 80) return text;
    const code = text.charCodeAt(76);
    return `${text.slice(0, code >= 0xD800 && code <= 0xDBFF ? 76 : 77)}...`;
};

/**
 * A configuration value as a message writes it: a number as itself (NaN, -Infinity) and a bigint with its n
 * (10n); else as JSON, or as String gives it where JSON writes nothing (undefined, a function, a symbol); and
 * an object neither can write (one that refers to itself, holds a bigint, or has a toJSON, getter or proxy
 * that throws) by what it is. Never throws, so a mistaken value is reported, typed, whatever it is.
 */
function shownConfigValue(value: unknown): string {
    if (typeof value === 'number') return String(value);
    if (typeof value === 'bigint') return `${value}n`;
    try { return JSON.stringify(value) ?? String(value); } catch { /* written below */ }
    return 'an object that cannot be written out';
}

/**
 * Lookup table for warning messages.
 */
const WARNING_MESSAGES: Record<OfficeWarningType, string | ((...args: any[]) => string)> = {
    [OfficeWarningType.INVALID_CONFIG_VALUE]: (info: { option: string; value: unknown; expected: string; fallback: unknown }) => {
        const shown = shownConfigValue(info.value);
        const value = shownTo80(shown);
        const fallback = info.fallback === '' || info.fallback === undefined ? 'the default is' : `the default (${JSON.stringify(info.fallback)}) is`;
        return `Invalid ${info.option}: ${value}. Expected ${info.expected}; ${fallback} used instead.`;
    },
    [OfficeWarningType.TEMPLATE_NOT_A_TAG]: (info: { texts: { text: string; part: string; reason: string }[]; found: number }) => {
        const why: Record<string, string> = {
            'space-in-name': 'a name has no spaces', 'unknown-filter': 'a word after | is no filter',
            'missing-delimiter': 'a delimiter is a character short', 'full-width': 'its delimiters are typed full-width',
        };
        // At most 1,000 are listed: past them, only counted.
        const found = Math.max(info.found ?? 0, info.texts.length);
        const shown = info.texts.slice(0, 5).map(t => `'${t.text}' in ${t.part} (${why[t.reason] ?? t.reason})`).join(', ');
        const more = found > 5 ? `, and ${found - 5} more` : '';
        const cut = found > info.texts.length ? ` (only the first ${info.texts.length} are listed)` : '';
        const one = found === 1;
        return `${one ? 'Text that looks like a template tag is' : `${found} texts that look like template tags are`} not ${one ? 'a tag' : 'tags'} and ${one ? 'is' : 'are'} written as typed${cut}: ${shown}${more}.`;
    },
    [OfficeWarningType.TEMPLATE_NUMBERS_COMPARED_AS_TEXT]: (info: TemplateNumbersAsTextDetails) => {
        const answer = info.differs === true ? ', which here answers otherwise than comparing them as numbers'
            : info.differs === false ? ' (here the same as comparing them as numbers, which other values need not be)' : '';
        const entry = info.entry === undefined ? '' : ` (data entry ${info.entry})`;
        return `Template tag '${info.tag}' in ${info.part} compares two texts that read as numbers, '${info.left}' ${info.operator} '${info.right}', as text, character by character${answer}${entry}. Set numericTextComparison to 'number' to compare such texts as numbers, or to 'text' to keep comparing them as text without this warning.`;
    },
    [OfficeWarningType.TEMPLATE_CHECK_LIMITED]: (info: { limit: number; entry?: number; batch?: boolean; unchecked?: number }) => info.batch
        ? `A template check stopped after ${info.limit} problems in all, in data entry ${info.entry}${info.unchecked ? `; the ${info.unchecked} entr${info.unchecked === 1 ? 'y' : 'ies'} after it ${info.unchecked === 1 ? 'was' : 'were'} not checked` : ''}. Fix these and check again.`
        : `A template check stopped after ${info.limit} problems${info.entry !== undefined ? ` in data entry ${info.entry}` : ''}; there may be more. Fix these and check again.`,
    [OfficeWarningType.UNRECOGNIZED_CONFIG_OPTION]: (info: { keys: string[], renames?: Record<string, string> }) => {
        const detail = info.keys.map(k => {
            const replacement = info.renames?.[k];
            if (!replacement) return `'${k}'`;
            // A replacement starting with '(' is prose, not a config path: the option was removed with
            // no one-to-one successor, so the sentence is rendered as written instead of being quoted
            // as a key the caller could switch to.
            return replacement.startsWith('(') ? `'${k}' ${replacement}` : `'${k}' (use '${replacement}' instead)`;
        }).join(', ');
        return `Unrecognized config option${info.keys.length === 1 ? '' : 's'}: ${detail}. ${info.keys.length === 1 ? 'It was' : 'They were'} ignored and had no effect. Check for a typo or a key renamed in a major release.`;
    },
};

/**
 * Creates a formatted error message for a specific error type.
 * 
 * @param type - The type of error
 * @param info - Optional additional information (e.g., filepath, extension)
 * @returns The formatted error message string
 */
const createOfficeError = (type: OfficeErrorType, info?: any): string => {
    const msg = ERROR_MESSAGES[type];
    const message = typeof msg === 'function' ? msg(info) : msg;
    return message;
};

/**
 * Core reporting logic for all issues.
 * Ensures consistent logging and callback execution.
 */
const reportIssue = (
    issue: OfficeIssue,
    config?: IssueHandler
): void => {
    if (config?.onWarning) {
        config.onWarning(issue);
    } else if (!config) {
        // No handler and no config context at all: fall back to the console.
        const formatted = ERRORHEADER + issue.message;
        if (issue.type === 'error') {
            console.error(formatted, issue.details || '');
        } else {
            console.warn(formatted, issue.details || '');
        }
    }
};

/**
 * Creates, reports (to `config.onWarning`, or the console when no config is given) and returns a
 * template error.
 */
export const getOfficeError = (type: OfficeErrorType, config?: IssueHandler, info?: any): TemplateError => {
    const error = buildOfficeError(type, info);
    reportIssue(error.officeIssue, config);
    return error;
};

/** A template error, built without reporting it anywhere. */
export const buildOfficeError = (type: OfficeErrorType, info?: any): TemplateError => {
    const issue: OfficeIssue = {
        type: 'error',
        code: type,
        message: createOfficeError(type, info),
        details: info
    };
    return new TemplateError(issue);
};

/**
 * Reports a non-fatal issue to `config.onWarning`, or to the console only when no config context is
 * supplied at all.
 */
export const logWarning = (type: OfficeWarningType, config?: IssueHandler, info?: any): void => {
    const msg = WARNING_MESSAGES[type];
    const message = typeof msg === 'function' ? msg(info) : msg;
    const issue: OfficeIssue = {
        type: 'warning',
        code: type,
        message,
        details: info
    };
    reportIssue(issue, config);
};
