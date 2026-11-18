/**
 * The template engine's defaults and option readers, taken from officeParser 8.9.0's `src/defaults.ts`
 * (`DEFAULT_TEMPLATE_CONFIG`, the parser's `decompressionLimits`) and `src/utils/configUtils.ts`
 * (`decompressionLimitsOf`, `PROTOTYPE_POLLUTION_KEYS`), MIT, Harsh Ankur. Values unchanged.
 *
 * @module template/defaults
 */

import { logWarning } from './errors.js';
import { DecompressionLimits, IssueHandler, OfficeWarningType, TemplateConfig, TemplateDateNames } from './types.js';

/** The decompression limits' defaults (officeParser's `DEFAULT_OFFICE_PARSER_CONFIG.decompressionLimits`). */
export const DEFAULT_DECOMPRESSION_LIMITS: Required<DecompressionLimits> = {
    maxUncompressedBytes: 512 * 1024 * 1024,
    maxZipEntries: 10000,
    maxTableCells: 1000000,
    maxXmlElements: 2000000,
    maxRawContentLength: 64 * 1024 * 1024,
    maxRepeatedContent: 16 * 1024 * 1024,
};

/**
 * Defaults for a template call: the tag delimiters, what a missing field does, what a picture writes
 * (separators, English month and day names, dates in UTC), and what a render may add by repetition.
 */
export const DEFAULT_TEMPLATE_CONFIG: Required<Pick<TemplateConfig, 'delimiters' | 'onMissing' | 'utcOffsetMinutes'>>
    & { numberSeparators: Required<NonNullable<TemplateConfig['numberSeparators']>>; dateNames: Required<TemplateDateNames>; maxRepeatedContent: number } = {
    delimiters: { start: '{{', end: '}}' },
    onMissing: 'keep',
    numberSeparators: { decimal: '.', group: ',' },
    dateNames: {
        months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
        monthsShort: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
        days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
        daysShort: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
        am: 'AM',
        pm: 'PM',
    },
    utcOffsetMinutes: 0,
    maxRepeatedContent: DEFAULT_DECOMPRESSION_LIMITS.maxRepeatedContent,
};

/**
 * Keys that must never be copied from a caller-supplied config onto one of our objects: a config that
 * arrived via `JSON.parse` can carry `__proto__` as an own enumerable property.
 */
export const PROTOTYPE_POLLUTION_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

/**
 * `decompressionLimits` with each limit as given when it is a number of 0 or more (Infinity turns a bound
 * off); another value is reported (INVALID_CONFIG_VALUE) and the default used.
 */
export function decompressionLimitsOf(limits: DecompressionLimits | undefined, config: IssueHandler): Required<DecompressionLimits> {
    const defaults = DEFAULT_DECOMPRESSION_LIMITS;
    const out = { ...defaults } as Required<DecompressionLimits>;
    // Not an object of limits at all (`'x'`, a list): reported, and every default used.
    if (limits !== undefined && limits !== null && (typeof limits !== 'object' || Array.isArray(limits))) {
        logWarning(OfficeWarningType.INVALID_CONFIG_VALUE, config, { option: 'decompressionLimits', value: limits, expected: 'an object of limits by name', fallback: undefined });
        return out;
    }
    for (const key of Object.keys(defaults) as (keyof Required<DecompressionLimits>)[]) {
        const value = limits?.[key];
        if (value === undefined) continue; // left out (or set to undefined): the default
        if (typeof value === 'number' && value >= 0) { out[key] = value; continue; }
        logWarning(OfficeWarningType.INVALID_CONFIG_VALUE, config, { option: `decompressionLimits.${key}`, value, expected: 'a number of 0 or more', fallback: defaults[key] });
    }
    return out;
}
